/**
 * @file src/tray_windows.c
 * @brief System tray implementation for Windows.
 */
// standard includes
#include <windows.h>
#include <strsafe.h>
// clang-format off
// build fails if shellapi.h is included before windows.h
#include <shellapi.h>
// clang-format on
#ifndef ARRAYSIZE
#define ARRAYSIZE(a) (sizeof(a)/sizeof((a)[0]))
#endif
// std C
#include <stdarg.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

// local includes
#include "tray.h"

#define WM_TRAY_CALLBACK_MESSAGE (WM_USER + 1)  ///< Tray callback message.
#define WC_TRAY_CLASS_NAME "TRAY"  ///< Tray window class name.
#define ID_TRAY_FIRST 1000  ///< First tray identifier.
#define ID_TRAY_RETRY_TIMER 1  ///< Timer that retries notification icon registration.
#define TRAY_RETRY_INTERVAL_MS 5000  ///< Interval between icon registration retries.
#define TRAY_RETRY_LOG_PERIOD 60  ///< Log a retry failure at WARNING once per this many attempts.
#define TRAY_NOTIFICATION_REPLAY_TTL_MS (3 * 60 * 1000)  ///< Replay a remembered notification after re-registration only within this window.

/**
 * @brief Icon information.
 */
struct icon_info {
  const char *path;  ///< Path to the icon
  HICON icon;  ///< Regular icon
  HICON large_icon;  ///< Large icon
  HICON notification_icon;  ///< Notification icon
};

/**
 * @brief Icon type.
 */
enum IconType {
  REGULAR = 1,  ///< Regular icon
  LARGE,  ///< Large icon
  NOTIFICATION  ///< Notification icon
};

static WNDCLASSEXA wc;
static NOTIFYICONDATAA nid;
static HWND hwnd;
static HMENU hmenu = NULL;
static void (*notification_cb)() = 0;
static UINT wm_taskbarcreated;
static struct tray *g_tray = NULL;  // remember last tray so we can re-apply after Explorer restarts

static BOOL icon_added = FALSE;  // whether the shell currently has our notification icon
static unsigned int icon_add_failures = 0;
static ULONGLONG notification_posted_ms = 0;  // GetTickCount64() when the app last posted notification text

static struct icon_info *icon_infos;
static HMENU _tray_menu(struct tray_menu *m, UINT *id);
static HICON _fetch_icon(const char *path, enum IconType icon_type);
static int tray_try_add_icon(void);
static void tray_apply_state(struct tray *tray, BOOL is_replay);

static tray_log_callback g_tray_log_cb = NULL;

void tray_set_log_callback(tray_log_callback cb) {
  g_tray_log_cb = cb;
}

static void tray_log(enum tray_log_level level, const char *fmt, ...) {
  if (!g_tray_log_cb || !fmt) {
    return;
  }
  char buffer[1024];
  va_list args;
  va_start(args, fmt);
  vsnprintf(buffer, sizeof(buffer), fmt, args);
  va_end(args);
  buffer[sizeof(buffer) - 1] = '\0';
  g_tray_log_cb(level, buffer);
}

static void tray_log_last_error(enum tray_log_level level, const char *context) {
  DWORD err = GetLastError();
  char message[512] = {0};
  DWORD flags = FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
  DWORD len = FormatMessageA(flags, NULL, err, 0, message, (DWORD)sizeof(message), NULL);
  if (len == 0 || message[0] == '\0') {
    tray_log(level, "%s failed (err=%lu; no extended error message)", context, (unsigned long)err);
  } else {
    // Trim trailing newlines from FormatMessageA
    while (len > 0 && (message[len - 1] == '\n' || message[len - 1] == '\r')) {
      message[len - 1] = '\0';
      --len;
    }
    tray_log(level, "%s failed (err=%lu): %s", context, (unsigned long)err, message);
  }
}

// Safe copy that always NUL-terminates
static void safe_copy_sz(char *dst, size_t dstcch, const char *src) {
  if (!dst || dstcch == 0) return;
  if (!src) { dst[0] = '\0'; return; }
  StringCchCopyA(dst, dstcch, src);
}

static int icon_info_count;

static DWORD tray_apply_icon_and_tip(struct tray *tray, DWORD flags) {
  nid.hIcon = NULL;
  if (tray != NULL && tray->icon != NULL && tray->icon[0] != '\0') {
    HICON icon = _fetch_icon(tray->icon, REGULAR);
    if (icon != NULL) {
      nid.hIcon = icon;
      flags |= NIF_ICON;
    }
  }

  if (tray != NULL && tray->tooltip != NULL && tray->tooltip[0] != '\0') {
    safe_copy_sz(nid.szTip, ARRAYSIZE(nid.szTip), tray->tooltip);
    flags |= NIF_TIP;
#ifdef NIF_SHOWTIP
    flags |= NIF_SHOWTIP;
#endif
  } else {
    nid.szTip[0] = '\0';
  }

  return flags;
}

static int tray_add_notify_icon(struct tray *tray, enum tray_log_level failure_level) {
  nid.uFlags = tray_apply_icon_and_tip(tray, NIF_MESSAGE);
  nid.uCallbackMessage = WM_TRAY_CALLBACK_MESSAGE;
  if (!Shell_NotifyIconA(NIM_ADD, &nid)) {
    // The shell may still be tracking a half-registered icon for this identity
    // (e.g. a previous instance that died mid-update). Clear it and try once more.
    Shell_NotifyIconA(NIM_DELETE, &nid);
    if (!Shell_NotifyIconA(NIM_ADD, &nid)) {
      tray_log_last_error(failure_level, "Shell_NotifyIconA(NIM_ADD)");
      return -1;
    }
  }

  nid.uVersion = NOTIFYICON_VERSION_4;
  if (!Shell_NotifyIconA(NIM_SETVERSION, &nid)) {
    tray_log_last_error(TRAY_LOG_WARNING, "Shell_NotifyIconA(NIM_SETVERSION)");
  }

  return 0;
}

static void tray_schedule_icon_retry(void) {
  if (hwnd != NULL) {
    SetTimer(hwnd, ID_TRAY_RETRY_TIMER, TRAY_RETRY_INTERVAL_MS, NULL);
  }
}

// Try to (re-)register the notification icon. The shell can refuse NIM_ADD for
// long stretches (around logon, Explorer crashes, installer-driven restarts), so
// failures are not fatal: a timer keeps retrying until the shell accepts. Failure
// logs are throttled to one WARNING per TRAY_RETRY_LOG_PERIOD attempts so a
// persistently broken shell does not flood the log.
static int tray_try_add_icon(void) {
  if (g_tray == NULL || hwnd == NULL) {
    return -1;
  }

  enum tray_log_level level = (icon_add_failures % TRAY_RETRY_LOG_PERIOD == 0) ? TRAY_LOG_WARNING : TRAY_LOG_DEBUG;
  if (tray_add_notify_icon(g_tray, level) < 0) {
    ++icon_add_failures;
    icon_added = FALSE;
    tray_schedule_icon_retry();
    return -1;
  }

  if (icon_add_failures > 0) {
    tray_log(TRAY_LOG_INFO, "Tray icon registered after %u failed attempts", icon_add_failures);
  }
  icon_add_failures = 0;
  icon_added = TRUE;
  KillTimer(hwnd, ID_TRAY_RETRY_TIMER);
  tray_apply_state(g_tray, TRUE);
  return 0;
}

// Explorer broadcasts TaskbarCreated at medium integrity. When this process runs
// elevated (or as SYSTEM), UIPI silently drops that broadcast unless we opt in,
// which would leave the icon permanently missing after an Explorer restart.
static void tray_allow_taskbar_created(HWND wnd) {
  typedef BOOL(WINAPI * change_window_message_filter_ex_t)(HWND, UINT, DWORD, void *);
  HMODULE user32 = GetModuleHandleA("user32.dll");
  if (user32 == NULL) {
    return;
  }
  change_window_message_filter_ex_t change_filter =
    (change_window_message_filter_ex_t) GetProcAddress(user32, "ChangeWindowMessageFilterEx");
  if (change_filter == NULL) {
    return;
  }
  if (!change_filter(wnd, wm_taskbarcreated, 1 /* MSGFLT_ALLOW */, NULL)) {
    tray_log_last_error(TRAY_LOG_WARNING, "ChangeWindowMessageFilterEx(TaskbarCreated)");
  }
}

static LRESULT CALLBACK _tray_wnd_proc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  switch (msg) {
    case WM_CLOSE:
      DestroyWindow(hwnd);
      return 0;
    case WM_DESTROY:
      PostQuitMessage(0);
      return 0;
    case WM_TIMER:
      if (wparam == ID_TRAY_RETRY_TIMER) {
        if (icon_added) {
          KillTimer(hwnd, ID_TRAY_RETRY_TIMER);
        } else {
          tray_try_add_icon();
        }
        return 0;
      }
      break;
    case WM_COMMAND: {
      if (HIWORD(wparam) == 0) {
        const UINT cmd_id = LOWORD(wparam);
        MENUITEMINFOA item_info;
        memset(&item_info, 0, sizeof(item_info));
        item_info.cbSize = sizeof(item_info);
        item_info.fMask = MIIM_DATA | MIIM_STATE;
        if (GetMenuItemInfoA(hmenu, cmd_id, FALSE, &item_info) && item_info.dwItemData != 0) {
          struct tray_menu *menu = (struct tray_menu *) item_info.dwItemData;
          if (menu->checkbox) {
            menu->checked = !menu->checked;
            item_info.fMask = MIIM_STATE;
            item_info.fState = menu->checked ? MFS_CHECKED : 0;
            SetMenuItemInfoA(hmenu, cmd_id, FALSE, &item_info);
          }
          if (menu->cb) {
            menu->cb(menu);
          }
        }
      }
      return 0;
    }
    case WM_TRAY_CALLBACK_MESSAGE: {
      switch (LOWORD(lparam)) {
        case WM_LBUTTONUP:
        case WM_RBUTTONUP:
        case WM_CONTEXTMENU: {
          POINT p;
          GetCursorPos(&p);
          SetForegroundWindow(hwnd);
          WORD cmd = (WORD)TrackPopupMenu(
            hmenu,
            TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY,
            p.x, p.y, 0, hwnd, NULL);
          if (cmd) {
            SendMessage(hwnd, WM_COMMAND, cmd, 0);
          }
          // Ensure the menu dismisses properly (MSDN guidance)
          PostMessage(hwnd, WM_NULL, 0, 0);
          return 0;
        }

        case NIN_BALLOONUSERCLICK:
          if (notification_cb) {
            notification_cb();
          }
          return 0;
      }
      break;
    }
  }

  // Handle Explorer restarts: the old registration is gone, so re-add the icon
  // and re-apply state (tray_try_add_icon keeps retrying on failure).
  if (msg == wm_taskbarcreated) {
    icon_added = FALSE;
    tray_try_add_icon();
    return 0;
  }

  return DefWindowProc(hwnd, msg, wparam, lparam);
}

static HMENU _tray_menu(struct tray_menu *m, UINT *id) {
  HMENU hmenu = CreatePopupMenu();
  for (; m != NULL && m->text != NULL; m++, (*id)++) {
    if (strcmp(m->text, "-") == 0) {
      InsertMenuA(hmenu, *id, MF_SEPARATOR, 0, NULL);
    } else {
      MENUITEMINFOA item;
      memset(&item, 0, sizeof(item));
      item.cbSize = sizeof(MENUITEMINFOA);
      item.fMask = MIIM_ID | MIIM_TYPE | MIIM_STATE | MIIM_DATA;
      item.fType = 0;
      item.fState = 0;
      if (m->submenu != NULL) {
        item.fMask |= MIIM_SUBMENU;
        item.hSubMenu = _tray_menu(m->submenu, id);
      }
      if (m->disabled) {
        item.fState |= MFS_DISABLED;
      }
      if (m->checked) {
        item.fState |= MFS_CHECKED;
      }
      item.wID = *id;
      item.dwTypeData = (LPSTR) m->text;
      item.dwItemData = (ULONG_PTR) m;

      InsertMenuItemA(hmenu, *id, TRUE, &item);
    }
  }
  return hmenu;
}

/**
 * @brief Create icon information.
 * @param path Path to the icon.
 * @return Icon information.
 */
struct icon_info _create_icon_info(const char *path) {
  struct icon_info info;
  info.path = strdup(path);

  // These must be separate invocations otherwise Windows may opt to only return large or small icons.
  // MSDN does not explicitly state this anywhere, but it has been observed on some machines.
  ExtractIconExA(path, 0, &info.large_icon, NULL, 1);
  ExtractIconExA(path, 0, NULL, &info.icon, 1);

  info.notification_icon = LoadImageA(NULL, path, IMAGE_ICON, GetSystemMetrics(SM_CXICON) * 2, GetSystemMetrics(SM_CYICON) * 2, LR_LOADFROMFILE);
  return info;
}

/**
 * @brief Initialize icon cache.
 * @param paths Paths to the icons.
 * @param count Number of paths.
 */
void _init_icon_cache(const char **paths, int count) {
  icon_info_count = count;
  icon_infos = malloc(sizeof(struct icon_info) * icon_info_count);

  for (int i = 0; i < count; ++i) {
    icon_infos[i] = _create_icon_info(paths[i]);
  }
}

/**
 * @brief Destroy icon cache.
 */
void _destroy_icon_cache() {
  for (int i = 0; i < icon_info_count; ++i) {
    if (icon_infos[i].icon) DestroyIcon(icon_infos[i].icon);
    if (icon_infos[i].large_icon) DestroyIcon(icon_infos[i].large_icon);
    if (icon_infos[i].notification_icon) DestroyIcon(icon_infos[i].notification_icon);
    free((void *) icon_infos[i].path);
  }

  free(icon_infos);
  icon_infos = NULL;
  icon_info_count = 0;
}

/**
 * @brief Fetch cached icon.
 * @param icon_record Icon record.
 * @param icon_type Icon type.
 * @return Icon.
 */
HICON _fetch_cached_icon(struct icon_info *icon_record, enum IconType icon_type) {
  switch (icon_type) {
    case REGULAR:
      return icon_record->icon;
    case LARGE:
      return icon_record->large_icon;
    case NOTIFICATION:
      return icon_record->notification_icon;
  }
}

/**
 * @brief Fetch icon.
 * @param path Path to the icon.
 * @param icon_type Icon type.
 * @return Icon.
 */
HICON _fetch_icon(const char *path, enum IconType icon_type) {
  // Find a cached icon by path
  for (int i = 0; i < icon_info_count; ++i) {
    if (strcmp(icon_infos[i].path, path) == 0) {
      return _fetch_cached_icon(&icon_infos[i], icon_type);
    }
  }

  // Expand cache, fetch, and retry
  icon_info_count += 1;
  icon_infos = realloc(icon_infos, sizeof(struct icon_info) * icon_info_count);
  icon_infos[icon_info_count - 1] = _create_icon_info(path);

  return _fetch_cached_icon(&icon_infos[icon_info_count - 1], icon_type);
}

int tray_init(struct tray *tray) {
  wm_taskbarcreated = RegisterWindowMessageA("TaskbarCreated");

  _init_icon_cache(tray->allIconPaths, tray->iconPathCount);
  g_tray = tray;

  memset(&wc, 0, sizeof(wc));
  wc.cbSize = sizeof(WNDCLASSEXA);
  wc.lpfnWndProc = _tray_wnd_proc;
  wc.hInstance = GetModuleHandle(NULL);
  wc.lpszClassName = WC_TRAY_CLASS_NAME;
  if (!RegisterClassExA(&wc)) {
    tray_log_last_error(TRAY_LOG_ERROR, "RegisterClassExA");
    _destroy_icon_cache();
    g_tray = NULL;
    return -1;
  }

  // Hidden top-level window (NOT message-only) is safest for Shell_NotifyIcon callbacks.
  hwnd = CreateWindowExA(0, WC_TRAY_CLASS_NAME, NULL, 0, 0, 0, 0, 0, NULL, NULL, GetModuleHandle(NULL), NULL);
  if (hwnd == NULL) {
    tray_log_last_error(TRAY_LOG_ERROR, "CreateWindowExA");
    _destroy_icon_cache();
    g_tray = NULL;
    UnregisterClassA(WC_TRAY_CLASS_NAME, GetModuleHandle(NULL));
    return -1;
  }
  UpdateWindow(hwnd);
  tray_allow_taskbar_created(hwnd);

  memset(&nid, 0, sizeof(nid));
  nid.cbSize = sizeof(NOTIFYICONDATAA);
  nid.hWnd = hwnd;
  nid.uID = 1; // non-zero id

  // A rejected NIM_ADD is not fatal: keep the window and message loop alive so
  // the retry timer and TaskbarCreated can register the icon once the shell is
  // willing. Tearing down here would leave the tray permanently missing.
  icon_added = FALSE;
  icon_add_failures = 0;
  tray_try_add_icon();
  return 0;
}

int tray_loop(int blocking) {
  MSG msg;
  if (blocking) {
    // Get thread-wide messages so we receive WM_QUIT too
    BOOL r = GetMessageA(&msg, NULL, 0, 0);
    if (r <= 0) {
      if (r == -1) {
        tray_log_last_error(TRAY_LOG_ERROR, "GetMessageA");
      }
      return -1; // error or WM_QUIT
    }
    TranslateMessage(&msg);
    DispatchMessageA(&msg);
    return 0;
  } else {
    // Drain all pending messages safely
    while (PeekMessageA(&msg, NULL, 0, 0, PM_REMOVE)) {
      if (msg.message == WM_QUIT) {
        return -1;
      }
      TranslateMessage(&msg);
      DispatchMessageA(&msg);
    }
    return 0;
  }
}

void tray_update(struct tray *tray) {
  tray_apply_state(tray, FALSE);
}

// Applies the given state to the shell icon. is_replay marks re-registration
// paths (TaskbarCreated, retry timer, NIM_MODIFY failure) that re-apply the
// remembered g_tray rather than a fresh update from the app.
static void tray_apply_state(struct tray *tray, BOOL is_replay) {
  if (tray == NULL || hwnd == NULL) {
    return;
  }

  g_tray = tray; // remember the last state for re-adding after Explorer restarts
  if (!icon_added) {
    // No icon registered yet; the retry path re-applies g_tray once NIM_ADD succeeds.
    return;
  }

  UINT id = ID_TRAY_FIRST;
  HMENU prevmenu = hmenu;
  hmenu = _tray_menu(tray->menu, &id);
  SendMessage(hwnd, WM_INITMENUPOPUP, (WPARAM) hmenu, 0);

  // Rebuild flags each update to avoid stale bits carrying over
  DWORD flags = tray_apply_icon_and_tip(tray, NIF_MESSAGE);

  // Balloon/toast (legacy surface mapped to Win10+ toasts)
  BOOL has_title = (tray->notification_title && tray->notification_title[0]);
  BOOL has_text  = (tray->notification_text  && tray->notification_text[0]);
  if (is_replay) {
    // Re-registration re-applies the remembered state, and NIF_INFO would
    // re-show the last toast. Taskbar hosts (DisplayFusion multi-monitor
    // taskbars, Explorer restarts) broadcast TaskbarCreated on every taskbar
    // rebuild, so an unexpired toast replayed each time reads as notification
    // spam. Only replay while the toast is still fresh.
    const BOOL fresh = notification_posted_ms != 0 &&
                       GetTickCount64() - notification_posted_ms <= TRAY_NOTIFICATION_REPLAY_TTL_MS;
    if (!fresh) {
      has_title = FALSE;
      has_text = FALSE;
    }
  } else {
    notification_posted_ms = (has_title || has_text) ? GetTickCount64() : 0;
  }
  if (has_title || has_text) {
    safe_copy_sz(nid.szInfoTitle, ARRAYSIZE(nid.szInfoTitle),
                 has_title ? tray->notification_title : "");
    safe_copy_sz(nid.szInfo, ARRAYSIZE(nid.szInfo),
                 has_text ? tray->notification_text : "");
    nid.dwInfoFlags = NIIF_NONE;

    // Prefer a user-provided notification icon; else fall back to the tray icon.
    HICON hLarge = NULL;
    if (tray->notification_icon && tray->notification_icon[0]) {
      hLarge = _fetch_icon(tray->notification_icon, NOTIFICATION);
    }
    if (!hLarge && nid.hIcon) {
      hLarge = nid.hIcon;
    }
#if defined(NIIF_LARGE_ICON)
    if (hLarge) {
      nid.hBalloonIcon = hLarge;
      nid.dwInfoFlags  = NIIF_USER | NIIF_LARGE_ICON;
    }
#endif
    flags |= NIF_INFO;
  } else {
    // Clear any previous info text to avoid the shell re-showing old balloons
    nid.szInfoTitle[0] = '\0';
    nid.szInfo[0]      = '\0';
    nid.dwInfoFlags    = NIIF_NONE;
    nid.hBalloonIcon   = NULL;
  }

  // Keep the callback up-to-date regardless of Focus Assist state
  notification_cb = tray->notification_cb;

  // Apply the freshly computed flags for this modification (prevents stale NIF_* carry-over)
  nid.uFlags = flags;
  if (!Shell_NotifyIconA(NIM_MODIFY, &nid)) {
    tray_log_last_error(TRAY_LOG_WARNING, "Shell_NotifyIconA(NIM_MODIFY)");
    // The shell no longer has our icon (e.g. Explorer restarted without us seeing
    // TaskbarCreated). Re-register it and re-apply this update.
    icon_added = FALSE;
    if (tray_add_notify_icon(tray, TRAY_LOG_WARNING) == 0) {
      icon_added = TRUE;
      nid.uFlags = flags;
      Shell_NotifyIconA(NIM_MODIFY, &nid);
    } else {
      ++icon_add_failures;
      tray_schedule_icon_retry();
    }
  }

  if (prevmenu != NULL) {
    DestroyMenu(prevmenu);
  }
}

void tray_exit(void) {
  g_tray = NULL;
  Shell_NotifyIconA(NIM_DELETE, &nid);
  _destroy_icon_cache();
  if (hwnd != NULL) {
    KillTimer(hwnd, ID_TRAY_RETRY_TIMER);
    DestroyWindow(hwnd);
    hwnd = NULL;
  }
  icon_added = FALSE;
  icon_add_failures = 0;
  notification_posted_ms = 0;
  if (hmenu != 0) {
    DestroyMenu(hmenu);
    hmenu = NULL;
  }
  notification_cb = NULL;
  memset(&nid, 0, sizeof(nid));
  UnregisterClassA(WC_TRAY_CLASS_NAME, GetModuleHandle(NULL));
}
