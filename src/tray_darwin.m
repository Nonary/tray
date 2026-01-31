/**
 * @file src/tray_darwin.m
 * @brief System tray implementation for macOS.
 */
// standard includes
#include <stdarg.h>
#include <stdio.h>
#include <string.h>

// lib includes
#include <Cocoa/Cocoa.h>

// local includes
#include "tray.h"

/**
 * @class AppDelegate
 * @brief The application delegate that handles menu actions.
 */
@interface AppDelegate: NSObject <NSApplicationDelegate>
/**
 * @brief Callback function for menu item actions.
 * @param sender The object that sent the action message.
 * @return void
 */
- (IBAction)menuCallback:(id)sender;
@end

@implementation AppDelegate {
}

- (IBAction)menuCallback:(id)sender {
  struct tray_menu *m = [[sender representedObject] pointerValue];
  if (m != NULL && m->cb != NULL) {
    m->cb(m);
  }
}

@end

static NSApplication *app;
static NSStatusBar *statusBar;
static NSStatusItem *statusItem;

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

static NSMenu *_tray_menu(struct tray_menu *m) {
  NSMenu *menu = [[NSMenu alloc] init];
  [menu setAutoenablesItems:FALSE];

  for (; m != NULL && m->text != NULL; m++) {
    if (strcmp(m->text, "-") == 0) {
      [menu addItem:[NSMenuItem separatorItem]];
    } else {
      NSMenuItem *menuItem = [[NSMenuItem alloc]
        initWithTitle:[NSString stringWithUTF8String:m->text]
               action:@selector(menuCallback:)
        keyEquivalent:@""];
      [menuItem setEnabled:(m->disabled ? FALSE : TRUE)];
      [menuItem setState:(m->checked ? 1 : 0)];
      [menuItem setRepresentedObject:[NSValue valueWithPointer:m]];
      [menu addItem:menuItem];
      if (m->submenu != NULL) {
        [menu setSubmenu:_tray_menu(m->submenu) forItem:menuItem];
      }
    }
  }
  return menu;
}

int tray_init(struct tray *tray) {
  AppDelegate *delegate = [[AppDelegate alloc] init];
  app = [NSApplication sharedApplication];
  [app setDelegate:delegate];
  statusBar = [NSStatusBar systemStatusBar];
  statusItem = [statusBar statusItemWithLength:NSVariableStatusItemLength];
  if (statusBar == nil || statusItem == nil) {
    tray_log(TRAY_LOG_ERROR, "Failed to initialize NSStatusBar/NSStatusItem");
    return -1;
  }
  tray_update(tray);
  [app activateIgnoringOtherApps:TRUE];
  return 0;
}

int tray_loop(int blocking) {
  NSDate *until = (blocking ? [NSDate distantFuture] : [NSDate distantPast]);
  NSEvent *event = [app nextEventMatchingMask:ULONG_MAX
                                    untilDate:until
                                       inMode:[NSString stringWithUTF8String:"kCFRunLoopDefaultMode"]
                                      dequeue:TRUE];
  if (event) {
    [app sendEvent:event];
  }
  return 0;
}

void tray_update(struct tray *tray) {
  NSImage *image = [[NSImage alloc] initWithContentsOfFile:[NSString stringWithUTF8String:tray->icon]];
  NSSize size = NSMakeSize(16, 16);
  if (image == nil) {
    tray_log(TRAY_LOG_WARNING, "Failed to load tray icon image");
    return;
  }
  [image setSize:NSMakeSize(16, 16)];
  statusItem.button.image = image;
  [statusItem setMenu:_tray_menu(tray->menu)];
}

void tray_exit(void) {
  [app terminate:app];
}
