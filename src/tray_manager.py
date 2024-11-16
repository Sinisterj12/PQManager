"""
!!! IMPORTANT - DO NOT MODIFY WITHOUT CAREFUL CONSIDERATION !!!
This module handles all system tray functionality and window state management.
Any changes to the minimize/restore logic could break the application's tray behavior.

Last verified working version: 1.0.0
Tested behaviors:
- X button minimizes to tray
- Minimize button (dash) minimizes to tray
- "Minimize to Tray" button works
- Restore from tray works
- Clean exit from tray works
"""

import pystray
from pystray import MenuItem as item
from PIL import Image
import logging
import sys
import os
from packaging import version

__version__ = "1.0.0"  # Current working version
MINIMUM_COMPATIBLE_VERSION = "1.0.0"  # Minimum version known to work

class TrayManager:
    def __init__(self, root_window):
        current = version.parse(__version__)
        min_compatible = version.parse(MINIMUM_COMPATIBLE_VERSION)
        
        if current < min_compatible:
            raise RuntimeError(
                f"TrayManager version {__version__} is not compatible. "
                f"Minimum required version is {MINIMUM_COMPATIBLE_VERSION}"
            )
            
        self.root = root_window
        self.tray_icon = None
        self.icon_path = self._get_icon_path()
        
        # Bind window events
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        self.root.bind("<Unmap>", lambda e: self.minimize_to_tray() if e.widget is self.root else None)
        
    def _get_icon_path(self):
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, "assets", "Logo.ico")

    def create_tray_icon(self):
        """Create and display the system tray icon"""
        if self.tray_icon:
            return
        
        try:
            icon_image = Image.open(self.icon_path)
            
            def left_click(icon, query):
                self._show_window(icon, None)
            
            menu = (
                item('Show', self._show_window, default=True),  # Make Show the default action
                item('Exit', self._exit_application)
            )
            
            self.tray_icon = pystray.Icon(
                "PQManager",
                icon_image,
                "PQ Manager v1.1",
                menu
            )
            
            # Set default action for left click
            self.tray_icon.on_click = left_click
            
            self.tray_icon.run_detached()
            logging.info("Tray icon created successfully")
            
        except Exception as e:
            logging.error(f"Error creating tray icon: {e}")
            raise

    def _show_window(self, icon, item):
        """Show the main window and remove tray icon"""
        try:
            if self.tray_icon:
                self.tray_icon.visible = False
                self.tray_icon.stop()
                self.tray_icon = None
            
            self.root.after(0, self._restore_window)
        except Exception as e:
            logging.error(f"Error showing window: {e}")

    def _restore_window(self):
        """Restore the main window and bring it to front"""
        try:
            self.root.deiconify()
            self.root.state('normal')
            self.root.lift()
            self.root.focus_force()
            
            # Force window to top
            self.root.attributes('-topmost', True)
            self.root.update()
            self.root.attributes('-topmost', False)
        except Exception as e:
            logging.error(f"Error restoring window: {e}")

    def minimize_to_tray(self, event=None):
        """Minimize the window to system tray"""
        try:
            if not self.tray_icon:
                self.create_tray_icon()
            self.root.withdraw()
            logging.info("Window minimized to tray successfully")
        except Exception as e:
            logging.error(f"Error minimizing to tray: {e}")

    def _exit_application(self, icon, item):
        """Clean exit from the application"""
        try:
            # Stop the tray icon first
            if self.tray_icon:
                self.tray_icon.visible = False
                self.tray_icon.stop()
                self.tray_icon = None
            
            # Schedule the actual exit
            self.root.after(0, self._perform_exit)
        except Exception as e:
            logging.error(f"Error during exit: {e}")
            # Force exit if needed
            self.root.quit()

    def _perform_exit(self):
        """Perform the actual exit operations"""
        try:
            # Destroy the root window
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            logging.error(f"Error during window destruction: {e}")
            # Force exit if needed
            import sys
            sys.exit(0)

    def cleanup(self):
        """Clean up resources"""
        if self.tray_icon:
            try:
                self.tray_icon.visible = False
                self.tray_icon.stop()
                self.tray_icon = None
            except Exception as e:
                logging.error(f"Error during cleanup: {e}")
