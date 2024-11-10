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
        
        # Bind window events when TrayManager is initialized
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        self.root.bind("<Unmap>", self._handle_minimize)
        
    def _handle_minimize(self, event):
        """Handle the minimize event"""
        # Check if window is being minimized and not already withdrawn
        if self.root.state() == 'iconic' and not self.tray_icon:
            self.minimize_to_tray()
        return "break"
        
    def _get_icon_path(self):
        """Get the correct path for the icon file"""
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
            menu = (
                item('Show', self._show_window),
                item('Exit', self._exit_application)
            )
            
            self.tray_icon = pystray.Icon(
                "PQManager",
                icon_image,
                "Printer Queue Manager",
                menu
            )
            self.tray_icon.run_detached()
            logging.info("Tray icon created successfully")
            
        except Exception as e:
            logging.error(f"Error creating tray icon: {e}")
            raise

    def _show_window(self, icon, item):
        """Show the main window and remove tray icon"""
        if self.tray_icon:
            self.tray_icon.visible = False
            self.tray_icon.stop()
            self.tray_icon = None
        self.root.after(0, self._restore_window)

    def _restore_window(self):
        """Restore the main window"""
        self.root.deiconify()
        self.root.state('normal')  # Ensure window is in normal state
        self.root.lift()
        self.root.focus_force()

    def minimize_to_tray(self, event=None):
        """Minimize the window to system tray"""
        if not self.tray_icon:
            self.create_tray_icon()
        self.root.withdraw()
        return "break"

    def _exit_application(self, icon, item):
        """Clean exit of the application"""
        self.cleanup()
        self.root.quit()

    def cleanup(self):
        """Clean up resources"""
        if self.tray_icon:
            try:
                self.tray_icon.visible = False
                self.tray_icon.stop()
                self.tray_icon = None
            except Exception as e:
                logging.error(f"Error during cleanup: {e}")
