import tkinter as tk
from tkinter import ttk
import win32print
import logging
import logging.handlers
import json
import os
import sys
import tempfile
import atexit
import win32event
import win32api
import winerror
import win32con
import win32gui
from src.tray_manager import TrayManager
from src.config_manager import ConfigManager
from src.printer_manager import PrinterManager
import time

def is_already_running():
    """Check if another instance is already running"""
    try:
        handle = win32event.CreateMutex(None, True, "PQManager")
        return win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS
    except:
        return False

class Application:
    def __init__(self):
        # Create necessary directories
        os.makedirs('logs', exist_ok=True)

        # Remove any existing handlers from the root logger
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)

        # Add log rotation handler
        log_file = 'logs/printer_queue.log'
        handler = logging.handlers.RotatingFileHandler(
            log_file,
            maxBytes=1024*1024,  # 1MB per file
            backupCount=5  # Keep 5 backup files
        )
        
        # Create a custom formatter with colors
        class ColoredFormatter(logging.Formatter):
            COLORS = {
                'ERROR': '\033[91m',  # Red
                'INFO': '\033[92m',   # Green
                'WARNING': '\033[93m', # Yellow
                'DEBUG': '\033[94m',   # Blue
                'RESET': '\033[0m'     # Reset
            }

            def format(self, record):
                # Check if it's a new day compared to the last log
                now = time.strftime('%Y-%m-%d')
                if not hasattr(self, 'last_date'):
                    self.last_date = now
                elif self.last_date != now:
                    self.last_date = now
                    # Add a separator line between days
                    with open(log_file, 'a') as f:
                        f.write('\n' + '='*50 + f'\n{now}\n' + '='*50 + '\n')

                # Add colors to the level name
                if record.levelname in self.COLORS:
                    record.levelname = f"{self.COLORS[record.levelname]}{record.levelname}{self.COLORS['RESET']}"
                
                return super().format(record)

        # Create and set the formatter
        formatter = ColoredFormatter(
            '%(asctime)s [%(levelname)s] %(message)s',
            datefmt='%H:%M:%S'  # Shorter time format, date is shown in separators
        )
        handler.setFormatter(formatter)

        # Configure the root logger
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(handler)
        
        # Log startup with version
        logging.info("PQManager v1.1 starting up")

        # Initialize the main window
        self.root = tk.Tk()
        self.root.title("PQ Manager v1.1")
        self.root.geometry("300x200")
        self.root.resizable(False, False)
        
        # Disable error popups
        def custom_excepthook(exc_type, exc_value, exc_traceback):
            logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
        sys.excepthook = custom_excepthook
        
        # Set up window close protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Initialize managers
        self.config_manager = ConfigManager()
        self.printer_manager = PrinterManager(self.config_manager)
        self.tray_manager = TrayManager(self.root)

        # Global variables
        self.printer_var = tk.StringVar()
        self.is_monitoring = False
        self.monitor_id = None

        # Create GUI
        self.create_gui()
        
        # Load saved settings and start monitoring
        self.load_saved_settings()
        
        # Hide window if started minimized
        if len(sys.argv) > 1 and '--minimized' in sys.argv:
            logging.info("Starting minimized")
            self.root.withdraw()
            self.tray_manager.create_tray_icon()

    def on_closing(self):
        """Handle window closing"""
        try:
            # Just minimize to tray instead of closing
            self.tray_manager.minimize_to_tray()
        except Exception as e:
            logging.error(f"Error during window closing: {e}")

    def monitor_queue(self):
        """Monitor the print queue for the selected printer"""
        try:
            if self.is_monitoring and self.printer_var.get() and self.printer_var.get() != "Select Printer":
                printer_name = self.printer_var.get()
                # Only clear queue if there are 2 or more jobs
                queue_length = self.printer_manager.get_queue_length(printer_name)
                if queue_length > 1:  
                    if self.printer_manager.clear_queue(printer_name):
                        pass
            
            # Schedule next check every 5 minutes
            if self.is_monitoring:
                self.monitor_id = self.root.after(300000, self.monitor_queue)  
                
        except Exception as e:
            logging.error(f"Error in monitor_queue: {e}")
            # Ensure monitoring continues even after error
            if self.is_monitoring:
                self.monitor_id = self.root.after(300000, self.monitor_queue)

    def start_monitoring(self):
        """Start monitoring with status check"""
        if not self.is_monitoring:
            selected_printer = self.printer_var.get()
            if selected_printer and selected_printer != "Select Printer":
                logging.info(f"Starting monitoring for {selected_printer}")
                self.is_monitoring = True
                self.monitor_queue()  # Start immediately
                # Update settings individually to preserve other settings
                self.config_manager.update_setting('selected_printer', selected_printer)
                self.config_manager.update_setting('auto_start_monitoring', True)

    def stop_monitoring(self):
        """Stop monitoring"""
        if self.is_monitoring:
            self.is_monitoring = False
            if self.monitor_id:
                self.root.after_cancel(self.monitor_id)
                self.monitor_id = None
            logging.info("Monitoring stopped")

    def on_printer_select(self, event=None):
        """Handle printer selection"""
        selected_printer = self.printer_var.get()
        if selected_printer and selected_printer != "Select Printer":
            # Stop any existing monitoring
            self.stop_monitoring()
            
            # Update settings individually to preserve other settings
            self.config_manager.update_setting('selected_printer', selected_printer)
            self.config_manager.update_setting('auto_start_monitoring', True)
            
            # Start new monitoring
            self.start_monitoring()
            
    def create_gui(self):
        """Create the GUI"""
        frame = ttk.Frame(self.root)
        frame.pack(expand=True, fill='both', padx=10, pady=10)

        # Printer selection
        printer_label = ttk.Label(frame, text="Select Receipt Printer:")
        printer_label.pack(pady=(0, 5))

        printer_dropdown = ttk.Combobox(frame, textvariable=self.printer_var, width=30)
        printer_dropdown['values'] = self.printer_manager.get_printers()
        printer_dropdown.pack(pady=(0, 10))
        printer_dropdown.set("Select Printer")
        printer_dropdown.bind('<<ComboboxSelected>>', self.on_printer_select)

        # Queue management
        def clear_queue():
            selected_printer = self.printer_var.get()
            if selected_printer and selected_printer != "Select Printer":
                logging.info(f"Manually clearing queue for {selected_printer}")
                if self.printer_manager.clear_queue(selected_printer):
                    logging.info("Queue cleared successfully")
                else:
                    logging.warning("Failed to clear queue completely")

        clear_button = ttk.Button(
            frame, 
            text="Clear Print Queue", 
            command=clear_queue
        )
        clear_button.pack(pady=5)

        # Bottom attribution
        author_label = ttk.Label(frame, text="by James", font=('Segoe UI', 8, 'italic'))
        author_label.pack(side='right', padx=5, pady=(10, 0))

    def load_saved_settings(self):
        """Load saved printer and monitoring status"""
        try:
            # Get available printers
            printers = self.printer_manager.get_printers()
            
            # Load saved printer from config
            saved_printer = self.config_manager.get_setting('selected_printer')
            if saved_printer and saved_printer in printers:
                self.printer_var.set(saved_printer)
                # Only log if we're switching from default
                if saved_printer != "Select Printer":
                    logging.info(f"Loaded saved printer: {saved_printer}")
            
            # Start monitoring if auto-start is enabled
            if self.config_manager.get_setting('auto_start_monitoring', False):
                self.start_monitoring()
        except Exception as e:
            logging.error(f"Error loading settings: {e}")

    def run(self):
        """Start the application"""
        # Start normally - let the window be visible until user minimizes it
        self.root.mainloop()

    def cleanup(self):
        """Clean up resources before exit"""
        try:
            self.stop_monitoring()
            if self.monitor_id:
                self.root.after_cancel(self.monitor_id)
                self.monitor_id = None
        except Exception as e:
            logging.error(f"Error during cleanup: {e}")

    def show_window(self):
        """Show the main window"""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def hide_window(self):
        """Hide the main window"""
        self.root.withdraw()

def main():
    # Setup logging first
    os.makedirs('logs', exist_ok=True)
    logging.basicConfig(
        filename='logs/printer_queue.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Check for existing instance
    if is_already_running():
        return
        
    try:
        app = Application()
        if len(sys.argv) > 1 and '--minimized' in sys.argv:
            app.root.withdraw()
            app.tray_manager.create_tray_icon()
        app.run()
    except Exception as e:
        logging.error(f"Error in main: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
