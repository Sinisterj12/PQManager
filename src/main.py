import tkinter as tk
from tkinter import ttk
import win32print
import logging
import json
import os
from src.tray_manager import TrayManager
from src.config_manager import ConfigManager
from src.printer_manager import PrinterManager
import time

class Application:
    def __init__(self):
        # Create necessary directories
        os.makedirs('logs', exist_ok=True)

        # Setup logging
        logging.basicConfig(
            filename='logs/printer_queue.log',
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

        # Initialize the main window
        self.root = tk.Tk()
        self.root.title("PQ Manager v1.1")
        self.root.geometry("300x200")
        self.root.resizable(False, False)

        # Initialize managers
        self.config_manager = ConfigManager()
        self.printer_manager = PrinterManager(self.config_manager)
        self.tray_manager = TrayManager(self.root)

        # Global variables
        self.printer_var = tk.StringVar()
        self.is_monitoring = False
        self.last_queue_check = 0
        self.queue_check_interval = 15  # seconds

        # Create GUI
        self.create_gui()
        
        # Load saved settings and start monitoring
        self.load_saved_settings()

    def monitor_queue(self):
        """Monitor printer queue with enhanced error handling"""
        if self.is_monitoring and self.root.winfo_exists():
            current_time = time.time()
            
            if current_time - self.last_queue_check >= self.queue_check_interval:
                selected_printer = self.printer_var.get()
                if selected_printer and selected_printer != "Select Printer":
                    # Check printer status first
                    if self.printer_manager.check_printer_status(selected_printer):
                        # Only clear queue if printer is responsive
                        queue_length = self.printer_manager.get_queue_length(selected_printer)
                        if queue_length > 0:
                            self.printer_manager.clear_queue(selected_printer)
                            logging.info(f"Cleared {queue_length} jobs from {selected_printer}")
                    self.last_queue_check = current_time
            
            # Schedule next check
            self.root.after(1000, self.monitor_queue)

    def start_monitoring(self):
        """Start monitoring with status check"""
        if not self.is_monitoring:
            self.is_monitoring = True
            self.monitor_queue()
            logging.info("Started printer monitoring")

    def on_printer_select(self, event=None):
        """Handle printer selection with enhanced error handling"""
        selected_printer = self.printer_var.get()
        if selected_printer and selected_printer != "Select Printer":
            if self.printer_manager.check_printer_status(selected_printer):
                self.config_manager.update_setting('selected_printer', selected_printer)
                self.config_manager.update_setting('auto_start_monitoring', True)
                self.start_monitoring()
                logging.info(f"Selected and verified printer: {selected_printer}")
            else:
                logging.warning(f"Selected printer {selected_printer} is not ready")

    def create_gui(self):
        """Create the GUI with improved layout"""
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
        clear_button = ttk.Button(
            frame, 
            text="Clear Print Queue", 
            command=lambda: self.printer_manager.clear_queue(self.printer_var.get())
        )
        clear_button.pack(pady=5)

        # Bottom attribution
        author_label = ttk.Label(frame, text="by James", font=('Segoe UI', 8, 'italic'))
        author_label.pack(side='right', padx=5, pady=(10, 0))

    def load_saved_settings(self):
        """Load saved settings and initialize printer"""
        settings = self.config_manager.load_settings()
        saved_printer = settings.get('selected_printer')
        
        if saved_printer and saved_printer in self.printer_manager.get_printers():
            self.printer_var.set(saved_printer)
            if settings.get('auto_start_monitoring', True):
                self.start_monitoring()
        
        # Start minimized if configured
        if settings.get('start_minimized', True):  # Default to True for stores
            self.root.after(1000, self.tray_manager.minimize_to_tray)

    def run(self):
        """Start the application"""
        self.root.mainloop()

def main():
    app = Application()
    app.run()

if __name__ == "__main__":
    main()
