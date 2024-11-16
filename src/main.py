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
            level=logging.INFO,
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
        self.monitor_id = None

        # Create GUI
        self.create_gui()
        
        # Load saved settings and start monitoring
        self.load_saved_settings()

    def on_closing(self):
        """Handle window closing"""
        try:
            # Stop monitoring
            self.stop_monitoring()
            
            # Minimize to tray instead of closing
            self.tray_manager.minimize_to_tray()
        except Exception as e:
            logging.error(f"Error during window closing: {e}")

    def monitor_queue(self):
        """Monitor printer queue with enhanced error handling"""
        if not self.is_monitoring:
            return

        try:
            selected_printer = self.printer_var.get()
            if selected_printer and selected_printer != "Select Printer":
                logging.info(f"Checking queue for {selected_printer}")
                
                # Get queue length
                queue_length = self.printer_manager.get_queue_length(selected_printer)
                logging.info(f"Current queue length: {queue_length}")
                
                if queue_length > 0:
                    logging.info(f"Attempting to clear {queue_length} jobs...")
                    # Try clearing up to 3 times if needed
                    for attempt in range(3):
                        if self.printer_manager.clear_queue(selected_printer):
                            logging.info("Queue cleared successfully")
                            break
                        else:
                            # Check if queue is actually empty despite error
                            new_length = self.printer_manager.get_queue_length(selected_printer)
                            if new_length == 0:
                                logging.info("Queue is empty despite reported error")
                                break
                            elif attempt < 2:  # Don't log on last attempt
                                logging.warning(f"Failed to clear queue (attempt {attempt + 1}), retrying...")
                                time.sleep(1)  # Small delay between attempts
            else:
                logging.info("No printer selected")
                
        except Exception as e:
            logging.error(f"Error in monitor_queue: {e}")

        # Schedule next check if still monitoring
        if self.is_monitoring:
            self.monitor_id = self.root.after(10000, self.monitor_queue)

    def start_monitoring(self):
        """Start monitoring with status check"""
        if not self.is_monitoring:
            selected_printer = self.printer_var.get()
            if selected_printer and selected_printer != "Select Printer":
                logging.info(f"Starting monitoring for {selected_printer}")
                self.is_monitoring = True
                # Start immediate check
                self.monitor_queue()

    def stop_monitoring(self):
        """Stop monitoring"""
        if self.is_monitoring:
            self.is_monitoring = False
            # Cancel any pending monitor calls
            if self.monitor_id:
                self.root.after_cancel(self.monitor_id)
                self.monitor_id = None
            logging.info("Monitoring stopped")

    def on_printer_select(self, event=None):
        """Handle printer selection with enhanced error handling"""
        selected_printer = self.printer_var.get()
        if selected_printer and selected_printer != "Select Printer":
            logging.info(f"Printer selected: {selected_printer}")
            
            # Stop any existing monitoring
            self.stop_monitoring()
            
            # Save settings
            settings = self.config_manager.load_settings()
            settings['selected_printer'] = selected_printer
            settings['auto_start_monitoring'] = True
            self.config_manager.save_settings(settings)
            
            # Start new monitoring
            self.start_monitoring()
            
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
        """Load saved settings and initialize printer"""
        try:
            settings = self.config_manager.load_settings()
            saved_printer = settings.get('selected_printer')
            
            if saved_printer:
                available_printers = self.printer_manager.get_printers()
                if saved_printer in available_printers:
                    self.printer_var.set(saved_printer)
                    logging.info(f"Loaded saved printer: {saved_printer}")
                    if settings.get('auto_start_monitoring', True):
                        self.start_monitoring()
        except Exception as e:
            logging.error(f"Error loading settings: {e}")

    def run(self):
        """Start the application"""
        self.root.mainloop()

def main():
    app = Application()
    app.run()

if __name__ == "__main__":
    main()
