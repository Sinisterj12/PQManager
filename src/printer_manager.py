import win32print
import logging
import json
import os
from typing import List

class PrinterQueueManager:
    """Manages printer queue operations."""
    
    def __init__(self):
        self.selected_printer = None
        
    def get_printers(self) -> List[str]:
        """Get list of available printers."""
        try:
            return [printer[2] for printer in 
                   win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        except Exception as e:
            logging.error(f"Error getting printers: {e}")
            return []

    def clear_queue(self) -> None:
        """Clear the selected printer's queue."""
        if not self.selected_printer:
            logging.warning("No printer selected.")
            return

        try:
            printer_handle = win32print.OpenPrinter(self.selected_printer)
            jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
            for job in jobs:
                win32print.SetJob(
                    printer_handle, 
                    job['JobId'], 
                    0, 
                    None, 
                    win32print.JOB_CONTROL_DELETE
                )
            win32print.ClosePrinter(printer_handle)
            logging.info(f"Queue cleared for {self.selected_printer}")
        except Exception as e:
            logging.error(f"Error clearing queue: {e}")

    def load_saved_printer(self) -> None:
        """Load the saved printer from the configuration file."""
        try:
            if os.path.exists('printer_config.json'):
                with open('printer_config.json', 'r') as f:
                    config = json.load(f)
                    self.selected_printer = config.get('printer')
                    logging.info(f"Loaded saved printer: {self.selected_printer}")
        except Exception as e:
            logging.error(f"Error loading saved printer: {e}")
