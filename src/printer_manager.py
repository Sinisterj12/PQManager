import win32print
import win32con
import win32api
import time
import logging
from datetime import datetime
from win32com.shell import shell, shellcon
import os
import winreg
import json
import sys

class PrinterManager:
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.last_error_time = {}  # Track last error time per printer
        self.error_cooldown = 300  # 5 minutes between same error notifications
        self.setup_autostart()
        
    def setup_autostart(self):
        """Setup application to run at startup using Registry"""
        try:
            # Get the path to the executable
            if getattr(sys, 'frozen', False):
                app_path = sys.executable
            else:
                app_path = os.path.abspath(sys.argv[0])
                
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE
            )
            
            try:
                existing_path = winreg.QueryValueEx(key, "PQManager")[0]
                if existing_path != f'"{app_path}"':
                    winreg.SetValueEx(key, "PQManager", 0, winreg.REG_SZ, f'"{app_path}"')
            except WindowsError:
                winreg.SetValueEx(key, "PQManager", 0, winreg.REG_SZ, f'"{app_path}"')
                
            winreg.CloseKey(key)
            logging.info("Autostart registry entry verified")
        except Exception as e:
            logging.error(f"Failed to setup autostart: {e}")

    def clear_queue(self, printer_name):
        """Clear the print queue with enhanced error handling"""
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            try:
                jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                if jobs:
                    for job in jobs:
                        try:
                            win32print.SetJob(printer_handle, job['JobId'], 0, None, win32print.JOB_CONTROL_DELETE)
                            logging.info(f"Cleared job {job['JobId']} from {printer_name}")
                        except Exception as e:
                            self._handle_error(printer_name, f"Failed to clear job {job['JobId']}: {e}")
                    
                    # Verify queue is clear
                    remaining_jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                    if remaining_jobs:
                        self._handle_error(printer_name, "Some jobs could not be cleared")
                        
            finally:
                win32print.ClosePrinter(printer_handle)
                
        except Exception as e:
            self._handle_error(printer_name, f"Printer access error: {e}")

    def _handle_error(self, printer_name, error_msg):
        """Handle printer errors with notifications"""
        current_time = time.time()
        error_key = f"{printer_name}:{error_msg}"
        
        # Check if we should show this error (cooldown period)
        if error_key not in self.last_error_time or \
           (current_time - self.last_error_time[error_key]) > self.error_cooldown:
            
            self.last_error_time[error_key] = current_time
            logging.error(f"{printer_name}: {error_msg}")
            
            # Show balloon tip notification
            try:
                from win32gui import Shell_NotifyIcon, NIM_MODIFY, NIIF_WARNING
                from win32gui import NIF_INFO, NIF_ICON, NIF_TIP, NIF_MESSAGE
                # Note: actual notification implementation will be handled by tray_manager
            except Exception as e:
                logging.error(f"Failed to show notification: {e}")

    def check_printer_status(self, printer_name):
        """Check printer status and attempt reconnection if needed"""
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(printer_handle, 2)
            win32print.ClosePrinter(printer_handle)
            
            status = printer_info['Status']
            if status == 0:
                return True  # Printer is ready
                
            # Log specific printer state for troubleshooting
            status_messages = []
            if status & win32print.PRINTER_STATUS_PAPER_JAM:
                status_messages.append("Paper jam detected")
            if status & win32print.PRINTER_STATUS_PAPER_OUT:
                status_messages.append("Out of paper")
            if status & win32print.PRINTER_STATUS_PAPER_PROBLEM:
                status_messages.append("Paper problem")
            if status & win32print.PRINTER_STATUS_OFFLINE:
                status_messages.append("Printer offline")
            
            if status_messages:
                self._handle_error(printer_name, ", ".join(status_messages))
            return False
            
        except Exception as e:
            self._handle_error(printer_name, f"Status check failed: {e}")
            return False

    def get_queue_length(self, printer_name):
        """Get number of jobs in print queue"""
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
            win32print.ClosePrinter(printer_handle)
            return len(jobs)
        except Exception as e:
            self._handle_error(printer_name, f"Failed to get queue length: {e}")
            return 0

    def load_saved_printer(self):
        """Load the saved printer from the configuration file."""
        try:
            if os.path.exists('printer_config.json'):
                with open('printer_config.json', 'r') as f:
                    config = json.load(f)
                    printer_name = config.get('printer')
                    logging.info(f"Loaded saved printer: {printer_name}")
                    return printer_name
        except Exception as e:
            logging.error(f"Error loading saved printer: {e}")

    def get_printers(self):
        """Get list of available printers."""
        try:
            return [printer[2] for printer in 
                   win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        except Exception as e:
            logging.error(f"Error getting printers: {e}")
            return []

def main():
    config_manager = None  # Replace with actual config manager
    printer_manager = PrinterManager(config_manager)
    saved_printer = printer_manager.load_saved_printer()
    if saved_printer:
        printer_manager.clear_queue(saved_printer)
    else:
        available_printers = printer_manager.get_printers()
        if available_printers:
            printer_manager.clear_queue(available_printers[0])

if __name__ == "__main__":
    main()
