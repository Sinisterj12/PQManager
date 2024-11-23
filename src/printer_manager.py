import win32print
import win32api
import win32con
import logging
import time
import ctypes
import sys

# Import win32timezone conditionally
try:
    import win32timezone
except ImportError:
    logging.warning("win32timezone not available, some functionality may be limited")

class PrinterManager:
    def __init__(self, config_manager):
        self.config_manager = config_manager
        # Only log if not admin
        if not self.is_admin():
            logging.warning("Application is not running with administrator rights")
        self._verify_win32_modules(silent=True)

    def _verify_win32_modules(self, silent=False):
        """Verify all required win32 modules are available"""
        required_modules = ['win32print', 'win32api', 'win32con']
        missing_modules = []
        for module in required_modules:
            try:
                __import__(module)
                if not silent:
                    logging.info(f"Successfully loaded {module}")
            except ImportError as e:
                missing_modules.append(module)
                logging.error(f"Failed to load {module}: {e}")
        
        if missing_modules:
            logging.error(f"Missing required modules: {', '.join(missing_modules)}")

    def get_queue_length(self, printer_name):
        """Get number of jobs in print queue with better error handling"""
        if not printer_name or printer_name == "Select Printer":
            return 0

        try:
            handle = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
            printer_handle = None
            try:
                printer_handle = win32print.OpenPrinter(printer_name, handle)
                if not printer_handle:
                    logging.error(f"Failed to get printer handle for {printer_name}")
                    return 0
                
                jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                count = len(jobs)
                # Only log if there are multiple jobs in the queue
                if count > 1:
                    logging.info(f"Found {count} jobs in queue for {printer_name}")
                return count
                
            finally:
                if printer_handle:
                    win32print.ClosePrinter(printer_handle)
                    
        except Exception as e:
            logging.error(f"Failed to get queue length: {e}")
            if "win32timezone" in str(e):
                return self._get_queue_length_basic(printer_name)
            return 0

    def _get_queue_length_basic(self, printer_name):
        """Fallback method for getting queue length without win32timezone"""
        try:
            handle = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
            printer_handle = win32print.OpenPrinter(printer_name, handle)
            try:
                jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                return len(jobs)
            finally:
                win32print.ClosePrinter(printer_handle)
        except Exception as e:
            logging.error(f"Basic queue length check failed: {e}")
            return 0

    def _get_job_status_string(self, status):
        """Convert job status to readable string"""
        status_flags = []
        if status & win32print.JOB_STATUS_PAUSED:
            status_flags.append("Paused")
        if status & win32print.JOB_STATUS_ERROR:
            status_flags.append("Error")
        if status & win32print.JOB_STATUS_DELETING:
            status_flags.append("Deleting")
        if status & win32print.JOB_STATUS_SPOOLING:
            status_flags.append("Spooling")
        if status & win32print.JOB_STATUS_PRINTING:
            status_flags.append("Printing")
        if status & win32print.JOB_STATUS_OFFLINE:
            status_flags.append("Offline")
        if status & win32print.JOB_STATUS_PAPEROUT:
            status_flags.append("Paper Out")
        if status & win32print.JOB_STATUS_PRINTED:
            status_flags.append("Printed")
        if status & win32print.JOB_STATUS_DELETED:
            status_flags.append("Deleted")
        if status & win32print.JOB_STATUS_BLOCKED_DEVQ:
            status_flags.append("Blocked")
        if status & win32print.JOB_STATUS_USER_INTERVENTION:
            status_flags.append("Needs User Intervention")
        return ", ".join(status_flags) if status_flags else "Unknown"

    def clear_queue(self, printer_name):
        """Clear the print queue with enhanced error handling"""
        if not printer_name or printer_name == "Select Printer":
            return False

        try:
            handle = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
            printer_handle = win32print.OpenPrinter(printer_name, handle)
            
            try:
                # Get printer info to check if it's a receipt printer
                printer_info = win32print.GetPrinter(printer_handle, 2)
                is_receipt_printer = any(keyword in printer_info['pDriverName'].lower() 
                                      for keyword in ['receipt', 'pos', 'thermal', 'epson', 'star'])
                
                jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                if not jobs:
                    return True

                # Only log when we find multiple jobs
                job_count = len(jobs)
                if job_count > 1:
                    logging.info(f"Found and clearing {job_count} jobs from {printer_name}")
                
                success = True
                
                for job in jobs:
                    try:
                        job_id = job['JobId']
                        
                        # For receipt printers, be more aggressive with completed jobs
                        if is_receipt_printer and (job['Status'] & win32print.JOB_STATUS_PRINTED):
                            try:
                                win32print.SetJob(printer_handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
                                continue
                            except Exception as e:
                                logging.error(f"Failed to clear completed receipt job {job_id}: {e}")
                        
                        # Try to cancel first
                        try:
                            win32print.SetJob(printer_handle, job_id, 0, None, win32print.JOB_CONTROL_CANCEL)
                            if is_receipt_printer:
                                time.sleep(0.2)
                        except Exception as e:
                            if "parameter is incorrect" not in str(e).lower():
                                logging.error(f"Failed to cancel job {job_id}: {e}")
                        
                        # Then try to delete
                        try:
                            win32print.SetJob(printer_handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
                            if is_receipt_printer:
                                time.sleep(0.2)
                        except Exception as e:
                            if "parameter is incorrect" not in str(e).lower():
                                logging.error(f"Failed to delete job {job_id}: {e}")
                                success = False
                        
                    except Exception as e:
                        if "parameter is incorrect" not in str(e).lower():
                            success = False
                            logging.error(f"Failed to process job {job_id}: {e}")

                # Verify queue is actually empty
                remaining_jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                if remaining_jobs:
                    if len(remaining_jobs) > 1:  # Only log if multiple jobs remained
                        logging.error(f"{len(remaining_jobs)} jobs could not be cleared from {printer_name}")
                    success = False
                elif job_count > 1:  # Only log success if we cleared multiple jobs
                    logging.info(f"Successfully cleared {job_count} jobs from {printer_name}")
                
                return success
                
            finally:
                win32print.ClosePrinter(printer_handle)
                
        except Exception as e:
            logging.error(f"Failed to clear print queue: {e}")
            return False

    def check_queue(self, printer_name):
        """Check and clear the print queue for the specified printer"""
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            try:
                # Get printer status
                printer_info = win32print.GetPrinter(printer_handle, 2)
                status = printer_info['Status']
                
                if status & win32print.PRINTER_STATUS_OFFLINE:
                    logging.warning(f"Printer {printer_name} is offline")
                    return False
                    
                # Check for jobs
                jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                if jobs:
                    logging.info(f"Found {len(jobs)} jobs in queue for {printer_name}")
                    # Clear the queue
                    win32print.SetPrinter(printer_handle, 0, None, win32print.PRINTER_CONTROL_PURGE)
                    logging.info(f"Successfully cleared {len(jobs)} jobs from {printer_name}")
                    return True
                return False
                    
            finally:
                win32print.ClosePrinter(printer_handle)
                
        except Exception as e:
            logging.error(f"Error checking/clearing queue for {printer_name}: {str(e)}")
            return False

    def is_admin(self):
        """Check if running with admin rights"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def get_printers(self):
        """Get list of available printers"""
        try:
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printer_info = win32print.EnumPrinters(flags, None, 1)
            printers = [printer[2] for printer in printer_info]
            logging.info(f"Found printers: {printers}")
            return printers
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
