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
        logging.info("Initializing PrinterManager")
        # Check admin rights on startup
        if not self.is_admin():
            logging.warning("Application is not running with administrator rights")
        self._verify_win32_modules()

    def _verify_win32_modules(self):
        """Verify all required win32 modules are available"""
        required_modules = ['win32print', 'win32api', 'win32con']
        missing_modules = []
        for module in required_modules:
            try:
                __import__(module)
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
                logging.info(f"Queue length for {printer_name}: {count}")
                
                if count > 0:
                    # Log job details
                    for job in jobs:
                        status_str = self._get_job_status_string(job['Status'])
                        logging.info(f"Job {job['JobId']}: Status={status_str}, Document={job.get('pDocument', 'Unknown')}")
                return count
                
            finally:
                if printer_handle:
                    win32print.ClosePrinter(printer_handle)
                    
        except Exception as e:
            logging.error(f"Failed to get queue length: {e}")
            if "win32timezone" in str(e):
                logging.info("Attempting to continue without win32timezone...")
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
            logging.warning("No printer selected for queue clearing")
            return False

        if not self.is_admin():
            logging.error("Cannot clear queue: Administrator rights required")
            return False

        try:
            logging.info(f"Attempting to open printer: {printer_name}")
            handle = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
            printer_handle = win32print.OpenPrinter(printer_name, handle)
            
            try:
                jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                if not jobs:
                    logging.info(f"No jobs in queue for {printer_name}")
                    return True

                logging.info(f"Found {len(jobs)} jobs to clear")
                success = True
                
                for job in jobs:
                    try:
                        job_id = job['JobId']
                        status_str = self._get_job_status_string(job['Status'])
                        logging.info(f"Processing job {job_id} (Status: {status_str})")
                        
                        # Try to cancel first
                        try:
                            logging.info(f"Attempting to cancel job {job_id}")
                            win32print.SetJob(printer_handle, job_id, 0, None, win32print.JOB_CONTROL_CANCEL)
                            time.sleep(0.1)
                        except Exception as e:
                            if "parameter is incorrect" not in str(e).lower():
                                logging.warning(f"Non-critical error canceling job {job_id}: {e}")
                        
                        # Then try to delete
                        try:
                            logging.info(f"Attempting to delete job {job_id}")
                            win32print.SetJob(printer_handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
                            logging.info(f"Successfully cleared job {job_id}")
                        except Exception as e:
                            if "parameter is incorrect" not in str(e).lower():
                                logging.error(f"Failed to delete job {job_id}: {e}")
                                success = False
                            else:
                                logging.info(f"Job {job_id} appears to be already cleared")
                        
                    except Exception as e:
                        if "parameter is incorrect" not in str(e).lower():
                            success = False
                            logging.error(f"Failed to process job {job_id}: {e}")

                # Verify queue is actually empty
                remaining_jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                if remaining_jobs:
                    logging.warning(f"{len(remaining_jobs)} jobs could not be cleared")
                    success = False
                else:
                    logging.info("Queue verified empty - all jobs cleared successfully")
                    success = True  # If queue is empty, consider it a success regardless of errors

                return success

            finally:
                win32print.ClosePrinter(printer_handle)
                logging.info("Printer handle closed")

        except Exception as e:
            logging.error(f"Printer access error: {e}")
            if isinstance(e, win32print.error) and e.winerror == 5:
                logging.error("Access denied - Administrator rights required")
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
