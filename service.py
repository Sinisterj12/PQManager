import win32serviceutil
import win32service
import win32event
import win32print
import servicemanager
import logging
import os
import sys
import json

# Get the absolute path to the config and log files
if hasattr(sys, 'frozen'):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, 'printer_config.json')
LOG_FILE = os.path.join(BASE_DIR, 'printer_service.log')

# Set up logging with more details
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='a'  # append mode
)

class PrinterQueueService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PrinterQueueService"
    _svc_display_name_ = "Printer Queue Manager Service"
    _svc_description_ = "Monitors and clears printer queues automatically"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.is_running = True
        logging.info("Service initializing...")
        self.selected_printer = self.load_config()
        logging.info(f"Loaded printer from config: {self.selected_printer}")

    def load_config(self):
        try:
            logging.info(f"Looking for config file at: {CONFIG_FILE}")
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    printer = config.get('printer', '')
                    logging.info(f"Found printer in config: {printer}")
                    return printer
            else:
                logging.error(f"Config file not found at: {CONFIG_FILE}")
        except Exception as e:
            logging.error(f"Error loading config: {str(e)}")
        return ''

    def clear_printer_queue(self):
        if not self.selected_printer:
            logging.warning("No printer selected in config")
            return

        try:
            printer_handle = win32print.OpenPrinter(self.selected_printer)
            jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
            for job in jobs:
                win32print.SetJob(printer_handle, job['JobId'], 0, None, win32print.JOB_CONTROL_DELETE)
            win32print.ClosePrinter(printer_handle)
            logging.info(f"Queue cleared for {self.selected_printer}")
        except Exception as e:
            logging.error(f"Error clearing queue: {e}")

    def SvcDoRun(self):
        try:
            if not self.selected_printer:
                logging.error("No printer selected. Service stopping.")
                self.SvcStop()
                return
            
            logging.info(f"Service starting with printer: {self.selected_printer}")
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            
            while self.is_running:
                self.clear_printer_queue()
                win32event.WaitForSingleObject(self.stop_event, 15000)
        except Exception as e:
            logging.error(f"Service error: {e}")
            self.SvcStop()

    def SvcStop(self):
        logging.info("Service stop requested")
        self.is_running = False
        win32event.SetEvent(self.stop_event)
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(PrinterQueueService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(PrinterQueueService)