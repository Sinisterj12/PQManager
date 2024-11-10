import tkinter as tk
from tkinter import ttk
import win32print
import logging
import json
import os
from src.tray_manager import TrayManager
from src.config_manager import ConfigManager

# Create necessary directories
os.makedirs('logs', exist_ok=True)

# Setup logging
logging.basicConfig(
    filename='logs/printer_queue.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Initialize config manager
config_manager = ConfigManager()

# Initialize the main window
root = tk.Tk()
root.title("Printer Queue Manager")
root.geometry("300x200")
root.resizable(False, False)

# Global variables
printer_var = tk.StringVar()
tray_manager = TrayManager(root)
is_monitoring = False

# Function to get the list of printers
def get_printers():
    return [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]

# Function to clear the printer queue for the selected printer
def clear_queue():
    selected_printer = printer_var.get()
    if selected_printer:
        try:
            printer_handle = win32print.OpenPrinter(selected_printer)
            jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
            for job in jobs:
                win32print.SetJob(printer_handle, job['JobId'], 0, None, win32print.JOB_CONTROL_DELETE)
            win32print.ClosePrinter(printer_handle)
            logging.info(f"Queue cleared for {selected_printer}.")
        except Exception as e:
            logging.error(f"Error clearing queue: {e}")
    else:
        logging.warning("No printer selected.")

# Function to monitor and clear the queue periodically
def monitor_queue():
    if is_monitoring and root.winfo_exists():
        clear_queue()
        root.after(15000, monitor_queue)  # 15000 ms = 15 seconds

# Function to start monitoring after a printer is selected
def start_monitoring():
    global is_monitoring
    if not is_monitoring:
        is_monitoring = True
        monitor_queue()

# Function to handle printer selection and start monitoring
def on_printer_select(event):
    selected_printer = printer_var.get()
    if selected_printer and selected_printer != "Select Printer":
        if config_manager.update_setting('selected_printer', selected_printer):
            logging.info(f"Saved printer selection: {selected_printer}")
            start_monitoring()
        else:
            logging.error("Failed to save printer selection")

# Function to load the saved printer on startup
def load_saved_printer():
    settings = config_manager.load_settings()
    saved_printer = settings.get('selected_printer')
    if saved_printer and saved_printer in printer_dropdown['values']:
        printer_var.set(saved_printer)
        logging.info(f"Loaded saved printer: {saved_printer}")
        if settings.get('auto_start_monitoring', True):
            start_monitoring()

# Create GUI elements
printer_dropdown = ttk.Combobox(root, textvariable=printer_var)
printer_dropdown['values'] = get_printers()
printer_dropdown.pack(pady=10)
printer_dropdown.set("Select Printer")
printer_dropdown.bind('<<ComboboxSelected>>', on_printer_select)

# Button to manually clear the queue
clear_button = tk.Button(root, text="Clear Printer Queue", command=clear_queue)
clear_button.pack(pady=10)

# Minimize to tray button
tray_button = tk.Button(root, text="Minimize to Tray", command=lambda: tray_manager.minimize_to_tray())
tray_button.pack(pady=10)

# Load saved printer settings
load_saved_printer()
