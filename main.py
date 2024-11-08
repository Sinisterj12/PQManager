import tkinter as tk
from tkinter import ttk
import win32print
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import threading
import logging

# Setup logging
logging.basicConfig(
    filename='printer_queue.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Initialize the main window
root = tk.Tk()
root.title("Printer Queue Manager")
root.geometry("300x200")
root.resizable(False, False)

# Global variables
printer_var = tk.StringVar()
tray_icon = None  # Global variable to hold the tray icon instance
is_monitoring = False  # Variable to track if monitoring has started

# Function to get the list of printers
def get_printers():
    printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
    return printers

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

# Function to minimize to tray
def minimize_to_tray():
    root.withdraw()  # Hide the main window
    create_tray_icon()  # Show the tray icon

# Override the close button to minimize to tray
root.protocol("WM_DELETE_WINDOW", minimize_to_tray)

# Function to create a system tray icon
def create_tray_icon():
    global tray_icon
    icon_image = Image.new('RGB', (64, 64), (0, 128, 255))  # Brighter blue icon
    draw = ImageDraw.Draw(icon_image)
    draw.rectangle([16, 16, 48, 48], fill=(255, 255, 255))  # White rectangle in center
    
    tray_icon = pystray.Icon(
        "PrinterQueueManager",
        icon_image,
        "Printer Queue Manager",  # Hover text
        menu=pystray.Menu(
            item('Show', show_window),
            item('Exit', exit_app)
        )
    )
    threading.Thread(target=tray_icon.run, daemon=True).start()  # Make thread daemon

# Function to show the window again
def show_window(icon=None, item=None):
    icon.stop()  # Stop the tray icon
    root.deiconify()  # Show the main window

# Function to exit the app from the system tray
def exit_app(icon=None, item=None):
    global tray_icon, is_monitoring
    is_monitoring = False  # Stop the monitoring loop
    if tray_icon:
        tray_icon.stop()  # Stop the tray icon
    root.quit()  # Close the main GUI event loop
    root.destroy()  # Ensure window is destroyed

# Printer selection dropdown
printer_dropdown = ttk.Combobox(root, textvariable=printer_var)
printer_dropdown['values'] = get_printers()
printer_dropdown.pack(pady=10)
printer_dropdown.set("Select Printer")

# Function to handle printer selection and start monitoring
def on_printer_select(event):
    if printer_var.get() != "Select Printer":
        start_monitoring()

printer_dropdown.bind("<<ComboboxSelected>>", on_printer_select)

# Button to manually clear the queue
clear_button = tk.Button(root, text="Clear Printer Queue", command=clear_queue)
clear_button.pack(pady=10)

# Minimize to tray button
tray_button = tk.Button(root, text="Minimize to Tray", command=minimize_to_tray)
tray_button.pack(pady=10)

# Start the main GUI loop
root.mainloop()
