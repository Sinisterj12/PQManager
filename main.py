import tkinter as tk
from tkinter import ttk
import win32print
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import threading
import logging
import json
import os

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
    if tray_icon:
        return  # If icon exists, don't create a new one
    
    try:
        import os
        import sys
        
        # Get the correct path whether running as script or exe
        if getattr(sys, 'frozen', False):
            # Running as exe
            base_path = sys._MEIPASS
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        icon_path = os.path.join(base_path, "Logo.ico")
        icon_image = Image.open(icon_path)
        
        menu = (
            item('Show', show_window),
            item('Exit', exit_app)
        )
        
        tray_icon = pystray.Icon(
            "PQManager",
            icon_image,
            "Printer Queue Manager",
            menu
        )
        tray_icon.run_detached()
    except Exception as e:
        logging.error(f"Error creating tray icon: {e}")

# Modify the minimize function
def minimize_to_tray():
    root.withdraw()  # Hide the main window
    if not tray_icon:  # Only create tray icon if it doesn't exist
        create_tray_icon()

# Modify the show_window function
def show_window(icon=None, item=None):
    global tray_icon
    if tray_icon:
        try:
            tray_icon.visible = False  # Hide the icon first
            tray_icon.stop()  # Then stop it
            tray_icon = None
        except Exception as e:
            logging.error(f"Error hiding tray icon: {e}")
    
    root.after(100, root.deiconify)  # Small delay before showing window
    root.lift()
    root.focus_force()

def exit_app(icon=None, item=None):
    global tray_icon, is_monitoring
    is_monitoring = False
    
    # Cancel any pending monitoring tasks first
    try:
        root.after_cancel(monitor_queue)
    except:
        pass
    
    # Stop the tray icon
    if tray_icon:
        try:
            tray_icon.visible = False
            tray_icon.stop()
            tray_icon = None
        except Exception as e:
            logging.error(f"Error stopping tray icon: {e}")
    
    # Destroy the root window
    try:
        root.quit()
        root.update()  # Process any remaining events
        root.destroy()
    except Exception as e:
        logging.error(f"Error destroying window: {e}")
    
    # Kill all threads
    for thread in threading.enumerate():
        if thread is not threading.current_thread():
            try:
                thread.join(timeout=0.5)
            except:
                pass
    
    # Force terminate
    import os
    os._exit(0)

# Printer selection dropdown
printer_dropdown = ttk.Combobox(root, textvariable=printer_var)
printer_dropdown['values'] = get_printers()
printer_dropdown.pack(pady=10)
printer_dropdown.set("Select Printer")

# Function to handle printer selection and start monitoring
def on_printer_select(event):
    selected_printer = printer_var.get()
    if selected_printer and selected_printer != "Select Printer":
        try:
            with open('printer_config.json', 'w') as f:
                config = {'printer': selected_printer}
                json.dump(config, f)
            logging.info(f"Saved printer selection: {selected_printer}")
            start_monitoring()
        except Exception as e:
            logging.error(f"Error saving printer selection: {e}")

# Add this function to load the saved printer on startup
def load_saved_printer():
    try:
        if os.path.exists('printer_config.json'):
            with open('printer_config.json', 'r') as f:
                config = json.load(f)
                saved_printer = config.get('printer')
                if saved_printer in printer_dropdown['values']:
                    printer_var.set(saved_printer)
                    logging.info(f"Loaded saved printer: {saved_printer}")
                    start_monitoring()
    except Exception as e:
        logging.error(f"Error loading saved printer: {e}")

# Call this after creating the printer dropdown
load_saved_printer()

# Button to manually clear the queue
clear_button = tk.Button(root, text="Clear Printer Queue", command=clear_queue)
clear_button.pack(pady=10)

# Minimize to tray button
tray_button = tk.Button(root, text="Minimize to Tray", command=minimize_to_tray)
tray_button.pack(pady=10)

# Start the main GUI loop
root.mainloop()
