# Printer Queue Manager

A Windows application that monitors and clears printer queues automatically, helping prevent printer jams and stuck print jobs.

## Features
- Automatic printer queue monitoring (every 15 seconds)
- System tray integration for minimal interference
- Manual queue clearing option
- Remembers your selected printer
- Minimizes to system tray (works with both minimize button and close button)
- Clean exit functionality from system tray

## Requirements
- Windows 10/11
- Python 3.x
- Required packages (automatically installed via requirements.txt):
  - pywin32 (for printer management)
  - pystray (system tray functionality)
  - Pillow (image handling for tray icon)
  - packaging (version management)

## Installation
1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/PQManager.git
   cd PQManager
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. Run the application:
   ```bash
   python main.py
   ```

2. Select your printer from the dropdown menu
3. The application will:
   - Automatically monitor and clear the selected printer's queue
   - Minimize to system tray when using the minimize button or close button
   - Continue running in the background

## System Tray Features
- Right-click the tray icon to:
  - Show the main window
  - Exit the application
- Left-click to show the main window

## Project Structure
```
PQManager/
├── assets/          # Contains application icons (Logo.ico)
├── config/          # Configuration files (settings.json)
├── logs/            # Application logs (printer_queue.log)
├── src/             # Source code
│   ├── __init__.py
│   ├── main.py
│   ├── config_manager.py
│   ├── printer_manager.py
│   └── tray_manager.py
├── .gitignore
├── main.py          # Application launcher
├── README.md
└── requirements.txt
```

## Development
- Main application code is in `src/main.py`
- System tray functionality in `src/tray_manager.py`
- Configuration management in `src/config_manager.py`
- Printer operations in `src/printer_manager.py`

## Contributing
1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License
[Your chosen license]

## Support
For issues or feature requests, please create an issue in the GitHub repository.