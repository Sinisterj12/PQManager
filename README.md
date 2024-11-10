# Printer Queue Manager

A Windows service that automatically monitors and clears printer queues to prevent printing issues.

## Features
- Runs as a Windows service for automatic queue management
- System tray integration for easy control
- Manual queue clearing option
- Automatic printer queue monitoring (15-second intervals)
- Persistent printer selection
- Clean exit handling

## Requirements
### For Users
- Windows 10/11
- Administrative privileges for installation

### For Developers
- Python 3.x
- Required packages:
  - pywin32
  - pystray
  - Pillow
  - tkinter (included with Python)

## Installation
### User Installation
1. Download the latest release (PQManagerSetup.exe)
2. Run the installer as administrator
3. Select your printer from the system tray application

### Developer Installation
1. Clone the repository
2. Create virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. After installation, the service starts automatically
2. Click the system tray icon to:
   - Select your printer
   - Manually clear queue
   - View service status
   - Exit application

## Building from Source
1. Ensure all requirements are installed
2. Run build script:
   ```bash
   build.bat
   ```
3. Installer will be created in `installer/Output/PQManagerSetup.exe`

## Support
For issues or feature requests, please create an issue in the GitHub repository.