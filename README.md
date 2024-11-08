# Printer Queue Manager

A Windows utility that automatically monitors and clears printer queues to prevent printing issues.

## Features
- Automatic printer queue monitoring and clearing
- System tray integration
- Manual queue clearing option
- Multiple printer support
- Clean exit handling

## Requirements
- Windows OS
- Python 3.x (if running from source)
- Required packages (if running from source):
  - tkinter
  - win32print
  - pystray
  - Pillow

## Installation
1. Download the latest release
2. Extract PQManager.exe and Logo.ico to your preferred location
3. Run PQManager.exe

## Usage
1. Select your printer from the dropdown
2. Click "Minimize to Tray" to run in background
3. Right-click the tray icon to show/exit
4. Use "Clear Printer Queue" button for manual clearing

## Building from Source 