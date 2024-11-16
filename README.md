# PQ Manager (Printer Queue Manager)

A lightweight Windows application designed to automatically manage and clear printer queues, specifically optimized for grocery store environments.

## Features

- 🖨️ Automatic printer queue monitoring and clearing
- 🔄 10-second monitoring cycle
- 🎯 Multiple printer support
- 💾 Settings persistence
- 🚀 Automatic startup option
- 🔒 Administrative rights handling
- 📊 Detailed logging for troubleshooting

## Requirements

- Windows OS
- Python 3.12 or higher
- Administrator rights (for printer management)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/PQManager.git
   cd PQManager
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Building the Executable

To build the standalone executable:

```bash
pyinstaller PQManager.spec
```

The executable will be created in the `dist` folder.

## Usage

1. Run the application as Administrator
2. Select your receipt printer from the dropdown
3. The application will automatically:
   - Monitor the printer queue every 10 seconds
   - Clear any stuck print jobs
   - Minimize to system tray

## Configuration

Settings are stored in: `%APPDATA%\PQManager\settings.json`
- Selected printer
- Monitoring interval
- Startup preferences

## Logging

Logs are stored in: `logs/printer_queue.log`
- Queue operations
- Error messages
- Printer status

## Development

### Project Structure
```
/PQManager
├── main.py             # Entry point
├── bootstrap.py        # Import and environment setup
├── src/
│   ├── main.py         # Core application logic
│   ├── printer_manager.py  # Printer queue operations
│   ├── tray_manager.py     # System tray handling
│   └── config_manager.py   # Settings persistence
├── assets/
│   └── Logo.ico        # Application icon
└── dist/
    └── PQManager.exe   # Compiled executable
```

### Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

[MIT License](LICENSE)

## Author

James

## Support

For support, please check the logs at `logs/printer_queue.log` first, then open an issue on GitHub.