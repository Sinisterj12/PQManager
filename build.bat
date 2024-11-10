@echo off
echo Building Printer Queue Manager...

REM Create executable
pyinstaller --onefile --windowed --icon=Logo.ico --add-data "Logo.ico;." --hidden-import win32timezone main.py --name PQManager

REM Build service executable
pyinstaller --onefile --icon=Logo.ico --hidden-import win32timezone service.py --name PQManagerService

echo Build complete!
pause