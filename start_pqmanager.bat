@echo off
echo Setting up PQManager auto-start...

REM Check for admin rights
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges...
) else (
    echo This script requires administrator privileges.
    echo Please right-click and select "Run as administrator"
    pause
    exit /b 1
)

REM Create shortcut in All Users Startup folder
echo Creating startup shortcut...
powershell -Command "$WS = New-Object -ComObject WScript.Shell; $Shortcut = $WS.CreateShortcut('%ProgramData%\Microsoft\Windows\Start Menu\Programs\StartUp\PQManager.lnk'); $Shortcut.TargetPath = '%~dp0PQManager.exe'; $Shortcut.Arguments = '--minimized'; $Shortcut.WorkingDirectory = '%~dp0'; $Shortcut.Save()"

REM Verify shortcut was created
if exist "%ProgramData%\Microsoft\Windows\Start Menu\Programs\StartUp\PQManager.lnk" (
    echo Successfully created startup shortcut.
    echo PQManager will now start automatically when Windows starts.
) else (
    echo Failed to create startup shortcut.
    pause
    exit /b 1
)

REM Start the program now
echo Starting PQManager...
start "" "%~dp0PQManager.exe" --minimized

echo.
echo PQManager will now start automatically with Windows.
echo The program has also been started.
pause
