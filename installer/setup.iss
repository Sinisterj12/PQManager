[Setup]
AppName=Printer Queue Manager
AppVersion=1.0
DefaultDirName={pf}\PrinterQueueManager
DefaultGroupName=Printer Queue Manager
OutputDir=Output
OutputBaseFilename=PQManagerSetup
SetupIconFile=assets\Logo.ico
Compression=lzma
SolidCompression=yes
UninstallDisplayIcon={app}\Logo.ico
PrivilegesRequired=admin

[Dirs]
Name: "{app}"; Permissions: everyone-full

[Files]
Source: "..\dist\PQManager.exe"; DestDir: "{app}"
Source: "..\dist\PQManagerService.exe"; DestDir: "{app}"
Source: "..\Logo.ico"; DestDir: "{app}"
Source: "..\printer_config.json"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Icons]
Name: "{group}\Printer Queue Manager"; Filename: "{app}\PQManager.exe"; IconFilename: "{app}\Logo.ico"
Name: "{commondesktop}\Printer Queue Manager"; Filename: "{app}\PQManager.exe"; IconFilename: "{app}\Logo.ico"

[Run]
; First stop and remove any existing service
Filename: "net"; Parameters: "stop PrinterQueueService"; Flags: runhidden; WorkingDir: "{app}"; StatusMsg: "Stopping service..."
Filename: "{app}\PQManagerService.exe"; Parameters: "remove"; Flags: runhidden waituntilterminated; StatusMsg: "Removing old service..."

; Wait a moment for service cleanup
Filename: "{sys}\ping.exe"; Parameters: "-n 3 127.0.0.1"; Flags: runhidden; StatusMsg: "Cleaning up..."

; Install and start new service
Filename: "{app}\PQManagerService.exe"; Parameters: "--startup delayed install"; Flags: runhidden waituntilterminated; StatusMsg: "Installing service..."
Filename: "net"; Parameters: "start PrinterQueueService"; Flags: runhidden; StatusMsg: "Starting service..."
Filename: "{app}\PQManager.exe"; Description: "Launch application"; Flags: postinstall nowait