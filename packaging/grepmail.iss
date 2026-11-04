; Inno Setup script for GrepMail (Windows, per-user install, no administrator rights needed)
; Build the program first:  see packaging\build_windows.bat  (creates dist\GrepMail\)
#define AppName "GrepMail"
#define AppVersion "2.1.0"

[Setup]
AppId={{6B0F3E1A-5D1C-4E8B-9A57-0A1B2C3D4E5F}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
PrivilegesRequired=lowest
OutputDir=..\installer
OutputBaseFilename=GrepMail-Setup-{#AppVersion}
SetupIconFile=..\grepmail.ico
UninstallDisplayIcon={app}\GrepMail.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked
Name: "autostart"; Description: "Start GrepMail when I sign in to Windows"; Flags: unchecked

[Files]
Source: "..\dist\GrepMail\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\GrepMail.exe"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\GrepMail.exe"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "GrepMail"; ValueData: """{app}\GrepMail.exe"""; Tasks: autostart; Flags: uninsdeletevalue

[Run]
Filename: "{app}\GrepMail.exe"; Description: "Launch GrepMail"; Flags: nowait postinstall skipifsilent

; The uninstaller removes only the program. The database and settings are left untouched.
