; installer.iss (FULL) — builds Setup.exe from dist folder

#define MyAppName "VARDA Control Center"
#define MyAppVersion "1.0.0"
#define MyAppExeName "VARDA Control Center.exe"

[Setup]
SetupIconFile=varda.ico
AppId={{B3B29F03-7A6A-4C3A-9D45-8C4E9B5E7A11}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename=VARDA_Control_Center_Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
DisableProgramGroupPage=yes

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "dist\VARDA Control Center\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "varda.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{autoprograms}\VARDA Control Center"; Filename: "{app}\VARDA Control Center.exe"; IconFilename: "{app}\varda.ico"
Name: "{autodesktop}\VARDA Control Center"; Filename: "{app}\VARDA Control Center.exe"; Tasks: desktopicon; IconFilename: "{app}\varda.ico"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent