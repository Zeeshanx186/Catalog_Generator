; installer.iss — Inno Setup script for BOM Manual Downloader
; Download Inno Setup free at: https://jrsoftware.org/isinfo.php
; Compile with: iscc installer.iss

#define AppName      "BOM Manual Downloader"
#define AppVersion   "1.0.0"
#define AppPublisher "Your Organisation"
#define AppExeName   "BomDownloader.exe"
#define AppURL       "https://github.com/yourorg/bom-downloader"

[Setup]
AppId={{A3F1B2C4-E5D6-4789-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\BomDownloader
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=dist
OutputBaseFilename=BomDownloader-Setup-{#AppVersion}
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
WizardResizable=no
DisableDirPage=no
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesInstallIn64BitMode=x64os
MinVersion=10.0
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startupicon"; Description: "Launch at &Windows startup"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Application files from PyInstaller output
Source: "dist\BomDownloader\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}";                   Filename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}";         Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";             Filename: "{app}\{#AppExeName}"; Tasks: desktopicon
Name: "{userstartup}\{#AppName}";             Filename: "{app}\{#AppExeName}"; Tasks: startupicon

[Registry]
; Register as a handler for BOM file types (optional)
Root: HKCU; Subkey: "Software\Classes\.bom\OpenWithProgids"; ValueType: string; ValueName: "BomDownloader.Document"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Classes\BomDownloader.Document"; ValueType: string; ValueData: "BOM Manual Downloader File"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\BomDownloader.Document\DefaultIcon"; ValueType: string; ValueData: "{app}\{#AppExeName},0"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\BomDownloader.Document\shell\open\command"; ValueType: string; ValueData: """{app}\{#AppExeName}"" ""%1"""; Flags: uninsdeletekey

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; \
  Flags: nowait postinstall skipifsilent

[Code]
// Show a friendly message if .NET / WebView2 Runtime is missing
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
