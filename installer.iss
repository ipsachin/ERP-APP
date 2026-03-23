[Setup]
AppId={{8B1D0E0B-4F15-4CB7-BB7A-3A91A3A6D8E2}
AppName=Liquimech ERP
AppVersion=1.0.0
AppPublisher=Liquimech
AppPublisherURL=https://liquimech.com.au
AppSupportURL=https://liquimech.com.au
AppUpdatesURL=https://liquimech.com.au
DefaultDirName={localappdata}\Liquimech\LiquimechERP
DefaultGroupName=Liquimech ERP
OutputDir=installer
OutputBaseFilename=LiquimechERP-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
WizardStyle=modern
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\Liquimech ERP.exe

[Files]
Source: "dist\Liquimech ERP\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\Liquimech ERP"; Filename: "{app}\Liquimech ERP.exe"
Name: "{autodesktop}\Liquimech ERP"; Filename: "{app}\Liquimech ERP.exe"

[Run]
Filename: "{app}\Liquimech ERP.exe"; Description: "Launch Liquimech ERP"; Flags: nowait postinstall skipifsilent
