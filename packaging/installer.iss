; Inno Setup 6 script — per-user install (no administrator rights), Persian UI fonts installed for the current user.
; Build: iscc packaging\installer.iss   (after PyInstaller has produced dist\Karnameh)

#define AppName "کارنامه پرسنلی"
#define AppVersion "0.1.0"
#define SrcDir "..\dist\Karnameh"
#define FontDir "..\dist\Karnameh\_internal\karnameh\static\fonts"

[Setup]
AppId={{6F2B3D0E-8E7A-4C55-9C0B-6A0D5B1F2E41}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Karnameh
DefaultDirName={localappdata}\Programs\Karnameh
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\dist
OutputBaseFilename=Karnameh-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\Karnameh.exe
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
; Farsi.isl: unofficial translation from jrsoftware/issrc (needs Inno Setup 6.5+).
; Inno Setup does not allow comments at the end of an entry line.
Name: "fa"; MessagesFile: "Farsi.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#SrcDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Fonts: per-user installation (Windows 10 1803+) when not running as administrator.
Source: "{#FontDir}\BNazanin.ttf"; DestDir: "{autofonts}"; FontInstall: "B Nazanin"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "{#FontDir}\BNaznnBd.ttf"; DestDir: "{autofonts}"; FontInstall: "B Nazanin Bold"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "{#FontDir}\BMITRA.TTF"; DestDir: "{autofonts}"; FontInstall: "B Mitra"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "{#FontDir}\BMitraBd.ttf"; DestDir: "{autofonts}"; FontInstall: "B Mitra Bold"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "{#FontDir}\Vazirmatn-FD-Regular.ttf"; DestDir: "{autofonts}"; FontInstall: "Vazirmatn FD Regular"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "{#FontDir}\Vazirmatn-FD-Bold.ttf"; DestDir: "{autofonts}"; FontInstall: "Vazirmatn FD Bold"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\docs\fa\راهنمای کاربر.html"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\Karnameh.exe"
Name: "{autoprograms}\راهنمای کارنامه پرسنلی"; Filename: "{app}\راهنمای کاربر.html"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\Karnameh.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Karnameh.exe"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
