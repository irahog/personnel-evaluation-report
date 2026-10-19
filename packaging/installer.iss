; Inno Setup 6 script — per-user install (no administrator rights), Persian UI fonts installed for the current user.
; Build: iscc packaging\installer.iss   (after PyInstaller has produced dist\PersonnelEvaluationReport)

#define AppName "کارنامه پرسنلی"
#define AppVersion "0.2.2"
#define SrcDir "..\dist\PersonnelEvaluationReport"
#define FontDir "..\dist\PersonnelEvaluationReport\_internal\personnel_evaluation_report\static\fonts"

[Setup]
AppId={{6F2B3D0E-8E7A-4C55-9C0B-6A0D5B1F2E41}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Personnel Evaluation Report
DefaultDirName={localappdata}\Programs\PersonnelEvaluationReport
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\dist
OutputBaseFilename=PersonnelEvaluationReport-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=app.ico
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\PersonnelEvaluationReport.exe
; 64-bit Intel/AMD Windows, or Windows 11 on ARM (x64 emulation). Refuses 32-bit and Windows 10 on ARM.
ArchitecturesAllowed=x64compatible
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
Source: "{#FontDir}\BTitrBd.ttf"; DestDir: "{autofonts}"; FontInstall: "B Titr Bold"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\docs\fa\راهنمای کاربر.html"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\PersonnelEvaluationReport.exe"
Name: "{autoprograms}\راهنمای کارنامه پرسنلی"; Filename: "{app}\راهنمای کاربر.html"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\PersonnelEvaluationReport.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\PersonnelEvaluationReport.exe"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
; In-app update (silent install): reopen the app afterwards.
Filename: "{app}\PersonnelEvaluationReport.exe"; Flags: nowait skipifnotsilent
