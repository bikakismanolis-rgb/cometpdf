#define MyAppName "CometPDF"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "CometPDF"
#define MyAppExeName "CometPDF.exe"

[Setup]
AppId={{C5E914B2-7B1A-4C24-93E7-3E48F604A150}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\CometPDF
DefaultGroupName=CometPDF
DisableProgramGroupPage=yes
OutputDir=release
OutputBaseFilename=CometPDFSetup-1.0.0
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=cometpdf.ico
UninstallDisplayIcon={app}\CometPDF.exe
PrivilegesRequired=lowest
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"; Flags: checkedonce
Name: "rightclick"; Description: "Add 'Convert to PDF' to the Windows right-click menu"; GroupDescription: "Windows Explorer:"; Flags: checkedonce

[Files]
Source: "dist\CometPDF.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "cometpdf.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "THIRD_PARTY_NOTICES.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "PRODUCT_PLAN.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\CometPDF"; Filename: "{app}\CometPDF.exe"; IconFilename: "{app}\cometpdf.ico"
Name: "{autodesktop}\CometPDF"; Filename: "{app}\CometPDF.exe"; IconFilename: "{app}\cometpdf.ico"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Classes\*\shell\CometPDF"; ValueType: string; ValueName: ""; ValueData: "Convert to PDF"; Flags: uninsdeletekey; Tasks: rightclick
Root: HKCU; Subkey: "Software\Classes\*\shell\CometPDF"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\cometpdf.ico"; Tasks: rightclick
Root: HKCU; Subkey: "Software\Classes\*\shell\CometPDF\command"; ValueType: string; ValueName: ""; ValueData: """{app}\CometPDF.exe"" --convert ""%1"" --open"; Tasks: rightclick

[Run]
Filename: "{app}\CometPDF.exe"; Description: "Launch CometPDF"; Flags: nowait postinstall skipifsilent
