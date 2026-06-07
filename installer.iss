[Setup]
AppName=Fakturkomat
AppVersion=1.0
AppPublisher=Zabka
DefaultDirName={localappdata}\Fakturkomat
DisableProgramGroupPage=yes
OutputDir=.
OutputBaseFilename=FakturkomatSetup
SetupIconFile=logo.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\Fakturkomat.exe

[Languages]
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"

[Files]
Source: "Fakturkomat.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo.ico";        DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userdesktop}\Fakturkomat"; Filename: "{app}\Fakturkomat.exe"; IconFilename: "{app}\logo.ico"

[Run]
Filename: "{app}\Fakturkomat.exe"; Description: "Uruchom Fakturkomat"; Flags: nowait postinstall skipifsilent
