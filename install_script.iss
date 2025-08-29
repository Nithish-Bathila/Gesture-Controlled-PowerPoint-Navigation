[Setup]
AppName=Gesture-Controlled PowerPoint Navigation
AppVerName=Gesture-Controlled PowerPoint Navigation 2.9
AppVersion=2.9
AppPublisher=P S A Nithish Bathila
DefaultDirName={autopf}\GestureControl
DefaultGroupName=Gesture Control
OutputDir=dist
OutputBaseFilename=GestureControlInstaller_2.9
SetupIconFile=icon.ico
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
WizardStyle=modern
AllowNoIcons=yes
CreateAppDir=yes
UninstallDisplayIcon={app}\GestureControl.exe
UninstallDisplayName=Gesture-Controlled PowerPoint Navigation
VersionInfoCopyright=© 2025 Nithish Bathila. All rights reserved.
VersionInfoProductName=Gesture-Controlled PowerPoint Navigation

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startupicon"; Description: "Start the application with Windows"; GroupDescription: "Startup options:"; Flags: unchecked

[Files]
Source: "dist\GestureControl.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.png"; DestDir: "{app}"; Flags: ignoreversion

; Include user guide images needed by the app
Source: "welcome.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "gestures.png"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Gesture Control"; Filename: "{app}\GestureControl.exe"; IconFilename: "{app}\icon.ico"; WorkingDir: "{app}"
Name: "{commondesktop}\Gesture Control"; Filename: "{app}\GestureControl.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon; WorkingDir: "{app}"
Name: "{userstartup}\Gesture Control"; Filename: "{app}\GestureControl.exe"; IconFilename: "{app}\icon.ico"; Tasks: startupicon; WorkingDir: "{app}"

[Run]
; Run the app immediately after installation (except if silent install)
Filename: "{app}\GestureControl.exe"; Description: "Launch Gesture Control"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\*.*"
Type: files; Name: "{userstartup}\Gesture Control.lnk"
