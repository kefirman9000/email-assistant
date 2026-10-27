; Inno Setup script for the Email Assistant installer.
; Built automatically by .github/workflows/build-installer.yml.
; APP_VERSION comes from version.json via an environment variable.

#define MyAppName "Email Assistant"
#define MyAppExe "Email Assistant.exe"
#define MyAppVersion GetEnv("APP_VERSION")

[Setup]
AppId={{8F3C2A51-7D4E-4B6A-9E2F-5C1D0A7B3E94}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher=Email Assistant
AppPublisherURL=https://github.com/kefirman9000/email-assistant
AppSupportURL=https://github.com/kefirman9000/email-assistant
; Installs just for the current user: no administrator rights needed
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableReadyPage=yes
OutputDir=..\dist
OutputBaseFilename=EmailAssistantSetup
SetupIconFile=..\assets\icon.ico
UninstallDisplayIcon={app}\{#MyAppExe}
UninstallDisplayName={#MyAppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; Close the app if it's running while installing an update
CloseApplications=yes
RestartApplications=no

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"

[Files]
Source: "..\dist\{#MyAppName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"; Tasks: desktopicon

[Run]
; Also runs after a silent update, so the app reopens on its own
Filename: "{app}\{#MyAppExe}"; Description: "Open {#MyAppName}"; Flags: nowait postinstall
