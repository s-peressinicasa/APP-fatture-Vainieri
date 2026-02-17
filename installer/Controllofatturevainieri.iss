#define MyAppName "Controllo Fatture Vainieri"
#define MyAppExeName "ControlloFattureVainieri.exe"
#define MyAppPublisher "Peressini casa"
#define MyAppURL "https://github.com/"
#define MyDefaultDir "{autopf}\ControlloFattureVainieri"

; Versione da env (GitHub Actions la imposta), fallback manuale
#define MyAppVersion GetEnv("APP_VERSION")
#if MyAppVersion == ""
  #define MyAppVersion "0.1.0"
#endif

[Setup]
AppId={{C2B72343-14E2-4EE8-9D62-2F7F5C6A2F8B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={#MyDefaultDir}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=ControlloFattureVainieri-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; chiusura automatica se l'app è aperta
CloseApplications=yes
RestartApplications=yes

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Files]
Source: "..\dist\ControlloFattureVainieri\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Disinstalla {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Avvia {#MyAppName}"; Flags: nowait postinstall skipifsilent
