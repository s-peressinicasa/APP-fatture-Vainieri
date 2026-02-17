; -- ControlloFattureVainieri.iss --
; Compilazione (PowerShell, da root APP):
;   $env:APP_VERSION="0.1.2"
;   & "C:\...\ISCC.exe" "installer\ControlloFattureVainieri.iss"

#define MyAppName "Controllo Fatture Vainieri"
#define MyAppPublisher "Peressini casa"
#define MyAppExeName "ControlloFattureVainieri.exe"

; Legge versione da variabile ambiente (fallback se non impostata)
#define MyAppVersion GetEnv("APP_VERSION")
#if MyAppVersion == ""
  #define MyAppVersion "0.0.0"
#endif

[Setup]
AppId={{F18D9C66-3E3A-4A1F-B1B5-5AE7E77F6B29}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}

; ✅ mostra la pagina Start Menu
DisableProgramGroupPage=no

; output: APP\installer\Output\...
OutputDir=Output
OutputBaseFilename=ControlloFattureVainieri-Setup
Compression=lzma
SolidCompression=yes

; icona dell'installer (relativa a APP\installer\...)
SetupIconFile=..\assets\icon.ico

; icona mostrata in "Programmi e funzionalità"
UninstallDisplayIcon={app}\{#MyAppExeName}

WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; (opzionali ma utili)
CloseApplications=yes
RestartApplications=yes

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "Crea un'icona sul Desktop"; GroupDescription: "Icone aggiuntive:"; Flags: unchecked

[Files]
; ✅ ATTENZIONE: lo script è in APP\installer\, quindi dist è a ..\dist\
Source: "..\dist\ControlloFattureVainieri\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu (gruppo scelto dall’utente)
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{group}\Disinstalla {#MyAppName}"; Filename: "{uninstallexe}"

; Desktop opzionale (checkbox)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Avvia {#MyAppName}"; Flags: nowait postinstall skipifsilent
