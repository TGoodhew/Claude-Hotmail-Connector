; Inno Setup script for the Claude Hotmail Connector (Windows).
;
; Wraps the self-contained single-exe (build\hotmail-connector.exe, produced by
; `npm run build:exe`) in a familiar Next -> Finish installer that:
;   - installs the exe per-user (no admin / UAC),
;   - auto-configures Claude Desktop to launch it (`... setup`),
;   - offers to sign in to Microsoft at the end (`... login`),
;   - adds Start-menu shortcuts and a proper uninstaller.
;
; Build (needs Inno Setup 6, `iscc` on PATH):  npm run build:installer
; NOTE: the produced installer is UNSIGNED by design (see installer\welcome.txt).

#define AppName "Claude Hotmail Connector"
#define AppVersion "0.1.0"
#define AppExe "hotmail-connector.exe"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Tony Goodhew
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist-installer
OutputBaseFilename=HotmailConnectorSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#AppName}
; Reassure the user up front (shown before install; explains the SmartScreen warning).
InfoBeforeFile=welcome.txt

[Files]
Source: "..\build\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Sign in to Hotmail Connector"; Filename: "{app}\{#AppExe}"; Parameters: "login"
Name: "{group}\Reconfigure Claude Desktop"; Filename: "{app}\{#AppExe}"; Parameters: "setup"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"

[Run]
; Always: point Claude Desktop (and Code) at the installed exe.
Filename: "{app}\{#AppExe}"; Parameters: "setup"; StatusMsg: "Configuring Claude Desktop..."; Flags: runhidden
; Optional finish-page action: sign in now (opens the system browser).
Filename: "{app}\{#AppExe}"; Parameters: "login"; Description: "Sign in to your Microsoft account now"; Flags: postinstall skipifsilent

[UninstallRun]
; Best-effort: remove our entry from the Claude host configs on uninstall.
Filename: "{app}\{#AppExe}"; Parameters: "unsetup"; Flags: runhidden; RunOnceId: "unsetup"

[Messages]
; The finish page reminds the user to fully restart Claude Desktop.
FinishedLabel=Setup is complete.%n%nIMPORTANT: fully quit Claude Desktop from the system tray (closing the window only minimises it) and reopen it, so it loads the new connector.
