; Inno Setup script for the Claude Hotmail Connector (Windows).
;
; Wraps the self-contained single-exe (build\hotmail-connector.exe, produced by
; `npm run build:exe`) in a familiar Next -> Finish installer that:
;   - installs the exe per-user (no admin / UAC),
;   - asks for the user's OWN Entra app registration client id (no shared id is
;     bundled; spec v2 section 8.1) - or takes it from /CLIENTID=<id> when silent,
;   - auto-configures Claude Desktop to launch it (`... setup --client-id <id>`),
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
; Always: save the client id and point Claude Desktop at the installed exe.
Filename: "{app}\{#AppExe}"; Parameters: "setup{code:SetupClientIdArg}"; StatusMsg: "Configuring Claude Desktop..."; Flags: runhidden
; Optional finish-page action: sign in now (opens the system browser).
Filename: "{app}\{#AppExe}"; Parameters: "login"; Description: "Sign in to your Microsoft account now"; Flags: postinstall skipifsilent

[UninstallRun]
; Best-effort: remove our entry from the Claude host configs on uninstall.
Filename: "{app}\{#AppExe}"; Parameters: "unsetup"; Flags: runhidden; RunOnceId: "unsetup"

[Messages]
; The finish page reminds the user to fully restart Claude Desktop.
FinishedLabel=Setup is complete.%n%nIMPORTANT: fully quit Claude Desktop from the system tray (closing the window only minimises it) and reopen it, so it loads the new connector.

[Code]
var
  ClientIdPage: TInputQueryWizardPage;

{ settings.json written by `setup --client-id` (see src/settings.ts). }
function HasSavedClientId(): Boolean;
begin
  Result := FileExists(ExpandConstant('{localappdata}\claude-hotmail-connector\settings.json'));
end;

{ Entra application (client) ids are GUIDs: 8-4-4-4-12 hex digits. }
function IsGuid(const S: String): Boolean;
var
  I: Integer;
  C: Char;
begin
  Result := Length(S) = 36;
  if not Result then Exit;
  for I := 1 to 36 do
  begin
    C := S[I];
    if (I = 9) or (I = 14) or (I = 19) or (I = 24) then
      Result := C = '-'
    else
      Result := ((C >= '0') and (C <= '9')) or ((C >= 'a') and (C <= 'f')) or ((C >= 'A') and (C <= 'F'));
    if not Result then Exit;
  end;
end;

function EnteredClientId(): String;
begin
  Result := Trim(ClientIdPage.Values[0]);
end;

procedure InitializeWizard();
var
  Hint: String;
begin
  Hint := 'This connector signs in through a Microsoft Entra app registration that YOU own. ' +
    'If you do not have one yet, follow "Register a Microsoft Entra application" in the README ' +
    '(or run scripts\register-app.ps1), then paste its Application (client) ID here.';
  if HasSavedClientId() then
    Hint := Hint + #13#10#13#10 + 'A client ID is already saved on this PC. Leave the box empty to keep it.';
  ClientIdPage := CreateInputQueryPage(wpInfoBefore,
    'Your Microsoft app registration',
    'Enter the Application (client) ID of your own Entra app registration.',
    Hint);
  ClientIdPage.Add('Application (client) ID:', False);
  ClientIdPage.Values[0] := ExpandConstant('{param:CLIENTID|}');
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Id: String;
begin
  Result := True;
  if CurPageID = ClientIdPage.ID then
  begin
    Id := EnteredClientId();
    if (Id = '') and HasSavedClientId() then Exit;
    if not IsGuid(Id) then
    begin
      MsgBox('Please enter a valid Application (client) ID, e.g. 00000000-0000-0000-0000-000000000000.',
        mbError, MB_OK);
      Result := False;
    end;
  end;
end;

{ Used by [Run]: pass the entered id to `setup`, or nothing to keep the saved one. }
function SetupClientIdArg(Param: String): String;
begin
  if EnteredClientId() = '' then
    Result := ''
  else
    Result := ' --client-id "' + EnteredClientId() + '"';
end;
