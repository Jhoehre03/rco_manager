#define MyAppName "RCO Manager"
#define MyAppVersion "1.2.3"
#define MyAppPublisher "Jhoehre03"
#define MyAppExeName "RCO Manager.exe"
#define MyAppDir "dist\RCO Manager"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=RCO_Manager_Setup_v{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0

; Ícone do instalador (opcional — descomente se tiver um .ico)
; SetupIconFile=icon.ico

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Ícones adicionais:"; Flags: checkedonce

[Files]
; Executável principal
Source: "{#MyAppDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Pasta _internal (dependências do PyInstaller)
Source: "{#MyAppDir}\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

; Credenciais OAuth (necessário para Google Sheets)
Source: "{#MyAppDir}\oauth_credentials.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Iniciar {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
// ── Verificação do .NET 6.0 Desktop Runtime ──────────────────────────────────
function DotNet6Installed(): Boolean;
var
  Key: String;
  Names: TArrayOfString;
  I: Integer;
begin
  Result := False;
  Key := 'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App';
  if RegGetValueNames(HKLM, Key, Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      if Pos('6.', Names[I]) = 1 then
      begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;

// ── Verificação do WebView2 Runtime ──────────────────────────────────────────
function WebView2Installed(): Boolean;
var
  Version: String;
begin
  Result := RegQueryStringValue(HKLM,
    'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
    'pv', Version) and (Version <> '') and (Version <> '0.0.0.0');
  if not Result then
    Result := RegQueryStringValue(HKLM,
      'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
      'pv', Version) and (Version <> '') and (Version <> '0.0.0.0');
end;

// ── Download e instalação silenciosa ─────────────────────────────────────────
function DownloadAndInstall(URL, TempFile, Args: String; Description: String): Boolean;
var
  ResultCode: Integer;
  PSCmd: String;
begin
  Result := False;
  WizardForm.StatusLabel.Caption := 'Baixando ' + Description + '...';
  PSCmd := 'Invoke-WebRequest -Uri "' + URL + '" -OutFile "' + TempFile + '" -UseBasicParsing';
  if Exec('powershell.exe', '-NoProfile -ExecutionPolicy Bypass -Command "' + PSCmd + '"',
          '', SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0) then
  begin
    WizardForm.StatusLabel.Caption := 'Instalando ' + Description + '...';
    if Exec(TempFile, Args, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      Result := True;
    DeleteFile(TempFile);
  end;
end;

procedure InstallPrerequisites();
var
  TempDir: String;
begin
  TempDir := ExpandConstant('{tmp}');

  if not DotNet6Installed() then
  begin
    if MsgBox('.NET 6.0 Desktop Runtime não encontrado.' + #13#10 +
              'É necessário para o RCO Manager funcionar.' + #13#10#13#10 +
              'Deseja instalar agora? (~55 MB)',
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      DownloadAndInstall(
        'https://download.visualstudio.microsoft.com/download/pr/8a1e6a00-b3cc-4f79-b5b2-edcd96f48e17/90f5c7f3b2bdc0af8c24f1aa89e5f3de/windowsdesktop-runtime-6.0.36-win-x64.exe',
        TempDir + '\dotnet6.exe',
        '/install /quiet /norestart',
        '.NET 6.0 Runtime'
      );
    end;
  end;

  if not WebView2Installed() then
  begin
    if MsgBox('Microsoft Edge WebView2 Runtime não encontrado.' + #13#10 +
              'É necessário para o RCO Manager funcionar.' + #13#10#13#10 +
              'Deseja instalar agora? (~2 MB)',
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      DownloadAndInstall(
        'https://go.microsoft.com/fwlink/p/?LinkId=2124703',
        TempDir + '\webview2.exe',
        '/silent /install',
        'WebView2 Runtime'
      );
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    InstallPrerequisites();
end;
