#define MyAppName "Gestion des relevés bancaires"
#define MyAppVersion "1.0"
#define MyAppPublisher "Louis-Marie"

[Setup]
AppId={{A7D9D6A5-8B9C-4E3A-B5F2-123456789ABC}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={userdocs}\comptabilite
UninstallFilesDir={app}\programme
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
OutputDir=.
OutputBaseFilename=Installateur_GestionReleves
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Files]
Source: "TraitementReleve.ps1"; DestDir: "{app}\programme"; Flags: ignoreversion
Source: "Modele.ods"; DestDir: "{app}\programme"; Flags: ignoreversion
Source: "Notice.pdf"; DestDir: "{app}\programme"; Flags: ignoreversion

; XML inclus dans l'installateur, mais pas installé directement.
Source: "Faisons les comptes.xml"; Flags: dontcopy


[Icons]
Name: "{autoprograms}\Gestion des relevés bancaires\Notice"; \
    Filename: "{app}\programme\Notice.pdf"


[UninstallDelete]
; Ce fichier n'est pas installé par Inno Setup :
; il peut avoir été créé par InstallTask().
Type: files; Name: "{app}\programme\TraitementReleve.ps1"
Type: files; Name: "{app}\programme\Modele.ods" 
Type: files; Name: "{app}\programme\Notice.pdf"
Type: files; Name: "{app}\programme\Faisons les comptes.xml"


[Code]

procedure InstallTask();
var
  PowerShell: string;
  XmlSource: string;
  XmlPersonnalise: string;
  Script: string;
  ResultCode: Integer;
  Success: Boolean;
begin

  PowerShell := ExpandConstant(
  '{sys}\WindowsPowerShell\v1.0\powershell.exe'
  );
  { Extraction du XML générique dans le dossier temporaire }

  ExtractTemporaryFile('Faisons les comptes.xml');

  XmlSource := ExpandConstant(
    '{tmp}\Faisons les comptes.xml'
  );

  XmlPersonnalise := ExpandConstant(
    '{app}\programme\Faisons les comptes.xml'
  );

  { Personnalisation du XML avec les informations
    de l'utilisateur qui installe le programme }

  Script :=
    '$xml = Get-Content -Path ''' + XmlSource + ''' -Raw; ' +
    '$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value; ' +
    '$xml = $xml.Replace(''SID_UTILISATEUR'', $sid); ' +
    '$xml = $xml.Replace(''DATE_DEBUT'', (Get-Date -Format ''yyyy-MM-dd'') + ''T08:00:00''); ' +
    '$xml = $xml.Replace(''INTERVALLE_JOURS'', ''' + IntToStr(83) + '''); ' +
    '$xml = $xml.Replace(''CHEMIN_SCRIPT'', ''' +  ExpandConstant('{app}\programme\TraitementReleve.ps1') + '''); ' +
    '$xml = $xml.Replace(''DOSSIER_TRAVAIL'', ''' + ExpandConstant('{app}') + '''); ' +
    '[System.IO.File]::WriteAllText(''' + XmlPersonnalise + ''', $xml, [System.Text.Encoding]::Unicode);';

  if (not Exec(
    PowerShell,
    '-NoProfile -ExecutionPolicy Bypass -Command "' +
    Script +
    '"',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  )) or (ResultCode <> 0) then
  begin
    MsgBox(
      'Impossible de préparer le fichier de configuration de la tâche planifiée.'#13#10 +
      'Le fichier XML personnalisé a été conservé dans le dossier programme.',
      mbError,
      MB_OK
    );
    Exit;
  end;
  { Élévation temporaire pour créer la tâche }

  Script :=
    'Unregister-ScheduledTask ' +
    '-TaskName ''Faisons les comptes'' ' +
    '-Confirm:$false ' +
    '-ErrorAction SilentlyContinue; ' +
    'Register-ScheduledTask ' +
    '-TaskName ''Faisons les comptes'' ' +
    '-Xml (Get-Content -Path ''' +
      XmlPersonnalise +
      ''' -Raw) ' +
    '-Force;';

  Success := ShellExec(
    'runas',
    PowerShell,
    '-NoProfile -ExecutionPolicy Bypass -Command "' +
    Script +
    '"',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  );

  if not Success then
  begin
      {UAC refusée : le XML personnalisé est conservé pour
      permettre une création manuelle de la tâche. }

    MsgBox(
      'La création automatique de la tâche planifiée n''a pas été effectuée.'#13#10 +
      'Le fichier XML personnalisé a été conservé dans :'#13#10 +
      XmlPersonnalise + #13#10#13#10 +
      'Vous pouvez utiliser ce fichier pour créer la tâche manuellement.',
      mbInformation,
      MB_OK
    );
    Exit;
  end;

  if ResultCode <> 0 then
  begin
    MsgBox(
      'La tâche planifiée n''a pas pu être créée.'#13#10 +
      'Le fichier XML personnalisé a été conservé dans :'#13#10 +
      XmlPersonnalise + #13#10#13#10 +
      'Code d''erreur : ' + IntToStr(ResultCode),
      mbError,
      MB_OK
    );
    Exit;
  end;

  { La tâche a été créée : le XML personnalisé n'est plus nécessaire. }

  DeleteFile(XmlPersonnalise);

end;


procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    InstallTask();
end;


procedure CurUninstallStepChanged(
  CurUninstallStep: TUninstallStep
);
var
  PowerShell: string;
  ResultCode: Integer;
begin

  if CurUninstallStep = usUninstall then
  begin

    PowerShell := ExpandConstant(
      '{sys}\WindowsPowerShell\v1.0\powershell.exe'
    );

    Exec(
      PowerShell,
      '-NoProfile -ExecutionPolicy Bypass -Command ' +
      '"Unregister-ScheduledTask ' +
      '-TaskName ''Faisons les comptes'' ' +
      '-Confirm:$false ' +
      '-ErrorAction SilentlyContinue"',
      '',
      SW_HIDE,
      ewWaitUntilTerminated,
      ResultCode
    );

  end;

end;