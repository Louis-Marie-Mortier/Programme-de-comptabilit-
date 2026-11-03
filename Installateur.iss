#define MyAppName "Gestion des relevés bancaires"
#define MyAppVersion "1.1"
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
Source: "AjouterTransactionsOds.py"; DestDir: "{app}\programme"; Flags: ignoreversion
Source: "Modele.ods"; DestDir: "{app}\programme"; Flags: ignoreversion
Source: "Notice.pdf"; DestDir: "{app}\programme"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Gestion des relevés bancaires\Notice"; \
    Filename: "{app}\programme\Notice.pdf"
    
[Code]
var
  PageFrequence: TInputQueryWizardPage;
  IntervalleJours: Integer;


procedure InitializeWizard();
begin
  PageFrequence := CreateInputQueryPage(
    wpSelectDir,
    'Fréquence de rappel',
    'Choisissez la fréquence de rappel',
    'Le programme vous rappellera de télécharger votre relevé bancaire à cette fréquence.'
  );

  PageFrequence.Add('Nombre de jours entre deux rappels :', False);
  PageFrequence.Values[0] := '83';
end;


function NextButtonClick(CurPageID: Integer): Boolean;
var
  Valeur: Integer;
begin
  Result := True;

  if CurPageID = PageFrequence.ID then
  begin
    Valeur := StrToIntDef(Trim(PageFrequence.Values[0]), 0);

    if (Valeur < 1) or (Valeur > 90) then
    begin
      MsgBox(
        'Veuillez saisir un nombre compris entre 1 et 90 jours.',
        mbError,
        MB_OK
      );

      Result := False;
      Exit;
    end;

    IntervalleJours := Valeur;
  end;
end;


function GetTaskParameters(Param: String): String;
var
  ScriptPath: String;
  Command: String;
  NombreRemplacements: Integer;
begin
  ScriptPath := ExpandConstant('{app}\programme\TraitementReleve.ps1');

  Command :=
    '$ErrorActionPreference = ''Stop''; ' +

    '$action = New-ScheduledTaskAction ' +
    '-Execute ''powershell.exe'' ' +
    '-Argument ''-NoProfile -ExecutionPolicy Bypass -File "' + ScriptPath + '"''; ' +

    '$trigger = New-ScheduledTaskTrigger ' +
    '-Daily ' +
    '-DaysInterval ' + IntToStr(IntervalleJours) + ' ' +
    '-At ''08:00''; ' +

    '$principal = New-ScheduledTaskPrincipal ' +
    '-UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) ' +
    '-LogonType Interactive ' +
    '-RunLevel Limited; ' +

    '$settings = New-ScheduledTaskSettingsSet ' +
    '-StartWhenAvailable ' +
    '-AllowStartIfOnBatteries ' +
    '-DontStopIfGoingOnBatteries ' +
    '-ExecutionTimeLimit ([TimeSpan]::FromHours(1)) ' +
    '-MultipleInstances IgnoreNew ' +
    '-Priority 7; ' +

    'Register-ScheduledTask ' +
    '-TaskName ''Faisons les comptes'' ' +
    '-Action $action ' +
    '-Trigger $trigger ' +
    '-Principal $principal ' +
    '-Settings $settings ' +
    '-Description ''Le but est de penser à enregistrer les comptes régulièrement'' ' +
    '-Force ' +
    '-ErrorAction Stop';

  NombreRemplacements := StringChangeEx(Command, '"', '""', True);

  Result :=
    '-NoProfile -ExecutionPolicy Bypass -Command "' +
    Command +
    '"';
end;


function GetXmlParameters(Param: String): String;
var
  ScriptPath: String;
  XmlPath: String;
  Command: String;
  NombreRemplacements: Integer;
begin
  ScriptPath := ExpandConstant('{app}\programme\TraitementReleve.ps1');
  XmlPath := ExpandConstant('{app}\programme\Tache_Faisons_les_comptes.xml');

  Command :=
    '$ErrorActionPreference = ''Stop''; ' +
    '$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name; ' +
    '$escapedUser = [System.Security.SecurityElement]::Escape($user); ' +
    '$escapedPath = [System.Security.SecurityElement]::Escape(''' +
      ScriptPath +
    '''); ' +
    '$workingDirectory = [System.IO.Path]::GetDirectoryName(''' +
      ScriptPath +
    '''); ' +
    '$escapedWorkingDirectory = [System.Security.SecurityElement]::Escape($workingDirectory); ' +
    '$start = (Get-Date).Date.AddHours(8).ToString(''yyyy-MM-ddTHH:mm:ss''); ' +

    '$xml = ''<?xml version=''''1.0'''' encoding=''''UTF-16''''?>'' + ' +
    '''<Task version=''''1.4'''' xmlns=''''http://schemas.microsoft.com/windows/2004/02/mit/task''''>'' + ' +

    '''<RegistrationInfo>'' + ' +
    '''<Author>'' + $escapedUser + ''</Author>'' + ' +
    '''<Description>Le but est de penser à enregistrer les comptes régulièrement</Description>'' + ' +
    '''</RegistrationInfo>'' + ' +

    '''<Triggers>'' + ' +
    '''<CalendarTrigger>'' + ' +
    '''<StartBoundary>'' + $start + ''</StartBoundary>'' + ' +
    '''<Enabled>true</Enabled>'' + ' +
    '''<ScheduleByDay>'' + ' +
    '''<DaysInterval>'' + ' + IntToStr(IntervalleJours) + ' + ''</DaysInterval>'' + ' +
    '''</ScheduleByDay>'' + ' +
    '''</CalendarTrigger>'' + ' +
    '''</Triggers>'' + ' +

    '''<Principals>'' + ' +
    '''<Principal id=''''Author''''>'' + ' +
    '''<UserId>'' + $escapedUser + ''</UserId>'' + ' +
    '''<LogonType>InteractiveToken</LogonType>'' + ' +
    '''<RunLevel>LeastPrivilege</RunLevel>'' + ' +
    '''</Principal>'' + ' +
    '''</Principals>'' + ' +

    '''<Settings>'' + ' +
    '''<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>'' + ' +
    '''<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>'' + ' +
    '''<StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>'' + ' +
    '''<StartWhenAvailable>true</StartWhenAvailable>'' + ' +
    '''<AllowHardTerminate>true</AllowHardTerminate>'' + ' +
    '''<ExecutionTimeLimit>PT1H</ExecutionTimeLimit>'' + ' +
    '''<Priority>7</Priority>'' + ' +
    '''<Enabled>true</Enabled>'' + ' +
    '''<Hidden>false</Hidden>'' + ' +
    '''<RunOnlyIfIdle>false</RunOnlyIfIdle>'' + ' +
    '''<WakeToRun>false</WakeToRun>'' + ' +
    '''</Settings>'' + ' +

    '''<Actions Context=''''Author''''>'' + ' +
    '''<Exec>'' + ' +
    '''<Command>powershell.exe</Command>'' + ' +
    '''<Arguments>-NoProfile -ExecutionPolicy Bypass -File &quot;'' + $escapedPath + ''&quot;</Arguments>'' + ' +
    '''<WorkingDirectory>'' + $escapedWorkingDirectory + ''</WorkingDirectory>'' + ' +
    '''</Exec>'' + ' +
    '''</Actions>'' + ' +

    '''</Task>''; ' +

    '[System.IO.File]::WriteAllText(''' +
      XmlPath +
    ''', $xml, [System.Text.Encoding]::Unicode)';

  NombreRemplacements := StringChangeEx(Command, '"', '""', True);

  Result :=
    '-NoProfile -ExecutionPolicy Bypass -Command "' +
    Command +
    '"';
end;


procedure CreateScheduledTask;
var
  ResultCode: Integer;
  Success: Boolean;
  PowerShellPath: String;
begin
  PowerShellPath :=
    ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');

  Success := ShellExec(
    'runas',
    PowerShellPath,
    GetTaskParameters(''),
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  );

  if (not Success) or (ResultCode <> 0) then
  begin
    if Exec(
      PowerShellPath,
      GetXmlParameters(''),
      '',
      SW_HIDE,
      ewWaitUntilTerminated,
      ResultCode
    ) then
    begin
      MsgBox(
        'La création automatique de la tâche planifiée n''a pas pu être effectuée.' +
        Chr(13) + Chr(10) + Chr(13) + Chr(10) +
        'Un fichier XML permettant de créer la tâche manuellement a été créé dans :' +
        Chr(13) + Chr(10) +
        ExpandConstant('{app}\programme\Tache_Faisons_les_comptes.xml'),
        mbInformation,
        MB_OK
      );
    end
    else
    begin
      MsgBox(
        'Impossible de créer la tâche planifiée automatiquement et impossible de créer le fichier XML de secours.',
        mbError,
        MB_OK
      );
    end;
  end;
end;


procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    CreateScheduledTask;
end;

procedure SupprimerTachePlanifiee();
var
  ResultCode: Integer;
  Parametres: String;
begin
  Parametres :=
    '-NoProfile -ExecutionPolicy Bypass -Command ' +
    '"Unregister-ScheduledTask -TaskName ''Faisons les comptes'' -Confirm:$false -ErrorAction Stop"';

  if ShellExec(
    'runas',
    ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    Parametres,
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  ) then
  begin
    if ResultCode <> 0 then
    begin
      MsgBox(
        'La tâche planifiée « Faisons les comptes » n''a pas pu être supprimée automatiquement.' +
        Chr(13) + Chr(10) + Chr(13) + Chr(10) +
        'Vous pourrez la supprimer manuellement depuis le Planificateur de tâches.',
        mbInformation,
        MB_OK
      );
    end;
  end
  else
  begin
    MsgBox(
      'La tâche planifiée « Faisons les comptes » n''a pas pu être supprimée automatiquement.' +
      Chr(13) + Chr(10) + Chr(13) + Chr(10) +
      'Vous pourrez la supprimer manuellement depuis le Planificateur de tâches.',
      mbInformation,
      MB_OK
    );
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    SupprimerTachePlanifiee();
end;