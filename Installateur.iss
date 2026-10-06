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

[Icons]
Name: "{autoprograms}\Gestion des relevés bancaires\Notice"; \
    Filename: "{app}\programme\Notice.pdf"

[Run]
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; \
    Parameters: "{code:GetTaskParameters}"; \
    Flags: runhidden waituntilterminated

[UninstallRun]
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; \
    Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""Unregister-ScheduledTask -TaskName 'Faisons les comptes' -Confirm:$false -ErrorAction SilentlyContinue"""; \
    Flags: runhidden waituntilterminated
    
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
  XmlPath: String;
  Command: String;
  NombreRemplacements: Integer;
begin
  ScriptPath := ExpandConstant('{app}\programme\TraitementReleve.ps1');
  XmlPath := ExpandConstant('{app}\programme\Tache_Faisons_les_comptes.xml');

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

    'try { ' +
      'Register-ScheduledTask ' +
      '-TaskName ''Faisons les comptes'' ' +
      '-Action $action ' +
      '-Trigger $trigger ' +
      '-Principal $principal ' +
      '-Settings $settings ' +
      '-Description ''Le but est de penser à enregistrer les comptes régulièrement'' ' +
      '-Force ' +
      '-ErrorAction Stop; ' +
    '} ' +
    'catch { ' +

      '$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name; ' +
      '$escapedUser = [System.Security.SecurityElement]::Escape($user); ' +
      '$escapedScript = [System.Security.SecurityElement]::Escape($env:COMPUTERNAME); ' +
      '$escapedPath = [System.Security.SecurityElement]::Escape(''' + ScriptPath + '''); ' +

      '$start = (Get-Date).Date.AddDays(1).AddHours(8).ToString(''yyyy-MM-ddTHH:mm:ss''); ' +

      '$xml = ''<?xml version="1.0" encoding="UTF-16"?>'' + ' +
      '''<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">'' + ' +
      '''<RegistrationInfo><Author>'' + $escapedUser + ''</Author>'' + ' +
      '''<Description>Le but est de penser à enregistrer les comptes régulièrement</Description></RegistrationInfo>'' + ' +
      '''<Triggers><CalendarTrigger>'' + ' +
      '''<StartBoundary>'' + $start + ''</StartBoundary>'' + ' +
      '''<Enabled>true</Enabled>'' + ' +
      '''<ScheduleByDay><DaysInterval>'' + ' + IntToStr(IntervalleJours) + ' + ''</DaysInterval></ScheduleByDay>'' + ' +
      '''</CalendarTrigger></Triggers>'' + ' +
      '''<Principals><Principal id="Author">'' + ' +
      '''<UserId>'' + $escapedUser + ''</UserId>'' + ' +
      '''<LogonType>InteractiveToken</LogonType>'' + ' +
      '''<RunLevel>LeastPrivilege</RunLevel>'' + ' +
      '''</Principal></Principals>'' + ' +
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
      '''<Actions Context="Author"><Exec>'' + ' +
      '''<Command>powershell.exe</Command>'' + ' +
      '''<Arguments>-NoProfile -ExecutionPolicy Bypass -File &quot;'' + $escapedPath + ''&quot;</Arguments>'' + ' +
      '''<WorkingDirectory>'' + [System.Security.SecurityElement]::Escape((Split-Path ''' + ScriptPath + ''')) + ''</WorkingDirectory>'' + ' +
      '''</Exec></Actions>'' + ' +
      '''</Task>''; ' +

      '[System.IO.File]::WriteAllText(''' + XmlPath + ''', $xml, [System.Text.Encoding]::Unicode); ' +

      'Write-Host ''Impossible de créer automatiquement la tâche.''; ' +
      'Write-Host ''Le fichier XML a été créé : ' + XmlPath + '''; ' +
      'exit 1; ' +
    '}';

  NombreRemplacements := StringChangeEx(Command, '"', '""', True);

  Result :=
    '-NoProfile -ExecutionPolicy Bypass -Command "' +
    Command +
    '"';
end;