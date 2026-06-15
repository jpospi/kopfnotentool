#define MyAppName "KopfnotenTool"
#define MyAppExeName "KopfnotenTool.exe"
#define MyAppVersion "1.0.0"

[Setup]
AppId={{A83266D9-97C8-4A6D-BA81-2C188F9099E9}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=KopfnotenTool-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Desktop-Verknüpfung erstellen"; GroupDescription: "Zusätzliche Symbole:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "KopfnotenTool starten"; Flags: nowait postinstall skipifsilent

[Code]
var
  DataRootPage: TInputDirWizardPage;
  ImportDirPage: TInputDirWizardPage;
  OutputWordDirPage: TInputDirWizardPage;
  OutputExcelDirPage: TInputDirWizardPage;
  DbDirPage: TInputDirWizardPage;
  BackupDirPage: TInputDirWizardPage;

function JsonEscape(Value: String): String;
begin
  Result := Value;
  StringChangeEx(Result, '\', '\\', True);
  StringChangeEx(Result, '"', '\"', True);
end;

procedure InitializeWizard;
var
  DefaultRoot: String;
begin
  DefaultRoot := ExpandConstant('{userdocs}\KopfnotenToolData');

  DataRootPage := CreateInputDirPage(
    wpSelectDir,
    'Datenpfad',
    'Bitte Speicherort fuer Anwendungsdaten festlegen',
    'Hier werden temp, logs, templates und Standarddaten gespeichert.',
    False,
    ''
  );
  DataRootPage.Add('');
  DataRootPage.Values[0] := DefaultRoot;

  ImportDirPage := CreateInputDirPage(
    DataRootPage.ID,
    'Importpfad',
    'Standardpfad fuer Excel-Import',
    'Dieser Pfad wird als Startordner beim lokalen Excel-Import verwendet.',
    False,
    ''
  );
  ImportDirPage.Add('');
  ImportDirPage.Values[0] := DataRootPage.Values[0] + '\input_excel';

  OutputWordDirPage := CreateInputDirPage(
    ImportDirPage.ID,
    'Word-Exportpfad',
    'Standardpfad fuer Word-Export',
    'Hier werden erzeugte Serienbrief-Dateien abgelegt.',
    False,
    ''
  );
  OutputWordDirPage.Add('');
  OutputWordDirPage.Values[0] := DataRootPage.Values[0] + '\output_word';

  OutputExcelDirPage := CreateInputDirPage(
    OutputWordDirPage.ID,
    'Excel-Exportpfad',
    'Standardpfad fuer Excel-Auswertungen',
    'Hier werden Excel-Listen und Auswertungen gespeichert.',
    False,
    ''
  );
  OutputExcelDirPage.Add('');
  OutputExcelDirPage.Values[0] := DataRootPage.Values[0] + '\output_excel';

  DbDirPage := CreateInputDirPage(
    OutputExcelDirPage.ID,
    'Datenbankpfad',
    'Ordner fuer die aktive Datenbank',
    'Die Datenbank-Datei heisst: kopfnoten_secure.db',
    False,
    ''
  );
  DbDirPage.Add('');
  DbDirPage.Values[0] := DataRootPage.Values[0] + '\output_database';

  BackupDirPage := CreateInputDirPage(
    DbDirPage.ID,
    'Datenbank-Sicherung',
    'Ordner fuer Datenbank-Backups',
    'Hierhin werden Sicherungskopien ueber das Menue "Datenbank sichern" geschrieben.',
    False,
    ''
  );
  BackupDirPage.Add('');
  BackupDirPage.Values[0] := DataRootPage.Values[0] + '\db_backup';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigPath: String;
  DataRoot, ImportDir, OutputWordDir, OutputExcelDir, DbDir, BackupDir: String;
  JsonContent: String;
begin
  if CurStep = ssPostInstall then
  begin
    DataRoot := DataRootPage.Values[0];
    ImportDir := ImportDirPage.Values[0];
    OutputWordDir := OutputWordDirPage.Values[0];
    OutputExcelDir := OutputExcelDirPage.Values[0];
    DbDir := DbDirPage.Values[0];
    BackupDir := BackupDirPage.Values[0];

    ForceDirectories(DataRoot);
    ForceDirectories(ImportDir);
    ForceDirectories(OutputWordDir);
    ForceDirectories(OutputExcelDir);
    ForceDirectories(DbDir);
    ForceDirectories(BackupDir);
    ForceDirectories(AddBackslash(DataRoot) + 'logs');
    ForceDirectories(AddBackslash(DataRoot) + 'temp');
    ForceDirectories(AddBackslash(DataRoot) + 'templates');

    JsonContent :=
      '{'#13#10 +
      '  "data_root": "' + JsonEscape(DataRoot) + '",'#13#10 +
      '  "import_dir": "' + JsonEscape(ImportDir) + '",'#13#10 +
      '  "output_word_dir": "' + JsonEscape(OutputWordDir) + '",'#13#10 +
      '  "output_excel_dir": "' + JsonEscape(OutputExcelDir) + '",'#13#10 +
      '  "templates_dir": "' + JsonEscape(AddBackslash(DataRoot) + 'templates') + '",'#13#10 +
      '  "logs_dir": "' + JsonEscape(AddBackslash(DataRoot) + 'logs') + '",'#13#10 +
      '  "temp_dir": "' + JsonEscape(AddBackslash(DataRoot) + 'temp') + '",'#13#10 +
      '  "database_path": "' + JsonEscape(AddBackslash(DbDir) + 'kopfnoten_secure.db') + '",'#13#10 +
      '  "backup_dir": "' + JsonEscape(BackupDir) + '",'#13#10 +
      '  "sph_config_path": "' + JsonEscape(AddBackslash(DataRoot) + 'sph_config.json') + '"'#13#10 +
      '}'#13#10;

    ConfigPath := ExpandConstant('{app}\kopfnotentool.paths.json');
    SaveStringToFile(ConfigPath, JsonContent, False);
  end;
end;
