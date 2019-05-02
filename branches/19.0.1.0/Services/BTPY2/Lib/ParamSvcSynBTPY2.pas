unit ParamSvcSynBTPY2;

interface

uses
  Windows
  , Forms
  , Classes
  , Controls
  , StdCtrls
  , Buttons
  , TntButtons
  , Hctrls
  , Spin
  , Mask
  , ExtCtrls
  , ComCtrls
  , TntComCtrls
  , Grids
  , TntGrids
  , CommonTools
  ;

type
  TParamSvcSyncrho = class(TForm)
    Tabs: THPageControl2;
    TSParamGeneraux: TTabSheet;
    TSUpdateFrequency: TTabSheet;
    lSecondTimeout: TLabel;
    lSecondTimeoutUnit: TLabel;
    SecondTimeout: TSpinEdit;
    glExecutionPeriodDays: TGroupBox;
    Lundi: TCheckBox;
    Mardi: TCheckBox;
    Mercredi: TCheckBox;
    Jeudi: TCheckBox;
    Vendredi: TCheckBox;
    Samedi: TCheckBox;
    Dimanche: TCheckBox;
    lExecutionPeriodHours: TGroupBox;
    lExecutionPeriodStart: TLabel;
    lExecutionPeriodEnd: TLabel;
    ExecutionPeriodStart: TMaskEdit;
    ExecutionPeriodEnd: TMaskEdit;
    LogLevel: TRadioGroup;
    LogType: TRadioGroup;
    pTitle: TPanel;
    LGenerateIn: TLabel;
    GenerateIn: TEdit;
    Reload: THBitBtn;
    pButtons: TPanel;
    bValider: THBitBtn;
    bFermer: THBitBtn;
    tFolders: TTabSheet;
    gExchange: TGroupBox;
    Download: TCheckBox;
    Upload: TCheckBox;
    lDebugFilesDirectory: TLabel;
    gFrequency1: THGrid;
    LogMaxSize: TSpinEdit;
    DebugFilesDirectory: THCritMaskEdit;
    lFolderChoice: TLabel;
    NewFolder: THBitBtn;
    DelFolder: THBitBtn;
    cFolderChoice: TListBox;
    gFolderParam: TGroupBox;
    LastSynchro: TMaskEdit;
    lLastSynchro: TLabel;
    lBTPUser: TLabel;
    BTPUser: TMaskEdit;
    lBTPServer: TLabel;
    lBTP: TLabel;
    BTPServer: TMaskEdit;
    lBTPFolder: TLabel;
    BTPFolder: TMaskEdit;
    lY2Server: TLabel;
    lY2: TLabel;
    Y2Server: TMaskEdit;
    lY2Folder: TLabel;
    Y2Folder: TMaskEdit;
    AddNewFolder: THBitBtn;
    CancelNewFolder: THBitBtn;
    UpdateFolder: THBitBtn;
    tService: TTabSheet;
    gSvcState: TGroupBox;
    SvcStateRefresh: THBitBtn;
    ServiceState: TLabel;
    gSvcManagement: TGroupBox;
    SvcInstall: THBitBtn;
    SvcUnInstall: THBitBtn;
    SvcStart: THBitBtn;
    SvcStop: THBitBtn;
    gSvcInstalParam: TGroupBox;
    SvcInstallType: TRadioGroup;
    gSvcInstallConnection: TGroupBox;
    SvcInstallAccount: TEdit;
    lInstallAccount: TLabel;
    SvcInstallPwd: TEdit;
    lInstallPwd: TLabel;
    gSvcInstallPath: TGroupBox;
    SvcInstallPath: THCritMaskEdit;
    SvcInstallPwdConfirm: TEdit;
    lInstallPwdConfirm: TLabel;
    SvcAccountTipe: TRadioGroup;
    LogLevelDebug: TRadioGroup;
    ConnectBTP: THBitBtn;
    ConnectY2: THBitBtn;
    LogSee: TGroupBox;
    SeeLogs: THBitBtn;

    procedure FormShow(Sender: TObject);
    procedure bValiderClick(Sender: TObject);
    procedure bFermerClick(Sender: TObject);
    procedure ReloadClick(Sender: TObject);
    procedure LogLevelClick(Sender: TObject);
    procedure LogTypeClick(Sender: TObject);
    procedure UploadClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure gFrequency1DblClick(Sender: TObject);
    procedure cFolderChoiceClick(Sender: TObject);
    procedure NewFolderClick(Sender: TObject);
    procedure DelFolderClick(Sender: TObject);
    procedure AddNewFolderClick(Sender: TObject);
    procedure CancelNewFolderClick(Sender: TObject);
    procedure UpdateFolderClick(Sender: TObject);
    procedure SecondTimeoutExit(Sender: TObject);
    procedure SecondTimeoutClick(Sender: TObject);
    procedure SecondTimeoutKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SecondTimeoutKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ExecutionPeriodStartExit(Sender: TObject);
    procedure ExecutionPeriodEndExit(Sender: TObject);
    procedure SvcInstallClick(Sender: TObject);
    procedure SvcUnInstallClick(Sender: TObject);
    procedure SvcStartClick(Sender: TObject);
    procedure SvcStopClick(Sender: TObject);
    procedure SvcStateRefreshClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SvcAccountTipeClick(Sender: TObject);
    procedure gFrequency1MouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure TestConnectClick(Sender: TObject);
    procedure SeeLogsClick(Sender: TObject);
  private
    PathFileName  : string;
    AlreadyExist  : boolean;
    OnNewFolder   : boolean;
    TslTableList  : TStringList;
    TslFolderList : TStringList;

    function GetIniName : string;
    procedure ShowTables;
    procedure ShowFolders(pItemIndex : integer);
    procedure LoadIniIfExist;
    function SetMnTimeOut : string;
    procedure ShowLogType;
    procedure ShowMaxSizeType;
    procedure ShowDebugFilesDirectory;
    procedure ManageFolderParams(pEnabled : boolean);
    procedure ManageOtherControls(pEnabled : boolean);
    procedure ClearFolderParams;
    procedure TimeSlotIsValid(IsStart : boolean; pValue : string);
    function GetServiceState : tServiceState;
    function GetSvcPathName : string;
    procedure SetCaptionServiceState;
    procedure ServiceButtonManagement;

  public
    { Déclarations publiques }
  end;

var
  ParamSvcSyncrho: TParamSvcSyncrho;

implementation

uses
  Messages
  , SysUtils
  , Variants
  , Graphics
  , Dialogs
  , ConstServices
  , IniFiles
  , HmsgBox
  , ed_Tools
  , AdoDB
  , WinSvc
  , ShellApi
  ;

const
  cWarningMsg     = 'ATTENTION, les valeurs non enregistrées seront perdues.';
  cServiceName    = 'SvcSyncBTPY2';
  cExeServiceName = 'SvcSynBTPY2';

{$R *.dfm}
procedure TParamSvcSyncrho.bValiderClick(Sender: TObject);
var
  FileNameBackup : string;
  Section        : string;
  Id             : string;
  Value          : string;
  MsgConfirm     : string;
  Cpt            : integer;
  CanContinue    : boolean;
  IsSvcRunning   : boolean;
  SettingFile    : TInifile;

  function GetExecutionPeriodDays : string;
  begin
    Result := '';
    if Lundi.Checked    then Result := Format('%s,1', [Result]);
    if Mardi.Checked    then Result := Format('%s,2', [Result]);
    if Mercredi.Checked then Result := Format('%s,3', [Result]);
    if Jeudi.Checked    then Result := Format('%s,4', [Result]);
    if Vendredi.Checked then Result := Format('%s,5', [Result]);
    if Samedi.Checked   then Result := Format('%s,6', [Result]);
    if Dimanche.Checked then Result := Format('%s,7', [Result]);
    Result := copy(Result, 2, length(Result));
  end;

begin
  if TslFolderList.count = 0 then
  begin
    PGIError('Il n''y a pas pas de dossier(s) paramétré(s).');
  end else
  begin
    MsgConfirm     := '';
    FileNameBackup := ExtractFilePath(PathFileName) + FormatDateTime('yyyymmdd', Now) + GetIniName;
    IsSvcRunning   := (GetServiceState <> tssUnInstall);
    if AlreadyExist then MsgConfirm := Format('-> Le fichier existe déjà et sera sauvegardé sous le nom %s.%s', [FileNameBackup, #13#10]);
    if IsSvcRunning then MsgConfirm := Format('%s-> Le service doit être redémarré après validation pour prendre en compte les modifications.%s', [MsgConfirm, #13#10]);
    if PGIAsk(Format('Veuillez confirmer la génération de %s avec le paramétrage défini ?%s%s%s', [PathFileName, #13#10, #13#10, MsgConfirm]), 'Confirmation') = mrYes then
    begin
      { Sauvegarde de l'ancien paramétrage si existe }
      CanContinue := True;
      if AlreadyExist then
      begin
        if FileExists(FileNameBackup) then
          DeleteFile(FileNameBackup);
        if not RenameFile(PathFileName, FileNameBackup) then
          CanContinue := (PGIAsk(Format('Erreur lors de la sauvegarde de l''ancien fichier.%sVoulez-vous continuer (l''ancien paramétrage ne sera pas sauvegardé) ?', [#13#10])) = mrYes);
      end;
      if CanContinue then
      begin
        SettingFile := TIniFile.Create(PathFileName);
        try
          Section := 'GLOBALSETTINGS';
          SettingFile.WriteString(Section, 'SecondTimeout'       , SecondTimeout.Text);
          SettingFile.WriteString(Section, 'ExecutionPeriodDays' , GetExecutionPeriodDays);
          SettingFile.WriteString(Section, 'ExecutionPeriodStart', ExecutionPeriodStart.Text);
          SettingFile.WriteString(Section, 'ExecutionPeriodEnd'  , ExecutionPeriodEnd.Text);
          SettingFile.WriteString(Section, 'LogLevel'            , IntToStr(LogLevel.ItemIndex));
          SettingFile.WriteString(Section, 'OneLogPerDay'        , Tools.iif(LogType.ItemIndex=0, '1', '0'));
          SettingFile.WriteString(Section, 'LogMoMaxSize'        , IntToStr(LogMaxSize.Value));
          SettingFile.WriteString(Section, 'DebugEvents'         , IntToStr(LogLevelDebug.ItemIndex));
          SettingFile.WriteString(Section, 'DebugFilesDirectory' , DebugFilesDirectory.Text);
          SettingFile.WriteString(Section, 'Download'            , Tools.iif(Download.Checked, '1', '0'));
          SettingFile.WriteString(Section, 'Upload'              , Tools.iif(Upload.Checked, '1', '0'));
          Section := 'SERVICESETTINGS';
          SettingFile.WriteString(Section, 'SvcInstallPath'      , SvcInstallPath.Text);
          SettingFile.WriteString(Section, 'SvcInstallType'      , IntToStr(SvcInstallType.ItemIndex));
          SettingFile.WriteString(Section, 'SvcAccountTipe'      , IntToStr(SvcAccountTipe.ItemIndex));
          SettingFile.WriteString(Section, 'SvcInstallAccount'   , SvcInstallAccount.Text);
          Section := 'UPDATEFREQUENCY';
          for Cpt := 0 to pred(TslTableList.count) do
          begin
            Value := TslTableList[Cpt];
            SettingFile.WriteString(Section, copy(Value, 1, pos('=', Value) -1), copy(Value, pos('=', Value) +1, length(Value)));
          end;
          for Cpt := 0 to pred(TslFolderList.count) do
          begin
            Value   := TslFolderList[Cpt];
            Section := copy(Value, 1, pos('=', Value)-1);
            Value   := copy(Value, pos('=', Value)+1, length(Value));
            Id      := Tools.ReadTokenSt_(Value, '|');
            SettingFile.WriteString(Section, Id, Value);
          end;
        finally
          Reload.Visible := True;
          AlreadyExist   := True;
          LoadIniIfExist;
          SettingFile.Free;
        end;
      end;
    end;
  end;
end;

procedure TParamSvcSyncrho.bFermerClick(Sender: TObject);
begin
  if PGIAsk(Format('Voulez-vous quitter ?%s%s', [#13#10, cWarningMsg])) = mrYes then
    Application.Terminate;
end;

procedure TParamSvcSyncrho.FormCreate(Sender: TObject);
begin
  ParamSvcSyncrho.Position := poScreenCenter;
end;

procedure TParamSvcSyncrho.FormShow(Sender: TObject);
begin
  DeleteMenu(GetSystemMenu(Handle,False),SC_CLOSE,MF_BYCOMMAND);
  GenerateIn.Text        := TServicesLog.GetServicesAppDataPath(True);
  PathFileName           := Format('%s\%s', [GenerateIn.Text, GetIniName]);
  AlreadyExist           := (FileExists(PathFileName));
  Reload.Visible         := AlreadyExist;
  TslTableList           := TStringList.Create;
  TslFolderList          := TStringList.Create;
  OnNewFolder            := False;
  Tabs.ActivePageIndex   := 0;
  SetCaptionServiceState;
  ServiceButtonManagement;
  LoadIniIfExist;
  SetMnTimeOut;
  ShowLogType;
  if not IsUserAnAdmin then
  begin
    ServiceState.Left       := 9;
    ServiceState.Caption    := 'L''application doit être exécutée en mode "Administrateur" pour pouvoir gérer le service.';
    SvcStateRefresh.Visible := False;
    gSvcManagement.Visible  := False;
  end else
    gSvcState.Caption := Format('  Etat du service %s  ', [cServiceName]);
end;

procedure TParamSvcSyncrho.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := False;
end;

procedure TParamSvcSyncrho.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(TslTableList);
  FreeAndNil(TslFolderList);
end;

function TParamSvcSyncrho.GetIniName : string;
begin
  Result := 'SvcSynBTPY2.ini';
end;

procedure TParamSvcSyncrho.ShowTables;
var
  Cpt       : integer;
  CptCol    : integer;
  CheckCol  : integer;
  Value     : string;
  TableName : string;
begin
  if TslTableList.count > 0 then
  begin
    gFrequency1.RowCount := TslTableList.count + 1;
    for Cpt := 0 to 7 do
    begin
           if Cpt = 0 then gFrequency1.ColWidths[Cpt] := 10
      else if Cpt = 1 then gFrequency1.ColWidths[Cpt] := 100
                      else gFrequency1.ColWidths[Cpt] := 85;
      if Cpt >= 2 then gFrequency1.ColAligns[Cpt] := taCenter;
    end;
    for Cpt := 0 to pred(TslTableList.count) do
    begin
      Value     := TslTableList[Cpt];
      TableName := copy(Value, 1, pos('=', Value)-1);
      Value     := copy(Value, pos('=', Value)+1, length(Value));
      gFrequency1.Cells[1, Cpt+1] := TableName;
      case Tools.CaseFromString(Value, ['Everytime', 'Daily', 'Monthly', 'Annual', 'Once', 'Never']) of
        {Everytime} 0 : CheckCol := 2;
        {Daily}     1 : CheckCol := 3;
        {Monthly}   2 : CheckCol := 4;
        {Annual}    3 : CheckCol := 5;
        {Once}      4 : CheckCol := 6;
        {Never}     5 : CheckCol := 7;
      else
        CheckCol := -1;
      end;
      for CptCol := 2 to 7 do
        gFrequency1.Cells[CptCol, Cpt+1] := Tools.iif(CptCol = CheckCol, 'X', '');
    end;
  end;
end;

procedure TParamSvcSyncrho.ShowFolders(pItemIndex : integer);
var
  LastFolder    : string;
  CurrentFolder : string;
  Value         : string;
  Cpt           : integer;
begin
  cFolderChoice.Items.Clear;
  LastFolder := '';
  ClearFolderParams;
  for Cpt := 0 to pred(TslFolderList.count) do
  begin
    Value := TslFolderList[Cpt];
    CurrentFolder := copy(Value, 1 ,pos('=', Value)-1);
    if LastFolder <> CurrentFolder then
    begin
      LastFolder := CurrentFolder;
      cFolderChoice.AddItem(' ' + CurrentFolder, nil);
    end;
  end;
  cFolderChoice.ItemIndex := pItemIndex;
  cFolderChoiceClick(nil);  
end;

procedure TParamSvcSyncrho.LoadIniIfExist;
var
  SettingFile         : TInifile;
  Section             : string;
  ExecutionPeriodDays : string;
  iReadInteger        : integer;
  Cpt                 : integer;
  SectionExist        : boolean;

  function OkDays(Number : integer) : boolean;
  begin
    Result := (pos(IntToStr(Number), ExecutionPeriodDays) > 0);
  end;

  procedure AddTableList(TableName, DefaultValue : string);
  begin
    TslTableList.Add(Format('%s=%s', [TableName, Tools.iif(AlreadyExist, SettingFile.ReadString(Section, TableName, DefaultValue), DefaultValue)]));
  end;

  procedure AddFolderList(Exist : boolean; ParamName, DefaultValue : string);
  begin
    TslFolderList.Add(Format('%s=%s|%s', [Section, ParamName, Tools.iif(Exist, SettingFile.ReadString(Section , ParamName, DefaultValue), DefaultValue)]));
  end;

begin
  SettingFile := TIniFile.Create(Tools.iif(AlreadyExist, PathFileName, ''));
  try
    TslTableList.Clear;
    TslFolderList.Clear;
    Section := 'GLOBALSETTINGS';
    SecondTimeout.Value       := Tools.iif(AlreadyExist, SettingFile.ReadInteger(Section, 'SecondTimeout'      , 900), 900);
    ExecutionPeriodDays       := Tools.iif(AlreadyExist, SettingFile.ReadString(Section , 'ExecutionPeriodDays', ''), '');
    Lundi.Checked             := Tools.iif(AlreadyExist, OkDays(1), True);
    Mardi.Checked             := Tools.iif(AlreadyExist, OkDays(2), True);
    Mercredi.Checked          := Tools.iif(AlreadyExist, OkDays(3), True);
    Jeudi.Checked             := Tools.iif(AlreadyExist, OkDays(4), True);
    Vendredi.Checked          := Tools.iif(AlreadyExist, OkDays(5), True);
    Samedi.Checked            := Tools.iif(AlreadyExist, OkDays(6), True);
    Dimanche.Checked          := Tools.iif(AlreadyExist, OkDays(7), True);
    ExecutionPeriodStart.Text := Tools.iif(AlreadyExist, SettingFile.ReadString(Section , 'ExecutionPeriodStart', '00:00:01'), '00:00:01');
    ExecutionPeriodEnd.Text   := Tools.iif(AlreadyExist, SettingFile.ReadString(Section , 'ExecutionPeriodEnd'  , '23:59:59'), '23:59:59');
    LogLevel.ItemIndex        := Tools.iif(AlreadyExist, SettingFile.ReadInteger(Section, 'LogLevel', 1), 1);
    LogLevelDebug.ItemIndex   := Tools.iif(AlreadyExist, SettingFile.ReadInteger(Section, 'DebugEvents', 1), 1);
    iReadInteger              := Tools.iif(AlreadyExist, SettingFile.ReadInteger(Section, 'OneLogPerDay', 0), 0);
    LogType.ItemIndex         := Tools.iif(iReadInteger=1, 0, 1);
    LogMaxSize.Value          := Tools.iif(AlreadyExist, SettingFile.ReadInteger(Section , 'LogMoMaxSize', 0), 0);
    Download.Checked          := Tools.iif(AlreadyExist, (SettingFile.ReadInteger(Section, 'Download', 0) = 1), True);
    UpLoad.Checked            := Tools.iif(AlreadyExist, (SettingFile.ReadInteger(Section, 'Upload', 0) = 1), True);
    DebugFilesDirectory.Text  := Tools.iif(AlreadyExist, SettingFile.ReadString(Section  , 'DebugFilesDirectory', ''), '');
    Section := 'SERVICESETTINGS';
    SvcInstallPath.Text       := Tools.iif(AlreadyExist, SettingFile.ReadString(Section  , 'SvcInstallPath', ''), '');
    SvcInstallType.ItemIndex  := Tools.iif(AlreadyExist, SettingFile.ReadInteger(Section , 'SvcInstallType', 0), 0);
    SvcAccountTipe.ItemIndex  := Tools.iif(AlreadyExist, SettingFile.ReadInteger(Section , 'SvcAccountTipe', 0), 0);
    SvcInstallAccount.Text    := Tools.iif(AlreadyExist, SettingFile.ReadString(Section  , 'SvcInstallAccount', ''), '');
    Section := 'UPDATEFREQUENCY';
    AddTableList('CHANCELL' , 'Everytime');
    AddTableList('CHOIXCOD' , 'Everytime');
    AddTableList('CHOIXEXT' , 'Everytime');
    AddTableList('CODEPOST' , 'Monthly');
    AddTableList('COMMUN'   , 'Daily');
    AddTableList('CORRESP'  , 'Everytime');
    AddTableList('DEVISE'   , 'Everytime');
    AddTableList('ETABLISS' , 'Everytime');
    AddTableList('EXERCICE' , 'Daily');
    AddTableList('GENERAUX' , 'Everytime');
    AddTableList('JOURNAL'  , 'Daily');
    AddTableList('MODEPAIE' , 'Daily');
    AddTableList('MODEREGL' , 'Daily');
    AddTableList('PARAMSOC' , 'Daily');
    AddTableList('PAYS'     , 'Daily');
    AddTableList('RELANCE'  , 'Everytime');
    AddTableList('TXCPTTVA' , 'Everytime');
    AddTableList('ECRITURE' , 'Everytime');
    SectionExist := True;
    Cpt          := 0;
    while SectionExist do
    begin
      Inc(Cpt);
      Section      := 'FOLDER' + IntToStr(Cpt);
      SectionExist := (SettingFile.ReadString(Section, 'BTPUser', '') <> '');
      if SectionExist then
      begin
        AddFolderList(SectionExist, 'LastSynchro', '01/01/1900 00:00:00');
        AddFolderList(SectionExist, 'BTPUser'    , '');
        AddFolderList(SectionExist, 'BTPServer'  , '');
        AddFolderList(SectionExist, 'BTPFolder'  , '');
        AddFolderList(SectionExist, 'Y2Server'   , '');
        AddFolderList(SectionExist, 'Y2Folder'   , '');
      end;
    end;
  finally
    SettingFile.Free;
  end;
  ShowTables;
  ShowFolders(0);
end;

function TParamSvcSyncrho.SetMnTimeOut: string;
begin
  lSecondTimeoutUnit.Caption := Format('secondes (%s Minutes)', [FormatFloat('#,##0', SecondTimeout.Value /60)]);
end;

procedure TParamSvcSyncrho.ReloadClick(Sender: TObject);
begin
  if PGIAsk(Format('Voulez-vous recharger le paramétrage défini dans %s ?%s%s', [PathFileName, #13#10, cWarningMsg])) = mrYes then
    LoadIniIfExist;  

end;

procedure TParamSvcSyncrho.LogLevelClick(Sender: TObject);
begin
  ShowMaxSizeType;
  ShowLogType;
  LogLevelDebug.Visible := (LogLevel.ItemIndex > 0);
//  LogSee.Visible        := (LogLevel.ItemIndex > 0);
  if LogLevel.ItemIndex = 0 then
    LogLevelDebug.ItemIndex := 0;
end;

procedure TParamSvcSyncrho.LogTypeClick(Sender: TObject);
begin
  ShowMaxSizeType;
end;

procedure TParamSvcSyncrho.UploadClick(Sender: TObject);
begin
  ShowDebugFilesDirectory;
end;

procedure TParamSvcSyncrho.ShowLogType;
begin
  LogType.Visible := LogLevel.ItemIndex > 0;
  if not LogType.Visible then
  begin
    LogType.ItemIndex  := 0;
    LogMaxSize.Visible := False;
    LogMaxSize.Value   := 0;
  end;
end;

procedure TParamSvcSyncrho.ShowMaxSizeType;
begin
  LogMaxSize.Visible := (LogType.ItemIndex = 1);
  if not LogMaxSize.Visible then
    LogMaxSize.Value := 0;
end;

procedure TParamSvcSyncrho.ShowDebugFilesDirectory;
begin
  DebugFilesDirectory.Visible  := Upload.Checked;
  lDebugFilesDirectory.Visible := Upload.Checked;
  if not DebugFilesDirectory.Visible then
    DebugFilesDirectory.Text := '';
end;

procedure TParamSvcSyncrho.gFrequency1DblClick(Sender: TObject);
var
  TableName : string;
  Value     : string;
  Cpt       : integer;
  lIndex    : integer;

  function GetParamName(ColName : integer) : string;
  begin
    case ColName of
      2 : Result := 'Everytime';
      3 : Result := 'Daily';
      4 : Result := 'Monthly';
      5 : Result := 'Annual';
      6 : Result := 'Once';
      7 : Result := 'Never';
    else ;
      Result := '';
    end;
  end;

begin
  TableName  := gFrequency1.Cells[1, gFrequency1.Row];
  for Cpt := 2 to 7 do
  begin
    Value := Tools.iif(Cpt = gFrequency1.Col, 'X', '');
    gFrequency1.Cells[Cpt, gFrequency1.Row] := Value;
    if Value = 'X'  then
    begin
      lIndex := TslTableList.IndexOfName(TableName);
      TslTableList[lIndex] := Format('%s=%s', [TableName, GetParamName(Cpt)]);
    end;
  end;
end;

procedure TParamSvcSyncrho.cFolderChoiceClick(Sender: TObject);
var
  SelectdFolder : string;
  Value         : string;
  TslFolder     : string;
  TslParam      : string;
  Cpt           : integer;
begin
  SelectdFolder := cFolderChoice.Items.ValueFromIndex[cFolderChoice.ItemIndex];
  for Cpt := 0 to pred(TslFolderList.count) do
  begin
    Value     := TslFolderList[Cpt];
    TslFolder := copy(Value, 1, pos('=', Value)-1);
    Value     := copy(Value, pos('=', Value)+1, length(Value));
    TslParam  := Tools.ReadTokenSt_(Value, '|');
    if TslFolder = SelectdFolder then
    begin
      case Tools.CaseFromString(TslParam, ['LastSynchro', 'BTPUser', 'BTPServer', 'BTPFolder', 'Y2Server', 'Y2Folder']) of
        {LastSynchro} 0 : LastSynchro.Text := Value;
        {BTPUser}     1 : BTPUser.Text     := Value;
        {BTPServer}   2 : BTPServer.Text   := Value;
        {BTPFolder}   3 : BTPFolder.Text   := Value;
        {Y2Server}    4 : Y2Server.Text    := Value;
        {Y2Folder}    5 : Y2Folder.Text    := Value;
      end;
    end;
  end;
end;

procedure TParamSvcSyncrho.NewFolderClick(Sender: TObject);
begin
  OnNewFolder := True;
  cFolderChoice.AddItem(' FOLDER' + IntToStr(cFolderChoice.Count+1), nil);
  cFolderChoice.ItemIndex := cFolderChoice.Count-1;
  ClearFolderParams;
  LastSynchro.text := '01/01/1900 00:00:00';
  ManageFolderParams(True);
  ManageOtherControls(False);
  if LastSynchro.CanFocus then
    LastSynchro.SetFocus;
end;

procedure TParamSvcSyncrho.UpdateFolderClick(Sender: TObject);
begin
  ManageFolderParams(True);
  ManageOtherControls(False);
  if LastSynchro.CanFocus then
    LastSynchro.SetFocus;
end;

procedure TParamSvcSyncrho.DelFolderClick(Sender: TObject);
var
  FolderName : string;
  Value      : string;
  DoRecalcul : boolean;
  FolderNum  : integer;

  lIndexOf   : integer;
  Cpt        : integer;
begin
  FolderNum  := cFolderChoice.ItemIndex+1;
  FolderName := Format('FOLDER%s', [IntToStr(FolderNum)]);
  if PGIAsk(Format('Voulez-vous supprimer le dossier %s ?', [FolderName])) = mrYes then
  begin
    DoRecalcul := (cFolderChoice.Count > cFolderChoice.ItemIndex+1);
    lIndexOf   := TslFolderList.IndexOfName(FolderName);
    while lIndexOf > -1 do
    begin
      TslFolderList.Delete(lIndexOf);
      lIndexOf := TslFolderList.IndexOfName(FolderName);
    end;
    if DoRecalcul then
    begin
      inc(FolderNum);
      FolderName := Format('FOLDER%s', [IntToStr(FolderNum)]);
      lIndexOf   := TslFolderList.IndexOfName(FolderName);
      for Cpt := lIndexOf to pred(TslFolderList.count) do
      begin
        Value     := TslFolderList[Cpt];
        FolderNum := StrToInt(Copy(Value, 7, pos('=', Value)-7))-1;
        Value     := Format('FOLDER%s=%s', [IntToStr(FolderNum), copy(Value, pos('=', Value)+1, length(Value))]);
        TslFolderList[Cpt] := Value;
      end;
    end;
    ManageFolderParams(False);
    ManageOtherControls(True);
    ShowFolders(0);
  end;
end;

procedure TParamSvcSyncrho.AddNewFolderClick(Sender: TObject);
var
  NumError      : integer;
  CurrentFolder : integer;
  Cpt           : integer;
  CptCurrFolder : integer;
  Msg           : string;
  Value         : string;

  function SetTslValue(Param : string; FolderNum : integer; Value : string) : string;
  begin
    Result := Format('FOLDER%s=%s|%s', [IntToStr(FolderNum), Param, Value]);
  end;

begin
       if LastSynchro.Text = '' then NumError := 1
  else if BTPUser.Text     = '' then NumError := 2
  else if BTPServer.Text   = '' then NumError := 3
  else if BTPFolder.Text   = '' then Numerror := 4
  else if Y2Server.Text    = '' then NumError := 5
  else if Y2Folder.Text    = '' then NumError := 6
  else                               NumError := 0;
  case NumError of
    1 : Msg := 'La date de synchronisation';
    2 : Msg := 'L''utilisateur';
    3 : Msg := 'Le serveur BTP';
    4 : Msg := 'Le dossier BTP';
    5 : Msg := 'Le serveur Y2';
    6 : Msg := 'Le dossier Y2';
  else
    Msg := '';
  end;
  if NumError > 0 then
  begin
    PGIError(Format('%s est obligatoire.', [Msg]));
    case NumError of
      1 : LastSynchro.SetFocus;
      2 : BTPUser.SetFocus;
      3 : BTPServer.SetFocus;
      4 : BTPFolder.SetFocus;
      5 : Y2Server.SetFocus;
      6 : Y2Folder.SetFocus;
    end;
  end else
  begin
    if OnNewFolder then
    begin
      CurrentFolder := cFolderChoice.Count;
      TslFolderList.Add(SetTslValue('LastSynchro', CurrentFolder, LastSynchro.Text));
      TslFolderList.Add(SetTslValue('BTPUser'    , CurrentFolder, BTPUser.Text));
      TslFolderList.Add(SetTslValue('BTPServer'  , CurrentFolder, BTPServer.Text));
      TslFolderList.Add(SetTslValue('BTPFolder'  , CurrentFolder, BTPFolder.Text));
      TslFolderList.Add(SetTslValue('Y2Server'   , CurrentFolder, Y2Server.Text));
      TslFolderList.Add(SetTslValue('Y2Folder'   , CurrentFolder, Y2Folder.Text));
    end else
    begin
      CurrentFolder := cFolderChoice.ItemIndex+1;
      CptCurrFolder := 0;
      for Cpt := 0 to pred(TslFolderList.count) do
      begin
        Value := TslFolderList[Cpt];
        if copy(Value, 1 ,pos('=', Value)-1) = Format('FOLDER%s', [IntToStr(CurrentFolder)]) then
        begin
          inc(CptCurrFolder);
          case CptCurrFolder of
            1 : TslFolderList[Cpt] := SetTslValue('LastSynchro', CurrentFolder, LastSynchro.Text);
            2 : TslFolderList[Cpt] := SetTslValue('BTPUser'    , CurrentFolder, BTPUser.Text);
            3 : TslFolderList[Cpt] := SetTslValue('BTPServer'  , CurrentFolder, BTPServer.Text);
            4 : TslFolderList[Cpt] := SetTslValue('BTPFolder'  , CurrentFolder, BTPFolder.Text);
            5 : TslFolderList[Cpt] := SetTslValue('Y2Server'   , CurrentFolder, Y2Server.Text);
            6 : TslFolderList[Cpt] := SetTslValue('Y2Folder'   , CurrentFolder, Y2Folder.Text);
          end;
        end;
      end;
    end;
    OnNewFolder := False;
    ManageFolderParams(False);
    ManageOtherControls(True);
    ShowFolders(cFolderChoice.ItemIndex);
  end;
end;

procedure TParamSvcSyncrho.CancelNewFolderClick(Sender: TObject);
begin
  if PGIAsk('Voulez-vous annuler la saisie en cours ?') = mrYes then
  begin
    OnNewFolder := False;
    ClearFolderParams;
    ManageFolderParams(False);
    ManageOtherControls(True);
    ShowFolders(0);
  end;
end;

procedure TParamSvcSyncrho.ManageFolderParams(pEnabled : boolean);
begin
  LastSynchro.Enabled     := pEnabled;
  BTPUser.Enabled         := pEnabled;
  BTPServer.Enabled       := pEnabled;
  BTPFolder.Enabled       := pEnabled;
  Y2Server.Enabled        := pEnabled;
  Y2Folder.Enabled        := pEnabled;
  AddNewFolder.Visible    := pEnabled;
  CancelNewFolder.Visible := pEnabled;
  ConnectBTP.Visible      := pEnabled;
  ConnectY2.Visible       := pEnabled;
end;

procedure TParamSvcSyncrho.ManageOtherControls(pEnabled: boolean);
begin
  cFolderChoice.Enabled := pEnabled;
  NewFolder.Enabled     := pEnabled;
  UpdateFolder.Enabled  := pEnabled;
  DelFolder.Enabled     := pEnabled;
  Reload.Enabled        := pEnabled;
  bValider.Enabled      := pEnabled;
end;

procedure TParamSvcSyncrho.ClearFolderParams;
begin
  LastSynchro.Text := '';
  BTPUser.Text     := '';
  BTPServer.Text   := '';
  BTPFolder.Text   := '';
  Y2Server.Text    := '';
  Y2Folder.Text    := '';
end;

procedure TParamSvcSyncrho.SecondTimeoutExit(Sender: TObject);
begin
  SetMnTimeOut;
  if     (StrToInt(SecondTimeout.Text) < 900)
     and (PGIAsk(Format('Il est déconseillé de paramétrer un déclenchement inférieur à 15 mn.%sVoulez-vous continuer ?', [#13#10])) <> mrYes)
  then
  begin
    SecondTimeout.Text := '900';
    SetMnTimeOut;
  end;
end;

procedure TParamSvcSyncrho.SecondTimeoutClick(Sender: TObject);
begin
  SetMnTimeOut;
end;


procedure TParamSvcSyncrho.SecondTimeoutKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  SetMnTimeOut;
end;

procedure TParamSvcSyncrho.SecondTimeoutKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  SetMnTimeOut;
end;

procedure TParamSvcSyncrho.ExecutionPeriodStartExit(Sender: TObject);
begin
  TimeSlotIsValid(True, ExecutionPeriodStart.Text);
end;

procedure TParamSvcSyncrho.ExecutionPeriodEndExit(Sender: TObject);
begin
  TimeSlotIsValid(False, ExecutionPeriodEnd.Text);
end;

procedure TParamSvcSyncrho.TimeSlotIsValid(IsStart: boolean; pValue: string);
var
  Error : boolean;
begin
  Error := ((pValue < '00:00:00') or ((pValue > '23:59:59')));
  if Error then
  begin
    PGIError('La plage horaire doit être comprise entre 00:00:00 et 23:59:59');
    if IsStart then
    begin
      ExecutionPeriodStart.Text := '00:00:00';
      ExecutionPeriodStart.SetFocus
    end else
    begin
      ExecutionPeriodEnd.Text := '23:59:59';
      ExecutionPeriodEnd.SetFocus;
    end;
  end;
end;

function TParamSvcSyncrho.GetServiceState : tServiceState;
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(cServiceName);
  try
    Result := Svcm.GetState;
  finally
    Svcm.Free;
  end;
end;

function TParamSvcSyncrho.GetSvcPathName : string;
begin
  Result := Format('%s\%s.exe', [SvcInstallPath.Text, cExeServiceName]);
end;

procedure TParamSvcSyncrho.SetCaptionServiceState;
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(cServiceName);
  try
    ServiceState.Caption := Svcm.GetStateLabel;
  finally
    Svcm.Free;
  end;
end;

procedure TParamSvcSyncrho.ServiceButtonManagement;
var
  SvcState : tServiceState;
begin
  SvcState                := GetServiceState;
  gSvcInstalParam.Visible := Tools.iif((SvcState=tssUnknown) or (SvcState=tssUnInstall), True, False); // Inconnu, Non installé
  SvcInstall.Enabled      := Tools.iif((SvcState=tssUnknown) or (SvcState=tssUnInstall), True, False); // Inconnu, Non installé
  SvcUnInstall.Enabled    := Tools.iif((SvcState=tssStopped) or (SvcState=tssPaused), True, False);    // Arrêté, En pause
  SvcStart.Enabled        := Tools.iif((SvcState=tssStopped) or (SvcState=tssPaused), True, False);    // Arrêté, En pause
  SvcStop.Enabled         := Tools.iif((SvcState=tssRunning), True, False);                            // En cours d''exécution
  SetCaptionServiceState;
end;

procedure TParamSvcSyncrho.SvcInstallClick(Sender: TObject);
var
  SvcM        : SvcManagement;
  SvcType     : tServiceStartType;
  CanContinue : boolean;
begin
  if SvcInstallPath.Text = '' then
  begin
    PGIError('Veuillez sélectionner l''emplacement du service.');
    CanContinue := False;
  end else
  if (SvcAccountTipe.ItemIndex = 1) then
  begin
    CanContinue := (    (SvcInstallAccount.Text    <> '')
                    and (SvcInstallPwd.Text        <> '')
                    and (SvcInstallPwdConfirm.Text <> '')
                    and ((SvcInstallPwd.Text = SvcInstallPwdConfirm.Text))
                   );
    if not CanContinue then
    begin
      if SvcInstallAccount.Text = '' then
      begin
        PGIError('Veuillez saisir le compte.');
        SvcInstallAccount.SetFocus;
      end else
      if SvcInstallPwd.Text = '' then
      begin
        PGIError('Veuillez saisir le mot de passe.');
        SvcInstallPwd.SetFocus;
      end else
      if SvcInstallPwdConfirm.Text = '' then
      begin
        PGIError('Veuillez saisir la confirmation du mot de passe.');
        SvcInstallPwdConfirm.SetFocus;
      end else
      if (SvcInstallPwd.Text <> SvcInstallPwdConfirm.Text) then
      begin
        PGIError('Les mots de passe ne correspondent pas.');
        SvcInstallPwd.SetFocus;
      end;
    end;
  end else
    CanContinue := True;
  if CanContinue then
  begin
    Svcm := SvcManagement.Create(cServiceName);
    try
      case SvcInstallType.ItemIndex of
        0 : SvcType := tsstManuel;
        1 : SvcType := tsstAutomatic;
        2 : SvcType := tsstDisabled;
      else
        SvcType := tsstUnknown;
      end;
      Svcm.Install(GetSvcPathName, cServiceName, Tools.iif((SvcAccountTipe.ItemIndex = 0), '', SvcInstallAccount.Text), Tools.iif((SvcAccountTipe.ItemIndex = 0), '', SvcInstallPwd.Text), SvcType);
    finally
      SvcInstallPwd.Text         := '';
      SvcInstallPwdConfirm .Text := '';
      Svcm.Free;
    end;
    ServiceButtonManagement;
  end;
end;

procedure TParamSvcSyncrho.SvcUnInstallClick(Sender: TObject);
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(cServiceName);
  try
    Svcm.UnInstall(cServiceName);
  finally
    Svcm.Free;
  end;
  ServiceButtonManagement;
end;

procedure TParamSvcSyncrho.SvcStartClick(Sender: TObject);
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(cServiceName);
  try
    Svcm.Start(cServiceName);
  finally
    Svcm.Free;
  end;
  ServiceButtonManagement;
end;

procedure TParamSvcSyncrho.SvcStopClick(Sender: TObject);
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(cServiceName);
  try
    Svcm.Stop(cServiceName);
  finally
    Svcm.Free;
  end;
  ServiceButtonManagement;
end;

procedure TParamSvcSyncrho.SvcStateRefreshClick(Sender: TObject);
begin
  SetCaptionServiceState;
  ServiceButtonManagement;
end;

procedure TParamSvcSyncrho.SvcAccountTipeClick(Sender: TObject);
var
  Choice : integer;
begin
  Choice := SvcAccountTipe.ItemIndex;
  lInstallAccount.Visible      := (Choice = 1);
  SvcInstallAccount.Visible    := (Choice = 1);
  lInstallPwd.Visible          := (Choice = 1);
  SvcInstallPwd.Visible        := (Choice = 1);
  lInstallPwdConfirm.Visible   := (Choice = 1);
  SvcInstallPwdConfirm.Visible := (Choice = 1);
  if (Choice = 1) and (SvcInstallAccount.CanFocus) then
    SvcInstallAccount.SetFocus;
end;

procedure TParamSvcSyncrho.gFrequency1MouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ;
end;

procedure TParamSvcSyncrho.TestConnectClick(Sender: TObject);
var
  ServerName : string;
  FolderName : string;
  Msg        : WaitingMessage;
begin
  ServerName := Tools.iif(TMaskEdit(sender).Name = 'ConnectBTP', BTPServer.Text, Y2Server.Text);
  FolderName := Tools.iif(TMaskEdit(sender).Name = 'ConnectBTP', BTPFolder.Text, Y2Folder.Text);
  if (ServerName = '') or (FolderName = '') then
    PGIError('Les noms du serveur et du dossiers doivent être renseingnés.')
  else
  begin
    Msg := WaitingMessage.Create('Test en cours', Format('Connexion %s/%s', [ServerName, FolderName]), True);
    try
      if Tools.GetParamSocSecur_('SO_SOCIETE', '', ServerName, FolderName) = '' then
        PGIError(Format('Erreur sur la connexion %s/%s', [ServerName, FolderName]))
      else
        PGIBox(Format('Connexion réussie sur %s/%s', [ServerName, FolderName]));
    finally
      Msg.Free;
    end;
  end;
end;

procedure TParamSvcSyncrho.SeeLogsClick(Sender: TObject);
begin
  ;
end;

end.

