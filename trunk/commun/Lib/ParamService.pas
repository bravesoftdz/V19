unit ParamService;

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
    gServicesLst: TGroupBox;
    SvcList: THGrid;
    gServiceMangement: TGroupBox;
    gSvcState: TGroupBox;
    gSvcManagement: TGroupBox;
    pButtons: TPanel;
    bFermer: THBitBtn;
    lSvcState: TLabel;
    lSvcStart: TLabel;
    ServiceState: TEdit;
    ServiceTypeStart: TEdit;
    lSvcAccount: TLabel;
    SvcAccount: TEdit;
    lSvcLocation: TLabel;
    SvcLocalisation: TEdit;
    Panel1: TPanel;
    SvcStateRefresh: THBitBtn;
    SvcUnInstall: THBitBtn;
    SvcStart: THBitBtn;
    SvcStop: THBitBtn;
    gSvcInstallPath: TGroupBox;
    SvcInstallPath: THCritMaskEdit;
    SvcInstallType: TRadioGroup;
    gSvcInstallConnection: TGroupBox;
    lInstallAccount: TLabel;
    lInstallPwd: TLabel;
    lInstallPwdConfirm: TLabel;
    SvcInstallAccount: TEdit;
    SvcInstallPwd: TEdit;
    SvcInstallPwdConfirm: TEdit;
    SvcAccountTipe: TRadioGroup;
    pnl2: TPanel;
    SvcInstall: THBitBtn;
    SvcSystem: THBitBtn;
    pnl3: TPanel;
    SvcLstRefresh: THBitBtn;
    ChkLSESrv: TCheckBox;
    lSvcSearch: TLabel;
    SvcSearch: TEdit;
    SvcLstSearch: THBitBtn;

    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SvcInstallClick(Sender: TObject);
    procedure SvcUnInstallClick(Sender: TObject);
    procedure SvcStartClick(Sender: TObject);
    procedure SvcStopClick(Sender: TObject);
    procedure SvcStateRefreshClick(Sender: TObject);
    procedure SvcAccountTipeClick(Sender: TObject);
    procedure bFermerClick(Sender: TObject);
    procedure SvcListClick(Sender: TObject);
    procedure SvcLstRefreshClick(Sender: TObject);
    procedure SvcLstSearchClick(Sender: TObject);
    function SearchInList(Value : string; StartOn : integer) : boolean;
    procedure ClearInstallDatas;
    procedure SvcSystemClick(Sender: TObject);
    function IsLseService(FileName : string) : boolean;
  private
    TslTableList  : TStringList;
    TslFolderList : TStringList;

    function GetSelectedService : string;
    function GetServiceState : tServiceState;
    function GetLocalisation(SvcName : string) : string;
    procedure SetCaptionServiceState;
    procedure ServiceButtonManagement;
    procedure ShowButtons;
    procedure ServiceGetList;

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
  //cExeServiceName = 'SvcSynBTPY2';

{$R *.dfm}
procedure TParamSvcSyncrho.FormCreate(Sender: TObject);
begin
  ParamSvcSyncrho.Position := poScreenCenter;
end;

procedure TParamSvcSyncrho.FormShow(Sender: TObject);
begin
  SvcList.SortEnabled := True;
  SvcList.ColWidths[0] := 10;
  SvcList.ColWidths[1] := 500;
  SvcList.ColWidths[2] := 200;
  if not IsUserAnAdmin then
  begin
    PGIBox('L''application doit être exécutée en mode "Administrateur".');
    SvcStateRefresh.Visible := False;
    gSvcManagement.Visible  := False;
    ShowButtons;
  end else
  begin
    gSvcState.Caption := Format('  Etat du service %s  ', [GetSelectedService]);
    ServiceGetList;
  end;
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

function TParamSvcSyncrho.GetSelectedService : string;
begin
  Result := SvcList.Cells[2, SvcList.Row]
end;

function TParamSvcSyncrho.GetServiceState : tServiceState;
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(GetSelectedService);
  try
    Result := Svcm.GetState;
  finally
    Svcm.Free;
  end;
end;

function TParamSvcSyncrho.GetLocalisation(SvcName : string) : string;
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(SvcName);
  try
    Result := Svcm.GetLocation;
  finally
    Svcm.Free;
  end;
end;

procedure TParamSvcSyncrho.SetCaptionServiceState;
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(GetSelectedService);
  try
    ServiceState.Text     := Svcm.GetStateLabel;
    ServiceTypeStart.Text := Svcm.GetStartTypeLabel(Svcm.GetStartType);
    SvcAccount.Text       := Svcm.GetConnectionAccount;
    SvcLocalisation.Text  := Svcm.GetLocation;
  finally
    Svcm.Free;
  end;
end;

procedure TParamSvcSyncrho.ServiceButtonManagement;
var
  SvcState : tServiceState;
begin
  SvcState             := GetServiceState;
  SvcUnInstall.Enabled := Tools.iif((SvcState=tssStopped) or (SvcState=tssPaused), True, False);    // Arrêté, En pause
  SvcStart.Enabled     := Tools.iif((SvcState=tssStopped) or (SvcState=tssPaused), True, False);    // Arrêté, En pause
  SvcStop.Enabled      := Tools.iif((SvcState=tssRunning), True, False);                            // En cours d''exécution
  SetCaptionServiceState;
end;

procedure TParamSvcSyncrho.ShowButtons;
var
  IsSystem : boolean;
begin
//  gSvcInstalParam.Visible := Tools.iif((SvcState=tssUnknown) or (SvcState=tssUnInstall), True, False); // Inconnu, Non installé
//  SvcInstall.Enabled      := Tools.iif((SvcState=tssUnknown) or (SvcState=tssUnInstall), True, False); // Inconnu, Non installé
  IsSystem             := (pos(UpperCase(Tools.GetSystDir), UpperCase(SvcLocalisation.Text)) > 0);
  SvcSystem.Visible    := IsSystem;
  SvcUnInstall.Visible := (not IsSystem);
  SvcStart.Visible     := (not IsSystem);
  SvcStop.Visible      := (not IsSystem);
  if not IsSystem then
    ServiceButtonManagement;
end;

procedure TParamSvcSyncrho.SvcInstallClick(Sender: TObject);
var
  SvcM        : SvcManagement;
  SvcType     : tServiceStartType;
  CanContinue : boolean;
  ServiceName : string;
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
    ServiceName := ExtractFileName(SvcInstallPath.Text);
    ServiceName := copy(ServiceName, 1, pos('.', ServiceName)-1);
    case Tools.CaseFromString(ServiceName, ['SvcSynBTPY2', 'SvcSynBTPVerdonExp', 'SvcSynBTPVerdonImp']) of
      {} 0 : ServiceName := 'SvcSyncBTPY2';
      {} 1 : ServiceName := 'SvcSyncBTPVerdonExp';
      {} 2 : ServiceName := 'SvcSyncBTPVerdonImp';
    end;
    if SearchInList(ServiceName, 1) then
    begin
      PGIError(Format('Le service %s est déjà installé.', [ServiceName]));
    end else
    begin
      Svcm := SvcManagement.Create(ServiceName);
      try
        case SvcInstallType.ItemIndex of
          0 : SvcType := tsstManuel;
          1 : SvcType := tsstAutomatic;
          2 : SvcType := tsstDisabled;
        else
          SvcType := tsstUnknown;
        end;
        Svcm.Install(SvcInstallPath.Text, ServiceName, Tools.iif((SvcAccountTipe.ItemIndex = 0), '', SvcInstallAccount.Text), Tools.iif((SvcAccountTipe.ItemIndex = 0), '', SvcInstallPwd.Text), SvcType);
      finally
        SvcInstallPwd.Text         := '';
        SvcInstallPwdConfirm .Text := '';
        Svcm.Free;
      end;
      ServiceGetList;
      SearchInList(ServiceName, 1);
      ShowButtons;
    end;
    ClearInstallDatas;
  end;
end;

procedure TParamSvcSyncrho.SvcUnInstallClick(Sender: TObject);
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(GetSelectedService);
  try
    Svcm.UnInstall(GetSelectedService);
  finally
    Svcm.Free;
  end;
  ServiceGetList;
  ShowButtons;
end;

procedure TParamSvcSyncrho.SvcStartClick(Sender: TObject);
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(GetSelectedService);
  try
    Svcm.Start(GetSelectedService);
  finally
    Svcm.Free;
  end;
  ShowButtons;
end;

procedure TParamSvcSyncrho.SvcStopClick(Sender: TObject);
var
  SvcM : SvcManagement;
begin
  Svcm := SvcManagement.Create(GetSelectedService);
  try
    Svcm.Stop(GetSelectedService);
  finally
    Svcm.Free;
  end;
  ShowButtons;
end;

procedure TParamSvcSyncrho.SvcStateRefreshClick(Sender: TObject);
begin
  SetCaptionServiceState;
  ShowButtons;
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

procedure TParamSvcSyncrho.ServiceGetList;

const
	cnMaxServices = 4096;

type
	TSvcA = array[0..cnMaxServices]
	of TEnumServiceStatus;
	PSvcA = ^TSvcA;

var
  SvcM           : SvcManagement;
	Cpt            : integer;
	schm           : SC_Handle;
	nBytesNeeded   : DWord;
	nServices      : DWord;
	nResumeHandle  : DWord;
  dwServiceType  : DWord;
  dwServiceState : DWord;
	ssa            : PSvcA;
  SvcName        : string;
  SvcLongName    : string;
  SvcValues      : string;
  SvcCop         : string;
  StartType      : tServiceStartType;
  TslSvcLst      : TStringList;
  AddInList      : boolean;
  WaitMsg        : WaitingMessage;
begin
  WaitMsg := WaitingMessage.Create('Liste des services.', 'Recherche en cours.');
  try
    SvcList.RowCount := 2;
    TslSvcLst        := TStringList.Create;
    try
      dwServiceType    := SERVICE_WIN32;
      dwServiceState   := SERVICE_STATE_ALL;
      schm             := OpenSCManager(PChar(Tools.tGetComputerName), Nil, SC_MANAGER_ALL_ACCESS);
      try
        if(schm > 0)then
        begin
          nResumeHandle := 0;
          New(ssa);
          EnumServicesStatus(schm, dwServiceType, dwServiceState, ssa^[0], SizeOf(ssa^), nBytesNeeded, nServices, nResumeHandle );
          for Cpt := 1 to pred(nServices) do
          begin
            SvcName     := ssa^[Cpt].lpServiceName;
            SvcLongName := ssa^[Cpt].lpDisplayName;
            if SvcName <> '' then
            begin
              Svcm := SvcManagement.Create(SvcName);
              try
                StartType := Svcm.GetStartType;
                if (StartType = tsstAutomatic) or (StartType = tsstDifferedAutomatic) or (StartType = tsstManuel) then
                begin
                  if ChkLSESrv.Checked then
                  begin
                    SvcCop    := Tools.GetFileProperty(GetLocalisation(SvcName), tfpLegalCopyright);
                    AddInList := (Pos('LSE', SvcCop) > 0);
                  end else
                    AddInList := True;
                  if AddInList then
                    TslSvcLst.Add(Format('%s;%s', [SvcLongName, SvcName]));
                end;
              finally
                Svcm.Free;
              end;
            end;
          end;
          Dispose(ssa);
        end;
      finally
        CloseServiceHandle(schm);
        for Cpt := 0 to pred(TslSvcLst.Count) do
        begin
          SvcValues               := TslSvcLst[Cpt];
          SvcList.Cells[1, Cpt+1] := Tools.ReadTokenSt_(SvcValues, ';');
          SvcList.Cells[2, Cpt+1] := Tools.ReadTokenSt_(SvcValues, ';');
          SvcList.RowCount        := SvcList.RowCount + 1;
        end;
        SvcList.SortGrid(1, False);
      end;
    finally
      FreeAndNil(TslSvcLst);
    end;
    SvcStateRefreshClick(nil);
  finally
    WaitMsg.Free;
  end;
end;

procedure TParamSvcSyncrho.bFermerClick(Sender: TObject);
begin
  if PGIAsk('Voulez-vous quitter ?') = mrYes then
    Application.Terminate;
end;

procedure TParamSvcSyncrho.SvcListClick(Sender: TObject);
begin
  SetCaptionServiceState;
  ShowButtons;
end;

procedure TParamSvcSyncrho.SvcLstRefreshClick(Sender: TObject);
var
  Cpt : integer;
begin
  for Cpt := pred(SvcList.RowCount) Downto 1 do
    SvcList.Rows[Cpt].Clear;
  ServiceGetList;
end;

procedure TParamSvcSyncrho.SvcLstSearchClick(Sender: TObject);
begin
  SearchInList(SvcSearch.Text, SvcList.Row +1);
end;

function TParamSvcSyncrho.SearchInList(Value: string; StartOn : integer) : boolean;
var
  Cpt : integer;
begin
  Result := False;
  if Value <> '' then
  begin
    for Cpt := StartOn to pred(SvcList.RowCount) do
    begin
      Result := ((pos(UpperCase(Value), UpperCase(SvcList.Cells[1, Cpt])) > 0) or (pos(UpperCase(Value), UpperCase(SvcList.Cells[2, Cpt])) > 0));
      if Result then
      begin
        SvcList.Row := Cpt;
        break;
      end;
    end;
  end;
end;

procedure TParamSvcSyncrho.ClearInstallDatas;
begin
  SvcInstallPath.Text      := '';
  SvcInstallType.ItemIndex := 0;
  SvcAccountTipe.ItemIndex := 0;
  SvcInstallAccount.Text   := '';
end;

procedure TParamSvcSyncrho.SvcSystemClick(Sender: TObject);
begin
  SvcSystem.Visible    := False;
  SvcUnInstall.Visible := True;
  SvcStart.Visible     := True;
  SvcStop.Visible      := True;
  ServiceButtonManagement;
end;

function TParamSvcSyncrho.IsLseService(FileName : string) : boolean;

	type
		PLandCodepage = ^TLandCodepage;
		TLandCodepage = record
		wLanguage,
		wCodePage: word;
	end;

	var
		dummy : cardinal;
		len   : cardinal;
		buf   : pointer;
    pntr  : pointer;
		lang  : string;
    lFileName : string;
//    Cop : string;
begin
  Result := False;
  lFileName := copy(FileName, 1, pos('.exe', FileName) + 4);
	len    := GetFileVersionInfoSize(PChar(FileName), dummy);
	if len > 0 then
  begin
    GetMem(buf, len);
    try
  //		if not GetFileVersionInfo(PChar(FileName), 0, len, buf) then
  //		RaiseLastOSError;
  //		if not VerQueryValue(buf, '\VarFileInfo\Translation\', pntr, len) then
  //		RaiseLastOSError;
      lang := Format('%.4x%.4x', [PLandCodepage(pntr)^.wLanguage, PLandCodepage(pntr)^.wCodePage]);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\LegalCopyright'), pntr, len) then
        if Pos('LSE', PChar(pntr)) > 0 then
          Result := True;
      if Result then
       ;

  (*
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\CompanyName'), pntr, len){ and (@len <> nil)} then
      result.CompanyName := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\FileDescription'), pntr, len){ and (@len <> nil)} then
      result.FileDescription := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\FileVersion'), pntr, len){ and (@len <> nil)} then
      result.FileVersion := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\InternalName'), pntr, len){ and (@len <> nil)} then
      result.InternalName := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\LegalCopyright'), pntr, len){ and (@len <> nil)} then
      result.LegalCopyright := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\LegalTrademarks'), pntr, len){ and (@len <> nil)} then
      result.LegalTrademarks := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\OriginalFileName'), pntr, len){ and (@len <> nil)} then
      result.OriginalFileName := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\ProductName'), pntr, len){ and (@len <> nil)} then
      result.ProductName := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\ProductVersion'), pntr, len){ and (@len <> nil)} then
      result.ProductVersion := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\Comments'), pntr, len){ and (@len <> nil)} then
      result.Comments := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\PrivateBuild'), pntr, len){ and (@len <> nil)} then
      result.PrivateBuild := PChar(pntr);
      if VerQueryValue(buf, PChar('\StringFileInfo\' + lang + '\SpecialBuild'), pntr, len){ and (@len <> nil)} then
      result.SpecialBuild := PChar(pntr);
  *)
    finally
      FreeMem(buf);
    end;
  end;
end;

end.

