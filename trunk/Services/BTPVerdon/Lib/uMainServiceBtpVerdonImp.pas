unit uMainServiceBtpVerdonImp;

interface

uses
  Windows
  , Messages
  , SysUtils
  , Classes
  , Graphics
  , Controls
  , SvcMgr
  , Dialogs
  , ConstServices
  , UtilBTPVerdon
  ;

type
  TSvcSyncBTPVerdonImp = class(TService)
    procedure ServiceAfterInstall(Sender: TService);
    {$IFDEF APPSRVWITHCBP}
    procedure ServiceExecute(Sender: TService);
    {$ENDIF APPSRVWITHCBP}
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
  private
    IniPath                   : string;
    AppPath                   : string;
    LogPath                   : string;
    TslParametersTypeMatching : TStringList;

    function ReadSettings(Import : TImportTreatment) : string;
    procedure ClearTablesValues(Import : TImportTreatment);
    procedure SetFieldsList(Import : TImportTreatment);

  public
    function GetServiceController: TServiceController; override;
  end;

var
  SvcSyncBTPVerdonImp: TSvcSyncBTPVerdonImp;

implementation

uses
  Registry
  , CommonTools
  , ActiveX
  , WinSVC
  , ShellAPI
  , IniFiles
  , StrUtils
  , AdoDB
  , DateUtils
  ;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  SvcSyncBTPVerdonImp.Controller(CtrlCode);
end;

function TSvcSyncBTPVerdonImp.ReadSettings(Import : TImportTreatment) : string;
var
  SettingFile : TInifile;
  Section     : string;
  Msg         : string;
  Cpt         : integer;

  function IsActive(lTn : T_TablesName) : boolean;
  var
    TableName : string;
  begin
    TableName := TUtilBTPVerdon.GetTMPTableName(lTn);
    if SettingFile.ReadString(Section, TableName, 'na') = 'na' then
      Result := True
    else
      Result := (SettingFile.ReadInteger(Section, TableName, 0) = 1);
  end;

  function SetMsg(Text, Value : string) : string;
  begin
    Result := Format('%s%s%s = %s', [Msg, #13#10, Text, Value]);
  end;

begin
  Msg         := '';
  SettingFile := TIniFile.Create(IniPath);
  try
    { Paramètres généraux }
    Section := 'GLOBALSETTINGS';
    Import.LogValues.LogLevel             := SettingFile.ReadInteger(Section, 'LogLevel', 0);                 Msg := SetMsg('LogValues.LogLevel'            , IntToStr(Import.LogValues.LogLevel));
    Import.LogValues.LogMoMaxSize         := SettingFile.ReadInteger(Section, 'LogMoMaxSize', 0);             Msg := SetMsg('LogValues.LogMoMaxSize'        , FloatToStr(Import.LogValues.LogMoMaxSize));
    Import.LogValues.DebugEvents          := SettingFile.ReadInteger(Section, 'DebugEvents', 0);              Msg := SetMsg('LogValues.DebugEvents'         , IntToStr(Import.LogValues.DebugEvents));
    Import.LogValues.OneLogPerDay         := (SettingFile.ReadInteger(Section, 'OneLogPerDay', 0) = 1);       Msg := SetMsg('LogValues.OneLogPerDay'        , BoolToStr(Import.LogValues.OneLogPerDay));
    Import.LogValues.LogPath              := LogPath;                                                         Msg := SetMsg('LogValues.LogPath'             , Import.LogValues.LogPath);
    Import.LogValues.TrueValue            := SettingFile.ReadString(Section, 'TrueValue', 'Vrai');            Msg := SetMsg('LogValues.TrueValue'           , Import.LogValues.TrueValue);
    Import.LogValues.FalseValue           := SettingFile.ReadString(Section, 'FalseValue', 'Faux');           Msg := SetMsg('LogValues.FalseValue'          , Import.LogValues.FalseValue);
    Import.LogValues.ExecutionPeriodDays  := SettingFile.ReadString(Section, 'ExecutionPeriodDays'    , '');  Msg := SetMsg('LogValues.ExecutionPeriodDays' , Import.LogValues.ExecutionPeriodDays);
    Import.LogValues.ExecutionPeriodStart := SettingFile.ReadString(Section, 'ExecutionPeriodStart'    , ''); Msg := SetMsg('LogValues.ExecutionPeriodStart', Import.LogValues.ExecutionPeriodStart);
    Import.LogValues.ExecutionPeriodEnd   := SettingFile.ReadString(Section, 'ExecutionPeriodEnd'    , '');   Msg := SetMsg('LogValues.ExecutionPeriodEnd'  , Import.LogValues.ExecutionPeriodEnd);
    { Paramètres dossier }
    Section := 'FOLDER';
    Import.FolderValues.BTPUserAdmin := SettingFile.ReadString(Section, 'BTPUser'     , ''); Msg := SetMsg('Import.FolderValues.BTPUserAdmin' , Import.FolderValues.BTPUserAdmin);
    Import.FolderValues.BTPServer    := SettingFile.ReadString(Section, 'Server'      , ''); Msg := SetMsg('Import.FolderValues.BTPServer'    , Import.FolderValues.BTPServer );
    Import.FolderValues.BTPDataBase  := SettingFile.ReadString(Section, 'BTPFolder'   , ''); Msg := SetMsg('Import.FolderValues.BTPDataBase'  , Import.FolderValues.BTPDataBase);
    Import.FolderValues.TMPServer    := SettingFile.ReadString(Section, 'Server'      , ''); Msg := SetMsg('Import.FolderValues.TMPServer'    , Import.FolderValues.TMPServer);
    Import.FolderValues.TMPDataBase  := SettingFile.ReadString(Section, 'TMPBDDFolder', ''); Msg := SetMsg('Import.FolderValues.TMPDataBase'  , Import.FolderValues.TMPDataBase);
    { Temps d'exécution }
    Section := 'IMPORT_EXECUTIONTIME';
    Import.ParametresValues.TimeOut   := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnParameters)  , 0); Msg := SetMsg('Import.ParamsValues.TimeOut'       , IntToStr(Import.ParametresValues.TimeOut));
    Import.SalariesValues.TimeOut     := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnSalaries)    , 0); Msg := SetMsg('Import.EmployeesValues.TimeOut '   , IntToStr(Import.SalariesValues.TimeOut));
    Import.ArticlesValues.TimeOut     := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnArticles)    , 0); Msg := SetMsg('Import.ArticlesValues.TimeOut '    , IntToStr(Import.ArticlesValues.TimeOut));
    Import.ModeRegleValues.TimeOut    := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnModeRegle)   , 0); Msg := SetMsg('Import.ModeRegleValues.TimeOut '   , IntToStr(Import.ModeRegleValues.TimeOut));
    Import.ReglementValues.TimeOut    := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnReglement)   , 0); Msg := SetMsg('Import.ReglementValues.TimeOut '   , IntToStr(Import.ReglementValues.TimeOut));
    Import.HresSalariesValues.TimeOut := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnHresSalaries), 0); Msg := SetMsg('Import.HresSalariesValues.TimeOut ', IntToStr(Import.HresSalariesValues.TimeOut));
    Import.ConsoStockValues.TimeOut   := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnConsoStock)  , 0); Msg := SetMsg('Import.ConsoStockValues.TimeOut '  , IntToStr(Import.ConsoStockValues.TimeOut));
    { Dernière synchronisation }
    Section := 'IMPORT_LASTSYNCHRO';
    Import.ParametresValues.LastSynchro   := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnParameters)  , ''); Msg := SetMsg('Import.ParamsValues.LastSynchro'      , Import.ParametresValues.LastSynchro );
    Import.SalariesValues.LastSynchro     := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnSalaries)    , ''); Msg := SetMsg('Import.EmployeesValues.LastSynchro'   , Import.SalariesValues.LastSynchro);
    Import.ArticlesValues.LastSynchro     := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnArticles)    , ''); Msg := SetMsg('Import.ArticlesValues.LastSynchro'    , Import.ArticlesValues.LastSynchro);
    Import.ModeRegleValues.LastSynchro    := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnModeRegle)   , ''); Msg := SetMsg('Import.ModeRegleValues.LastSynchro'   , Import.ModeRegleValues.LastSynchro);
    Import.ReglementValues.LastSynchro    := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnReglement)   , ''); Msg := SetMsg('Import.ReglementValues.LastSynchro'   , Import.ReglementValues.LastSynchro);
    Import.HresSalariesValues.LastSynchro := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnHresSalaries), ''); Msg := SetMsg('Import.HresSalariesValues.LastSynchro', Import.HresSalariesValues.LastSynchro);
    Import.ConsoStockValues.LastSynchro   := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnConsoStock), '');   Msg := SetMsg('Import.ConsoStockValues.LastSynchro'  , Import.ConsoStockValues.LastSynchro);
    { Garder les données après import }
    Section := 'IMPORT_KEEPDATA';
    Import.ParametresValues.KeepData   := (SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnParameters)  , 0) = 1); Msg := SetMsg('Import.ParamsValues.IsActive'      , BoolToStr(Import.ParametresValues.IsActive));
    Import.SalariesValues.KeepData     := (SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnSalaries)    , 0) = 1); Msg := SetMsg('Import.EmployeesValues.IsActive'   , BoolToStr(Import.SalariesValues.IsActive));
    Import.ArticlesValues.KeepData     := (SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnArticles)    , 0) = 1); Msg := SetMsg('Import.ArticlesValues.IsActive'    , BoolToStr(Import.ArticlesValues.IsActive));
    Import.ModeRegleValues.KeepData    := (SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnModeRegle)   , 0) = 1); Msg := SetMsg('Import.ModeRegleValues.IsActive'   , BoolToStr(Import.ModeRegleValues.IsActive));
    Import.ReglementValues.KeepData    := (SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnReglement)   , 0) = 1); Msg := SetMsg('Import.ReglementValues.IsActive'   , BoolToStr(Import.ReglementValues.IsActive));
    Import.HresSalariesValues.KeepData := (SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnHresSalaries), 0) = 1); Msg := SetMsg('Import.HresSalariesValues.IsActive', BoolToStr(Import.HresSalariesValues.IsActive));
    Import.ConsoStockValues.KeepData   := (SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnConsoStock), 0) = 1);   Msg := SetMsg('Import.ConsoStockValues.IsActive'  , BoolToStr(Import.ConsoStockValues.IsActive));
    { Table est active }
    Section := 'IMPORT_ISACTIVE';
    Import.ParametresValues.IsActive   := IsActive(tnParameters);   Msg := SetMsg('Import.ParamsValues.IsActive'      , BoolToStr(Import.ParametresValues.IsActive));
    Import.SalariesValues.IsActive     := IsActive(tnSalaries);     Msg := SetMsg('Import.EmployeesValues.IsActive'   , BoolToStr(Import.SalariesValues.IsActive));
    Import.ArticlesValues.IsActive     := IsActive(tnArticles);     Msg := SetMsg('Import.ArticlesValues.IsActive'    , BoolToStr(Import.ArticlesValues.IsActive));
    Import.ModeRegleValues.IsActive    := IsActive(tnModeRegle);    Msg := SetMsg('Import.ModeRegleValues.IsActive'   , BoolToStr(Import.ModeRegleValues.IsActive));
    Import.ReglementValues.IsActive    := IsActive(tnReglement);    Msg := SetMsg('Import.ReglementValues.IsActive'   , BoolToStr(Import.ReglementValues.IsActive));
    Import.HresSalariesValues.IsActive := IsActive(tnHresSalaries); Msg := SetMsg('Import.HresSalariesValues.IsActive', BoolToStr(Import.HresSalariesValues.IsActive));
    Import.ConsoStockValues.IsActive   := IsActive(tnConsoStock);   Msg := SetMsg('Import.ConsoStockValues.IsActive'  , BoolToStr(Import.ConsoStockValues.IsActive));
    Section := 'IMPORT_PARAMETERSTYPEMATCHING';
    SettingFile.ReadSection(Section, TslParametersTypeMatching);
    for Cpt := 0 to pred(TslParametersTypeMatching.Count) do
      TslParametersTypeMatching[Cpt] := Format('%s=%s', [TslParametersTypeMatching[Cpt], SettingFile.ReadString(Section, TslParametersTypeMatching[Cpt], '')]);
  finally
    Result := Msg;
    SettingFile.Free;
  end;
end;

procedure TSvcSyncBTPVerdonImp.ClearTablesValues(Import : TImportTreatment) ;
begin
  Import.ParametresValues.FirstExec   := True;
  Import.ParametresValues.Count       := 0;
  Import.ParametresValues.TimeOut     := 0;
  Import.SalariesValues.FirstExec     := True;
  Import.SalariesValues.Count         := 0;
  Import.SalariesValues.TimeOut       := 0;
  Import.ArticlesValues.FirstExec     := True;
  Import.ArticlesValues.Count         := 0;
  Import.ArticlesValues.TimeOut       := 0;
  Import.ModeRegleValues.FirstExec    := True;
  Import.ModeRegleValues.Count        := 0;
  Import.ModeRegleValues.TimeOut      := 0;
  Import.ReglementValues.FirstExec    := True;
  Import.ReglementValues.Count        := 0;
  Import.ReglementValues.TimeOut      := 0;
  Import.HresSalariesValues.FirstExec := True;
  Import.HresSalariesValues.Count     := 0;
  Import.HresSalariesValues.TimeOut   := 0;
  Import.ConsoStockValues.FirstExec   := True;
  Import.ConsoStockValues.Count       := 0;
  Import.ConsoStockValues.TimeOut     := 0;
end;

procedure TSvcSyncBTPVerdonImp.SetFieldsList(Import : TImportTreatment);
begin
  Import.ParametresFOFieldsList := Tools.GetFieldsListFromPrefix('AFO', Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
  Import.ParametresCCFieldsList := Tools.GetFieldsListFromPrefix('CC' , Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
  Import.ParametresYXFieldsList := Tools.GetFieldsListFromPrefix('YX' , Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
  Import.ArticlesFieldsList     := Tools.GetFieldsListFromPrefix('GA' , Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
  Import.SalariesFieldsList     := Tools.GetFieldsListFromPrefix('ARS', Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
  Import.ModeRegleFieldsList    := Tools.GetFieldsListFromPrefix('MP' , Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
  Import.ReglementFieldsList    := Tools.GetFieldsListFromPrefix('GAC', Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
  Import.HresSalariesFieldsList := Tools.GetFieldsListFromPrefix('BCO', Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
  if Import.HresSalariesFieldsList <> '' then
    Import.ConsoStockFieldsList := Import.HresSalariesFieldsList
  else
    Import.ConsoStockFieldsList := Tools.GetFieldsListFromPrefix('BCO', Import.FolderValues.BTPServer, Import.FolderValues.BTPDataBase);
end;

function TSvcSyncBTPVerdonImp.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TSvcSyncBTPVerdonImp.ServiceAfterInstall(Sender: TService);
var
  Reg : TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Sender.Name, false) then
    try
      Reg.WriteString('Description', 'LSE-Synchronisation des données entre BTP et VERDON - Import.');
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

{$IFDEF APPSRVWITHCBP}
procedure TSvcSyncBTPVerdonImp.ServiceExecute(Sender: TService);
var
  Import        : TImportTreatment;
  Settings      : string;
  Cpt           : integer;
  ShowMsg       : boolean;
  LastMsgPeriod : TDateTime;

  function ExecParametres : boolean;
  begin
    Import.ParametresValues.Count     := 0;
    Import.ParametresValues.FirstExec := False;
    Import.Tn                         := tnParameters;
    Import.LastSynchro                := Import.ParametresValues.LastSynchro;
    Result                            := Import.ImportTreatment;
  end;

  function ExecArticles : boolean;
  begin
    Import.ArticlesValues.Count     := 0;
    Import.ArticlesValues.FirstExec := False;
    Import.Tn                       := tnArticles;
    Import.LastSynchro              := Import.ArticlesValues.LastSynchro;
    Result                          := Import.ImportTreatment;
  end;

  function ExecSalaries : Boolean;
  begin
    Import.SalariesValues.Count     := 0;
    Import.SalariesValues.FirstExec := False;
    Import.Tn                       := tnSalaries;
    Import.LastSynchro              := Import.SalariesValues.LastSynchro;
    Result                          := Import.ImportTreatment;
  end;

  function ExecModeRegle : boolean;
  begin
    Import.ModeRegleValues.Count     := 0;
    Import.ModeRegleValues.FirstExec := False;
    Import.Tn                        := tnModeRegle;
    Import.LastSynchro               := Import.ModeRegleValues.LastSynchro;
    Result                           := Import.ImportTreatment;
  end;

  function ExecEcrRegle : boolean;
  begin
    Import.ReglementValues.Count     := 0;
    Import.ReglementValues.FirstExec := False;
    Import.Tn                        := tnReglement;
    Import.LastSynchro               := Import.ReglementValues.LastSynchro;
    Result                           := Import.ImportTreatment;
  end;

  function ExecHresSalaries : boolean;
  begin
    Import.HresSalariesValues.Count     := 0;
    Import.HresSalariesValues.FirstExec := False;
    Import.Tn                           := tnHresSalaries;
    Import.LastSynchro                  := Import.HresSalariesValues.LastSynchro;
    Result                              := Import.ImportTreatment;
  end;

  function ExecConsoStock : boolean;
  begin
    Import.ConsoStockValues.Count     := 0;
    Import.ConsoStockValues.FirstExec := False;
    Import.Tn                         := tnConsoStock;
    Import.LastSynchro                := Import.ConsoStockValues.LastSynchro;
    Result                            := Import.ImportTreatment;
  end;

begin
  AppPath := TServicesLog.GetServicesAppDataPath(True);
  if AppPath <> '' then
  begin
    IniPath := Format('%s\%s.%s', [AppPath, ServiceName_BTPVerdonIniFile, 'ini']);
    LogPath := AppPath;
    AppPath := TServicesLog.GetFilePath(ServiceName_BTPVerdonImp, 'exe');
    if not FileExists(IniPath) then
    begin
      LogMessage(Format('Impossible d''initialiser le service %s. Le fichier de configuration "%s" est inexistant.', [ServiceName_BTPVerdonExp, IniPath]), EVENTLOG_ERROR_TYPE);
    end else
    begin
      Coinitialize(nil);
      try
        TslParametersTypeMatching := TStringList.Create;
        try
          Import := TImportTreatment.Create;
          try
            ClearTablesValues(Import);
            Settings := ReadSettings(Import);
            if (Import.LogValues.DebugEvents > 0) then LogMessage(Format('Settings :%s%s', [#13#10, Settings]), EVENTLOG_INFORMATION_TYPE);
            for Cpt := 0 to pred(TslParametersTypeMatching.count) do
              Import.TslParametersTypeMatching.Add(TslParametersTypeMatching[Cpt]);
            Import.TreatmentType := 'IMPORT';
            Import.AdoQryBTP     := AdoQry.Create;
            Import.AdoQryBTP.Qry := TADOQuery.Create(nil);
            Import.AdoQryTMP     := AdoQry.Create;
            Import.AdoQryTMP.Qry := TADOQuery.Create(nil);
            Import.AssignAdoQry(Import.AdoQryBTP, Import.AdoQryTMP, Import.FolderValues, Import.LogValues);
            SetFieldsList(Import);
            {$IFNDEF TSTSRV}
            LastMsgPeriod := Now;
            while not Terminated do
            begin
              inc(Import.ParametresValues.Count);
              inc(Import.ArticlesValues.Count);
              inc(Import.SalariesValues.Count);
              inc(Import.ModeRegleValues.Count);
              inc(Import.ReglementValues.Count);
              inc(Import.HresSalariesValues.Count);
              inc(Import.ConsoStockValues.Count);
              ShowMsg := (Now > IncMinute(LastMsgPeriod, ShowMsgOutsidePeriodMinutes));
              if ShowMsg then
                LastMsgPeriod := IncMinute(LastMsgPeriod, ShowMsgOutsidePeriodMinutes);
              if TServicesLog.CanExecuteFromPeriod(Import.LogValues, ServiceName_BTPVerdonImp, ShowMsg) then
              begin
                if (Import.ParametresValues.IsActive)   and ((Import.ParametresValues.Count   >= Import.ParametresValues.TimeOut)   or (Import.ParametresValues.FirstExec))   then ExecParametres;
                if (Import.ArticlesValues.IsActive)     and ((Import.ArticlesValues.Count     >= Import.ArticlesValues.TimeOut)     or (Import.ArticlesValues.FirstExec))     then ExecArticles;
                if (Import.SalariesValues.IsActive)     and ((Import.SalariesValues.Count     >= Import.SalariesValues.TimeOut)     or (Import.SalariesValues.FirstExec))     then ExecSalaries;
                if (Import.ModeRegleValues.IsActive)    and ((Import.ModeRegleValues.Count    >= Import.ModeRegleValues.TimeOut)    or (Import.ModeRegleValues.FirstExec))    then ExecModeRegle;
                if (Import.ReglementValues.IsActive)    and ((Import.ReglementValues.Count    >= Import.ReglementValues.TimeOut)    or (Import.ReglementValues.FirstExec))    then ExecEcrRegle;
                if (Import.HresSalariesValues.IsActive) and ((Import.HresSalariesValues.Count >= Import.HresSalariesValues.TimeOut) or (Import.HresSalariesValues.FirstExec)) then ExecHresSalaries;
                if (Import.ConsoStockValues.IsActive)   and ((Import.ConsoStockValues.Count   >= Import.ConsoStockValues.TimeOut)   or (Import.ConsoStockValues.FirstExec))   then ExecConsoStock;
              end;
              Sleep(1000);
              ServiceThread.ProcessRequests(False);
            end;
            {$ELSE !TSTSRV}
            if TServicesLog.CanExecuteFromPeriod(Import.LogValues, ServiceName_BTPVerdonImp, True) then
            begin
              if (Import.ParametresValues.IsActive)   then ExecParametres;
              if (Import.ArticlesValues.IsActive)     then ExecArticles;
              if (Import.SalariesValues.IsActive)     then ExecSalaries;
              if (Import.ModeRegleValues.IsActive)    then ExecModeRegle;
              if (Import.ReglementValues.IsActive)    then ExecEcrRegle;
              if (Import.HresSalariesValues.IsActive) then ExecHresSalaries;
              if (Import.ConsoStockValues.IsActive)   then ExecConsoStock;
            end;
            {$ENDIF !TSTSRV}
          finally
            Import.Free;
          end;
        finally
          TslParametersTypeMatching.Free;
        end;        
      finally
        CoUninitialize();
      end;
    end;
  end else
    LogMessage(Format('Impossible de créer le répertoire %s.', [TServicesLog.GetServicesAppDataPath(False, 'VERDON')]), EVENTLOG_ERROR_TYPE);
end;
{$ENDIF APPSRVWITHCBP}

procedure TSvcSyncBTPVerdonImp.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  LogMessage(Format('Arrêt de %s.', [ServiceName_BTPVerdonImp]), EVENTLOG_INFORMATION_TYPE);
end;

procedure TSvcSyncBTPVerdonImp.ServiceStart(Sender: TService; var Started: Boolean);
begin
  LogMessage(Format('Démarrage de %s.', [ServiceName_BTPVerdonImp]), EVENTLOG_INFORMATION_TYPE);
end;

end.
