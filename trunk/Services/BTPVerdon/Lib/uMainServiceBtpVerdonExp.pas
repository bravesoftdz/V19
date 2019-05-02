unit uMainServiceBtpVerdonExp;

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
  , tThreadTiers
  , tThreadChantiers
  , tThreadDevis
  , tThreadLignesBR
  , tThreadIntervenants
  ;

type
  TSvcSyncBTPVerdonExp = class(TService)
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
  private
    uThreadTiers       : ThreadTiers;
    uThreadChantier    : ThreadChantiers;
    uThreadDevis       : ThreadDevis;
    uThreadLignesBR    : ThreadLignesBR;
    uThreadIntervenants : ThreadIntervenants;
    IniPath            : string;
    LogPath            : string;
    LogValues          : T_WSLogValues;
    TiersValues        : T_TiersValues;
    ChantierValues     : T_ChantierValues;
    DevisValues        : T_DevisValues;
    LignesBRValues     : T_LignesBRValues;
    FolderValues       : T_FolderValues;
    IntervenantsValues : T_IntervenantsValues;

    procedure ClearTablesValues;
    function ReadSettings : string;

  public
    function GetServiceController: TServiceController; override;
  end;

var
  SvcSyncBTPVerdonExp: TSvcSyncBTPVerdonExp;

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
  SvcSyncBTPVerdonExp.Controller(CtrlCode);
end;

procedure TSvcSyncBTPVerdonExp.ClearTablesValues;
begin
  TiersValues.FirstExec        := True;
  TiersValues.Count            := 0;
  TiersValues.TimeOut          := 0;
  ChantierValues.FirstExec     := True;
  ChantierValues.Count         := 0;
  ChantierValues.TimeOut       := 0;
  DevisValues.FirstExec        := True;
  DevisValues.Count            := 0;
  DevisValues.TimeOut          := 0;
  LignesBRValues.FirstExec     := True;
  LignesBRValues.Count         := 0;
  LignesBRValues.TimeOut       := 0;
  IntervenantsValues.FirstExec := True;
  IntervenantsValues.Count     := 0;
  IntervenantsValues.TimeOut   := 0;
end;

function TSvcSyncBTPVerdonExp.ReadSettings : string;
var
  SettingFile : TInifile;
  Section     : string;
  Msg         : string;

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
    LogValues.LogLevel             := SettingFile.ReadInteger(Section, 'LogLevel', 0);                  Msg := SetMsg('LogValues.LogLevel'            , IntToStr(LogValues.LogLevel));
    LogValues.LogMoMaxSize         := SettingFile.ReadInteger(Section, 'LogMoMaxSize', 0);              Msg := SetMsg('LogValues.LogMoMaxSize'        , FloatToStr(LogValues.LogMoMaxSize));
    LogValues.DebugEvents          := SettingFile.ReadInteger(Section, 'DebugEvents', 0);               Msg := SetMsg('LogValues.DebugEvents'         , IntToStr(LogValues.DebugEvents));
    LogValues.OneLogPerDay         := (SettingFile.ReadInteger(Section,'OneLogPerDay', 0) = 1);         Msg := SetMsg('LogValues.OneLogPerDay'        , BoolToStr(LogValues.OneLogPerDay));
    LogValues.LogPath              := LogPath;                                                          Msg := SetMsg('LogValues.LogPath'             , LogValues.LogPath);
    LogValues.TrueValue            := SettingFile.ReadString(Section, 'TrueValue', 'Vrai');             Msg := SetMsg('LogValues.TrueValue'           , LogValues.TrueValue);
    LogValues.FalseValue           := SettingFile.ReadString(Section, 'FalseValue', 'Faux');            Msg := SetMsg('LogValues.FalseValue'          , LogValues.FalseValue);
    LogValues.ExecutionPeriodDays  := SettingFile.ReadString(Section, 'ExecutionPeriodDays'    , '');   Msg := SetMsg('LogValues.ExecutionPeriodDays' , LogValues.ExecutionPeriodDays);
    LogValues.ExecutionPeriodStart := SettingFile.ReadString(Section, 'ExecutionPeriodStart'    , '');  Msg := SetMsg('LogValues.ExecutionPeriodStart', LogValues.ExecutionPeriodStart);
    LogValues.ExecutionPeriodEnd   := SettingFile.ReadString(Section, 'ExecutionPeriodEnd'    , '');    Msg := SetMsg('LogValues.ExecutionPeriodEnd'  , LogValues.ExecutionPeriodEnd);
    { Paramètres du dossier }
    Section := 'FOLDER';
    FolderValues.BTPUserAdmin := SettingFile.ReadString(Section, 'BTPUser'     , ''); Msg := SetMsg('FolderValues.BTPUserAdmin', FolderValues.BTPUserAdmin);
    FolderValues.BTPServer    := SettingFile.ReadString(Section, 'Server'      , ''); Msg := SetMsg('FolderValues.BTPServer'   , FolderValues.BTPServer);
    FolderValues.BTPDataBase  := SettingFile.ReadString(Section, 'BTPFolder'   , ''); Msg := SetMsg('FolderValues.BTPDataBase' , FolderValues.BTPDataBase);
    FolderValues.TMPServer    := SettingFile.ReadString(Section, 'Server'      , ''); Msg := SetMsg('FolderValues.TMPServer'   , FolderValues.TMPServer);
    FolderValues.TMPDataBase  := SettingFile.ReadString(Section, 'TMPBDDFolder', ''); Msg := SetMsg('FolderValues.TMPDataBase' , FolderValues.TMPDataBase);
    { Délai d'exécution }
    Section := 'EXPORT_EXECUTIONTIME';
    TiersValues.TimeOut        := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnTiers)   , 0);     Msg := SetMsg('TiersValues.TimeOut'       , IntToStr(TiersValues.TimeOut));
    ChantierValues.TimeOut     := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnChantier), 0);     Msg := SetMsg('ChantierValues.TimeOut'    , IntToStr(ChantierValues.TimeOut));
    DevisValues.TimeOut        := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnDevis), 0);        Msg := SetMsg('DevisValues.TimeOut'       , IntToStr(DevisValues.TimeOut));
    LignesBRValues.TimeOut     := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnLignesBR), 0);     Msg := SetMsg('LignesBRValues.TimeOut'    , IntToStr(LignesBRValues.TimeOut));
    IntervenantsValues.TimeOut := SettingFile.ReadInteger(Section, TUtilBTPVerdon.GetTMPTableName(tnIntervenants), 0); Msg := SetMsg('IntervenantsValues.TimeOut', IntToStr(IntervenantsValues.TimeOut));
    { Dernière synchronisation }
    Section := 'EXPORT_LASTSYNCHRO';
    TiersValues.LastSynchro        := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnTiers)   , '');     Msg := SetMsg('TiersValues.LastSynchro'       , TiersValues.LastSynchro);
    ChantierValues.LastSynchro     := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnChantier), '');     Msg := SetMsg('ChantierValues.LastSynchro'    , ChantierValues.LastSynchro);
    DevisValues.LastSynchro        := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnDevis), '');        Msg := SetMsg('DevisValues.LastSynchro'       , DevisValues.LastSynchro);
    LignesBRValues.LastSynchro     := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnLignesBR), '');     Msg := SetMsg('LignesBRValues.LastSynchro'    , LignesBRValues.LastSynchro);
    IntervenantsValues.LastSynchro := SettingFile.ReadString(Section, TUtilBTPVerdon.GetTMPTableName(tnIntervenants), ''); Msg := SetMsg('IntervenantsValues.LastSynchro', IntervenantsValues.LastSynchro);
    { Table active }
    Section := 'EXPORT_ISACTIVE';
    TiersValues.IsActive        := IsActive(tnTiers);        Msg := SetMsg('TiersValues.IsActive'       , Tools.iif(TiersValues.IsActive       , LogValues.TrueValue, LogValues.FalseValue));
    ChantierValues.IsActive     := IsActive(tnChantier);     Msg := SetMsg('ChantierValues.IsActive'    , Tools.iif(ChantierValues.IsActive    , LogValues.TrueValue, LogValues.FalseValue));
    DevisValues.IsActive        := IsActive(tnDevis);        Msg := SetMsg('DevisValues.IsActive'       , Tools.iif(DevisValues.IsActive       , LogValues.TrueValue, LogValues.FalseValue));
    LignesBRValues.IsActive     := IsActive(tnLignesBR);     Msg := SetMsg('LignesBRValues.IsActive'    , Tools.iif(LignesBRValues.IsActive    , LogValues.TrueValue, LogValues.FalseValue));
    IntervenantsValues.IsActive := IsActive(tnIntervenants); Msg := SetMsg('IntervenantsValues.IsActive', Tools.iif(IntervenantsValues.IsActive, LogValues.TrueValue, LogValues.FalseValue));
  finally
    SettingFile.Free;
  end;
end;

function TSvcSyncBTPVerdonExp.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TSvcSyncBTPVerdonExp.ServiceAfterInstall(Sender: TService);
var
  Reg : TRegistry;
begin                                                                               
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Sender.Name, false) then
    try
      Reg.WriteString('Description', 'LSE-Synchronisation des données entre BTP et VERDON - Export.');
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;                                                                              
  end;
end;

procedure TSvcSyncBTPVerdonExp.ServiceExecute(Sender: TService);
var
  AppPath       : string;
  Settings      : string;
  ShowMsg       : boolean;
  LastMsgPeriod : TDateTime;

  procedure StartLog(lTn : T_TablesName; LastSynchro : string);
  begin
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, lTn, '', LogValues, 0);
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, lTn, DupeString('*', 50), LogValues, 0);
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, True, LastSynchro), LogValues, 0);
  end;

  procedure CallThreadTiers;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnTiers, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    TUtilBTPVerdon.StartLog(ServiceName_BTPVerdonExp, tnTiers, LogValues, TiersValues.LastSynchro);
    TiersValues.FirstExec        := False;
    TiersValues.Count            := 0;
    uThreadTiers                 := ThreadTiers.Create(True);
    uThreadTiers.FreeOnTerminate := True;
    uThreadTiers.Priority        := tpNormal;
    uThreadTiers.lTn             := tnTiers;
    uThreadTiers.TableValues     := TiersValues;
    uThreadTiers.LogValues       := LogValues;
    uThreadTiers.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnTiers, Format('%sBefore Call uThreadTiers.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadTiers.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

  procedure CallThreadChantiers;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnChantier, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    StartLog(tnChantier, ChantierValues.LastSynchro);
    ChantierValues.FirstExec        := False;
    ChantierValues.Count            := 0;
    uThreadChantier                 := ThreadChantiers.Create(True);
    uThreadChantier.FreeOnTerminate := True;
    uThreadChantier.Priority        := tpNormal;
    uThreadChantier.lTn             := tnChantier;
    uThreadChantier.TableValues     := ChantierValues;
    uThreadChantier.LogValues       := LogValues;
    uThreadChantier.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnChantier, Format('%sBefore Call uThreadChantier.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadChantier.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

  procedure CallThreadDevis;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnDevis, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    StartLog(tnDevis, DevisValues.LastSynchro);
    DevisValues.FirstExec        := False;
    DevisValues.Count            := 0;
    uThreadDevis                 := ThreadDevis.Create(True);
    uThreadDevis.FreeOnTerminate := True;
    uThreadDevis.Priority        := tpNormal;
    uThreadDevis.lTn             := tnDevis;
    uThreadDevis.TableValues     := DevisValues;
    uThreadDevis.LogValues       := LogValues;
    uThreadDevis.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnDevis, Format('%sBefore Call uThreadDevis.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadDevis.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

  procedure CallThreadLignesBR;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnLignesBR, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    StartLog(tnLignesBR, DevisValues.LastSynchro);
    LignesBRValues.FirstExec        := False;
    LignesBRValues.Count            := 0;
    uThreadLignesBR                 := ThreadLignesBR.Create(True);
    uThreadLignesBR.FreeOnTerminate := True;
    uThreadLignesBR.Priority        := tpNormal;
    uThreadLignesBR.lTn             := tnLignesBR;
    uThreadLignesBR.TableValues     := LignesBRValues;
    uThreadLignesBR.LogValues       := LogValues;
    uThreadLignesBR.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnLignesBR, Format('%sBefore Call uThreadLignesBR.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadLignesBR.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

  procedure CallThreadIntervenants;
  begin
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnIntervenants, Format('%sWith Thread', [WSCDS_DebugMsg]), LogValues, 0);
    StartLog(tnIntervenants, DevisValues.LastSynchro);
    IntervenantsValues.FirstExec        := False;
    IntervenantsValues.Count            := 0;
    uThreadIntervenants                 := ThreadIntervenants.Create(True);
    uThreadIntervenants.FreeOnTerminate := True;
    uThreadIntervenants.Priority        := tpNormal;
    uThreadIntervenants.lTn             := tnLignesBR;
    uThreadIntervenants.TableValues     := IntervenantsValues;
    uThreadIntervenants.LogValues       := LogValues;
    uThreadIntervenants.FolderValues    := FolderValues;
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, tnIntervenants, Format('%sBefore Call uThreadLignesBR.Resume', [WSCDS_DebugMsg]), LogValues, 0);
    try
      uThreadIntervenants.Resume;
    except
      on E: Exception do
        LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
    end;
  end;

begin
  AppPath := TServicesLog.GetServicesAppDataPath(True);
  if AppPath <> '' then
  begin
    IniPath := Format('%s\%s.%s', [AppPath, ServiceName_BTPVerdonIniFile, 'ini']);
    LogPath := AppPath; 
    AppPath := TServicesLog.GetFilePath(ServiceName_BTPVerdonExp, 'exe');
    if not FileExists(IniPath) then
    begin
      LogMessage(Format('Impossible d''initialiser le service %s. Le fichier de configuration "%s" est inexistant.', [ServiceName_BTPVerdonExp, IniPath]), EVENTLOG_ERROR_TYPE);
    end else
    begin
      ClearTablesValues;
      Settings := ReadSettings;
      if (LogValues.DebugEvents > 0) then LogMessage(Format('Settings :%s%s', [#13#10, Settings]), EVENTLOG_INFORMATION_TYPE);
      {$IFNDEF TSTSRV}
      LastMsgPeriod := Now;
      while not Terminated do
      begin
        Inc(TiersValues.Count);
        Inc(ChantierValues.Count);
        Inc(DevisValues.Count);
        Inc(LignesBRValues.Count);
        Inc(IntervenantsValues.Count);
        ShowMsg := (Now > IncMinute(LastMsgPeriod, ShowMsgOutsidePeriodMinutes));
        if ShowMsg then
          LastMsgPeriod := IncMinute(LastMsgPeriod, ShowMsgOutsidePeriodMinutes);
        if TServicesLog.CanExecuteFromPeriod(LogValues, ServiceName_BTPVerdonExp, ShowMsg) then
        begin
          if (TiersValues.IsActive)        and ((TiersValues.Count        >= TiersValues.TimeOut)        or (TiersValues.FirstExec))        then CallThreadTiers;
          if (ChantierValues.IsActive)     and ((ChantierValues.Count     >= ChantierValues.TimeOut)     or (ChantierValues.FirstExec))     then CallThreadChantiers;
          if (DevisValues.IsActive)        and ((DevisValues.Count        >= DevisValues.TimeOut)        or (DevisValues.FirstExec))        then CallThreadDevis;
          if (LignesBRValues.IsActive)     and ((LignesBRValues.Count     >= LignesBRValues.TimeOut)     or (LignesBRValues.FirstExec))     then CallThreadLignesBR;
          if (IntervenantsValues.IsActive) and ((IntervenantsValues.Count >= IntervenantsValues.TimeOut) or (IntervenantsValues.FirstExec)) then CallThreadIntervenants;
        end;
        Sleep(1000);
        ServiceThread.ProcessRequests(False);
      end;
      {$ELSE !TSTSRV}
      if TServicesLog.CanExecuteFromPeriod(LogValues, ServiceName_BTPVerdonExp, True) then
      begin
        if (TiersValues.IsActive)        then CallThreadTiers;
        if (ChantierValues.IsActive)     then CallThreadChantiers;
        if (DevisValues.IsActive)        then CallThreadDevis;
        if (LignesBRValues.IsActive)     then CallThreadLignesBR;
        if (IntervenantsValues.IsActive) then CallThreadIntervenants;
      end;
      {$ENDIF !TSTSRV}
    end;
  end else
    LogMessage(Format('Impossible de créer le répertoire %s.', [TServicesLog.GetServicesAppDataPath(False, 'VERDON')]), EVENTLOG_ERROR_TYPE);
end;

procedure TSvcSyncBTPVerdonExp.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  FreeAndNil(uThreadTiers);
  FreeAndNil(uThreadChantier);
  FreeAndNil(uThreadDevis);
  FreeAndNil(uThreadLignesBR);
  LogMessage(Format('Arrêt de %s.', [ServiceName_BTPVerdonExp]), EVENTLOG_INFORMATION_TYPE);
end;

procedure TSvcSyncBTPVerdonExp.ServiceStart(Sender: TService; var Started: Boolean);
begin
  LogMessage(Format('Démarrage de %s.', [ServiceName_BTPVerdonExp]), EVENTLOG_INFORMATION_TYPE);
end;

end.

