unit uMainService;

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
  , UWinSystem
  ;

type
  TSvcSyncBTPY2 = class(TService)
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure ServiceStart(Sender: TService; var Started: Boolean);

  public
    function GetServiceController: TServiceController; override;
  end;

var
  SvcSyncBTPY2: TSvcSyncBTPY2;
  
implementation

uses
  Registry
  , uThreadExecute
  , CommonTools
  , uExecuteService
  , ActiveX
  , WinSVC
  , UConnectWSConst
  , ShellAPI
  , ConstServices
  , DateUtils
  ;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  SvcSyncBTPY2.Controller(CtrlCode);
end;

function TSvcSyncBTPY2.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TSvcSyncBTPY2.ServiceAfterInstall(Sender: TService);
var
  Reg : TRegistry;
begin                                                                               
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Sender.Name, false) then
    try
      Reg.WriteString('Description', 'LSE-Synchronisation des données comptables entre BTP et Y2.');
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;                                                                              
  end;
end;                                                                                    

procedure TSvcSyncBTPY2.ServiceExecute(Sender: TService);
var
  Count         : Integer;
  BTPY2Exec     : TSvcSyncBTPY2Execute;
  IniPath       : string;
  AppPath       : string;
  LogPath       : string;
  FirstExec     : boolean;
  ShowMsg       : boolean;
  ForceRetry    : boolean;
  LastMsgPeriod : TDateTime;
begin
  ForceRetry := false;
  AppPath := TServicesLog.GetServicesAppDataPath(True);
  if AppPath <> '' then
  begin
    IniPath := Format('%s\%s.%s', [AppPath, ServiceName_BTPY2, 'ini']);
    LogPath := AppPath;
    AppPath := TServicesLog.GetFilePath(ServiceName_BTPY2, 'exe');
    if not FileExists(IniPath) then
    begin
      LogMessage(Format('Impossible d''initialiser le service %s. Le fichier de configuration "%s" est inexistant.', [ServiceName_BTPY2, IniPath]), EVENTLOG_ERROR_TYPE);
    end else
    begin
      FirstExec := True;
      BTPY2Exec := TSvcSyncBTPY2Execute.Create;
      try
        BTPY2Exec.CreateObjects;
        try
          BTPY2Exec.IniFilePath := IniPath;
          BTPY2Exec.AppFilePath := AppPath;
          BTPY2Exec.LogFilePath := LogPath;
          if BTPY2Exec.InitApplication then
          begin
            Count         := 0;
            LastMsgPeriod := Now;
            while not Terminated do
            begin
              Inc(Count);
              ShowMsg := (Now > IncMinute(LastMsgPeriod, ShowMsgOutsidePeriodMinutes));
              if ShowMsg then
                LastMsgPeriod := IncMinute(LastMsgPeriod, ShowMsgOutsidePeriodMinutes);
              if (TServicesLog.CanExecuteFromPeriod(BTPY2Exec.LogValues, ServiceName_BTPY2, ShowMsg)) and ((Count >= BTPY2Exec.SecondTimeout) or (FirstExec) or  (ForceRetry and (Count >= 20))) then
              begin
                FirstExec := False;
                Count     := 0;
                try
                  LogMessage('Début d''exécution du service.', EVENTLOG_INFORMATION_TYPE);
                  ForceRetry := (not BTPY2Exec.ServiceExecute);
                  try
                    BTPY2Exec.LogFilePath := LogPath;
                  finally
                    LogMessage('Fin d''exécution du service.', EVENTLOG_INFORMATION_TYPE);
                  end;
                except
                  on E: Exception do
                    LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
                end;
              end;
              Sleep(1000);
              ServiceThread.ProcessRequests(False);
              {$IFDEF TSTSRV}
              ShellExecute(0,'NET','STOP Synchronisation BTP Y2"',nil,nil,SW_HIDE);
              {$ENDIF TSTSRV}
            end;
          end;
        finally
          BTPY2Exec.FreeObjects;
        end;
      finally
        BTPY2Exec.Free;
      end;
    end;
  end;
end;

procedure TSvcSyncBTPY2.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  CoUnInitialize;
  LogMessage(Format('Arrêt de %s.', [ServiceName_BTPY2]), EVENTLOG_INFORMATION_TYPE);
end;

procedure TSvcSyncBTPY2.ServiceStart(Sender: TService; var Started: Boolean);
begin
  LogMessage(Format('Démarrage de de %s.', [ServiceName_BTPY2]), EVENTLOG_INFORMATION_TYPE);
  Coinitialize(nil);
end;

end.
