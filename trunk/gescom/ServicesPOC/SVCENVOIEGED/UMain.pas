unit UMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, SvcMgr, Dialogs,UTob;

type
  TLSESVCENVOIBAST = class(TService)
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
  private
    { Déclarations privées }
  public
    function GetServiceController: TServiceController; override;
    { Déclarations publiques }
  end;

var
  LSESVCENVOIBAST: TLSESVCENVOIBAST;
  
implementation

uses
  Registry
  , CommonTools
  , ActiveX
  , WinSVC
  , ShellAPI
  , ConstServices
  , HEnt1
  , Hdb
  , UtilEnvEnvoiGed
  , Ulog
  , UconnectBSV
  ;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  LSESVCENVOIBAST.Controller(CtrlCode);
end;

function TLSESVCENVOIBAST.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TLSESVCENVOIBAST.ServiceAfterInstall(Sender: TService);
var
  Reg : TRegistry;
begin                                                                               
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Sender.Name, false) then
    try
      Reg.WriteString('Description', 'LSE Enregistrement des BAST dans GED BSV');
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TLSESVCENVOIBAST.ServiceExecute(Sender: TService);
var
  Count     : Integer;
  EnvEnvoiGed : TEnvEnvoiGed;
  IniPath   : string;
  AppPath   : string;
  LogPath   : string;
  FirstExec : boolean;
  TOBFILES : TOB;
  TOBFields : TOB;
begin
  IniPath := IncludeTrailingBackslash(TServicesLog.GetServicesAppDataPath(True,ServiceName_BASTVERSGED))+ServiceName_BASTVERSGED+'.ini';
  AppPath := TServicesLog.GetFilePath(ServiceName_BASTVERSGED, 'exe');
  LogPath := IncludeTrailingBackslash(TServicesLog.GetServicesAppDataPath(True,ServiceName_BASTVERSGED))+ServiceName_BASTVERSGED+'.log';
  if not FileExists(IniPath) then
  begin
    LogMessage(Format('Impossible d''initialiser le service %s. Le fichier de configuration "%s" est inexistant.', [ServiceName_BASTVERSGED,IniPath]), EVENTLOG_ERROR_TYPE);
  end else
  begin
    FirstExec := True;
    EnvEnvoiGed := TEnvEnvoiGed.create (IniPath);
    try
      if not EnvEnvoiGed.Status then
      begin
        LogMessage(Format('Impossible d''exécuter le service %s. Le fichier de configuration "%s" est incorrect.', [ServiceName_BASTVERSGED, IniPath]), EVENTLOG_ERROR_TYPE);
        Exit;
      end;
      try
        Count := 0;
        TOBFILES := TOB.create('LES FICHIERS',nil,-1);
        TOBFields := TOB.create('LES PARAMS',nil,-1);
        GetParamStockageBSV (TOBFields,'XBT',EnvEnvoiGed.Server,EnvEnvoiGed.Database,EnvEnvoiGed.ModeDebug,LogPath);
        if TOBFields.Detail.count = 0 then
        begin
          LogMessage('Veuillez parametrer le BAST pour la liaison avec BSV', EVENTLOG_ERROR_TYPE);
          Exit;
        end;
        //
//        if Tools.GetparamSocSecur ('
        while not Terminated do
        begin
          TOBFILES := TOB.create('LES FICHIERS',nil,-1);
          Inc(Count);
          if (Count >= EnvEnvoiGed.delay) or (FirstExec) then
          begin
            FirstExec := False;
            Count     := 0;
            try
              if EnvEnvoiGed.ModeDebug > 0 then  EcritLogs(LogPath,'4. Début d''exécution du service.');
              try
              finally
                if EnvEnvoiGed.ModeDebug > 0 then EcritLogs(LogPath,'5. Fin d''exécution du service.');
              end;
            except
              on E: Exception do
                LogMessage(Format('Fin exécution du service avec erreur : %s', [E.Message]), EVENTLOG_ERROR_TYPE);
            end;
          end;
          ServiceThread.ProcessRequests(True);
          Sleep(1000);
        end;
      finally
        TOBFILES.free;
        TOBFields.free;
      end;
    finally
      LogMessage('Deconnexion fin Service.',EVENTLOG_INFORMATION_TYPE);
    end;
  end;
end;

procedure TLSESVCENVOIBAST.ServiceStart(Sender: TService;var Started: Boolean);
begin
  LogMessage('Démarrage du service d''envoi des BAST dans GED.', EVENTLOG_INFORMATION_TYPE);
  Coinitialize(nil);
end;

procedure TLSESVCENVOIBAST.ServiceStop(Sender: TService;var Stopped: Boolean);
begin
  CoUnInitialize;
  LogMessage('Arrêt du service d''envoi des BAST dans GED.', EVENTLOG_INFORMATION_TYPE);
end;

end.
