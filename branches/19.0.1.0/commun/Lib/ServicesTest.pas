unit ServicesTest;

interface

 {$IFDEF APPSRV}
type
  T_SvcTestExecute = class (TObject)
    class procedure SvcBtpY2;
    class procedure SvcBtpToVerdon;
    class procedure SvcVerdonToBtp;
  end;
 {$ENDIF APPSRV}                                                          

implementation

uses
   SysUtils
  , CommonTools
  , hMsgBox
  {$IFDEF APPSRV}
  , Forms
  , uExecuteService
  , uMainService
  , uExecuteServiceBtpVerdonImp
  , uMainServiceBtpVerdonImp
  , uExecuteServiceBtpVerdonExp
  , uMainServiceBtpVerdonExp
  , ConstServices
  {$ENDIF APPSRV}
  ;

{ T_SvcTestExecute }

{$IFDEF APPSRV}
class procedure T_SvcTestExecute.SvcBtpY2;
var
  BTPY2Exec : TSvcSyncBTPY2Execute;
  AppName   : string;
  IniPath   : string;
  AppPath   : string;
  LogPath   : string;
begin
  { Test du service }
  AppPath := TServicesLog.GetServicesAppDataPath(True);
  if AppPath <> '' then
  begin
    AppName   := ExtractFilePath(Application.ExeName);
    BTPY2Exec := TSvcSyncBTPY2Execute.Create;
    try
      IniPath := Format('%s\%s.%s', [AppPath, ServiceName_BTPY2, 'ini']);
      LogPath := AppPath;
      AppPath := TServicesLog.GetFilePath(ServiceName_BTPY2, 'exe');
      if not FileExists(IniPath) then
      begin
        PGIError(Format('Impossible d''initialiser le service %s. Le fichier de configuration "%s" est inexistant.', [ServiceName_BTPY2, IniPath]));
      end else
      begin
        BTPY2Exec.IniFilePath := IniPath;
        BTPY2Exec.AppFilePath := AppPath;
        BTPY2Exec.LogFilePath := LogPath;
        BTPY2Exec.CreateObjects;
        try
          BTPY2Exec.InitApplication;
          try
           BTPY2Exec.ServiceExecute;
          finally
          end;
        finally
          BTPY2Exec.FreeObjects;
        end;
      end;
    finally
      BTPY2Exec.Free;
    end;
  end;
end;
{$ENDIF APPSRV}

{$IFDEF APPSRV}
class procedure T_SvcTestExecute.SvcBtpToVerdon;
var
  BTPVerdonExec : TSvcSyncBTPVerdonExp;
begin
  BTPVerdonExec := TSvcSyncBTPVerdonExp.Create(nil);
  try
    BTPVerdonExec.ServiceExecute(nil);
  finally
    BTPVerdonExec.Free;
  end;

end;
{$ENDIF APPSRV}

{$IFDEF APPSRV}
class procedure T_SvcTestExecute.SvcVerdonToBtp;
var
  BTPVerdonExec : TSvcSyncBTPVerdonImp;
begin
  BTPVerdonExec := TSvcSyncBTPVerdonImp.Create(nil);
  try
    BTPVerdonExec.ServiceExecute(nil);
  finally
    BTPVerdonExec.Free;
  end;
end;
{$ENDIF APPSRV}

end.
