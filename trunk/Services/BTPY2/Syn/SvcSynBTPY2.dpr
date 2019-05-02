program SvcSynBTPY2;

uses
  SvcMgr,
  uMainService in '..\Lib\uMainService.pas' {SvcSyncBTPY2: TService},
  uThreadExecute in '..\Lib\uThreadExecute.pas',
  uWSDataService in '..\Lib\uWSDataService.pas',
  CommonTools in '..\..\..\commun\Lib\CommonTools.pas',
  WinHttp_TLB in '..\..\..\CONNECTWS\WinHttp_TLB.pas',
  uLkJSON in '..\..\..\CONNECTWS\uLkJSON.pas',
  UConnectWSConst in '..\..\..\CONNECTWS\UConnectWSConst.pas',
  uExecuteService in '..\Lib\uExecuteService.pas',
  UConnectWSCEGID in '..\..\..\CONNECTWS\UConnectWSCEGID.pas',
  TRAFileUtil in '..\..\..\commun\Lib\TRAFileUtil.pas',
  Ulog in '..\..\..\commun\Lib\Ulog.pas',
  Zip in '..\..\..\commun\Lib\Zip.pas',
  ZipDlls in '..\..\..\commun\Lib\ZipDlls.pas',
  ConstServices in '..\..\..\commun\Lib\ConstServices.pas',
  UWinSystem in '..\..\..\COMMUN\LIB\UWinSystem.pas',
  XeroBase64 in '..\..\..\CONNECTWS\XeroBase64.pas',
  XeroTypes in '..\..\..\CONNECTWS\XeroTypes.pas',
  tWaitingMessage in '..\..\..\commun\Lib\tWaitingMessage.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'LSE Synchronisation BTP Y2';
  Application.CreateForm(TSvcSyncBTPY2, SvcSyncBTPY2);
  Application.Run;
end.                                                                                                
