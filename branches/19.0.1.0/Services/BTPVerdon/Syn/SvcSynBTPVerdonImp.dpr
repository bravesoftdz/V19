program SvcSynBTPVerdonImp;

uses
  SvcMgr,
  uMainServiceBtpVerdonImp in '..\Lib\uMainServiceBtpVerdonImp.pas' {SvcSyncBTPVerdonImp: TService},
  CommonTools in '..\..\..\commun\Lib\CommonTools.pas',
  Zip in '..\..\..\commun\Lib\Zip.pas',
  ZipDlls in '..\..\..\commun\Lib\ZipDlls.pas',
  uExecuteServiceBtpVerdonImp in '..\Lib\uExecuteServiceBtpVerdonImp.pas',
  Ulog in '..\..\..\commun\Lib\Ulog.pas',
  ConstServices in '..\..\..\commun\Lib\ConstServices.pas',
  UtilBTPVerdon in '..\Lib\UtilBTPVerdon.pas',
  WinHttp_TLB in '..\..\..\CONNECTWS\WinHttp_TLB.pas',
  UConnectWSConst in '..\..\..\CONNECTWS\UConnectWSConst.pas',
  uLkJSON in '..\..\..\CONNECTWS\uLkJSON.pas',
  UWinSystem in '..\..\..\commun\Lib\UWinSystem.pas',
  tWaitingMessage in '..\..\..\commun\Lib\tWaitingMessage.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TSvcSyncBTPVerdonImp, SvcSyncBTPVerdonImp);
  Application.Run;
end.
