program SvcSynBTPVerdonExp;

uses
  SvcMgr,
  uMainServiceBtpVerdonExp in '..\Lib\uMainServiceBtpVerdonExp.pas' {SvcSyncBTPVerdonImp: TService},
  CommonTools in '..\..\..\commun\Lib\CommonTools.pas',
  Zip in '..\..\..\commun\Lib\Zip.pas',
  ZipDlls in '..\..\..\commun\Lib\ZipDlls.pas',
  uExecuteServiceBtpVerdonExp in '..\Lib\uExecuteServiceBtpVerdonExp.pas',
  Ulog in '..\..\..\commun\Lib\Ulog.pas',
  ConstServices in '..\..\..\commun\Lib\ConstServices.pas',
  UtilBTPVerdon in '..\Lib\UtilBTPVerdon.pas',
  tThreadTiers in '..\Lib\tThreadTiers.pas',
  UConnectWSCEGID in '..\..\..\CONNECTWS\UConnectWSCEGID.pas',
  WinHttp_TLB in '..\..\..\CONNECTWS\WinHttp_TLB.pas',
  UConnectWSConst in '..\..\..\CONNECTWS\UConnectWSConst.pas',
  uLkJSON in '..\..\..\CONNECTWS\uLkJSON.pas',
  tThreadChantiers in '..\Lib\tThreadChantiers.pas',
  tThreadDevis in '..\Lib\tThreadDevis.pas',
  tThreadLignesBR in '..\Lib\tThreadLignesBR.pas',
  UWinSystem in '..\..\..\COMMUN\LIB\UWinSystem.pas',
  XeroBase64 in '..\..\..\CONNECTWS\XeroBase64.pas',
  XeroTypes in '..\..\..\CONNECTWS\XeroTypes.pas',
  tThreadIntervenants in '..\Lib\tThreadIntervenants.pas',
  tWaitingMessage in '..\..\..\commun\Lib\tWaitingMessage.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TSvcSyncBTPVerdonExp, SvcSyncBTPVerdonExp);
  Application.Run;
end.
