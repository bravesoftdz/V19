program SVCEnvoiBASTGed;

uses
  SvcMgr,
  Hent1,  
  UMain in '..\ServicesPOC\SVCENVOIEGED\UMain.pas' {LSESVCENVOIBAST: TService},
  CommonTools in '..\..\commun\Lib\CommonTools.pas',
  Zip in '..\..\commun\Lib\Zip.pas',
  ZipDlls in '..\..\COMMUN\LIB\ZipDlls.pas',
  UconnectBSV in '..\..\CONNECTWS\UconnectBSV.pas',
  WinHttp_TLB in '..\..\CONNECTWS\WinHttp_TLB.pas',
  XeroTypes in '..\..\CONNECTWS\XeroTypes.pas',
  XeroBase64 in '..\..\CONNECTWS\XeroBase64.pas',
  UdefServices in '..\..\CONNECTWS\UdefServices.pas',
  UCryptage in '..\..\WebServices\Commun\Lib\UCryptage.pas',
  UtilBSV in '..\LibBTP\UtilBSV.pas',
  UDefGlobals in '..\LibBTP\UDefGlobals.pas',
  UWinSystem in '..\..\COMMUN\LIB\UWinSystem.pas',
  UtilEnvEnvoiGed in '..\ServicesPOC\SVCENVOIEGED\UtilEnvEnvoiGed.pas',
  Ulog in '..\..\commun\Lib\Ulog.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Service envoie BAST dans GED POC';
  Application.CreateForm(TLSESVCENVOIBAST, LSESVCENVOIBAST);
  Application.Run;
end.
