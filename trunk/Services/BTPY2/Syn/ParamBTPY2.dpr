program ParamBTPY2;

uses
  Forms,
  ConstServices in '..\..\..\commun\Lib\ConstServices.pas',
  CommonTools in '..\..\..\commun\Lib\CommonTools.pas',
  Ulog in '..\..\..\commun\Lib\Ulog.pas',
  Zip in '..\..\..\commun\Lib\Zip.pas',
  ZipDlls in '..\..\..\commun\Lib\ZipDlls.pas',
  UWinSystem in '..\..\..\commun\Lib\UWinSystem.pas',
  ParamSvcSynBTPY2 in '..\Lib\ParamSvcSynBTPY2.pas',
  GIFImage in '..\..\..\commun\Lib\GifImage.pas',
  tWaitingMessage in '..\..\..\commun\Lib\tWaitingMessage.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TParamSvcSyncrho, ParamSvcSyncrho);
  Application.Run;
end.
