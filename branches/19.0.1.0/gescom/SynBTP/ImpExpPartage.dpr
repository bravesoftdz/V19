program ImpExpPartage;

uses
  SysUtils,
  Forms,
  UPrincipale in '..\impexppartage\UPrincipale.pas' {Fprincipale},
  UPartageExport in '..\impexppartage\UPartageExport.pas' {FExportData},
  UImportDatas in '..\impexppartage\UImportDatas.pas' {FImportDatas},
  UShareDB in '..\..\CommunBTP\LIB\UShareDB.pas',
  Ulog in '..\..\commun\Lib\Ulog.pas',
  UdefImportDatas in '..\impexppartage\UdefImportDatas.pas';

//

// FIN NEW

{$R *.RES}

begin
{$ifdef MEMCHECK}
  MemCheckLogFileName:=ChangeFileExt(Application.exename,'.log');
  MemChk;
{$endif}
  Application.Initialize;
  Application.Title := 'Import Export Partage CEGID';
  Application.CreateForm(TFprincipale, Fprincipale);
  Application.Run;
end.
