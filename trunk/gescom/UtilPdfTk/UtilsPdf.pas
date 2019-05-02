unit UtilsPdf;

interface
uses
  Classes,
  SysUtils,
  Windows,
  ShellApi,
  HEnt1
  ;

procedure RegroupePdf (Pdf1,Pdf2,PdfOut : string);

implementation

procedure RegroupePdf (Pdf1,Pdf2,PdfOut : string);
var Param : string;
begin
  Param := Format('pdftk.exe %s %s cat output %s',[Pdf1,Pdf2,PdfOut]);
  FileExecAndWait (Param);
end;

end.
