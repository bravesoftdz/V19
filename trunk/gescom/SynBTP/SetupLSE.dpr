program SetupLSE;

uses
  Forms,
  UMainControl in '..\..\Lanceur Install\UMainControl.pas' {Form1},
  UControlInstall in '..\..\Lanceur Install\UControlInstall.pas' {FControlInstall};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Setup LSE';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
