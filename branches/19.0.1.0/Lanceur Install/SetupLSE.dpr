program SetupLSE;

uses
  Forms,
  UMainControl in 'UMainControl.pas' {Form1},
  UControlInstall in 'UControlInstall.pas' {FControlInstall},
  HManifest in '..\COMMUN\Lib\HManifest.pas',
  UInfoParams in 'UInfoParams.pas' {InfoParams},
  UCommonParams in 'UCommonParams.pas';

{$R *.res}

var
  II : Integer;

begin
  ModeDebug := false;
  ModeForce := False;
  
  Application.Initialize;
  Application.Title := 'Setup LSE';
  if ParamCount > 0 then
  begin
    for II := 0 to ParamCount do
    begin
      if ParamStr(II)='-DEBUG' then ModeDebug := True;
      if ParamStr(II)='-FORCE' then ModeForce := True;
    end;
  end;
  Application.CreateForm(TForm1, Form1);
  ModeDebug := ModeDebug;
  ModeForce := ModeForce;
  Application.Run;
end.
