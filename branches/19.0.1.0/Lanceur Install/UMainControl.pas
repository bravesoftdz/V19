unit UMainControl;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    procedure LanceInstall;
    procedure ShowParams (Drive : string; DriveType : Cardinal;CurrentPath,TempPath,Params : string);
    function LSEVERSION10 : Boolean;
  public
    { Déclarations publiques }
  end;

var
  Form1: TForm1;

implementation
uses UControlInstall,
     registry,
     ShellAPI,
     UInfoParams,
     UCommonParams     
;

{$R *.dfm}

procedure TForm1.FormShow(Sender: TObject);
var XX : TFControlInstall;
begin
  if (not LSEVERSION10) and (not ModeForce) then
  begin
    XX := TFControlInstall.Create(application);
    XX.Show;
    TRY
      XX.Refresh;
      Sleep(10000);
    FINALLY
      XX.free;
    END;
  end else
  begin
    LanceInstall;
  end;
  Application.Terminate;
end;

procedure TForm1.LanceInstall;
var
  Buffer: array[0..1023] of Char;
  TmpExefile,TempPath,TmpXmlfile : string;
  FromExefile,CurrentPath,Params,Drive,FromXmlfile : string;
  DriveType : cardinal;
begin
  GetTempPath(Sizeof(Buffer)-1,Buffer);
  TempPath := format('%s',[buffer]);
  CurrentPath := ExtractFilePath(Application.ExeName); CurrentPath := Copy(CurrentPath,1,Length(CurrentPath)-1);
  Drive := ExtractFileDrive(Application.ExeName); DriveType := GetDriveType(PAnsiChar(Drive));
  TmpExefile := Format('%s\SetupCegid.exe',[buffer]);
  TmpXmlfile := Format('%s\Setup.xml',[buffer]);
  //
  FromExefile := 'System\LSE\SetupCegid.exe';
  FromXmlfile := 'Setup.xml';
  if FileExists(TmpExefile) then DeleteFile(TmpExefile);
  if FileExists(TmpXmlfile) then DeleteFile(TmpXmlfile);
  //
  CopyFile(PAnsiChar(FromExefile),PAnsiChar(TmpExefile),false);
  CopyFile(PAnsiChar(FromXmlfile),PAnsiChar(TmpXmlfile),false);
  //
  if (DriveType = DRIVE_FIXED)  then
  begin
    Params := '/z"'+CurrentPath+'"';
    if ModeDebug then ShowParams (Drive,DriveType,CurrentPath,TempPath,Params);
    ShellExecute(Application.Handle,'open',PAnsiChar(TmpExefile),PAnsiChar(Params),nil,SW_SHOW);
  end else
  begin
    Params := '/z"'+CurrentPath+'"';
    //
    if ModeDebug then ShowParams (DRIVE,DriveType,CurrentPath,TempPath,Params);
    ShellExecute(Application.Handle,'open',PAnsiChar(TmpExefile),PAnsiChar(Params),nil,SW_SHOW);
  end;

end;

function TForm1.LSEVERSION10: Boolean;
var Reg : TRegistry;
    Version,Majeur : string;
    PosPoint : Integer;
begin
  Result := true;
  Version := '';
  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SOFTWARE\Wow6432Node\Cegid\LSE Business\LSE Business Place BTP', false) then
    try
      Version := Reg.ReadString('Version');
    finally
      Reg.CloseKey;
    end;
    if Version = '' then
    begin
      Reg.RootKey := HKEY_LOCAL_MACHINE;
      if Reg.OpenKey('\SOFTWARE\Wow6432Node\Cegid\LSE Business\LSE Business Suite BTP', false) then
      try
        Version := Reg.ReadString('Version');
      finally
        Reg.CloseKey;
      end;
    end;
    if Version <> '' then
    begin
      Majeur := '19';
      PosPoint := Pos('.',Version);
      if PosPoint > 0 then Majeur := Copy(Version,1,PosPoint-1);
      if StrToInt(Majeur) <> 19 then Result := false;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TForm1.ShowParams(Drive : string; DriveType : Cardinal; CurrentPath, TempPath, Params: string);
var   XX: TInfoParams;
begin
   XX := TInfoParams.Create(Application);
   XX.DRIVE.Caption := Drive;
   
   if DriveType = DRIVE_FIXED then XX.DRIVETYPE.Caption := 'DISQUE LOCAL'
   else if DriveType = DRIVE_REMOTE then XX.DRIVETYPE.Caption := 'REPERTOIRE DISTANT'
   else if DriveType = DRIVE_CDROM then XX.DRIVETYPE.Caption := 'CDROM'
   else XX.DRIVETYPE.Caption := 'VOIS PAS !!';
   //
   XX.TEMPPATH.Caption := TempPath;
   XX.CURPATH.Caption := CurrentPath;
   XX.PARAMS.Caption := Params;
   XX.ShowModal;
   XX.free;

end;

end.
