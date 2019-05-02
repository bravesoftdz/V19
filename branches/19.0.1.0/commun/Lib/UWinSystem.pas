unit UWinSystem;

interface
uses
  Windows
  , Classes
  ,Registry
  ;
type
  TWinSystem = class (TObject)
    class function GetAppDataPath : string;
  end;

implementation


class function TWinSystem.GetAppDataPath : string;
var
  Reg : TRegistry;
begin                                                                               
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList', false) then
    try
      result := Reg.readString('ProgramData');
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;


end.


