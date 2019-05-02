unit UtilEnvEnvoiGed;

interface
uses
  Windows, Messages, SysUtils, Classes, HEnt1,
  DBCtrls, HDB, MajTable, IniFiles,
  {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF}
;
type
  TEnvEnvoiGed = class (TObject)
    private
      fServer : string;
      fDatabase : string;
      fdelay    : Integer;
      fModeDebug : Integer;
      fStatus : Boolean;
    public
      property Server : string read fServer;
      property Database : string read fDatabase;
      property delay : integer read fdelay;
      property ModeDebug : Integer read fModeDebug;
      property Status : Boolean read fStatus;
      constructor create (IniFile : string);
      destructor destroy; override;
  end;

implementation

{ TEnvEnvoiGed }

constructor TEnvEnvoiGed.create(IniFile: string);
var SettingFile  : TInifile;
    Section : string;
begin
  fStatus := false;
  Section := 'ENVOIGED';
  SettingFile := TIniFile.Create(IniFile);
  try
    fServer := SettingFile.ReadString(Section, 'SERVER', '');
    fDatabase := SettingFile.ReadString(Section, 'DATABASE', '');
    fDelay := SettingFile.ReadInteger(Section, 'DELAI', 0);
    fModeDebug := SettingFile.ReadInteger(Section, 'DEBUG', 0);
  finally
    if (fDatabase <> '')  and (fServer<>'')  then fStatus := True;
    if fdelay = 0 then fdelay := 60;
    SettingFile.free;
  end;
end;

destructor TEnvEnvoiGed.destroy;
begin

  inherited;
end;

end.
