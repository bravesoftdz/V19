unit Ulog;

interface
uses SysUtils
;

procedure ecritLog (const Emplacement : string; const SMessage : string; Nomfic : string=''); overload;
procedure ecritLogS (const LogFile : string; const SMessage : string); overload;
procedure EcritDatas (const DataFile : string; const SMessage : string); 
implementation

procedure ecritLog (const Emplacement : string; const SMessage : string; Nomfic : string='');
var f : TextFile;
    NomFile : string;
		nomLog : string;
begin
  if NomFic = '' then NomFile := 'WebService.log' else NomFile := Nomfic;
  nomLog := IncludeTrailingPathDelimiter(Emplacement)+NomFile;

  if not FileExists(nomLog) then
  begin
    AssignFile(f, nomLog);
    ReWrite(f);
    closeFile(f);
  end;

  AssignFile(f, nomLog);
  Append (f);
  writeln ( f, SMessage);
  Flush(f);
  CloseFile(f);
end;

procedure ecritLogS (const LogFile : string; const SMessage : string); overload;
var f : TextFile;
begin
  if not FileExists(LogFile) then
  begin
    AssignFile(f, LogFile);
    ReWrite(f);
    closeFile(f);
  end;

  AssignFile(f, LogFile);
  Append (f);
  writeln ( f, format('%s : %s',[FormatDateTime('dd/mm/yyyy hh:nn:ss',Now),SMessage]));
  Flush(f);
  CloseFile(f);

end;

procedure EcritDatas (const DataFile : string; const SMessage : string);
var f : TextFile;
begin
  if not FileExists(DataFile) then
  begin
    AssignFile(f, DataFile);
    ReWrite(f);
    closeFile(f);
  end;

  AssignFile(f, DataFile);
  Append (f);
  writeln ( f, format('%s',[SMessage]));
  Flush(f);
  CloseFile(f);
end;

end.















