unit BTRECUPHISTOACHAT;

interface

function BTLanceFiche_POCRecupHistoAchat(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

uses
  {$IFNDEF EAGLCLIENT}
  db
  ,mul
  ,FE_Main
   {$IFNDEF DBXPRESS}
  , dbTables
   {$ELSE}
  , uDbxDataSet {$ENDIF}
  , dbGrids
  , EdtREtat
  {$ELSE}
  , emul
  , MaineAGL
  , UtileAgl
  {$ENDIF}
  ;

function BTLanceFiche_POCRecupHistoAchat(Nat, Cod : String ; Range,Lequel,Argument : string) : string;
begin
  if (Nat = '') and (Cod = '') then
    Result := ''
  else
    Result := AGLLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

end.
