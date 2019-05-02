unit UtilBU;

interface
uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DB, EdtEtat, EdtDoc,
  EdtREtat, EdtRDoc,
  {$ENDIF}
  SysUtils, Dialogs, UtilPGI, AGLInit,
  Math, EntGC, Classes, HMsgBox,
  {$IFDEF CHR}HRReglement, {$ENDIF}
  {$IFDEF GRC}UtilRT,{$ENDIF}
  UtilGC, ParamSoc;

procedure InitBU;
function findBu (Code : string) : Boolean;
function findTOBBu (Code : string) : TOB;
function GetLibelleBu (Code : string) : string;


implementation
var TOBBU : TOB;

procedure InitBU;
begin
  TOBBU.CLearDetail;
end;

function findBu (Code : string) : Boolean;
var TBB : TOB;
    QQ : TQuery;
begin
  if Code = '' then
  begin
    Result := true;
    Exit;
  end;
  TBB := TOBBU.findFirst(['BBU_CODEBU'],[Code],true);
  if TBB = nil then
  begin
    QQ := OpenSql('SELECT * FROM BTBU WHERE BBU_CODEBU="'+Code+'"',true,-1, '', True);
    if not QQ.eof then
    begin
      TBB := TOB.create ('BTBU',TOBBU,-1);
      TBB.SelectDB ('',QQ);
    end;
    Ferme(QQ);
  end;
  Result := (TBB <> nil);
end;

function findTOBBu (Code : string) : TOB;
var TBB : TOB;
    QQ : TQuery;
begin
  TBB := TOBBU.findFirst(['BBU_CODEBU'],[Code],true);
  if TBB = nil then
  begin
    QQ := OpenSql('SELECT * FROM BTBU WHERE BBU_CODEBU="'+Code+'"',true,-1, '', True);
    if not QQ.eof then
    begin
      TBB := TOB.create ('BTBU',TOBBU,-1);
      TBB.SelectDB ('',QQ);
    end;
    Ferme(QQ);
  end;
  Result := TBB;
end;

function GetLibelleBu (Code : string) : string;
var TT : TOB;
begin
  result := '';
  if Code = '' then
  begin
    exit;
  end;
  TT := FindTOBBu(Code);
  if TT=nil then BEGIN Result := 'Business unit : Inconnu'; EXIT; END;
  result := 'Business unit '+TT.GeTString('BBU_LIBELLE');
end;


INITIALIZATION
  TOBBU := TOB.Create ('LES BUSINESS UNITS',nil,-1);
FINALIZATION
  TOBBU.free;


end.
