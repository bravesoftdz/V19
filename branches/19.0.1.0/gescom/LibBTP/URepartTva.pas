unit URepartTva;

interface
uses HEnt1, UTOB,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DB, EdtEtat, EdtDoc,
  EdtREtat, EdtRDoc, HCtrls,
  uEntCommun,
  SysUtils, Dialogs,
  Math, EntGC, Classes, HMsgBox,
  UtilGC;

type
  RepartTva = class
    class procedure AppliquePiece(TOBPiece,TOBVENTILTVA : TOB);
    class procedure constitueTOB(TOBVENTILTVA : TOB; CpteVte,CpteTVA, CodeTaxe : string; TauxTva,Base,Montant : Double);
    class procedure EcritTOB(TOBpiece,TOBVENTILTVA : TOB);
    class procedure DeleteAncien(TOBPiece : TOB);
  end;

implementation
uses FactTOB,UtilTOBpiece;

{ RepartTva }

class procedure RepartTva.AppliquePiece(TOBPiece, TOBVENTILTVA: TOB);
var II : Integer;
begin
  for II := 0 to TOBVENTILTVA.Detail.count -1 do
  begin
    TOBVENTILTVA.detail[II].SetString('BP4_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG'));
    TOBVENTILTVA.detail[II].SetString('BP4_SOUCHE',TOBPiece.GetString('GP_SOUCHE'));
    TOBVENTILTVA.detail[II].SetInteger('BP4_NUMERO',TOBPiece.GetInteger('GP_NUMERO'));
    TOBVENTILTVA.detail[II].SetInteger('BP4_INDICEG',TOBPiece.GetInteger('GP_INDICEG'));
  end;
end;

class procedure RepartTva.constitueTOB(TOBVENTILTVA: TOB;CpteVte, CpteTVA, CodeTaxe: string; TauxTva, Base, Montant: Double);
var TT : TOB;
begin
  TT := TOBVENTILTVA.FindFirst(['BP4_GENERAL','BP4_CPTTVA'],[CpteVte,CpteTva],true);
  if TT = nil then
  begin
    TT := TOB.Create ('BPIECEREPTVA',TOBVENTILTVA,-1);
    TT.SetString('BP4_GENERAL' , CpteVte);
    TT.SetString('BP4_CPTTVA'  , CpteTva);
    TT.SetString('BP4_CODETAXE', CodeTaxe);
    TT.SetDouble('BP4_TAUXTAXE', TauxTva);
  end;
  TT.SetDouble('BP4_BASEHT'     , TT.GetDouble('BP4_BASEHT')+Base);
  TT.SetDouble('BP4_MONTANTTAXE', TT.GetDouble('BP4_MONTANTTAXE')+Montant);
end;

class procedure RepartTva.DeleteAncien(TOBPiece: TOB);
var cledoc : r_cledoc;
    SQL : String;
begin
  cledoc := TOB2Cledoc (TOBPiece);
  SQL := 'DELETE FROM BPIECEREPTVA WHERE '+WherePiece(cledoc,ttdRepartTva,true);
  ExecuteSQL(SQL);
end;

class procedure RepartTva.EcritTOB(TOBpiece,TOBVENTILTVA: TOB);
begin
  DeleteAncien(TOBPiece);
  AppliquePiece (TOBpiece,TOBVENTILTVA);
  //
  TOBVENTILTVA.SetAllModifie(true);
  TOBVENTILTVA.InsertDB(nil);
end;

end.
