unit UAvtVerdon;

interface

Uses StdCtrls,
     Controls,
     Classes,
     Fe_Main,
     db, HDB,
     uDbxDataSet,
     uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UentCommun;

Type

  AVCVerdon = class
    class procedure SaisieAvancements (Affaire : string; Cledoc: R_CLEDOC; DateApplic : TDateTime; NumAvc : integer);
    class function ExisteDejaUn (Cledoc: R_CLEDOC) : Boolean;
    class procedure AfficheListeAvancements(Affaire : string; Cledoc: r_cledoc);
    class function ControleSaisie(Cledoc: r_Cledoc;DateSaisie: TDateTime): boolean;
    class function DemandeDateApplic (cledoc: R_CLEDOC;var DateApplic : TDateTime) : Boolean;
    class procedure ChargeTOBFromAvc (cledoc: R_CLEDOC; DateApplic: TDateTime; NumAvc : Integer; TOBAVC,TOBLigFac : TOB);
    class function ChargeFromPrevAvc (cledoc: R_CLEDOC; DateApplic: TDateTime; var NumAvc : Integer; TOBAVC,TOBLigFac : TOB) : boolean;
    class procedure ChargeInitAvc (cledoc: R_CLEDOC; DateApplic: TDateTime; var NumAvc : Integer; TOBAVC,TOBLigFac : TOB);
    class procedure InitLesNull(TOBLigFac : TOB; NumAvc : integer; DateApplic : TDateTime);
    class procedure InitNouvelleSituation (TOBLigFac : TOB; NumAvc : integer; DateApplic : TdateTime);
    class procedure InitTOBAVC(TOBAVC : TOB; Cledoc :R_CLEDOC; NumAvc : Integer; DateAplic : TDateTime);
    class function TraiteQteCumule(Saisie : string; TL : TOB) : boolean;
    class function TraitePourcentAvanc(Saisie : string; TL : TOB) : boolean;
    class procedure ValideLaSaisie(cledoc: R_CLEDOC; Affaire : string; DateApplic: TDateTime; NumAvc : Integer; TOBAVC,TOBLigFac : TOB);
    class procedure Supprime (cledoc: R_CLEDOC; DateApplic: TDateTime);
    class function DernierAvancement(cledoc:r_cledoc;DateApplic :TDateTime) : Boolean;
    class procedure ValideTOBAVC(TOBAVC: TOB; Cledoc: R_CLEDOC; Affaire : string; NumAvc: Integer; DateAplic: TDateTime);
    class function FindNumAvc(Affaire : string; DateMvtFIn : TDateTime): Integer;
  end;


implementation
uses AglInit,UtilTOBPiece,Variants;

{ AVCVerdon }

class function AVCVerdon.ControleSaisie(Cledoc: r_Cledoc;DateSaisie: TDateTime): boolean;
var SQL : string;
begin
  SQL := 'SELECT 1 FROM BAVCVERDON WHERE '+WherePiece(Cledoc,ttdAvcVerdon,false)+' AND BVV_DATEAPPLIC >= "'+USDATETIME(DateSaisie)+'"';
  Result := not ExisteSQL(SQL);
  if not result then
  begin
    PgiInfo('Des avancements sont saisis à une date ultérieure ou à la même date.#13#10 Vous ne pouvez pas saisir à cette date.');
  end;
end;


class procedure AVCVerdon.AfficheListeAvancements(Affaire : string; Cledoc: r_cledoc);
var Param : string;
begin
  Param := 'AFFAIRE='+Affaire+';BVV_NATUREPIECEG='+Cledoc.NaturePiece+';BVV_SOUCHE='+Cledoc.Souche+';BVV_NUMERO='+InttoStr(Cledoc.NumeroPiece)+';BVV_INDICEG='+InttoStr(Cledoc.Indice);
  AGLLanceFiche('BTP','BAVCVERDON_SEL',Param,'','ACTION=MODIFICATION');
end;

class function AVCVerdon.ExisteDejaUn(Cledoc: R_CLEDOC): Boolean;
var SQL : string;
begin
  SQL := 'SELECT 1 FROM BAVCVERDON WHERE '+WherePiece(Cledoc,ttdAvcVerdon,false);
  Result := ExisteSQL(SQL);
end;

class procedure AVCVerdon.SaisieAvancements(Affaire : string; Cledoc: R_CLEDOC;DateApplic: TDateTime; NumAvc : integer);
var TT : TOB;
begin
  TT := TOB.Create ('UN PARAM',nil,-1);
  try
    TT.AddChampSupValeur('AFFAIRE',Affaire);
    TT.AddChampSupValeur('NATUREPIECEG',Cledoc.NaturePiece);
    TT.AddChampSupValeur('SOUCHE',Cledoc.Souche);
    TT.AddChampSupValeur('NUMERO',Cledoc.NumeroPiece);
    TT.AddChampSupValeur('INDICE',Cledoc.Indice);
    TT.AddChampSupValeur('DATEAPPLIC',DateApplic);
    TT.AddChampSupValeur('NUMAVC',NumAvc);
    TheTOB := TT;
    AGLLanceFiche('BTP','BTSAISAVCVERDON','','','ACTION=MODIFICATION');
  finally
    TT.free;
  end;
end;

class function AVCVerdon.DemandeDateApplic(cledoc: R_CLEDOC; var DateApplic: TDateTime): Boolean;
var TOBDATES : TOB;
begin
  result := false;
  TOBDATES := TOB.Create ('UNE TOB',nil,-1);
  TRY
    TOBDATES.AddChampSupValeur('RETOUROK','-');
    TOBDates.AddChampSupValeur('CTRLEX','X');
    TOBDates.AddChampSupValeur('TYPEDATE','Date de saisie');
    TOBDates.AddChampSupValeur('DATFAC',iDate1900);
    TRY
      TheTOB := TOBDates;
      AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
      TheTOB := nil;
    FINALLY
      if TOBDATES.getValue('RETOUROK')='X' then
      begin
        DateApplic := TOBDates.GetDateTime('DATFAC');
      end;
    END;
    if TOBDATES.getValue('RETOUROK')<>'X' then Exit;
    result := AVCVerdon.ControleSaisie(Cledoc,DateApplic) ;
  FINALLY
    TOBDATES.Free;
  END;
end;


class procedure AVCVerdon.ChargeTOBFromAvc(cledoc: R_CLEDOC;DateApplic: TDateTime; NumAvc: Integer; TOBAVC, TOBLigFac: TOB);
var SQL : string;
    QQ : TQuery;
begin
  SQL := 'SELECT * FROM BAVCVERDON WHERE '+WherePiece(cledoc,ttdAvcVerdon,true)+ ' AND BVV_NUMAVC='+IntToStr(NumAvc);
  QQ := OpenSQL(SQl,True,1,'',true);
  if not QQ.eof then
  begin
    TOBAVC.SelectDB('',QQ);
  end else
  begin
    InitTOBAVC(TOBAVC,cledoc,NumAvc,DateApplic);
  end;
  Ferme(QQ);
  //
  SQL := 'SELECT LIGNE.GL_TYPELIGNE,LIGNE.GL_TYPEARTICLE,LIGNE.GL_CODEARTICLE,LIGNE.GL_LIBELLE,LIGNE.GL_NUMORDRE,LIGNE.GL_NUMLIGNE,LIGNE.GL_MONTANTHTDEV,LIGNE.GL_QTEFACT,LIGNE.GL_NIVEAUIMBRIC,LIGNE.GL_INDICENOMEN,'+
         'LIGNEFAC.* '+
         'FROM LIGNE '+
         'LEFT JOIN LIGNEFAC ON GL_NATUREPIECEG=BLF_NATUREPIECEG AND GL_SOUCHE=BLF_SOUCHE AND GL_NUMERO=BLF_NUMERO AND GL_INDICEG=BLF_INDICEG AND GL_NUMORDRE=BLF_NUMORDRE AND BLF_NUMAVC='+IntToStr(NumAvc)+' '+
         'WHERE '+WherePiece(cledoc,ttdLigne,false)+
         'ORDER BY GL_NUMLIGNE';
  QQ := OpenSQL(SQl,True,-1,'',true);
  if not QQ.eof then
  begin
    TOBLigFac.LoadDetailDb('LIGNEFAC','','',QQ,false);
  end;
  Ferme(QQ);
  //
  InitLesNull(TOBLigFac,NumAvc,DateApplic);
end;

class function AVCVerdon.ChargeFromPrevAvc(cledoc: R_CLEDOC;DateApplic: TDateTime; var NumAvc: Integer; TOBAVC,TOBLigFac: TOB): boolean;
var SQL : String;
    QQ : TQuery;
begin
  Result := false;
  SQL := 'SELECT 1 FROM BAVCVERDON WHERE '+WherePiece(cledoc,ttdAvcVerdon,true);
  if not ExisteSQL(SQL) then Exit;
  Result := True;
  //
  SQL := 'SELECT MAX(BVV_NUMAVC) FROM BAVCVERDON WHERE '+WherePiece(cledoc,ttdAvcVerdon,true);
  QQ := OpenSQL(SQL,True,1,'',true);
  if not QQ.eof then
  begin
    NumAvc := QQ.fields[0].AsInteger;
  end;
  Ferme(QQ);
  ChargeTOBFromAvc ( cledoc,DateApplic, NumAvc, TOBAVC,TOBLigFac);
  // Situation Suivante (avancement)
  NumAvc := NumAvc + 1;
  InitNouvelleSituation (TOBLigFac,NumAvc,DateApplic);
end;

class procedure AVCVerdon.ChargeInitAvc(cledoc: R_CLEDOC;DateApplic: TDateTime; var NumAvc: Integer; TOBAVC, TOBLigFac: TOB);
begin
  NumAvc := 0;
  ChargeTOBFromAvc ( cledoc,DateApplic, NumAvc, TOBAVC,TOBLigFac);
  // Situation Suivante (avancement)
  NumAvc := NumAvc + 1;
  InitNouvelleSituation (TOBLigFac,NumAvc,DateApplic);
end;

class procedure AVCVerdon.InitLesNull(TOBLigFac: TOB; NumAvc : integer; DateApplic : TDateTime);
var II : Integer;
    TL : TOB;
begin
  for II := 0 to TOBLigFac.detail.count -1 do
  begin
    TL := TOBLIgFac.detail[II];
    if TL.GetInteger('BLF_NUMAVC') = 0 then
    begin
      TL.SetDouble('BLF_MTMARCHE',TL.GetDouble('GL_MONTANTHTDEV'));
      TL.SetDouble('BLF_MTCUMULEFACT',0);
      TL.SetDouble('BLF_MTDEJAFACT',0);
      TL.SetDouble('BLF_MTSITUATION',0);
      TL.SetDouble('BLF_QTEMARCHE',TL.GetDouble('GL_QTEFACT'));
      TL.SetDouble('BLF_QTECUMULEFACT',0);
      TL.SetDouble('BLF_QTEDEJAFACT',0);
      TL.SetDouble('BLF_QTESITUATION',0);
      TL.SetDouble('BLF_POURCENTAVANC',0);
      TL.SetString('BLF_NATURETRAVAIL','');
      TL.SetString('BLF_FOURNISSEUR','');
      TL.SetInteger('BLF_NUMORDRE',TL.GetInteger('GL_NUMORDRE'));
      TL.SetInteger('BLF_UNIQUEBLO',0);
      TL.SetDouble('BLF_TOTALTTCDEV',0);
      TL.SetDouble('BLF_MTPRODUCTION',0);
      TL.SetString('BLF_CODEMARCHE','');
      TL.SetDateTime('BLF_DATEAPPLIC',DateApplic);
      TL.SetInteger('BLF_NUMAVC',NumAvc);
    end;
  end;
end;

class procedure AVCVerdon.InitNouvelleSituation(TOBLigFac: TOB; NumAvc : integer; DateApplic : TdateTime);
var II : Integer;
    TL : TOB;
begin
  for II := 0 to TOBLigFac.detail.count -1 do
  begin
    TL := TOBLIgFac.detail[II];
    if TL.GetSTring('GL_TYPELIGNE')<>'ART' then continue;
    TL.SetDouble('BLF_MTDEJAFACT',TL.GetDouble('BLF_MTDEJAFACT')+TL.GetDouble('BLF_MTSITUATION'));
    TL.SetDouble('BLF_MTCUMULEFACT',TL.GetDouble('BLF_MTDEJAFACT'));
    TL.SetDouble('BLF_MTSITUATION',0);
    TL.SetDouble('BLF_QTEDEJAFACT',TL.GetDouble('BLF_QTEDEJAFACT')+TL.GetDouble('BLF_QTESITUATION'));
    TL.SetDouble('BLF_QTECUMULEFACT',TL.GetDouble('BLF_QTEDEJAFACT'));
    TL.SetDouble('BLF_QTESITUATION',0);
    if TL.GetDouble('BLF_QTEMARCHE') <> 0 then TL.SetDouble('BLF_POURCENTAVANC',ARRONDI(TL.GetDouble('BLF_QTECUMULEFACT')/TL.GetDouble('BLF_QTEMARCHE')*100,2));
    TL.SetInteger('BLF_NUMAVC',NumAvc);
    TL.SetDateTime('BLF_DATEAPPLIC',DateApplic);
  end;
end;

class procedure AVCVerdon.InitTOBAVC(TOBAVC: TOB; Cledoc: R_CLEDOC;NumAvc: Integer; DateAplic: TDateTime);
begin
  TOBAVC.SetString('BVV_NATUREPIECEG',Cledoc.NaturePiece);
  TOBAVC.SetString('BVV_SOUCHE',Cledoc.Souche);
  TOBAVC.SetInteger('BVV_NUMERO',Cledoc.NumeroPiece );
  TOBAVC.SetInteger('BVV_INDICEG',Cledoc.Indice );
  TOBAVC.SetDateTime('BVV_DATEAPPLIC',DateAplic );
  TOBAVC.SetInteger('BVV_NUMAVC',NumAvc );
  TOBAVC.Setboolean('BVV_VALIDE',false );
end;

class procedure AVCVerdon.ValideTOBAVC(TOBAVC: TOB; Cledoc: R_CLEDOC; Affaire : string; NumAvc: Integer; DateAplic: TDateTime);
begin
  TOBAVC.SetString('BVV_AFFAIRE',Affaire);
  TOBAVC.SetString('BVV_NATUREPIECEG',Cledoc.NaturePiece);
  TOBAVC.SetString('BVV_SOUCHE',Cledoc.Souche);
  TOBAVC.SetInteger('BVV_NUMERO',Cledoc.NumeroPiece );
  TOBAVC.SetInteger('BVV_INDICEG',Cledoc.Indice );
  TOBAVC.SetDateTime('BVV_DATEAPPLIC',DateAplic );
  TOBAVC.SetInteger('BVV_NUMAVC',NumAvc );
end;

class function AVCVerdon.TraiteQteCumule(Saisie: string;TL: TOB): boolean;
begin
  Result := false;
  // si l'on saisie la quantité cumulé la valeur du mois = cumule - déja saisie
  if VALEUR(Saisie) > TL.GetDouble('BLF_QTEMARCHE') then
  begin
    if PGIask('Vous avez saisie une quantité supérieure à la quantité prévue.#13#10 Confirmez-vous ?') <> mrYes then
    begin
      Result := true;
      Exit;
    end;
  end;
  if VALEUR(Saisie) < TL.GetDouble('BLF_QTECUMULEFACT') then
  begin
    if PGIask('Vous avez saisie une quantité inférieure à la quantité précédente.#13#10 Confirmez-vous ?') <> mrYes then
    begin
      Result := true;
      Exit;
    end;
  end;
  TL.SetDouble('BLF_QTECUMULEFACT',VALEUR(Saisie));
  TL.SetDouble('BLF_QTESITUATION',TL.GetDouble('BLF_QTECUMULEFACT')-TL.GetDouble('BLF_QTEDEJAFACT'));
  if TL.GetDouble('BLF_QTEMARCHE') <> 0 then TL.SetDouble('BLF_POURCENTAVANC',ARRONDI((TL.GetDouble('BLF_QTECUMULEFACT')/TL.GetDouble('BLF_QTEMARCHE'))*100,2));
  TL.SetDouble('BLF_MTCUMULEFACT',ARRONDI(TL.GetDouble('BLF_MTMARCHE')*TL.GetDouble('BLF_POURCENTAVANC')/100,2));
  TL.SetDouble('BLF_MTSITUATION',TL.GetDouble('BLF_MTCUMULEFACT')-TL.GetDouble('BLF_MTDEJAFACT'));
end;

class function AVCVerdon.TraitePourcentAvanc(Saisie: string;TL: TOB): boolean;
begin
  Result := false;
  if VALEUR(Saisie) > 100 then
  begin
    if PGIask('Vous avez saisie un avancement supérieur à 100%#13#10 Confirmez-vous ?') <> mrYes then
    begin
      Result := true;
      Exit;
    end;
  end;
  if VALEUR(Saisie) > 100 then
  begin
    if PGIask('Vous avez saisie un avancement inférieur au précédent#13#10 Confirmez-vous ?') <> mrYes then
    begin
      Result := true;
      Exit;
    end;
  end;
  TL.SetDouble('BLF_POURCENTAVANC',VALEUR(Saisie));
  TL.SetDouble('BLF_QTECUMULEFACT',ARRONDI((VALEUR(Saisie)/100)*TL.GetDouble('BLF_QTEMARCHE'),V_PGI.OkDecQ));
  TL.SetDouble('BLF_QTESITUATION',TL.GetDouble('BLF_QTECUMULEFACT')-TL.GetDouble('BLF_QTEDEJAFACT'));
  if TL.GetDouble('BLF_QTEMARCHE') <> 0 then TL.SetDouble('BLF_POURCENTAVANC',ARRONDI((TL.GetDouble('BLF_QTECUMULEFACT')/TL.GetDouble('BLF_QTEMARCHE'))*100,2));
  TL.SetDouble('BLF_MTCUMULEFACT',ARRONDI(TL.GetDouble('BLF_MTMARCHE')*TL.GetDouble('BLF_POURCENTAVANC')/100,2));
  TL.SetDouble('BLF_MTSITUATION',TL.GetDouble('BLF_MTCUMULEFACT')-TL.GetDouble('BLF_MTDEJAFACT'));
end;

class procedure AVCVerdon.ValideLaSaisie(cledoc: R_CLEDOC;Affaire : string; DateApplic: TDateTime; NumAvc: Integer; TOBAVC, TOBLigFac: TOB);
var II : Integer;
    TL : TOB;
begin
  BEGINTRANS;
  //
  try
    Supprime (cledoc,DateApplic);
    //
    TOBAVC.SetDouble('BVV_MTMARCHE',0);
    TOBAVC.SetDouble('BVV_MTCUMULE',0);
    TOBAVC.SetDouble('BVV_MTMOIS',0);
    for II :=0 to  TOBLigFac.detail.count -1 do
    begin
      TL := TOBLigFac.detail[II];
      TL.SetString('BLF_NATUREPIECEG',cledoc.NaturePiece);
      TL.SetString('BLF_SOUCHE',cledoc.Souche);
      TL.SetInteger('BLF_NUMERO',cledoc.NumeroPiece);
      TL.SetInteger('BLF_INDICEG',cledoc.Indice);
      TL.SetInteger('BLF_NUMAVC',NumAvc);
      TL.SetDateTime('BLF_DATEAPPLIC',DateApplic);
      TL.SetInteger('BLF_NUMAVC',NumAvc);
      if TL.GetString('GL_TYPELIGNE')='ART' then
      begin
        TOBAVC.SetDouble('BVV_MTMARCHE',TOBAVC.GetDouble('BVV_MTMARCHE')+TL.GetDouble('BLF_MTMARCHE'));
        TOBAVC.SetDouble('BVV_MTCUMULE',TOBAVC.GetDouble('BVV_MTCUMULE')+TL.GetDouble('BLF_MTCUMULEFACT'));
        TOBAVC.SetDouble('BVV_MTMOIS',TOBAVC.GetDouble('BVV_MTMOIS')+TL.GetDouble('BLF_MTSITUATION'));
      end;
    end;
    if TOBAVC.GetDouble('BVV_MTMARCHE') <> 0 then
    begin
      TOBAVC.SetDouble('BVV_POURCENTCUM',ARRONDI((TOBAVC.GetDouble('BVV_MTCUMULE')/TOBAVC.GetDouble('BVV_MTMARCHE'))*100,2));
      TOBAVC.SetDouble('BVV_POURCENTMOIS',ARRONDI((TOBAVC.GetDouble('BVV_MTMOIS')/TOBAVC.GetDouble('BVV_MTMARCHE'))*100,2));
    end;
    TOBLigFac.SetAllModifie(true);
    TOBLigFac.InsertDB(nil);
    //
    ValideTOBAVC (TOBAVC,cledoc,Affaire,NumAvc,DateApplic);
    TOBAVC.SetAllModifie(True);
    TOBAVC.InsertDB(nil);
    //
    COMMITTRANS;
  except
    on E:Exception do
    begin
      PgiInfo(E.message);
      ROLLBACK;
    end;
  end;

end;

class procedure AVCVerdon.Supprime(cledoc: R_CLEDOC; DateApplic: TDateTime);
var SQL : STring;
begin
  SQL := 'DELETE FROM BAVCVERDON WHERE '+WherePiece(Cledoc,ttdAvcVerdon,false)+' AND BVV_DATEAPPLIC = "'+USDATETIME(DateApplic)+'"';
  ExecuteSQL(SQL);
  SQL := 'DELETE FROM LIGNEFAC WHERE '+WherePiece(Cledoc,ttdLignefac,false)+' AND BLF_DATEAPPLIC = "'+USDATETIME(DateApplic)+'"';
  ExecuteSQL(SQL);
end;

class function AVCVerdon.DernierAvancement(cledoc: r_cledoc;DateApplic: TDateTime): Boolean;
var SQL : string;
begin
  SQL := 'SELECT 1 FROM BAVCVERDON WHERE '+WherePiece(Cledoc,ttdAvcVerdon,false)+' AND BVV_DATEAPPLIC > "'+USDATETIME(DateApplic)+'"';
  Result := not ExisteSQL(SQL);
end;

class function AVCVerdon.FindNumAvc(Affaire: string;DateMvtFIn: TDateTime): Integer;
var QQ : TQuery;
    SQL : String;
begin
  Result := 0;
  SQL := 'SELECT MAX(BVV_NUMAVC) FROM BAVCVERDON WHERE BVV_AFFAIRE="'+Affaire+'" AND BVV_DATEAPPLIC <= "'+USDATETIME(DateMvtFin)+'"';
  QQ := OpenSql(SQL,True,1,'',false);
  if not QQ.Eof then
  begin
    Result := QQ.Fields [0].AsInteger;
  end;
  Ferme(QQ);
end;

end.
