unit UGenereDocument;
interface
uses
  Classes, SysUtils,
  uTob, hEnt1, hCtrls,
  EntGC, SaisUtil,
  wCommuns
  ,uEntCommun
  ,UtilConso,
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  FactSpec, AffaireUtil, BTPUtil,UtilPhases,CommonTools
  ;

type

  Tresult = class (TObject)
  private
    procedure init;
  public
    RefPiece : string;
    NatureDoc : string;
    NumeroDoc : Integer;
    ErrorResult : TIOErr;
    LibError : string;
    constructor create;
  end;

  TGenerePiece = class (TObject)
  private
    DEV : RDevise;
    TOBArticlePlusPlus : TOB;
    TOBPiecePrec : TOB;
    TOBPiece:TOB;
    TOBVTECOLLECTIF:TOB;
    TOBAffaireInterv:TOB;
    TOBPieceTrait:TOB;
    TOBBasesL:TOB;
    TOBBases:TOB;
    TOBEches:TOB;
    TOBPorcs:TOB;
    TOBTiers:TOB;
    TOBArticles:TOB;
    TOBConds:TOB;
    TOBDispo:TOB;
    TOBAdresses:TOB;
    TOBAcomptes:TOB;
    TOBAffaire:TOB;
    TOBCPTA:TOB;
    TOBANAP:TOB;
    TOBANAS:TOB;
    TOBOuvrage:TOB;
    TOBOuvragesP:TOB;
    TOBLIENOLE:TOB;
    TOBPieceRG:TOB;
    TOBBasesRG:TOB;
    TOBSSTRAIT:TOB;
    TOBEnteteScan : TOB;
    TOBREPARTTVA : TOB;
    ZTOBREPARTTVA : TOB;
    TOBVENTILTVA : TOB;
    TOBProv : TOB;   // liste des BLS (g�n�ration de facture)
    //
    GereReliquat : boolean;
    OldEcr, OldStk: RMVT;
    //
    fResult : TResult;
    VenteAchat : string;
    GestionConso : TGestionPhase;
    DocFromBSV : Boolean;
    fBAP : boolean;
    procedure creeTOB;
    procedure LibereTOB;
    procedure Reinit;
    procedure ConstitueEnvironnement (TOBProv : TOB);
    procedure RecalculeLeDocument;
    procedure RenumerotePiece;
    procedure UpdateDocumentsPrecedent;
    procedure ReinitMontantPieceTrait;
    procedure InitToutModif;
    procedure AddReferenceDoc(CleDoc: R_CLEDOC; TOBPP: TOB);
    procedure ConstitueTobArticles ;
    procedure ConstitueTobDispo;
    procedure ValideLaNumerotation;
    procedure ValideLeDocument;
    procedure GenereCompta;
    function GetNumeroPiece : integer;
    procedure ConstitueTobAffaire;
    procedure AjouteArticlesPLus;
    procedure ConstituePieceFromBast(TOBBAST : TOB; DateFac : TDateTime);
    procedure ChargeTiers(NatureAuxi, Tiers: string);
    function ChargeTOBA(RefUnique, stDepot: string): TOB;
    procedure ValideLaRepartTVA(TOBPiece, TOBREPARTTVA: TOB);
    procedure InverseAvoir;
    function IsAvoir(TOBpiece: TOB): Boolean;

  public
    constructor create; overload;
    destructor destroy; override;
    function GenereDocument (TOBProv : TOB) : Tresult;
    function GenereDocFromBAST (TOBBAST : TOB; DateFac : TdateTime) : Tresult;
    property result : Tresult read fResult;
    property TOBArticlePlus : TOB read TOBArticlePlusPlus write TOBArticlePlusPlus;
    property TOBPROVFAC : TOB read TOBProv write TOBProv;
    property GTOBREPARTVA : TOB read ZTOBREPARTTVA write ZTOBREPARTTVA;
    property BAP : boolean read fBAP write fBAP;
  end;


implementation
uses FactTOB,ParamSoc,UtilPGI,FactAdresse,FactCalc,FactUtil,FactComm,UtilTOBpiece,StockUtil,FactPiece,FactOuvrage,FactCpta,BTGENODANAL_TOF,
  TypInfo,ENt1,FactRG,UCumulCollectifs,FactArticle,USpecifPoc,UrepartTva;
{ TGenerePiece }


procedure TGenerePiece.AddReferenceDoc(CleDoc : R_CLEDOC;TOBPP : TOB);
var TOBL : TOB;
begin
  TOBL := TOBPP.FindFirst(['NATUREPIECEG','SOUCHE','NUMERO','INDICEG'],
                          [CleDoc.NaturePiece,CleDoc.Souche,CleDoc.NumeroPiece,CleDoc.Indice],true);
  if TOBL = nil then
  begin
    TOBL:= TOB.Create ('UN DOC',TOBPP,-1);
    TOBL.AddChampSupValeur('NATUREPIECEG',CleDoc.NaturePiece);
    TOBL.AddChampSupValeur('SOUCHE',CleDoc.Souche);
    TOBL.AddChampSupValeur('NUMERO',CleDoc.NumeroPiece);
    TOBL.AddChampSupValeur('INDICEG',CleDoc.Indice);
  end;
end;

procedure TGenerePiece.ConstitueTobAffaire;
var QQ: TQuery;
begin
  if TOBpiece.getString('GP_AFFAIRE') = '' then exit;
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'Lecture du chantier';

  QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBpiece.getString('GP_AFFAIRE')+'"',true,-1, '', True);
  TRY
    if not QQ.eof then TOBAffaire.selectdb ('',QQ) else Exit;
  FINALLY
    ferme (QQ);
  end;

  fResult.ErrorResult := oeOk;
  fResult.LibError := '';
end;

procedure TGenerePiece.ChargeTiers(NatureAuxi,Tiers : string);
var QQ : TQuery;
    sRib : string;
begin
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'Tiers inexistant';
  QQ := OpenSQL('SELECT * FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS WHERE T_TIERS="' + Tiers + '" AND T_NATUREAUXI="'+NatureAuxi+'"', True,-1, '', True);
  if not QQ.EOF then
  begin
    TOBTiers.SelectDB('', QQ);
    fResult.ErrorResult := oeOk;
  end;
  ferme (QQ);
  if fResult.ErrorResult <> OeOk then exit;
  //
  sRib := '';
  QQ := OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + TOBTiers.GetValue('T_AUXILIAIRE') + '" AND R_PRINCIPAL="X"', True,-1, '', True);
  if not QQ.EOF then
  begin
    sRib := EncodeRIB(QQ.FindField('R_ETABBQ').AsString, QQ.FindField('R_GUICHET').AsString,
                      QQ.FindField('R_NUMEROCOMPTE').AsString, QQ.FindField('R_CLERIB').AsString,
                      QQ.FindField('R_DOMICILIATION').AsString);
  end;

  Ferme(QQ);
  TOBTiers.SetString('RIB', sRIB);
  //
  if (TOBTiers.GetValue('T_FERME') = 'X') then
  begin
    fResult.ErrorResult := oeUnknown;
    fResult.LibError := 'Tiers Ferm�';
  end;
end;

procedure TGenerePiece.ConstitueEnvironnement(TOBProv : TOB);
var SavType,NatureAuxi : string;
    cledoc : r_cledoc;
    TOBL  :TOB;
    II : Integer;
    NumConsoPrec : double;
begin
  NumConsoPrec := 0;
  if TOBProv <> nil then
  begin
    TOBPiece.Dupliquer(TOBProv,True,true);
    TOBEnteteScan := TOB(TOBProv.data);
    // modif
    TOBPiece.SetString('GP_CREEPAR','BSV');
  end;
  VenteAchat := GetInfoParPiece(TOBPiece.GetString('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  if Venteachat = 'ACH' then NatureAuxi := 'FOU' else NatureAuxi := 'CLI';
  //
  GereReliquat := GetInfoParPiece(TOBPiece.GetString('GP_NATUREPIECEG'), 'GPP_RELIQUAT')='X';
  //
  for II := 0 to TOBPiece.Detail.count -1 do
  begin
    TOBL := TOBPiece.detail[II];
    if TOBL.FieldExists('NUMCONSOPREC') then NumConsoPrec := TOBL.GetDouble('NUMCONSOPREC');
    NewTOBLigneFille (TOBL);
    if TOBProv <> nil then
    begin
      AddLesSupLigne(TOBL, false);
    end;
    Savtype := TOBL.GetValue('BNP_TYPERESSOURCE');
    if TOBProv <> nil then InitLesSupLigne(TOBL,false);
    gestionConso.InitialisationLigne (TOBL);
    TOBL.PutValue('BNP_TYPERESSOURCE',Savtype);
    if (TOBL.getValue('GL_TYPELIGNE')='COM') and (TOBL.getValue('GL_TYPEARTICLE')='RRR') then
    begin
      TOBL.putValue('MODIFIABLE','-');
      TOBPiece.PutValue('ESCREMMULTIPLE','X');
    end;
    if TOBL.GetString('GL_PIECEPRECEDENTE')<>'' then
    begin
      DecodeRefPiece(TOBL.GetString('GL_PIECEPRECEDENTE'),cledoc);
    end;
    if NumConsoPrec <> 0 then
    begin
      if not TOBL.FieldExists ('NUMCONSOPREC') then TOBL.AddChampSup ('NUMCONSOPREC',false);
      TOBL.SetDouble('NUMCONSOPREC',NumConsoPrec);
    end;
  end;
  ConstitueTobAffaire;
  if fResult.ErrorResult <> OeOk then Exit;
  //
  ConstitueTobArticles;
  if fResult.ErrorResult <> OeOk then Exit;
  AjouteArticlesPLus;
  if fResult.ErrorResult <> OeOk then Exit;
  for II := 0 to TOBPiece.Detail.count -1 do
  begin
    TOBL := TOBPiece.detail[II];
    if TOBL.GetSTring('GL_TYPELIGNE')<>'ART' Then Continue;
    ChargeTOBA (TOBL.GEtSTring('GL_ARTICLE'),TOBL.GetSTring('GL_DEPOT'));
  end;
  if fResult.ErrorResult <> OeOk then Exit;
  ConstitueTobDispo;
  if fResult.ErrorResult <> OeOk then Exit;
  //
  ValideLaPeriode (TOBPiece);
  //
  ValideAnalytiques (TOBpiece,TOBANAP,TOBANAS);
  //
  if TOBProv <> nil then
  begin
    ChargeTiers(NatureAuxi,TOBProv.GetString('GP_TIERS'));
  end;
  //
  if fResult.ErrorResult <> OeOk then exit;
  if GetParamSoc('SO_GCPIECEADRESSE') then
  BEGIN
    TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Livraison}
    TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Facturation}
  END else
  BEGIN
    TOB.Create('ADRESSES',TOBAdresses,-1) ; {Livraison}
    TOB.Create('ADRESSES',TOBAdresses,-1) ; {Facturation}
  END ;
  //
  TiersVersAdresses(TOBTiers, TOBAdresses, TOBPiece);
  AffaireVersAdresses(TOBAffaire,TOBAdresses,TOBPiece);
  TOBAdresses.SetAllModifie(True);
  //
  DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, CleDoc.DatePiece);
  //
  if ZTOBREPARTTVA <> nil then
  begin
    II := 0;
    repeat
      ZTOBREPARTTVA.detail[II].ChangeParent(TOBREPARTTVA,-1);
    until ZTOBREPARTTVA.detail.count = 0;
  end;
end;

constructor TGenerePiece.create;
begin
  fBAP := false;
  GestionConso := TGestionPhase.create;
  fResult := Tresult.Create;
  FillChar(OldEcr,Sizeof(OldEcr),#0) ; FillChar(OldStk,Sizeof(OldStk),#0) ;
  creeTOB;
end;

procedure TGenerePiece.creeTOB;
begin
  TOBDispo:= TOB.Create('LES DISPO',nil,-1);
  TOBPiecePrec := TOB.Create ('LES PIECES PREC',nil,-1);
  TOBPiece := TOB.Create('PIECE', nil, -1);
  AddLesSupEntete(TOBPiece);
  TOBVTECOLLECTIF := TOB.Create ('LES VENTIL COLL',nil,-1);
  TOBAffaireInterv := TOB.Create ('LES CO-SOUSTRAITANTS',nil,-1);
  TOBPieceTrait := TOB.Create ('LES LIGNES EXTRENALISE',nil,-1);
  TOBBasesL := TOB.Create('LES BASES LIGNES', nil, -1);
  TOBBases := TOB.Create('BASES', nil, -1);
  TOBEches := TOB.Create('Les ECHEANCES', nil, -1);
  TOBPorcs := TOB.Create('PORCS', nil, -1);
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBTiers.AddChampSup('RIB', False);
  TOBArticles := TOB.Create('ARTICLES', nil, -1);
  TOBConds := TOB.Create('CONDS', nil, -1);
  TOBAdresses := TOB.Create('LESADRESSES', nil, -1);
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
  end else
  begin
    TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
    TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
  end;
  TOBAcomptes := TOB.Create('', nil, -1);
  TOBAffaire := TOB.Create('AFFAIRE', nil, -1);
  TOBCpta:=TOB.Create('',Nil,-1) ;
  TOBCpta.AddChampSup('LASTSQL',False) ; TOBCpta.PutValue('LASTSQL','') ;
  TOBANAP := TOB.Create('', nil, -1);
  TOBANAS := TOB.Create('', nil, -1);
  TOBOuvrage := TOB.Create('OUVRAGES', nil, -1);
  TOBOuvragesP := TOB.Create('LES OUVRAGES PLAT', nil, -1);
  TOBLIENOLE := TOB.Create('LIENS', nil, -1);
  TOBPieceRG := TOB.create('PIECERRET', nil, -1);
  TOBBasesRG := TOB.create('BASESRG', nil, -1);
  TOBSSTRAIT := TOB.Create ('LES SOUS TRAITS',nil,-1);
  TOBREPARTTVA := TOB.Create ('LES REPARTTVA',nil,-1);
  TOBVENTILTVA := TOB.Create('LES VENTILS TVA',nil,-1);
end;

destructor TGenerePiece.destroy;
begin
  GestionConso.free;
  TOBREPARTTVA.free;
  TOBVENTILTVA.free;
  LibereTOB;
  inherited;
end;

function TGenerePiece.GenereDocFromBAST (TOBBAST : TOB; DateFac : TdateTime) : Tresult;
begin
  VH_GC.ModeGestionEcartComptable := '';
  DocFromBSV := false;
  Result := fResult;
  fBAP := True;
  //
  Reinit;
  TRY
    ChargeTiers('FOU',TOBBAST.GetString('BM4_FOURNISSEUR'));
    if fResult.ErrorResult <> OeOk then Exit;
    //
    ConstituePieceFromBast(TOBBAST,DateFac);
    if fResult.ErrorResult <> OeOk then Exit;
    ConstitueEnvironnement (nil);
    if fResult.ErrorResult = OeOk then
    begin
      RenumerotePiece ;
      RecalculeLeDocument;
      InitToutModif;
      if fResult.ErrorResult = OeOk then
      begin
        BEGINTRANS;
        TRY
          ValideLeDocument;
          COMMITTRANS;
        EXCEPT
          ROLLBACK;
        end;
      end;
    end;
  FINALLY
    Result := fResult;
    VH_GC.ModeGestionEcartComptable := '';
  END;
end;

function TGenerePiece.GenereDocument(TOBPRov: TOB): Tresult;
begin
  VH_GC.ModeGestionEcartComptable := '';
  //
  DocFromBSV := True;
  Reinit;
  ConstitueEnvironnement (TOBPRov);
  TRY
    fBAP := TOBPiece.GetBoolean('GP_BAP'); // gestion des bons � payer
    if fResult.ErrorResult = OeOk then
    begin
      RenumerotePiece ;
      RecalculeLeDocument;
      InitToutModif;
      if fResult.ErrorResult = OeOk then
      begin
        BEGINTRANS;
        TRY
          UpdateDocumentsPrecedent;
          ValideLeDocument;
          COMMITTRANS;
        EXCEPT
          ROLLBACK;
        end;
      end;
    end;
  FINALLY
    Result := fResult;
    VH_GC.ModeGestionEcartComptable := '';
  END;
end;

procedure TGenerePiece.InitToutModif ;
Var NowFutur : TDateTime ;
BEGIN
  NowFutur:=NowH ;
  TOBPiece.SetAllModifie(True) ; TOBPiece.SetDateModif(NowFutur) ;
  TOBBases.SetAllModifie(True)  ;
  TOBBasesL.SetAllModifie(True)  ;
  TOBOuvragesP.SetAllModifie(True)  ;
  TOBEches.SetAllModifie(True)  ;
  TOBAcomptes.SetAllModifie(True)  ;
  TOBPorcs.SetAllModifie(True)  ;
  TOBTiers.SetAllModifie(True)  ;
  TOBAnaP.SetAllModifie(True)   ; TOBAnaS.SetAllModifie(True)   ;
  TOBPieceRG.SetAllModifie (true);
  TOBBasesRG.SetAllModifie (true);
  TOBVTECOLLECTIF.SetAllModifie(true);
  TOBREPARTTVA.SetAllModifie(true);
END ;

procedure TGenerePiece.LibereTOB;
begin
  TOBDispo.free;
  TOBPiecePrec.free;
  TOBPiece.free;
  TOBVTECOLLECTIF.free;
  TOBAffaireInterv.free;
  TOBPieceTrait.free;
  TOBBasesL.free;
  TOBBases.free;
  TOBEches.free;
  TOBPorcs.free;
  TOBTiers.free;
  TOBArticles.free;
  TOBConds.free;
  TOBAdresses.free;
  TOBAcomptes.free;
  TOBAffaire.free;
  TOBCPTA.free;
  TOBANAP.free;
  TOBANAS.free ;
  TOBOuvrage.free;
  TOBOuvragesP.free;
  TOBLIENOLE.free;
  TOBPieceRG.free;
  TOBBasesRG.free;
  TOBSSTRAIT.free;
end;


procedure TGenerePiece.ReinitMontantPieceTrait;
var TOBP : TOB;
	Indice : integer;
begin
	For Indice := 0 to TOBpieceTrait.detail.count -1 do
  begin
  	TOBP := TOBpieceTrait.detail[Indice];
    TOBP.PutValue('BPE_TOTALHTDEV',0) ;
    TOBP.PutValue('BPE_TOTALTTCDEV',0) ;
    TOBP.PutValue('BPE_MONTANTPA',0) ;
    TOBP.PutValue('BPE_MONTANTPR',0) ;
  end;
end;

procedure TGenerePiece.RecalculeLeDocument;
  procedure InitCalcul;
  var II : Integer;            
  begin
    ReinitMontantPieceTrait;
    ZeroFacture (TOBpiece);
    ZeroMontantPorts (TOBPorcs);
    TOBBases.ClearDetail;
    TOBBasesL.ClearDetail;
    for II := 0 to TOBPiece.detail.Count -1 do ZeroLigneMontant(TOBPiece.detail[II]);
    PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X'); // positionne le recalcul du document
    TOBVTECOLLECTIF.ClearDetail;
  end;

var II,Imax : Integer;
    Coef,CoefTaxe,SumTaxe,MaxBase,SumCalcB,SumCalcT,Sumbase : Double;
    TobB : TOB;
    ArtEcart,Atest : string;
    QQ : TQuery;
    SQL : string;
begin
  ArtEcart := '';
  if VH_GC.BTCODESPECIF = '001' then
  begin
    if (VenteAchat = 'ACH') and (TOBPiece.GetString('GP_NATUREPIECEG')='FF') and (DocFromBSV) then
    begin
      SQL := 'SELECT YTC_TEXTELIBRE1 FROM TIERS LEFT JOIN TIERSCOMPL ON T_AUXILIAIRE=YTC_AUXILIAIRE WHERE T_TIERS="'+TOBPiece.getString('GP_TIERS')+'" AND T_NATUREAUXI="FOU"';
      QQ := OpenSQL(SQL,True,1,'',true);
      if not QQ.eof then
      begin
        Atest := QQ.fields[0].AsString;
        if Atest <> '' then
        begin
          if ExisteSQL('SELECT 1 FROM ARTICLE WHERE GA_CODEARTICLE="'+Atest+'" AND GA_TYPEARTICLE IN ("MAR","ARP")') then
          begin
            ArtEcart := Atest;
          end;
        end;
      end;
      Ferme(QQ);
      //
    end;
  end;
  if Artecart = '' then
  begin
    SQL := 'SELECT BP2_ARTICLE FROM BSVSERVER';
    QQ := OpenSQL(SQL,True,1,'',true);
    if not QQ.eof then
    begin
      Atest := QQ.fields[0].AsString;
      if Atest <> '' then
      begin
        if ExisteSQL('SELECT 1 FROM ARTICLE WHERE GA_ARTICLE="'+Atest+'" AND GA_TYPEARTICLE IN ("MAR","ARP")') then
        begin
          ArtEcart := Trim(Copy(Atest,1,18));
        end;
      end;
    end;
    Ferme(QQ);
  end;
  if Artecart = '' then
  begin
    Atest := GetParamSocSecur('SO_BTECARTBSVACH','');
    if Atest <> '' then
    begin
      if ExisteSQL('SELECT 1 FROM ARTICLE WHERE GA_CODEARTICLE="'+Atest+'" AND GA_TYPEARTICLE IN ("MAR","ARP")') then
      begin
        ArtEcart := Atest;
      end;
    end;
  end;
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'Recalcul du document';
  InitCalcul;
  CalculFacture(TOBAffaire,TOBPiece,TOBPieceTrait,TOBSSTRAIT,TOBOUvrage, TOBOuvragesP,TOBBases, TOBBasesL,TOBTiers, TOBArticles, TOBPorcs, TOBPieceRG, TOBBasesRG,TOBVTECOLLECTIF, DEV);
  TOBEches.ClearDetail;
//
	if (VenteAchat = 'ACH') and ((TOBPiece.GetString('GP_NATUREPIECEG')='FF') or (TOBPiece.GetString('GP_NATUREPIECEG')='AF')) and (DocFromBSV) then
  begin
    if (TOBREPARTTVA <> nil) and (TOBREPARTTVA.detail.count > 0) then
    begin
      SumTaxe := 0; Sumbase := 0;
      for II := 0 to TOBREPARTTVA.Detail.Count - 1 do
      begin
        TobB := TOBREPARTTVA.Detail[II];
        SumTaxe := SumTaxe + TobB.GetDouble('BP8_MONTANTTAXE');
        SumBase := SumBase + TobB.GetDouble('BP8_BASEHT');
      end;
      SumTaxe := ARRONDI(SumTaxe,DEV.Decimale);
      SumBase := ARRONDI(SumBase,DEV.Decimale);
      // --
      // --
      if (Arrondi(TOBPiece.GetDouble('GP_TOTALTTCDEV'),DEV.Decimale) <> ARRONDI(SumTaxe+SumBase,DEV.Decimale)) or
         (ARRONDI(TOBPiece.GetDouble('GP_TOTALHTDEV'),DEV.Decimale)<> ARRONDI(SumBase,DEV.Decimale)) then
      begin
        if ArtEcart <> '' then
        begin
          AffecteEcartArticleDocAchat (TOBPiece,TOBTiers,TOBAffaire,TOBArticles,Arrondi(SumBase-TOBPiece.GetDouble('GP_TOTALHTDEV'),DEV.Decimale),ArtEcart);
          InitCalcul;
          CalculFacture(TOBAffaire,TOBPiece,TOBPieceTrait,TOBSSTRAIT,TOBOUvrage, TOBOuvragesP,TOBBases, TOBBasesL,TOBTiers, TOBArticles, TOBPorcs, TOBPieceRG, TOBBasesRG,TOBVTECOLLECTIF, DEV);
        end;
        //
        TOBPiece.setDouble('GP_TOTALTTCDEV',ARRONDI(SumTaxe+SumBase,DEV.Decimale));
        TOBPiece.setDouble('GP_TOTALHTDEV',SumBase);
        TOBPiece.PutValue ('GP_TOTALHT',DeviseToEuro (TOBPiece.GetValue ('GP_TOTALHTDEV'), DEV.Taux, DEV.Quotite)) ;
        TOBPiece.PutValue ('GP_TOTALTTC',DeviseToEuro (TOBPiece.GetValue ('GP_TOTALTTCDEV'), DEV.Taux, DEV.Quotite)) ;
        // ----
        VH_GC.ModeGestionEcartComptable := 'CPA'; {DBR CPA}
      end;
    end else
    begin
      //
      Imax := -1; MaxBase := 0; SumTaxe := 0;
      for II := 0 to TOBBases.Detail.Count - 1 do
      begin
        TobB := TOBBases.Detail[II];
        SumTaxe := SumTaxe + TobB.GetDouble('GPB_VALEURDEV');
        if TobB.GetDouble('GPB_BASEDEV') > MaxBase then
        begin
          Imax := II;
          MaxBase := TobB.GetDouble('GPB_BASEDEV');
        end;
      end;
      //
      if (TOBPiece.GetDouble('GP_TOTALTTCDEV')<> TOBEnteteScan.GetDouble('B10_TOTALTTC')) or
         (TOBPiece.GetDouble('GP_TOTALHTDEV')<> TOBEnteteScan.GetDouble('B10_TOTALHT')) or
         (ARRONDI(SumTaxe,DEV.decimale)<> TOBEnteteScan.GetDouble('B10_TOTALTAXE')) then
      begin
        // --
        if ArtEcart <> '' then
        begin
          AffecteEcartArticleDocAchat (TOBPiece,TOBTiers,TOBAffaire,TOBArticles,(TOBEnteteScan.GetDouble('B10_TOTALHT')-TOBPiece.GetDouble('GP_TOTALHTDEV')),ArtEcart);
          InitCalcul;
          CalculFacture(TOBAffaire,TOBPiece,TOBPieceTrait,TOBSSTRAIT,TOBOUvrage, TOBOuvragesP,TOBBases, TOBBasesL,TOBTiers, TOBArticles, TOBPorcs, TOBPieceRG, TOBBasesRG,TOBVTECOLLECTIF, DEV);
        end;
        // --
        if TOBEnteteScan.GetDouble('B10_TOTALTTC') = 0 then
        begin
          TOBPiece.setDouble('GP_TOTALTTCDEV',TOBEnteteScan.GetDouble('B10_TOTALHT'));
          TOBPiece.setDouble('GP_TOTALHTDEV',TOBEnteteScan.GetDouble('B10_TOTALHT'));
        end else
        begin
          TOBPiece.setDouble('GP_TOTALTTCDEV',TOBEnteteScan.GetDouble('B10_TOTALTTC'));
          TOBPiece.setDouble('GP_TOTALHTDEV',TOBEnteteScan.GetDouble('B10_TOTALHT'));
        end;
        if  TOBEnteteScan.GetDouble('B10_TOTALHT') = 0 then Coef := 1
                                                       else Coef := TOBPiece.GetDouble('GP_TOTALHTDEV')/ TOBEnteteScan.GetDouble('B10_TOTALHT');
        if TOBEnteteScan.GetDouble('B10_TOTALTAXE') = 0 then CoefTaxe := 1
                                                        else CoefTaxe := SumTaxe/ TOBEnteteScan.GetDouble('B10_TOTALTAXE');
        //
        SumCalcB := 0;
        SumCalcT := 0;
        for II := 0 to TOBBases.Detail.Count - 1 do
        begin
          TobB := TOBBases.Detail[II];
          TobB.SetDouble ('GPB_BASEDEV',ARRONDI(TobB.getDouble ('GPB_BASEDEV') * Coef,DEV.decimale));
          TobB.SetDouble ('GPB_VALEURDEV',ARRONDI(TobB.getDouble ('GPB_VALEURDEV') * coef,DEV.decimale));
          SumCalcB := SumCalcB + TobB.getDouble ('GPB_BASEDEV');
          SumCalcT := SumCalcT + TobB.getDouble ('GPB_VALEURDEV');
        end;
        SumCalcB := ARRONDI(SumCalcB,DEV.decimale);
        SumCalcT := ARRONDI(SumCalcT,DEV.decimale);
        //
        if (SumCalcB <> TOBEnteteScan.GetDouble('B10_TOTALHT')) and (Imax <> -1)  then
        begin
          TOBBases.detail[Imax].SetDouble('GPB_BASEDEV', TOBBases.detail[Imax].GetDouble('GPB_BASEDEV') + (TOBEnteteScan.GetDouble('B10_TOTALHT')-SumCalcB)) ;
        end;
        if (SumCalcT <> TOBEnteteScan.GetDouble('B10_TOTALTAXE')) and (Imax <> -1)  then
        begin
          TOBBases.detail[Imax].SetDouble('GPB_VALEURDEV', TOBBases.detail[Imax].GetDouble('GPB_VALEURDEV') + (TOBEnteteScan.GetDouble('B10_TOTALTAXE')-SumCalcT)) ;
        end;

        if DEV.Code<>V_PGI.DevisePivot then
        begin
          for II := 0 to TOBBases.detail.count -1 do
          begin
            TOBBases.detail[II].SetDouble ('GPB_BASETAXE',DeviseToEuro (TOBBases.detail[II].GetDouble ('GPB_BASEDEV'), DEV.Taux, DEV.Quotite)) ;
            TOBBases.detail[II].SetDouble ('GPB_VALEURTAXE',DeviseToEuro (TOBBases.detail[II].GetDouble ('GPB_VALEURDEV'), DEV.Taux, DEV.Quotite)) ;
          end;
          TOBPiece.PutValue ('GP_TOTALHT',DeviseToEuro (TOBPiece.GetValue ('GP_TOTALHTDEV'), DEV.Taux, DEV.Quotite)) ;
          TOBPiece.PutValue ('GP_TOTALTTC',DeviseToEuro (TOBPiece.GetValue ('GP_TOTALTTCDEV'), DEV.Taux, DEV.Quotite)) ;
        end else
        begin
          for II := 0 to TOBBases.detail.count -1 do
          begin
            TOBBases.detail[II].SetDouble ('GPB_BASETAXE',TOBBases.detail[II].GetDouble ('GPB_BASEDEV')) ;
            TOBBases.detail[II].SetDouble ('GPB_VALEURTAXE',TOBBases.detail[II].GetDouble ('GPB_VALEURDEV')) ;
          end;
          TOBPiece.PutValue ('GP_TOTALHT',TOBPiece.GetValue ('GP_TOTALHTDEV')) ;
          TOBPiece.PutValue ('GP_TOTALTTC', TOBPiece.GetValue('GP_TOTALTTCDEV'));
        end;
        // ----
        VH_GC.ModeGestionEcartComptable := 'CPA'; {DBR CPA}
      end;
      //
    end;
  end;
  //
  GereEcheancesGC(TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG,TOBPieceTrait,TOBPorcs,taCreat,DEV,False) ;
  //
  fResult.ErrorResult := oeOK;
  fResult.LibError := '';
end;

procedure TGenerePiece.Reinit;
begin
  GestionConso.InitReceptions;
  fResult.init;
  TOBPiecePrec.clearDetail; TOBPiecePrec.InitValeurs(false);
  TOBPiece.ClearDetail; TOBPiece.InitValeurs(false);
  TOBVTECOLLECTIF.ClearDetail; TOBVTECOLLECTIF.InitValeurs(false);
  TOBAffaireInterv.ClearDetail; TOBAffaireInterv.InitValeurs(false);
  TOBPieceTrait.ClearDetail; TOBPieceTrait.InitValeurs(false);
  TOBBasesL.ClearDetail; TOBBasesL.InitValeurs(false);
  TOBBases.ClearDetail; TOBBases.InitValeurs(false);
  TOBEches.ClearDetail; TOBEches.InitValeurs(false);
  TOBPorcs.ClearDetail; TOBPorcs.InitValeurs(false);
  TOBTiers.ClearDetail; TOBTiers.InitValeurs(false);
  TOBArticles.ClearDetail; TOBArticles.InitValeurs(false);
  TOBConds.ClearDetail; TOBConds.InitValeurs(false);
  TOBAdresses.ClearDetail; TOBAdresses.InitValeurs(false);
  TOBAcomptes.ClearDetail; TOBAcomptes.InitValeurs(false);
  TOBAffaire.ClearDetail; TOBAffaire.InitValeurs(false);
  TOBCPTA.ClearDetail; TOBCPTA.InitValeurs(false);
  TOBANAP.ClearDetail; TOBANAP.InitValeurs(false);
  TOBANAS.ClearDetail; TOBANAS.InitValeurs(false);
  TOBOuvrage.ClearDetail; TOBOuvrage.InitValeurs(false);
  TOBOuvragesP.ClearDetail; TOBOuvragesP.InitValeurs(false);
  TOBLIENOLE.ClearDetail; TOBLIENOLE.InitValeurs(false);
  TOBPieceRG.ClearDetail; TOBPieceRG.InitValeurs(false);
  TOBBasesRG.ClearDetail; TOBBasesRG.InitValeurs(false);
  TOBSSTRAIT.ClearDetail; TOBSSTRAIT.InitValeurs(false);
  TOBDispo.ClearDetail; TOBDispo.InitValeurs(false);
  VH_GC.ModeGestionEcartComptable := '';
  TOBREPARTTVA.ClearDetail;
end;

procedure TGenerePiece.RenumerotePiece;
var MaxNumOrdre : Integer;
    I,inl,Ino : Integer;
    TOBL : TOB;
begin
  MaxNumOrdre := 1;
  inl := -1; ino := -1;
  PutValueDetail(TOBPiece,'GP_RECALCULER','X');
  for i := 0 to TOBPiece.Detail.Count-1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, i+1);
    if i=0 then
    begin
      inl := TOBL.GetNumChamp('GL_NUMLIGNE') ;
      ino := TOBL.GetNumChamp('GL_NUMORDRE') ;
    end;
    TOBL.PutValeur(inl, i+1);
    Inc(MaxNumOrdre);
    TOBL.PutValeur(ino, MaxNumOrdre);
  end;
  TOBPiece.PutValue('GP_CODEORDRE', MaxNumOrdre);
end;

procedure TGenerePiece.UpdateDocumentsPrecedent;

  function IsPieceVivante (TOBP : TOB) : Boolean;
  var Req : string;
  begin
    Req := 'SELECT 1 FROM LIGNE WHERE GL_NATUREPIECEG="' + TOBP.GetString('NATUREPIECEG');
    Req := Req + '" AND GL_SOUCHE="' + TOBP.GetString('SOUCHE');
    Req := Req + '" AND GL_NUMERO= ' +  TOBP.GetString('NUMERO');
    Req := Req + '  AND GL_INDICEG=' + TOBP.GetString('INDICEG');
    Req := Req + '  AND GL_TYPELIGNE="ART" AND GL_VIVANTE="X"';
    Result := ExisteSQL(Req);
  end;

  function GerePiecePrecVivante (TOBPiece,TOBPP: TOB) : Boolean;
  var TOBL : TOB;
      II : Integer;
      cledoc : R_CLEDOC;
  begin
    Result := false;
    if TOBPP.Detail.count = 0 then Exit;
    for II := 0 to TOBPP.detail.count -1 do
    begin
      TOBL := TOBPP.detail[II];
      cledoc.NaturePiece := TOBL.GetString('NATUREPIECEG');
      cledoc.Souche := TOBL.GetString('SOUCHE');
      cledoc.NumeroPiece := TOBL.GetInteger('NUMERO');
      cledoc.Indice := TOBL.GetInteger('INDICEG');
      if GereReliquat then
      begin
        if not IsPieceVivante(TOBL) then
        begin
         if (ExecuteSQL('UPDATE PIECE SET GP_VIVANTE="-",GP_DEVENIRPIECE="'+encoderefPiece(TOBPiece)+'" WHERE '+WherePiece(cledoc,ttdPiece,false)) <= 0) then exit;
        end else
        begin
         if (ExecuteSQL('UPDATE PIECE SET GP_DEVENIRPIECE="'+encoderefPiece(TOBPiece)+'" WHERE '+WherePiece(cledoc,ttdPiece,false)) <=0) then exit;
        end;
      end else
      begin
        if (ExecuteSQL('UPDATE PIECE SET GP_VIVANTE="-",GP_DEVENIRPIECE="'+encoderefPiece(TOBPiece)+'" WHERE '+WherePiece(cledoc,ttdPiece,false)) <=0) then exit;
      end;
    end;
    Result := True;
  end;

  function SetPieceMorte (TOBPiece,TOBPP: TOB) : Boolean;
  var TOBL : TOB;
      II : Integer;
      SQL : string;
  begin
    Result := false;
    if TOBPP.Detail.count = 0 then Exit;
    for II := 0 to TOBPP.detail.count -1 do
    begin
      TOBL := TOBPP.detail[II];
      SQL := 'UPDATE PIECE SET GP_VIVANTE="-",GP_DEVENIRPIECE="'+EncoderefPiece(TOBPiece)+'" '+
             'WHERE '+
             'GP_NATUREPIECEG="'+TOBL.GetString('B12_NATUREORIGINE')+'" AND '+
             'GP_SOUCHE="'+TOBL.GetString('B12_SOUCHEORIGINE')+'" AND '+
             'GP_NUMERO='+IntToStr(TOBL.GetInteger('B12_NUMEROORIGINE'))+' AND '+
             'GP_INDICEG='+IntToStr(TOBL.GetInteger(' B12_INDICEORIGINE'));
      if (ExecuteSQL(SQL) <= 0) then exit;
      SQL := 'UPDATE LIGNE SET GL_VIVANTE="-" '+
             'WHERE '+
             'GL_NATUREPIECEG="'+TOBL.GetString('B12_NATUREORIGINE')+'" AND '+
             'GL_SOUCHE="'+TOBL.GetString('B12_SOUCHEORIGINE')+'" AND '+
             'GL_NUMERO='+IntToStr(TOBL.GetInteger('B12_NUMEROORIGINE'))+' AND '+
             'GL_INDICEG='+IntToStr(TOBL.GetInteger(' B12_INDICEORIGINE'));
      if (ExecuteSQL(SQL) <= 0) then exit;
    end;
    Result := True;
  end;

  procedure PositionneLigneMorte (TOBL : TOB;cledoc : r_cledoc);
  var QQ : TQuery;
      SQL : String;
  begin
    // on ne r�cup�re que la ligne
    SQL := 'SELECT * '+
           'FROM LIGNE '+
           'WHERE '+ WherePiece(CleDoc, ttdLigne, true,true);
    QQ := OpenSQL(SQL,false,1);
    if not QQ.eof then
    begin
      QQ.Edit;
      QQ.FindField('GL_VIVANTE').AsString := '-';
      QQ.Post;
    end;
    QQ.Close;
  end;


  function TraiteReliquatLigne (TOBL : TOB;cledoc : r_cledoc) : Boolean;
  var TOBLP : TOB;
      QQ : TQuery;
      SQL : String;
  begin
    Result := false;
    TOBLP := TOB.Create ('LIGNE',nil,-1);
    TRY
      // on ne r�cup�re que la ligne
      SQL := 'SELECT GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_QTEFACT,GL_QTERESTE,GL_QTERELIQUAT,GL_QTESTOCK,GL_MTRESTE,GL_MTRELIQUAT,GL_VIVANTE,(SELECT GA_RELIQUATMT FROM ARTICLE WHERE GA_ARTICLE=GL_ARTICLE) AS RELIQUATMT '+
             'FROM LIGNE '+
             'WHERE '+ WherePiece(CleDoc, ttdLigne, true,true);
      QQ := OpenSQL(SQL,true,1,'',true);
      if not QQ.eof then
      begin
        TOBLP.SelectDB('',QQ);
        ferme(QQ);
        if TOBLP.GetString('RELIQUATMT')='X' then
        begin
          if TOBLP.GetDouble('GL_MTRESTE') <= TOBL.GetDouble('GL_MONTANTHT') then
          begin
            TOBLP.SetDouble('GL_MTRESTE',0);
            TOBLP.SetBoolean('GL_VIVANTE',false);
          end else
          begin
            TOBLP.SetDouble('GL_MTRESTE',TOBLP.GetDouble('GL_MTRESTE')-TOBL.GetDouble('GL_MONTANTHT'));
          end;
          TOBL.SetDouble('GL_MTRESTE',TOBL.GetDouble('GL_MONTANTHT'));
        end else
        begin
          if TOBLP.GetDouble('GL_QTERESTE') <= TOBL.GetDouble('GL_QTEFACT') then
          begin
            TOBLP.SetDouble('GL_QTERESTE',0);
            TOBLP.SetBoolean('GL_VIVANTE',false);
          end else
          begin
            TOBLP.SetDouble('GL_QTERESTE',TOBLP.GetDouble('GL_QTERESTE')-TOBL.GetDouble('GL_QTEFACT'));
          end;
          //
          TOBL.SetDouble('GL_QTERELIQUAT',TOBLP.GetDouble('GL_QTERESTE'));
          TOBL.SetDouble('GL_QTESAIS',TOBL.GetDouble('GL_QTESTOCK'));
        end;
        //
        if (not TOBLP.UpdateDB(false)) then exit;
      end else
      begin
        ferme (QQ);
      end;
      Result := True;
    FINALLY
      TOBLP.Free;
    END;
  end;

var II : Integer;
    TOBL : TOB;
    cledoc : r_cledoc;
begin
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'Mise � jour pi�ce pr�c�dente';

  for II := 0 to TOBPiece.Detail.count -1 do
  begin
    TOBL := TOBPiece.detail[II];
    if TOBL.GetString('GL_PIECEPRECEDENTE')<>'' then
    begin
      DecodeRefPiece(TOBL.GetString('GL_PIECEPRECEDENTE'),cledoc);
      if GereReliquat then
      begin
        if not TraiteReliquatLigne (TOBL,cledoc) then
        begin
          raise Exception.Create('Mise � jour pi�ce pr�c�dente / ligne');
          exit;
        end;
      end else
      begin
        PositionneLigneMorte(TOBL,cledoc);
      end;
      AddReferenceDoc (cledoc,TOBPiecePrec);
    end;
  end;
  //
  if TOBPiecePrec.detail.count > 0 then
  begin
    if not GerePiecePrecVivante (TOBPiece,TOBPiecePrec) then
    begin
      raise Exception.Create('Mise � jour pi�ce pr�c�dente / Entete');
      exit;
    end;
  end;
  
  if TOBprov <> nil then
  begin
    if TOBProv.detail.count > 0 then
    begin
      if not SetPieceMorte (TOBPiece,TOBPROV) then
      begin
        raise Exception.Create('Mise � jour pi�ce pr�c�dente / Entete');
        exit;
      end;

    end;
  end;
  fResult.ErrorResult := oeOk;
  fResult.LibError := '';
end;

function TGenerePiece.GetNumeroPiece : integer;
var newSouche : string;
begin
  Result := 0;
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'R�cup�ration N� Document';
  NewSouche := GetSoucheG (TOBPiece.getValue('GP_NATUREPIECEG'),TOBPiece.getValue('GP_ETABLISSEMENT'),TOBPiece.getValue('GP_DOMAINE'));
  if NewSouche = '' then Exit;
  TOBPiece.setString('GP_SOUCHE',newSouche);
  result := SetNumberAttribution(TOBPiece.getValue('GP_NATUREPIECEG'),newSouche, TOBPiece.getValue('GP_DATEPIECE'),1);
  fResult.ErrorResult := oeOk;
  fResult.LibError := '';
end;

procedure TGenerePiece.ValideLaNumerotation;
var NaturePieceG: string;
    NewNum : Integer;
begin
  newNum := GetNumeroPiece;
  if fResult.ErrorResult <> OeOk then Exit;
  //
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'ATTRIBUTION N� Document';
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
  if not SetDefinitiveNumber(TOBPiece, TOBBases, TOBBasesL,TOBEches, nil, TOBAcomptes, TOBPIeceRG, TOBBasesRg, nil, NewNum) then Exit;
  if GetInfoParPiece(NaturePieceG, 'GPP_ACTIONFINI') = 'ENR' then TOBPiece.PutValue('GP_VIVANTE', '-');
  fResult.ErrorResult := oeOK;
  fResult.LibError := '';
end;

procedure TgenerePiece.GenereCompta ;
begin
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'Passation comptable';
  if Not PassationComptable(TOBPiece,TOBOuvrage, TOBOuvragesP,TOBBases,TOBBasesL,TOBEches,TOBPieceTrait,nil,
                            TOBTiers,TOBArticles,TOBCpta,TOBAcomptes,TOBPorcs,TOBPIECERG,TOBBasesRG,TOBAnaP,TOBAnaS,nil,TOBVTECOLLECTIF,TOBREPARTTVA,TOBVENTILTVA, DEV,OldEcr,OldStk,false,true,FBAP) then
  begin
    V_PGI.IOError := oeLettrage;
    exit;
  end;
  fResult.ErrorResult := oeOk;
  fResult.LibError := '';
end;

function TGenerePiece.ChargeTOBA(RefUnique : string; stDepot : string) : TOB; // DBR : D�pot unique charg�
var Q : TQuery;
    SQL : String;
    TobArt : TOB;
begin
  TOBART := TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],true);
  if TobArt= nil then
  begin
    SQL := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
           '"" AS REFARTSAISIE, '+
           '"" AS REFARTBARRE, '+
           '"" AS REFARTTIERS, '+
           '"" AS _FROMOUVRAGE, '+
           '"-" AS SUPPRIME, '+
           '"-" AS UTILISE '+
          'FROM ARTICLE A '+
          'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
          'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
          'WHERE GA_ARTICLE="'+RefUnique+'"';
    Q:=OpenSQL(SQL,True,-1,'',true) ;
    if Not Q.EOF then
    begin
      TobArt := CreerTOBArt(TOBArticles);
      TobArt.SelectDB('',Q);
      LoadTOBDispo(TobArt, True, '"' + stDepot + '"') ; // DBR : D�pot unique charg�
    end;
    Ferme(Q);
  end;
  Result:=TobArt;
end;

procedure TGenerePiece.ValideLaRepartTVA (TOBPiece,TOBREPARTTVA : TOB);
var II : Integer;
begin
  for II := 0 to TOBREPARTTVA.detail.count -1 do
  begin
    TOBREPARTTVA.detail[II].SetString('BP8_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG'));
    TOBREPARTTVA.detail[II].SetString('BP8_SOUCHE',TOBPiece.GetString('GP_SOUCHE'));
    TOBREPARTTVA.detail[II].SetString('BP8_NUMERO',TOBPiece.GetString('GP_NUMERO'));
    TOBREPARTTVA.detail[II].SetString('BP8_INDICEG',TOBPiece.GetString('GP_INDICEG'));
  end;
  if not TOBREPARTTVA.InsertDB(nil) then
  begin
    V_PGI.IOError := OeUnknown;
  end;
end;

procedure TGenerePiece.InverseAvoir;
begin
  InverseLesPieces(TOBPiece, 'PIECE');
  InverseLesPieces(TOBBases, 'PIEDBASE');
  InverseLesPieces(TOBEches, 'PIEDECHE');
  InverseLesPieces(TOBPorcs, 'PIEDPORT'); //mcd 07/06/02 port non pris en compte
  InverseLesPieces(TOBPieceRG, 'PIECERG');
  InverseLesPieces(TOBBasesRG, 'PIEDBASERG');
  InverseLesPieces(TOBPieceTrait, 'PIECETRAIT');
end;

function TGenerePiece.IsAvoir(TOBpiece : TOB) : Boolean;
begin
  result := (GetInfoParPiece(TOBpiece.GetString('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X');
end;


procedure TGenerePiece.ValideLeDocument;
var II : Integer;
begin
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'Validation de la pi�ce';
  //
  V_PGI.IOError := oeOk;
  //
  if IsAvoir(TOBpiece) then InverseAvoir;
  //
  ValideLaNumerotation;
  if fResult.ErrorResult <> OeOk then Exit;
  ValideLesArticlesFromOuv(TOBarticles,TOBOuvrage);
  //
  ValideLaRepartTVA (TOBPiece,TOBREPARTTVA);
  InsereLigneRG(TOBPIece, TOBPIeceRG, TOBBASES, TOBBASESRG, TOBTIers, TOBAffaire, nil, -1);
  if V_PGI.IoError=oeOk then ValideLesLignes(TOBPiece,TOBArticles,nil,nil,TOBOuvrage,TOBPieceRG,TOBBasesRG,GereReliquat,false,false,true) ;
  if V_PGI.IoError=oeOk then ValideLesLignesCompl(TOBPiece, nil);
  if V_PGI.IoError=oeOk then ValideLesAdresses(TOBPiece,nil,TOBAdresses) ;
  if V_PGI.IoError=oeOk then ValideLesArticles(TOBPiece,TOBArticles) ;
  if V_PGI.IoError=oeOk then ValideleTiers(TOBPiece,TOBTiers) ;
  if V_PGI.IoError=oeOk then ValideLesBases(TOBPiece,TobBases,TOBBasesL);
  if V_PGI.IoError=oeOk then ValideLesAdresses(TOBPiece, TOBPiece, TOBAdresses);
  if V_PGI.IoError=oeOk then ValideLesRetenues(TOBPiece, TOBPieceRG);
  if V_PGI.IoError=oeOk then ValideLesBasesRG(TOBPiece, TOBBasesRG);
  if V_PGI.IoError=oeOk then ValideLesPorcs(TOBPiece, TOBPorcs);

  if V_PGI.IoError=oeOk then GenereCompta ;
  if V_PGI.IOError = OeOk then
  begin
    PrepareInsertCollectif (TOBPiece,TOBVTECOLLECTIF);
    if not TOBVTECOLLECTIF.InsertDB(nil) then V_PGI.IOError := OeUnknown;
  end;
  if V_PGI.IOError = OeOk then
  begin
    RepartTva.EcritTOB(TOBPiece,TOBVENTILTVA);
  end;
  if V_PGI.IoError = oeOk then ValideLesBases(TOBPiece,TobBases,TOBBasesL);
  for II := 0 to TOBpiece.detail.count -1 do
  begin
    TRY
      ProtectionCoef (TOBPiece.detail[II]);
    EXCEPT
      on E: Exception do
      begin
        V_PGI.IOError := OeUnknown;
      end;
    END;
  end;

  TRY
    if V_PGI.IoError=oeOk then TOBPiece.InsertOrUpdateDB(false) ;
  EXCEPT
    on E: Exception do
    begin
      V_PGI.IOError := OeUnknown;
    end;
  END;
  if V_PGI.IoError = oeOk then TOBBasesL.InsertDB(nil);
  if V_PGI.IoError=oeOk then TOBBases.InsertDB(Nil) ;
  if V_PGI.IoError=oeOk then TOBEches.InsertOrUpdateDB(True) ;

  if V_PGI.IoError=oeOk then TOBAnaP.InsertDB(Nil) ;
  if V_PGI.IoError=oeOk then TOBAnaS.InsertDB(Nil) ;

  if (V_PGI.IOError = OeOk) Then
  begin
    TRY
      GestionConso.GenerelesPhases(TOBPiece,nil,true,false,false,taCreat);
    EXCEPT
      on E: Exception do
      begin
        V_PGI.IOError := oeUnknown;
      end;
    END;
  end;
  if (V_PGI.IOError = OeOk) Then GestionConso.clear;


  if (V_PGI.IOError = OeOk) Then
  begin
    TRY
      GenereLivraisonClients(TOBPiece,TaCreat,true,false,false,false);
    EXCEPT
      on E: Exception do
      begin
        V_PGI.IOError := oeUnknown;
      end;
    END;
  end;

  if V_PGI.IOError = OeOk then
  begin
    if IsLivraisonClient(TOBPiece) then
    begin
      UpdateStatusMoisOD (TOBPiece);
    end;
  end;



  if (V_PGI.IOError = OeOk) and (fResult.ErrorResult = OeOk) then
  begin
    fResult.ErrorResult := oeOk;
    fResult.LibError := '';
    fResult.RefPiece := EncodeRefPresqueCPGescom(TOBPiece);
    fResult.NatureDoc := TOBPiece.getString('GP_NATUREPIECEG');
    fResult.NumeroDoc := TOBPiece.getInteger('GP_NUMERO');
  end;
end;

procedure TGenerePiece.ConstitueTobArticles;
var SQL,SQLBIS : String;
    II : Integer;
    QQ : TQuery;
    DejaUn : Boolean;
begin
  DejaUn := false;
  if (TOBPiecePrec.detail.count = 0) and ((TOBProv=nil) or ((TOBProv<>nil) and (tobProv.detail.count=0))) then exit;
  //
  fResult.ErrorResult := oeUnknown;
  fResult.LibError  := 'Lecture articles';
  //
  // on recup d'un coup tous les articles des differents documents et ce en provenance des lignes ou des lignes d'ouvrages
  //
  SQL := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
         '"" AS REFARTSAISIE, '+
         '"" AS REFARTBARRE, '+
         '"" AS REFARTTIERS, '+
         '"" AS _FROMOUVRAGE, '+
         '"-" AS SUPPRIME, '+
         '"-" AS UTILISE '+
        'FROM ARTICLE A '+
        'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
        'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE ';
  for II := 0 to TOBPiecePrec.detail.count -1 do
  begin
    if TOBPiecePrec.detail[II].getString('NATUREPIECEG')='' then continue;
    DejaUn := True;
    if II > 0 then SQLBIS := SQLBIS + ' UNION ';
    SQLBIS := SQLBIS + 'SELECT GL_ARTICLE AS REFARTICLE FROM LIGNE '+
                      'WHERE '+
                      'GL_NATUREPIECEG="'+TOBPiecePrec.detail[II].getString('NATUREPIECEG')+ '" AND '+
                      'GL_SOUCHE="'+TOBPiecePrec.detail[II].getString('SOUCHE')+'" AND '+
                      'GL_NUMERO='+InttoStr(TOBPiecePrec.detail[II].GetInteger('NUMERO'))+ ' AND '+
                      'GL_INDICEG='+InttoStr(TOBPiecePrec.detail[II].GetInteger('INDICEG'));
    SQLBIS := SQLBIS + ' UNION ';
    SQLBIS := SQLBIS + 'SELECT BLO_ARTICLE AS REFARTICLE FROM LIGNEOUV '+
                      'WHERE '+
                      'BLO_NATUREPIECEG="'+TOBPiecePrec.detail[II].getString('NATUREPIECEG')+'" AND '+
                      'BLO_SOUCHE="'+TOBPiecePrec.detail[II].getString('SOUCHE')+'" AND '+
                      'BLO_NUMERO='+InttoStr(TOBPiecePrec.detail[II].GetInteger('NUMERO'))+ ' AND '+
                      'BLO_INDICEG='+InttoStr(TOBPiecePrec.detail[II].GetInteger('INDICEG'));
  end;
  if TOBProv <> nil then
  begin
    for II := 0 to TOBProv.detail.count -1 do
    begin
      if (II > 0) or (DejaUn) then SQLBIS := SQLBIS + ' UNION ';
      SQLBIS := SQLBIS + 'SELECT GL_ARTICLE AS REFARTICLE FROM LIGNE '+
                        'WHERE '+
                        'GL_NATUREPIECEG="'+TOBProv.detail[II].getString('B12_NATUREORIGINE')+ '" AND '+
                        'GL_SOUCHE="'+TOBProv.detail[II].getString('B12_SOUCHEORIGINE')+'" AND '+
                        'GL_NUMERO='+IntToStr(TOBProv.detail[II].GetInteger('B12_NUMEROORIGINE'))+ ' AND '+
                        'GL_INDICEG='+InttoStr(TOBProv.detail[II].GetInteger('B12_INDICEORIGINE'));
      SQLBIS := SQLBIS + ' UNION ';
      SQLBIS := SQLBIS + 'SELECT BLO_ARTICLE AS REFARTICLE FROM LIGNEOUV '+
                        'WHERE '+
                        'BLO_NATUREPIECEG="'+TOBProv.detail[II].getString('B12_NATUREORIGINE')+'" AND '+
                        'BLO_SOUCHE="'+TOBProv.detail[II].getString('B12_SOUCHEORIGINE')+'" AND '+
                        'BLO_NUMERO='+InttoStr(TOBProv.detail[II].GetInteger('B12_NUMEROORIGINE'))+ ' AND '+
                        'BLO_INDICEG='+InttoStr(TOBProv.detail[II].GetInteger('B12_INDICEORIGINE'));

    end;
  end;
  if DejaUn then
  begin
    SQL := SQL + 'WHERE GA_ARTICLE IN ('+ SQLBIS+')';
    QQ := OpenSQL (SQL,True,-1,'',True);
    if not QQ.eof then
    begin
      TOBArticles.LoadDetailDB('ARTICLE','','',QQ,false);
    end;
    ferme (QQ);
  end;
  fResult.ErrorResult := oeOk;
  fResult.LibError  := '';
end;

procedure TGenerePiece.ConstitueTobDispo;
var SQL,SQLART,SQLDEPOT : String;
    II : Integer;
    QQ : TQuery;
    TOBA,TOBDA : TOB;
    AuMoinsUn : boolean;
begin
  AuMoinsUn := false;
  if (TOBPiecePrec.detail.count = 0) and ((TOBProv=nil) or ((TOBProv<>nil) and (tobProv.detail.count=0))) then exit;
  fResult.ErrorResult := oeUnknown;
  fResult.LibError  := 'Lecture dispo';
  //
  // on recup d'un coup tous les dispo de tous les articles de tous les depots des documents constituants
  //
  SQL := 'SELECT * FROM DISPO ';

  for II := 0 to TOBPiecePrec.detail.count -1 do
  begin
    if TOBPiecePrec.detail[II].getString('NATUREPIECEG')='' then continue;
    AuMoinsUn := true;
    if II > 0 then
    begin
      SQLART := SQLART + ' UNION ';
      SQLDEPOT := SQLDEPOT + ' UNION ';
    end;
    // Article du document (ligne)
    SQLART := SQLART + 'SELECT GL_ARTICLE AS REFARTICLE FROM LIGNE '+
                      'WHERE '+
                      'GL_NATUREPIECEG="'+TOBPiecePrec.detail[II].getString('NATUREPIECEG')+ '" AND '+
                      'GL_SOUCHE="'+TOBPiecePrec.detail[II].getString('SOUCHE')+'" AND '+
                      'GL_NUMERO='+IntToStr(TOBPiecePrec.detail[II].getInteger('NUMERO'))+ ' AND '+
                      'GL_INDICEG='+IntToStr(TOBPiecePrec.detail[II].GetInteger('INDICEG'));
    SQLART := SQLART + ' UNION ';
    // Articles du document (ligne)
    SQLART := SQLART + 'SELECT BLO_ARTICLE AS REFARTICLE FROM LIGNEOUV '+
                      'WHERE '+
                      'BLO_NATUREPIECEG="'+TOBPiecePrec.detail[II].getString('NATUREPIECEG')+'" AND '+
                      'BLO_SOUCHE="'+TOBPiecePrec.detail[II].getString('SOUCHE')+'" AND '+
                      'BLO_NUMERO='+IntToStr(TOBPiecePrec.detail[II].GetInteger('NUMERO'))+ ' AND '+
                      'BLO_INDICEG='+IntToStr(TOBPiecePrec.detail[II].GetInteger('INDICEG'));
    // Articles du document (ligneouv)
    SQLDEPOT := SQLDEPOT + 'SELECT GL_DEPOT FROM LIGNE '+
                        'WHERE '+
                        'GL_NATUREPIECEG="'+TOBPiecePrec.detail[II].getString('NATUREPIECEG')+ '" AND '+
                        'GL_SOUCHE="'+TOBPiecePrec.detail[II].getString('SOUCHE')+'" AND '+
                        'GL_NUMERO='+IntToStr(TOBPiecePrec.detail[II].GetInteger('NUMERO'))+ ' AND '+
                        'GL_INDICEG='+IntToStr(TOBPiecePrec.detail[II].GetInteger('INDICEG'));
  end;
  if TOBProv <> nil then
  begin
    for II := 0 to TOBProv.detail.count -1 do
    begin
      AuMoinsUn := true;

      if II > 0 then
      begin
        SQLART := SQLART + ' UNION ';
        SQLDEPOT := SQLDEPOT + ' UNION ';
      end;
      // Article du document (ligne)
      SQLART := SQLART + 'SELECT GL_ARTICLE AS REFARTICLE FROM LIGNE '+
                        'WHERE '+
                        'GL_NATUREPIECEG="'+TOBProv.detail[II].getString('B12_NATUREPIEORIGNE')+ '" AND '+
                        'GL_SOUCHE="'+TOBProv.detail[II].getString('B12_SOUCHEORIGNE')+'" AND '+
                        'GL_NUMERO='+InttoStr(TOBProv.detail[II].getInteger('B12_NUMEROORIGNE'))+ ' AND '+
                        'GL_INDICEG='+InttoStr(TOBProv.detail[II].getInteger('B12_INDICEORIGNE'));
      SQLART := SQLART + ' UNION ';
      // Articles du document (ligne)
      SQLART := SQLART + 'SELECT BLO_ARTICLE AS REFARTICLE FROM LIGNEOUV '+
                        'WHERE '+
                        'BLO_NATUREPIECEG="'+TOBProv.detail[II].getString('B12_NATUREPIEORIGNE')+'" AND '+
                        'BLO_SOUCHE="'+TOBProv.detail[II].getString('B12_SOUCHEORIGNE')+'" AND '+
                        'BLO_NUMERO='+InttoStr(TOBProv.detail[II].getInteger('B12_NUMEROORIGNE'))+ ' AND '+
                        'BLO_INDICEG='+IntToStr(TOBProv.detail[II].getInteger('B12_INDICEORIGNE'));
      // Articles du document (ligneouv)
      SQLDEPOT := SQLDEPOT + 'SELECT GL_DEPOT FROM LIGNE '+
                          'WHERE '+
                          'GL_NATUREPIECEG="'+TOBProv.detail[II].getString('B12_NATUREPIEORIGNE')+ '" AND '+
                          'GL_SOUCHE="'+TOBProv.detail[II].getString('B12_SOUCHEORIGNE')+'" AND '+
                          'GL_NUMERO='+IntToStr(TOBProv.detail[II].getInteger('B12_NUMEROORIGNE'))+ ' AND '+
                          'GL_INDICEG='+IntToStr(TOBProv.detail[II].getInteger('B12_INDICEORIGNE'));
    end;
  end;
  if AuMoinsUn then
  begin
    SQL := SQL + 'WHERE GQ_ARTICLE IN ('+ SQLART+') AND GQ_DEPOT IN ('+SQLDEPOT+')';
    QQ := OpenSQL (SQL,True,-1,'',True);
    if not QQ.eof then
    begin
      TOBDispo.LoadDetailDB('DISPO','','',QQ,false);
    end;
    ferme (QQ);
    //
    for II := 0 TO TOBArticles.Detail.Count -1 do
    begin
      TOBA := TOBArticles.detail[II];
      TobDA := TOBDispo.FindFirst(['GQ_ARTICLE'], [TOBA.GetValue('GA_ARTICLE')], False);
      while TOBDA <> nil  do
      begin
        TOBDA.ChangeParent(TOBA,-1);
        DispoChampSupp (TOBDA);
        TobDA := TOBDispo.FindNext(['GQ_ARTICLE'], [TOBA.GetValue('GA_ARTICLE')], False);
      end;
    end;
    //
  end;
  fResult.ErrorResult := oeOk;
  fResult.LibError  := '';
end;

procedure TGenerePiece.AjouteArticlesPLus;
var II : Integer;
    TOBA,TOBAA : TOB;
begin
  fResult.ErrorResult := oeUnknown;
  fResult.LibError  := 'Ajout articles';
  if TOBArticlePlus <> nil then
  begin
    for II := 0 to TOBArticlePlus.Detail.count -1 do
    begin
      TOBA := TOBArticlePlus.detail[II];
      TOBAA := TOBArticles.FindFirst(['GA_ARTICLE'],[TOBA.GetString ('GA_ARTICLE')],true );
      if TOBAA = nil then
      begin
        TOBAA := TOB.create ('ARTICLE',TOBArticles,-1);
        TOBAA.Dupliquer(TOBA,True,true);
      end;
    end;
  end;
  fResult.ErrorResult := oeOk ;
  fResult.LibError  := '';
end;

procedure TGenerePiece.ConstituePieceFromBast(TOBBAST: TOB; DateFac : TDateTime);

  function GetTauxTaxe (TOBpiece : TOB; VenteAchat : string; RegimeTaxe ,FamilleTaxe : string) : double;
  var TOBT : TOB;
      Taux : Double;
  begin
    Taux := 0;
    TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',RegimeTaxe,FamilleTaxe],False) ;
    if TOBT<>Nil then
    BEGIN
       if (RegimeTaxe = 'INT') and (venteAchat='ACH') then
       begin
         Taux := 0;
       end else if (TOBPiece.getBoolean('GP_AUTOLIQUID')) then
       begin
         Taux := 0;
       end else
       begin
         if VenteAchat='VEN' then
         begin
          Taux:=TOBT.GetValue('TV_TAUXVTE')
         end else
         begin
          Taux:=TOBT.GetValue('TV_TAUXACH') ;
         end;
       end;
    END ;
    result := Taux;
  end;

  function  AjouteLigneArticle (TOBPiece,TLBAST,TParam : TOB ; var NbTrait : Integer ; PhaseTravaux : string =''; Sens : string='+') : boolean;
  var TOBL : TOB;
      TOBA : TOB;
      SQL : String;
      QQ : TQuery;
  begin
    TOBA := nil;
    result := false;
    SQL := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
           '"" AS REFARTSAISIE, '+
           '"" AS REFARTBARRE, '+
           '"" AS REFARTTIERS, '+
           '"" AS _FROMOUVRAGE, '+
           '"-" AS SUPPRIME, '+
           '"-" AS UTILISE '+
          'FROM ARTICLE A '+
          'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
          'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
          'WHERE GA_ARTICLE="'+TPARAM.getString('BM6_ARTICLE')+'"';
    QQ := OpenSQL(SQL,True,1,'',true);
    if not QQ.eof then
    begin
      TOBA := TOB.Create('ARTICLE',TOBArticles,-1);
      TOBA.SelectDB('',QQ);
      result := true;
    end;
    ferme (QQ);
    if not result then Exit;
    TOBL := NewTOBLigne(TobPiece, 0);
    PieceVersLigne(TobPiece,TOBL,True);
    TOBL.PutValue('GL_PERIODE', GetPeriode(TOBL.GetValue('GL_DATEPIECE')));
    TOBL.PutValue('GL_SEMAINE', NumSemaine(TOBL.GetValue('GL_DATEPIECE')));
    TOBL.PutValue('GL_ARTICLE', TOBA.GetValue('GA_ARTICLE'));
    TOBL.PutValue('GL_REFARTSAISIE', TOBA.GetValue('GA_CODEARTICLE'));
    TOBL.PutValue('GL_CODEARTICLE', TOBA.GetValue('GA_CODEARTICLE'));
    TOBL.PutValue('GL_TYPEREF', 'ART');
    ArticleVersLigne(TOBPiece,TOBA,nil,TOBL,TOBTiers);
    TOBL.PutValue('GL_LIBELLE', TLBAST.GetString('BM5_LIBELLE'));
    if TParam.getString('BM6_SENS')='+' then
    begin
      TOBL.SetDouble('GL_QTEFACT', 1);
    end else
    begin
      TOBL.SetDouble('GL_QTEFACT', -1);
    end;
    TOBL.SetDouble('GL_QTESTOCK', TOBL.GetDouble('GL_QTEFACT'));
    TOBL.SetDouble('GL_QTERESTE', TOBL.GetDouble('GL_QTEFACT'));
    TOBL.SetDouble('GL_PUHTDEV', TLBAST.Getdouble('BM5_MTSITUATION'));
    TOBL.SetDouble('GL_DPA', TLBAST.Getdouble('BM5_MTSITUATION'));
    TOBL.SetDouble('GL_DPR', TLBAST.Getdouble('BM5_MTSITUATION'));
    TOBL.PutValue('GL_COEFMARG', 0);
    TOBL.PutValue('GL_COEFFG', 1);
    TOBL.SetString('GL_TYPEDIM', 'NOR');
    TOBL.PutValue('GL_FAMILLETAXE1', TOBBAST.GetValue('BM4_FAMILLETAXE1'));
    TOBL.PutValue('GL_FAMILLENIV2', TOBBAST.GetValue('BM4_FAMILLENIV2'));
    TOBL.SetInteger('GL_IDENTIFIANTWOL', -1); // ajout LS pour g�n�ration des livraison chantier
    if PhaseTravaux <> '' then
    begin
      TOBL.SetString('BLP_PHASETRA', PhaseTravaux);
    end;
  end;

  function  AjouteLigneRestitutionRg (TOBPiece,TLBAST,TParam,TOBRRG : TOB ; var NbTrait : Integer ; PhaseTravaux : string =''; Sens : string='+') : boolean;
  var TOBL : TOB;
      TOBA : TOB;
      SQL : String;
      QQ : TQuery;
  begin
    TOBA := nil;
    result := false;
    SQL := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
           '"" AS REFARTSAISIE, '+
           '"" AS REFARTBARRE, '+
           '"" AS REFARTTIERS, '+
           '"" AS _FROMOUVRAGE, '+
           '"-" AS SUPPRIME, '+
           '"-" AS UTILISE '+
          'FROM ARTICLE A '+
          'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
          'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
          'WHERE GA_ARTICLE="'+TPARAM.getString('BM6_ARTICLE')+'"';
    QQ := OpenSQL(SQL,True,1,'',true);
    if not QQ.eof then
    begin
      TOBA := TOB.Create('ARTICLE',TOBArticles,-1);
      TOBA.SelectDB('',QQ);
      result := true;
    end;
    ferme (QQ);
    if not result then Exit;
    TOBL := NewTOBLigne(TobPiece, 0);
    PieceVersLigne(TobPiece,TOBL,True);
    TOBL.PutValue('GL_PERIODE', GetPeriode(TOBL.GetValue('GL_DATEPIECE')));
    TOBL.PutValue('GL_SEMAINE', NumSemaine(TOBL.GetValue('GL_DATEPIECE')));
    TOBL.PutValue('GL_ARTICLE', TOBA.GetValue('GA_ARTICLE'));
    TOBL.PutValue('GL_REFARTSAISIE', TOBA.GetValue('GA_CODEARTICLE'));
    TOBL.PutValue('GL_CODEARTICLE', TOBA.GetValue('GA_CODEARTICLE'));
    TOBL.PutValue('GL_TYPEREF', 'ART');
    ArticleVersLigne(TOBPiece,TOBA,nil,TOBL,TOBTiers);
    TOBL.PutValue('GL_LIBELLE', TLBAST.GetString('BM5_LIBELLE')+' TVA '+TOBRRG.GetString('TVARG'));
    if TParam.getString('BM6_SENS')='+' then
    begin
      TOBL.SetDouble('GL_QTEFACT', 1);
    end else
    begin
      TOBL.SetDouble('GL_QTEFACT', -1);
    end;
    TOBL.SetString('GL_FAMILLETAXE1',TOBRRG.GetString('TVARG'));
    TOBL.SetDouble('GL_QTESTOCK', TOBL.GetDouble('GL_QTEFACT'));
    TOBL.SetDouble('GL_QTERESTE', TOBL.GetDouble('GL_QTEFACT'));
    //
    TOBL.SetDouble('GL_PUHTDEV', TOBRRG.GetDouble('MONTANTRG'));
    TOBL.SetDouble('GL_PUTTCDEV', TOBRRG.GetDouble('MONTANTRG'));
    TOBL.SetDouble('GL_DPA', TOBL.GetDouble('GL_PUHTDEV'));
    TOBL.SetDouble('GL_DPR', TOBL.GetDouble('GL_PUHTDEV'));
    TOBL.PutValue('GL_COEFMARG', 0);
    TOBL.PutValue('GL_COEFFG', 1);
    TOBL.SetString('GL_TYPEDIM', 'NOR');
    TOBL.SetInteger('GL_IDENTIFIANTWOL', 0); // ajout LS pour �viter la g�n�ration de livraison chantier
    if PhaseTravaux <> '' then
    begin
      TOBL.SetString('BLP_PHASETRA', PhaseTravaux);
    end;
  end;

  function  AjouteLignePort (TOBPiece,TOBPorcs,TLBAST,TParam : TOB ; var NbTrait : Integer) : boolean;
  var CodePort : string;
      QQ : TQuery;
      TP,TOBPP : TOB;
  begin
    result := false;
    TP := TOB.Create('PORT', nil, -1);
    TRY
      if TParam <> nil then
      begin
        CodePort := TParam.GetString('BM6_ARTICLE');
      end else
      begin
        CodePort := TLBAST.GetString('BM5_ARTICLE');
      end;
      if CodePort = '' then Exit;
      QQ := OpenSQL('SELECT * FROM PORT WHERE GPO_CODEPORT="'+CodePort+'"',True,1,'',true);
      if not QQ.Eof then
      begin
        TP.SelectDB('', QQ);
        TOBPP := TOB.Create ('PIEDPORT',TOBPorcs,-1);
        TOBPP.SetString('GPT_CODEPORT', CodePort);
        ChargeTobPiedPort (TOBPP,TP,TOBPiece);
        TOBPP.SetString('GPT_LIBELLE',TLBAST.GetString('BM5_LIBELLE'));
        if (TParam.getString('BM6_SENS')='+') and (TP.GetString('GPO_RETENUEDIVERSE')='X') then
        begin
          TOBPP.SetDouble('GPT_BASETTCDEV',(TLBAST.getDouble('BM5_MTSITUATION')*(-1)));
        end else
        begin
          TOBPP.SetDouble('GPT_BASETTCDEV',TLBAST.getDouble('BM5_MTSITUATION'));
        end;
        TOBPP.SetDouble('GPT_BASETTC',DEVISETOPIVOTEx(TOBPP.GetDouble('GPT_BASETTCDEV'),DEV.Taux,DEV.Quotite,V_PGI.OkDecP));
        TOBPP.SetString('GPT_FAMILLETAXE1',TOBBAST.GetString('BM4_FAMILLETAXE1'));
        CalculMontantsPiedPort(TOBpiece, TOBPP,TOBBases);
        result := True;
      end;
      ferme (QQ);
    FINALLY
      TP.free;
    END;
  end;

  function  AjouteLigneRetenue (TOBPiece,TOBPorcs,TLBAST,TParam : TOB ; var NbTrait : Integer) : boolean;
  var CodePort : string;
      QQ : TQuery;
      TP,TOBPP : TOB;
  begin
    result := false;
    TP := TOB.Create('PORT', nil, -1);
    TRY
      CodePort := TLBAST.GetString('BM5_ARTICLE');
      if CodePort = '' then Exit;
      QQ := OpenSQL('SELECT * FROM PORT WHERE GPO_CODEPORT="'+CodePort+'"',True,1,'',true);
      if not QQ.Eof then
      begin
        TP.SelectDB('', QQ);
        TOBPP := TOB.Create ('PIEDPORT',TOBPorcs,-1);
        TOBPP.SetString('GPT_CODEPORT', CodePort);
        ChargeTobPiedPort (TOBPP,TP,TOBPiece);
        TOBPP.SetString('GPT_LIBELLE',TLBAST.GetString('BM5_LIBELLE'));
        TOBPP.SetDouble('GPT_BASETTCDEV',TLBAST.getDouble('BM5_MTSITUATION'));
        TOBPP.SetDouble('GPT_BASETTC',DEVISETOPIVOTEx(TOBPP.GetDouble('GPT_BASETTCDEV'),DEV.Taux,DEV.Quotite,V_PGI.OkDecP));
        TOBPP.SetString('GPT_FAMILLETAXE1',TOBBAST.GetString('BM4_FAMILLETAXE1'));
        CalculMontantsPiedPort(TOBpiece, TOBPP,TOBBases);
        result := True;
      end;
      ferme (QQ);
    FINALLY
      TP.free;
    END;
  end;

  procedure  GetLesRGPrecByTVA (TOBBAST,TOBOLDRG : TOB);
  var QQ: TQuery;
      SQL : String;
      II : Integer;
      TOBDOLBAST,TRG : TOB;
      TheTva : string;
      MtRg : double;
  begin
    TRG := nil;
    TOBDOLBAST := TOB.Create ('LES BAST',nil,-1);
    TRY
      SQL := 'SELECT BM4_NUMSITUATION,BM4_FAMILLETAXE1,BM4_AFFAIRE,BM4_CODEMARCHE,BM4_NUMSITUATION,BM4_PAIEMENTPOC,'+
             '( '+
              'SELECT BM5_MTSITUATION FROM BASTLIG WHERE '+
              'BM5_AFFAIRE=BM4_AFFAIRE AND '+
              'BM5_CODEMARCHE=BM4_CODEMARCHE AND '+
              'BM5_NUMSITUATION=BM4_NUMSITUATION AND '+
              'BM5_PAIEMENTPOC="X" AND '+
              'BM5_TYPELBAST="001" AND '+
              'BM5_TYPELIG="TOT"'+
             ') AS MTRGTTC '+
             'FROM BASTENT '+
             'WHERE '+
             'BM4_AFFAIRE="'+TOBBAST.GetString('BM4_AFFAIRE')+'" AND '+
             'BM4_CODEMARCHE="'+TOBBAST.GetString('BM4_CODEMARCHE')+'" AND '+
             'BM4_PAIEMENTPOC="X" AND '+
             'BM4_NUMSITUATION < '+TOBBAST.GetString('BM4_NUMSITUATION')+' '+
             'ORDER BY BM4_FAMILLETAXE1';
      QQ := OpenSQl(SQL,true,-1,'',True);
      if not QQ.eof then
      begin
        TOBDOLBAST.LoadDetailDB('BASTENT','','',QQ,false);
      end;
      ferme(QQ);
      TheTva := '';
      if TOBDOLBAST.Detail.count > 0 then
      begin
        for II:= 0 TO TOBDOLBAST.Detail.count -1 do
        begin
          (*
          MtRG := 0;
          SQL := 'SELECT BM5_MTSITUATION FROM BASTLIG WHERE '+
              'BM5_AFFAIRE="'+ TOBDOLBAST.Detail[II].GetString('BM4_AFFAIRE')+'" AND '+
              'BM5_CODEMARCHE="'+ TOBDOLBAST.Detail[II].GetString('BM4_CODEMARCHE')+'" AND '+
              'BM5_NUMSITUATION="'+ TOBDOLBAST.Detail[II].GetString('BM4_NUMSITUATION')+'" AND '+
              'BM5_PAIEMENTPOC="X" AND '+
              'BM5_TYPELBAST="001" AND '+
              'BM5_TYPELIG="TOT"';
          QQ := OpenSQL(SQL,True,1,'',true);
          if not QQ.eof then
          begin
            MtRg := QQ.Fields[0].asfloat;
          end;
          Ferme(QQ);
          *)
          if (TheTva <> TOBDOLBAST.detail[II].GetString('BM4_FAMILLETAXE1')) then
          begin
            TheTva := TOBDOLBAST.detail[II].GetString('BM4_FAMILLETAXE1');
            TRG := TOB.Create ('UNE RG A UN TAUX',TOBOLDRG,-1);
            TRG.AddChampSupValeur('TVARG',TheTva);
            TRG.AddChampSupValeur('MONTANTRG',0);
          end;
          TRG.SetDouble('MONTANTRG',TRG.GetDouble('MONTANTRG')+TOBDOLBAST.detail[II].GetDouble('MTRGTTC'));
        end;
      end;
    FINALLY
      TOBDOLBAST.free;
    END;
  end;


  function  AjouteLigneRG (TOBPiece,TOBRG,TOBBAST,TLBAST,TParam : TOB ; var NbTrait : Integer) : boolean;
  var TT : TOB;
      MtRgCum,TauxTaxe,MtCautionCum,MtCautionMois,MtRgMois,MtCautionPrec,MtRgPrec : double;
      resteCum,restePrec,MtRgDue : Double;
  begin
    result := false;
    if TLBAST.GetString('BM5_CODE')='RET0000000001' then
    begin
      //
      TT := TOB.Create ('PIECERG',TOBRG,-1);
      initLigneRg (TT,TOBPiece);
      TT.AddChampSupValeur('CAUTIONTOT',0);
      TT.AddChampSupValeur('RGCUM',TLBAST.GEtValue('BM5_MTDEJAFACT')+TLBAST.getDouble('BM5_MTSITUATION'));
      TT.putValue('PRG_TAUXRG',TOBBAST.GEtValue('BM4_TAUXRG'));
      TT.putValue('PRG_TYPERG','TTC');
      TT.putValue('PRG_NUMLIGNE',0);
      TT.putValue('PRG_MTMANUEL','X');
      TT.putValue('INDICERG',0);
      TT.SetDouble('PRG_MTTTCRGDEV',TLBAST.getDouble('BM5_MTSITUATION'));
      TT.SetDouble('PRG_MTTTCRG',DEVISETOPIVOTEx(TT.GetDouble('PRG_MTTTCRGDEV'),DEV.Taux,DEV.Quotite,V_PGI.okdecV ));
      TauxTaxe := GetTauxTaxe(TOBPiece,'ACH',TOBPiece.GetString('GP_REGIMETAXE'),TOBBAST.GetString('BM4_FAMILLETAXE1'));
      TT.SetDouble('PRG_MTHTRGDEV',ARRONDI(TT.GetDouble('PRG_MTTTCRGDEV')/ (1+TauxTaxe/100),DEV.Decimale));
      TT.SetDouble('PRG_MTHTRG',DEVISETOPIVOTEx(TT.GetDouble('PRG_MTHTRGDEV'),DEV.Taux,DEV.Quotite,V_PGI.okdecV ));
      result := true;
    end else if TLBAST.GetString('BM5_CODE')='RET0000000002' then
    begin
      if TOBRG.detail.count > 0 then
      begin
        TT := TOBRG.detail[0];
        if TT <> nil then
        begin
          MtCautionCum := Arrondi(TLBAST.GetDouble('BM5_MTDEJAFACT')+TLBAST.getDouble('BM5_MTSITUATION'),DEV.Decimale);
          MtRgCum := TT.GetDouble('RGCUM');
          ResteCum :=  ARRONDI(MtRgCum - MtCautionCum,DEV.Decimale); if ResteCum < 0 then resteCum := 0;
          //
          MtCautionMois := TLBAST.getDouble('BM5_MTSITUATION');
          MtRgMois := TT.GetDouble('PRG_MTTTCRGDEV');
          //
          MtCautionPrec := ARRONDI(MtCautionCum - MtCautionMois,DEV.Decimale); if MtCautionPrec < 0 then MtCautionPrec := 0;
          MtRgPrec := ARRONDI(MtRgCum - MtRgMois,DEV.Decimale);
          RestePrec := ARRONDI( MtRgPrec - MtCautionPrec,DEV.decimale);
          //

          MtRgDue := Arrondi(ResteCum - restePrec,DEV.decimale);
          if MtRgDue > 0 then
          begin
            TT.putValue('PRG_MTMANUEL','X');
            TT.SetDouble('PRG_MTTTCRGDEV',MtRgDue);
            TT.SetDouble('PRG_MTTTCRG',DEVISETOPIVOTEx(TT.GetDouble('PRG_MTTTCRGDEV'),DEV.Taux,DEV.Quotite,V_PGI.okdecV ));
            TauxTaxe := GetTauxTaxe(TOBPiece,'ACH',TOBPiece.GetString('GP_REGIMETAXE'),TOBBAST.GetString('BM4_FAMILLETAXE1'));
            TT.SetDouble('PRG_MTHTRGDEV',ARRONDI(TT.GetDouble('PRG_MTTTCRGDEV')/ (1+TauxTaxe/100),DEV.Decimale));
            TT.SetDouble('PRG_MTHTRG',DEVISETOPIVOTEx(TT.GetDouble('PRG_MTHTRGDEV'),DEV.Taux,DEV.Quotite,V_PGI.okdecV ));
          end else
          begin
            TT.free;
          end;
        end;
      end;
      result := true;
    end;
  end;

var TT,TB,TOBPARAM,TP,TOBOLDRG : TOB;
    cledoc : R_CLEDOC;
    II,JJ,NbTrait : Integer;
    PhaseTravaux : string;
begin
  PhaseTravaux := FindPhasePoc(TOBBAST.GetString ('BM4_AFFAIRE'),TOBBAST.GetString ('BM4_FOURNISSEUR'),TOBBAST.GetString ('BM4_CODEMARCHE'));
  TOBPARAM := TOB.Create('LES PARAMS',nil,-1);
  fResult.ErrorResult := oeUnknown;
  fResult.LibError := 'Param�trage du BAST non correct';
  NbTrait := 0;
  //
  TOBPARAM.LoadDetailFromSQL('SELECT * FROM BTYPELIGBAST',false);
  //
  if TOBBAST.getBoolean('BM4_RGLIBER') then
  begin
    cledoc.NaturePiece := 'B01';
  end else
  begin
    cledoc.NaturePiece := 'FF';
  end;
  TT := CreerTOBPieceVide (cledoc,TOBBAST.GetString ('BM4_FOURNISSEUR'),TOBBAST.GetString ('BM4_AFFAIRE'),VH^.EtablisDefaut,'',True,False,1);
  TRY
    TOBPiece.Dupliquer(TT,false,true);
    TOBPiece.SetDouble('GP_ESCOMPTE',0); // afin d'�viter de r�cup�rer le taux d'escompte de la fiche fournisseur
    TOBPiece.SetDateTime('GP_DATEPIECE',DateFac);
    TOBPiece.SetDateTime('GP_DATECREATION',V_PGI.DateEntree);
    TOBPiece.SetString('GP_REFEXTERNE',TOBBAST.GetString('BM3_NUMERODOC'));
    // Gestion pour la reprise d'ant�riorit�
    TOBPiece.SetString('GP_HORSCOMPTA',TOBBAST.GetString('BM4_HORSCOMPTA'));
    // ------
    TOBPiece.SetString('GP_CREEPAR','BST');
    TOBPiece.SetBoolean('GP_RESTITUERG',TOBBAST.getBoolean('BM4_RGLIBER'));
    //
    if (Tools.StringInList(TOBBAST.GetString('BM4_FAMILLETAXE1'),VH_GC.AutoLiquiTVAST)) then
    begin
      TOBPiece.SetBoolean('GP_AUTOLIQUID',true);
    end;
    //
    if TOBBAST.getBoolean('BM4_RGLIBER') then TOBPiece.SetBoolean('GP_AUTOLIQUID',true);;
    //
    for II := 0 to TOBBAST.Detail.count -1 do
    begin
      TB := TOBBASt.detail[II];
      if (TB.GetString('BM5_CODE')='') and (TB.GetString('BM5_ARTICLE')='') then Continue;
      if (TB.GetDouble('BM5_MTSITUATION')=0) then continue;
      if TB.GetString('BM5_CODE')<>'' then
      begin
        TP := TOBPARAM.FindFirst(['BM6_CATLIGNE','BM6_CODE'],[TB.GetString('BM5_TYPELBAST'),TB.GetString('BM5_CODE')],false);
        if TP = nil then continue;
        if (TP.GetString('BM6_ARTICLE')= '') and (Copy(TB.GetString('BM5_CODE'),1,3)<>'RET') then break;
        //
        if TP.GetString('BM6_TYPEERP')='ART' then
        begin
          if TP.GetString('BM6_CODE')='RET0000000003' then
          begin
            // Restitution des RGS
            TOBOLDRG := TOB.Create ('LES OLD BAST',nil,-1);
            TRY
              GetLesRGPrecByTVA (TOBBAST,TOBOLDRG);
              if TOBOLDRG.Detail.count > 0 then
              begin
                for JJ := 0 to TOBOLDRG.detail.count -1 do
                begin
                  AjouteLigneRestitutionRg (TOBPiece,TB,TP,TOBOLDRG.detail[JJ],NbTrait);
                end;
              end;
            FINALLY
              TOBOLDRG.free;
            END;
          end else
          begin
            if not AjouteLigneArticle (TOBPiece,TB,TP,NbTrait,PhaseTravaux) then break;
          end;
        end else if (Pos (TP.GetString('BM6_TYPEERP'),'RG;CAU')>0) then
        begin
            if not AjouteLigneRG (TOBPiece,TOBPieceRG,TOBBAST,TB,TP,NbTrait) then break;
        end else
        begin
          if not AjouteLignePort (TOBPiece,TOBPorcs,TB,TP,NbTrait) then break;
        end;
      end else if TB.GetString('BM5_ARTICLE')<>'' then  // cas de l'ajout de retenue non param�tr�e 
      begin
        if not AjouteLigneRetenue (TOBPiece,TOBPorcs,TB,TP,NbTrait) then break;
      end;
    end;
    fResult.ErrorResult := oeOk;
    fResult.LibError := '';
  FINALLY
    TOBPARAM.free;
    TT.Free;
  END;
end;

{ Tresult }

constructor Tresult.create;
begin
  init;
end;

procedure Tresult.init;
begin
  NatureDoc :='';
  NumeroDoc :=0;
  ErrorResult := OeOK;
  LibError :='';
end;

end.
