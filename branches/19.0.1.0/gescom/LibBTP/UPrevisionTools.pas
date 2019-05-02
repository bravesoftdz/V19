unit UPrevisionTools;

interface
uses
  Classes,
  SysUtils,
  uTob,
  hEnt1,
  hCtrls,
  EntGC,
  SaisUtil,
  wCommuns,
  uEntCommun,
  UtilConso,
  forms,
  Menus,
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  affaireutil,Paramsoc
  ;
type
  TTypeData = (TTdNormal,TTdFrais);
  
  TGenerePrev = class (TObject)
  private
    DEV : Rdevise;
    TOBA : TOB;
    fLivChantier : integer;
    ListPhase : Tlist;
    procedure ChargeReplPourcent;
    function DefiniLaBranche(TOBL,TOBDEPART: TOB): TOB;
    function InsereParag (Libelle,TypeD,Famille: STring; Indice,Niveau : Integer) : TOB;
    function AjouteLaFamille(LeType, LaFamille: string; TT: TOB;IsPrestation: boolean): TOB;
    function AjouteLeType(LeType: string; TOBDEPART : TOB): TOB;
    function findLaFamille(LeTYpe, LaFamille: string; TOBDEBUT: TOB): TOB;
    function findLeType(LeTYpe: string; TOBDEPART : TOB): TOB;
    procedure RemplacePourcent(TOBL, TOBPiece: TOB);
    function IntegreArticleGlobal(TOBPiece,TOBL,TOBDEBUT: TOB; TypeData : TTypeData ): TOB;
  public
    NumOrdre : Integer;
    TOBPiece : TOB;
    TOBOuvrage : TOB;
    TobFrais : TOB;
    TOBOuvrageFrais : TOB;
    TOBSST : TOB;
    TOBgenere : TOB;
    TOBPLAT : TOB;
    constructor create;
    destructor destroy ; override;
    function ConstituePrevision : boolean;
  end;

  TPrevisionTools = class
  private
    class function ChargeLesTobs(Cledoc : R_CLEDOC;TobPiece,TOBOuvrage,TobFrais,TOBOuvrageFrais,TOBSST : TOB) : boolean;
    class function ConstitueDocSortie(TOBGenere,TobPiece,TOBOuvrage,TobFrais,TOBOuvrageFrais,TOBSST : TOB) : boolean;
    class function EnregistrePrevision(TOBPiece,TOBGenere :TOB) : boolean;
    class function MajAffaireDevis (TOBpiece : TOB) : boolean;
  public
    class function GenerePrevisionChantier (NaturePiece,Souche: string; Numero,Indice : integer) : boolean;
    class procedure ReajusteLigneDevisPourCtrEtude(TOBL: TOB;CoefFr: double);
  end;

implementation
uses UtilTOBPiece
    , BTPUtil
    , UCotraitance
    , FactureBtp
    , FactTOB
    , FactUtil
    , FactComm
    , FactVariante
    , AGLInitBtp
    , FactOuvrage
    , BTStructChampSup
    , FactCalc
    , UPlannifchUtil
    , FactGrp
    ;

{ PrevisionTools }
class procedure TPrevisionTools.ReajusteLigneDevisPourCtrEtude(TOBL : TOB; CoefFr : double);
var MontantAch,MtFG,MtFR,MTFC,QTe : double;
begin
  MtFR := 0;
  Qte := TOBL.GetValue('GL_QTEFACT')/ TOBL.GetValue('GL_PRIXPOURQTE');
  MontantAch := TOBL.getValue('GL_MONTANTPA');
  MtFG := TOBL.GetValue('GL_MONTANTFG');
  MtFC := 0;
  if TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X' then
  begin
    MtFR := Arrondi((MontantAch+MTFG+MTFC) * Coeffr,4);
  end;
  TOBL.putValue('GL_COEFFC',0);
  TOBL.putValue('GL_COEFFR',COEFFR);
  TOBL.putValue('GL_MONTANTPAFR',MontantAch);
  TOBL.putValue('GL_MONTANTPAFG',MontantAch);
  TOBL.putValue('GL_MONTANTPAFC',0);
  TOBL.putValue('GL_MONTANTFG',MTFG);
  TOBL.putValue('GL_MONTANTFR',MTFR);
  TOBL.putValue('GL_MONTANTFC',MTFC);
  TOBL.putValue('GL_MONTANTPR',MontantAch + MtFG+ MTFC+MtFR);
  if QTe <> 0 then TOBL.putValue('GL_DPR',Arrondi(TOBL.GetValue('GL_MONTANTPR')/Qte,V_PGI.okdecP));
  if TOBL.GetValue('GL_MONTANTPR')<>0 then TOBL.putValue('GL_COEFMARG',Arrondi(TOBL.GetValue('GL_MONTANTHT')/TOBL.GetValue('GL_MONTANTPR'),4));
end;

class function TPrevisionTools.ChargeLesTobs(Cledoc: R_CLEDOC; TobPiece,TOBOuvrage, TobFrais, TOBOuvrageFrais, TOBSST: TOB): boolean;
var Q : Tquery;
    SQL : String;
    cledocAff : r_cledoc;
    ModeSaisie : TActionFiche;
begin
  Result := True;
  TRY
    ModeSaisie := taModif;
    Q:=OpenSQL ('SELECT * FROM PIECE WHERE '+ WherePiece(Cledoc,ttdPiece,False),True,-1,'',true) ;
    TobPiece.selectDB ('',Q);
    Ferme(Q) ;
    LoadLesSousTraitants(cledoc ,TOBSST);
    Sql := MakeSelectLigneBtp (true,false,false);
    Sql := Sql + ' WHERE ' + WherePiece(cledoc, ttdLigne,false) + ' ORDER BY GL_NUMLIGNE';
    Q:=OpenSQL (SQL,True,-1,'',true) ;
    TOBPiece.LoadDetailDB ('LIGNE','','',Q,True,true) ;
    Ferme(Q) ;
    if TOBPiece.detail.count > 0 then
    begin
      RetrouvePieceFraisBtp (cledocAff,TOBpiece,ModeSaisie);
      if (cledocAff.NumeroPiece <> 0) then
      begin
        Q:=OpenSQL ('SELECT * FROM PIECE WHERE '+WherePiece(cledocAff,ttdPiece,False),True,-1,'',true) ;
        TobFrais.selectDB ('',Q);
        Ferme(Q) ;
        Sql := MakeSelectLigneBtp (true,false,false);
        Sql := Sql + ' WHERE ' + WherePiece(cledocAff, ttdLigne,false) + ' ORDER BY GL_NUMLIGNE';
        Q:=OpenSQL (SQL,True,-1,'',true) ;
        TOBFrais.LoadDetailDB ('LIGNE','','',Q,True,true) ;
        Ferme(Q) ;
        ChargeLesOuvrages (TOBFrais,TOBOuvrageFrais,cledocAff);
      end;
    end;
    // --
    ChargeLesOuvrages (TOBPIece,TOBOuvrage,Cledoc);
  EXCEPT
    Result := false;
  end;
end;

class function TPrevisionTools.ConstitueDocSortie(TOBGenere, TobPiece,TOBOuvrage, TobFrais, TOBOuvrageFrais, TOBSST: TOB): boolean;
var TT : TGenerePrev;
begin
  Result := false;
  TT := TGenerePrev.create;
  TT.TOBPiece := TOBPiece;
  TT.TOBOuvrage := TOBOuvrage;
  TT.TobFrais := TobFrais;
  TT.TOBOuvrageFrais := TOBOuvrageFrais;
  TT.TOBgenere := TOBGenere;
  try
    if TT.ConstituePrevision then Result := true;
  finally
    TT.Free;
  end;
end;

class function TPrevisionTools.EnregistrePrevision(TOBPiece,TOBGenere : TOB): boolean;
var NewPiece : Boolean;
    TOBChantier,TOBresult : TOB;
begin
  Result := false;
  TOBResult := TOB.Create ('PIECE',nil,-1);
  TOBChantier := TOB.Create ('PIECE',nil,-1);
  TRY
    newPiece := ExistsChantier (TOBGenere,TOBChantier);
    if NewPiece then
    begin
      if not CreerPiecesFromLignes(TobGenere,'DEVTOCHAN',iDate1900,true,True,TOBResult) then exit;
    end else
    begin
      if not AjouteLignesToPiece (TOBGenere,TOBChantier,'DEVTOCHAN',True,true,TOBResult) then exit;
    end;
    Result := True;
  FINALLY
    TOBChantier.Free;
    TOBresult.free;
  end;

end;

class function TPrevisionTools.GenerePrevisionChantier(NaturePiece,Souche: string; Numero, Indice: integer): boolean;
var TobPiece,TOBOuvrage,TOBFrais,TOBOuvrageFrais : TOB;
    TOBBasesL,TOBSST,TOBGenere:TOB;
    cledoc : R_CLEDOC;
begin
  Result := false;
  FillChar(cledoc,SizeOf(cledoc),#0);
  cledoc.NaturePiece := NaturePiece;
  cledoc.Souche := Souche;
  cledoc.NumeroPiece := Numero;
  cledoc.Indice := Indice;
  // génération de la prévision de chantier globale
  TobPiece := TOB.Create ('PIECE',NIL, -1);
  TOBOuvrage := TOB.Create ('LIGNEOUV',nil,-1);
  TobFrais := TOB.Create ('PIECE',NIL, -1);
  TOBOuvrageFrais := TOB.Create ('LIGNEOUV',nil,-1);
  TOBBasesL := TOB.create ('LES LIGNEBASE',nil,-1);
  TOBSST := TOB.Create ('LES SOUS-TRAITS',nil,-1);
  TOBGenere := TOB.Create ('PIECE',nil,-1);
  try
    if not ChargeLesTobs(Cledoc,TobPiece,TOBOuvrage,TobFrais,TOBOuvrageFrais,TOBSST) then Exit;
    if not ConstitueDocSortie(TOBGenere,TobPiece,TOBOuvrage,TobFrais,TOBOuvrageFrais,TOBSST) then Exit;
    if not EnregistrePrevision(TobPiece,TOBGenere) then Exit;
    if not MajAffaireDevis (TOBPiece) then exit;
    Result := True;
  finally
    TobGenere.free;
    TobPiece.free;
    TOBOuvrage.free;
    TobFrais.free;
    TOBOuvrageFrais.free;
    TOBBasesL.free;
    TOBSST.free;
  end;
end;

{ TGenerePrev }

function TGenerePrev.InsereParag (Libelle,TypeD,Famille: STring; Indice,Niveau : Integer) : TOB;
var TOBL : TOB;
begin
  inc(NumOrdre);
  if Indice <> -1 then Inc(Indice);
  TOBL:=NewTobLigne(TOBGenere,Indice);
  InitialiseLigne (TOBL,TOBL.GetIndex,0);
  PieceVersLigne (TOBPiece,TOBL);
  TOBL.SetDateTime('GL_DATELIVRAISON',IDate2099);
  TOBL.SetString('ZZTRI0',TypeD);
  TOBL.SetString('ZZTRI1',Famille);
  TOBL.SetString('GL_TYPELIGNE','DP'+IntToStr(Niveau));
  TOBL.SetInteger('GL_NIVEAUIMBRIC',Niveau) ;
  TOBL.SetInteger('GL_NUMLIGNE',0);
  TOBL.SetInteger('GL_NUMORDRE',NumOrdre);
  TOBL.SetString('GL_LIBELLE',Libelle);
  TOBL.SetString('GL_TYPEDIM','NOR') ;
  TOBL.SetInteger('GL_IDENTIFIANTWOL',fLivChantier);
  TOBL.SetInteger('IDENTIFIANTWOL',fLivChantier);
  TOBL.AddChampSupValeur('GP_REFINTERNE',TOBPiece.GetValue('GP_REFINTERNE'));
  TOBL.AddChampSupValeur('GP_REPRESENTANT',TOBPiece.GetValue('GP_REPRESENTANT'));
  TOBL.AddChampSupValeur ('PIECEORIGINE',EncodeRefPiece (TOBPiece,0,false));
  TOBL.AddChampSupValeur ('GP_NUMADRESSEFACT',TOBPiece.GetValue('GP_NUMADRESSEFACT'));
  TOBL.AddChampSupValeur ('GP_NUMADRESSELIVR',TOBPiece.GetValue('GP_NUMADRESSELIVR'));
  result := TOBL;
  //
  inc(NumOrdre);
  if Indice <> -1 then inc(Indice);
  TOBL:=NewTobLigne(TOBGenere,Indice);
  InitialiseLigne (TOBL,TOBL.GetIndex,0);
  PieceVersLigne (TOBPiece,TOBL);
  TOBL.SetDateTime('GL_DATELIVRAISON',IDate2099);
  TOBL.SetString('ZZTRI0','');
  TOBL.SetString('ZZTRI1','');
  TOBL.SetString('GL_TYPELIGNE','TP'+InttoStr(Niveau));
  TOBL.SetInteger('GL_NIVEAUIMBRIC',Niveau) ;
  TOBL.SetInteger('GL_NUMLIGNE',0);
  TOBL.SetInteger('GL_NUMORDRE',NumOrdre);
  TOBL.SetString('GL_LIBELLE','Total '+Libelle);
  TOBL.SetString('GL_TYPEDIM','NOR') ;
  TOBL.SetInteger('GL_IDENTIFIANTWOL',fLivChantier);
  TOBL.SetInteger('IDENTIFIANTWOL',fLivChantier);
  TOBL.AddChampSupValeur('GP_REFINTERNE',TOBPiece.GetValue('GP_REFINTERNE'));
  TOBL.AddChampSupValeur('GP_REPRESENTANT',TOBPiece.GetValue('GP_REPRESENTANT'));
  TOBL.AddChampSupValeur ('PIECEORIGINE',EncodeRefPiece (TOBPiece,0,false));
  TOBL.AddChampSupValeur ('GP_NUMADRESSEFACT',TOBPiece.GetValue('GP_NUMADRESSEFACT'));
  TOBL.AddChampSupValeur ('GP_NUMADRESSELIVR',TOBPiece.GetValue('GP_NUMADRESSELIVR'));
end;

function TGenerePrev.ConstituePrevision: boolean;
var II,Indice1 : integer;
    TOBL,LaPhase,TD : TOB;
    CodeArticle : string;
begin
  Result := false;
  ChargeReplPourcent;
  TD := nil;
  //
  if (TOBPiece = nil) or (TOBOuvrage = nil) then Exit;
  DEV.Code:=TOBPiece.GetValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
  // Traitement des lignes du devis
  TOBGenere.ClearDetail;
  for II := 0 to TobPiece.detail.Count -1 do
  begin
    if II = 0 then
    begin
      TD := InsereParag ('Détail des Travaux','','',-1,1);
    end;
    TOBL := TOBPiece.detail[II];
    if IsVariante (TOBL) then continue;
    //
    if (TOBL.GetString('GLC_NATURETRAVAIL')='001') then continue;
    // --
    if ((TOBL.GetValue('GL_TYPEARTICLE')='OUV') OR (TOBL.GetValue('GL_TYPEARTICLE')='ARP')) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
    begin
      TRY
        TOBPLAT.ClearDetail;
        MiseAplatouv (TOBPiece,TOBL,TOBOuvrage,TOBPLat,true,FALSE,true,true);
        for indice1 := 0 to TOBPlat.detail.count -1 do
        begin
          // VARIANTE
          if IsVariante (TOBPLat.detail[Indice1]) then continue;
          if TOBPLat.detail[Indice1].GetValue('GL_QTEFACT') = 0 then continue;
          //gestion de la non-prise en compte des lignes Cotraitantes
          if (TOBPLat.detail[Indice1].GetString('GLC_NATURETRAVAIL')='001') then continue;
          if (TOBPLat.detail[Indice1].GetValue('GL_TYPEARTICLE')='PRE') then
          begin
            TOBPLat.detail[Indice1].PutValue('BNP_TYPERESSOURCE',RenvoieTypeRes(TOBPLat.detail[Indice1].GetValue('GL_ARTICLE')));
          end;
          LaPhase := DefiniLaBranche (TOBPlat.detail[Indice1],TD);
          if TOBL.GEtValue('GLC_DOCUMENTLIE')<> '' then
          begin
            TOBPLat.detail[Indice1].PutValue('GLC_DOCUMENTLIE',TOBL.GEtValue('GLC_DOCUMENTLIE'));
          end;
          // -- Recalcul des elements de la ligne en enlevant les frais de chantier
          TPrevisionTools.ReajusteLigneDevisPourCtrEtude(TOBPLat.detail[Indice1],TOBPiece.getValue('GP_COEFFR'));
          // --
          CodeArticle := TOBPlat.detail[Indice1].GetValue('GL_CODEARTICLE');
          if (TOBPLat.detail[Indice1].getvalue('GL_TYPEARTICLE')='POU') then RemplacePourcent(TOBPlat.detail[Indice1],TOBPiece);
          IntegreArticleGlobal (TOBPiece,TOBPlat.detail[Indice1], LaPhase,TTdNormal);
        end;
      FINALLY
        TOBPlat.clearDetail;
      END;
    end else if (TOBL.GetValue('GL_TYPEARTICLE')='MAR') OR
                (TOBL.GetValue('GL_TYPEARTICLE')='PRE') or
                (TOBL.GetValue('GL_TYPEARTICLE')='FRA') or
                ((TOBL.GetValue('GL_TYPEARTICLE')='ARP') and (TOBL.GetValue('GL_INDICENOMEN')=0)) then
    BEGIN
      if TOBL.GetValue('GL_QTEFACT') = 0 then continue;
      TPrevisionTools.ReajusteLigneDevisPourCtrEtude(TOBL,TOBPiece.getValue('GP_COEFFR'));
      LaPhase := DefiniLaBranche (TOBL,TD);
      IntegreArticleGlobal (TOBPiece,TOBL,LaPhase,TTdNormal);
    end else if (TOBL.GetValue('GL_TYPEARTICLE')='POU') then
    begin
      // remplacement de l'article de type pourcentage en article std marchandise
      RemplacePourcent(TOBL,TOBpiece);
      TPrevisionTools.ReajusteLigneDevisPourCtrEtude(TOBL,TOBPiece.getValue('GP_COEFFR'));
      LaPhase := DefiniLaBranche (TOBL,TD);
      IntegreArticleGlobal (TOBPiece,TOBL,LaPhase,TTdNormal);
    end;

  end;
  // traitement des frais de chantier
  for II := 0 to TOBFrais.detail.Count -1 do
  begin
    if II = 0 then
    begin
      TD := InsereParag ('Frais de chantier','','',-1,1);
    end;
    TOBL := TOBFrais.detail[II];
    if IsVariante (TOBL) then continue;
    //
    if (TOBL.GetString('GLC_NATURETRAVAIL')='001') then continue;
    // --
    if ((TOBL.GetValue('GL_TYPEARTICLE')='OUV') OR (TOBL.GetValue('GL_TYPEARTICLE')='ARP')) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
    begin
      TRY
        TOBPLAT.ClearDetail;
        MiseAplatouv (TOBFrais,TOBL,TOBOuvrageFrais,TOBPLat,true,FALSE,true,true);
        for indice1 := 0 to TOBPlat.detail.count -1 do
        begin
          // VARIANTE
          if IsVariante (TOBPLat.detail[Indice1]) then continue;
          if TOBPLat.detail[Indice1].GetValue('GL_QTEFACT') = 0 then continue;
          //gestion de la non-prise en compte des lignes Cotraitantes
          if (TOBPLat.detail[Indice1].GetString('GLC_NATURETRAVAIL')='001') then continue;
          if (TOBPLat.detail[Indice1].GetValue('GL_TYPEARTICLE')='PRE') then
          begin
            TOBPLat.detail[Indice1].PutValue('BNP_TYPERESSOURCE',RenvoieTypeRes(TOBPLat.detail[Indice1].GetValue('GL_ARTICLE')));
          end;
          LaPhase := DefiniLaBranche (TOBPlat.detail[Indice1],TD);
          if TOBL.GEtValue('GLC_DOCUMENTLIE')<> '' then
          begin
            TOBPLat.detail[Indice1].PutValue('GLC_DOCUMENTLIE',TOBL.GEtValue('GLC_DOCUMENTLIE'));
          end;
          // -- Recalcul des elements de la ligne en enlevant les frais de chantier
          TPrevisionTools.ReajusteLigneDevisPourCtrEtude(TOBPLat.detail[Indice1],TOBPiece.getValue('GP_COEFFR'));
          // --
          CodeArticle := TOBPlat.detail[Indice1].GetValue('GL_CODEARTICLE');
          if (TOBPLat.detail[Indice1].getvalue('GL_TYPEARTICLE')='POU') then RemplacePourcent(TOBPlat.detail[Indice1],TOBPiece);
          IntegreArticleGlobal (TOBFrais,TOBPlat.detail[Indice1], LaPhase, TTdFrais);
        end;
      FINALLY
        TOBPlat.clearDetail;
      END;
    end else if (TOBL.GetValue('GL_TYPEARTICLE')='MAR') OR
                (TOBL.GetValue('GL_TYPEARTICLE')='PRE') or
                (TOBL.GetValue('GL_TYPEARTICLE')='FRA') or
                ((TOBL.GetValue('GL_TYPEARTICLE')='ARP') and (TOBL.GetValue('GL_INDICENOMEN')=0)) then
    BEGIN
      if TOBL.GetValue('GL_QTEFACT') = 0 then continue;
      TPrevisionTools.ReajusteLigneDevisPourCtrEtude(TOBL,TOBPiece.getValue('GP_COEFFR'));
      LaPhase := DefiniLaBranche (TOBL,TD);
      IntegreArticleGlobal (TobFrais,TOBL,LaPhase,TTdFrais);
    end else if (TOBL.GetValue('GL_TYPEARTICLE')='POU') then
    begin
      // remplacement de l'article de type pourcentage en article std marchandise
      RemplacePourcent(TOBL,TOBpiece);
      TPrevisionTools.ReajusteLigneDevisPourCtrEtude(TOBL,TOBPiece.getValue('GP_COEFFR'));
      LaPhase := DefiniLaBranche (TOBL,TD);
      IntegreArticleGlobal (TobFrais,TOBL,LaPhase,TTdFrais);
    end;

  end;

  result := true;
end;

constructor TGenerePrev.create;
begin
  TOBA := TOB.Create ('ARTICLE',nil,-1);
  TOBPLAT := TOB.Create ('LES OUV PLATS',nil,-1);
  ListPhase := TList.Create;
  MemoriseChampsSupLigneETL ('PBT' ,True);
  MemoriseChampsSupLigneOUV ('PBT');
  MemoriseChampsSupPIECETRAIT;
  fLivChantier := 0; // livraison dépot
  if GetParamSoc ('SO_BTLIVCHANTIER') then
  begin
    fLivChantier := -1;
  end;
end;

destructor TGenerePrev.destroy;
begin
  TOBA.free;
  TOBPLAT.Free;
  ListPhase.free;
  InitStructureETL;
  inherited;
end;

function TGenerePrev.AjouteLaFamille (LeType,LaFamille : string;TT : TOB; IsPrestation : boolean) : TOB;
var II,Indice : Integer;
    TOBXX: TOB;
    LibeLle : string;
begin
  Indice := -1;
  for II := TT.GetIndex+1 to TOBgenere.detail.Count -1 do
  begin
    TOBXX := TOBgenere.detail[II];
    if (TOBXX.GetString('ZZTRI1') <> LeType) or
       (TOBXX.GetString('ZZTRI2') > LaFamille) then
    begin
      Indice := II-1;
    end;
  end;
  if IsPrestation then
  begin
    LibeLle := RechDom('BTNATPRESTATION',LaFamille,false);
  end else
  begin
    if LaFamille <> '@@@' then LibeLle := RechDom('GCFAMILLENIV1',LaFamille,false)
                          else Libelle := 'Non défini';
  end;
  result := InsereParag (Libelle,LeType,LaFamille, Indice,3);
end;

function TGenerePrev.AjouteLeType (LeType : string; TOBDEPART : TOB) : TOB;

  function GetLibelleType(LeType : string) : string;
  begin
    if LeType = '000' then
    begin
      Result := 'Matériaux';
    end else if LeType = '001' then
    begin
      Result := 'Main d''oeuvre';
    end else if LeType = '002' then
    begin
      Result := 'Interim.';
    end else if LeType = '003' then
    begin
      Result := 'Sous traitance';
    end else if LeType = '004' then
    begin
      Result := 'Location';
    end else if LeType = '005' then
    begin
      Result := 'Matériels';
    end else if LeType = '006' then
    begin
      Result := 'Outillage';
    end else if LeType = '007' then
    begin
      Result := 'Autre';
    end;
  end;

var II,Indice : Integer;
    TOBXX : TOB;
begin
  Indice := TOBGenere.detail.Count -1;
  for II := TOBDEPART.GetIndex+1 to TOBGenere.detail.Count -1 do
  begin
    TOBXX := TOBGenere.detail[II];
    if TOBXX.GetString('ZZTRI0') > LeType then
    begin
      Indice := II;
      break;
    end;
  end;
  result := InsereParag (GetLibelleType(LeType),LeType,'',Indice,2);
end;

procedure TGenerePrev.ChargeReplPourcent;
var QQ : TQuery;
    TheArticle : string;
begin
  TheArticle := GetParamSoc ('SO_BTREPLPOURCENT');

  TOBA := TOB.Create ('ARTICLE',nil,-1);
  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_CODEARTICLE="'+TheArticle+'"',true,-1,'',true);
  TOBA.selectDb ('',QQ);
  ferme(QQ);
end;

function TGenerePrev.DefiniLaBranche (TOBL,TOBDEPART : TOB) : TOB;

  function GetNaturePrest (CodeArticle : string) : string;
  var QQ : TQuery;
  begin
    Result := 'Inconnu...';
    QQ := OpenSQL('SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE="'+CodeArticle+'"',True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.Fields[0].AsString;
    end;
    Ferme(QQ);
  end;

var TT,TF : TOB;
    LeType,LaFamille : string;
    IsPrestation : Boolean;
begin
  LeType := TOBL.GetString('ZZTRI0');
  IsPrestation := (TOBL.getSTring('GL_TYPEARTICLE')='PRE');
  if IsPrestation then
  begin
    LaFamille := GetNaturePrest(TOBL.GetString('GL_ARTICLE'));
  end else
  begin
    LaFamille := TOBL.GetString('GL_FAMILLENIV1');
    if LaFamille='' then LaFamille:='@@@';
  end;
  TT := findLeType(LeType,TOBDEPART);
  if TT = nil then
  begin
    TT := AjouteLeType (LeType,TOBDEPART);
    TF := AjouteLaFamille (LeType,LaFamille,TT, Isprestation);
  end else
  begin
    TF := FindLaFamille (LeType,LaFamille,TT);
    if TF = nil then
    begin
      TF := AjouteLaFamille (LeType,LaFamille,TT,Isprestation);
    end;
  end;
  Result := TF;
end;

function TGenerePrev.findLaFamille(LeTYpe,LaFamille : string; TOBDEBUT : TOB) : TOB;
var II : Integer;
    TOBD : TOB;
begin
  Result := nil;
  for II := TOBDEBUT.GetIndex+1 to TOBgenere.detail.Count -1 do
  begin
    TOBD := TOBgenere.detail[II];
    if TOBD.GetInteger('GL_NIVEAUIMBRIC') < 3 then break;
    if TOBD.GetInteger('GL_NIVEAUIMBRIC') > 3 then Continue;
    if TOBD.GetString('ZZTRI1') = LaFamille then
    begin
      Result := TOBD;
      Break;
    end;
  end;
end;

function TGenerePrev.findLeType(LeType : string; TOBDEPART : TOB) : TOB;
var II : Integer;
    TOBT : TOB;
begin
  result:= nil;
  if TOBGenere.detail.Count <= 2 then exit;
  for II := TOBDEPART.GetIndex+1 to TOBgenere.Detail.count -1 do
  begin
    TOBT := TOBgenere.detail[II];
    if TOBT.GetInteger('GL_NIVEAUIMBRIC') > 2 then Continue;
    if (TOBT.GetString('ZZTRI0') = LeType)  then
    begin
      result:= TOBT;
      Break;
    end;
  end;
end;

function TGenerePrev.IntegreArticleGlobal (TOBPiece,TOBL,TOBDEBUT: TOB; TypeData : TTypeData): TOB;

  function GetNaturePrest (CodeArticle : string) : string;
  var QQ : TQuery;
  begin
    Result := 'Inconnu...';
    QQ := OpenSQL('SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE="'+CodeArticle+'"',True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.Fields[0].AsString;
    end;
    Ferme(QQ);
  end;

  function AddLigne ( TOBPiece,TOBDEBUT,TOBL : TOB; Indice : Integer; LeType,LaFamille : string) : TOB;
  var TOBS : TOB;
      TypeArt : string;
      Fournisseur : string;
  begin
    inc(Numordre);
    if TOBL.GetString('GLC_NATURETRAVAIl')='002' then
    begin
      Fournisseur := TOBL.GetString('GL_FOURNISSEUR');
    end else
    begin
      Fournisseur := '';
    end;
    TypeArt := TOBL.GetValue('GL_TYPEARTICLE');
    TOBS:=NewTobLigne(TOBGenere,Indice);
    InitialiseLigne (TOBS,TOBS.GetIndex,0);
    PieceVersLigne (TOBPiece,TOBS);
    TOBS.Dupliquer(TOBL,false,true);
    //
    TOBS.SetDouble('GL_COEFFC',0);
    TOBS.SetDouble('GL_COEFFG',0);
    TOBS.SetDouble('GL_COEFMARG',0);
    //
    TOBS.SetInteger('GL_NUMLIGNE',0);
    TOBS.SetInteger('GL_NUMORDRE',NumOrdre);
    TOBS.SetInteger('GL_NIVEAUIMBRIC',3);
    TOBS.PutValue('IDENTIFIANTWOL',fLivChantier);
    TOBS.PutValue('GL_IDENTIFIANTWOL',fLivChantier);
    TOBS.SetString('ZZTRI0',LeType);
    TOBS.SetString('ZZTRI1',LaFamille);
    TOBS.AddChampSupValeur ('PIECEORIGINE',EncodeRefPiece (TOBPiece,0,false));
    TOBS.AddChampSupValeur ('GP_NUMADRESSEFACT',TOBPiece.GetValue('GP_NUMADRESSEFACT'));
    TOBS.AddChampSupValeur ('GP_NUMADRESSELIVR',TOBPiece.GetValue('GP_NUMADRESSELIVR'));
    TOBS.AddChampSupValeur('GP_REPRESENTANT',TOBPiece.GetValue('GP_REPRESENTANT'));
    //
    TOBS.AddChampSupValeur ('GP_REFINTERNE',TOBPiece.GetValue('GP_REFINTERNE'));
    TOBS.PutValue('GL_INDICENOMEN',0);
    if TypeData = ttdFrais then TOBS.SetString('GL_TYPELIGNE','ARV');
    result := TOBS;
  end;

  procedure CumuleLigne (TOBI,TOBL : TOB);
  begin
    TOBI.PutValue ('GL_QTEFACT',TOBL.GetValue('GL_QTEFACT')+TOBI.GetValue('GL_QTEFACT'));
    TOBI.PutValue ('GL_MONTANTPA',TOBL.GetValue('GL_MONTANTPA')+TOBI.GetValue('GL_MONTANTPA'));
    TOBI.PutValue ('GL_MONTANTPR',TOBL.GetValue('GL_MONTANTPR')+TOBI.GetValue('GL_MONTANTPR'));
    TOBI.PutValue ('GL_MONTANTPV',TOBL.GetValue('GL_MONTANTPR')+TOBI.GetValue('GL_MONTANTPR'));
    TOBI.PutValue ('GL_MONTANTHTDEV',TOBL.GetValue('GL_MONTANTHTDEV')+TOBI.GetValue('GL_MONTANTHTDEV'));
    TOBI.PutValue ('GL_MONTANTHT',TOBL.GetValue('GL_MONTANTHT')+TOBI.GetValue('GL_MONTANTHT'));
    TOBI.PutValue ('GL_TOTALHTDEV',TOBL.GetValue('GL_TOTALHTDEV')+TOBI.GetValue('GL_TOTALHTDEV'));
    TOBI.PutValue ('GL_TOTALHT',TOBL.GetValue('GL_TOTALHT')+TOBI.GetValue('GL_TOTALHT'));
    //
    TOBI.SetDouble('GL_COEFFC',0);
    TOBI.SetDouble('GL_COEFFG',0);
    TOBI.SetDouble('GL_COEFMARG',0);
    //
    if TOBI.GetValue('GL_QTEFACT') <> 0 then TOBI.PutValue ('GL_DPA',Arrondi(TOBI.GetValue('GL_MONTANTPA') / TOBI.GetValue('GL_QTEFACT'),V_PGI.okdecQ));
    if TOBI.GetValue('GL_QTEFACT') <> 0 then TOBI.PutValue ('GL_PUHTDEV',Arrondi(TOBI.GetValue('GL_MONTANTHTDEV')/ TOBI.GetValue('GL_QTEFACT'),V_PGI.OkDecP));
    if TOBI.GetValue('GL_QTEFACT') <> 0 then TOBI.PutValue ('GL_PUHT',Arrondi(TOBI.GetValue('GL_MONTANTHT')/ TOBI.GetValue('GL_QTEFACT'),V_PGI.OkDecP));
    if TOBI.GetValue('GL_QTEFACT') <> 0 then TOBI.PutValue ('GL_DPR',Arrondi(TOBI.GetValue('GL_MONTANTPR')/ TOBI.GetValue('GL_QTEFACT'),V_PGI.OkDecP));
  end;

var
    LaFamille,LeType: string;
    Article,Fournisseur : string;
    II,IndInsert : Integer;
    TOBP,TOBX : TOB;
    TypeLigne : string;
begin
  Result := nil;
  TOBX := nil;
  IndInsert := -1;
  LeType := TOBL.GetString('ZZTRI0');
  if TOBL.getSTring('GL_TYPEARTICLE')='PRE' then
  begin
    LaFamille := GetNaturePrest(TOBL.GetString('GL_ARTICLE'));
  end else
  begin
    LaFamille := TOBL.GetString('GL_FAMILLENIV1');
    if LaFamille = '' then LaFamille := '@@@';
  end;
  Article := TOBL.GetString('GL_ARTICLE');
  if TOBL.GetValue('GLC_NATURETRAVAIL')='002' then
  begin
    Fournisseur := TOBL.GetValue('GL_FOURNISSEUR');
  end else
  begin
    Fournisseur := '';
  end;
  for II := TOBDEBUT.GetIndex+1 to TOBGenere.Detail.count -1 do
  begin
    TOBP := TOBGenere.detail[II];
    if (TOBP.GetString('ZZTRI0')<> LeType) or (TOBP.GetString('ZZTRI1')<> LaFamille) then
    begin
      IndInsert := II+1;
      break;
    end;
    if (TOBP.GETString('GL_ARTICLE')=Article) and (TOBP.GETString('GL_FOURNISSEUR')=Fournisseur) then
    begin
      TOBX := TOBP;
      break;
    end;
    if (TOBP.GETString('GL_ARTICLE')>Article) then
    begin
      IndInsert := II+1;
      break;
    end;
  end;
  if TOBX = nil then
  begin
     TOBX := AddLigne(TOBPiece,TOBDEBUT,TOBL,IndInsert,LeType,LaFamille);
  end else
  begin
    CumuleLigne(TOBX,TOBL);
  end;
end;

procedure TGenerePrev.RemplacePourcent(TOBL,TOBPiece: TOB);
begin
  TOBL.PutValue('GL_CODEARTICLE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_ARTICLE',TOBA.GetValue('GA_ARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_TYPEARTICLE',TOBA.GetValue('GA_TYPEARTICLE'));
  TOBL.PutValue('GL_QTEFACT',1);
  TOBL.PutValue('GL_QUALIFQTEVTE',TOBA.GetVAlue('GA_QUALIFUNITEVTE'));
  TOBL.PutValue('GL_PRIXPOURQTE',1);
  TOBL.Putvalue('GL_RECALCULER','X');
  if TOBPiece.GetValue('GP_FACTUREHT')='X' then
  begin
    TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('GL_MONTANTHTDEV'));
    CalculeLigneHT (TOBL,nil,TOBPiece,DEV,DEV.decimale);
  end else
  begin
    TOBL.PutValue('GL_PUTTCDEV',TOBL.GetValue('GL_MONTANTTTCDEV'));
    CalculeLigneTTC (TOBL,nil,TOBPiece,DEV,DEV.decimale);
  end;
end;

class function TPrevisionTools.MajAffaireDevis(TOBpiece: TOB): boolean;
var Cledoc : r_cledoc;
begin
  cledoc := TOB2cledoc (TOBpiece);
  if ExecuteSql ('UPDATE AFFAIRE SET AFF_PREPARE="X", AFF_DATEMODIF="' + USDATETIME(NowH) + '" '+
              'WHERE '+
              'AFF_AFFAIRE=('+
              'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(cledoc,ttdPiece,false)+')')=0  then
end;

end.
