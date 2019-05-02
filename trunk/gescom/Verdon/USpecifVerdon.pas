unit USpecifVerdon;

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
  affaireutil,FE_Main,Paramsoc
  ;

type
  TTypeEnvoi = (TTevMt,TTevST);
  
  TOptionVerdonSais = class (TObject)
    private
      fActif : Boolean;
      FF : TForm;
      POPGS : TPopupMenu;
      fMaxItems : Integer;
      fCreatedPop : Boolean;
      MesMenuItem : array[0..2] of TMenuItem;
      //
      procedure DefiniMenuPop(Parent: Tform);
    procedure VerdonVoirStockWTT(Sender: TObject);
    public
      constructor create (Parent : Tform);
      Destructor destroy; override;
  end;

  tToolsVerdon = class
  private
    class procedure EnvoiMail(TOBP : TOB; NatureEnvoie : TTypeEnvoi);
    class function CreerAffaireVerdon(Numero : integer; CodeTiers, Etablissement, Devise, LabelAff : string) : string;
    class procedure UpdateTable(Prefix, CodeAff, Aff1, Aff2, Aff3, Aven, NaturePiece, Souche : string; Numero, Indice : integer);
    class function AffecteAffaire(CodeTiers, NaturePiece, Souche : string; Numero, Indice : integer) : string;
    class procedure ReaffecteConsoSurAffaire(CodeAffaire,NaturePiece,Souche : string; Numero,Indice : integer);
  public
    class function TraiteAcceptationVerdon(CodeTiers, NaturePiece, Souche : string; Numero, Indice : integer) : boolean;
    class procedure ChargeDestinatairesVerdon;
    class procedure LibereDestinatairesVerdon;
  end;

var TOBDestVERDON : TOB;

implementation

uses
  Facture
  , FactTOB
  , UtilsMail
  , ErrorsManagement
  , FormsName
  , CommonTools
  , CalcOLEGenericBTP
  , UPrevisionTools
  , Messages
  ;
  
class procedure tToolsVerdon.LibereDestinatairesVerdon;
begin
  TOBDestVERDON.free;
  TOBDestVERDON := nil;
end;

class procedure tToolsVerdon.ChargeDestinatairesVerdon;
var QQ : Tquery;
begin
  if TOBDestVERDON = nil then
  begin
    TOBDestVERDON := TOB.Create ('LES DESTINATAIRES',nil,-1);
  end;
  QQ := OpenSQL('SELECT * FROM BTVERDDESTMAIL',True,-1,'',True);
  if not QQ.eof then
  begin
    TOBDestVERDON.LoadDetailDB('BVERDESTMAIL','','',QQ,false);
  end;
  Ferme(QQ);
end;

class function tToolsVerdon.CreerAffaireVerdon(Numero : integer; CodeTiers, Etablissement, Devise, LabelAff : string) : string;
var
  P0             : THCritMaskEdit;
  P2             : THCritMaskEdit;
  P1             : THCritMaskEdit;
  P3             : THCritMaskEdit;
  Av             : THCritMaskEdit;
  CodeAffaire    : string;
  TobAFF         : TOB;
  IP  : Integer;
begin
  TobAFF := TOB.Create('AFFAIRE',  nil, -1);
  P0 := THCritMaskEdit.Create(Application);  P0.text := 'A';
  P1 := THCritMaskEdit.Create(Application);  P1.text := '';
  P2 := THCritMaskEdit.Create(Application);  P2.text := '';
  P3 := THCritMaskEdit.Create(Application);  P3.text := '';
  Av := THCritMaskEdit.Create(Application);  Av.text := '';
  try
    CodeAffaire := DechargeCleAffaire (P0,P1,P2,P3,Av,CodeTiers,taCreat,true,True,False,IP);
    if P1.Text = '' then
    begin
      Result := '';
      Exit;
    end;
    TobAFF.InitValeurs;
    TobAFF.SetString('AFF_AFFAIRE'       , CodeAffaire);
    TobAFF.SetString('AFF_AFFAIREREF'    , CodeAffaire);
    TobAFF.SetString('AFF_AFFAIREINIT'   , CodeAffaire);
    TobAFF.SetString('AFF_AFFAIRE0'      , P0.Text);
    TobAFF.SetString('AFF_AFFAIRE1'      , P1.text);
    TobAFF.SetString('AFF_AFFAIRE2'      , P2.text);
    TobAFF.SetString('AFF_AFFAIRE3'      , P3.text);
    TobAFF.SetString('AFF_AVENANT'       , Av.text);
    TobAFF.SetString('AFF_RECONDUCTION'  , VH_GC.AFFReconduction);
    TobAFF.SetString('AFF_MANDATAIRE'    ,'');
    TobAFF.SetString('AFF_PERIODICITE'   , GetParamSocSecur('SO_AfPeriodicte', 'M'));
    TobAFF.SetString('AFF_GENERAUTO'     , VH_GC.AFFGenerAuto );
    TobAFF.SetString('AFF_TYPEPREVU'     , 'GLO');
    TobAFF.SetString('AFF_NUMDERGENER'   , '');
    TobAFF.SetString('AFF_PROFILGENER'   , GetParamSocSecur('SO_AFPROFILGENER', ''));
    TobAFF.SetString('AFF_TERMEECHEANCE' , GetParamSocSecur('SO_AFTERMEECHE', ''));
    TobAFF.SetString('AFF_METHECHEANCE'  , GetParamSocSecur('SO_AFMETHECHE', ''));
    TobAFF.SetString('AFF_STATUTAFFAIRE' , 'AFF');
    TobAFF.SetString('AFF_ETATAFFAIRE'   , 'ENC');
    TobAFF.SetString('AFF_REGROUPEFACT'  , 'AUC');
    TobAFF.SetString('AFF_CREATEUR'      , V_PGI.User);
    TobAFF.SetString('AFF_UTILISATEUR'   , V_PGI.User);
    TobAFF.SetString('AFF_CREERPAR'      , 'SAI');
    TobAFF.SetString('AFF_ETABLISSEMENT' , Etablissement);
    TobAFF.SetString('AFF_DEVISE'        , Devise);
    TobAFF.SetString('AFF_REPRISEACTIV'  , 'TOU');
    TobAFF.SetString('AFF_TIERS'         , Codetiers);
    TobAFF.SetString('AFF_DESCRIPTIF'    , LabelAff);
    TobAFF.SetString('AFF_LIBELLE'       , LabelAff);
    TobAFF.SetDateTime('AFF_DATESSTRAIT' , iDate1900);
    TobAFF.SetDateTime('AFF_DATECREATION', V_PGI.DateEntree);
    TobAFF.SetDateTime('AFF_DATEDEBUT'   , V_PGI.DateEntree);
    TobAFF.SetDateTime('AFF_DATEFIN'     , idate2099);
    TobAFF.SetDateTime('AFF_DATEDEBGENER', idate1900);
    TobAFF.SetDateTime('AFF_DATEFINGENER', idate2099);
    TobAFF.SetDateTime('AFF_DATELIMITE'  , idate2099);
    TobAFF.SetDateTime('AFF_DATERESIL'   , idate2099);
    TobAFF.SetDateTime('AFF_DATECLOTTECH', idate2099);
    TobAFF.SetDateTime('AFF_DATEGARANTIE', idate2099);
    TobAFF.SetDateTime('AFF_DATECUTOFF'  , idate1900);
    TobAFF.SetDateTime('AFF_DATESIGNE'   , idate2099);
    TobAFF.SetInteger('AFF_INTERVALGENER', GetParamSocSecur('SO_AFINTERVAL',1));
    TobAFF.SetInteger('AFF_COEFFREVALO'  , 1);
    TobAFF.SetBoolean('AFF_SSTRAITANCE'  , False);
    TobAFF.SetBoolean('AFF_CALCTOTHTGLO' , True);
    TobAFF.SetBoolean('AFF_ADMINISTRATIF', False);
    TobAFF.SetBoolean('AFF_MODELE'       , False);
    TobAFF.SetBoolean('AFF_AFFAIREHT'    , False);
    TobAFF.SetBoolean('AFF_REGSURCAF'    , False);
    TobAFF.SetBoolean('AFF_SAISIECONTRE' , False);
    TobAFF.SetBoolean('AFF_AFFCOMPLETE'  , True);
    TobAFF.SetBoolean('AFF_AFFCOMPLETE'  , True);
    TobAFF.SetDouble('MONTANTGLOBAL'     , 0);
    TobAFF.SetDouble('MONTANTDEJAFACT'   , 0);
    TobAFF.SetDouble('MONTANTAFACT'      , 0);
    if TobAFF.InsertDB(nil) then
      Result := Format('%s;%s;%s;%s;%s', [CodeAffaire, P1.Text, P2.text, P3.text, Av.text])
    else
      Result := '';
  finally
    FreeAndNil(TobAFF);
    P0.Free;
    P1.free;
    P2.free;
    P3.Free;
    Av.Free;
  end;
end;

class procedure tToolsVerdon.UpdateTable(Prefix, CodeAff, Aff1, Aff2, Aff3, Aven, NaturePiece, Souche : string; Numero, Indice : integer);
var
  Sql : string;
begin
  if Prefix <> '' then
  begin
    Sql := Format('UPDATE %s'
                + ' SET %s_AFFAIRE  = "%s"'
                + '   , %s_AFFAIRE1 = "%s"'
                + '   , %s_AFFAIRE2 = "%s"'
                + '   , %s_AFFAIRE3 = "%s"'
                + '   , %s_AVENANT  = "%s"'
                + ' WHERE %s_NATUREPIECEG = "%s"'
                + '   AND %s_SOUCHE       = "%s"'
                + '   AND %s_NUMERO       = %s'
                + '   AND %s_INDICEG      = %s'
               , [  Tools.GetTableNameFromTtn(Tools.GetTtnFromPrefix(Prefix))
                  , Prefix, CodeAff
                  , Prefix, Aff1
                  , Prefix, Aff2
                  , Prefix, Aff3
                  , Prefix, Aven
                  , Prefix, Naturepiece
                  , Prefix, Souche
                  , Prefix, IntToStr(Numero)
                  , Prefix, IntToStr(Indice)
                 ]);
    ExecuteSql(Sql);
  end;
end;

class procedure tToolsVerdon.EnvoiMail (TOBP : TOB; NatureEnvoie : TTypeEnvoi);
var XX : TGestionMail;
begin
  XX := TGestionMail.Create(Application);

  if NatureEnvoie=TTevMt then
      XX.Sujet := 'Acceptation de devis - Depassement montant'
  else
      XX.Sujet := 'Acceptation de devis - Sous traitance présente';
  //    
  XX.Corps := hTStringList.Create;
  XX.Corps.Clear ;

  XX.Copie         := '';
  XX.TypeContact   := '';
  XX.Fournisseur   := '';
  XX.FichierSource := '';
  XX.FichierTempo  := '';
  XX.Fichiers      := '';
  XX.TypeDoc       := '';
  //Pourrait être déterminé par le type d'enregistrement traité ou par le type de planning (????)
  XX.Tiers         := '';
  XX.Contact       := '';
  XX.Destinataire  := '';
  if NatureEnvoie=TTevMt then
  begin
    XX.QualifMail    := 'VVM';
  end else if NatureEnvoie=TTevST then
  begin
    XX.QualifMail    := 'VVS';
  end;
  XX.TobRapport    := TOBP;
  XX.GestionParam  := True;
  XX.ListeDestinataires := TOBDestVERDON;

  XX.AppelEnvoiMail (false,False);

  FreeAndNil(XX);

end;

class function tToolsVerdon.AffecteAffaire(CodeTiers, NaturePiece, Souche : string; Numero, Indice : integer) : string;
var
  RetChoixAffaire : string;
  ActionAffaire   : string;
  LibelleAffaire  : string;

  procedure UpdateLesTables(CodeAff, Aff1, Aff2, Aff3, Aven : string);
  begin
    begintrans;
    try
      UpdateTable('GP' , CodeAff, Aff1, Aff2, Aff3, Aven, NaturePiece, Souche, Numero, Indice);
      UpdateTable('GL' , CodeAff, Aff1, Aff2, Aff3, Aven, NaturePiece, Souche, Numero, Indice);
      UpdateTable('BLO', CodeAff, Aff1, Aff2, Aff3, Aven, NaturePiece, Souche, Numero, Indice);
      UpdateTable('BOP', CodeAff, Aff1, Aff2, Aff3, Aven, NaturePiece, Souche, Numero, Indice);
      Result := CodeAff;
      CommitTrans;
    except
      Result := '';
      Rollback;
    end;
  end;

  function CreerAffaire : string;
  var
    Sql      : string;
    Etabliss : string;
    Devise   : string;
    RetAff   : string;
    CodeAff  : string;
    Aff1     : string;
    Aff2     : string;
    Aff3     : string;
    Aven     : string;
    Qry      : TQuery;
  begin
    Sql := Format('SELECT GP_ETABLISSEMENT, GP_DEVISE FROM PIECE WHERE GP_NATUREPIECEG = "%s" AND GP_SOUCHE = "%s" AND GP_NUMERO = %s AND GP_INDICEG = %s'
               , [  NaturePiece
                  , Souche
                  , IntToStr(Numero)
                  , IntToStr(Indice)
                 ]);
    Qry := OpenSql(Sql, True);
    try
      Etabliss := Qry.FindField('GP_ETABLISSEMENT').AsString;
      Devise   := Qry.FindField('GP_DEVISE').AsString;
    finally
      Ferme(Qry);
    end;
    RetAff := CreerAffaireVerdon(Numero, CodeTiers, Etabliss, Devise, LibelleAffaire);
    if RetAff <> '' then
    begin
      CodeAff := Tools.ReadTokenSt_(RetAff, ';');
      Aff1    := Tools.ReadTokenSt_(RetAff, ';');
      Aff2    := Tools.ReadTokenSt_(RetAff, ';');
      Aff3    := Tools.ReadTokenSt_(RetAff, ';');
      Aven    := Tools.ReadTokenSt_(RetAff, ';');
      UpdateLesTables(CodeAff, Aff1, Aff2, Aff3, Aven);
      Result := CodeAff;
    end else
      Result := '';
  end;

  function AssocierAffaire : string;
  var
    Parms   : string;
    Range   : string;
    CodeAff : string;
    SqlAFF  : string;
    Aff1    : string;
    Aff2    : string;
    Aff3    : string;
    Aven    : string;
    Qry     : TQuery;

  begin
    Result  := '';
    Parms   := 'ACTION=RECH';
    Range   := Format('AFF_TIERS=%s;AFF_AFFAIRE0=A', [CodeTiers]);
    CodeAff := AGLLanceFiche('BTP', 'BTAFFAIRE_MUL', Range, '', Parms);
    if CodeAff <> '' then
    begin
      CodeAff := copy(CodeAff, 1, pos(';', CodeAff)-1);
      SqlAFF  := Format('SELECT AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_AVENANT FROM AFFAIRE WHERE AFF_AFFAIRE = "%s"', [CodeAFF]);
      Qry     := OpenSql(SqlAFF, True);
      try
        Aff1 := Qry.FindField('AFF_AFFAIRE1').AsString;
        Aff2 := Qry.FindField('AFF_AFFAIRE2').AsString;
        Aff3 := Qry.FindField('AFF_AFFAIRE3').AsString;
        Aven := Qry.FindField('AFF_AVENANT').AsString;
      finally
        Ferme(Qry);
      end;
      UpdateLesTables(CodeAff, Aff1, Aff2, Aff3, Aven);
      Result := CodeAff;
    end;
  end;

begin
  Result := '';
  RetChoixAffaire := OpenForm.AddGetAffaire(Format('TIERS=%s', [CodeTiers]));
  if RetChoixAffaire <> '' then
  begin
    ActionAffaire   := Tools.ReadTokenSt_(RetChoixAffaire, ';');
    LibelleAffaire  := Tools.ReadTokenSt_(RetChoixAffaire, ';');
    { ActionAffaire : A=Annuler, C=Créer, L=Lier}
    if ActionAffaire <> 'A' then
    begin
      case Tools.CaseFromString(ActionAffaire, ['C', 'L']) of
        {Créer}    0 : Result := CreerAffaire;
        {Associer} 1 : Result := AssocierAffaire;
      end;
    end;
  end;
end;

class procedure tToolsVerdon.ReaffecteConsoSurAffaire(CodeAffaire,NaturePiece,Souche : string; Numero,Indice : integer);
var SQl : string;
    A0,A1,A2,A3,A4 : string;
begin
  BTPCodeAffaireDecoupe(CodeAffaire,A0,A1,A2,A3,A4, TaModif,false);
  //
  SQL := 'SELECT 1 FROM CONSOMMATIONS WHERE '+
         'BCO_NATUREPIECEG="'+NaturePiece+'" AND '+
         'BCO_SOUCHE="'+Souche+'" AND '+
         'BCO_NUMERO='+IntToStr(Numero)+' AND '+
         'BCO_INDICEG='+InttoStr(Indice)+' AND '+
         'BCO_AFFAIRE LIKE "AACDC%"';
  if ExisteSQL(SQL) then
  begin
    SQL := 'UPDATE CONSOMMATIONS SET BCO_AFFAIRE="'+CodeAffaire+'",'+
           'BCO_AFFAIRE0="'+A0+'",'+
           'BCO_AFFAIRE1="'+A1+'",'+
           'BCO_AFFAIRE2="'+A2+'",'+
           'BCO_AFFAIRE3="'+A3+'",'+
           'BCO_AVENANT="'+A4+'" '+
           'WHERE '+
           'BCO_NATUREPIECEG="'+NaturePiece+'" AND '+
           'BCO_SOUCHE="'+Souche+'" AND '+
           'BCO_NUMERO='+InttoStr(Numero)+' AND '+
           'BCO_INDICEG='+InttoStr(Indice)+' AND '+
           'BCO_AFFAIRE LIKE "AACDC%"';
    if ExecuteSQL(SQL) = 0 then
    begin
      Raise Exception.create('Impossible de réajuster les consommations');
    end;
  end;
end;

class function tToolsVerdon.TraiteAcceptationVerdon(CodeTiers, NaturePiece, Souche : string; Numero, Indice : integer) : boolean;
var
  TT          : TOB;
  QQ          : TQuery;
  SQl         : string;
  CodeAffaire : string;
  MtAccepte   : double;
  MtDevis     : double;
  StPresent   : Boolean;
begin
  BEGINTRANS;
  TRY
    CodeAffaire := AffecteAffaire(CodeTiers, NaturePiece, Souche, Numero, Indice);
    Result      := (CodeAffaire <> '');
    if Result then
    begin
      Result := TPrevisionTools.GenerePrevisionChantier(NaturePiece,Souche,Numero,Indice);
      if not Result then
      begin
        raise Exception.Create('Erreur lors de la génération de la prévision de chantier');
        Exit;
      end;
      ReaffecteConsoSurAffaire(CodeAffaire,NaturePiece,Souche,Numero,Indice);
      StPresent := false;
      MtAccepte := 0;
      MtDevis   := 0;
      TT := TOB.Create ('PIECE',nil,-1);
      try
        SQL := 'SELECT * FROM PIECE '+
               'LEFT JOIN TIERS ON T_NATUREAUXI="CLI" AND GP_TIERS=T_TIERS '+
               'LEFT JOIN AFFAIRE ON GP_AFFAIRE=AFF_AFFAIRE '+
               'WHERE '+
               'GP_NATUREPIECEG="'+naturepiece+'" AND '+
               'GP_SOUCHE="'+Souche+'" AND '+
               'GP_NUMERO='+InttoStr(Numero)+' AND '+
               'GP_INDICEG='+InttoStr(Indice);
        QQ := OpenSql (SQL,True,1,'',true);
        begin
          if not QQ.eof then
          begin
            TT.SelectDB('',QQ);
            MtDevis := TT.GetDouble('GP_TOTALHTDEV');
          end;
        end;
        ferme (QQ);
        //
        SQL := 'SELECT SUM(P2.GP_TOTALHTDEV) FROM PIECE P2 '+
               'WHERE '+
               'P2.GP_AFFAIRE=(SELECT P1.GP_AFFAIRE FROM PIECE P1 WHERE '+
               'P1.GP_NATUREPIECEG="'+naturepiece+'" AND '+
               'P1.GP_SOUCHE="'+Souche+'" AND '+
               'P1.GP_NUMERO='+InttoStr(Numero)+' AND '+
               'P1.GP_INDICEG='+InttoStr(Indice)+') AND '+
               '(SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE=P2.GP_AFFAIREDEVIS)="ACP"';
        QQ := OpenSql (SQL,True,1,'',true);
        begin
          if not QQ.eof then
          begin
            MtAccepte := QQ.fields[0].AsFloat;
          end;
        end;
        ferme (QQ);
        //
        SQL := 'SELECT 1 AS EXIST FROM LIGNE '+
               'LEFT JOIN ARTICLE ON GL_ARTICLE=GA_ARTICLE '+
               'WHERE '+
               'GA_NATUREPRES IN (SELECT BNP_NATUREPRES FROM NATUREPREST WHERE BNP_TYPERESSOURCE="ST") AND '+
               'GL_NATUREPIECEG="'+naturepiece+'" AND '+
               'GL_SOUCHE="'+Souche+'" AND '+
               'GL_NUMERO='+InttoStr(Numero)+' AND '+
               'GL_INDICEG='+InttoStr(Indice);
        if ExisteSQL(SQL) then
        begin
          StPresent := True;
        end;
        //
        if ARRONDI(MtAccepte+MTDevis,V_PGI.okdecV) > 25000 then
        begin
          // Envoie mail pour montant depassé
          EnvoiMail (TT,TTevMt);
        end else if StPresent then
        begin
          EnvoiMail (TT,TTevST);
          // Envoie mail pour sous traitance      
        end;
      finally
        TT.free;
      end;
    end else
    begin
      if CodeAffaire <> '' then
        raise  Exception.Create('Impossible de créer/affecter le chantier');
    end;
    COMMITTRANS;
  EXCEPT
    Result := false;
    ROLLBACK;
  END;
end;

constructor TOptionVerdonSais.create(Parent: Tform);
var ThePop : Tcomponent;
begin
  fActif := false;
  if VH_GC.BTCODESPECIF = '002' then
  begin
    fActif := True;
    FF := Parent;
    ThePop := Parent.Findcomponent  ('POPBTP');
    if ThePop = nil then
    BEGIN
      // pas de menu BTP trouve ..on le cree
      POPGS := TPopupMenu.Create(Parent);
      POPGS.Name := 'POPBTP';
      fCreatedPop := true;
    END else
    BEGIN
      fCreatedPop := false;
      POPGS := TPopupMenu(thePop);
    END;
    DefiniMenuPop(Parent);
  end;
end;

procedure TOptionVerdonSais.DefiniMenuPop (Parent : Tform);
var Indice : integer;

begin
  fMaxItems := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with MesMenuItem[fMaxItems] do
      begin
      Caption := '-';
      end;
    inc (fMaxItems);
  end;
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'VERDONVOIRSTOCK';
    Caption := TraduireMemoire ('Voir le stock(WTT)');
    OnClick := VerdonVoirStockWTT;
    end;
  inc (fMaxItems);
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Caption := '-';
    end;
  inc (fMaxItems);

  for Indice := 0 to fMaxItems -1 do
    begin
      if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
    end;
end;

destructor TOptionVerdonSais.destroy;
var Indice : integer;
begin
  if not fActif then Exit;
  for Indice := 0 to fMaxItems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
  if fcreatedPop then POPGS.free;
  //
  inherited;
end;

procedure TOptionVerdonSais.VerdonVoirStockWTT (Sender : TObject);
begin
  TFFacture(FF).VerdonVoirStockClick(Sender);
end;
end.
