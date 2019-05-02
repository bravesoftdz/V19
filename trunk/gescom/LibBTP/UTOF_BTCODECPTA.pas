{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/08/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCODECPTA ()
Mots clefs ... : TOF;BTCODECPTA
*****************************************************************}
Unit UTOF_BTCODECPTA ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
     db,
     {$IFNDEF DBXPRESS}
     dbtables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF}
     fe_main,
     mul,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HPanel,
     HTB97,
     HEnt1,
     HMsgBox,
     uTOB,
     Paramsoc,
     LookUp,
     UtilsGrille,
     UtilsEtat,
     Vierge,
     HSysMenu,
     HRichEdt,
     HRichOLE,
     UTOF;

Type
  TOF_BTCODECPTA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    //
    Action        : TActionFiche;
    //Variable nécessaire pour la gestion de l'état
    OptionEdition : TOptionEdition;
    fEtat         : THValComboBox;
    TheType       : String;
    TheNature     : String;
    TheTitre      : String;
    TheModele     : String;
    //
    VenteAchat  : THValComboBox;
    FamCptaArt  : THValComboBox;
    FamCptaTiers: THValComboBox;
    FamCptaAff  : THValComboBox;
    RegimeTaxe  : THValComboBox;
    FamTaxe     : THValComboBox;
    Etablissement: THValComboBox;
    //
    Rang        : THEdit;
    CptHTAchat  : THEdit;
    CptHTVente  : THEdit;
    CptHtStock  : THEdit;
    CptHTVarStk : THEdit;
    //
    TCptaArt    : THLabel;
    TCptaTiers  : THLabel;
    TCptaAff    : THLabel;
    TCptHTAchat : THLabel;
    TCptHTVente : THLabel;
    TCptStock   : THLabel;
    TCptVarStk  : THLabel;
    //
    OkCptaArt   : Boolean;
    OkCptaTiers : Boolean;
    OkCptaAff   : Boolean;
    AvecStock   : Boolean;
    NoAchat     : Boolean;
    AvecImmoDiv : Boolean;
    //
    BImprimer   : TToolbarButton97;
    BDelete     : TToolbarButton97;
    BVENTILACH  : TToolbarButton97;
    BVENTILVTE  : TToolbarButton97;
    BVENTILSTK  : TToolbarButton97;
    //
    GroupBoxEsc : TGroupBox;
    GroupBoxRem : TGroupBox;
    //
    TobCodeCPTA : TOB;
    TobEdition  : TOB;
    //
    PageCTRL    : TPageControl;
    //
    procedure AffichageInitEcran(OkAff: Boolean);
    procedure ChargeTobWithZoneEcran;
    procedure ChargeZoneEcranWithTob;
    function CompteExiste(Tablette, Valeur: string): boolean;
    procedure Controlechamp(Champ, Valeur: String);
    procedure CreateEdition;
    procedure CreateTOB;
    procedure DestroyTOB;
    procedure GetObjects;
    procedure OnChangeVenteAchat(Sender: TObject);
    procedure OnElipsisCptAchClick(Sender: Tobject);
    procedure OnElipsisCptStkClick(Sender: Tobject);
    procedure OnElipsisCptVarStkClick(Sender: Tobject);
    procedure OnElipsisCptVteClick(Sender: Tobject);
    procedure OnImprimeFiche(Sender: TObject);
    procedure RazZoneEcran;
    procedure SetScreenEvents;
    procedure VentilationAchat(Sender: TObject);
    procedure VentilationStock(Sender: TObject);
    procedure VentilationVente(Sender: TObject);
    procedure VentilCompteHT(Nature: String);
    function VerifTout: Boolean;
    procedure CalculRang;
    //

  end ;

Const
	// libellés des messages
  TexteMessage: array[1..4] of string  = (
          {1}   'Vous devez renseigner un compte comptable'
          {2},  'Ce compte n''existe pas'
          {3},  'La suppression de la ventilation comptable a échoué'
          {4},  'La suppression de la ventilation analytique a échoué'
                                         );





Implementation

Uses  EntGC,
      Ventil;

procedure TOF_BTCODECPTA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA.OnDelete ;
var StSQL : string;
begin
  Inherited ;

  if PGIAsk('Confirmez-vous la suppression de cette ventilation comptable ? ', 'Ventilation comptable') = Mryes then
  begin
    //suppression pure et simple de l'enregistrement avec confirmation
    StSQL := 'DELETE CODECPTA WHERE GCP_RANG="' + Rang.Text + '"';
    if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[3], ' Ventilation comptable');
    //Suppression des ventilation analytique associées
    StSQL := 'DELETE FROM VENTIL WHERE (V_NATURE like "HA%" OR V_NATURE like "HV%" OR V_NATURE like "ST%") ';
    StSQL := StSQL + 'AND V_COMPTE="'+ Rang.Text +'"';
    if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[4], ' Ventilation comptable');
  end;

  Ecran.ModalResult := 1;

end ;

procedure TOF_BTCODECPTA.OnUpdate ;
begin
  Inherited ;

  if not VerifTout then
  Begin
    Ecran.ModalResult := 0;
    Exit;
  End;

  ChargeTobWithZoneEcran;

  TobCodeCPTA.SetAllModifie (True);
  TobCodeCPTA.InsertOrUpdateDB(True);

  Ecran.ModalResult := 1;

end ;

procedure TOF_BTCODECPTA.OnLoad ;
Var StSql : string;
    QQ    : TQuery;
begin
  Inherited ;

    If (Rang.Text = '') or (Action=TaCreat) then exit;

    StSql := 'SELECT * FROM CODECPTA WHERE GCP_RANG=' + Rang.text;
    QQ := OpenSQL(StSql, False);

    If Not QQ.eof then
    begin
      TobCodeCPTA.SelectDB('CODECPTA', QQ);
      //
      ChargeZoneEcranWithTob;
      //
      If VenteAchat.Value = 'ACH' then
      begin
          BVENTILVTE.Visible := False;
          BVENTILACH.Visible := True;
      end;
      If VenteAchat.Value = 'VEN' then
      begin
        BVENTILVTE.Visible := True;
        BVENTILACH.Visible := False;
      end;
    end;

    Ferme(QQ);

end ;

procedure TOF_BTCODECPTA.OnArgument (S : String ) ;
var Critere : string;
    Champ   : string;
    Valeur  : string;
    x       : Integer;
begin
  Inherited ;
  //
  TFvierge(Ecran).FormResize := false;
  //
  AvecImmoDiv := GetParamSocSecur('SO_GCCPTAIMMODIV', False) ;
  AvecStock   := GetParamSocSecur('SO_GCINVPERM', False) ;
  OkCptaArt   := GetParamSocSecur('SO_GCVENTCPTAART', False);
  OkCptaTiers := GetParamSocSecur('SO_GCVENTCPTATIERS', False);
  OkCptaAff   := GetParamSocSecur('SO_GCVENTCPTAAFF', False);

  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  CreateEdition;
  //
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ  := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  //Remise à blanc des zone ecran
  RAZZoneEcran;

  //Affichage de l'écran initial
  AffichageInitEcran(False);

end ;

procedure TOF_BTCODECPTA.OnClose ;
begin
  Inherited ;

  DestroyTOB;

end ;

procedure TOF_BTCODECPTA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA.GetObjects;
begin

  if Assigned(GetControl('GCP_VENTEACHAT'))     then VenteAchat   := THValComboBox(GetControl('GCP_VENTEACHAT'));
  if Assigned(GetControl('GCP_COMPTAARTICLE'))  then FamCptaArt   := THValComboBox(GetControl('GCP_COMPTAARTICLE'));
  if Assigned(GetControl('GCP_COMPTATIERS'))    then FamCptaTiers := THValComboBox(GetControl('GCP_COMPTATIERS'));
  if Assigned(GetControl('GCP_COMPTAAFFAIRE'))  then FamCptaAff   := THValComboBox(GetControl('GCP_COMPTAAFFAIRE'));
  if Assigned(GetControl('GCP_ETABLISSEMENT'))  then Etablissement:= THValComboBox(GetControl('GCP_ETABLISSEMENT'));
  if Assigned(GetControl('GCP_REGIMETAXE'))     then RegimeTaxe   := THValComboBox(GetControl('GCP_REGIMETAXE'));
  if Assigned(GetControl('GCP_FAMILLETAXE'))    then FamTaxe      := THValComboBox(GetControl('GCP_FAMILLETAXE'));

  if Assigned(GetControl('TGCP_COMPTAARTICLE')) then TCptaArt     := THLabel(GetControl('TGCP_COMPTAARTICLE'));
  if Assigned(GetControl('TGCP_COMPTATIERS'))   then TCptaTiers   := THLabel(GetControl('TGCP_COMPTATIERS'));
  if Assigned(GetControl('TGCP_COMPTAAFFAIRE')) then TCptaAff     := THLabel(GetControl('TGCP_COMPTAAFFAIRE'));

  if Assigned(GetControl('GCP_RANG'))           then Rang         := THEdit(GetControl('GCP_RANG'));
  if Assigned(GetControl('GCP_CPTEGENEACH'))    then CptHTAchat   := THEdit(GetControl('GCP_CPTEGENEACH'));
  if Assigned(GetControl('GCP_CPTEGENEVTE'))    then CptHTVente   := THEdit(GetControl('GCP_CPTEGENEVTE'));
  if Assigned(GetControl('GCP_CPTEGENESTOCK'))  then CptHTStock   := THEdit(GetControl('GCP_CPTEGENESTOCK'));
  if Assigned(GetControl('GCP_CPTEGENEVARSTK')) then CptHTVarStk  := THEdit(GetControl('GCP_CPTEGENEVARSTK'));

  if Assigned(GetControl('TGCP_CPTEGENEACH'))   then TCptHTAchat  := THLabel(GetControl('TGCP_CPTEGENEACH'));
  if Assigned(GetControl('TGCP_CPTEGENEVTE'))   then TCptHTVente  := THLabel(GetControl('TGCP_CPTEGENEVTE'));
  if Assigned(GetControl('TGCP_CPTEGENESTOCK')) then TCptStock    := THLabel(GetControl('TGCP_CPTEGENESTOCK'));
  if Assigned(GetControl('TGCP_CPTEGENEVARSTK'))then TCptVarStk   := THLabel(GetControl('TGCP_CPTEGENEVARSTK'));


  if Assigned(GetControl('GROUPBOXESC'))        then GroupBoxEsc  := TGroupBox(GetControl('GROUPBOXESC'));
  if Assigned(GetControl('GROUPBOXREM'))        then GroupBoxRem  := TGroupBox(GetControl('GROUPBOXREM'));

  if Assigned(GetControl('BIMPRIMER'))          then BImprimer    := TToolbarButton97(GetControl('BIMPRIMER'));
  If Assigned(GetControl('BDelete'))            then BDelete      := TToolbarButton97(Getcontrol('Bdelete'));
  if assigned(Getcontrol('BVENTILACH'))         then BVENTILACH   := TToolbarButton97(Getcontrol('BVENTILACH'));
  if assigned(Getcontrol('BVENTILVTE'))         then BVENTILVTE   := TToolbarButton97(Getcontrol('BVENTILVTE'));
  if assigned(Getcontrol('BVENTILSTK'))         then BVENTILSTK   := TToolbarButton97(Getcontrol('BVENTILSTK'));

end;

procedure TOF_BTCODECPTA.SetScreenEvents;
begin

  VenteAchat.OnChange := OnChangeVenteAchat;
  //
  BImprimer.OnClick   := OnImprimeFiche;
  BVENTILACH.OnClick  := VentilationAchat;
  BVENTILVTE.OnClick  := VentilationVente;
  BVENTILSTK.OnClick  := VentilationStock;
  //
  CptHTVente.OnElipsisClick := OnElipsisCptVteClick;
  CptHTAchat.OnElipsisClick := OnElipsisCptAchClick;
  CptHtStock.OnElipsisClick := OnElipsisCptStkClick;
  CptHTVarStk.OnElipsisClick:= OnElipsisCptVarStkClick;
  //

end;


Procedure TOF_BTCODECPTA.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if Champ='RANG'         then Rang.text := Valeur;

end;

procedure TOF_BTCODECPTA.RazZoneEcran;
begin

  If Action=TaCreat then Rang.text := '';
  //
  VenteAchat.Value   := '';
  //
  FamCptaArt.Value   := '';
  FamCptaAff.Value   := '';
  FamCptaTiers.Value := '';
  //
  Etablissement.Value:= '';
  FamTaxe.Value      := '';
  RegimeTaxe.Value   := '';
  //
  CptHTAchat.text   := '';
  CptHTVente.text   := '';
  CptHtStock.text   := '';
  CptHTVarStk.text  := '';
  //

end;

Procedure TOF_BTCODECPTA.ChargeZoneEcranWithTob;
begin

  Rang.text           := TobCodeCPTA.GetString('GCP_RANG');
  VenteAchat.Value    := TobCodeCPTA.GetString('GCP_VENTEACHAT');
  //
  FamCptaArt.Value    := TobCodeCPTA.GetString('GCP_COMPTAARTICLE');
  FamCptaAff.Value    := TobCodeCPTA.GetString('GCP_COMPTAAFFAIRE');
  FamCptaTiers.Value  := TobCodeCPTA.GetString('GCP_COMPTATIERS');
  //
  Etablissement.Value := TobCodeCPTA.GetString('GCP_ETABLISSEMENT');
  FamTaxe.Value       := TobCodeCPTA.GetString('GCP_FAMILLETAXE');
  RegimeTaxe.Value    := TobCodeCPTA.GetString('GCP_REGIMETAXE');
  //
  CptHTAchat.Text     := TobCodeCPTA.GetString('GCP_CPTEGENEACH');
  CptHTVente.text     := TobCodeCPTA.GetString('GCP_CPTEGENEVTE');
  CptHtStock.text     := TobCodeCPTA.GetString('GCP_CPTEGENESTOCK');
  CptHTVarStk.text    := TobCodeCPTA.GetString('GCP_CPTEGENEVARSTK');

end;

Procedure TOF_BTCODECPTA.ChargeTobWithZoneEcran;
begin

  if Action = TaCreat then CalculRang;

  TobCodeCPTA.PutValue('GCP_RANG', Rang.text);

  TobCodeCPTA.PutValue('GCP_VENTEACHAT',     VenteAchat.Value);
  //
  TobCodeCPTA.PutValue('GCP_COMPTAARTICLE',  FamCptaArt.Value);
  TobCodeCPTA.PutValue('GCP_COMPTAAFFAIRE',  FamCptaAff.Value);
  TobCodeCPTA.PutValue('GCP_COMPTATIERS',    FamCptaTiers.Value);
  //
  TobCodeCPTA.PutValue('GCP_ETABLISSEMENT',  Etablissement.Value);
  TobCodeCPTA.PutValue('GCP_FAMILLETAXE',    FamTaxe.Value);
  TobCodeCPTA.PutValue('GCP_REGIMETAXE',     RegimeTaxe.Value);
  //
  TobCodeCPTA.PutValue('GCP_CPTEGENEACH',    CptHTAchat.text);
  TobCodeCPTA.PutValue('GCP_CPTEGENEVTE',    CptHTVente.text);
  TobCodeCPTA.PutValue('GCP_CPTEGENESTOCK',  CptHtStock.text);
  TobCodeCPTA.PutValue('GCP_CPTEGENEVARSTK', CptHTVarStk.text);

end;

Procedure TOF_BTCODECPTA.AffichageInitEcran(OkAff : Boolean);
begin

  FamCptaArt.Visible  := OkCptaArt;
  FamCptaAff.Visible  := OkCptaAff;
  FamCptaTiers.Visible := OkCptaTiers;

  TCptaArt.visible    := OkCptaArt;
  TCptaTiers.visible  := OkCptaTiers;
  TCptaAff.Visible    := OkCptaAff;

  CptHtStock.Visible  := AvecStock;
  CptHTVarStk.Visible := AvecStock;
  TCptStock.Visible   := AvecStock;
  TCptVarStk.Visible  := AvecStock;
  BVENTILSTK.Visible  := AvecStock;

  GROUPBOXESC.Visible := False;
  GroupBoxRem.Visible := False;

  Ecran.Height := Ecran.Height - (GroupBoxRem.Height + GroupBoxEsc.Height);

  BVENTILVTE.Visible := True;
  BVENTILACH.Visible := True;

  If Action = TaCreat then
    BDelete.Visible := false
  else
    BDelete.Visible := True;

end;

Procedure TOF_BTCODECPTA.CreateTOB;
begin

  TobCodeCPTA   := Tob.Create('CODECPTA', nil, -1);

  // Edition
  TobEdition    := TOB.Create(' EDITCODECPTA', nil, -1);

end;

procedure TOF_BTCODECPTA.DestroyTOB;
begin

  FreeAndNil(TobCodeCPTA);
  //
  FreeAndNil(TobEdition);
  FreeAndNil(OptionEdition);

end;

Function TOF_BTCODECPTA.VerifTout : Boolean ;
begin

  Result:= True ;

  if (CptHTAchat.Text = '') And (CptHTVente.text = '') then
  begin
    PGIError(TexteMessage[1], ' Ventilation comptable');
    Result:=false;
    if NoAchat then
      CptHtVente.SetFocus
    else
      CptHTAchat.SetFocus;
    Exit;
  end;

  If (Not NoAchat) and (CptHTAchat.text <> '') and (not CompteExiste('TZGCHARGE',CptHTAchat.text)) then
  Begin
    if (Not AvecImmoDiv) or
      ((Not CompteExiste('TZGIMMO',CptHTAchat.text)) and
       (Not CompteExiste('TZGDIVERS',CptHTAchat.text))) then
    begin
      PGIError(TexteMessage[2], ' Ventilation comptable');
      Result:=false;
      CptHTAchat.SetFocus;
      Exit;
    end ;
  End;

  If (CptHTVente.text <> '') and (not CompteExiste('TZGPRODUIT',CptHTVente.text)) then
  begin
    if (Not AvecImmoDiv) or
      ((Not CompteExiste('TZGIMMO',CptHTVente.text)) and
       (Not CompteExiste('TZGDIVERS',CptHTVente.text))) then
    Begin
      PGIError(TexteMessage[2], ' Ventilation comptable');
      Result:=false;
      CptHTVente.SetFocus;
      Exit;
    end;
  end;

  if AvecStock then
  Begin
    if (CptHtStock.text <> '') and (not CompteExiste('TZGDIVERS', CptHtStock.text)) then
    Begin
      PGIError(TexteMessage[2], ' Ventilation comptable');
      Result:=False ;
      CptHtStock.SetFocus;
      Exit;
    end;
    //
    if (CptHTVarStk.text <> '') and
       (not CompteExiste('TZGCHARGE',CptHTVarStk.text)) and
       (not CompteExiste('TZGPRODUIT',CptHTVarStk.text)) then
    Begin
      if (Not AvecImmoDiv) or
        ((Not CompteExiste('TZGIMMO',CptHTVarStk.text)) and
         (Not CompteExiste('TZGDIVERS',CptHTVarStk.text))) then
      Begin
        PGIError(TexteMessage[2], ' Ventilation comptable');
        Result:=False ;
        CptHTVarStk.SetFocus;
        Exit;
      End;
    End;
  end;

end ;

Function TOF_BTCODECPTA.CompteExiste (Tablette, Valeur : string) : boolean;
var St : string;
BEGIN

  Result := False;
  St := RechDom (Tablette, Valeur, False);
  if (St <> '') AND (St <> 'Error') then Result := True;

END;

procedure TOF_BTCODECPTA.CreateEdition;
begin
  //
  TheType       := 'E';
  TheNature     := 'PAR';
  TheTitre      := 'Ventilation comptable';
  TheModele     := 'VCP';
  //
  OptionEdition := TOptionEdition.Create(TheType,TheNature,TheModele, TheTitre, '', True, True, True, False, False, PageCTRL, fEtat);
  //
  OptionEdition.Apercu    := True;
  OptionEdition.DeuxPages := False;
  OptionEdition.Spages    := PageCTRL;

end;

Procedure TOF_BTCODECPTA.OnElipsisCptVteClick(Sender : Tobject);
Var Criteres : string;
begin

  if AvecImmoDiv then
    Criteres  := '(G_NATUREGENE="PRO" OR G_NATUREGENE="IMO" OR G_NATUREGENE="DIV")'
  else
    Criteres  := 'G_NATUREGENE="PRO"' ;

  LookupList(CptHTVente,TraduireMemoire('Comptes de produits'),'GENERAUX','G_GENERAL','G_LIBELLE',Criteres,'G_GENERAL',TRUE, 0) ;

end;

Procedure TOF_BTCODECPTA.OnElipsisCptAchClick(Sender : Tobject);
Var Criteres : string;
begin

  if AvecImmoDiv then
    Criteres:='(G_NATUREGENE="CHA" OR G_NATUREGENE="IMO" OR G_NATUREGENE="DIV")'
  else
    Criteres:='G_NATUREGENE="CHA"' ;

  LookupList(CptHTAchat,TraduireMemoire('Comptes de charges'),'GENERAUX','G_GENERAL','G_LIBELLE',Criteres,'G_GENERAL',TRUE, 0) ;

end;

Procedure TOF_BTCODECPTA.OnElipsisCptStkClick(Sender : Tobject);
Var Criteres : string;
begin

  if AvecStock then
  begin
    Criteres:='G_NATUREGENE="DIV"' ;
    LookupList(CptHTStock,TraduireMemoire('Comptes divers de stock'),'GENERAUX','G_GENERAL','G_LIBELLE',Criteres,'G_GENERAL',TRUE, 0) ;
  end ;

end;

Procedure TOF_BTCODECPTA.OnElipsisCptVarStkClick(Sender : Tobject);
Var Criteres : string;
begin

  if AvecStock then
  begin
    if AvecImmoDiv then
      Criteres:='(G_NATUREGENE="CHA" OR G_NATUREGENE="PRO" OR G_NATUREGENE="IMO" OR G_NATUREGENE="DIV")'
    else
      Criteres:='G_NATUREGENE="CHA" OR G_NATUREGENE="PRO" ' ;

   LookupList(CptHTVarStk,TraduireMemoire('Comptes de variation de stock'),'GENERAUX','G_GENERAL','G_LIBELLE',Criteres,'G_GENERAL',TRUE, 0) ;
  end ;

end;


Procedure TOF_BTCODECPTA.OnImprimeFiche(Sender : TObject);
Var StSQL : String;
    QQ    : TQuery;
begin

  TobEdition.ClearDetail;

  StSQL := 'SELECT * FROM CODECPTA';
  QQ := OpenSQL(StSQL, False);

  If not QQ.eof then
  begin
    TobEdition.LoadDetailDB('EDITION VENTILCPTA','','',QQ, False);
  end;

  Ferme(QQ);

  if OptionEdition.LanceImpression('', TobEdition) < 0 then V_PGI.IoError:=oeUnknown;

end;

Procedure TOF_BTCODECPTA.OnChangeVenteAchat(Sender : TObject);
begin

  if VenteAchat.Value = 'ACH' then
  begin
    BVENTILACH.Visible := True;
    BVENTILVTE.Visible := False;
    BVENTILSTK.Visible := False;
  end;

  if VenteAchat.Value = 'VEN' then
  begin
    BVENTILACH.Visible := False;
    BVENTILVTE.Visible := True;
    BVENTILSTK.Visible := False;
  end;

  if VenteAchat.Value = 'CON' then
  begin
    BVENTILACH.Visible := False;
    BVENTILVTE.Visible := False;
    BVENTILSTK.Visible := False;
  end;

end;

Procedure TOF_BTCODECPTA.VentilationAchat(Sender : TObject);
begin
  VentilCompteHT('HA');
end;

Procedure TOF_BTCODECPTA.VentilationVente(Sender : TObject);
begin

  VentilCompteHT('HV');

end;

Procedure TOF_BTCODECPTA.VentilationStock(Sender : TObject);
begin

  VentilCompteHT('ST')

end;

procedure TOF_BTCODECPTA.VentilCompteHT(Nature : String) ;
Var StAxes : String ;
begin

  if (Nature = 'HV') and (CptHTVente.text = '') then Exit;
  //
  if (Nature = 'HA') and (CptHTAchat.text = '') then Exit;
  //
  if (Nature = 'ST') and (CptHTStock.text = '') then Exit;

  if EstSerie(S7) then
    StAxes:='12345'
  else
  BEGIN
    StAxes:='1' ;
    {$IFNDEF CCS3}
    if VH_GC.GCventAxe2 then StAxes:=StAxes+'2' ;
    if VH_GC.GCventAxe3 then StAxes:=StAxes+'3' ;
    {$ENDIF}
  END ;

  if (Nature = 'HV') then ParamVentil(Nature,CptHTVente.text, StAxes, taCreat, FALSE);
  if (Nature = 'HA') then ParamVentil(Nature,CptHTAchat.text, StAxes, taCreat, FALSE);
  if (Nature = 'ST') then ParamVentil(Nature,CptHTStock.text, StAxes, taCreat, FALSE);

end ;

procedure TOF_BTCODECPTA.CalculRang;
var StSQl : string;
    QQ    : TQuery;
begin

  StSQl := 'select Max(GCP_RANG) as MaxRang From CodeCPTA';
  QQ := OpenSQL(StSQl, False);
  If not QQ.eof then
  begin
    Rang.text := InttoStr(QQ.FindField('MaxRang').AsInteger + 1)
  end;

  Ferme(QQ);

end;

Initialization
  registerclasses ( [ TOF_BTCODECPTA ] ) ; 
end.

