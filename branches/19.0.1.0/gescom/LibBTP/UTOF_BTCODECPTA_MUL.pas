{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/08/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCODECPTA_MUL ()
Mots clefs ... : TOF;BTCODECPTA_MUL
*****************************************************************}
Unit UTOF_BTCODECPTA_MUL ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
     uTob,
{$ELSE}
     DBCtrls, Db,
     {$IFNDEF DBXPRESS}
     dbTables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF DBXPRESS}
     fe_main,
     HDB,
     Mul,
{$ENDIF EAGLCLIENT}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     //Ajout
     HTB97,
     uEntCommun,
     uTOFComm,
     ParamSoc,
     UTOF;

Type
  TOF_BTCODECPTA_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  Private
    //Définition des variables utilisées dans le Uses
    Action      : TActionFiche;
    ModeRecherche : Boolean;
    //
    VenteAchat  : THValComboBox;
    FamCptaArt  : THValComboBox;
    FamCptaTiers: THValComboBox;
    FamCptaAff  : THValComboBox;
    RegimeTaxe  : THValComboBox;
    FamTaxe     : THValComboBox;
    //
    TCptaArt    : THLabel;
    TCptaTiers  : THLabel;
    TCptaAff    : THLabel;
    //
    Grille      : THDBGrid;
    BInsert     : TToolBarButton97;
    //
    procedure Controlechamp(Champ, Valeur: String);
    procedure GetObjects;
    procedure GrilleOnDblclick(Sender: TObject);
    procedure InsertOnclick(Sender: TObject);
    procedure SetScreenEvents;
    //
  end ;

Implementation

procedure TOF_BTCODECPTA_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA_MUL.OnArgument (S : String ) ;
var x       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin

  Inherited ;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
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

  //Gesion des Paramètres société
  FamCptaArt.Visible    := GetParamSocSecur('SO_GCVENTCPTAART', False);
  FamCptaTiers.Visible  := GetParamSocSecur('SO_GCVENTCPTATIERS', False);
  FamCptaAff.Visible    := GetParamSocSecur('SO_GCVENTCPTAAFF', False);

  TCptaArt.visible      := FamCptaArt.Visible;
  TCptaTiers.Visible    := FamCptaTiers.Visible;
  TCptaAff.Visible      := FamCptaAff.Visible;

end ;

procedure TOF_BTCODECPTA_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCODECPTA_MUL.GetObjects;
begin

  if Assigned(GetControl('Fliste'))             then Grille   := THDBGrid(ecran.FindComponent('Fliste'));

  if Assigned(GetControl('BInsert'))            then BInsert  := TToolBarButton97(Getcontrol('BInsert'));

  if Assigned(GetControl('GCP_VENTEACHAT'))     then VenteAchat  := THValComboBox(GetControl('GCP_VENTEACHAT'));
  if Assigned(GetControl('GCP_COMPTAARTICLE'))  then FamCptaArt  := THValComboBox(GetControl('GCP_COMPTAARTICLE'));
  if Assigned(GetControl('GCP_COMPTATIERS'))    then FamCptaTiers:= THValComboBox(GetControl('GCP_COMPTATIERS'));
  if Assigned(GetControl('GCP_COMPTAAFFAIRE'))  then FamCptaAff  := THValComboBox(GetControl('GCP_COMPTAAFFAIRE'));
  if Assigned(GetControl('GCP_ETABLISSEMENT'))  then VenteAchat  := THValComboBox(GetControl('GCP_ETABLISSEMENT'));
  if Assigned(GetControl('GCP_REGIMETAXE'))     then RegimeTaxe  := THValComboBox(GetControl('GCP_REGIMETAXE'));
  if Assigned(GetControl('GCP_FAMILLETAXE'))    then FamTaxe     := THValComboBox(GetControl('GCP_FAMILLETAXE'));

  if Assigned(GetControl('TGCP_COMPTAARTICLE')) then TCptaArt   := THLabel(GetControl('TGCP_COMPTAARTICLE'));
  if Assigned(GetControl('TGCP_COMPTATIERS'))   then TCptaTiers := THLabel(GetControl('TGCP_COMPTATIERS'));
  if Assigned(GetControl('TGCP_COMPTAAFFAIRE')) then TCptaAff   := THLabel(GetControl('TGCP_COMPTAAFFAIRE'));

end;

procedure TOF_BTCODECPTA_MUL.SetScreenEvents;
begin

  Grille.OnDblClick := GrilleOnDblclick;

  BInsert.OnClick   := InsertOnClick;

end;

Procedure TOF_BTCODECPTA_MUL.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if Valeur='CREATION'          then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end
  else if Champ = 'RECH' then
  begin
    if valeur='X'                 then ModeRecherche:=true;
  end;  

end;

Procedure TOF_BTCODECPTA_MUL.GrilleOnDblclick(Sender : TObject);
Var Argument : String;
    Rang     : Integer;
begin

  Rang := Grille.Datasource.Dataset.FindField('GCP_RANG').AsInteger;

  if Moderecherche then
  begin
  	TFMul(Ecran).Retour := IntToStr(Rang);
    TFMul(Ecran).Close;
  end
  else
  begin
    //if Rang = 0 then
    //  Argument := 'ACTION=CREATION'
    //else
    Argument := 'RANG=' + IntToStr(Rang) + ';ACTION=MODIFICATION';

    AGLLanceFiche('BTP','BTCODECPTA','','',Argument);

    TFMul(Ecran).BChercheClick (Self);
  end;

end;

procedure TOF_BTCODECPTA_MUL.InsertOnclick(Sender: TObject);
Var Argument : String;
begin

  Argument := 'ACTION=CREATION';

  AGLLanceFiche('BTP','BTCODECPTA','','',Argument);

  TFMul(Ecran).BChercheClick (Self);

end;

Initialization
  registerclasses ( [ TOF_BTCODECPTA_MUL ] ) ;
end.

