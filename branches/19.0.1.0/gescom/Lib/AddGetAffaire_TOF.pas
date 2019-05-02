{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMWS ()
Mots clefs ... : TOF;BTPARAMWS
*****************************************************************}
Unit AddGetAffaire_TOF ;

Interface

Uses
  StdCtrls
  , Controls
  , Classes
  {$IFNDEF EAGLCLIENT}
  , db
  , uDbxDataSet
  , mul
  , FE_main
  , uTob
  {$ELSE EAGLCLIENT}
  , eMul
  , MaineAGL
  {$ENDIF EAGLCLIENT}
  , forms
  , sysutils
  , ComCtrls
  , HCtrls
  , HMsgBox
  , UTOF
  , HTB97
  ;

function BTLanceFicheAddGetAffaire(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_AddGetAffaire = Class (TOF)
  private
    CodeTiers : string;

    procedure bValider_OnClick(Sender : Tobject);
    procedure bChoix_OnClick(Sender : Tobject);

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

uses
  UtilBTPVerdon
  , ParamSoc
  , HEnt1
  , TntStdCtrls
  , utilPGI
  , Vierge
  ;

function BTLanceFicheAddGetAffaire(Nat, Cod : String ; Range,Lequel,Argument : string) : string;
begin
  if (Nat <> '') and (Cod <> '') then
  begin
    V_PGI.ZoomOle := True;
    Result        := AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
    V_PGI.ZoomOle := False;
  end else
    Result := '';
end;

procedure TOF_AddGetAffaire.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AddGetAffaire.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AddGetAffaire.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AddGetAffaire.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AddGetAffaire.OnArgument(S : String ) ;
begin
  Inherited ;
  TToolbarButton97(GetControl('BValider')).OnClick := bValider_OnClick;
  THRadioGroup(GetControl('BCHOIX')).OnClick       := bChoix_OnClick;
  CodeTiers := GetArgumentString(S, 'TIERS');
  TFVierge(Ecran).Caption := Format('%s - %s', [CodeTiers, TFVierge(Ecran).Caption]);
end;

procedure TOF_AddGetAffaire.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AddGetAffaire.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AddGetAffaire.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AddGetAffaire.bValider_OnClick(Sender: Tobject);
var
  Choix : string;
begin
  case THRadioGroup(GetControl('BCHOIX')).ItemIndex of
    {Annuler}  0 : Choix := 'A';
    {Créer}    1 : Choix := 'C';
    {Associer} 2 : Choix := 'L';
  end;
  TFVierge(Ecran).Retour := Format('%s;%s', [Choix, THEdit(GetControl('LIBELLE')).Text]);
end;

procedure TOF_AddGetAffaire.bChoix_OnClick(Sender: Tobject);
begin
  TGroupBox(GetControl('GLABEL')).Visible := (THRadioGroup(GetControl('BCHOIX')).ItemIndex = 1); 
end;

Initialization
  registerclasses ( [ TOF_AddGetAffaire ] ) ;
end.                                                     

