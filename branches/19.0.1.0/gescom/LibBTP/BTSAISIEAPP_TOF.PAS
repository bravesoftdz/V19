{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/09/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISIEAPP ()
Mots clefs ... : TOF;BTSAISIEAPP
*****************************************************************}
Unit BTSAISIEAPP_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HTB97,
     HMsgBox,
     HRichOLE,
     HRichEdt,
     UTOF ;

Type
  TOF_BTSAISIEAPP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBAPP : TOB;
    //
    Action      : TActionFiche;
    //
    Appel       : THEdit;
    Appel0      : THEdit;
    Appel1      : THEdit;
    Appel2      : THEdit;
    Appel3      : THEdit;
    Avenant     : THEdit;
    //
    CodeTiers   : THLabel;
    NomTiers    : THLabel;
    DateSouhait : THLabel;
    //
    Descriptif  : THRichEditOLE;
    //
    BSelect     : TToolbarButton97;
    BEfface     : TToolbarButton97;
    //
    procedure AffEffaceOnClick(Sender: TObject);
    procedure AffSelectOnClick(Sender: TObject);
    procedure AppelOnExit(Sender: TObject);
    procedure ChargeZoneEcran;
    procedure CreateTOB;
    procedure GetObjects;
    procedure SetScreenEvents;
    procedure ChargeZoneTOB;
    //
  end ;

Implementation
uses AffaireUtil,AppelsUtil,FACTUtil;

procedure TOF_BTSAISIEAPP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnUpdate ;
begin
  Inherited ;
  
  if not ControleSaisieOK (Appel1.text, Appel2.text, Appel3.text) then
  begin
    PgiInfo('Merci de renseigner le code intervention');
    ecran.ModalResult := mrNone;
    exit;
  end;

  ChargeZoneTob;

end ;

procedure TOF_BTSAISIEAPP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnArgument (S : String ) ;
begin
  Inherited ;

  //Chargement des zones ecran dans des zones programme
  GetObjects;
  SetScreenEvents;

  SetControlText('AFF_APPEL0','A');

  CreateTOB;

  ChargeZoneEcran;

  ChargeClefAppel (Appel0,Appel1,Appel2,Appel3,Avenant, BSelect, Action, Appel.text, False);

end ;

procedure TOF_BTSAISIEAPP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.GetObjects;
begin

  Appel   := THEdit(Getcontrol('AFF_APPEL'));
  Appel0  := THEdit(Getcontrol('AFF_APPEL0'));
  Appel1  := THEdit(Getcontrol('AFF_APPEL1'));
  Appel2  := THEdit(Getcontrol('AFF_APPEL2'));
  Appel3  := THEdit(Getcontrol('AFF_APPEL3'));
  Avenant := THEdit(Getcontrol('AFF_AVENANT'));
  //
  Descriptif  := THRichEditOLE(GetControl('BAA_DESCRIPTIF'));
  //
  CodeTiers   := THLabel(Getcontrol('CODETIERS'));
  NomTiers    := THLabel(Getcontrol('NOMTIERS'));
  DateSouhait := THLabel(Getcontrol('DATYESQOUHAIT'));
  //
  BSelect     := TToolbarButton97(Getcontrol('BSELECT'));
  BEfface     := TToolbarButton97(Getcontrol('BEFFACE'));

end;

Procedure TOF_BTSAISIEAPP.CreateTOB;
begin

  TOBAPP := LaTOB;

end;

procedure TOF_BTSAISIEAPP.SetScreenEvents;
begin
  //
  Appel1.Onexit   := AppelOnExit;
  Appel2.OnExit   := AppelOnExit;
  Appel3.OnExit   := AppelOnExit;
  //
  BSelect.OnClick := AffSelectOnClick;
  BEfface.OnClick := AffEffaceOnclick;
  //
end;

Procedure TOF_BTSAISIEAPP.ChargeZoneEcran;
begin

  Appel0.Text   := 'W';
  Appel1.Text   := '';
  Appel2.Text   := '';
  Appel3.Text   := '';
  Avenant.Text  := '';

  CodeTiers.caption   := 'Code tiers : ' +  TOBAPP.GetString('BAA_TIERS');
  NomTiers.Caption    := 'Nom Tiers : ' + TOBAPP.GetString('BAA_LIBADR');
  DateSouhait.Caption := 'Date Souhaitée : ' + TOBAPP.GetString('BAA_DATESOUHAIT');
  //
  //StringToRich(Descriptif, TOBAPP.GetValue('BAA_DESCRIPTIF'));
  Descriptif.text     := TOBAPP.GetString('BAA_DESCRIPTIF');
  //
end;

Procedure TOF_BTSAISIEAPP.ChargeZoneTOB;
begin

  TOBAPP.PutValue('AFFAIRE',  Appel.text);
  TOBAPP.PutValue('AFF_AFFAIRE0', Appel0.text);
  TOBAPP.PutValue('AFF_AFFAIRE1', Appel1.text);
  TOBAPP.PutValue('AFF_AFFAIRE2', Appel2.text);
  TOBAPP.PutValue('AFF_AFFAIRE3', Appel3.text);
  TOBAPP.PutValue('AFF_AVENANT',  Avenant.text);

  TOBAPP.SetString('SAISIEOK','X')

end;

Procedure TOF_BTSAISIEAPP.AppelOnExit(Sender : TObject);

Var C0		: String;
    C1		: String;
    C2		: String;
    C3 		: String;
    C4	  : String;
    Tiers : String;
    IP    : Integer;
begin

	C0 := Appel0.text;
	C1 := Appel1.Text;
	C2 := Appel2.text;
	C3 := Appel3.text;
	C4 := Avenant.Text;
  //
	Tiers := TOBAPP.GetString('BAA_TIERS');

	Appel.text := CodeAffaireRegroupe(C0, C1, C2, C3, C4, taCreat, True, True,false);

  //control existence
  if Appel.text = '' then
  begin
    exit;
  end;

  SetControlText('AFF_APPEL0', C0);
  SetControlText('AFF_APPEL1', C1);
  SetControlText('AFF_APPEL2', C2);
  SetControlText('AFF_APPEL3', C3);
  SetControlText('APP_AVENANT', C4);

  Appel.text := DechargeCleAffaire(Appel0, Appel1, Appel2, Appel3, Avenant, Tiers, Action, False, True, false, IP);

end;

Procedure TOF_BTSAISIEAPP.AffEffaceOnClick(Sender : TObject);
begin

  Appel0.Text := 'W';
  Appel1.Text := '';
  Appel2.Text := '';
  Appel3.Text := '';
  Avenant.Text  := '';

end;

Procedure TOF_BTSAISIEAPP.AffSelectOnClick(Sender : TObject);
Var StChamps  : String;
begin

  StChamps  := Appel.text;

  if GetAffaireEnteteSt(Appel0, Appel1, Appel2, Appel3, Avenant, nil, StChamps, false, false, false, True, true,'') then Appel.text := StChamps;

  ChargeClefAppel (Appel0,Appel1,Appel2,Appel3,Avenant, BSelect, Action, Appel.text, False);

end;

Initialization
  registerclasses ( [ TOF_BTSAISIEAPP ] ) ;
end.
