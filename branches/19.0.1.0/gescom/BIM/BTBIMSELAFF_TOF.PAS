{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/07/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTBIMSELAFF ()
Mots clefs ... : TOF;BTBIMSELAFF
*****************************************************************}
Unit BTBIMSELAFF_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     uDbxDataSet, 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     UTofBTCodeAffaire,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97, 
     UTOF ; 

Type
  TOF_BTBIMSELAFF = Class (TOF_BTCODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
  private
    procedure BSTOPCLick (Sender : TObject);
  end ;

Implementation

procedure TOF_BTBIMSELAFF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMSELAFF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMSELAFF.OnUpdate ;
begin
  Inherited ;
  if THEdit(GetControl('AFF_AFFAIRE')).text <> '' then LaTOB.SetString('AFFAIRE',THEdit(GetControl('AFF_AFFAIRE')).text);
end ;

procedure TOF_BTBIMSELAFF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMSELAFF.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolbarButton97 (GetControl('BSTOP')).onClick := BSTOPCLick;
end ;

procedure TOF_BTBIMSELAFF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMSELAFF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMSELAFF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMSELAFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  Aff := THEdit(Getcontrol('AFF_AFFAIRE'));
  Aff0 := THEdit(Getcontrol('AFF_AFFAIRE0'));
  Aff1 := THEdit(Getcontrol('AFF_AFFAIRE1'));
  Aff2 := THEdit(Getcontrol('AFF_AFFAIRE2'));
  Aff3 := THEdit(Getcontrol('AFF_AFFAIRE3'));
  Aff4 := THEdit(Getcontrol('AFF_AVENANT'));
end;

procedure TOF_BTBIMSELAFF.BSTOPCLick(Sender: TObject);
begin
  LaTOB.SetBoolean('ANNULATION',True);
  Ecran.Close;
end;

Initialization
  registerclasses ( [ TOF_BTBIMSELAFF ] ) ; 
end.
