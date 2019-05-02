{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 11/03/2019
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPREVISION_MUL ()
Mots clefs ... : TOF;BTPREVISION_MUL
*****************************************************************}
Unit BTPREVISION_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, HDB, 
     uDbxDataSet, 
     mul,
     Fe_Main,
{$else}
     eMul, 
{$ENDIF}
     uTob, 
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     UentCommun,
     UTofAfBaseCodeAffaire ;

Type

  TOF_BTPREVISION_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
  private
    Fliste : THDBGrid;
    procedure FlisteDblClick(Sender: TObject);
  end ;

Implementation

uses AglInit, UAvtVerdon;

procedure TOF_BTPREVISION_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVISION_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVISION_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVISION_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVISION_MUL.OnArgument (S : String ) ;
begin
  fMulDeTraitement  := true;
  Inherited ;
  Fliste := THDbGrid (GetControl('FLISTE'));
  Fliste.OnDblClick := FlisteDblClick;
  Ecran.Caption := 'Saisie des avancements';
  UpdateCaption(Ecran);
end ;

procedure TOF_BTPREVISION_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVISION_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVISION_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVISION_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,Tiers_: THEdit);
begin
  Aff:=THEdit(GetControl('GP_AFFAIRE'));
  Aff0:=THEdit(GetControl('AFFAIRE0'));
  Aff1:=THEdit(GetControl('GP_AFFAIRE1'));
  Aff2:=THEdit(GetControl('GP_AFFAIRE2'));
  Aff3:=THEdit(GetControl('GP_AFFAIRE3'));
  Aff4:=THEdit(GetControl('GP_AVENANT'));
  Tiers:=THEdit(GetControl('GP_TIERS'));
end;

procedure TOF_BTPREVISION_MUL.FlisteDblClick(Sender: TObject);
var Cledoc : r_cledoc;
    DateApplic : TdateTime;
    Affaire : string;
begin
  DateApplic := IDate1900;
  fillchar(cledoc,sizeof(cledoc),#0);
  TRY
    Cledoc.NaturePiece  := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
    Cledoc.Souche := Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
    Cledoc.NumeroPiece := Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
    Cledoc.Indice := Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
    Affaire := Fliste.datasource.dataset.FindField('GP_AFFAIRE').AsString;
    if not AVCVerdon.ExisteDejaUn (cledoc) then
    begin
      if not AVCVerdon.DemandeDateApplic (Cledoc,DateApplic) then Exit;
      AVCVerdon.SaisieAvancements (Affaire,Cledoc,DateApplic,0);
    end else
    begin
      AVCVerdon.AfficheListeAvancements(Affaire,Cledoc);
    end;
  FINALLY
    Refreshdb;
  END;

end;

Initialization
  registerclasses ( [ TOF_BTPREVISION_MUL ] ) ;
end.
