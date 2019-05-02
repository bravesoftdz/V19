{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/05/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTVENTILCOLLMUL ()
Mots clefs ... : TOF;BTVENTILCOLLMUL
*****************************************************************}
Unit BTVENTILCOLLMUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     DBGrids,
     mul,
{$else}
     eMul,
{$ENDIF}
     fe_main, 
     HDB,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF ;

Type
  TOF_BTVENTILCOLLMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    fliste : THDBGrid;
    procedure FlisteDblClick (Sender : TObject);
    procedure BinsertClick (Sender : TObject);
    procedure RefreshGrid;
  end ;

Implementation

procedure TOF_BTVENTILCOLLMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLLMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLLMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLLMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLLMUL.OnArgument (S : String ) ;
begin
  Inherited ;
  Fliste := THDbGrid (GetControl('FLISTE'));
  Fliste.OnDblClick := FlisteDblClick;
  TToolbarButton97 (GetControl('Binsert')).onclick := BinsertClick;
end ;

procedure TOF_BTVENTILCOLLMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLLMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLLMUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLLMUL.FlisteDblClick(Sender: TObject);
var TypeV,ComptaArt,ComptaTiers,ComptaAffaire,Etablissement,Regime,CodePort : string;
    TheParam :string;
begin
  //
  if FListe.datasource.DataSet.RecordCount = 0  then exit;
  //
  TypeV := Fliste.datasource.dataset.FindField('BVC_NATUREV').AsString;
  CodePort := Fliste.datasource.dataset.FindField('BVC_CODEPORT').AsString;
  ComptaArt := Fliste.datasource.dataset.FindField('BVC_COMPTAARTICLE').AsString;
  ComptaTiers := Fliste.datasource.dataset.FindField('BVC_COMPTATIERS').AsString;
  ComptaAffaire := Fliste.datasource.dataset.FindField('BVC_COMPTAAFFAIRE').AsString;
  Etablissement := Fliste.datasource.dataset.FindField('BVC_ETABLISSEMENT').AsString;
  Regime := Fliste.datasource.dataset.FindField('BVC_REGIMETAXE').AsString;
  //
  TheParam := Format('BVC_NATUREV=%s;BVC_CODEPORT=%s;BVC_COMPTAARTICLE=%s;BVC_COMPTATIERS=%s;BVC_COMPTAAFFAIRE=%s;BVC_ETABLISSEMENT=%s;BVC_REGIMETAXE=%s',[TypeV,CodePort,ComptaArt,ComptaTiers,ComptaAffaire,Etablissement,Regime]);
  AGLLanceFiche ('BTP','BTVENTILCOLL','','','ACTION=MODIFICATION;'+TheParam);
  RefreshGrid;
end;

procedure TOF_BTVENTILCOLLMUL.BinsertClick(Sender: TObject);
begin
  AGLLanceFiche ('BTP','BTVENTILCOLL','','','ACTION=CREATION');
  RefreshGrid;
end;

procedure TOF_BTVENTILCOLLMUL.RefreshGrid;
begin
  TToolbarButton97(GetControl('BCHerche')).Click;
end;

Initialization
  registerclasses ( [ TOF_BTVENTILCOLLMUL ] ) ; 
end.
