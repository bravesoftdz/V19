{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/03/2019
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BAVCVERDON_SEL ()
Mots clefs ... : TOF;BAVCVERDON_SEL
*****************************************************************}
Unit BAVCVERDON_SEL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     HTB97, 
{$IFNDEF EAGLCLIENT}
     db, 
     uDbxDataSet, 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     HDB,
     uTob, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     uTOFComm, 
     UTOF ; 

Type
  TOF_BAVCVERDON_SEL = Class (tTOFComm)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Procedure InsClick (Sender  :Tobject);
    procedure FlisteDblClick (Sender : Tobject);
  end ;

Implementation
uses UAvtVerdon,uEntCommun,ExtCtrls, Spin;

procedure TOF_BAVCVERDON_SEL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BAVCVERDON_SEL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BAVCVERDON_SEL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BAVCVERDON_SEL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BAVCVERDON_SEL.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolbarButton97 (GetControl('Binsert')).OnClick := InsClick;
  TToolbarButton97 (GetControl('fListe')).OnDblClick := FlisteDblClick;
end ;

procedure TOF_BAVCVERDON_SEL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BAVCVERDON_SEL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BAVCVERDON_SEL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BAVCVERDON_SEL.InsClick(Sender: Tobject);
var cledoc : r_cledoc;
    dateApplic : TDateTime;
    affaire : string;
begin
  cledoc.naturepiece := THValComboBox(GetControl('BVV_NATUREPIECEG')).Value;
  cledoc.Souche := THValComboBox(GetControl('BVV_SOUCHE')).Value;
  cledoc.NumeroPiece  := THSpinEdit(GetControl('BVV_NUMERO')).Value;
  cledoc.Indice  := THSpinEdit(GetControl('BVV_INDICEG')).Value;
  Affaire := THEdit(GetControl('AFFAIRE')).Text;
  //
  if not AVCVerdon.DemandeDateApplic (Cledoc,DateApplic) then Exit;
  AVCVerdon.SaisieAvancements (Affaire,Cledoc,DateApplic,0);
  RefreshDb;
end;

procedure TOF_BAVCVERDON_SEL.FlisteDblClick(Sender: Tobject);
var cledoc : r_cledoc;
    dateApplic : TDateTime;
    affaire : string;
    NumAvc : Integer;
begin
  cledoc.naturepiece := THValComboBox(GetControl('BVV_NATUREPIECEG')).Value;
  cledoc.Souche := THValComboBox(GetControl('BVV_SOUCHE')).Value;
  cledoc.NumeroPiece  := THSpinEdit(GetControl('BVV_NUMERO')).Value;
  cledoc.Indice  := THSpinEdit(GetControl('BVV_INDICEG')).Value;
  Affaire := THEdit(GetControl('AFFAIRE')).Text;
  DateApplic := THDBGrid(GetControl('Fliste')).datasource.dataset.FindField('BVV_DATEAPPLIC').AsDateTime;
  NumAvc := THDBGrid(GetControl('Fliste')).datasource.dataset.FindField('BVV_NUMAVC').AsInteger;
  //
  AVCVerdon.SaisieAvancements (Affaire,Cledoc,DateApplic,NumAvc);
  RefreshDb;
end;

Initialization
  registerclasses ( [ TOF_BAVCVERDON_SEL ] ) ;
end.
