{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/11/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTEXTEMEMO ()
Mots clefs ... : TOF;BTTEXTEMEMO
*****************************************************************}
Unit BTTexteMemo_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     Fiche,
     FichList,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     HRichOle,
     Windows,
     Math,
     Grids,
     HTB97,
     UtilsGrille,
     UTob,
     Vierge,
     HSysMenu,
     HRichEdt,
     UTOF;

Type
  TOF_BTTEXTEMEMO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
private
  //
  BPremier    : TToolbarButton97;
  BSuivant    : TToolbarButton97;
  BDernier    : TToolbarButton97;
  BPrecedent  : TToolbarButton97;
  //
  BInsert     : TToolbarButton97;
  BDelete     : TToolbarButton97;
  BValide     : TToolbarButton97;
  BAnnuler    : TToolbarButton97;
  BFermer     : TToolbarButton97;
  //
  Grille      : THGrid;
  GTxtMemo    : TGestionGS;
  //
  Code        : THEdit;
  Libelle     : THEdit;
  BlocNote    : THRichEditOLE;
  MeMoSave    : THRichEditOLE;
  //
  TobTxtMemo  : TOB;
  TobLigMemo  : TOB;
  //
  Action      : TActionFiche;
  //
    procedure AfficheBoutons(OK_Affiche: Boolean);
    procedure BannulerOnclick(Sender: Tobject);
    procedure BDernierOnclick(Sender: TObject);
    procedure BlocNoteOnchange(Sender: Tobject);
    procedure BPrecedentOnclick(Sender: TObject);
    procedure BPremierOnclick(Sender: TObject);
    procedure BSuivantOnclick(Sender: TObject);
    procedure ChargeEcranToTob;
    procedure ChargeZoneEcran(Arow: Integer);
    procedure CreateTob;
    procedure GetObjects;
    procedure GrilleExit(Sender: TObject);
    procedure GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrilleRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
    procedure InitGrille;
    procedure LibelleOnchange(Sender: Tobject);
    procedure RAZZoneEcran;
    procedure SetScreenEvents;
  //
  end ;

Implementation
uses  BTPUtil,
      RtfCounter;

procedure TOF_BTTEXTEMEMO.OnNew ;
begin
  Inherited ;

  Action := taCreat;

  RAZZoneEcran;

  Code.enabled := True;

  TobLigMemo := TOB.Create('COMMENTAIRE', TobTxtMemo, -1);

  AfficheBoutons(False);

  Grille.Enabled := false;

  Code.SetFocus;

end ;

procedure TOF_BTTEXTEMEMO.OnDelete ;
var NumLig : Integer;
begin
  Inherited ;

  if PGIAsk('Confirmez-vous la suppression du texte mémorisé', 'Textes mémorisés')= MrNo then Exit;

  NumLig := Grille.Row;

  TobLigMemo.DeleteDB;
  TobTxtMemo.detail.delete(NumLig-1);

  //chargement de la grille
  GTxtMemo.ChargementGrille;

  if TobTxtMemo.detail.count = 0 then
  Begin
    Grille.ClearSelected;
    Onnew;
    Exit;
  end
  else if TobTxtMemo.detail.count < NumLig then
    NumLig := TobTxtMemo.detail.count;

  Grille.Row := NumLig;
  ChargeZoneEcran(Grille.Row);

end ;

procedure TOF_BTTEXTEMEMO.OnUpdate ;
var NumLig : Integer;
begin
  Inherited ;

  if action = TaConsult then exit;

  NumLig := Grille.Row;

  ChargeEcranToTob;

  TobLigMemo.SetAllModifie(True);
  TobLigMemo.InsertOrUpdateDB(False);

  AfficheBoutons(True);

  //chargement de la grille
  GTxtMemo.ChargementGrille;

  If action = TaCreat then NumLig := Grille.RowCount -1;

  Action := TaConsult;
  Grille.Enabled := True;

  //On se positionne sur la première ligne de la grille
  Grille.Row := NumLig;
  ChargeZoneEcran(Grille.Row);

end ;

procedure TOF_BTTEXTEMEMO.OnLoad ;
Var StSQL : String;
    QQ    : TQuery;
begin
  Inherited ;

  //chargement du Matériel
  StSQL := 'SELECT * FROM COMMENTAIRE';
  QQ := OpenSQL(StSQL, False);

  If not QQ.Eof then
  begin
    TobTxtMemo.LoadDetailDB('COMMENTAIRE', '', '', QQ, False);
    //
    Ferme(QQ);

    //chargement de la grille
    GTxtMemo.ChargementGrille;

    //On se positionne sur la première ligne de la grille
    Grille.Row := 1;
    ChargeZoneEcran(Grille.row);
  end
  else
  begin
    Ferme(QQ);
    //Se mettre directement en création
    OnNew;
  end;

end ;

procedure TOF_BTTEXTEMEMO.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  TFvierge(Ecran).FormResize := True;

  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  InitGrille;
  //
  BAnnuler.Left := BFermer.Left;
  BAnnuler.Top  := BFermer.Top;
  BAnnuler.Visible := False;
  BFermer.visible  := True;
  //
  //Gestion des évènement des zones écran
  SetScreenEvents;

  //Remise à blanc des zone ecran
  RAZZoneEcran;

  //
  Action := taConsult;

end ;

procedure TOF_BTTEXTEMEMO.GetObjects;
begin

  Code        := THEdit(Getcontrol('GCT_CODE'));
  Libelle     := THEdit(Getcontrol('GCT_LIBELLE'));
  BlocNote    := THRichEditOLE(Getcontrol('GCT_CONTENU'));

  Grille      := THGrid(Getcontrol('GCTGRILLE'));

  BPremier    := TToolbarButton97(Getcontrol('BPremier'));
  BPrecedent  := TToolbarButton97(Getcontrol('BPrecedent'));
  BSuivant    := TToolbarButton97(Getcontrol('BSuivant'));
  BDernier    := TToolbarButton97(Getcontrol('BDernier'));
  //
  BDelete     := TToolbarButton97(Getcontrol('BDelete'));
  BInsert     := TToolbarButton97(Getcontrol('BInsert'));
  Bvalide     := TToolbarButton97(Getcontrol('Bvalide'));
  BAnnuler    := TToolbarButton97(GetControl('BAnnuler'));
  BFermer     := TToolbarButton97(GetControl('BFerme'));
  //
end;

procedure TOF_BTTEXTEMEMO.CreateTob;
begin

  TobTxtMemo := TOB.Create('COMMENTAIRE',nil,-1);

end;

procedure TOF_BTTEXTEMEMO.InitGrille;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GTxtMemo := TGestionGS.Create;

  GTxtMemo.Ecran := TFVierge(Ecran);
  GTxtMemo.GS    := Grille;
  GTxtMemo.TOBG  := TobTxtMemo;

  GTxtMemo.NomListe  := 'BTCOMMENTAIRE';

  GTxtMemo.ChargeInfoListe;

  if GTxtMemo.NomListe <> '' then GTxtMemo.DessineGrille

end;

Procedure TOF_BTTEXTEMEMO.SetScreenEvents;
begin

  Grille.OnRowEnter   := GrilleRowEnter;
  Grille.OnRowExit    := GrilleRowExit;
  Grille.OnExit       := GrilleExit;
  //
  Libelle.OnChange    := LibelleOnchange;
  BlocNote.OnChange   := BlocNoteOnchange;
  //
  BPremier.OnClick    := BPremierOnclick;
  BDernier.OnClick    := BDernierOnclick;
  BSuivant.OnClick    := BSuivantOnclick;
  BPrecedent.OnClick  := BPrecedentOnclick;
  //
  BAnnuler.OnClick    := BannulerOnclick;
end;

procedure TOF_BTTEXTEMEMO.RAZZoneEcran;
begin

  Code.Text     := '';
  Libelle.Text  := '';

  BlocNote.Clear;

end;

procedure TOF_BTTEXTEMEMO.OnClose ;
begin
  Inherited ;

  FreeAndNil(TobTxtMemo);

end ;

procedure TOF_BTTEXTEMEMO.BannulerOnclick (Sender : Tobject) ;
begin
  Inherited ;

  if Action = taCreat then
  begin
    if PGIAsk('Désirez-vous abandonner la création de votre texte mémorisé', 'Textes mémorisés')= MrNo then Exit;
    TobLigMemo.ClearDetail;
    TobTxtMemo.detail.delete(TobtxtMemo.Detail.count-1);
    Grille.enabled := true;
  end
  else
  begin
    if PGIAsk('Désirez-vous abandonner vos modifications', 'Textes mémorisés')= MrNo then Exit;
  end;

  Action := taConsult;

  AfficheBoutons(True);

  ChargeZoneEcran(Grille.row);

end ;

procedure TOF_BTTEXTEMEMO.ChargeZoneEcran(Arow : Integer);
begin

  if TobTxtMemo = nil then Exit;
  If TobTxtMemo.detail.Count =0 then Exit;

  FreeAndNil(MemoSave);
  BlocNote.clear;

  TobLigMemo := TobTxtMemo.detail[Arow-1];

  Code.enabled := False;

  Code.Text     := TobLigMemo.GetString('GCT_CODE');
  Libelle.Text  := TobLigMemo.GetString('GCT_LIBELLE');

  MemoSave         := THRichEditOLE.Create(Application);
  MemoSave.Parent  := Ecran;
  MeMoSave.Visible := False;
  StringtoRich(MemoSave,TobLigMemo.GetString('GCT_CONTENU'));

  StringtoRich(BlocNote,TobLigMemo.GetString('GCT_CONTENU'));

  AppliqueFontDefaut(BlocNote);

end;

procedure TOF_BTTEXTEMEMO.ChargeEcranToTob;
begin

  TobLigMemo.PutValue('GCT_CODE', code.text);
  TobLigMemo.PutValue('GCT_LIBELLE', Libelle.text);
  TobLigMemo.PutValue('GCT_TYPECOMMENT', '');

  if (Length(BlocNote.text) <> 0) and (BlocNote.text <> sLineBreak) then
  begin
    TobLigMemo.PutValue('GCT_CONTENU', ExRichToString(BlocNote));
  end
  else
    TobLigMemo.PutValue('GCT_CONTENU', '');

end;

Procedure TOF_BTTEXTEMEMO.LibelleOnchange(Sender : Tobject);
begin

  if Action = taCreat then Exit;

  if TobLigMemo = nil then exit;

  if Libelle.Text = '' then Exit;

  if Libelle.Text <> TobLigMemo.GetString('GCT_LIBELLE') then
  begin
    Action := taModif;
    AfficheBoutons(False);
  end;

end;

Procedure TOF_BTTEXTEMEMO.BlocNoteOnchange(Sender : Tobject);
begin

  if Action = taCreat then Exit;

  if TobLigMemo = nil then exit;

  If BlocNote.Text = '' then Exit;

  if BlocNote.Text <> MemoSave.text then
  begin
    Action            := taModif;
    AfficheBoutons(False);
  end;

end;

Procedure TOF_BTTEXTEMEMO.AfficheBoutons(OK_Affiche : Boolean);
Begin

    BInsert.Visible   := Ok_Affiche;
    BDelete.Visible   := Ok_Affiche;
    BPremier.Visible  := Ok_Affiche;
    BDernier.Visible  := Ok_Affiche;
    BSuivant.Visible  := Ok_Affiche;
    BPrecedent.Visible:= Ok_Affiche;

    BAnnuler.Visible  := Not Ok_Affiche;

    Bfermer.Visible   := Ok_Affiche;

end;

procedure TOF_BTTEXTEMEMO.GrilleExit(Sender: TObject);
begin

end;

procedure TOF_BTTEXTEMEMO.GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

  //chargement des zones à l'écran
  ChargeZoneEcran(Ou)

end;

procedure TOF_BTTEXTEMEMO.GrilleRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

  if Action = taModif then
  begin
    if PGIAsk('Désirez-vous abandonner vos modifications', 'Textes mémorisés') = MrNo then
    begin
      Grille.row := Ou;
    end
    else
    begin
      Action := taConsult;
      AfficheBoutons(True);
    end;
  end;

end;

procedure TOF_BTTEXTEMEMO.BPremierOnclick(Sender : TObject);
begin

  Grille.Row := 1;
  ChargeZoneEcran(Grille.row);

end;

procedure TOF_BTTEXTEMEMO.BDernierOnclick(Sender : TObject);
begin

  Grille.Row := Grille.RowCount-1;
  ChargeZoneEcran(Grille.row);

end;

procedure TOF_BTTEXTEMEMO.BSuivantOnclick(Sender : TObject);
Var Numlig  : Integer;
begin

  Numlig := Grille.Row;

  if Numlig = Grille.RowCount-1 then
    Numlig := 1
  else
    Numlig := Numlig + 1;

  Grille.Row := NumLig;
  ChargeZoneEcran(Grille.row);

end;

procedure TOF_BTTEXTEMEMO.BPrecedentOnclick(Sender : TObject);
Var Numlig  : Integer;
begin

  Numlig := Grille.Row;

  if Numlig = 1 then
    Numlig :=  Grille.RowCount-1
  else
    Numlig := Numlig - 1;

  Grille.Row := NumLig;
  ChargeZoneEcran(Grille.row);

end;


Initialization
  registerclasses ( [ TOF_BTTEXTEMEMO ] ) ;
end.

