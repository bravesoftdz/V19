{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/03/2019
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISAVCVERDON ()
Mots clefs ... : TOF;BTSAISAVCVERDON
*****************************************************************}
Unit BTSAISAVCVERDON_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     uTob,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul,
{$else}
     eMul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uEntCommun,
     Vierge,
     UTOF,
     Graphics,
     Grids,
     ImgList,
     HTB97,
     UTofAfBaseCodeAffaire,
     UFactgestionAff ;

const
    STChaine = 'WW;GL_TYPEARTICLE;GL_CODEARTICLE;GL_LIBELLE;BLF_QTEMARCHE;BLF_QTECUMULEFACT;BLF_POURCENTAVANC';

Type
  TOF_BTSAISAVCVERDON = Class (TOF_AFBASECODEAFFAIRE)
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
    fAction : TActionFiche;
    stCell :string;
    cledoc : r_cledoc;
    DateApplic : TDateTime;
    NumAvc : Integer;
    TOBAVC,TOBLigFac : TOB;
    Affaire : string;
    GS : THGrid;
    ImTypeArticle : TImageList;
    SG_LIBELLE : Integer;
    SG_POURCENTAVANC : Integer;
    SG_QTECUMULE : Integer;
    fGestionAff: TAffichageDoc;
    Fparcours : TToolWindow97;
    TVParag : TTreeView;

    procedure ConstitueGrid;
    procedure ChargeLagrid;
    procedure SetEventGrid (State : boolean);
    procedure GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ChargeListImage;
    function RecupTypeGraph(TOBL: TOB): integer;
    procedure ZoneSuivanteOuOk(Grille: THgrid; var ACol, ARow: integer; var Cancel: boolean);
    function ZoneAccessible(Grille: THgrid; var ACol,ARow: integer): boolean;
    procedure AfficheLaligne(ARow : integer);
    function ControleSaisieOk : Boolean;
    procedure ValideLaSaisie;
    procedure BdeleteClick (Sender : TObject);
    function IsArticle(TOBL: TOB): boolean;
    function IsCommentaire(TOBL: TOB): boolean;
    function IsDebutParagraphe(TOBL: TOB): boolean;
    function IsFinParagraphe(TOBL: TOB): boolean;
    function IsParagraphe(TOBL: TOB): boolean;
    function IsVariante(TOBL: TOB): boolean;
    procedure SetRowHeight(TOBL: TOB; Arow: integer);
    function IsSousTotal(TOBL: TOB): boolean;
    procedure BVerouilleClick (Sender : TObject);
    procedure Valide;
    procedure FParcoursClose (sender : Tobject);
    procedure FParcoursOpen (sender : Tobject);
    procedure PrepareTreeView;
    procedure TVParagClick(Sender: TObject);
    procedure GotoLigne(var Acol, Arow: Integer);

  end ;

Implementation
uses windows,Messages,UAvtVerdon,CalcOLEGenericBTP,
     ExtCtrls, uTOFComm, Hpanel;

procedure TOF_BTSAISAVCVERDON.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISAVCVERDON.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISAVCVERDON.Valide;
var cancel : Boolean;
    Acol,Arow : Integer;
begin
  Acol := GS.Col;
  Arow := GS.Row;
  GSCellExit(Self,Acol,Arow,cancel);
  if not ControleSaisieOk then BEGIN Ecran.ModalResult := 0; Exit; end;
  ValideLaSaisie;
end;

procedure TOF_BTSAISAVCVERDON.OnUpdate ;
begin
  Inherited ;
  Valide;
end ;

procedure TOF_BTSAISAVCVERDON.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISAVCVERDON.OnArgument (S : String ) ;
var Aff0,Aff1,Aff2,Aff3,Aff4 : string;
    QQ : TQuery;
    cancel : boolEan;
    Arow,Acol : Integer;
begin
  fMulDeTraitement := True;
  Inherited ;

  fAction := taModif;
  //
  Fparcours := TToolWindow97.Create(ecran);
  Fparcours.Parent := Ecran;
  Fparcours.ClientWidth := THPanel(GetControl('PPARCOURS')).Width+10;
  Fparcours.ClientHeight := THPanel(GetControl('PPARCOURS')).Height+10;
  Fparcours.Top := Ecran.top+ THPanel(GetControl('PPARCOURS')).top;
  Fparcours.Left := Ecran.left+THPanel(GetControl('PPARCOURS')).left;
  THPanel(GetControl('PPARCOURS')).parent := Fparcours;
  THPanel(GetControl('PPARCOURS')).Align := Alclient;
  THPanel(GetControl('PPARCOURS')).visible := true;
  Fparcours.Caption := TraduireMemoire('Arborescence du document');
  Fparcours.visible := false;
  Fparcours.OnClose := FParcoursClose;
  TVParag := TTreeView (GETControl('TVPARAG'));
  //
  fGestionAff := TAffichageDoc.create(ecran);
  fGestionAff.gestion := TtaNormal;
  fGestionAff.Action := taModif;
  //
  ImTypeArticle := TImageList.Create(ecran);
  ChargeListImage;
  //
  TOBAVC := TOB.Create('BAVCVERDON',nil,-1);
  TOBLigFac := TOB.Create('LES SLIGNES FAC',nil,-1);
  fillchar(cledoc,SizeOf(cledoc),#0);
  cledoc.NaturePiece := '';
  Affaire := '';
  DateApplic := IDate1900;
  NumAvc := 0;
  if LaTOB <> nil then
  begin
    Affaire := LaTOB.GetString('AFFAIRE');
    BTPCodeAffaireDecoupe (Affaire,Aff0,Aff1,Aff2,Aff3,Aff4,tacreat,false);
    SetControlText('AFF_AFFAIRE0',aff0);
    SetControlText('AFF_AFFAIRE1',aff1);
    SetControlText('AFF_AFFAIRE2',aff2);
    SetControlText('AFF_AFFAIRE3',aff3);
    SetControlText('AFF_AVENANT',aff4);
    SetControlText('AFF_DESIGNATION','');
    QQ := OpenSql('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire+'"',True,1,'',true);
    if not QQ.eof then
    begin
      SetControlText('AFF_DESIGNATION',QQ.fields[0].AsString);
    end;
    Ferme(QQ);
    //
    cledoc.NaturePiece := laTOB.GetString('NATUREPIECEG');
    cledoc.Souche := laTOB.GetString('SOUCHE');
    cledoc.NumeroPiece := laTOB.GetInteger('NUMERO');
    cledoc.Indice := laTOB.GetInteger('INDICE');
    DateApplic := LaTOB.GetDateTime('DATEAPPLIC');
    NumAvc := LaTOB.GetInteger('NUMAVC');
    //
    LaTOB.PutEcran(ecran);
  end;
  if cledoc.NaturePiece ='' then
  begin
    PGIError('merci de passer les paramètres d''ouverture de fiche');
    PostMessage(TFVierge(Ecran).Handle,WM_CLOSE,0,0) ;
    Exit;
  end;
  if NumAvc = 0 then
  begin
    if not AVCVerdon.ChargeFromPrevAvc (cledoc,DateApplic,NumAvc,TOBAVC,TOBLigFac) then
    begin
      AVCVerdon.ChargeInitAvc (cledoc,DateApplic,NumAvc,TOBAVC,TOBLigFac);
    end;
  end else
  begin
    AVCVerdon.ChargeTOBFromAvc (cledoc,DateApplic,NumAvc,TOBAVC,TOBLigFac);
  end;
  if TOBLigFac.detail.count = 0 then
  begin
    PGIError('Cette prévision est vide !');
    PostMessage(TFVierge(Ecran).Handle,WM_CLOSE,0,0) ;
    Exit;
  end;
  //
  if TOBAVC.GetBoolean('BVV_VALIDE') then fAction := taConsult
  else if not AVCVerdon.DernierAvancement(cledoc,DateApplic) then fAction := taConsult;
  //
  GS := THGrid(GetControl('GS'));
  ConstitueGrid;
  ChargeLagrid;
  if fAction = TaModif then
  begin
    GS.row :=1; GS.col := 1;
    Arow := 1;
    Acol := 1;
    GSRowEnter(Self,Arow,cancel,false);
    GSCellEnter(Self,Acol,Arow,cancel);
    GS.row := Arow;
    GS.Col := Acol;
    SetEventGrid (true);
    if (NumAvc <> 0) and (AVCVerdon.DernierAvancement(cledoc,DateApplic)) then
    begin
      TToolbarButton97(GetControl('Bdelete')).visible := true;
      TToolbarButton97(GetControl('Bdelete')).onClick := BdeleteClick;
    end;
    if (NumAvc <> 0) and (not TOBAVC.GetBoolean('BVV_VALIDE')) then
    begin
      TToolbarButton97(GetControl('BVERROU')).visible := true;
      TToolbarButton97(GetControl('BVERROU')).onClick := BVerouilleClick;
    end;
  end else
  begin
    GS.Options := GS.Options -[GoEditing]+[GoRowSelect];
    GS.row := 1;
    TToolbarButton97(GetControl('BValider')).Visible := false;
    GS.PostDrawCell := PostDrawCell;
    GS.GetCellCanvas := GetCellCanvas;
  end;
  PrepareTreeView;  
  TToolbarButton97 (GetCOntrol('BPARCOURS')).OnClick := FParcoursOpen;
  TVParag.OnClick := TVParagClick;
end ;

procedure TOF_BTSAISAVCVERDON.OnClose ;
begin
  Fparcours.Free;
  fGestionAff.Free;
  TOBAVC.free;
  TOBLigFac.free;
  ImTypeArticle.Free;
  Inherited ;
end ;

procedure TOF_BTSAISAVCVERDON.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISAVCVERDON.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISAVCVERDON.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,Tiers_: THEdit);
begin
  inherited;
  Aff:=THEdit(GetControl('AFF_AFFAIRE'));
  Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4:=THEdit(GetControl('AFF_AVENANT'));
end;

procedure TOF_BTSAISAVCVERDON.ConstitueGrid;
var ICol,i : Integer;
    SS,SD,FF : String;
begin
  FF := '#';

  if V_PGI.OkDecQ > 0 then
  begin
    FF := '#0.';
    for i := 1 to V_PGI.OkDecQ - 1 do
    begin
      FF := FF + '0';
    end;
    FF := FF + '0';
  end;

  //GS.VidePile(false);
  GS.RowCount := TOBLigFac.detail.Count +1;
  GS.colCount := 7;
  Icol := 0;
  SS := STChaine;
  repeat
    SD := READTOKENST(SS);
    if SD = '' then break;

    if SD='WW' then
    begin
      GS.cells[ICol,0] := ' ';
      GS.ColWidths [iCol] := 20;
      GS.ColEditables [iCol] := false;
      GS.AutoResizeColumn(icol, 20);
    end else if SD = 'GL_TYPEARTICLE' then
    begin
      GS.cells[ICol,0] := ' ';
      GS.ColWidths [iCol] := 36;
      GS.ColEditables [iCol] := false;
      GS.AutoResizeColumn(icol, 36);
    end else if SD = 'GL_CODEARTICLE' then
    begin
      GS.cells[ICol,0] := 'Article';
      GS.ColWidths [iCol] := 60;
      GS.ColEditables [iCol] := false;
    end else if SD = 'GL_LIBELLE' then
    begin
      GS.cells[ICol,0] := 'Désignation';
      GS.ColLengths [iCol] := 70;
      GS.ColWidths [iCol] := 110;
      GS.ColEditables [iCol] := false;
      SG_LIBELLE := Icol;
    end else if SD = 'BLF_QTEMARCHE' then
    begin
      GS.cells[ICol,0] := 'Qté Prévue';
      GS.ColLengths [iCol] := 0;
      GS.ColWidths [iCol] := 50;
      GS.ColEditables [iCol] := false;
      GS.ColAligns[iCol]  := taRightJustify;
      GS.ColFormats[Icol] := FF+';'+FF+'; ;';
    end else if SD = 'BLF_QTECUMULEFACT' then
    begin
      GS.cells[ICol,0] := 'Qté Avancement';
      GS.ColLengths [iCol] := 12;
      GS.ColWidths [iCol] := 50;
      GS.ColEditables [iCol] := true;
      GS.ColAligns[iCol]  := taRightJustify;
      GS.ColFormats[Icol] := FF+';'+FF+'; ;';
      SG_QTECUMULE := icol;
    end else if SD = 'BLF_POURCENTAVANC' then
    begin
      GS.cells[ICol,0] := '% avanc.';
      GS.ColLengths [iCol] := 0;
      GS.ColWidths [iCol] := 50;
      GS.ColEditables [iCol] := true;
      GS.ColAligns[iCol]  := taRightJustify;
      GS.ColFormats[Icol] := '#0.00;#0.00; ;';
      SG_POURCENTAVANC := Icol;
    end;
    inc(ICol);
  until SD='';
end;

procedure TOF_BTSAISAVCVERDON.ChargeLagrid;
var II : Integer;
begin
  for II := 0 to TOBLigFac.detail.count -1 do
  begin
    TOBLigFac.detail[II].PutLigneGrid(GS,II+1,false,false,STChaine);
  end;
  TFvierge(Ecran).hmTrad.ResizeGridColumns(GS);
end;

procedure TOF_BTSAISAVCVERDON.AfficheLaligne(ARow : integer);
begin
  TOBLigFac.detail[Arow-1].PutLigneGrid(GS,Arow,false,false,STChaine);
end;


procedure TOF_BTSAISAVCVERDON.SetEventGrid(State: boolean);
begin
  if State then
  begin
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    GS.OnCellEnter := GSCellEnter;
    GS.OnCellExit := GSCellExit;
    GS.PostDrawCell := PostDrawCell;
    GS.GetCellCanvas := GetCellCanvas;
    TFVierge(ecran).OnKeyDown := FormKeyDown;
  end else
  begin
    GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
    GS.OnCellEnter := nil;
    GS.OnCellExit := nil;
    GS.PostDrawCell := nil;
    GS.GetCellCanvas := nil;
    TFVierge(ecran).OnKeyDown := nil;
  end;
end;

function TOF_BTSAISAVCVERDON.IsParagraphe (TOBL : TOB) : boolean;
var TheValue : string;
begin
  Result := false;
  TheValue := TOBL.GetValue('GL_TYPELIGNE');
  if (copy(TheValue,1,2)='DP') or (copy(TheValue,1,2)='DV') or (copy(TheValue,1,2)='TP') or (copy(TheValue,1,2)='TV') then result := true;
end;

function TOF_BTSAISAVCVERDON.IsDebutParagraphe (TOBL : TOB) : boolean;
var TheValue : string;
begin
  Result := false;
  TheValue := TOBL.GetValue('GL_TYPELIGNE');
  if (copy(TheValue,1,2)='DP') or (copy(TheValue,1,2)='DV') then result := true;
end;

function TOF_BTSAISAVCVERDON.IsFinParagraphe (TOBL : TOB) : boolean;
var TheValue : string;
begin
  Result := false;
  TheValue := TOBL.GetValue('GL_TYPELIGNE');
  if (copy(TheValue,1,2)='TP') or (copy(TheValue,1,2)='TV') then result := true;
end;

function TOF_BTSAISAVCVERDON.IsArticle (TOBL : TOB) : boolean;
var TheValue : string;
begin
  Result := false;
  TheValue := TOBL.GetValue('GL_TYPELIGNE');
  if (TheValue='ART') or (TheValue='ARV') then result := true;
end;

function TOF_BTSAISAVCVERDON.IsCommentaire (TOBL : TOB) : boolean;
var TheValue : string;
begin
  Result := false;
  TheValue := TOBL.GetValue('GL_TYPELIGNE');
  if (TheValue='COM') or (TheValue='COV') then result := true;
end;

function TOF_BTSAISAVCVERDON.IsSousTotal (TOBL : TOB) : boolean;
var TheValue : string;
begin
  Result := false;
  TheValue := TOBL.GetValue('GL_TYPELIGNE');
  if (TheValue='TOT') then result := true;
end;

function TOF_BTSAISAVCVERDON.IsVariante (TOBL : TOB) : boolean;
var TypeLigne : string;
begin
  Result := false;
  TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
  if (TypeLigne = 'ARV') or
     (TypeLigne = 'COV') or
     (TypeLigne = 'SDV') or
     (copy (TypeLigne,1,2) = 'DV') or
     (copy (TypeLigne,1,2) = 'TV')  then result := true;
end;



procedure TOF_BTSAISAVCVERDON.GetCellCanvas(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
begin

  if ACol < GS.FixedCols then
    Exit;
  if ARow < GS.Fixedrows then
    Exit;
  //
  TOBL := TOBLigFac.detail[ARow - 1];

  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then AppliqueStyleSousTot(Canvas, fGestionAff, false);
  //
  if IsDebutParagraphe(TOBL) then
  begin
    AppliqueStyleParag(Canvas, fGestionAff, TOBL.GetValue('GL_NIVEAUIMBRIC'), false);
    exit;
  end;
  {Lignes fin de paragraphe}// Modif BTP
  if IsFinParagraphe(TOBL) then
  begin
    AppliqueStyleParag(Canvas, fGestionAff, TOBL.GetValue('GL_NIVEAUIMBRIC'), false);
    exit;
  end;
  if (IsArticle(TOBL)) and (TOBL.GetString('GL_TYPEARTICLE') = 'ARP') and (not IsVariante(TOBL)) then
  begin
    AppliqueStyleArticle(Canvas, fGestionAff, false);
    Exit;
  end;
  if (IsArticle(TOBL)) and (Pos(TOBL.GetString('GL_TYPEARTICLE'), 'MAR;PRE') >= 0) and (not IsVariante(TOBL)) then
  begin
    AppliqueStyleArticle(Canvas, fGestionAff, false);
    Exit;
  end;
  if IsCommentaire(TOBL) then
  begin
    AppliqueStyleCommentaire(Canvas, fGestionAff, false);
    Exit;
  end;
  if TOBL.GetValue('GL_TYPELIGNE') = 'RL' then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsBold];
    Exit;
  end;


end;

procedure TOF_BTSAISAVCVERDON.SetRowHeight(TOBL: TOB; Arow: integer);
var
  Canvas: TCanvas;
  Height: integer;
begin
  Canvas := TCanvas.create;
  Canvas.Brush.Color := GS.Canvas.Brush.Color;
  Canvas.Font.name := GS.Canvas.Font.name;
  Canvas.Font.Style := GS.canvas.font.Style;
  Canvas.Font.Size := GS.canvas.font.Size;
    //
  if ISArticle(TOBL) or (TOBL.GetString('GL_TYPEARTICLE') = 'ARP') then
  begin
    AppliqueStyleArticle(GS.Canvas, fGestionAff);
  end
  else if IsSousTotal(TOBL) then
  begin
    AppliqueStyleSousTot(GS.Canvas, fGestionAff);
  end
  else if IsParagraphe(TOBL) then
  begin
    AppliqueStyleParag(GS.Canvas, fGestionAff, TOBL.GetValue('GL_NIVEAUIMBRIC'));
  end
  else if IsVariante(TOBL) then
  begin
    AppliqueStyleVariante(GS.Canvas, fGestionAff);
  end
  else if IsCommentaire(TOBL) then
  begin
    AppliqueStyleCommentaire(GS.Canvas, fGestionAff);
  end;

    Height := GS.Canvas.TextHeight('W') + 2;
  if Height < 18 then
    Height := 18;
  GS.RowHeights[Arow] := Height;
  //
  GS.Canvas.Brush.Color := Canvas.Brush.Color;
  GS.Canvas.Font.name := Canvas.Font.name;
  GS.canvas.font.Style := Canvas.Font.Style;
  GS.canvas.font.Size := Canvas.Font.Size;
  Canvas.free;
end;

procedure TOF_BTSAISAVCVERDON.PostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);

  procedure TireLigne(DebX, DebY, FinX, FinY: integer);
  begin
    GS.Canvas.MoveTo(DebX, DebY);
    GS.Canvas.LineTo(FinX, FinY);
  end;

var Arect : TRect;
    TOBL : TOB;
    NumGraph : Integer;
    DecalagePar : integer;
    TheText : string;
begin
  ARect := GS.CellRect(ACol, ARow);
  if (ARow < GS.FixedRows) then Exit;
  TOBL := TOBLigFac.detail[Arow-1];
  if ACol = 0 then SetRowHeight(TOBL, ARow);
  if (ACol < GS.FixedCols) then Exit;
  if ((ACol = 1) and (ARow >= GS.fixedRows)) then
  begin
    if TOBL = nil then
      Exit;
    Canvas.FillRect(ARect);
    NumGraph := RecupTypeGraph(TOBL);
    if NumGraph >= 0 then
    begin
      ImTypeArticle.DrawingStyle := dsTransparent;
      ImTypeArticle.Draw(Canvas, ARect.left + 1, ARect.top, NumGraph);
    end;
    TireLigne(ARect.Left-1, ARect.top, ARect.left-1, ARect.Bottom);
    TireLigne(ARect.right+1, ARect.top, ARect.Right+1, ARect.Bottom);
  end;
  if ((ACol = SG_LIBELLE) and (ARow >= GS.fixedRows)) then
  begin
    if TOBL = nil then Exit;
    DecalagePar := 0;
    if TOBL.GetInteger('GL_NIVEAUIMBRIC') > 0 then
    begin
      if IsParagraphe(TOBL) then
      begin
        DecalagePar := ((TOBL.GetInteger('GL_NIVEAUIMBRIC') - 1) * GetDecalageParagraphe(fGestionAff)) * Canvas.TextWidth('w');
      end
      else
      begin
        DecalagePar := (TOBL.GetInteger('GL_NIVEAUIMBRIC') * GetDecalageParagraphe(fGestionAff)) * Canvas.TextWidth('w');
      end;
    end;

    Canvas.FillRect(ARect);
    TheText := TOBL.getvalue('GL_LIBELLE');
    Canvas.FillRect(ARect);
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.TextOut(ARect.left + DecalagePar + 1, ARect.Top + 2, TheText);
    TireLigne(ARect.Left-1, ARect.top, ARect.left-1, ARect.Bottom);
    TireLigne(ARect.right+1, ARect.top, ARect.Right+1, ARect.Bottom);

  end;
end;

procedure TOF_BTSAISAVCVERDON.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  if not cancel then
  begin
    stCell := GS.Cells[Acol,Arow];
  end;
end;

procedure TOF_BTSAISAVCVERDON.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  if GS.Cells[Acol,Arow] = stCell then Exit;
  TOBL := TOBLigFac.detail[Arow-1];
  if Acol = SG_LIBELLE then
  begin
    GS.Cells[Acol,Arow]:= stCell;
  end else if Acol = SG_QTECUMULE then
  begin
    cancel := AVCVerdon.TraiteQteCumule (GS.Cells[ACol,Arow],TOBL);
    if Cancel then GS.Cells[Acol,Arow]:= stCell;
  end else if Acol = SG_POURCENTAVANC then
  begin
    cancel := AVCVerdon.TraitePourcentAvanc (GS.Cells[ACol,Arow],TOBL);
    if Cancel then GS.Cells[Acol,Arow]:= stCell;
  end;
  AfficheLaligne(ARow);
end;

procedure TOF_BTSAISAVCVERDON.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BTSAISAVCVERDON.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BTSAISAVCVERDON.ChargeListImage;
var UneImage : Timage;
    NomImg : string;
    indice : integer;
begin
  Indice := 1;
  NomImg := 'IMG'+IntToStr(Indice-1);
  UneImage := TImage(GetCOntrol(NomImg));
  repeat
    if UneImage <> nil then
    begin
      if ImTypeArticle.AddMasked  (UneImage.Picture.Bitmap,TColor($FF00FF)) < 0 then exit;
      //
      inc(Indice);
      //
      NomImg := 'IMG'+IntToStr(Indice-1);
      UneImage := TImage(GetCOntrol(NomImg));
    end;
  until UneImage = nil;
end;

function TOF_BTSAISAVCVERDON.RecupTypeGraph(TOBL: TOB): integer;
var TypeArticle,TypeLigne : string;
    NiveauImbric,IndiceNOmen : integer;
begin
  TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
  NiveauImbric := TOBL.GetValue('GL_NIVEAUIMBRIC');
  TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');

  if TypeArticle = 'MAR' Then Result := 7
  else if (TypeArticle = 'OUV') then Result := 1
  else if (TypeArticle = 'POU') then Result := 6
  else if (TypeArticle = 'COM') and (IndiceNomen = 0) then Result := 2
  // VARIANTE
  else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV')) and (NiveauImbric > 1) then Result := 4
  else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV')) and (NiveauImbric = 1 ) then Result := 3
  // --
  else if TypeArticle ='PRE' then Result := 5
  else if TypeArticle ='FRA' then Result := 8
  else if TypeArticle = 'ARP' then Result := 9
  else Result := -1;
end;

procedure TOF_BTSAISAVCVERDON.ZoneSuivanteOuOk(Grille : THgrid;var ACol, ARow : integer;var  Cancel : boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := Grille.SynEnabled;
  Grille.SynEnabled := False;
  Sens := -1;
  ChgLig := (Grille.Row <> ARow);
  ChgSens := False;
  if Grille.Row > ARow then Sens := 1 else if ((Grille.Row = ARow) and (ACol <= Grille.Col)) then Sens := 1;
  ACol := Grille.Col;
  ARow := Grille.Row;
  ii := 0;
  while not ZoneAccessible(Grille,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      Lim := Grille.RowCount ;
      // ---
      if ((ACol = Grille.ColCount - 1) and (ARow >= Lim)) then
      begin
        Sens := -1;
        continue;
      end;
      if ChgLig then
      begin
        ACol := Grille.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < Grille.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := Grille.FixedCols;
      end;
    end else
    begin
      if ((ACol = Grille.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.ColCount;
        ChgLig := False;
      end;
      if ACol > Grille.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := Grille.ColCount - 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

function TOF_BTSAISAVCVERDON.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
begin
  result := true;
  if (Grille.ColWidths[acol] = 0) then BEGIN result := false; exit; END;
  if Arow < GS.FixedRows then exit;
  if Acol < GS.fixedCols then exit;
  if GS.RowHeights[ARow] <= 0 then begin result := false; Exit;  end;
  if Arow > TOBLigFac.detail.count then
  begin
    result := false;
    exit;
  end;
  if ((ACol <> SG_LIBELLE) AND (ACol <> SG_POURCENTAVANC) AND (ACol <> SG_QTECUMULE)) then  BEGIN result := false; exit; END;
end;


function TOF_BTSAISAVCVERDON.ControleSaisieOk: Boolean;
var II : Integer;
    Question,Question1 : boolean;
begin
  Result := false;
  Question := false; Question1 := false;
  for II := 0 to TOBLigFac.detail.Count -1 do
  begin
    if TOBLigFac.detail[II].GetDouble('BLF_POURCENTAVANC')> 100 then
    begin
      Question := True;
    end;
    if TOBLigFac.detail[II].GetDouble('BLF_QTESITUATION')< 0 then
    begin
      Question1 := True;
    end;
    if TOBLigFac.detail[II].GetDouble('BLF_QTESITUATION')<>0 then
    begin
      Result := true;
    end;
  end;
  if not Result then
  begin
    PGIInfo('Aucun avancement n''a été saisi');
    Exit;
  end;
  if Question then
  begin
    Result := (PGIAsk('Certains avancements sont supérieur à 100%.#13#10 Confirmez-vous la saisie?')=mrYes);
    if not Result then Exit;
  end;
  if Question1 then
  begin
    Result := (PGIAsk('Certains avancements sont inférieurs aux précédents.#13#10 Confirmez-vous la saisie?')=mrYes);
  end;
end;

procedure TOF_BTSAISAVCVERDON.ValideLaSaisie;
begin
  AVCVerdon.ValideLaSaisie(cledoc,Affaire,DateApplic,NumAvc,TOBAVC,TOBLigFac);
end;

procedure TOF_BTSAISAVCVERDON.BdeleteClick(Sender: TObject);
begin
  if PGIAsk('ATENTION : Vous allez supprimer cet avancement.#13#10 Etes-vous sur ?')<> Mryes then Exit;
  AVCVerdon.Supprime(cledoc,DateApplic);
  SendMessage(Ecran.handle,WM_CLOSE,0,0);
end;

procedure TOF_BTSAISAVCVERDON.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var vide,OkG : boolean;
begin
  OkG := (TFVierge(Ecran).ActiveControl = ThGrid(GEtControl('GS')));

  Vide := (Shift = []);
  case Key of
    VK_RETURN:
      if (OkG) and (Vide) then Key := VK_TAB;
  end;
end;

procedure TOF_BTSAISAVCVERDON.BVerouilleClick(Sender: TObject);
begin
  if not TOBAVC.GetBoolean('BVV_VALIDE') then
  begin
    if PGIAsk('Etes-vous sur de vouloir arreter cet avancement?')<>mryes then Exit;
    TOBAVC.SetBoolean('BVV_VALIDE',True);
    Valide;
    SendMessage(Ecran.handle,WM_CLOSE,0,0);
  end;
end;

procedure TOF_BTSAISAVCVERDON.FParcoursClose(sender: Tobject);
begin
  TToolbarButton97 (GetControl('BPARCOURS')).Down := false;
end;

procedure TOF_BTSAISAVCVERDON.FParcoursOpen(sender: Tobject);
begin
	Fparcours.visible := TToolbarButton97 (GetControl('BPARCOURS')).Down;
end;

procedure TOF_BTSAISAVCVERDON.PrepareTreeView;
  function IsRefDevis(TOBL: TOB): boolean;
  begin
    result := false;
    if (Pos(TOBL.GetValue('GL_NATUREPIECEG'), 'FBT;FBP') > 0) and (TOBL.GEtVAlue('GL_TYPEARTICLE') = 'EPO') then
      result := true;
  end;

var
  TOBL: TOB;
  TOBC, TOBCL: TOB;
  i: Integer;
  Tn, RootTN: TTreeNode;
  Titre: string;
begin
  TOBC := TOB.Create('LA TOB ', nil, -1);
  try
    // chargement treeview
    TVParag.items.clear;
    Titre := RechDom('GCNATUREPIECEG',cledoc.NaturePiece,false) + ' N° ' + TOBAVC.getString('BVV_NUMERO');
    RootTN := TVParag.items.add(nil, Titre);
    TOBCL := TOB.Create('UN ITEM', TOBC, -1);
    TOBCL.data := RootTN;

    for i := 0 to TOBLigFac.Detail.Count - 1 do
    begin
      TOBL := TOBLigFac.Detail[i];
      // VARIANTE
      if (IsDebutParagraphe(TOBL)) then
      begin
        Tn := TVParag.Items.AddChild(TTreeNode(TOBC.detail[TOBC.detail.count - 1].data), TOBL.GetValue('GL_LIBELLE'));
        Tn.Data := TOBL;
        TOBCL := TOB.Create('UN ITEM', TOBC, -1);
        TOBCL.data := Tn;
      end
      else if IsFinParagraphe(TOBL) then
      begin
        TOBC.detail[TOBC.detail.count - 1].free;
      end;
    end;
  finally
    TOBC.free;
  end;
  TVParag.FullExpand;
end;

procedure TOF_BTSAISAVCVERDON.GotoLigne (var Acol,Arow : Integer);
var cancel : boolean;
begin
  if fAction <> taConsult then
  begin
    SetEventGrid(false);
    GS.Col := Acol;
  end;
  GS.row := Arow;
  if fAction <> taConsult then SetEventGrid (true);

end;

procedure TOF_BTSAISAVCVERDON.TVParagClick(Sender: TObject);
var
  Tn: TTreeNode;
  TOBL: TOB;
  Acol, Arow: integer;
begin
  Tn := TVParag.selected;
  TOBL := TOB(Tn.data);
  if (TOBL <> nil) then
  begin
    Arow := TOBL.GetInteger('GL_NUMLIGNE');
    Acol := SG_LIBELLE;
    GotoLigne(Acol,Arow);
    GS.Row := Arow;
    GS.col := Acol;
    if fACtion <> taConsult then StCell := GS.Cells[GS.Col, GS.Row];
  end;
end;



Initialization
  registerclasses ( [ TOF_BTSAISAVCVERDON ] ) ;
end.
