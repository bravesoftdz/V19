{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/11/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISPAIMENTPOC ()
Mots clefs ... : TOF;BTSAISPAIMENTPOC
*****************************************************************}
Unit BTSAISPAIMENTPOC_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,Fe_Main,
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
     Vierge,
     Graphics,
     SAISUTIL,
     Windows,
     Messages,
     AGLInit,
     HTB97,
     CommonTools,
     UTOF ;

Type
  TOF_BTSAISPAIMENTPOC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBPiece : TOB;
    TOBSAISIE : TOB;
    TOBParam : TOB;
    TOBPieceTrait : TOb;
    TOBSOUSTrait : TOB;
    GS : THGrid;
    stcell : string;
    fColNames,Falignement,Ftitre,fLargeur : string;
    ColSaisie : Integer;
    DEV : Rdevise;
    procedure BInsertClick (Sender : TObject);
    procedure BDeleteClick (Sender : TObject);
    procedure BfermeClick (Sender : TObject);
    procedure DefiniGrid;
    procedure DefiniGrille(Grid: ThGrid; fColNames, Falignement, Ftitre,fLargeur: string);
    procedure ChargeLaGrille;
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ZoneSuivanteOuOk(Grille: THgrid; var ACol, ARow: integer; var Cancel: boolean);
    function ZoneAccessible(Grille: THgrid; var ACol,ARow: integer): boolean;
    procedure AjouteLigne(TOBL: TOB; Sens: string='+');
    procedure RaffraichitTotal;
  end ;

function DemandeReglementSousTrait (TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TObpieceTrait,TOBPieceRG,TOBSousTrait,TOBPorcs: TOB; MontantPaiement : double;  DEV : RDevise) : boolean;

Implementation
uses FactComm,UCotraitance,UspecifPOC,BTStructChampSup,ParamSoc;

function DemandeReglementSousTrait (TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TObpieceTrait,TOBPieceRG,TOBSousTrait,TOBPorcs: TOB; MontantPaiement : double;  DEV : RDevise) : boolean;
var TOBParam : TOB;
begin
  Result := True;
  if TOBpieceTrait.detail.count = 0 then Exit;
  TOBParam := TOB.Create('LES PARAMS',nil,-1);
  TRY
    TOBParam.AddChampSupValeur('OKOK','X');
    TOBParam.AddChampSupValeur('MODIF','-');
    TOBParam.AddChampSupValeur('NETAPAYER',MontantPaiement);
    TOBParam.AddChampSupValeur('NBDEC',DEV.decimale);
    TheTOB := TOBParam;
    TOBParam.Data := TObpieceTrait;
    TOBpieceTrait.data := TOBSousTrait;
    AGLLanceFiche('BTP','BTSAISPAIMENTPOC','','','ACTION=MODIFICATION');
    TheTOB := nil;
    TOBpieceTrait.data := nil;
    if TOBParam.getstring('MODIF')='X' then
    begin
      TOBEches.clearDetail;
      GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBPieceTrait,TOBPorcs, taCreat, DEV, false);
    end;
    if TOBParam.GetString('OKOK') <> 'X' then Result := false;
  FINALLY
    TOBParam.free;
  end;
end;

procedure TOF_BTSAISPAIMENTPOC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.OnUpdate ;
var II : Integer;
    TOBS , TOBL, TOBI  : TOB;
begin
  Inherited ;
  if ARRONDI(THNumEdit(getCOntrol('NETAPAYER')).Value,DEV.Decimale) <> ARRONDI(THNumEdit(getCOntrol('ENPAIEMENT')).Value,DEV.decimale) then
  begin
    PGIError('Le montant en règlement ne corrrespond pas au montant à payer par le client');
    ecran.ModalResult := 0;
    exit;
  end;
  //
  II := 0;
  repeat
    TOBS := TOBSAISIE.detail[II]; 
    TOBL := TOBPieceTrait.FindFirst(['BPE_FOURNISSEUR'],[TOBS.GetString('BPE_FOURNISSEUR')],true);
    if TOBL = nil then
    begin
      TOBS.changeParent (TOBPieceTrait,-1);
      TOBS.SetBoolean ('BPE_REGLSAISIE',true);
      if TOBS.GetString('BPE_FOURNISSEUR') <> '' then
      begin
        TOBI := TOBSOUSTrait.FindFirst(['BPI_TIERSFOU'],[TOBS.GetString('BPE_FOURNISSEUR')],true);
        if TOBI <> nil then
        begin
          if TOBS.GetDouble('BPE_MONTANTREGL') <> 0 then
          begin
            TOBI.SetString('BPI_TYPEPAIE','001');
            TOBS.SetString('TYPEPAIE','001');
          end else
          begin
            TOBI.SetString('BPI_TYPEPAIE','002');
            TOBS.SetString('TYPEPAIE','002');
          end;
        end;
      end;
    end else
    begin
      if TOBS.GetDouble('BPE_MONTANTREGL')<>TOBL.GetDouble('BPE_MONTANTREGL') then
      begin
        TOBL.SetDouble('BPE_MONTANTREGL',TOBS.GetDouble('BPE_MONTANTREGL'));
        TOBL.SetBoolean ('BPE_REGLSAISIE',true);
        if TOBS.GetString('BPE_FOURNISSEUR') <> '' then
        begin
          TOBI := TOBSOUSTrait.FindFirst(['BPI_TIERSFOU'],[TOBS.GetString('BPE_FOURNISSEUR')],true);
          if TOBI <> nil then
          begin

            if TOBS.GetDouble('BPE_MONTANTREGL') <> 0 then
            begin
              TOBI.SetString('BPI_TYPEPAIE','001');
              TOBL.SetString('TYPEPAIE','001');
            end else
            begin
              TOBI.SetString('BPI_TYPEPAIE','002');
              TOBL.SetString('TYPEPAIE','002');
            end;
          end;
        end;
      end;
      Inc(II);
    end;
  until II >= TOBSAISIE.detail.count;
  TOBParam.SetString('MODIF','X');
end ;

procedure TOF_BTSAISPAIMENTPOC.OnLoad ;
begin
  Inherited ;
end ;



procedure TOF_BTSAISPAIMENTPOC.OnArgument (S : String ) ;
var cancel :boolean;
    Arow,Acol : Integer;
begin
  Inherited ;
  TOBSAISIE := TOB.Create('LES LIGNES',nil,-1);
  //
  TOBParam := LaTOB;
  DEV.decimale := TOBParam.Getinteger('NBDEC');
  TOBPieceTrait := TOB(TOBParam.Data);
  TOBSOUSTRAIT := TOB(TOBPieceTrait.Data);
  TOBPIece := TOB(TOBSOUSTRAIT.Data);
  TOBPieceTrait.Data := nil;
  TOBParam.Data := nil;
  TOBSOUSTRAIT.data := nil;
  TOBSAISIE.Dupliquer(TOBPieceTrait,True,true);
  GS := THGrid(GetControl('GS'));
  THNumEdit(getCOntrol('NETAPAYER')).Value := Arrondi(TOBParam.GetDouble('NETAPAYER'),DEV.Decimale);
  THNumEdit(getCOntrol('ENPAIEMENT')).Value := Arrondi(TOBParam.GetDouble('NETAPAYER'),DEV.Decimale);
  TToolbarButton97 (GetControl('Bferme')).onclick := BfermeClick;
  DefiniGrid;
  ChargeLaGrille;
  Acol := ColSaisie;
  Arow := 1;
  SetFocusControl('GS'); 
  GS.Row := 1;
  GS.Col := ColSaisie;
  GS.OnCellEnter := GSCellEnter;
  GS.OnCellExit := GSCellExit;
  GS.OnKeyDown := GSKeyDown;
  TToolbarButton97(GetControl('BInsert')).OnClick := BInsertClick;
  TToolbarButton97(GetControl('BDelete')).OnClick := BDeleteClick;
  GSCellEnter(self,Acol,Arow,cancel);
end ;

procedure TOF_BTSAISPAIMENTPOC.OnClose ;
begin
  TOBSAISIE.free;
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.DefiniGrid;
begin
  fColNames := 'SEL;BPE_FOURNISSEUR;LIBELLE;BPE_MONTANTREGL;';
  Falignement := 'D.0  ---;G.0  ---;G.0  ---;D/2O -X-;';
  Ftitre := 'Num;Sous-Traitant;Raison sociale;Règlement client;';
  fLargeur := '5;18;60;20;';
  DefiniGrille (GS,fColNames,Falignement,Ftitre,fLargeur);
end;

procedure TOF_BTSAISPAIMENTPOC.DefiniGrille (Grid: ThGrid; fColNames,Falignement,Ftitre,fLargeur : string);
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,lelement : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec,NbCols : integer;
    indice,ind : integer;
    FFQTE : string;
begin
  //
  if Grid.rowcount < 2 then Grid.rowCount := 2;
  st := fColNames;
  NbCols := 0;
  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  Grid.ColCount := Nbcols;
  //
  st := fColNames;
  lesalignements := Falignement;
  lestitres := Ftitre;
  leslargeurs := fLargeur;

  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    OkVisu  := True;
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    grid.cells[Indice,0] := leTitre;
    grid.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G'      then grid.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then grid.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then grid.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if (Nam = 'BPE_MONTANTREGL') then
    begin
      grid.ColEditables[indice] := True;
      ColSaisie := indice;
    end else
    begin
      grid.ColEditables[indice] := false;
      grid.ColLengths [indice] := 0;
    end;
    //
    if OkVisu then
    begin
  		grid.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
    end else
    begin
    	grid.ColWidths[indice] := -1;
    end;

    if (OkLib) or (okImg) then
    begin
    	grid.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	grid.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end else if (Dec<>0) or (Sep) then
    begin
      FFQTE := '#';
      if Sep then FFQte := FFQTE+' ##';
      if Dec > 0 then
      begin
        FFQTE := FFQTE+'0.';
        for ind := 1 to Dec do
        begin
          FFQTE := FFQTE + '0';
        end;
      end else FFQTE := FFQTE + '0';
      GS.ColFormats[indice] := FFQTE ;
    end;

  end ;
end;


procedure TOF_BTSAISPAIMENTPOC.ChargeLaGrille;
begin
	GS.RowCount := TOBSAISIE.Detail.count+1;
  if GS.rowCount < 2 then GS.rowCount := 2;
	TOBSAISIE.PutGridDetail(GS,false,false,fColNames);
  TFvierge(ecran).HMTrad.ResizeGridColumns(GS);
end;

procedure TOF_BTSAISPAIMENTPOC.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteOuOk  (GS,Acol,Arow,Cancel);
  if not cancel then
  begin
    stcell := GS.Cells[Acol,Arow];  
  end;
end;

procedure TOF_BTSAISPAIMENTPOC.AjouteLigne(TOBL : TOB;SenS : string ='+');
begin
  if Sens = '+' then
  begin
    THNumEdit(getCOntrol('ENPAIEMENT')).Value := arrondi(THNumEdit(getCOntrol('ENPAIEMENT')).Value + TOBL.GetDouble('BPE_MONTANTREGL'),DEV.decimale);
  end else
  begin
    THNumEdit(getCOntrol('ENPAIEMENT')).Value := arrondi(THNumEdit(getCOntrol('ENPAIEMENT')).Value - TOBL.GetDouble('BPE_MONTANTREGL'),DEV.decimale);
  end;
end;

procedure TOF_BTSAISPAIMENTPOC.RaffraichitTotal;
begin
  if THNumEdit(getCOntrol('ENPAIEMENT')).Value <> THNumEdit(getCOntrol('NETAPAYER')).Value then
  begin
    THNumEdit(getCOntrol('ENPAIEMENT')).Font.Color := clred;
  end else
  begin
    THNumEdit(getCOntrol('ENPAIEMENT')).Font.Color := clblack;
  end;
end;

procedure TOF_BTSAISPAIMENTPOC.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  if GS.cells[Acol,Arow] = stcell then Exit;
  if Acol = ColSaisie then
  begin
    TOBL := TOBSAISIE.detail[Arow-1];
    AjouteLigne(TOBL,'-');

    TOBL.SetDouble('BPE_MONTANTREGL',VALEUR(GS.cells[Acol,Arow]));
    AjouteLigne(TOBL);

    RaffraichitTotal;
    TOBL.PutLigneGrid(GS,ARow,False,False,fColNames);
    StCell := GS.cells[Acol,Arow];
  end;
end;



function TOF_BTSAISPAIMENTPOC.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
begin
  result := true;
  if (Grille.ColWidths[acol] = 0) or (not Grille.ColEditables[acol] ) then BEGIN result := false; exit; END;
  if Arow < GS.FixedRows then exit;
  if Acol < GS.fixedCols then exit;
end;

procedure TOF_BTSAISPAIMENTPOC.ZoneSuivanteOuOk(Grille : THgrid;var ACol, ARow : integer;var  Cancel : boolean);
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


procedure TOF_BTSAISPAIMENTPOC.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var OkG, Vide: Boolean;
  TOBL : TOB;
  Diff : double;
begin
  OkG := (Screen.ActiveControl = GS);
  Vide := (Shift = []);
  case Key of
    VK_RETURN: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        SendMessage(GS.Handle, WM_KeyDown, VK_TAB, 0);
      end;
    VK_F11: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        TOBL := TOBSAISIE.detail[GS.row-1];
        Diff := TOBL.GetDouble('BPE_MONTANTREGL') + THNumEdit(getCOntrol('NETAPAYER')).Value - THNumEdit(getCOntrol('ENPAIEMENT')).Value;
        GS.cells[GS.Col,GS.row] := STRFPOINT(diff);
        SendMessage(GS.Handle, WM_KeyDown, VK_TAB, 0);
      end;
  end;
end;

procedure TOF_BTSAISPAIMENTPOC.BfermeClick(Sender: TObject);
begin
  TOBParam.SetString('OKOK','-');
  Ecran.Close;
end;

procedure TOF_BTSAISPAIMENTPOC.BDeleteClick(Sender: TObject);
var TOBL : TOB;
    LastRow : integer;
  Acol,Arow : Integer;
  Cancel : Boolean;
begin
  Lastrow := GS.row;
  TOBL := TOBSaisie.detail[GS.Row-1];
  if (TOBL.GetString('BPE_FOURNISSEUR')='')  then
  begin
    PGIInfo('On ne peut pas retirer votre société du paiement');
    Exit;
  end;
  if (TOBL.GetDouble('BPE_MONTANTFAC')<>0)  then
  begin
    PGIInfo('On ne peut pas retirer un sous traitant présent dans la facturation');
    Exit;
  end;
	GS.CacheEdit;
  GS.SynEnabled := false;
  GS.OnCellEnter := nil;
  GS.OnCellExit := nil;
  AjouteLigne (TOBL,'-');
  TOBSaisie.detail[GS.Row-1].Free;
  ChargeLaGrille;
  GS.Row := 1;
  GS.Col := ColSaisie;
  Acol := GS.col;
  Arow := GS.row;
  GS.SynEnabled := true;
	GS.ShowEditor;
  GS.OnCellEnter := GSCellEnter;
  GS.OnCellExit := GSCellExit;
  GSCellEnter(self,Acol,Arow,cancel);
end;

procedure TOF_BTSAISPAIMENTPOC.BInsertClick(Sender: TObject);

  procedure AddSousTraitant(Fournisseur,CodeMarche : string);
  var TT,TTE : TOB;
  begin
    TT := TOB.Create('PIECEINTERV',TOBSOUSTrait,-1);
    AddChampsSupTraitCreat (TT);
    AddChampsSupTrait (TT);
    TT.putValue('BPI_NATUREPIECEG',TOBpiece.GetString('GP_NATUREPIECEG'));
    TT.putValue('BPI_SOUCHE',TOBpiece.GetString('GP_SOUCHE'));
    TT.putValue('BPI_NUMERO',TOBpiece.GetInteger('GP_NUMERO') );
    TT.putValue('BPI_INDICEG',TOBpiece.GetInteger('GP_INDICEG'));
    TT.SetInteger('BPI_ORDRE',TOBSOUSTrait.detail.count -1);
    //
    TT.putValue('BPI_TYPEINTERV','Y00');
    TT.putValue('BPI_TYPEPAIE','001');
    TT.putValue('BPI_DATECONTRAT',V_PGI.DateEntree);
    TT.SetString('BPI_FAMILLETAXE',GetInfoMarcheST(TOBpiece.GetString('GP_AFFAIRE'),Fournisseur,CodeMarche,'FAMILLETAXE1'));
    TT.SetBoolean('BPI_AUTOLIQUID',Tools.StringInList(TT.GetString('BPI_FAMILLETAXE'), GetParamSocSecur('SO_CODETVALIQUIDST','')));
    TT.SetString('BPI_TIERSFOU',Fournisseur);
    TT.SetString('BPI_CODEMARCHE',CodeMarche);
    TT.SetDouble('NBUSED',1);
    //
    TTE := TOBSAISIE.FindFirst(['BPE_FOURNISSEUR'],[Fournisseur],true);
    if TTE = nil then
    begin
      TTE := TOB.Create('PIECETRAIT',TOBSAISIE,-1);
      AddlesChampsSupPieceTrait (TTE);

      TTE.putValue('BPE_NATUREPIECEG',TOBpiece.GetString('GP_NATUREPIECEG'));
      TTE.putValue('BPE_SOUCHE',TOBpiece.GetString('GP_SOUCHE'));
      TTE.putValue('BPE_NUMERO',TOBpiece.GetInteger('GP_NUMERO') );
      TTE.putValue('BPE_INDICEG',TOBpiece.GetInteger('GP_INDICEG'));
      TTE.SetString('BPE_FOURNISSEUR',Fournisseur);
      TTE.SetString ('LIBELLE',GetLibelleFou(Fournisseur));
      TTE.SetString('BPE_AFFAIRE',TOBPiece.GetString('GP_AFFAIRE'));
      TTE.SetString('BPE_TYPEINTERV',TT.GetString('BPI_TYPEINTERV'));
    end;

  end;

  procedure  AjouteSousTraitantPOC(Fournisseur,CodeMarche : string);
  var TT : TOB;
  begin
    TT := TOBSOUSTrait.FindFirst(['BPI_TYPEINTERV','BPI_TIERSFOU','BPI_CODEMARCHE'],['Y00',Fournisseur,CodeMarche],true);
    if TT = nil then
    begin
      AddSousTraitant(Fournisseur,CodeMarche);
    end else
    begin
      TT.SetDouble('NBUSED',TT.GetDouble('NBUSED')+1);
    end;
  end;

var TOBParam : TOB;
    CodeMarche,Fournisseur,ModePaie,Affaire : string;
begin
  Affaire := TOBpiece.GetValue('GP_AFFAIRE');
  CodeMarche := '';
  Fournisseur := '';
  // Sélection du Sous traitant
  TOBParam := TOB.Create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('ACTION','SELECTION');
  TOBParam.AddChampSupValeur ('SOUSTRAIT','');
  TOBParam.AddChampSupValeur ('CODEMARCHE','');
  TOBParam.AddChampSupValeur ('AFFAIRE',Affaire);
  TOBParam.Data := TObpieceTrait;
  TObpieceTrait.Data := TOBSOUSTrait;
  TheTOB := TOBParam;
  Fournisseur := '';
  ModePaie := '001';
  TRY
  	AGLLanceFiche('BTP','BTMARCHEST_MUL','BM1_AFFAIRE='+Affaire,'','');
    // --
    Fournisseur := TOBParam.GetValue('SOUSTRAIT');
    CodeMarche := TOBParam.GetValue('CODEMARCHE');
  FINALLY
    TObpieceTrait.Data := nil;
    TheTOB := nil;
  	TOBParam.free;
  end;
  if (Fournisseur <> '' ) and (CodeMarche <> '') then
  begin
    AjouteSousTraitantPOC (Fournisseur,CodeMarche);
    ChargeLaGrille;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTSAISPAIMENTPOC ] ) ;
end.
