{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/05/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTVENTILCOLL ()
Mots clefs ... : TOF;BTVENTILCOLL
*****************************************************************}
Unit BTVENTILCOLL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul,
     fe_main,
{$else}
     eMul,
{$ENDIF}
     HTB97, 
     Hpanel,
     Vierge, 
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     UTOF ;

Type

  TOF_BTVENTILCOLL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    STCellCur : string;
    GS : THGrid;
    DataModified : Boolean;
    TOBDatas : TOB;
    TOBTVAS : TOB;
    ColCollectifAc,ColCollectif : Integer;
    function GetControlValue(ControlName : string) : string;
    function WhereDatas : string;
    procedure ConstitueDatas;
    procedure AffecteValeurs;
    procedure SaveDatas;
    procedure DeleteDatas;
    procedure AjouteDatas (Sender : TObject);
    function BeforeAjouteDatas : Boolean;
    procedure RemplitDatas;
    procedure ChangeNature (Sender : TObject);
    procedure ConstitueGrid;
    procedure AfficheGrid;
    procedure SetEvents(Etat : boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BdeleteClick (sender : TObject);
  end ;

const
  ListeSaisie = ' ;LIBTAXE;BVC_COLLECTIFAC;BVC_COLLECTIF;';

Implementation

procedure TOF_BTVENTILCOLL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLL.OnUpdate ;
var II : integer;
    cancel : boolean;
    Acol,Arow : Integer; 
begin
  Inherited ;
  Acol := GS.Col;
  Arow := GS.row;
  GSCellExit(Self,Acol,Arow,cancel);
  II := 0;
  repeat
    if (TOBDatas.detail[II].GetString ('BVC_COLLECTIF')='') and (TOBDatas.detail[II].GetString ('BVC_COLLECTIFAC')='') then
    begin
      TOBDatas.detail[II].free;
    end else
    begin
      inc(II);
    end;
  until II >= TOBDatas.detail.count;
  SaveDatas;
  DataModified := false;
end ;

procedure TOF_BTVENTILCOLL.OnLoad ;
begin
  Inherited ;
  //
  if TFVierge(Ecran).TypeAction = TaModif then
  begin
    THPanel(GetControl('PCRIT')).enabled := false;
    SetControlVisible('BDelete',true);
    RemplitDatas;
    ChangeNature (self);
  end else
  begin
    TToolbarButton97 (GetControl('BFIND')).OnClick := AjouteDatas;
    THValComboBox (GetControl('BVC_NATUREV')).Onchange := ChangeNature;
    SetControlText('BVC_CODEPORT','');
    SetControlText('BVC_COMPTAARTICLE','');
    SetControlText('BVC_COMPTATIERS','');
    SetControlText('BVC_COMPTAAFFAIRE','');
    SetControlText('BVC_ETABLISSEMENT','');
    SetControlText('BVC_REGIMETAXE','');
  end;
end ;

procedure TOF_BTVENTILCOLL.OnArgument (S : String ) ;
var QQ : TQuery;
    critere,ChampMul,ValMul : string;
    x : Integer;
begin
  Inherited ;

  repeat
    Critere := Trim(ReadTokenSt(s));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'ACTION' then
        begin
          if ValMul = 'MODIFICATION' then TFVierge(Ecran).TypeAction := taModif;
        end else if GetControl(Champmul) <> nil then
        begin
          SetControlText(ChampMul,ValMul);
        end;
      end;
    end;
  until Critere = '';

  TOBDatas := TOB.Create ('LES VENTILS COLL',nil,-1);
  TOBTVAS := TOB.Create('LES TAXES',nil,-1);

  QQ := OpenSQL('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="TX1"',True,-1,'',true);
  if not QQ.Eof then
  begin
    TOBTVAS.LoadDetailDB('CHOIXCOD','','',QQ,false);
  end;
  ferme (QQ);
  //
  GS := THGrid(GetControl('GS'));
end ;

procedure TOF_BTVENTILCOLL.OnClose ;
begin
  if DataModified then
  begin
    if PGIAsk('Attention. Vous avez modifié des données.#13#10 Etes-vous sur(e) d''abandonner vos modifications' ) <> mrYes then
    begin
      TFVierge(Ecran).ModalResult := 0;
      Exit;
    end;
  end;

  TOBDatas.free;
  TOBTVAS.free;
  Inherited ;
end ;

procedure TOF_BTVENTILCOLL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCOLL.ConstitueDatas;
var II : Integer;
    TT : TOB;
begin
  for II := 0 to TOBTVAS.detail.Count -1 do
  begin
    TT := TOB.Create ('BVENTILCOLL',TOBDatas,-1);
    TT.SetString('BVC_NATUREV',GetControlValue('BVC_NATUREV'));
    TT.SetString('BVC_CODEPORT',GetControlValue('BVC_CODEPORT'));
    TT.SetString('BVC_COMPTAARTICLE',GetControlValue('BVC_COMPTAARTICLE'));
    TT.SetString('BVC_COMPTATIERS',GetControlValue('BVC_COMPTATIERS'));
    TT.SetString('BVC_COMPTAAFFAIRE',GetControlValue('BVC_COMPTAAFFAIRE'));
    TT.SetString('BVC_ETABLISSEMENT',GetControlValue('BVC_ETABLISSEMENT'));
    TT.SetString('BVC_REGIMETAXE',GetControlValue('BVC_REGIMETAXE'));
    TT.SetString('BVC_FAMILLETAXE',TOBTVAS.detail[II].GetString('CC_CODE'));
    TT.AddChampSupValeur('LIBTAXE',TOBTVAS.detail[II].GetString('CC_LIBELLE'));
  end;
end;

procedure TOF_BTVENTILCOLL.AffecteValeurs;
var QQ : TQuery;
    SQL : String;
    TT : TOB;
begin
  SQL := 'SELECT * FROM BVENTILCOLL '+
         'WHERE '+ WhereDatas;
  QQ := OpenSQL(SQL,True,-1,'',true );
  if not QQ.eof then
  begin
    QQ.First;
    repeat
      TT := TOBDatas.FindFirst(['BVC_COMPTAARTICLE','BVC_COMPTATIERS','BVC_COMPTAAFFAIRE','BVC_ETABLISSEMENT','BVC_REGIMETAXE','BVC_FAMILLETAXE'],
                               [GetControlValue('BVC_COMPTAARTICLE'),GetControlValue('BVC_COMPTATIERS'),
                                GetControlValue('BVC_COMPTAAFFAIRE'),GetControlValue('BVC_ETABLISSEMENT'),GetControlValue('BVC_REGIMETAXE'),
                                QQ.FindField('BVC_FAMILLETAXE').AsString],true);
      if TT <> nil then
      begin
        TT.SetString('BVC_COLLECTIF',QQ.FindField('BVC_COLLECTIF').AsString);
        TT.SetString('BVC_COLLECTIFAC',QQ.FindField('BVC_COLLECTIFAC').AsString);
      end;
      QQ.Next;
    until QQ.Eof;
  end;
  Ferme(QQ);
  TOBDatas.SetAllModifie(false);
end;

procedure TOF_BTVENTILCOLL.SaveDatas;
begin
  BEGINTRANS;
  try
    DeleteDatas;
    TOBDatas.SetAllModifie(True);
    TOBDatas.InsertDB(nil);
    COMMITTRANS;
  EXCEPT
    PGIError('les données n''ont pu être sauvegardées');
    ROLLBACK;
  END;
end;

procedure TOF_BTVENTILCOLL.DeleteDatas;
var SQL : string;
begin
  SQL := 'DELETE FROM BVENTILCOLL WHERE '+ WhereDatas;
  ExecuteSQL(SQL);
end;

procedure TOF_BTVENTILCOLL.RemplitDatas;
var cancel : Boolean;
    Acol,Arow : Integer;
begin
  TOBDatas.ClearDetail;
  ConstitueDatas;
  AffecteValeurs;
  ConstitueGrid;
  AfficheGrid;
  GS.visible := true;

  THPanel(GetControl('PCRIT')).enabled := false;
  SetEvents(True);
  cancel := false;
  GS.Row := 1;
  GS.Col := 2;
  Acol := 2;
  Arow := 1;
  GSCellEnter(Self,Acol,Arow,cancel);
  GS.ShowEditor;
  DataModified := false;
end;

procedure TOF_BTVENTILCOLL.AjouteDatas (Sender : TObject);
begin
  if not BeforeAjouteDatas then Exit;
  RemplitDatas;
end;

function TOF_BTVENTILCOLL.BeforeAjouteDatas: Boolean;
var SQL : string;
begin
  Result := true;
  if GetControlValue('BVC_NATUREV')='' then
  begin
    PGIInfo('Vous devez renseigner la nature');
    Result := false;
    Exit;
  end;
  
  if (GetControlValue('BVC_NATUREV')='003') and (GetControlValue('BVC_CODEPORT')='') then
  begin
    PGIInfo('Vous devez renseigner une avance/retenue');
    Result := false;
    Exit;
  end;

  SQL := 'SELECT 1 FROM BVENTILCOLL WHERE '+ WhereDatas;
  if ExisteSQL(SQL) then
  begin
    PGIInfo('Cette définition existe dèjà');
    Result := false;
    Exit;
  end;
end;

function TOF_BTVENTILCOLL.GetControlValue(ControlName: string): string;
begin
  Result := THValComboBox (GetControl(ControlName)).Value;
end;

function TOF_BTVENTILCOLL.WhereDatas: string;
begin
Result := 'BVC_NATUREV="'+GetControlValue('BVC_NATUREV')+'" AND '+
          'BVC_CODEPORT="'+GetControlValue('BVC_CODEPORT')+'" AND '+
          'BVC_COMPTAARTICLE="'+GetControlValue('BVC_COMPTAARTICLE')+'" AND '+
          'BVC_COMPTATIERS="'+GetControlValue('BVC_COMPTATIERS')+'" AND '+
          'BVC_COMPTAAFFAIRE="'+GetControlValue('BVC_COMPTAAFFAIRE')+'" AND '+
          'BVC_ETABLISSEMENT="'+GetControlValue('BVC_ETABLISSEMENT')+'" AND '+
          'BVC_REGIMETAXE="'+GetControlValue('BVC_REGIMETAXE')+'"';
end;

procedure TOF_BTVENTILCOLL.ChangeNature(Sender: TObject);
var TypeD : string;
begin
  TypeD := GetControlValue('BVC_NATUREV');
  if typeD = '003' then
  begin
    SetControlvisible('BVC_COMPTAARTICLE',false); SetControlvisible('LCOMPTAARTICLE',false);
    SetControlvisible('BVC_COMPTATIERS',false); SetControlvisible('LCOMPTATIERS',false);
    SetControlvisible('BVC_COMPTAAFFAIRE',false); SetControlvisible('LCOMPTAAFFAIRE',false);
    //
    SetControlText('BVC_COMPTAARTICLE','');
    SetControlText('BVC_COMPTATIERS','');
    SetControlText('BVC_COMPTAAFFAIRE','');
    SetControlvisible('BVC_CODEPORT',true);
    SetControlvisible('LCODEPORT',true);
    //
    THLabel(GetControl('LCODEPORT')).Top := THLabel(GetControl('LCOMPTAARTICLE')).Top;
    THLabel(GetControl('LCODEPORT')).left := THLabel(GetControl('LCOMPTAARTICLE')).left;
    //
    THValComboBox(GetControl('BVC_CODEPORT')).Top := THValComboBox(GetControl('BVC_COMPTAARTICLE')).Top;
    THValComboBox(GetControl('BVC_CODEPORT')).left := THValComboBox(GetControl('BVC_COMPTAARTICLE')).left;
    //
  end else if TypeD = '002' then
  begin
    SetControlvisible('BVC_CODEPORT',false); SetControlText('BVC_CODEPORT',''); SetControlvisible('LCODEPORT',false);
    SetControlvisible('BVC_COMPTAARTICLE',false); SetControlvisible('LCOMPTAARTICLE',false);
    SetControlvisible('BVC_COMPTATIERS',false); SetControlvisible('LCOMPTATIERS',false);
    SetControlvisible('BVC_COMPTAAFFAIRE',false); SetControlvisible('LCOMPTAAFFAIRE',false);
    SetControlText('BVC_COMPTAARTICLE','');
    SetControlText('BVC_COMPTATIERS','');
    SetControlText('BVC_COMPTAAFFAIRE','');
  end else if TypeD = '001' then
  begin
    SetControlvisible('BVC_CODEPORT',false); SetControlText('BVC_CODEPORT',''); SetControlvisible('LCODEPORT',false);
    SetControlvisible('BVC_COMPTAARTICLE',true); SetControlvisible('LCOMPTAARTICLE',true);
    SetControlvisible('BVC_COMPTATIERS',true); SetControlvisible('LCOMPTATIERS',true);
    SetControlvisible('BVC_COMPTAAFFAIRE',true); SetControlvisible('LCOMPTAAFFAIRE',true);
  end;
end;

procedure TOF_BTVENTILCOLL.AfficheGrid;
var II : Integer;
begin
  GS.RowCount := TOBDatas.detail.count +1;
  for II := 0 to TOBDatas.detail.count -1 do
  begin
    TOBDatas.detail[II].PutLigneGrid(GS,II+1,false,false,ListeSaisie)
  end;
  TFVierge(Ecran).HMTrad.ResizeGridColumns(GS);
end;

procedure TOF_BTVENTILCOLL.ConstitueGrid;
begin
  GS.ColCount := 4;
  //
  GS.cells[0,0] := '';
  GS.ColWidths[0] := 10;
  //
  GS.Cells[1,0] := 'Taxe';
  GS.ColWidths[1] := 70;
  GS.ColEditables [1] := false;
  GS.ColLengths [1] := 0;
  //
  GS.Cells[2,0] := 'Coll. Achat';
  GS.ColWidths[2] := 35;
  GS.ColFormats[2]:='CB=BTGCOLLECTIF||<<Aucun>>';
  ColCollectifAc := 2;
  //
  GS.Cells[3,0] := 'Coll. Vente';
  GS.ColWidths[3] := 35;
  GS.ColFormats[3]:='CB=BTGCOLLECTIF||<<Aucun>>';
  ColCollectif := 3;
end;

procedure TOF_BTVENTILCOLL.SetEvents(Etat: boolean);
begin
  GS.OnCellExit := GSCellExit;
  GS.OnCellExit := GSCellExit;
  GS.OnRowExit := GSRowExit;
  GS.OnRowEnter := GSRowEnter;
  TToolbarButton97 (GetControl('Bdelete')).OnClick := BdeleteClick;
end;

procedure TOF_BTVENTILCOLL.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TheValue : string;
begin
  if (ACol <> ColCollectifAc) and (Acol <> ColCollectif) then exit;
  TheValue := GS.CellValues [Acol,Arow];
  if Acol = ColCollectifAc then
  begin
    TOBDatas.detail[Arow-1].SetString('BVC_COLLECTIFAC',GS.CellValues [Acol,Arow]);
    DataModified := True;
  end else if Acol = ColCollectif then
  begin
    TOBDatas.detail[Arow-1].SetString('BVC_COLLECTIF',GS.CellValues [Acol,Arow]);
    DataModified := True;
  end;
end;

procedure TOF_BTVENTILCOLL.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  if (GS.col <> ColCollectifAc) and (GS.col <> ColCollectif) then exit; 
  STCellCur := GS.CellValues [GS.col,GS.row];
  GS.ShowCombo(GS.col,GS.row);
end;

procedure TOF_BTVENTILCOLL.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  TOBDatas.detail[Ou-1].PutLigneGrid(GS,Ou,false,false,ListeSaisie)
end;

procedure TOF_BTVENTILCOLL.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  TOBDatas.detail[GS.row-1].PutLigneGrid(GS,GS.row,false,false,ListeSaisie)
end;

procedure TOF_BTVENTILCOLL.BdeleteClick(sender: TObject);
begin
  if PGIAsk('Etes-vous sur(e) de vouloir supprimer ces informations ?')<> mryes then Exit;
  DeleteDatas;
  DataModified := false;
  Ecran.Close;
end;

Initialization
  registerclasses ( [ TOF_BTVENTILCOLL ] ) ;
end.
