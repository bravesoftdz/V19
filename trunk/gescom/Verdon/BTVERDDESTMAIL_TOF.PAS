{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/09/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTVERDDESTMAIL ()
Mots clefs ... : TOF;BTVERDDESTMAIL
*****************************************************************}
Unit BTVERDDESTMAIL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul,FE_Main,
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
     AglInit,
     Lookup,
     Vierge,
     HTB97 ;

Const
  LesColonnes = ' ;BV0_UTILISAT;US_LIBELLE;US_EMAIL;';
Type

  TOF_BTVERDDESTMAIL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    GS : THGrid;
    fListDest : TOB;
    fStructTOB : TOB;
    procedure ConstitueGrid;
    procedure RemplitTOb;
    procedure AfficheDatas;
    procedure BinsertClick (Sender : TObject);
    procedure BDeleteClick (Sender : TObject);
    procedure GSDblClick(Sender: TObject);
    procedure memoriseChampsSup;
    procedure AddChampsSup(TOBL: TOB);
    procedure SetInfoSup(TT : TOB);
    function addTOBDESTMAIL: TOB;
    procedure AfficheLaLigne(TOBL: TOB; Lig: Integer);
    function controleEmailOk : boolean;
    procedure NumeroteInMemory;
    procedure DeleteOld;
    procedure WriteAll;
    function AllreadyPresent (User : string) : boolean;
  end ;

Implementation
uses UtilchampsSup;

procedure TOF_BTVERDDESTMAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTVERDDESTMAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTVERDDESTMAIL.OnUpdate ;
begin
  Inherited ;
  if not controleEmailOk then
  begin
    if PgiAsk ('ATTENTION : Certains utilisateur n''ont pas de Mail définis.#13#10 Confirmez-vous la validation?') <> Mryes then
    begin
      ecran.ModalResult := 0;
      Exit;
    end;
  end;
  NumeroteInMemory;
  BEGINTRANS;
  TRY
    DeleteOld;
    WriteAll;
    COMMITTRANS;
  EXCEPT
    ROLLBACK;
  end;

end ;

procedure TOF_BTVERDDESTMAIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTVERDDESTMAIL.OnArgument (S : String ) ;
begin
  Inherited ;
  GS := THGrid(GetControl('GS'));
  fListDest := TOB.Create ('LES DESTINATAIRES',nil,-1);
  fStructTOB := TOB.Create ('LA STRUCT',nil,-1);
  MemoriseChampsSup;
  ConstitueGrid;
  RemplitTOb;
  AfficheDatas;
  TToolbarButton97 (GetControl('Binsert')).OnClick := BinsertClick;
  TToolbarButton97 (GetControl('Bdelete')).OnClick := BDeleteClick;
  GS.OnDblClick := GSDblClick;
end ;

procedure TOF_BTVERDDESTMAIL.OnClose ;
begin
  fListDest.Free;
  fStructTOB.Free;
  Inherited ;
end ;

procedure TOF_BTVERDDESTMAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTVERDDESTMAIL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTVERDDESTMAIL.RemplitTOb;
var QQ : TQuery;
begin
  QQ := OpenSQL('SELECT * FROM BTVERDDESTMAIL',True,-1,'',True);
  if not QQ.eof then
  begin
    fListDest.LoadDetailDB('BVERDESTMAIL','','',QQ,true);
  end;
  Ferme(QQ);
  //
  if fListDest.detail.count = 0 then
  begin
    addTOBDESTMAIL;
  end;
end;

procedure TOF_BTVERDDESTMAIL.AddChampsSup (TOBL : TOB);
var Indice : integer;
		TOBEtl : TOB;
    NomChamps : string;
begin
	if fStructTOB.detail.count = 0 then exit;
  TOBEtl := fStructTOB.detail[0];
  //
  for Indice := 0 to TOBETL.NombreChampSup -1 do
  begin
  	NomChamps := TOBEtl.GetNomChamp(1000+Indice);
    if not TOBL.FieldExists(NomChamps) then
    begin
  		TOBL.AddChampSupValeur (NomChamps,TOBEtl.Getvaleur(1000+Indice));
    end;
  end;
end;


procedure TOF_BTVERDDESTMAIL.AfficheDatas;
var II : Integer;
begin
  GS.RowCount := fListDest.detail.count +1; 
  for II := 0 to fListDest.detail.count -1 do
  begin
    AfficheLaLigne (fListDest.detail[II],II+1);
  end;
  TFVierge(ecran).HMTrad.ResizeGridColumns(GS);
end;

procedure TOF_BTVERDDESTMAIL.AfficheLaLigne(TOBL: TOB; Lig: Integer);
begin
  TOBL.PutLigneGrid(GS,Lig,False,False,LesColonnes);
end;

procedure TOF_BTVERDDESTMAIL.ConstitueGrid;
begin
  GS.ColCount := 4;
  //
  GS.Cells[0,0] := '';
  GS.ColWidths[1]:=20;
  //
  GS.Cells[1,0] := 'Utilisateur';
  GS.ColWidths[1]:=10 * GS.Canvas.TextHeight ('W');
  GS.ColAligns[1]:=taLeftJustify ;
  GS.ColLengths[1]:=10;
  //
  GS.Cells[2,0] := 'Nom';
  GS.ColWidths[2]:=35 * GS.Canvas.TextHeight ('W');
  GS.ColAligns[2]:=taLeftJustify ;
  GS.ColLengths[2]:=35;
  //
  GS.Cells[3,0] := 'Adresse Mail';
  GS.ColWidths[3]:=70 * GS.Canvas.TextHeight ('W');
  GS.ColAligns[3]:=taLeftJustify ;
  GS.ColLengths[3]:=70;
end;

procedure TOF_BTVERDDESTMAIL.BinsertClick(Sender: TObject);
var TT : TOB;
begin
  THEdit(GetControl('NEWDEST')).Text := '';
  LookupList (THEdit(GetControl('NEWDEST')),'Sélection destinataire','UTILISAT','US_UTILISATEUR','US_LIBELLE','','US_UTILISATEUR',false,3170);
  if THEdit(GetControl('NEWDEST')).Text <> '' then
  begin
    if AllreadyPresent(THEdit(GetControl('NEWDEST')).Text) then
    begin
      PgiError ('IMPOSSIBLE : Cet utilisateur est déjà présent dans la liste');
      exit;
    end;
    if fListDest.detail.count = 0 then
    begin
      TT := addTOBDESTMAIL;
    end else
    begin
      TT :=fListDest.detail[fListDest.detail.count-1];
    end;
    if TT.GetString('BV0_UTILISAT')<>'' then
    begin
      TT := addTOBDESTMAIL;
    end;
    TT.setString('BV0_UTILISAT',THEdit(GetControl('NEWDEST')).Text);
    SetInfoSup(TT);
    AfficheDatas;
  end;
end;

procedure TOF_BTVERDDESTMAIL.BDeleteClick(Sender: TObject);
begin
  if PgiAsk('Etes-vous sur de supprimer cet utilisateur de la liste ?')<> mryes then exit; 
  fListDest.detail[GS.row-1].free;
  GS.DeleteRow(GS.row);
end;

function TOF_BTVERDDESTMAIL.addTOBDESTMAIL : TOB;
begin
  result := TOB.Create ('BVERDESTMAIL',fListDest,-1);
  result.setString('BV0_UTILISAT','');
  AddChampsSup (result);
end;

procedure TOF_BTVERDDESTMAIL.memoriseChampsSup;
var TOBEtalon : TOB;
		Indice : integer;
    NomChamps : string;
    QQ : TQuery;
begin
	if fStructTOB.detail.count > 0 then fStructTOB.clearDetail;
  QQ := OpenSQL('SELECT * FROM BTVERDDESTMAIL WHERE 1=2',True,-1,'',True);
  //
	TOBEtalon := TOB.Create ('LA LIGNE ETALON',fStructTOB,-1);
  //
  for Indice := 0 to QQ.FieldCount -1 do
  begin
  	NomChamps := QQ.Fields[Indice].FieldName;
		if ExtractPrefixe (NomChamps) = 'BV0' then continue; // pour ne prendre que les champs sup
  	TOBEtalon.AddChampSupValeur (NomChamps,getValInitChampsSup(NomChamps));
  end;
  Ferme(QQ);
end;


procedure TOF_BTVERDDESTMAIL.SetInfoSup(TT: TOB);
var QQ: TQuery;
    II : Integer;
begin
    QQ := OpenSQL('SELECT * FROM UTILISAT WHERE US_UTILISATEUR="'+TT.GetString('BV0_UTILISAT')+'"',True,1,'',true);
    if not QQ.eof then
    begin
      for II := 0 to QQ.FieldCount -1 do
      begin
        if TT.FieldExists(QQ.Fields [II].FieldName) then
        begin
          TT.PutValue(QQ.Fields [II].FieldName,QQ.Fields [II].Value);
        end;
      end;
    end;
    Ferme(QQ);
end;

function TOF_BTVERDDESTMAIL.controleEmailOk: boolean;
var I : Integer;
begin
  Result := True;
  for I :=0 to fListDest.Detail.count -1 do
  begin
    if fListDest.detail[I].GetString('US_EMAIL') = '' then
    begin
      Result := false;
      break;
    end;
  end;
end;

procedure TOF_BTVERDDESTMAIL.NumeroteInMemory;
var I : Integer;
begin
  for I := 0 to fListDest.detail.count -1 do
  begin
    fListDest.detail[I].SetInteger('BV0_ID',I+1);  
  end;
end;

procedure TOF_BTVERDDESTMAIL.DeleteOld;
begin
  ExecuteSQL('DELETE FROM BVERDESTMAIL');
end;

procedure TOF_BTVERDDESTMAIL.WriteAll;
var I : Integer;
begin
  if fListDest.Detail.count = 0 then Exit; 
  for I := 0 to fListDest.detail.count -1 do fListDest.detail[I].SetAllModifie(true);
  if not fListDest.InsertDB(nil) then V_PGI.Ioerror := oeUnknown; 
end;

procedure TOF_BTVERDDESTMAIL.GSDblClick(Sender: TObject);
var user : string;
    TT : TOB;
begin
  TT := fListDest.detail[GS.row-1];
  User := TT.GetString('BV0_UTILISAT');
  AglLAnceFiche('YY','YYUTILISAT','',User,'ACTION=MODIFICATION');
  SetInfoSup (TT);
  AfficheLaLigne(TT,GS.row);
end;

function TOF_BTVERDDESTMAIL.AllreadyPresent(User: string): boolean;
var II : Integer;
begin
  Result := false;
  for II :=0  to fListDest.detail.count -1 do
  begin
    if fListDest.detail[II].GetString('BV0_UTILISAT')=User then
    begin
      Result := True;
      break;
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTVERDDESTMAIL ] ) ;
end.
