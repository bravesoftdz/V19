{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/07/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTBIMIMPORT ()
Mots clefs ... : TOF;BTBIMIMPORT
*****************************************************************}
Unit BTBIMIMPORT_TOF ;

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
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     HRichEdt,
     HRichOLE,
     HPanel,
     HsysMenu,
     ParamSoc,
     Vierge,
     Windows,
     AGLInit,
     UTOF ;

Type

  TOF_BTBIMIMPORT = Class (TOF)
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
    TOBIMPBIM : TOB;
    procedure AfficheLaGrille;
    procedure GetComponents;
    procedure SetGridStructure;
    procedure BVoirListeClick (Sender : TObject);
    procedure GetInfoFileBIM(NomFic : string);
    procedure AddChampsSup(TOBL: TOB);
    procedure BIntegreFichiers (Sender : TObject);
  end ;

const ZONES = ';LIBTYPE;LIBELLE;CODEAFFAIRE;DESCRIPTIF;NOM';

Implementation
uses UBimetre,Messages,galPatience;

procedure TOF_BTBIMIMPORT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMIMPORT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMIMPORT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMIMPORT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMIMPORT.OnArgument (S : String ) ;
begin
  Inherited ;
  SetControlText('EMPLIMPORT',GetParamSocSecur('SO_BTEMPLBIMIMPORT',''));
  if GetParamSocSecur('SO_BTBIMCLIENTDEF','') = '' then
  begin
    PgiInfo('Merci de renseigner le client / défaut pour l''intégration');
    PostMessage(Ecran.Handle,WM_CLOSE,0,0) ;
    Exit ;
  end;
  //
  TOBIMPBIM := TOB.Create ('LES FICHIERS BIM',nil,-1);
  GetComponents;
  SetGridStructure;
  TToolbarButton97 (GetControl('BVOIRLISTE')).OnClick := BVoirListeClick;
  TToolbarButton97 (GetControl('BLANCEIMPORT')).OnClick := BIntegreFichiers;
end ;

procedure TOF_BTBIMIMPORT.OnClose ;
begin
  TOBIMPBIM.Free;
  Inherited ;
end ;

procedure TOF_BTBIMIMPORT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMIMPORT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMIMPORT.GetComponents;
begin
  GS := Thgrid (GetControl('GS'));
end;

procedure TOF_BTBIMIMPORT.AddChampsSup(TOBL : TOB);
begin
  TOBL.AddChampSupValeur('NOMFIC','');
  TOBL.AddChampSupValeur('TYPE','');
  TOBL.AddChampSupValeur('LIBTYPE','');
  TOBL.AddChampSupValeur('LIBELLE','');
  TOBL.AddChampSupValeur('CODEAFFAIRE','');
  TOBL.AddChampSupValeur('DESCRIPTIF','');
  TOBL.AddChampSupValeur('CODEBBIBLIO','');
  TOBL.AddChampSupValeur('NOTE','');
  TOBL.AddChampSupValeur('DATECREATION',iDate1900);
  TOBL.AddChampSupValeur('DATEDOC',iDate1900);
  TOBL.AddChampSupValeur('CODECLIENT','');
  TOBL.AddChampSupValeur('NOM','');
  TOBL.AddChampSupValeur('ADR','');
  TOBL.AddChampSupValeur('CP','');
  TOBL.AddChampSupValeur('VILLE','');
  TOBL.AddChampSupValeur('AFFAIRE','');
  TOBL.AddChampSupValeur('ANNULATION','-');
end;

procedure TOF_BTBIMIMPORT.SetGridStructure;
begin
  GS.ColCount := 6;
  GS.FixedCols := 1;
  GS.RowCount := 2;
  GS.FixedRows := 1;
  GS.ColWidths [0] := 10;
  GS.Cells[1,0] := 'Type Document';
  GS.ColWidths [1] := 30 * GS.Canvas.TextWidth('W');
  GS.Cells[2,0] := 'Libellé';
  GS.ColWidths [2] := 35 * GS.Canvas.TextWidth('W');
  GS.Cells[3,0] := 'Affaire';
  GS.ColWidths [3] := 70 * GS.Canvas.TextWidth('W');
  GS.Cells[4,0] := 'Descriptif';
  GS.ColWidths [4] := 70 * GS.Canvas.TextWidth('W');
  GS.Cells[5,0] := 'Client';
  GS.ColWidths [5] := 35 * GS.Canvas.TextWidth('W');
end;

procedure TOF_BTBIMIMPORT.BVoirListeClick(Sender: TObject);
var Rec : TSearchRec;
    FilesRec : string;
begin
  GS.VidePile(false);
  SetGridStructure;
  TOBIMPBIM.ClearDetail;
  if not DirectoryExists(GetControlText('EMPLIMPORT')) then
  begin
    PGIInfo('Ce répertoire n''existe pas');
    Exit;
  end;
  FilesRec:= IncludeTrailingBackslash(GetControlText('EMPLIMPORT')) + '\*.bimetre';
  if SysUtils.FindFirst (FilesRec,faAnyFile,Rec) = 0 then
  begin
    repeat
      if (rec.name <> '.') and (rec.name <> '..') then
      begin
        GetInfoFileBIM(Rec.Name);
      end;
    until FindNext(REC) <> 0;
  end;
  SysUtils.FindClose (Rec);
  AfficheLaGrille;
  TFVierge(Ecran).HmTrad.ResizeGridColumns(GS);
  if TOBIMPBIM.Detail.count > 0 then
  begin
    THpanel(GetControl('PBAS')).enabled := True;
  end;
end;

procedure TOF_BTBIMIMPORT.GetInfoFileBIM(NomFic: string);
var TOBL : TOB;
begin
  TOBL := TOB.Create ('UN FICHIER BIM',TOBIMPBIM,-1);
  AddChampsSup(TOBL);
  TOBL.SetString('NOMFIC',NomFic);
  RecupInfosEnteteBIm(GetControlText('EMPLIMPORT'),NomFic,TOBL);
end;

procedure TOF_BTBIMIMPORT.AfficheLaGrille;
var II : Integer;
begin
  GS.RowCount := TOBIMPBIM.detail.count +1; if GS.RowCount = 1 then GS.RowCount := 2;
  For II := 0 to TOBIMPBIM.detail.count -1 do
  begin
    TOBIMPBIM.detail[II].PutLigneGrid(GS,II+1,false,false,ZONES);
  end;
  GS.Row := 1;
end;

procedure TOF_BTBIMIMPORT.BIntegreFichiers(Sender: TObject);

  function DeplaceFichier (Repert : string; TOBL : TOB; status : boolean) : boolean;
  var RepertDest : string;
      FileSrc,FileDest : string;
  begin
    result := True;
    if Status then RepertDest := IncludeTrailingBackslash(Repert)+'OK'
              else RepertDest := IncludeTrailingBackslash(Repert)+'ERR';
    FileSrc := IncludeTrailingBackslash(Repert)+TOBL.GetString('NOMFIC');
    FileDest := IncludeTrailingBackslash(RepertDest)+TOBL.GetString('NOMFIC');
    if not DirectoryExists(RepertDest) then
    begin
      result := CreateDir (RepertDest);
    end;
    if result then Result := MoveFile(PansiChar(FileSrc),PAnsiChar(FileDest));
  end;

  function SelectionneAffaire (TOBL : TOB) : Boolean;
  begin
    Result := True;
    TheTOB := TOBL;
    TOBL.SetBoolean('ANNULATION',false);
    AGLLanceFiche('BTP','BTBIMSELAFF','','','ACTION=MODIFICATION');
    TheTOB := nil;
    if TOBL.GetBoolean('ANNULATION') then Result := false;
  end;

var II : Integer;
    TOBL : TOB;
    XX : TFPatience;
begin
  if GS.nbSelected = 0 then
  begin
    PGIInfo('Vous devez sélectionner les fichiers à importer');
    Exit;
  end;

  if THValComboBox(GetControl('NATUREPIECEG')).Value = '' then
  begin
    PGIInfo('Vous devez définir le type de document à générer');
    THValComboBox(GetControl('NATUREPIECEG')).SetFocus;
    Exit;
  end;

  if GetParamSocSecur('SO_BTARTICLEDIV','') = '' then
  begin
    PgiError('Vous devez renseigner l''article par défaut des appels d''offres');
    exit;
  end;
  //
  THRichEditOLE(GetControl('RAPPORT')).Visible := True;
  THRichEditOLE(GetControl('RAPPORT')).Clear;
  
  TRY
    for II := 1 to GS.RowCount -1 do
    begin
      if GS.IsSelected(GS.row) then
      begin
        TOBL := TOBIMPBIM.detail[GS.row-1];
        if TOBL.getString('TYPE')='1' then
        begin
          if SelectionneAffaire (TOBL) then
          begin
            XX := FenetrePatience('Intégration du fichier BIMétré '+TOBL.GetString('NOMFIC'),aoMilieu, False,true);
            XX.lcreation.visible := false ;
            XX.StartK2000 ;
            XX.Refresh;
            TRY
              if ConstitueDocument(THValComboBox(GetControl('NATUREPIECEG')).Value, GetControlText('EMPLIMPORT'),TOBL,THRichEditOLE(GetControl('RAPPORT'))) then
              begin
                if not DeplaceFichier(GetControlText('EMPLIMPORT'),TOBL,true) then
                begin
                  THRichEditOLE(GetControl('RAPPORT')).lines.Add(Format(' ---- le fichier %s ne sera pas transféré du fait d''un problème de droit d''accès ----- ',[TOBL.GetString('NOMFIC')]));
                end;
              end else
              begin
                THRichEditOLE(GetControl('RAPPORT')).lines.Add(Format(' ---- le fichier %s n''a pu être inétégré ----- ',[TOBL.GetString('NOMFIC')]));
                if not DeplaceFichier(GetControlText('EMPLIMPORT'),TOBL,false) then
                begin
                  THRichEditOLE(GetControl('RAPPORT')).lines.Add(Format(' ---- le fichier %s ne sera pas transféré du fait d''un problème de droit d''accès ----- ',[TOBL.GetString('NOMFIC')]));
                end;
              end;
            FINALLY
              XX.StopK2000 ;
              XX.free;
            END;
          end else
          begin
            THRichEditOLE(GetControl('RAPPORT')).lines.Add(Format(' ---- Importation du fichier %s annulé ----- ',[TOBL.GetString('NOMFIC')]));
          end;
        end;
      end;
    end;
  FINALLY
    GS.ClearSelected;
    BVoirListeClick(self); 
  end;
end;


Initialization
  registerclasses ( [ TOF_BTBIMIMPORT ] ) ;
end.
