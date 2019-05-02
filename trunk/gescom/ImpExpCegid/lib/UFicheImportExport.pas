unit UFicheImportExport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, ExtCtrls, StdCtrls, Mask, Hctrls, ComCtrls, TntComCtrls,
  UTOF,
  UTOZ,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  HEnt1,
  HMsgBox,
  ed_tools,
  DateUtils,
  UTOB, TntExtCtrls, HPanel, Grids, TntGrids, HSysMenu, TntStdCtrls;
  
const
    MaxFileSize = 200;

type
  TmodeTrait = (tmtExport,TmtImport);
  
  TFexpImpCegid = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    BImprimer: TToolbarButton97;
    PGCTRL: THPageControl2;
    TPSHTCAR: TTabSheet;
    TBSHTCONTROL: TTabSheet;
    EXPORTFIC: THCritMaskEdit;
    Label1: TLabel;
    NOMFIC: THCritMaskEdit;
    Label2: TLabel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    RDTEXPORT: TRadioButton;
    Label3: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label4: TLabel;
    RDTIMPORT: TRadioButton;
    GroupBox1: TGroupBox;
    Label8: TLabel;
    IMPORTFIC: THCritMaskEdit;
    CHKZIP: TCheckBox;
    GroupBox2: TGroupBox;
    Label9: TLabel;
    CHKREPERT: TCheckBox;
    IMPORTDIR: THCritMaskEdit;
    CHKMSGSTRUCT: TCheckBox;
    CHKSTOPONERROR: TCheckBox;
    CBVIDAGEEXP: TCheckBox;
    CBVIDAGEIMP: TCheckBox;
    CBCOMPTA: TCheckBox;
    CBPAYE: TCheckBox;
    CBDOSSIER: TCheckBox;
    HPanel1: THPanel;
    BLANCETRAIT: TToolbarButton97;
    GS: THGrid;
    hmtrad: THSystemMenu;
    RAPPORT: TTabSheet;
    TRACE: TListBox;
    TX: THCritMaskEdit;
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RDTEXPORTClick(Sender: TObject);
    procedure CHKZIPClick(Sender: TObject);
    procedure CHKREPERTClick(Sender: TObject);
    procedure RDTIMPORTClick(Sender: TObject);
    procedure CBCOMPTAClick(Sender: TObject);
    procedure CBPAYEClick(Sender: TObject);
    procedure CBDOSSIERClick(Sender: TObject);
    procedure BLANCETRAITClick(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
    procedure TXExit(Sender: TObject);
  private
    T_Mere : TOB;
    T : TOB;
    fRepert : string;
    ZipFile : string;
    { Déclarations privées }
    fModeTrait : TmodeTrait;
    WithOutZip : boolean;
    ExpDir,ExpFile : string;
    procedure ExportDossier;
    procedure ImportDOS(FileN,ExportN: string);
    function RendSQLTable (TraitExport : Boolean=False): string;
    function PGCreatZipFile(Archive: string; CodeSession: TModeOpenZip;var TheTOZ: TOZ; Ecran: TFORM): Boolean;
    function TraitementTable(TheTOZ: TOZ; MaTable,Domaine, Predefini, Prefixe: string; Numversion: integer; DDebut, DFin: TdateTime; CodeEx,LibEx: string; Pagine,premier: boolean ; NumPage,NbEnreg,CurrentLine,NbRecords,NbPage : integer): boolean;
    function PGZipFile(Fichier: string; TheTOZ: TOZ;Ecran: TFORM): Boolean;
    procedure PGFermeZipFile(TheTOZ: TOZ);
    procedure PGVideDirectory(FileN: string);
    function ExportComptaOnly: Boolean;
    function ExportPaieCompta: Boolean;
    function ExportPaieOnly: Boolean;
    function GetStoredProcedureName(TableName : string) : string;
    procedure CreateStoredProcedure(TableName, SortedFields : string);
    function GetProcessMemorySize(_sProcessName: string; var _nMemSize: Cardinal): Boolean;
    procedure AddTrace(Text : string);
    procedure PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
  public
    { Déclarations publiques }
  end;


procedure LanceImportExport;

implementation

uses
  CommonTools
  , Math
  , psAPI
  ;

{$R *.dfm}

procedure LanceImportExport;
  var FexpImpCegid: TFexpImpCegid;
begin
  FexpImpCegid := TFexpImpCegid.create(Application);
  TRY
    FexpimpCegid.ShowModal;
  FINALLY
    FexpImpCegid.free;
  END;
end;

procedure TFexpImpCegid.BValiderClick(Sender: TObject);
begin
  GS.RowCount := 1;
  WithOutZip := false;
  if RDTEXPORT.Checked then
  begin
    ExpDir := EXPORTFIC.Text;
    ExpFile := trim(NOMFIC.text);

    if ExpDir = '' then
    begin
      PgiBox('Vous n''avez pas renseigné le réperoire de transfert !', self.caption);
      ModalResult := mrNone;
      exit;
    end;
    if (not CBPAYE.Checked) and (not CBCOMPTA.Checked) and (not CBDOSSIER.Checked ) then
    begin
      PgiBox('Merci de définir le type d''export.', self.caption);
      ModalResult := mrNone;
      exit;
    end;
    if ExpFile = '' then
    begin
      if PGIAsk('Vous n''avez pas renseigné le nom du fichier archive !#13#10 Confirmez-vous ?') <> mrYes then
      begin
        ModalResult := mrNone;
        exit;
      end;
      if CBVIDAGEEXP.Checked then
      begin
        PGIInfo('VOus ne pouvez pas vider le répertoire sans avoir renseigné un fichier de destination');
        ModalResult := mrNone;
        exit;
      end;
      WithOutZip := true;
    end;
    PGCTRL.ActivePage := TBSHTCONTROL;
    ExportDossier;
  end
  else
  begin
    if PgiAsk('Attention: Cet import écrasera de manière définitive les données de ce dossier.#13#10 Voulez vous continuer le traitement ?', self.caption) <> Mryes then Exit;
    if (IMPORTFIC.text = '') and (CHKZIP.checked) then
    begin
      PgiError('Vous devez renseigner le nom du fichier à importer !', Self.caption);
      exit;
    end;
    if (IMPORTDIR.text = '') and (CHKREPERT.checked) then
    begin
      PgiError('Vous devez renseigner le nom du répertoire ou les fichiers sont stockés !', Self.caption);
      exit;
    end;
    if (PGCTRL  <> nil) and (TBSHTCONTROL <> nil) then PGCTRL.ActivePage := TBSHTCONTROL;
    //
    IMPORTDOS(IMPORTFIC.Text,IMPORTDIR.Text);
  end;
  ModalResult := 0;
end;

procedure TFexpImpCegid.BFermeClick(Sender: TObject);
begin
  close;
end;

procedure TFexpImpCegid.BImprimerClick(Sender: TObject);
begin
//
end;

procedure TFexpImpCegid.HelpBtnClick(Sender: TObject);
begin
//
end;

procedure TFexpImpCegid.FormCreate(Sender: TObject);
begin
  T_Mere := TOB.Create('Ma TOB', nil, -1);
  T := TOB.Create('La TOB des tables', nil, -1);
end;

procedure TFexpImpCegid.FormShow(Sender: TObject);
begin
//
	CHKZIP.Enabled := false;
  CHKREPERT.Enabled := false;
  IMPORTFIC.Enabled := false;
  IMPORTDIR.Enabled := false;
  PGCTRL.ActivePage := TPSHTCAR;
  GS.PostDrawCell := PostDrawCell;
end;

procedure TFexpImpCegid.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  T_Mere.free;
  T.free;
end;

procedure TFexpImpCegid.RDTEXPORTClick(Sender: TObject);
begin
  CHKMSGSTRUCT.Enabled := false;
  CHKSTOPONERROR.Enabled := false;
  CHKZIP.Enabled := false;
  CHKREPERT.Enabled := false;
  CBVIDAGEIMP.Enabled := false;
  IMPORTFIC.Enabled := false;
  IMPORTDIR.Enabled := false;
  //
  CBVIDAGEEXP.Enabled := true;
  NOMFIC.Enabled := true;
  EXPORTFIC.Enabled := true;
  CBCOMPTA.enabled := true;
  CBPAYE.enabled := true;

end;

procedure TFexpImpCegid.CHKZIPClick(Sender: TObject);
begin
  if CHKZIP.Checked then
  begin
    CHKREPERT.checked := false;
    IMPORTFIC.enabled := true;
    IMPORTDIR.Enabled := false;
  end else
  begin
    CHKREPERT.checked := true;
    IMPORTFIC.enabled := false;
    IMPORTDIR.Enabled := true;
  end;

end;

procedure TFexpImpCegid.CHKREPERTClick(Sender: TObject);
begin
  if CHKREPERT.checked then
  begin
    CHKZip.Checked := false;
    IMPORTDIR.enabled := true;
  end else
  begin
    CHKZip.Checked := true;
    IMPORTDIR.enabled := false;
  end;
end;


procedure TFexpImpCegid.ExportDossier;
var
  T1: TOB;
  St : string;
  Q: TQuery;
  Nbre, II : integer;
  TotLignes : integer;
  NbRecords, NbrecordsOk : Integer;
  TableSize : double;
  NomTable : string;
begin
  //
  BValider.Visible := false;
  GS.ColCount := 2;
  GS.Cells[0,0] := 'Liste des tables à traiter';
  GS.Cells[0,1] := 'Nb Enreg/Partie';
  TRACE.Items.Clear;
  fModeTrait := tmtExport;
  frepert := ExpDir;
  ZipFile := IncludeTrailingPathDelimiter (ExpDir) + ExpFile + '.zip';
  T.ClearDetail;
  //
  TRY
    st := RendSQLTable (True);
    Q := OPENSQL(ST, TRUE);
    try
      T.LoadDetailDb('DETABLES', '', '', Q, FALSE);
    finally
      Ferme(Q);
    end;
    TotLignes := pred(T.detail.count);
    //
    II := 0;
    InitMoveProgressForm(nil, 'Recherche des tables', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
    try
      repeat
        MoveCurProgressForm(Format('Recherche des tables (%s/%s)', [IntToStr(II), IntToStr(TotLignes)]));
        T1 := T.Detail[II];
        if T1 <> nil then
        begin
          st := 'SELECT COUNT (*) NBRE FROM ' + T1.GetValue('DT_NOMTABLE');
          Q := OPENSQL(ST, TRUE);
          if not Q.EOF then
          begin
            Nbre := Q.FindField('NBRE').AsInteger;
            if Nbre <= 0 then
            begin
              T1.Free
            end else
            begin
              Inc(II);
            end;
          end;
          Ferme(Q);
        end;
      until II > T.detail.count -1;
    finally
      FiniMoveProgressForm;
    end;
    GS.rowCount := T.detail.count  +1;
    GS.FixedRows := 1;
    InitMoveProgressForm(nil, 'Calcul de la taille des tables.', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
    try
      for II := 0 to T.detail.count -1 do
      begin
        MoveCurProgressForm(Format('Calcul de la taille des tables (%s/%s)', [IntToStr(II), IntToStr(pred(T.detail.count))]));
        T1 := T.detail[II];
        //
        NomTable := T1.GetString('DT_NOMTABLE');
        TableSize   := Tools.GetTableSize(NomTable);
        NbRecords   := Tools.GetTableInf(titRecordsQty, NomTable);
        T1.SetInteger('NBENREG',-1);
        T1.SetInteger('NBENREGTOTAL',NbRecords);
        if TableSize > MaxFileSize then
        begin
          if TableSize > MaxFileSize then
          begin
            NbrecordsOk := Round(NbRecords / (TableSize / MaxFileSize));
          end else
          begin
            NbrecordsOk := NbRecords;
          end;
          if NbrecordsOk > 200000 then NbrecordsOk := 200000;
          (*
          NbrecordsOk := Round(TableSize/MaxFileSize);
          if NbrecordsOk > 500000 then NbrecordsOk := 500000;
          *)
          T1.SetInteger('NBENREG',NbrecordsOk);
          T1.SetInteger('NBTOTAL',NbRecords);
        end;
        //
        T1.PutLigneGrid(GS,II+1,false,false,'DT_NOMTABLE;NBENREG;');
        GS.Objects[0,II+1] := T1;
      end;
    finally
      FiniMoveProgressForm;
    end;
    hmtrad.ResizeGridColumns(GS);
    GS.AllSelected := True;
    BValider.Visible := false;
    BLANCETRAIT.Visible := True;
  FINALLY
  END;
end;

procedure TFexpImpCegid.ImportDOS(FileN,ExportN: string);
var
  rep,  Numvers, LaVersion: Integer;
  T1, T2 : TOB;
  Q: TQUERY;
  LeNom,  st, MonFic: string;
  TheTOZ: TOZ;
  ret : Integer;
  sr: TsearchRec;

begin
  TRY
    fModeTrait := TmtImport;
    T.ClearDetail;
    T_Mere.ClearDetail;
    //
    if FileN<> '' then frepert := ExtractFilePath(FileN)
                  else fRepert := IncludeTrailingBackslash (ExportN);
    TheTOZ := nil;
    if (FileN <> '') then
    begin
      AddTrace('Phase 1 - unzip du fichier...');

      if not PGCreatZipFile(FileN, moOpen, TheTOZ, Self) then
      begin
        PgiError('Traitement abandonné', Self.caption);
        exit;
      end
      else PGFermeZipFile(TheTOZ);
      AddTrace('Phase 1 - unzip du fichier...OK');
    end;
    //
    ST := RendSQLTable;
    Q := OPENSQL(ST, TRUE);
    T.LoadDetailDb('DETABLES', '', '', Q, FALSE);
    Ferme(Q);
    //
    InitMoveProgressForm(nil, 'Analyse des données', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
    ret := FindFirst(fRepert + '*.SEB', 0 + faAnyFile, sr);
    TRY
      Rep := MrYes;
      while (ret = 0) and (rep = mrYes) do
      begin
        //
        T_Mere.ClearDetail;
        //
        MonFic := fRepert + sr.Name;
        MoveCurProgressForm('Analyse des données du fichier ' + sr.Name);
        //
        TOBLoadFromBinFile(MonFic, nil, T_Mere);
        T1 := T_Mere.Detail[0];
        if Assigned(T1) then
        begin
          LeNom := T1.GetValue('TABLE');
          LaVersion := T1.GetValue('VERSION');
          T2 := T.FindFirst(['DT_NOMTABLE'], [LeNom], FALSE);
          if T2 <> nil then
          begin
            if (T2.getString('VERIFIED')<> 'X') then
            begin
              NumVers := T2.GetValue('DT_NUMVERSION');
              if (LaVersion <> NumVers) and (NumVers <> 0) and (T2.getString('VERIFIED')<> 'X') and (CHKMSGSTRUCT.checked) then
              begin
                if PgiAsk ('ATTENTION : la structure d''origine de la table '+LeNom+' n''est pas identique à celle de ce dossier.#13#10 Confirmez-vous quand même l''import ?',self.caption) <> mryes then
                begin
                  rep := MrCancel;
                end else T2.SetString('VERIFIED','X');
              end;
            end;
          end;
        end;
        ret := FindNext(sr);
      end;
    FINALLY
      sysutils.FindClose(sr);
    END;

    if Rep <> MrYes then
    begin
      PgiError('Des anomalies de structure ont été détectées ou bien vous avez abandonné le traitement', self.caption);
      FiniMoveProgressForm();
      AddTrace('Traitement abandonné');
      exit;
    end else
    begin
      FiniMoveProgressForm();
    end;
    BValider.Visible := false;
    BLANCETRAITClick(self);
    //
  FINALLY
  END;
end;

function TFexpImpCegid.PGCreatZipFile(Archive: string;CodeSession: TModeOpenZip; var TheTOZ: TOZ; Ecran: TFORM): Boolean;
begin
  result := false;
  TheTOZ := nil;
  if WithOutZip then BEGIN Result := True; Exit; end;
  TheToz := TOZ.Create;
  try
    if TheToz.OpenZipFile(Archive, CodeSession) then
    begin
      if CodeSession = moCreate then
        if TheToz.OpenSession(osAdd) then result := TRUE
        else HShowMessage('0;Erreur;Soit le fichier : ' + Archive + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '');

      if CodeSession = moOpen then
        if TheToz.OpenSession(osExt) then result := TRUE
        else HShowMessage('0;Erreur;Soit le fichier : ' + Archive + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '');
    end
    else
    begin
      HShowMessage('0;Erreur;Erreur création du fichier archive : ' + Archive + ' impossible;E;O;O;O', '', '');
      Exit;
    end;
  except
    on E: Exception do
    begin
      PgiError('TozError : ' + E.Message, Ecran.Caption);
      TheToz.Free;
    end;
  end;
end;

procedure TFexpImpCegid.PGFermeZipFile(TheTOZ: TOZ);
begin
  if WithOutZip then exit;
  TheToz.CloseSession;
end;

procedure TFexpImpCegid.PGVideDirectory(FileN: string);
var
  sr: TsearchRec;
  ret: Integer;
  MonFic: string;
begin

  ret := FindFirst(FileN + '\*.TEB', 0 + faAnyFile, sr);
  while ret = 0 do
  begin
    MonFic := FileN + '\' + sr.Name;
    if FileExists(MonFic) then DeleteFile(PChar(MonFic));
    ret := FindNext(sr);
  end;
  sysutils.FindClose(sr);
  ret := FindFirst(FileN + '\*.SEB', 0 + faAnyFile, sr);
  while ret = 0 do
  begin
    MonFic := FileN + '\' + sr.Name;
    if FileExists(MonFic) then DeleteFile(PChar(MonFic));
    ret := FindNext(sr);
  end;
  sysutils.FindClose(sr);
end;

function TFexpImpCegid.PGZipFile(Fichier: string; TheTOZ: TOZ;Ecran: TFORM): Boolean;
begin
  if WithOutZip then BEGIN RESULT := True;  Exit; END;
  result := false;
  try
    if TheToz.ProcessFile(Fichier, '') then result := TRUE;
  except
    on E: Exception do
    begin
      PgiError('TozError : ' + E.Message, Ecran.caption);
      TheToz.Free;
    end;
  end;
end;

function TFexpImpCegid.ExportPaieCompta : Boolean;
begin
  Result := CBDOSSIER.Checked;
end;

function TFexpImpCegid.ExportComptaOnly : Boolean;
begin
  Result := (not CBPAYE.Checked) and (CBCOMPTA.Checked);
end;

function TFexpImpCegid.ExportPaieOnly : Boolean;
begin
  Result := (CBPAYE.Checked) and (not CBCOMPTA.Checked);
end;

function TFexpImpCegid.RendSQLTable (TraitExport : Boolean=False): string;

  function GetStartQry : string;
  begin
    Result := 'SELECT DT_NOMTABLE'
            + '     , DT_PREFIXE'
            + '     , DT_LIBELLE'
            + '     , DT_NUMVERSION'
            + '     , DT_DOMAINE'
            + '     , DT_CLE1'
            + '     , 0 AS NBENREG '
            + '     , 0 AS NBTOTAL '
            + '     , 0 AS NBENREGTOTAL '
            + '     , "-" AS REINIT'
            + '     , "-" AS VERIFIED'
            + '     , IIF(EXISTS(SELECT DH_NOMCHAMP FROM DECHAMPS WHERE DH_NOMCHAMP=DT_PREFIXE||"_PREDEFINI"),"X","-") AS PREDEFINI'
            + ' FROM DETABLES'
            + ' WHERE '
            ;
  end;

begin
  if (not TraitExport) or ((TraitExport) and (ExportPaieCompta)) then
  begin
    result := GetStartQry
            + '( '
            + '(DT_DOMAINE IN ("C","Y","T","D","0","P")) OR '
            + '(DT_NOMTABLE = "CHOIXCOD") OR '
            + '(DT_NOMTABLE = "MENU") OR '
            + '(DT_NOMTABLE = "PARAMSOC") OR '
            + '(DT_NOMTABLE = "ETABLISS") OR '
            + '(DT_NOMTABLE = "ETABLCOMPL") OR '
            + '(DT_NOMTABLE = "SOCIETE") OR '
            + '(DT_NOMTABLE = "DESEQUENCES") OR '
            + '(DT_NOMTABLE = "LIENSOLE")'
            + ') AND '
            + 'DT_NOMTABLE <> "YMULTIDOSSIER" AND '
            + 'DT_NOMTABLE <> "DOSSIER" AND '
            + 'DT_NOMTABLE NOT LIKE "TRAD%" AND '
            + 'DT_NOMTABLE <> "LOG" AND '
            + 'DT_NOMTABLE <> "LIENDONNEES" AND '
            + 'DT_NOMTABLE <> "RESSOURCEPR" AND '
            + 'DT_NOMTABLE <> "COMMUN" AND '
            + 'DT_NOMTABLE <> "CLAVIERECRAN" AND '
            + 'DT_NOMTABLE <> "EVARTRACK" AND '
            + 'DT_NOMTABLE <> "JNALEVENT" AND '
            + 'DT_NOMTABLE <> "MACRO" AND '
            + 'DT_NOMTABLE <> "MODEDATA" AND '
            + 'DT_NOMTABLE <> "COURRIER" AND '
            + 'DT_NOMTABLE <> "MODELES" AND '
            + 'DT_NOMTABLE <> "RTDOCUMENT" AND '
            + 'DT_NOMTABLE <> "SOUCHE" AND '
            + 'DT_NOMTABLE <> "YMYBOBS" AND '
            + 'DT_NOMTABLE <> "YMYCPTX" AND '
            + 'DT_NOMTABLE <> "FORMESP"'
            + 'ORDER BY DT_NOMTABLE'
            ;
  end else if (ExportComptaOnly) then
  begin
    result := GetStartQry
            + '( '
            + '(DT_DOMAINE IN ("C","Y","T","D","0")) OR '
            + '(DT_NOMTABLE = "CHOIXCOD") OR '
            + '(DT_NOMTABLE = "MENU") OR '
            + '(DT_NOMTABLE = "PARAMSOC") OR '
            + '(DT_NOMTABLE = "ETABLISS") OR '
            + '(DT_NOMTABLE = "ETABLCOMPL") OR '
            + '(DT_NOMTABLE = "SOCIETE") OR '
            + '(DT_NOMTABLE = "DESEQUENCES") OR '
            + '(DT_NOMTABLE = "LIENSOLE")'
            + ') AND '
            + 'DT_NOMTABLE <> "YMULTIDOSSIER" AND '
            + 'DT_NOMTABLE <> "DOSSIER" AND '
            + 'DT_NOMTABLE NOT LIKE "TRAD%" AND '
            + 'DT_NOMTABLE <> "LOG" AND '
            + 'DT_NOMTABLE <> "LIENDONNEES" AND '
            + 'DT_NOMTABLE <> "RESSOURCEPR" AND '
            + 'DT_NOMTABLE <> "COMMUN" AND '
            + 'DT_NOMTABLE <> "CLAVIERECRAN" AND '
            + 'DT_NOMTABLE <> "EVARTRACK" AND '
            + 'DT_NOMTABLE <> "JNALEVENT" AND '
            + 'DT_NOMTABLE <> "MACRO" AND '
            + 'DT_NOMTABLE <> "MODEDATA" AND '
            + 'DT_NOMTABLE <> "COURRIER" AND '
            + 'DT_NOMTABLE <> "MODELES" AND '
            + 'DT_NOMTABLE <> "RTDOCUMENT" AND '
            + 'DT_NOMTABLE <> "SOUCHE" AND '
            + 'DT_NOMTABLE <> "YMYBOBS" AND '
            + 'DT_NOMTABLE <> "YMYCPTX" AND '
            + 'DT_NOMTABLE <> "FORMESP"'
            + 'ORDER BY DT_NOMTABLE'
            ;
  end else if (ExportPaieOnly) then
  begin
    Result := GetStartQry
            +  '('
            +  '(DT_DOMAINE IN ("P","Y")) OR '
            +  '(DT_NOMTABLE = "CHOIXCOD") OR (DT_NOMTABLE = "PARAMSOC")  OR (DT_NOMTABLE = "ETABLISS") OR '
            +  '(DT_NOMTABLE = "CALENDRIER") OR (DT_NOMTABLE = "JOURFERIE") OR '
            +  '(DT_NOMTABLE = "TIERS") OR (DT_NOMTABLE = "RIB") OR '
            +  '(DT_NOMTABLE = "ANNUAIRE") OR '
            +  '(DT_NOMTABLE = "BANQUECP") OR '
            +  '(DT_NOMTABLE = "VENTIL") OR (DT_NOMTABLE = "VENTANA") OR (DT_NOMTABLE = "AXE") OR (DT_NOMTABLE = "SECTION")'
            +  ') AND '
            + 'DT_NOMTABLE <> "YMULTIDOSSIER" AND '
            + 'DT_NOMTABLE <> "DOSSIER" AND '
            + 'DT_NOMTABLE NOT LIKE "TRAD%" AND '
            + 'DT_NOMTABLE <> "LOG" AND '
            + 'DT_NOMTABLE <> "LIENDONNEES" AND '
            + 'DT_NOMTABLE <> "RESSOURCEPR" AND '
            + 'DT_NOMTABLE <> "COMMUN" AND '
            + 'DT_NOMTABLE <> "CLAVIERECRAN" AND '
            + 'DT_NOMTABLE <> "EVARTRACK" AND '
            + 'DT_NOMTABLE <> "JNALEVENT" AND '
            + 'DT_NOMTABLE <> "MACRO" AND '
            + 'DT_NOMTABLE <> "MODEDATA" AND '
            + 'DT_NOMTABLE <> "COURRIER" AND '
            + 'DT_NOMTABLE <> "MODELES" AND '
            + 'DT_NOMTABLE <> "RTDOCUMENT" AND '
            + 'DT_NOMTABLE <> "SOUCHE" AND '
            + 'DT_NOMTABLE <> "YMYBOBS" AND '
            + 'DT_NOMTABLE <> "YMYCPTX" AND '
            + 'DT_NOMTABLE <> "FORMESP"'
            + 'ORDER BY DT_NOMTABLE'
  end;
end;

function TFexpImpCegid.TraitementTable(TheTOZ: TOZ; MaTable,Domaine, Predefini, Prefixe: string; Numversion: integer; DDebut, DFin: TdateTime; CodeEx,LibEx: string; Pagine,premier: boolean ; NumPage,NbEnreg,CurrentLine,NbRecords,NbPage : integer): boolean;
var
  SQl : String;
  QQ : TQuery;
  T_MERE,T_Fic,T_STFIC,T_STMERE : TOB;
  TheFic,TheStFic : string;
  SPName : string;
  MSg : string;
begin
  Result   := True;
  SPName   := Tools.iif(NumPage > 0, GetStoredProcedureName(MaTable), '');
  MoveCurProgressForm(Format('%s/%s - Table %s (%s enregistrement(s))%s', [IntTostr(CurrentLine), IntTostr(GS.nbSelected), MaTable, IntToStr(NbRecords), Tools.iif(NumPage > 0, Format(' - Page %s/%s', [IntToStr(NumPage), IntToStr(NbPage)]), '')]));
  T_STMEre := TOB.Create ('UNE STRUCTURE',nil,-1);
  try
    T_Mere := TOB.Create('Ma tob', nil, -1);
    try
      try
        T_STFIC := TOB.Create(MaTable+'_', T_STMere, -1);
        T_STFIC.AddChampSupValeur('TABLE', MaTable);
        T_STFIC.AddChampSupValeur('VERSION', Numversion);
        T_STFIC.AddChampSupValeur('VIDAGE', 'X');
        T_Fic := TOB.Create(MaTable+'_', T_Mere, -1);
        T_Fic.AddChampSupValeur('TABLE', MaTable);
        T_Fic.AddChampSupValeur('VERSION', Numversion);
        T_Fic.AddChampSupValeur('VIDAGE', 'X');
        if Pos(maTable,'MENU;PARAMSOC;CHOIXCOD;CHOIXEXT;') > 0 then T_Fic.SetString('VIDAGE', '-');
        if (Predefini = 'X') or (Domaine='Y') then T_Fic.SetString('VIDAGE','-');
        TheStFic := ExpDir + '\' + MaTable+'.SEB';
        if MaTable = 'PARAMSOC' then
        begin
          if ExportPaieCompta then
          begin
            SQL := 'SELECT * FROM PARAMSOC WHERE (SOC_TREE LIKE "001;001;%") OR (SOC_TREE LIKE "001;002;%") OR (SOC_TREE LIKE "001;005;%") OR (SOC_TREE LIKE "001;018;%")';
          end else if ExportComptaOnly then
          begin
            SQL := 'SELECT * FROM PARAMSOC WHERE (SOC_TREE LIKE "001;001;%") OR (SOC_TREE LIKE "001;002;%") OR (SOC_TREE LIKE "001;018;%")';
          end else
          begin
            SQL := 'SELECT * FROM PARAMSOC WHERE (SOC_TREE LIKE "001;005;%")';
          end;
          TheStFic := ExpDir + '\' + MaTable+'.SEB';
          TheFic := ExpDir + '\' + MaTable +'.TEB';
        end else if MaTable = 'MENU' then
        begin
          if ExportPaieCompta then
          begin
            SQL := 'SELECT * FROM MENU '+
                   'WHERE ('+
                   '(mn_1=0 and mn_2 in (41,42,43,44,46,47,48,49,200,303,310,347,639,371,372,373,374,375,376,377)) or '+
                   '(mn_1 in (41,42,43,44,46,47,48,49,200,303,310,347,639,371,372,373,374,375,376,377)) or '+
                   '(mn_1=0 and mn_2 in (8,9,11,12,13,14,16,17,18,21,27,324,330,336,340,361)) or '+
                   '(mn_1 in (8,9,11,12,13,14,16,17,18,21,26,27,324,330,336,340,361))'+
                   ')';
          end else if ExportComptaOnly then
          begin
            SQL := 'SELECT * FROM MENU '+
                   'WHERE '+
                   '('+
                   '(mn_1=0 and mn_2 in (8,9,11,12,13,14,16,17,18,21,27,324,330,336,340,361)) or '+
                   '(mn_1 in (8,9,11,12,13,14,16,17,18,21,26,27,324,330,336,340,361))'+
                   ')';

          end else if ExportPaieOnly then
          begin
            SQL := 'SELECT * FROM MENU '+
                   'WHERE '+
                   '('+
                   '(mn_1=0 and mn_2 in (41,42,43,44,46,47,48,49,200,303,310,347,639,371,372,373,374,375,376,377)) or '+
                   '(mn_1 in (41,42,43,44,46,47,48,49,200,303,310,347,639,371,372,373,374,375,376,377))'+
                   ')';
          end;
          TheStFic := ExpDir + '\' + MaTable+'.SEB';
          TheFic := ExpDir + '\' + MaTable +'.TEB';
        end else if SPName <> '' then
        begin
          TheStFic := ExpDir + '\' + MaTable+'.SEB';
          TheFic   := ExpDir + '\' + MaTable +'#'+InttoStr(NumPage)+'.TEB';
          SQL      := Format('exec [dbo].%s %s, %s', [SPName, IntToStr(NumPage), IntToStr(nbEnreg)]);//  +IntToStr(NumPage)+','+IntToStr(nbEnreg);
        end else
        begin
          TheStFic := ExpDir + '\' + MaTable+'.SEB';
          SQL := 'SELECT * FROM '+maTable +' WHERE 1=1';
          TheFic := ExpDir + '\' + MaTable +'.TEB';
          if Predefini='X' then
          begin
            SQL := SQL + ' AND NOT '+Prefixe+'_PREDEFINI IN ("CEG","STD")';
          end;
        end;
        QQ := OPENSQL(SQL, TRUE);
        if not QQ.eof then
        begin
          T_Fic.LoadDetailDb(MaTable, '', '', QQ, FALSE);
        end;
        Ferme(QQ);
        if premier then
        begin
          if FileExists(TheSTFic) then DeleteFile(PChar(TheSTFic));
          T_STFic.SaveToBinFile(TheSTFic, FALSE, TRUE, TRUE, FALSE);
          if not PGZipFile(TheSTFic, TheTOZ, self) then result := FALSE;
        end;
        if FileExists(TheFic) then DeleteFile(PChar(TheFic));
        T_Fic.SaveToBinFile(TheFic, FALSE, TRUE, TRUE, FALSE);
        if not PGZipFile(TheFic, TheTOZ, self) then result := FALSE;
      except
        On E:Exception do
        begin
          Msg := 'Erreur '+E.Message+' sur table '+MaTable;
          PGIInfo(Msg);
          AddTrace(Msg);
          raise;
        end;
      end;
    finally
      FreeAndNil(T_Mere);
    end;
  finally
    FreeAndNil(T_STMEre);
  end;
end;

procedure TFexpImpCegid.RDTIMPORTClick(Sender: TObject);
begin
  CHKMSGSTRUCT.Enabled := true;
  CHKSTOPONERROR.Enabled := true;
  CHKZIP.Enabled := true;
  CHKREPERT.Enabled := true;
  CBVIDAGEIMP.Enabled := true;
  IMPORTFIC.Enabled := true;

  //
  CBVIDAGEEXP.Enabled := false;
  NOMFIC.Enabled := false;
  EXPORTFIC.Enabled := false;
  CBCOMPTA.enabled := False;
  CBPAYE.enabled := False;
end;

procedure TFexpImpCegid.CBCOMPTAClick(Sender: TObject);
begin
  CBDOSSIER.OnClick := nil;
  CBDOSSIER.Checked := false;
  CBDOSSIER.OnClick := CBDOSSIERClick;
end;

procedure TFexpImpCegid.CBPAYEClick(Sender: TObject);
begin
  CBDOSSIER.OnClick := nil;
  CBDOSSIER.Checked := false;
  CBDOSSIER.OnClick := CBDOSSIERClick;
end;

procedure TFexpImpCegid.CBDOSSIERClick(Sender: TObject);
begin
  CBCOMPTA.OnClick := nil;
  CBPAYE.OnClick := nil;
  CBCOMPTA.Checked := false;
  CBPAYE.Checked := false;
  CBCOMPTA.OnClick := CBCOMPTAClick;
  CBPAYE.OnClick := CBPAYEClick;
end;

procedure TFexpImpCegid.BLANCETRAITClick(Sender: TObject);
var
  T1              : TOB;
  T2              : TOB;
  LibEx           : string;
  Prefixe         : string;
  CodeEx          : string;
  MaTable         : string;
  msg             : string;
  PREDEFINI       : string;
  FieldsKey       : string;
  Domaine         : string;
  Monfic          : string;
  RapportName     : string;
  Version         : integer;
  I               : integer;
  ret             : integer;
  NbRecords       : integer;
  NbRecordsTot    : integer;   
  NbRecordsOk     : double;
  NbPage          : integer;
  Cpt             : integer;
  CurrentLine     : integer;
  ExportOk        : Boolean;
  OkOk            : Boolean;
  TheToz          : TOZ;
  DDebut          : Tdatetime;
  DFin            : Tdatetime;
  sr              : TsearchRec;
  FileName : string;
  first : boolean;
  PartNumber : Integer;
begin
  CurrentLine       := 0;
  PGCTRL.ActivePage := RAPPORT;
  ExportOk          := False;
  AddTrace('--------------------------------');
  AddTrace('Heure départ '+DateTimeToStr(NowH));
  AddTrace('--------------------------------');
  try
    if fModeTrait = tmtExport then
    begin
      TheTOZ := nil;
      if not PGCreatZipFile(ZipFile, moCreate, TheTOZ, self) then
      begin
        T.free;
        T_Mere.free;
        AddTrace('Erreur sur fichier archive : ' +ZipFile);
        exit;
      end;
      InitMoveProgressForm(nil, 'Lecture des données', 'Veuillez patienter SVP ...', T.detail.count, FALSE, TRUE);
      TRY
        for I := 1 to GS.RowCount - 1 do
        begin
          if not GS.IsSelected(I) then continue;
          inc(CurrentLine);
          T1 := TOB(GS.objects[0,I]);
          if T1 <> nil then
          begin
            MaTable         := T1.GetString('DT_NOMTABLE');
            Prefixe         := T1.GetString('DT_PREFIXE');
            Version         := T1.GetInteger('DT_NUMVERSION');
            PREDEFINI       := T1.GetString('PREDEFINI');
            DOMAINE         := T1.GetString('DT_DOMAINE');
            FieldsKey       := T1.GetString('DT_CLE1');
            DDebut          := IDate1900;
            DFin            := iDate2099;
            NbRecordsOk     := T1.GetInteger('NBENREG');
            NbRecords       := T1.GetInteger('NBTOTAL');
            NbRecordsTot    := T1.GetInteger('NBENREGTOTAL');
            if NbRecordsOk <>-1 then
            begin
              CreateStoredProcedure(MaTable, FieldsKey);
              try
                NbPage := Floor(Nbrecords/NbrecordsOk);
                if (NbPage * NbrecordsOk) < Nbrecords then Inc(NbPage);
                for Cpt := 1 to NbPage do
                begin
                  try
                    if not TraitementTable(TheTOZ, MaTable, Domaine, PREDEFINI, Prefixe, Version, DDebut, DFin, CodeEx, LibEx, True, (Cpt=1), Cpt, Round(NbrecordsOk), CurrentLine, Round(NbRecordsTot), NbPage) then
                    begin
                      Msg := Format('Erreur dans table %s - Page %s/%s', [MaTable, IntToStr(Cpt), IntToStr(NbPage)]);
                      PgiError(Msg, self.caption);
                      AddTrace(Msg);
                    end else
                      AddTrace(Format('Table %s - Page %s/%s --> Traité Ok', [MaTable, IntToStr(Cpt), IntToStr(NbPage)]));
                  except
                    On E:Exception do
                    begin
                      Msg := Format('Erreur dans table %s - Page %s/%s', [MaTable, IntToStr(Cpt), IntToStr(NbPage)]);
                      PGIInfo('Erreur '+E.Message+' sur table '+MaTable );
                      raise;
                    end;
                  end;
                end;
              finally
                Tools.DropStoredProcedure(GetStoredProcedureName(MaTable));
              end;
            end else
            begin
              TRY
                if not TraitementTable(TheTOZ, MaTable, Domaine, PREDEFINI, Prefixe, version, DDebut, DFin, '', LibEx, False, True, 0, 0, CurrentLine, Round(NbRecordsTot), 0) then
                begin
                  Msg := 'Erreur dans table '+MaTable;
                  PgiError(Msg, self.caption);
                  AddTrace(Msg);
                end else
                  AddTrace(Format('Table %s --> Traité Ok', [MaTable]));
              EXCEPT
                On E:Exception do
                begin
                  Msg :='Erreur '+E.Message+' sur table '+MaTable;
                  PGIInfo(Msg);
                  AddTrace(Msg);
                  raise;
                end;
              END;
            end;
          end;
        end;
        ExportOk := True;
      FINALLY
        MoveCurProgressForm('Fin de traitement des données');
        if not Exportok then
        begin
          PgiError('Le traitement est abandonné', self.caption)
        end else
        begin
          PGFermeZipFile(TheTOZ);
          if CBVIDAGEEXP.checked then PGVideDirectory(ExpDir);
        end;
        if Assigned(TheToz) then TheToz.Free;
        try
          AddTrace('Ecriture du fichier OK');
        except
          PGIBox('Une erreur est survenue lors de l''écriture du fichier', self.caption);
          AddTrace('Une erreur est survenue lors de l''écriture du fichier');
        end;
        FiniMoveProgressForm();
      end;
    end else
    begin
      ret := FindFirst(fRepert + '*.TEB', 0 + faAnyFile, sr);
      InitMoveProgressForm(nil, 'Lecture des données', 'Veuillez patienter SVP ...', T.detail.count, FALSE, TRUE);

      OkOk := TRUE;
      TRY
        while (ret = 0) and (OkOk) do
        begin
          T_mere.ClearDetail;
          //
          First:= true;
          FileName := Copy(ExtractFileName(sr.name),1,Pos('.',sr.Name)-1);
          if Pos('#',FileName) > 0 then
          begin
            PartNumber  := valeurI(Copy(FileName,Pos('#',FileName)+1,255));
            if  PartNumber > 1 then first := false;
          end;
          //
          MonFic := fRepert + sr.Name;
          TOBLoadFromBinFile(MonFic, nil, T_Mere);
          T1 := T_Mere.Detail[0];
          MoveCurProgressForm('Traitement des données du fichier ' + sr.Name);
          try
            T2 := T.FindFirst(['DT_NOMTABLE'], [ T1.GetValue('TABLE')], FALSE);
            if T2 = nil then
            begin
              ret := FindNext(sr);
              continue; // la table n''existe pas dans la destination....
            end;
            if (T1.getString('VIDAGE')='X') then
            begin
              if (T2.GetString('REINIT')<> 'X') and (first) then
              begin
                ExecuteSql('TRUNCATE TABLE '+T1.GetValue('TABLE'));
                T2.SetString('REINIT','X');
              end;
              T_Mere.SetAllModifie(true);
              OkOk := T_Mere.InsertDBByNivel(TRUE);
            end else
            begin
              T_Mere.SetAllModifie(true);
              OkOk := T_Mere.InsertOrUpdateDB(TRUE);
            end;
            AddTrace('Traitement du fichier ' + sr.Name + ' OK');
            ret := FindNext(sr);
          except
            On E:Exception do
            begin
              Msg :='Erreur '+E.Message+' sur table '+ T1.GetValue('TABLE');
              AddTrace(Msg);
              if (CHKSTOPONERROR.checked) then
              begin
                PgiError('ATTENTION : Le traitement des données du fichier '+sr.name+' n''est pas correct.#13#10 votre base de données n''est pas utilisable en l''état !', self.Caption);
                OkOk := FALSE;
                break;
              end;
            end;
          end;
        end;
      FINALLY
        MoveCurProgressForm('Nettoyage final');
        T_Mere.ClearDetail;
        FiniMoveProgressForm();
        sysutils.FindClose(sr);
      end;
      if not okok then
      begin
        PGIInfo('Importation abandonnée');
        BLANCETRAIT.Visible := false;
        BValider.Visible := true;
      end else
      begin
        fRepert := Copy(fRepert, 1, Strlen(PChar(fRepert)) - 1);
        if CBVIDAGEIMP.Checked then PGVideDirectory(fRepert);
        EXECUTESQL ('DELETE FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="PSA_PRENOM"');
      end;
    end;
  finally
    AddTrace('--------------------------------');
    AddTrace('Heure fin '+DateTimeToStr(NowH));
    AddTrace('--------------------------------');
    RapportName := Format('%s\%s_Rapport%s.txt', [fRepert, FormatDateTime('yyyymmdd', Now), Tools.iif(fModeTrait = tmtExport, 'Export', 'Import')]);
    Trace.Items.SaveToFile(RapportName);
    PgiBox(Format('%s terminé (rapport enregistré dans %s)', [Tools.iif(fModeTrait = tmtExport, 'Export', 'Import'), RapportName]), self.caption);
    BLANCETRAIT.Visible := false;
    BValider.Visible := true;
  end;
end;

function TFexpImpCegid.GetStoredProcedureName(TableName : string) : string;
begin
  Result := Tools.iif(TableName <> '', Format('sp_paging_%s', [TableName]), '');
end;

procedure TFexpImpCegid.CreateStoredProcedure(TableName, SortedFields : string);
var
  Sql       : string;
  SPName    : string;
begin
  if TableName <> '' then
  begin
    SPName := GetStoredProcedureName(TableName);
    if SPName <> '' then
    begin
      Tools.DropStoredProcedure(SPName);
      Sql := Format('CREATE PROCEDURE [dbo].[%s] @pageno as int, @records as int'
                 +  ' AS'
                 +  ' BEGIN'
                 +  ' SET NOCOUNT ON;'
                 +  ' declare @offsetcount as int'
                 +  ' set @offsetcount=(@pageno-1)*@records'
                 +  ' select * from %s order by %s offset @offsetcount rows fetch Next @records rows only'
                 +  ' END'
                   , [SPName, TableName, SortedFields]);
      ExecuteSQL(Sql);
    end;
  end;
end;

function TFexpImpCegid.GetProcessMemorySize(_sProcessName: string; var _nMemSize: Cardinal): Boolean;
var
  l_nWndHandle, l_nProcID, l_nTmpHandle: HWND;
  l_pPMC: PPROCESS_MEMORY_COUNTERS;
  l_pPMCSize: Cardinal;
begin
  l_nWndHandle := FindWindow(nil, PChar(_sProcessName)); // Application. handle;

  if l_nWndHandle = 0 then
  begin
    Result := False;
    Exit;
  end;

  l_pPMCSize := SizeOf(PROCESS_MEMORY_COUNTERS);

  GetMem(l_pPMC, l_pPMCSize);
  l_pPMC^.cb := l_pPMCSize;

  GetWindowThreadProcessId(l_nWndHandle, @l_nProcID);
  l_nTmpHandle := OpenProcess(PROCESS_ALL_ACCESS, False, l_nProcID);

  if (GetProcessMemoryInfo(l_nTmpHandle, l_pPMC, l_pPMCSize)) then
    _nMemSize := l_pPMC^.WorkingSetSize
  else
    _nMemSize := 0;

  FreeMem(l_pPMC);

  Result := True;
end;

procedure TFexpImpCegid.AddTrace(Text : string);
begin
  Trace.Items.Insert(0, Text);
end;

procedure TFexpImpCegid.GSDblClick(Sender: TObject);
var Arect : TRect;
    T1 : TOB;
begin
  //
  T1 := T.Detail[GS.row-1];
  ARect := GS.CellRect(1, GS.Row);
  TX.Parent := GS;
  TX.Top := Arect.top;
  TX.Left := Arect.Left;
  TX.Text := IntToStr(T1.GetInteger('NBENREG'));
  TX.visible := True;
  TX.SetFocus;
end;

procedure TFexpImpCegid.TXExit(Sender: TObject);
var T1 : TOB;
begin
  if not IsNumeric(TX.Text) then BEGIN TX.SetFocus; Exit; end;
  T1 := T.Detail[GS.row-1];
  T1.setInteger('NBENREG',StrToInt(TX.text));
  TX.Visible := false;
  T1.PutLigneGrid(GS,GS.row,false,false,'DT_NOMTABLE;NBENREG;');
end;

procedure TFexpImpCegid.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var Arect : TRect;
    T1 : TOB;
begin
//
  if csDestroying in ComponentState then Exit;
  if GS.RowHeights[ARow] <= 0 then Exit;
  if ARow = 0 then Exit;
  T1 := T.detail[Arow-1];
  ARect := GS.CellRect(ACol, ARow);
  GS.Canvas.Pen.Style := psSolid;
  GS.Canvas.Pen.Color := clgray;
  GS.Canvas.Brush.Style := BsSolid;
  if (Acol = 1)  Then
  begin
    if T1.GetInteger('NBENREG')=-1 then
    begin
    	Canvas.FillRect(ARect);
      GS.Canvas.Brush.Style := bsSolid;
      GS.Canvas.TextOut (Arect.left+1,Arect.Top +2,'Tout');
    end;
  end;

end;

end.

