unit RecupTarifFour;

interface

uses
  Windows
  , Messages
  , SysUtils
  , Classes
  , Graphics
  , Controls
  , Forms
  , Dialogs
  , assist
  , HSysMenu
  , hmsgbox
  , StdCtrls
  , HTB97
  , ComCtrls
  , ExtCtrls
  , Hctrls
  , Hent1
  , UIUtil
  , HPanel
  , Mask
  , UTOB
  , Math
  , ParamSoc
  {$IFDEF YOYO}
  , Rapport
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  {$ELSE EAGLCLIENT}
    {$IFNDEF DBXPRESS}
  , dbTables
    {$ELSE}
   , uDbxDataSet
     {$ENDIF}
  {$ENDIF EAGLCLIENT}
  , HStatus
  , UtilArticle
  , Ed_Tools
  , TarifAutoFour
  , TarifUtil
  , UtilGC
  , UtilTarif
  , TntStdCtrls
  , TntComCtrls
  , TntExtCtrls
  , ADODB
  ;

procedure EntreeRecupTarifFour (StParFou : string);

type
  TFRecupTarifFour = class(TFAssist)
    TINTRO: THLabel;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    PTITRE: THPanel;
    HLabel1: THLabel;
    GBCreationFiches: TGroupBox;
    CBCreationCatalogue: TCheckBox;
    CBCreationArticle: TCheckBox;
    TNomChamp: THLabel;
    TLongueur: THLabel;
    TOffset: THLabel;
    TCritere: THLabel;
    TBorneSup: THLabel;
    TBorneInf: THLabel;
    HLabel2: THLabel;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    ListRecap: TListBox;
    TRecap: THLabel;
    GBPremier: TGroupBox;
    HChamp1: THValComboBox;
    HLongueur1: THCritMaskEdit;
    HOffset1: THCritMaskEdit;
    HBorneInf1: THCritMaskEdit;
    HBorneSup1: THCritMaskEdit;
    GBDeuxieme: TGroupBox;
    HChamp2: THValComboBox;
    HLongueur2: THCritMaskEdit;
    HOffset2: THCritMaskEdit;
    HBorneInf2: THCritMaskEdit;
    HBorneSup2: THCritMaskEdit;
    GBTroisieme: TGroupBox;
    HChamp3: THValComboBox;
    HLongueur3: THCritMaskEdit;
    HOffset3: THCritMaskEdit;
    HBorneInf3: THCritMaskEdit;
    HBorneSup3: THCritMaskEdit;
    TRGF_PARFOU: THLabel;
    TGRF_TIERS: THLabel;
    GRF_PARFOU: THCritMaskEdit;
    GRF_TIERS: THCritMaskEdit;
    LGRF_TIERS: THLabel;
    HRecap: THMsgBox;
    OpenDialogButton: TOpenDialog;
    TGRF_FICHIER: THLabel;
    GRF_FICHIER: THCritMaskEdit;
    TDATESUP: THLabel;
    DATESUP: THCritMaskEdit;
    GRF_MULTIFOU: THCheckbox;
    CBRecInit: TCheckBox;
    chkFamSsFam: TCheckBox;
    
    procedure bFinClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure GRF_FICHIERElipsisClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure HChamp1Change(Sender: TObject);
    procedure HChamp2Change(Sender: TObject);
    procedure HChamp3Change(Sender: TObject);
    procedure DATESUPDblClick(Sender: TObject);
    procedure CBCreationArticleClick(Sender: TObject);
    procedure CBCreationCatalogueClick(Sender: TObject);
    
  private
    TheFournisseur        : string;
    TobT                  : TOB;
    TOBErreurs            : TOB;
    TobTablettes          : TOB;
    TobGPF                : TOB;
    TobGCA                : TOB;
    TobGA                 : TOB;
    TobParFou             : TOB;
    TobParFouLig          : TOB;
    TslGCAFields          : TStringList;
    TslGAFields           : TStringList;
    TslGA2Fields          : TStringList;
    TslBCBFields          : TStringList;
    TslBFTFields          : TStringList;
    TslBSFFields          : TStringList;
    TslGNEFields          : TStringList;
    TslGNLFields          : TStringList;
    TslCurrentKey         : TstringList;
    TslPARFOULIG          : TStringList;
    TslCacheIndexes       : TStringList;
    TslFieldsTablet       : TstringList;
    TslInsertTablet       : TstringList;
    TslUpdateTablet       : TstringList;
    TslCacheTobGCA        : TstringList;
    TslCacheTobGA         : TstringList;
    TslGAValuesToUpgraded : TstringList;
    TslInsertGCA          : TstringList;
    TslUpdateGCA          : TstringList;
    TslInsertGA           : TstringList;
    TslUpdateGA           : TstringList;
    TslInsertGA2          : TstringList;
    TslUpdateGA2          : TstringList;
    TslInsertBFT          : TstringList;
    TslInsertBSF          : TstringList;
    TslInsertGNE          : TstringList;
    TslInsertGNL          : TstringList;
    TslInsertBCB          : TstringList;
    GAParametersPrices    : array [0..3]of string;
    TabletExist           : boolean;
    CreateFamily          : boolean;
    CreateBCB             : boolean;
    ExistBCBPrincipal     : boolean;
    NeedValorization      : boolean;
    CreateGA2             : boolean;
    InsertGCALines        : integer;
    InsertGALines         : integer;
    InsertBFTLines        : integer;
    InsertBSFLines        : integer;
    InsertGNELines        : integer;
    InsertGNLLines        : integer;
    InsertBCBLines        : integer;
    UpdateGCALines        : integer;
    UpdateGALines         : integer;

    procedure ActiveChamp (HCrit : THCritMaskEdit);
    procedure DepileTOBLigne ;
    procedure DesactiveChamp (HCrit : THCritMaskEdit);
    Procedure InitialiseFiche;
    procedure RempliCombo;
    function ControleSaisie : boolean;
    function ControleSelection (stChamp, stBorneSup, stBorneInf, stLongueur, stOffset : string; Message1, Message2 : integer) : boolean;
    procedure ListeRecap;
    procedure RecapChamp(stChamp, stLongueur, stOffset, stBorneInf, stBorneSup : string);
    procedure Valorization(TobGA, TobGCA : TOB);
    procedure AffichageErreurs;
    function GetSeparator : string;
    function ReplaceChar(Value, OldPatern, NewPatern : string) : string;
    function GetTslFields(pPrefix : string) : TStringList;
    function GetFieldsQty(Prefix : string) : integer;
    function GetSpPrepare(Prefix : string; IsInsert : boolean) : string;
    procedure SetSpExec(TslData : TStringList; IsInsert : boolean; pEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, Prefix : string);
    procedure InsertUpdateInDB(pCnx : TADOConnection; Prefix : string; IsInsert : boolean; TslData : TStringList);
    function SupplierIsValid(WithMsg : boolean=True) : boolean;
    function GPFParamIsValid : boolean;
    function SetNeedValorization : boolean;
    procedure LoadCache;
    function GetExistSql(ReferenceGCA, ReferenceGA : string) : string;
    procedure SetTslFieldsList(ReferenceGCA, ReferenceGA, Prefix : string; TslData : TStringList);
    function GetDefaultValue(ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, Prefix, FieldName, FieldType : string) : string;
    function ExtractValue(FieldName, FieldType, pEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode : string) : string;
    procedure InsertTablet(pCnx : TADOConnection);
    procedure ExecuteInsertIfNeed(Prefix : string; LinesQty : integer; pCnx : TADOConnection);
    procedure ExecuteUpdateIfNeed(Prefix : string; LinesQty : integer; pCnx : TADOConnection);
    function RecupereCatalogue : boolean;
  public
  end;


implementation

uses
  TiersUtil
  , RECALCPIECE_RAP_TOF                                                                                       
  , CbpMCD
  , CbpEnumerator
  , CommonTools
  , UApplication
  , HeureUtil
  , ADOInt
  , TablesDefaultValues
  , DateUtils
  , FormsName
  ;

const
  MaxQryExec       = 5000;
  PrefixGCA        = 'GCA';
  PrefixGA         = 'GA';
  PrefixGA2        = 'GA2';
  PrefixBCB        = 'BCB';
  PrefixBFT        = 'BFT';
  PrefixBSF        = 'BSF';
  PrefixGNE        = 'GNE';
  PrefixGNL        = 'GNL';
  FieldsToUpgraded = ';GA_DPR;GA_PAHT;GA_PVHT;GA_PVTTC;';

{$R *.DFM}

procedure EntreeRecupTarifFour (StParFou : string);
var FF : TFRecupTarifFour;
    PPANEL  : THPanel ;
BEGIN
  SourisSablier;
  FF := TFRecupTarifFour.Create(Application) ;
  FF.GRF_PARFOU.Text := StParFou;

  PPANEL := FindInsidePanel ; // permet de savoir si la forme dépend d'un PANEL
  if PPANEL = Nil then        // Le PANEL est le premier ecran affiché
  BEGIN
    try
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
    SourisNormale ;
  end else
  BEGIN
    InitInside (FF, PPANEL);
    FF.Show ;
  end ;
end ;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}

procedure TFRecupTarifFour.ActiveChamp (HCrit : THCritMaskEdit);
begin
  HCrit.Enabled := True;
  HCrit.Color := clWindow;
end;

Procedure TFRecupTarifFour.DepileTOBLigne ;
var Index : integer;
BEGIN
  for Index := TobParFouLig.Detail.Count - 1 Downto 0 do
  BEGIN
  	TobParFouLig.Detail[Index].Free ;
  end;
end;

procedure TFRecupTarifFour.DesactiveChamp (HCrit : THCritMaskEdit);
begin
  HCrit.Enabled := False;
  HCrit.Color := clActiveBorder;
  HCrit.Text := '';
end;

Procedure TFRecupTarifFour.InitialiseFiche;
BEGIN
  TobParFou.InitValeurs;
  TobParFouLig.InitValeurs;

  LoadLesTobParFou (GRF_PARFOU.Text, TobParFou, TobParFouLig, True);
  RempliCombo;

  HChamp1.Text := 'Aucun';
  HChamp2.Enabled := False;
  HChamp3.Enabled := False;
  HChamp2.Text := 'Aucun';
  HChamp3.Text := 'Aucun';

  DesactiveChamp (HLongueur1);
  DesactiveChamp (HLongueur2);
  DesactiveChamp (HLongueur3);

  DesactiveChamp (HOffset1);
  DesactiveChamp (HOffset2);
  DesactiveChamp (HOffset3);

  DesactiveChamp (HBorneInf1);
  DesactiveChamp (HBorneInf2);
  DesactiveChamp (HBorneInf3);

  DesactiveChamp (HBorneSup1);
  DesactiveChamp (HBorneSup2);
  DesactiveChamp (HBorneSup3);

  CBCreationArticle.Checked := False;
  CBCreationCatalogue.Checked := False;

  GRF_TIERS.Text := TobParFou.GetValue ('GRF_TIERS');
  LGRF_TIERS.Caption := GetValChamp ('TIERS','T_LIBELLE','T_TIERS="' + GRF_TIERS.Text + '"');
  GRF_FICHIER.Text := TobParFou.GetValue ('GRF_FICHIER');
  GRF_MULTIFOU.Checked := TobParFou.GetBoolean ('GRF_MULTIFOU');
  if not GRF_MULTIFOU.Checked then GRF_MULTIFOU.Visible := false;
  chkFamSsFam.checked := GetParamSocSecur('SO_MODFAMSSFAM',true);
end;

procedure TFRecupTarifFour.RempliCombo;
var Index, IndexItem : integer;
    LibelleChamp : string;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  if TobParFou.getValue('GRF_TYPEENREG')<>'X' then
  begin
	  TobParFouLig.Detail.Sort('GFL_OFFSET');
  end;
  for Index := 0 to TobParFouLig.Detail.Count - 1 do
  begin
    if TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP') <> 'PASS' then
    begin
      if pos ('GCA_', TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP')) > 0 then
      begin
        Table := Mcd.getTable ('CATALOGU');
      end else
      begin
        Table := Mcd.getTable ('ARTICLE');
      end;
      FieldList := Table.Fields;
      FieldList.Reset();
      While FieldList.MoveNext do
      begin
        if (FieldList.Current as IFieldCOM).name = TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP') then
        LibelleChamp :=(FieldList.Current as IFieldCOM).Libelle;
      end;
      IndexItem := HChamp1.Items.Add (LibelleChamp);
      HChamp1.Items.objects[IndexItem] := TobParFouLig.Detail [Index];
      IndexItem := HChamp2.Items.Add (LibelleChamp);
      HChamp2.Items.Objects[IndexItem] := TobParFouLig.Detail [Index];
      IndexItem := HChamp3.Items.Add (LibelleChamp);
      HChamp3.Items.objects[IndexItem] := TobParFouLig.Detail [Index];
    end;
  end;
  HChamp1.Items.add ('Autre');
  HChamp2.Items.add ('Autre');
  HChamp3.Items.add ('Autre');

  HChamp1.Items.add ('Aucun');
  HChamp2.Items.add ('Aucun');
  HChamp3.Items.add ('Aucun');
end;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}

procedure TFRecupTarifFour.bFinClick(Sender: TObject);
begin
  inherited;
  ListeRecap;
  if ControleSaisie then
  begin
    if fileexists (GRF_FICHIER.Text) then
    begin
      TOBErreurs        := TOB.Create('LES ANOMALIES',nil,-1);
      try
  			if not GRF_MULTIFOU.Checked then
    			TheFournisseur := GRF_TIERS.Text;
        if RecupereCatalogue then
        begin
          if TOBErreurs.detail.count > 0 then AffichageErreurs;
          if JaiLeDroitTag(60301) then
          begin
            if PGIAsk('Traitement de récupération terminé. Voulez-vous consulter le journal d''évènements ?') = mrYes then
              OpenForm.JnalEvent('TYPEEVENT=MTF');
          end else
            Msg.Execute(9, Caption, '');
        end;
        Close;
      finally
        TOBErreurs.Free;
      end;
    end else
      Msg.Execute(7, Caption, '');
  end else
    ModalResult := 0;
end;

procedure TFRecupTarifFour.bSuivantClick(Sender: TObject);
var i_NumEcran : integer;
begin
  i_NumEcran := strtoint(Copy(P.ActivePage.Name, length(P.ActivePage.Name), 1));
  case i_NumEcran of
  	2 : if not ControleSaisie then Exit;
  end;
  inherited;
  if bFin.Enabled then ListeRecap;
end;

procedure TFRecupTarifFour.DATESUPDblClick(Sender: TObject);
begin
  inherited;
	GetDateRecherche (TForm(Self), DATESUP) ;
end;

procedure TFRecupTarifFour.FormClose(Sender: TObject;
var
  Action: TCloseAction);
begin
  inherited;
  DepileTOBLigne;
  FreeAndNil(TobParFouLig);
  FreeAndNil(TobParFou);
  FreeAndNil(TobT);
  FreeAndNil(TobTablettes);
  FreeAndNil(TobGPF);
  FreeAndNil(TobGCA);
  FreeAndNil(TobGA);
  FreeAndNil(TslGCAFields);
  FreeAndNil(TslGAFields);
  FreeAndNil(TslGA2Fields);
  FreeAndNil(TslBCBFields);
  FreeAndNil(TslBFTFields);
  FreeAndNil(TslBSFFields);
  FreeAndNil(TslGNEFields);
  FreeAndNil(TslGNLFields);
  FreeAndNil(TslCurrentKey);
  FreeAndNil(TslPARFOULIG);
  FreeAndNil(TslCacheIndexes);
  FreeAndNil(TslFieldsTablet);
  FreeAndNil(TslInsertTablet);
  FreeAndNil(TslUpdateTablet);
  FreeAndNil(TslCacheTobGCA);
  FreeAndNil(TslCacheTobGA);
  FreeAndNil(TslGAValuesToUpgraded);
  FreeAndNil(TslInsertGCA);
  FreeAndNil(TslUpdateGCA);
  FreeAndNil(TslInsertGA);
  FreeAndNil(TslUpdateGA);
  FreeAndNil(TslInsertGA2);
  FreeAndNil(TslUpdateGA2);          
  FreeAndNil(TslInsertBFT);
  FreeAndNil(TslInsertBSF);
  FreeAndNil(TslInsertGNE);
  FreeAndNil(TslInsertGNL);
  FreeAndNil(TslInsertBCB);
end;

procedure TFRecupTarifFour.GRF_FICHIERElipsisClick(Sender: TObject);
begin
  inherited;
  if OpenDialogButton.Execute then
  begin
    if OpenDialogButton.FileName <> '' then
    begin
    	GRF_FICHIER.Text := OpenDialogButton.Filename;
    end;
  end;
end;

procedure TFRecupTarifFour.FormCreate(Sender: TObject);
begin
  inherited;
  TheFournisseur        := '';
  TobParFou             := TOB.Create('PARFOU', Nil, -1);
  TobParFouLig          := TOB.Create('', Nil, -1);
  TobT                  := TOB.Create('TIERS',nil,-1);
  TobTablettes          := TOB.Create('_LESTABLETTES', nil, -1);
  TobGPF                := TOB.Create('_LESPROFILS', nil, -1);
  TobGCA                := TOB.Create('_CATALOGU', nil, -1);
  TobGA                 := TOB.Create('_ARTICLE', nil, -1);
  TslGCAFields          := TStringList.create;
  TslGAFields           := TStringList.create;
  TslGA2Fields          := TStringList.create;
  TslBCBFields          := TStringList.Create;
  TslBFTFields          := TStringList.create;
  TslBSFFields          := TStringList.create;
  TslGNEFields          := TStringList.create;
  TslGNLFields          := TStringList.create;
  TslCurrentKey         := TStringList.create;
  TslPARFOULIG          := TStringList.create;
  TslCacheIndexes       := TStringList.create;
  TslFieldsTablet       := TStringList.create;
  TslInsertTablet       := TStringList.create;
  TslUpdateTablet       := TStringList.create;
  TslCacheTobGCA        := TStringList.create;
  TslCacheTobGA         := TStringList.create;
  TslGAValuesToUpgraded := TStringList.create;
  TslInsertGCA          := TStringList.create;
  TslUpdateGCA          := TStringList.create;
  TslInsertGA           := TStringList.create;
  TslUpdateGA           := TStringList.create;
  TslInsertGA2          := TStringList.create;
  TslUpdateGA2          := TStringList.create;
  TslInsertBFT          := TStringList.create;
  TslInsertBSF          := TStringList.create;
  TslInsertGNE          := TStringList.create;
  TslInsertGNL          := TStringList.create;
  TslInsertBCB          := TStringList.create;
end;

procedure TFRecupTarifFour.FormShow(Sender: TObject);
begin
  inherited;
	InitialiseFiche;
end;

procedure TFRecupTarifFour.HChamp1Change(Sender: TObject);
begin
  inherited;
  if HChamp1.Text = 'Aucun' then
  begin
    DesactiveChamp (HBorneInf1);
    DesactiveChamp (HBorneSup1);
    DesactiveChamp (HLongueur1);
    DesactiveChamp (HOffset1);
    HChamp2.Enabled := False;
    HChamp2.Text := 'Aucun';
    HChamp2Change (HChamp1);
  end else
  begin
    HChamp2.Enabled := True;
    if HChamp1.Text = 'Autre' then
    begin
      if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
      begin
      	ActiveChamp (HLongueur1);
      end;
      ActiveChamp (HOffset1);
      ActiveChamp (HBorneInf1);
      ActiveChamp (HBorneSup1);
    end else
    begin
      ActiveChamp (HBorneInf1);
      ActiveChamp (HBorneSup1);
      DesactiveChamp (HLongueur1);
      DesactiveChamp (HOffset1);
    end;
  end;
end;

procedure TFRecupTarifFour.HChamp2Change(Sender: TObject);
begin
  inherited;
  if HChamp2.Text = 'Aucun' then
  begin
    DesactiveChamp (HBorneInf2);
    DesactiveChamp (HBorneSup2);
    DesactiveChamp (HLongueur2);
    DesactiveChamp (HOffset2);
    HChamp3.Enabled := False;
    HChamp3.Text := 'Aucun';
    HChamp3Change (HChamp2);
  end else
  begin
    HChamp3.Enabled := True;
    if HChamp2.Text = 'Autre' then
    begin
      if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
      begin
      	ActiveChamp (HLongueur2);
      end;
      ActiveChamp (HOffset2);
      ActiveChamp (HBorneInf2);
      ActiveChamp (HBorneSup2);
    end else
    begin
      ActiveChamp (HBorneInf2);
      ActiveChamp (HBorneSup2);
      DesactiveChamp (HLongueur2);
      DesactiveChamp (HOffset2);
    end;
  end;
end;

procedure TFRecupTarifFour.HChamp3Change(Sender: TObject);
begin
  inherited;
  if HChamp3.Text = 'Aucun' then
  begin
    DesactiveChamp (HBorneInf3);
    DesactiveChamp (HBorneSup3);
    DesactiveChamp (HLongueur3);
    DesactiveChamp (HOffset3);
  end else
  begin
    if HChamp3.Text = 'Autre' then
    begin
      if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
      begin
        ActiveChamp (HLongueur3);
      end;
      ActiveChamp (HOffset3);
      ActiveChamp (HBorneInf3);
      ActiveChamp (HBorneSup3);
    end else
    begin
      ActiveChamp (HBorneInf3);
      ActiveChamp (HBorneSup3);
      DesactiveChamp (HLongueur3);
      DesactiveChamp (HOffset3);
    end;
  end;
end;

function TFRecupTarifFour.ControleSaisie : boolean;
begin
  Result := True;
  if not ControleSelection (HChamp1.Text, HBorneSup1.Text, HBorneInf1.Text, HLongueur1.Text,
  													HOffset1.Text, 0, 7) then
  begin
  	Result := False
  end else if not ControleSelection (HChamp2.Text, HBorneSup2.Text, HBorneInf2.Text, HLongueur2.Text,
  																	 HOffset2.Text, 2, 3) then
  begin
  	Result := False
  end else if not ControleSelection (HChamp3.Text, HBorneSup3.Text, HBorneInf3.Text,
  																	HLongueur3.Text, HOffset3.Text, 4, 5) then
  begin
  	Result := False;
  end;
end;

function TFRecupTarifFour.ControleSelection (stChamp, stBorneSup, stBorneInf, stLongueur,
                                             stOffset : string; Message1, Message2 : integer) : boolean;
begin
  Result := True;
  if stChamp <> 'Aucun' then
  begin
    if stBorneSup < stBorneInf then
    begin
      Msg.Execute (Message1 + 1, Caption, '');
      Result := False;
    end else
    begin
      if stChamp = 'Autre' then
      begin
        if (((stLongueur = '') or (Not IsNumeric (stLongueur))) and
            (TobParFou.GetValue ('GRF_TYPEENREG') = 'X')) or
            ((stOffset = '') or (Not IsNumeric (stOffset))) then
        begin
          Msg.Execute (Message2 + 1, Caption, '');
          Result := False;
        end;
      end;
    end;
  end;
end;

procedure TFRecupTarifFour.ListeRecap;
BEGIN
  ListRecap.Items.Clear;
  ListRecap.Items.Add(PTITRE.Caption);
  ListRecap.Items.Add('Fichier récupéré : ' + GRF_FICHIER.Text);
  ListRecap.Items.Add(TDateSup.Caption + ' : ' + DATESUP.Text);
	ListRecap.Items.Add(ExtractLibelle (CBCreationCatalogue.Caption) + Tools.iif(CBCreationCatalogue.checked, HRecap.Mess[0], HRecap.Mess[1]));
  ListRecap.Items.Add(ExtractLibelle (CBCreationArticle.Caption) + Tools.iif(CBCreationArticle.checked, HRecap.Mess[0], HRecap.Mess[1]));
  if HChamp1.Text <> 'Aucun' then
  begin
    ListRecap.Items.Add('Sélection des enregistrements');
    RecapChamp(HChamp1.Text, HLongueur1.Text, HOffset1.Text, HBorneInf1.Text,
    HBorneSup1.Text);
    if HChamp2.Text <> 'Aucun' then
    begin
      RecapChamp (HChamp2.Text, HLongueur2.Text, HOffset2.Text, HBorneInf2.Text,
      HBorneSup2.Text);
      if HChamp3.Text <> 'Aucun' then
      begin
        RecapChamp (HChamp3.Text, HLongueur3.Text, HOffset3.Text, HBorneInf3.Text,
        HBorneSup3.Text);
      end;
    end;
  end else
  begin
  	ListRecap.Items.Add('Aucun élément de sélection des enregistrements');
  end;
  ListRecap.Items.Add('');
end;

procedure TFRecupTarifFour.RecapChamp (stChamp, stLongueur, stOffset, stBorneInf,
                                       stBorneSup : string);
var stChaine : string;
begin
  if stChamp <> 'Autre' then
  begin
  	ListRecap.Items.Add ('Sélection sur le champ : ' + stChamp);
  end else
  begin
    if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
    begin
      ListRecap.Items.Add ('Sélection à la position : ' + stOffset);
      stChaine := '    sur une longeur de : ' + stLongueur + ' caractère';
      if StrToInt(stLongueur) > 1 then stChaine := stChaine + 's';
      ListRecap.Items.Add (stChaine);
    end else
    begin
      ListRecap.Items.Add ('Sélection sur le champ n° : ' + stOffset);
    end;
  end;
  ListRecap.Items.Add ('    Borne inférieure : ' + stBorneInf);
  ListRecap.Items.Add ('    Borne supérieure : ' + stBorneSup);
end;

procedure TFRecupTarifFour.Valorization(TobGA, TobGCA : TOB);
var
  TobGF        : TOB;
  TobTl        : TOB;
  MTPAF        : double;
  TypeArticle  : string;
  Fournisseur  : string;
  Sql          : string;
  Qry          : TQuery;
  PrixPourQte  : double;
  IsUniteAchat : boolean;
begin
  if assigned(TobGA) then
  begin
    IsUniteAchat := True;
    TypeArticle  := TobGA.GetString('GA_TYPEARTICLE');
    Fournisseur  := TobGA.GetString('GA_FOURNPRINC');
    if (Fournisseur = '') and (TobParFou.GetBoolean('GRF_FOURPRINC')) then
      Fournisseur := TobParFou.GetString('GRF_TIERS');
    if      (Fournisseur <> '')
        and (   (TypeArticle = 'MAR')
             or (TypeArticle = 'ARP')
             or (TypeArticle = 'PRE'))
    then
    begin
      TobGF  := TOB.Create('TARIF',nil,-1);
      try
        TobTl := TobT.FindFirst(['T_TIERS'], [Fournisseur], True);
        if not assigned(TobTl) then
        begin
          TobTl := TOB.Create('TIERS', TobT, -1);
          TobTl.InitValeurs;
          Sql := Format('SELECT * FROM TIERS WHERE T_TIERS = "%s"', [Fournisseur]);
          Qry := OpenSql(Sql, True);
          try
            TobTl.SelectDB('', Qry, True);
          finally
            Ferme(Qry);
          end;
        end;
        if assigned(TobGCA) then
        begin
          GetTarifGlobal(TobGA.GetString('GA_ARTICLE'), TobGA.GetString('GA_TARIFARTICLE'), TobGA.GetString('GA2_SOUSFAMTARART'), 'ACH', TobGA, TobTl, TobGF, true);
          if TobGF.GetDouble('GF_PRIXUNITAIRE') <> 0 then
          begin
            MTPAF := TobGF.GetDouble('GF_PRIXUNITAIRE');
          end else
          begin
            PrixPourQte := TobGCA.GetDouble('GCA_PRIXPOURQTEAC');
            if PrixPourQte = 0 then PrixPourQte := 1;
            MTPAF := Tools.iif(TobGCA.GetDouble('GCA_PRIXBASE') <> 0, (TobGCA.GetDouble('GCA_PRIXBASE')/PrixPourQte), (TobGCA.GetDouble('GCA_PRIXVENTE')/PrixPourQte));
          end;
          if MTPAF = 0 then
          begin
            MTPAF := TobGA.GetDouble('GA_PAHT');
            IsUniteAchat := false;
          end;
          MTPAF := arrondi(MTPAF * (1-(TobGF.GetDouble('GF_REMISE')/100)), V_PGI.OkDecP );
          if IsUniteAchat then
            MTPAF := PassageUAUV(TobGA, TobGCA, MTPAF);
          if MTPAF <> TobGA.GetDouble('GA_PAHT') then
            TobGA.SetDouble('GA_PAHT', MTPAF);
        end;
        RecalculPrPV(TobGA,TobGCA);
      finally
        TobGF.free;
      end;
    end;
  end;
end;

procedure TFRecupTarifFour.AffichageErreurs;
begin
	RapportsErreurs (TOBErreurs);
end;

function TFRecupTarifFour.GetSeparator : string;
begin
  case Tools.CaseFromString(TobParFou.GetString('GRF_SEPARATEUR'), ['AUT', 'TAB', 'PIP', 'PTV']) of
    {AUT} 0 : Result := TobParFou.GetString('GRF_SEPTEXTE');
    {TAB} 1 : Result := #9;
    {PIP} 2 : Result := '|';
    {PTV} 3 : Result := ';';
  else
    Result := '';
  end;
end;

procedure TFRecupTarifFour.CBCreationCatalogueClick(Sender: TObject);
begin
  inherited;
  if CBCreationCatalogue.Checked then
  begin
    CBCreationArticle.Enabled := False;
    chkFamSsFam.Enabled       := False;
    CBCreationArticle.Checked := False;
    chkFamSsFam.Checked       := False;
  end else
  begin
    CBCreationArticle.Enabled := True;
    chkFamSsFam.Enabled       := True;
  end;
end;

procedure TFRecupTarifFour.CBCreationArticleClick(Sender: TObject);
begin
  inherited;
  if CBCreationArticle.Checked then
  begin
    CBCreationCatalogue.Enabled := False;
    CBCreationCatalogue.Checked := False;
  end else
    CBCreationCatalogue.Enabled := True;
end;

function TFRecupTarifFour.ReplaceChar(Value, OldPatern, NewPatern : string) : string;
begin
  if (OldPatern <> '') and (pos(OldPatern, Value) > 0) then
    Result := StringReplace(Value, OldPatern, NewPatern, [rfReplaceAll])
  else
    Result := Value;
end;

function TFRecupTarifFour.GetTslFields(pPrefix : string) : TStringList;
begin
  case Tools.CaseFromString(pPrefix, [PrefixGCA, PrefixGA, PrefixGA2, PrefixBFT, PrefixBSF, PrefixGNE, PrefixGNL, PrefixBCB]) of
    {GCA} 0 : Result := TslGCAFields;
    {GA}  1 : Result := TslGAFields;
    {GA2} 2 : Result := TslGA2Fields;
    {BFT} 3 : Result := TslBFTFields;
    {BSF} 4 : Result := TslBSFFields;
    {GNE} 5 : Result := TslGNEFields;
    {GNL} 6 : Result := TslGNLFields;
    {BCB} 7 : Result := TslBCBFields;
  else
    Result := nil;
  end;
end;

function TFRecupTarifFour.GetFieldsQty(Prefix : string) : integer;
begin
  case Tools.CaseFromString(Prefix, [PrefixGCA, PrefixGA, PrefixGA2, PrefixBFT, PrefixBSF, PrefixGNE, PrefixGNL, PrefixBCB]) of
    {GCA} 0 : Result := TslGCAFields.Count;
    {GA}  1 : Result := TslGAFields.Count;
    {GA2} 2 : Result := TslGA2Fields.Count;
    {BFT} 3 : Result := TslBFTFields.Count;
    {BSF} 4 : Result := TslBSFFields.Count;
    {GNE} 5 : Result := TslGNEFields.Count;
    {GNL} 6 : Result := TslGNLFields.Count;
    {BCB} 7 : Result := TslBCBFields.Count;
  else
    Result := 0;
  end;
end;

function TFRecupTarifFour.GetSpPrepare(Prefix : string; IsInsert : boolean) : string;
var
  Cpts          : integer;
  MaxQ          : integer;
  CptU          : integer;
  CptParam      : integer;
  IndexOfKey    : integer;
  Params        : string;
  InsertParams  : string;
  UpdateParams  : string;
  FieldName     : string;
  FieldsList    : string;
  TslFieldsList : TStringList;
begin
  MaxQ         := GetFieldsQty(Prefix);   // Nombre de champs dans la tables
  Params       := '';
  InsertParams := '';
  UpdateParams := '';
  Result := 'declare @o1 int'
          + ' set @o1 = 0'
          + ' exec sp_prepare @o1 output,N''';
  if IsInsert then
  begin
    for Cpts := 0 to Pred(MaxQ) do
    begin
      Params       := Format('%s,@P%s nvarchar(200)', [Params, IntToStr(Cpts + 1)]);
      InsertParams := Format('%s,@P%s', [InsertParams, IntToStr(Cpts + 1)]);
    end;
  end else
  begin
    Params   := '';
    CptParam := TslCurrentKey.count;
    { Ajout des paramètres correspondant aux champs de la clés qui doivent se trouver au début }
    for CptU := 0 to pred(TslCurrentKey.count) do
      Params := Format('%s,@P%s nvarchar(200)', [Params, IntToStr(CptU+1)]);
    for CptU := 0 to pred(TslPARFOULIG.count) do
    begin
      FieldName  := copy(TslPARFOULIG[CptU], 1, pos('=', TslPARFOULIG[CptU])-1);
      IndexOfKey := TslCurrentKey.IndexOfName(FieldName);
      if     (copy(FieldName, 1, pos('_', FieldName)-1) = Prefix) // Champ de la table en cours de traitement
         and (IndexOfKey = -1)                                    // Pas un champ de la clé
      then
      begin
        inc(CptParam);
        Params       := Format('%s,@P%s nvarchar(200)', [Params, IntToStr(CptParam)]);
        UpdateParams := Format('%s,%s=@P%s', [UpdateParams, FieldName, IntToStr(CptParam)]);
      end;
    end;
  end;
  Result       := Result + copy(Params, 2, length(Params)) + '''';
  InsertParams := copy(InsertParams, 2, length(InsertParams));
  UpdateParams := copy(UpdateParams, 2, length(UpdateParams));
  if IsInsert then
  begin
    TslFieldsList := GetTslFields(Prefix);
    for Cpts := 0 to pred(TslFieldsList.count) do
      FieldsList := Format('%s, %s', [FieldsList, copy(TslFieldsList[Cpts], 1, pos('=', TslFieldsList[Cpts])-1)]);
    FieldsList := copy(FieldsList, 2, length(FieldsList));
    Result := Result + Format(' ,N''INSERT INTO %s (%s) VALUES (%s)''', [PrefixeToTable(Prefix), FieldsList, InsertParams]);
  end else
  begin
    Result := Result + Format(' ,N''UPDATE %s SET %s WHERE ', [PrefixeToTable(Prefix), UpdateParams]);
    case Tools.CaseFromString(Prefix, [PrefixGCA, PrefixGA, PrefixGA2, PrefixBFT, PrefixBSF, PrefixGNE, PrefixGNL]) of
      {GCA} 0 : Result := Result + ' GCA_REFERENCE    = @P1 AND GCA_TIERS = @P2''';
      {GA}  1 : Result := Result + ' GA_ARTICLE       = @P1''';
      {GA2} 2 : Result := Result + ' GA2_ARTICLE      = @P1''';
      {BFT} 3 : Result := Result + ' BFT_FAMILLETARIF = @P1''';
      {BSF} 4 : Result := Result + ' BSF_FAMILLETARIF = @P1 AND BSF_SOUSFAMTARART = @P2''';
      {GNE} 5 : Result := Result + ' GNE_ARTICLE      = @P1 AND GNE_NOMENCLATURE  = @P2''';
      {GNL} 6 : Result := Result + ' GNL_NOMENCLATURE = @P1 AND GNL_NUMLIGNE      = @P2''';
    else
      Result := '';
    end;
  end;
end;

procedure TFRecupTarifFour.SetSpExec(TslData : TStringList; IsInsert : boolean; pEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, Prefix : string);
var
  sParams    : string;
  FieldNT    : string;
  FieldName  : string;
  FieldType  : string;
  FieldValue : string;
  CptA       : integer;
  MaxQ       : integer;
  CptU       : integer;
  IndexOfKey : integer;
  TslFields  : TStringList;

  procedure TabletManagement;
  var
    TobTabletDTB   : TOB;
    AlreadyExist   : boolean;
    tPrefix        : string;
    tFieldTypeName : string;
    tFieldCodeName : string;
    tType          : string;
    tLabel         : string;
    TabletIndex    : integer;

    function GetTabletLabel : string;
    var
      TobTabletFILE : TOB;
      CptTL         : integer;
      tValue        : string;
    begin
      TabletIndex := TslFieldsTablet.IndexOfName('FILE' + FieldName);
      if TabletIndex > -1 then
      begin
        TobTabletFILE := TobTablettes.Detail[TabletIndex];
        for CptTL := 0 to pred(TobTabletFILE.detail.count) do
        begin
          tValue := TobTabletFILE.Detail[CptTL].GetString('VALUE');
          if trim(copy(tValue, 1, 3)) = FieldValue then
          begin
            Result := trim(copy(tValue, 4, length(tValue)));
            break;
          end;
        end;
      end else
        Result := '';
    end;

  begin
    if TabletExist then 
    begin
      TabletIndex := TslFieldsTablet.IndexOfName('DTB' + FieldName);
      if (TabletIndex > -1) and (Tools.GetTypeFieldFromStringType(FieldType) = ttfCombo) then
      begin
        TobTabletDTB   := TobTablettes.Detail[TabletIndex];
        tPrefix        := TobTabletDTB.GetString('DO_PREFIXE');
        tFieldTypeName := Format('%s_TYPE', [tPrefix]);
        tFieldCodeName := Format('%s_CODE', [tPrefix]);
        tType          := TobTabletDTB.GetString('DO_TYPE');
        tLabel         := GetTabletLabel;
        AlreadyExist   := (assigned(TobTabletDTB.FindFirst([tFieldTypeName, tFieldCodeName], [tType, FieldValue], True)));
        Tools.iif(not AlreadyExist, TslInsertTablet, TslUpdateTablet).Add(Format('%s=%s|%s|%s|%s|%s', [tPrefix, tType, FieldValue, tLabel, '', '']));
      end;
    end;
  end;

  procedure AddTobGAGCA;
  var
    Tobl          : TOB;
    Sql           : string;
    FieldValueDte : string;
    Qry           : TQuery;
    tIndexOf      : integer;

    procedure TobPutValue;
    begin
      {$IFNDEF APPSRV}
      if Tools.GetTypeFieldFromStringType(FieldType) = ttfDate then
      begin
        FieldValueDte := Tools.DecodeDateTimeFromQry(FieldValue);
        Tools.TobPutValue(TobL, FieldName, FieldValueDte);
      end else
        Tools.TobPutValue(TobL, FieldName, FieldValue);
      {$ENDIF APPSRV}
    end;

  begin
    { Il faut ajouter dans une tob pour avoir les données nécessaire pour la revalorisation }
    if NeedValorization then
    begin
      case Tools.CaseFromString(Prefix, [PrefixGCA, PrefixGA]) of
        {PrefixGCA} 0 :
        begin
          tIndexOf := Tools.GetIndexOnSortedTsl(TslCacheTobGCA, ReferenceGCA + TheFournisseur);
          if tIndexOf = -1 then
          begin
            TslCacheTobGCA.Add(Format('%s=X', [ReferenceGCA + TheFournisseur]));
            Tobl := TOB.Create('CATALOGU', TobGCA, -1);
          end else
            TobL := TobGCA.Detail[tIndexOf];
          if IsInsert then
          begin
            if tIndexOf = -1 then
            begin
              TobL.SetString('GCA_REFERENCE', ReferenceGCA);
              TobL.SetString('GCA_TIERS'    , TheFournisseur);
            end;
            TobPutValue;
          end else
          if tIndexOf = -1 then
          begin
            Sql := Format('SELECT *'
                        + ' FROM CATALOGU'
                        + ' WHERE GCA_REFERENCE = "%s"'
                        + '   AND GCA_TIERS     = "%s"'
                       , [ReferenceGCA, TheFournisseur]);
            Qry := OpenSql(Sql, True);
            try
              Tobl.SelectDB('', Qry, True);
            finally
              Ferme(Qry);
            end;
          end;
        end;
        {PrefixGA}  1 :
        begin
          tIndexOf := Tools.GetIndexOnSortedTsl(TslCacheTobGA, ReferenceGA);
          if tIndexOf = -1 then
          begin
            TslCacheTobGA.Add(Format('%s=X', [ReferenceGA]));
            Tobl := TOB.Create('ARTICLE', TobGA, -1);
          end else
            TobL := TobGA.Detail[tIndexOf];
          if IsInsert then
          begin
            if tIndexOf = -1 then
              TobL.SetString('GA_ARTICLE', ReferenceGA);
            TobPutValue;
          end else
          if tIndexOf = -1 then
          begin
            Sql := Format('SELECT ARTICLE.*'
                        + '    , GA2_SOUSFAMTARART'
                        + ' FROM ARTICLE'
                        + ' LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE = GA_ARTICLE'
                        + ' WHERE GA_ARTICLE = "%s"'
                        , [ReferenceGA]);
            Qry := OpenSql(Sql, True);
            try
              Tobl.SelectDB('', Qry, True);
            finally
              Ferme(Qry);
            end;
          end;
        end;
      end;
    end;
  end;

  procedure AddRecoverableFields(pFieldName : string);
  var
    FieldsToUpgradedl : string;
    lFieldName        : string;
    GaCalculateValues : string;
    TobGAl            : TOB;
    TobGCAl           : TOB;
    Cpt               : integer;
    lIndice           : integer;

    procedure AddsParams(Position : integer);
    var
      CptP        : integer;
      lFieldValue : string;
    begin
      for CptP := 0 to Position  do
        lFieldValue := Tools.ReadTokenSt_(GaCalculateValues, ';');

      lFieldValue := StringReplace(lFieldValue, ',', '.', [rfReplaceAll]);
      sParams := Format('%s,''%s''', [sParams, lFieldValue]);
    end;

  begin
    if Prefix = PrefixGA then
    begin
      Cpt               := -1;
      FieldsToUpgradedl := FieldsToUpgraded;
      while FieldsToUpgradedl <> '' do
      begin
        lFieldName := Tools.ReadTokenSt_(FieldsToUpgradedl, ';');
        if lFieldName = pFieldName then
        begin
          inc(Cpt);
          { Le champ est présent dans le fichier, on prend la valeur du fichier }
          if GAParametersPrices[Cpt] = 'X' then
          begin
            FieldType  := copy(FieldNT, pos('=', FieldNT)+1, length(FieldNT));
            FieldValue := ExtractValue(lFieldName, FieldType, pEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
            sParams    := Format('%s,''%s''', [sParams, FieldValue]);
          end else
          { Le champ n'est pas présent dans le fichier, on calcule la valeur s'il existe des tarifs }
          if NeedValorization then
          begin
            lIndice := Tools.GetIndexOnSortedTsl(TslGAValuesToUpgraded, ReferenceGA);
            if lIndice = -1 then
            begin
              TobGAl  := TobGA.FindFirst(['GA_ARTICLE'], [ReferenceGA], True);
              TobGCAl := TobGCA.FindFirst(['GCA_REFERENCE', 'GCA_TIERS'], [ReferenceGCA, TheFournisseur], True);
              Valorization(TobGAl, TobGCAl);
              GaCalculateValues := Format('%s;%s;%s;%s', [TobGAl.GetString('GA_DPR'), TobGAl.GetString('GA_PAHT'), TobGAl.GetString('GA_PVHT'), TobGAl.GetString('GA_PVTTC')]);
              TslGAValuesToUpgraded.Add(Format('%s=%s', [ReferenceGA, GACalculateValues]));
            end else
              GaCalculateValues := copy(TslGAValuesToUpgraded[lIndice], pos('=', TslGAValuesToUpgraded[lIndice]) + 1, length(TslGAValuesToUpgraded[lIndice]));
            case Tools.CaseFromString(lFieldName, ['GA_DPR', 'GA_PAHT', 'GA_PVHT', 'GA_PVTTC']) of
              {GA_DPR}   0 : AddsParams(0);
              {GA_PAHT}  1 : AddsParams(1);
              {GA_PVHT}  2 : AddsParams(2);
              {GA_PVTTC} 3 : AddsParams(3);
            end;
          end else
            sParams := Format('%s,''%s''', [sParams, '0']);
        end;
      end;
    end;
  end;

begin
  TslCurrentKey.Clear;
  case Tools.CaseFromString(Prefix, [PrefixGCA, PrefixGA, PrefixGA2, PrefixBFT, PrefixBSF, PrefixGNE, PrefixGNL, PrefixBCB]) of
    {PrefixGCA} 0 :
    begin
      TslCurrentKey.Add(Format('GCA_REFERENCE=%s', [ReferenceGCA]));
      TslCurrentKey.Add(Format('GCA_TIERS=%s'    , [TheFournisseur]));
    end;
    {PrefixGA}  1 :
    begin
      TslCurrentKey.Add(Format('GA_ARTICLE=%s', [ReferenceGA]));
    end;
    {PrefixGA2} 2 :
    begin
      TslCurrentKey.Add(Format('GA2_ARTICLE=%s', [ReferenceGA]));
    end;
    {PrefixBFT} 3 :
    begin
      TslCurrentKey.Add(Format('BFT_FAMILLETARIF=%s', [FamillyTarif]));
    end;
    {PrefixBSF} 4 :
    begin
      TslCurrentKey.Add(Format('BSF_FAMILLETARIF=%s' , [FamillyTarif]));
      TslCurrentKey.Add(Format('BSF_SOUSFAMTARART=%s', [SubFamillyTarif]));
    end;
    {PrefixGNE} 5 :
    begin
      TslCurrentKey.Add(Format('GNE_ARTICLE=%s'     , [ReferenceGA]));
      TslCurrentKey.Add(Format('GNE_NOMENCLATURE=%s', [ReferenceGA]));
    end;
    {PrefixGNL} 6 :
    begin
      TslCurrentKey.Add(Format('GNL_NOMENCLATURE=%s', [ReferenceGA]));
      TslCurrentKey.Add(Format('GNL_NUMLIGNE=%s'    , ['1']));
    end;
    {PrefixBCB} 7 :
    begin
      TslCurrentKey.Add(Format('BCB_NATURECAB=%s' , [TobParFou.GetString('GRF_TYPEARTICLE')]));
      TslCurrentKey.Add(Format('BCB_IDENTIFCAB=%s', [ReferenceGA]));
      TslCurrentKey.Add(Format('BCB_CODEBARRE=%s' , [Barcode]));
    end;
  end;
  if TslData.Count = 0 then
    TslData.Add(GetSpPrepare(Prefix, IsInsert));
  sParams := '';
  MaxQ    := GetFieldsQty(Prefix);   // Nombre de champs dans la tables
  if IsInsert then
  begin
    TslFields := GetTslFields(Prefix);
    if assigned(TslFields) then
    begin
      for CptA := 0 to Pred(MaxQ) do
      begin
        FieldNT    := TslFields[CptA];
        FieldName  := copy(FieldNT, 1, pos('=', FieldNT)-1);
        FieldType  := copy(FieldNT, pos('=', FieldNT)+1, length(FieldNT));
        FieldValue := ExtractValue(FieldName, FieldType, pEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
        AddTobGAGCA;
        if pos(';' + FieldName + ';', FieldsToUpgraded) = 0 then
          sParams    := Format('%s,''%s''', [sParams, FieldValue])
        else
          AddRecoverableFields(FieldName);
        TabletManagement;
      end;
    end;
  end else
  begin
    { Ajout des valeurs des champs de la clés qui doivent se trouver au début }
    for CptU := 0 to pred(TslCurrentKey.count) do
      sParams := Format('%s,''%s''', [sParams, copy(TslCurrentKey[CptU], pos('=', TslCurrentKey[CptU])+1, length(TslCurrentKey[CptU]))]);
    for CptU := 0 to pred(TslPARFOULIG.count) do
    begin
      FieldNT    := TslPARFOULIG[CptU];
      FieldName  := copy(FieldNT, 1, pos('=', FieldNT)-1);
      FieldType  := copy(FieldNT, pos('=', FieldNT)+1, length(FieldNT));
      while pos('|', FieldType) > 0 do
        FieldType := copy(FieldType, pos('|', FieldType) + 1, length(FieldType));
      IndexOfKey := TslCurrentKey.IndexOfName(FieldName);
      if     (copy(FieldName, 1, pos('_', FieldName)-1) = Prefix) // Champ de la table en cours de traitement
         and (IndexOfKey = -1)                                    // Pas un champ de la clé
      then
      begin
        FieldValue := ExtractValue(FieldName, FieldType, pEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
        AddTobGAGCA;
        if pos(';' + FieldName + ';', FieldsToUpgraded) = 0 then
          sParams    := Format('%s,''%s''', [sParams, FieldValue])
        else
          AddRecoverableFields(FieldName);
        TabletManagement;
      end;
    end;
  end;
  sParams := copy(sParams, 2, length(sParams));
  TslData.Add(Format('exec sp_execute @o1, %s', [sParams]));
end;

procedure TFRecupTarifFour.InsertUpdateInDB(pCnx : TADOConnection; Prefix : string; IsInsert : boolean; TslData : TStringList);
var
  IULines : string;

    function DropOrCreateIndexes(tiaAction : tIndexesAction; TableName : string) : boolean;
    var
      Cpti         : integer;
      sType        : string;
      IndexesValue : string;
    begin
      Result := True;
      BeginTrans;
      try
        for Cpti := 0 to pred(TslCacheIndexes.count) do
        begin
          if copy(TslCacheIndexes[Cpti], 2, pos('=', TslCacheIndexes[Cpti])-2) = TableName then
          begin
            sType        := copy(TslCacheIndexes[Cpti], 1, 1);
            IndexesValue := copy(TslCacheIndexes[Cpti], pos('=', TslCacheIndexes[Cpti])+1, length(TslCacheIndexes[Cpti]));
            {$IFNDEF APPSRV}
            Result := Tools.DropOrCreateNonClusteredIndexes(tiaAction, TableName, sType, IndexesValue);
            {$ENDIF APPSRV}
            if not Result then
              Break;
          end;
        end;
      finally
        if Result then
        begin
          CommitTrans;
        end else
        begin
          case tiaAction of
            tia_Drop   : ListRecap.Items.Add(Format('Erreur lors de la suppression des indexes de la table %s. Traitement abandonné.', [TableName]));
            tia_Create : ListRecap.Items.Add('Erreur lors de la création des indexes.');
          end;
          Rollback;
        end;
      end;
    end;

begin
  if TslData.Count > 1 then
  begin
    IULines := IntToStr(TslData.Count-2);
    MoveCurProgressForm(Format('%s des données (%s ligne(s)).', [Tools.iif(IsInsert, 'Création', 'Mise à jour'), IULines]));
    try
      if DropOrCreateIndexes(tia_Drop, PrefixeToTable(Prefix)) then
      begin
        try
          try
            pCnx.Execute(TslData.Text);
          except
            on E : Exception do
              ListRecap.Items.Add(Format('*** Erreur lors de %s de %s ligne(s). Erreur : %s', [Tools.iif(IsInsert, 'création', 'mise à jour'), IULines, E.message]));
          end;
        finally
          DropOrCreateIndexes(tia_Create, PrefixeToTable(Prefix));
        end;
      end;
    finally
      TslData.Clear;
      TslData.Add(GetSpPrepare(Prefix, IsInsert));
    end;
  end;
end;

function TFRecupTarifFour.SupplierIsValid(WithMsg : boolean=True) : boolean;
begin
  if TheFournisseur <> '' then
  begin
    Result := ExisteSql(Format('SELECT T_TIERS FROM TIERS WHERE T_TIERS = "%s"', [TheFournisseur]));
    if (not Result) and (WithMsg) then
      PGIError(Format('Le fournisseur "%s" est inexistant.', [TheFournisseur]));
  end else
    Result := True;
end;

function TFRecupTarifFour.GPFParamIsValid : boolean;
var
  GpfComptaArticle : string;
  GpfCodeTaxe      : string;
  MsgArr           : array [0..4] of string;
  Qry              : TQuery;

  function ValueExist(FieldName : string; MsgNumber : integer) : boolean;
  var
    FieldValue : string;
    Sql        : string;
  begin
    FieldValue := TobParFou.GetString(FieldName);
    case Tools.CaseFromString(FieldName, ['GRF_TYPEARTICLE', 'GRF_FAMILLETAXE1', 'GRF_COMPTAARTICLE']) of
      {GRF_TYPEARTICLE}   0 : Sql := Format('SELECT CO_CODE FROM COMMUN   WHERE CO_TYPE = "TYA" AND CO_CODE = "%s"', [FieldValue]);
      {GRF_FAMILLETAXE1}  1 : Sql := Format('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE = "TX1" AND CC_CODE = "%s"', [FieldValue]);
      {GRF_COMPTAARTICLE} 2 : Sql := Format('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE = "GCA" AND CC_CODE = "%s"', [FieldValue]);
    end;
    Result := ExisteSql(Sql);
    if not Result then
      PGIError(Format(MsgArr[MsgNumber], [FieldValue]));
  end;

begin
  MsgArr[0] := 'Le code du profile article "%s" est inexistant.';
  MSgArr[1] := 'Le type d''article à générer "%s" est inexistant.';
  MSgArr[2] := 'La famille de taxe "%s" est inexistante.';
  MSgArr[3] := 'La famille comptable "%s" est inexistante.';
  GpfComptaArticle := '';
  GpfCodeTaxe      := '';
  Result := SupplierIsValid(True);
  if (Result) and (TobParFou.GetString('GRF_PROFILARTICLE') <> '') then
  begin
    Qry := OpenSql(Format('SELECT GPF_COMPTAARTICLE, GPF_CODETAXE FROM PROFILART WHERE GPF_PROFILARTICLE = "%s"', [TobParFou.GetString('GRF_PROFILARTICLE')]), True);
    try
      Result := (not Qry.Eof);
      if Result then
      begin
        GpfComptaArticle := Qry.Fields[0].AsString;
        GpfCodeTaxe      := Qry.Fields[1].AsString;
      end else
        PGIError(Format(MsgArr[0], [TobParFou.GetString('GPF_PROFILARTICLE')]));
    finally
      Ferme(Qry);
    end;
  end;
  if Result                               then Result := ValueExist('GRF_TYPEARTICLE', 1);
  if (Result) and (GpfCodeTaxe = '')      then Result := ValueExist('GRF_FAMILLETAXE1', 2);
  if (Result) and (GpfComptaArticle = '') then Result := ValueExist('GRF_COMPTAARTICLE', 3);
end;

function TFRecupTarifFour.SetNeedValorization : boolean;
begin
  Result := (    (ExisteSql('SELECT GF_TARIFARTICLE FROM TARIF WHERE GF_TARIFARTICLE <> '' AND GF_NATUREAUXI = "FOU"')) // Il existe des tarifs d'achat
             and (   (GAParametersPrices[0] = '-')                                                                      // Au moins un des 4 champs à valoriser n'est pas présent dans le fichier
                  or (GAParametersPrices[1] = '-')
                  or (GAParametersPrices[2] = '-')
                  or (GAParametersPrices[3] = '-')
                 )
             and (TobParFou.GetBoolean('GRF_FOURPRINC'))                                                                // Le fournisseur paramétré est défini comme principal
             and (   (TobParFou.GetString('GRF_TYPEARTICLE') = 'MAR')                                                   // Le type d'article est MAR ou ARP ou PRE
                  or (TobParFou.GetString('GRF_TYPEARTICLE') = 'ARP')
                  or (TobParFou.GetString('GRF_TYPEARTICLE') = 'PRE')
                 )
             and (   (TslPARFOULIG.IndexOfName('GCA_PRIXBASE')  > - 1)                                                  // Un des champs prix est paramétré
                  or (TslPARFOULIG.IndexOfName('GCA_PRIXVENTE') > - 1)
                  or (TslPARFOULIG.IndexOfName('GCA_DPA')       > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_DPA')        > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_PAHT')       > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_PRHT')       > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_PMAP')       > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_PMRP')       > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_DPR')        > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_PVHT')       > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_PVTTC')      > - 1)
                  or (TslPARFOULIG.IndexOfName('GA_PAUA')       > - 1)
                 )
             );
end;

procedure TFRecupTarifFour.LoadCache;
var
  CptP          : integer;
  CptT          : integer;
  TobL          : TOB;
  Sql           : string;
  FieldName     : string;
  tPrefix       : string;
  tType         : string;
  TableName     : string;
  TabletFile    : string;
  TobTabletteL  : TOB;
  TobTabletteLF : TOB;
  Qry           : TQuery;
  TslTabletc    : TStringList;
  TslTabletFile : TStringList;

  procedure AddTableIndexe(TableName : string);
 {$IFNDEF APPSRV}
  var
    IndexesValue : string;
 {$ENDIF APPSRV}
  begin
    {$IFNDEF APPSRV}
    IndexesValue := Tools.GetTableNonClusteredIndexes(TableName);
    if IndexesValue <> '' then TslCacheIndexes.Add(Format('t%s=%s', [TableName, IndexesValue]));
    IndexesValue := Tools.GetAdditionalNonClusteredIndexes(TableName);
    if IndexesValue <> '' then TslCacheIndexes.Add(Format('a%s=%s', [TableName, IndexesValue]));
    {$ENDIF APPSRV}
  end;

  procedure AddTabletInTob(pOrigin : string; pTobL : TOB);
  begin
    pTobL.AddChampSupValeur('DO_PREFIXE', tPrefix);
    pTobL.AddChampSupValeur('DO_TYPE'   , tType);
    pTobL.AddChampSupValeur('ORIGIN'    , pOrigin);
    pTobL.AddChampSupValeur('FIELDNAME' , FieldName);
    pTobL.AddChampSupValeur('START'     , TobL.GetString('GFL_OFFSET'));
    pTobL.AddChampSupValeur('LENGTH'    , TobL.GetString('GFL_LONGUEUR'));
  end;

begin
  TslTabletc := TStringList.Create;
  try
    TslTabletc.Sorted := True;
    TslTabletFile    := TStringList.Create;
    try
      TslTabletFile.Sorted := True;
      { Champs paramétrés }
      for CptP := 0 to pred(TobParFouLig.detail.count) do
      begin
        TobL := TobParFouLig.detail[CptP];
        if TobL.GetString('GFL_CHAMP') <> 'GA_DATESUPPRESSION' then // Exception pour ce champ, il ne faut pas le prendre en compte car il est paramétré dans l'en-tête
        begin
          TslPARFOULIG.Add(Format('%s=%s|%s|%s|%s', [TobL.GetString('GFL_CHAMP'), TobL.GetString('GFL_OFFSET'), TobL.GetString('GFL_LONGUEUR'), TobL.GetString('GFL_NBDECIMALE'), TobL.GetString('DH_TYPECHAMP')]));
          { Prépare la tob des tablettes si nécessaire }
          if Tools.GetTypeFieldFromStringType(TobL.GetString('DH_TYPECHAMP')) = ttfCombo then
          begin
            FieldName := TobL.GetString('GFL_CHAMP');
            Sql       := Format('SELECT DO_PREFIXE, DO_TYPE FROM DECOMBOS WHERE DO_PREFIXE IN ("CC", "CO", "YX") AND DO_NOMCHAMP LIKE "%s%s%s" ', ['%', Tools.GetSuffix(FieldName), '%']);
            Qry       := OpenSql(Sql, True);
            try
              tPrefix := Tools.iif(not Qry.Eof, Qry.Fields[0].AsString, '');
              tType   := Tools.iif(not Qry.Eof, Qry.Fields[1].AsString, '');
            finally
              Ferme(Qry);
            end;
            if tPrefix <> '' then
            begin
              TabletExist := True;
              if Tools.GetIndexOnSortedTsl(TslTabletc, tPrefix + tType) = -1 then
              begin
                TslTabletc.Add(Format('%s=X', [tPrefix + tType]));
                { Charge les données de la tablette dans la bases }
                TslFieldsTablet.Add(Format('%s=X', ['DTB' + FieldName]));
                TableName    := PrefixeToTable(tPrefix);
                TobTabletteL := TOB.Create('_TABLETTES', TobTablettes, -1);
                AddTabletInTob('DTB', TobTabletteL);
                Sql := Format('SELECT * FROM %s WHERE %s_TYPE = "%s"', [TableName, tPrefix, tType]);
                Qry := OpenSql(Sql, True);
                try
                  TobTabletteL.LoadDetailDB(TableName, '', '', Qry, False);
                finally
                  Ferme(Qry);
                end;
                { Charge les données de la tablette dans le fichier associé s'il en existe un }
                TabletFile := ExtractFilePath(GRF_FICHIER.text) + TobL.GetString('GFL_FICHIER');
                if (TabletFile <> '') and (FileExists(TabletFile)) then
                begin
                  TslTabletFile.LoadFromFile(TabletFile);
                  if TslTabletFile.Count > 0 then
                  begin
                    TslFieldsTablet.Add(Format('%s=X', ['FILE' + FieldName]));
                    TobTabletteL := TOB.Create('_TABLETTES', TobTablettes, -1);
                    AddTabletInTob('FILE', TobTabletteL);
                    for CptT := 0 to pred(TslTabletFile.Count) do
                    begin
                      TobTabletteLF := TOB.Create('_FILE', TobTabletteL, -1);
                      TobTabletteLF.AddChampSupValeur('VALUE', TslTabletFile[CptT]);
                    end;
                    TslTabletFile.Clear;
                  end;
                end;
              end;
            end;
          end;
        end;
      end;
      { Liste des champs et types des tables }
      {$IFNDEF APPSRV}
      if CBCreationCatalogue.checked then
        Tools.GetFieldsListAndTypeFromPrefix(PrefixGCA, TslGCAFields);
      if CBCreationArticle.checked then
      begin
        Tools.GetFieldsListAndTypeFromPrefix(PrefixGA , TslGAFields);
        Tools.GetFieldsListAndTypeFromPrefix(PrefixGA2, TslGA2Fields);
        Tools.GetFieldsListAndTypeFromPrefix(PrefixGNE, TslGNEFields);
        Tools.GetFieldsListAndTypeFromPrefix(PrefixGNL, TslGNLFields);
        Tools.GetFieldsListAndTypeFromPrefix(PrefixBCB, TslBCBFields);
      end;
      if CreateFamily then
      begin
        Tools.GetFieldsListAndTypeFromPrefix(PrefixBFT, TslBFTFields);
        Tools.GetFieldsListAndTypeFromPrefix(PrefixBSF, TslBSFFields);
      end;
      {$ENDIF APPSRV}
      { Indexes des tables si nécessaire }
      if CBCreationCatalogue.checked then
        AddTableIndexe(PrefixeToTable(PrefixGCA));
      if CBCreationArticle.checked then
      begin
        AddTableIndexe(PrefixeToTable(PrefixGA));
        AddTableIndexe(PrefixeToTable(PrefixGA2));
        AddTableIndexe(PrefixeToTable(PrefixGNE));
        AddTableIndexe(PrefixeToTable(PrefixGNL));
        if TslPARFOULIG.IndexOfName('BCB_CODEBARRE') > -1 then
          AddTableIndexe(PrefixeToTable(PrefixBCB));
        if CreateFamily then
        begin
          AddTableIndexe(PrefixeToTable(PrefixBFT));
          AddTableIndexe(PrefixeToTable(PrefixBSF));
        end;
      end;
    finally
      FreeAndNil(TslTabletFile);
    end;
  finally
    FreeAndNil(TslTabletc);
  end;
end;

function TFRecupTarifFour.GetExistSql(ReferenceGCA, ReferenceGA : string) : string;

  function GetJoinGA : string;
  begin
    Result := ' LEFT JOIN ARTICLE ON GA_ARTICLE = GCA_ARTICLE';
  end;

  function GetJoinGA2 : string;
  begin
    Result := ' LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE = GA_ARTICLE';
  end;

  function GetJoinBFT : string;
  begin
    Result := ' LEFT JOIN BTFAMILLETARIF ON BFT_FAMILLETARIF = GA_TARIFARTICLE';
  end;

  function GetJoinBSF : string;
  begin
    Result := ' LEFT JOIN BTSOUSFAMILLETARIF ON BSF_FAMILLETARIF = GA_TARIFARTICLE AND BSF_SOUSFAMTARART = GA2_SOUSFAMTARART';
  end;

  function GetJoinBCB : string;
  begin
    Result := Tools.iif(CreateBCB, ' LEFT JOIN BTCODEBARRE ON BCB_IDENTIFCAB = GA_ARTICLE', '');
  end;

  function GetWherePrincipalBCG : string;
  begin
    Result := Tools.iif(CreateBCB, ', (SELECT BCB_CABPRINCIPAL FROM BTCODEBARRE WHERE BCB_IDENTIFCAB=GA_ARTICLE AND BCB_CABPRINCIPAL = "X")', ', NULL');
  end;

  function GetWhereGCA : string;
  begin
    Result := ' WHERE GCA_REFERENCE = "%s" AND GCA_TIERS = "%s"';
  end;

  function GetWhereGA : string;
  begin
    Result := ' WHERE GA_ARTICLE = "%s"';
  end;
    
begin
  { Catalogue article et famille-sfamille }
  if (CBCreationCatalogue.Checked) and (CBCreationArticle.Checked) and (CreateFamily) then
    Result := Format('SELECT GCA_REFERENCE'
                   + '     , GA_ARTICLE'
                   + '     , BFT_FAMILLETARIF'
                   + '     , BSF_FAMILLETARIF'
                   + '     , GA_TARIFARTICLE'
                   + '     , GA2_SOUSFAMTARART'
                   + '     , GA_TYPEARTICLE'
                   + Tools.iif(CreateBCB, ', BCB_CODEBARRE', ', NULL AS BCB_CODEBARRE')
                   + GetWherePrincipalBCG + ' AS BCBEXIST'
                   + ' FROM CATALOGU'
                   + GetJoinGA
                   + GetJoinGA2
                   + GetJoinBCB
                   + GetJoinBFT
                   + GetJoinBSF
                   + GetWhereGCA
                   , [ReferenceGCA, TheFournisseur])
  { Catalogue et article }
  else if (CBCreationCatalogue.Checked) and (CBCreationArticle.Checked) then
    Result := Format('SELECT GCA_REFERENCE'
                   + '     , GA_ARTICLE'
                   + '     , NULL AS BFT_FAMILLETARIF'
                   + '     , NULL AS BSF_FAMILLETARIF'
                   + '     , NULL AS GA_TARIFARTICLE'
                   + '     , NULL AS GA2_SOUSFAMTARART'
                   + '     , GA_TYPEARTICLE'
                   + Tools.iif(CreateBCB, ', BCB_CODEBARRE', ', NULL AS BCB_CODEBARRE')
                   + GetWherePrincipalBCG + ' AS BCBEXIST'
                   + ' FROM CATALOGU'
                   + GetJoinGA
                   + GetJoinBCB
                   + GetWhereGCA
                   , [ReferenceGCA, TheFournisseur])
  { Article et famille-sfamille }
  else if (CBCreationArticle.checked) and (CreateFamily) then
    Result := Format('SELECT NULL AS GCA_REFERENCE'
                   + '     , GA_ARTICLE'
                   + '     , BFT_FAMILLETARIF'
                   + '     , BSF_FAMILLETARIF'
                   + '     , GA_TARIFARTICLE'
                   + '     , GA2_SOUSFAMTARART'
                   + '     , GA_TYPEARTICLE'
                   + Tools.iif(CreateBCB, ', BCB_CODEBARRE', ', NULL AS BCB_CODEBARRE')
                   + GetWherePrincipalBCG + ' AS BCBEXIST'
                   + ' FROM ARTICLE'
                   + GetJoinGA2
                   + GetJoinBCB
                   + GetJoinBFT
                   + GetJoinBSF
                   + GetWhereGA
                   , [ReferenceGA])
  { Catalogue seul }
  else if (CBCreationCatalogue.checked) then
    Result := Format('SELECT GCA_REFERENCE'
                   + '     , NULL AS GA_ARTICLE'
                   + '     , NULL AS BFT_FAMILLETARIF'
                   + '     , NULL AS BSF_FAMILLETARIF'
                   + '     , NULL AS GA_TARIFARTICLE'
                   + '     , NULL AS GA2_SOUSFAMTARART'
                   + '     , NULL AS GA_TYPEARTICLE'
                   + '     , NULL AS BCB_CODEBARRE'
                   + '     , NULL AS BCBEXIST'
                   + ' FROM CATALOGU'
                   + GetWhereGCA
                   , [ReferenceGCA, TheFournisseur])
  { Article seul }
  else if (CBCreationArticle.checked) then
    Result := Format('SELECT NULL AS GCA_REFERENCE'
                   + '     , GA_ARTICLE'
                   + '     , NULL AS BFT_FAMILLETARIF'
                   + '     , NULL AS BSF_FAMILLETARIF'
                   + '     , NULL AS GA_TARIFARTICLE'
                   + '     , NULL AS GA2_SOUSFAMTARART'
                   + '     , GA_TYPEARTICLE'
                   + Tools.iif(CreateBCB, ', BCB_CODEBARRE', ', NULL AS BCB_CODEBARRE')
                   + GetWherePrincipalBCG + ' AS BCBEXIST'
                   + ' FROM ARTICLE'
                   + GetJoinBCB
                   + GetWhereGA
                   , [ReferenceGA])
  else
    Result := '';
end;

procedure TFRecupTarifFour.SetTslFieldsList(ReferenceGCA, ReferenceGA, Prefix : string; TslData : TStringList);
var
  TobFields : TOB;
  Sql       : string;
begin
  TobFields := TOB.Create('_FIELDS', nil, -1);
  try
    Sql := Format('SELECT DH_NOMCHAMP FROM DECHAMPS WHERE DH_PREFIXE = "%s" ORDER BY DH_NUMCHAMP', [Prefix]);
    TobFields.LoadDetailFromSQL(Sql);
   {$IFNDEF APPSRV}
    Tools.TobToTStringList(TobFields, TslData);
   {$ENDIF APPSRV}
    TobFields.ClearDetail;
  finally
    FreeAndNil(TobFields);
  end;
end;

function TFRecupTarifFour.GetDefaultValue(ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, Prefix, FieldName, FieldType : string) : string;
var
  TobGPFl     : TOB;
  Qry         : TQuery;
  ProfilCode  : string;
  FieldsList  : string;
  FieldsArray : array of string;
  FindProfil  : boolean;

  function GetDefaultValueFromType : string;
  begin
    {$IFNDEF APPSRV}
    if Tools.GetTypeFieldFromStringType(FieldType) = ttfDate then
      Result := Tools.CastDateTimeForQry(IDate1900)
    else
      Result := Tools.GetDefaultValueFromtTypeField(Tools.GetFieldType(FieldName), False);
    {$ENDIF APPSRV}
  end;

begin
  Result := '';
  case Tools.CaseFromString(Prefix, [PrefixGCA, PrefixGA, PrefixGA2, PrefixBFT, PrefixBSF, PrefixGNE, PrefixGNL, PrefixBCB]) of
    0 : {PrefixGCA}
    begin
      { Affecte la valeur par défaut }
      Result := DefaultValuesGCA.GetDefaultValue(FieldName);
      { Si pas la valeur par défaut }
      if Result = '' then
      begin
        FieldsList := DefaultValuesGCA.GetFieldsList(cfltCatalogImport);
        if Tools.StringInList(FieldName, FieldsList) then
        begin
          SetLength(FieldsArray, 0);
          SetLength(FieldsArray, Tools.CountOccurenceString(FieldsList, ';')+1);
          Tools.SetArray(FieldsArray, FieldsList);
          case Tools.CaseFromString(FieldName, FieldsArray) of
            {GCA_REFERENCE} 0 : Result := ReferenceGCA;
            {GCA_TIERS}     1 : Result := TheFournisseur;
            {GCA_DATESUP}   2 : Result := Tools.CastDateTimeForQry(StrToDate(DATESUP.Text));
            {GCA_ARTICLE}   3 : Result := Tools.iif(ReferenceGA<>'',CodeArticleUnique2(ReferenceGA, ''),'');
          end;
        end else
          Result := GetDefaultValueFromType;
      end;
    end;
    1,2 : {PrefixGA, PrefixGA2}
    begin
      { Affecte la valeur par défaut }
      Result := DefaultValuesGA.GetDefaultValue(FieldName);
      { Si pas la valeur par défaut }
      if Result = '' then
      begin
        FieldsList := DefaultValuesGA.GetFieldsList(cfltCatalogImport);
        if Tools.StringInList(FieldName, FieldsList) then
        begin
          SetLength(FieldsArray, 0);
          SetLength(FieldsArray, Tools.CountOccurenceString(FieldsList, ';')+1);
          Tools.SetArray(FieldsArray, FieldsList);
          case Tools.CaseFromString(FieldName, FieldsArray) of
            {GA_ARTICLE,GA2_ARTICLE}         0,1  : Result := ReferenceGA;
            {GA_CODEARTICLE,GA2_CODEARTICLE} 2,3  : Result := Trim(copy(ReferenceGA, 1, 18));
            {GA_TYPEARTICLE}                 4    : Result := TobParFou.GetString('GRF_TYPEARTICLE');
            {GA_LIBELLE}                     5    : Result := Trim(copy(ReferenceGA, 1, 18));
            {GA_COMPTAARTICLE}               6    : Result := TobParFou.GetString('GRF_COMPTAARTICLE');
            {GA_DATESUPPRESSION}             7    : Result := Tools.CastDateTimeForQry(StrToDate(DATESUP.Text));
            {GA_TENUESTOCK}                  8    : Result := TobParFou.GetString('GRF_TENUESTOCK');
            {GA_FAMILLETAXE1}                9    : Result := TobParFou.GetString('GRF_FAMILLETAXE1');
            {GA_REMISEPIED}                  10   : Result := TobParFou.GetString('GRF_REMISEPIED');
            {GA_ESCOMPTABLE}                 11   : Result := TobParFou.GetString('GRF_ESCOMPTABLE');
            {GA_FOURNPRINC}                  12   : Result := Tools.iif(TobParFou.GetBoolean('GRF_FOURPRINC'), TheFournisseur, '');
            {GA_CREERPAR}                    13   : Result := 'IMP';
          end;
        end else
          Result := GetDefaultValueFromType;
      end;
      if (Result = '') and (ProfilCode <> '') then
      begin
        TobGPFl    := TobGPF.FindFirst(['GPF_PROFILARTICLE'], [ProfilCode], True);
        FindProfil := assigned(TobGPFl);
        if not FindProfil then
        begin
          TobGPFl := TOB.Create('PROFILART', TobGPF, -1);
          Qry := OpenSql(Format('SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE = "%s"', [ProfilCode]), True);
          try
            FindProfil := (not Qry.Eof);
            if FindProfil then
              TobGPFl.SelectDB('', Qry);
          finally
            Ferme(Qry);
          end;
        end;
        if FindProfil then
        begin
          FieldsList := DefaultValuesGA.GetFieldsList(cfltProfile);
          SetLength(FieldsArray, 0);
          SetLength(FieldsArray, Tools.CountOccurenceString(FieldsList, ';')+1);
          Tools.SetArray(FieldsArray, FieldsList);
          case Tools.CaseFromString(FieldName, FieldsArray) of
            {GA_FAMILLENIV1}     0  : Result := TobGPFl.GetString('GPF_FAMILLENIV1');
            {GA_FAMILLENIV2}     1  : Result := TobGPFl.GetString('GPF_FAMILLENIV2');
            {GA_FAMILLENIV3}     2  : Result := TobGPFl.GetString('GPF_FAMILLENIV3');
            {GA_COMPTAARTICLE}   3  : Result := TobGPFl.GetString('GPF_COMPTAARTICLE');
            {GA_TENUESTOCK}      4  : Result := TobGPFl.GetString('GPF_ENUESTOCK');
            {GA_CALCPRIXPR}      5  : Result := TobGPFl.GetString('GPF_CALCPRIXPR');
            {GA_COEFFG}          6  : Result := TobGPFl.GetString('GPF_COEFCALCPR');
            {GA_DPRAUTO}         7  : Result := TobGPFl.GetString('GPF_CALCAUTOPR');
            {GA_LOT}             8  : Result := TobGPFl.GetString('GPF_LOT');
            {GA_NUMEROSERIE}     9  : Result := TobGPFl.GetString('GPF_NUMEROSERIE');
            {GA_CONTREMARQUE}    10 : Result := TobGPFl.GetString('GPF_CONTREMARQUE');
            {GA_REMISEPIED}      11 : Result := TobGPFl.GetString('GPF_REMISEPIED');
            {GA_REMISELIGNE}     12 : Result := TobGPFl.GetString('GPF_REMISELIGNE');
            {GA_ESCOMPTABLE}     13 : Result := TobGPFl.GetString('GPF_ESCOMPTABLE');
            {GA_FAMILLETAXE1}    14 : Result := TobGPFl.GetString('GPF_CODETAXE');
            {GA_COMMISSIONNABLE} 15 : Result := TobGPFl.GetString('GPF_COMMISSIONNABLE');
            {GA_CALCPRIXHT}      16 : Result := TobGPFl.GetString('GPF_CALCPRIXHT');
            {GA_CALCPRIXTTC}     17 : Result := TobGPFl.GetString('GPF_CALCPRIXTTC');
            {GA_COEFCALCHT}      18 : Result := TobGPFl.GetString('GPF_COEFCALCHT');
            {GA_COEFCALCTTC}     19 : Result := TobGPFl.GetString('GPF_COEFCALCTTC');
            {GA_CALCAUTOHT}      20 : Result := TobGPFl.GetString('GPF_CALCAUTOHT');
            {GA_CALCAUTOTTC}     21 : Result := TobGPFl.GetString('GPF_CALCAUTOTTC');
            {GA_TARIFARTICLE}    22 : Result := TobGPFl.GetString('GPF_TARIFARTICLE');
            {GA2_SOUSFAMTARART}  23 : Result := TobGPFl.GetString('GPF_SOUSFAMTARART');
            {GA_PAYSORIGINE}     24 : Result := TobGPFl.GetString('GPF_PAYSORIGINE');
            {GA_ARRONDIPRIX}     25 : Result := TobGPFl.GetString('GPF_ARRONDIPRIX');
            {GA_ARRONDIPRIXTTC}  26 : Result := TobGPFl.GetString('GPF_ARRONDIPRIXTTC');
            {GA_PRIXUNIQUE}      27 : Result := TobGPFl.GetString('GPF_PRIXUNIQUE');
            {GA_COLLECTION}      28 : Result := TobGPFl.GetString('GPF_COLLECTION');
            {GA_FOURNPRINC}      29 : Result := TobGPFl.GetString('GPF_FOURNPRINC');
            {GA_LIBREART1}       30 : Result := TobGPFl.GetString('GPF_LIBREART1');
            {GA_LIBREART2}       31 : Result := TobGPFl.GetString('GPF_LIBREART2');
            {GA_LIBREART3}       32 : Result := TobGPFl.GetString('GPF_LIBREART3');
            {GA_LIBREART4}       33 : Result := TobGPFl.GetString('GPF_LIBREART4');
            {GA_LIBREART5}       34 : Result := TobGPFl.GetString('GPF_LIBREART5');
            {GA_LIBREART6}       35 : Result := TobGPFl.GetString('GPF_LIBREART6');
            {GA_LIBREART7}       36 : Result := TobGPFl.GetString('GPF_LIBREART7');
            {GA_LIBREART8}       37 : Result := TobGPFl.GetString('GPF_LIBREART8');
            {GA_LIBREART9}       38 : Result := TobGPFl.GetString('GPF_LIBREART9');
            {GA_LIBREARTA}       39 : Result := TobGPFl.GetString('GPF_LIBREARTA');
            {GA_VALLIBRE1}       40 : Result := TobGPFl.GetString('GPF_VALLIBRE1');
            {GA_VALLIBRE2}       41 : Result := TobGPFl.GetString('GPF_VALLIBRE2');
            {GA_VALLIBRE3}       42 : Result := TobGPFl.GetString('GPF_VALLIBRE3');
            {GA_DATELIBRE1}      43 : Result := TobGPFl.GetString('GPF_DATELIBRE1');
            {GA_DATELIBRE2}      44 : Result := TobGPFl.GetString('GPF_DATELIBRE2');
            {GA_DATELIBRE3}      45 : Result := TobGPFl.GetString('GPF_DATELIBRE3');
            {GA_CHARLIBRE1}      46 : Result := TobGPFl.GetString('GPF_CHARLIBRE1');
            {GA_CHARLIBRE2}      47 : Result := TobGPFl.GetString('GPF_CHARLIBRE2');
            {GA_CHARLIBRE3}      48 : Result := TobGPFl.GetString('GPF_CHARLIBRE3');
            {GA_BOOLLIBRE1}      49 : Result := TobGPFl.GetString('GPF_BOOLLIBRE1');
            {GA_BOOLLIBRE2}      50 : Result := TobGPFl.GetString('GPF_BOOLLIBRE2');
            {GA_BOOLLIBRE3}      51 : Result := TobGPFl.GetString('GPF_BOOLLIBRE3');
          end;
        end;
      end;
    end;
    3 : {PrefixBFT}
    begin
      FieldsList := DefaultValuesBFT.GetFieldsList(cfltCatalogImport);
      if Tools.StringInList(FieldName, FieldsList) then
      begin
        SetLength(FieldsArray, 0);
        SetLength(FieldsArray, Tools.CountOccurenceString(FieldsList, ';')+1);
        Tools.SetArray(FieldsArray, FieldsList);
        case Tools.CaseFromString(FieldName, FieldsArray) of
          {BFT_FAMILLETARIF} 0 : Result := FamillyTarif;
          {BFT_LIBELLE}      1 : Result := Format('Famille : %s', [FamillyTarif]);
        end;
      end else
        Result := GetDefaultValueFromType;
    end;
    4 : {PrefixBSF}
    begin
      FieldsList := DefaultValuesBSF.GetFieldsList(cfltCatalogImport);
      if Tools.StringInList(FieldName, FieldsList) then
      begin
        SetLength(FieldsArray, 0);
        SetLength(FieldsArray, Tools.CountOccurenceString(FieldsList, ';')+1);
        Tools.SetArray(FieldsArray, FieldsList);
        case Tools.CaseFromString(FieldName, FieldsArray) of
          {BSF_FAMILLETARIF}  0 : Result := FamillyTarif;
          {BSF_SOUSFAMTARART} 1 : Result := SubFamillyTarif;
          {BSF_LIBELLE}       2 : Result := Format('Sous-Famille : %s', [SubFamillyTarif]);
        end;
      end else
        Result := GetDefaultValueFromType;
    end;
    5 : {PrefixGNE}
    begin
      { Affecte la valeur par défaut }
      Result := DefaultValueGNE.GetDefaultValue(FieldName);
      { Si pas la valeur par défaut }
      if Result = '' then
      begin
        FieldsList := DefaultValueGNE.GetFieldsList(cfltCatalogImport);
        if Tools.StringInList(FieldName, FieldsList) then
        begin
          SetLength(FieldsArray, 0);
          SetLength(FieldsArray, Tools.CountOccurenceString(FieldsList, ';')+1);
          Tools.SetArray(FieldsArray, FieldsList);
          case Tools.CaseFromString(FieldName, FieldsArray) of
            {GNE_NOMENCLATURE} 0 : Result := ReferenceGA;
            {GNE_LIBELLE}      1 : Result := ReferenceGA;
            {GNE_ARTICLE}      2 : Result := ReferenceGA;
          end;
        end else
          Result := GetDefaultValueFromType;
      end;
    end;
    6 : {PrefixGNL}
    begin
      { Affecte la valeur par défaut }
      Result := DefaultValueGNL.GetDefaultValue(FieldName);
      { Si pas la valeur par défaut }
      if Result = '' then
      begin
        FieldsList := DefaultValueGNL.GetFieldsList(cfltCatalogImport);
        if Tools.StringInList(FieldName, FieldsList) then
        begin
          SetLength(FieldsArray, 0);
          SetLength(FieldsArray, Tools.CountOccurenceString(FieldsList, ';')+1);
          Tools.SetArray(FieldsArray, FieldsList);
          case Tools.CaseFromString(FieldName, FieldsArray) of
            {GNL_NOMENCLATURE} 0 : Result := ReferenceGA;
            {GNL_NUMLIGNE}     1 : Result := '1';
            {GNL_LIBELLE}      2 : Result := Trim(copy(ReferenceGA, 1, 18));
            {GNL_CODEARTICLE}  3 : Result := Trim(copy(ReferenceGA, 1, 18));
            {GNL_ARTICLE}      4 : Result := ReferenceGA;
            {GNL_QTE}          5 : Result := '1';
          end;
        end else
          Result := GetDefaultValueFromType;
      end;
    end;
    7 : {PrefixBCB}
    begin
      { Affecte la valeur par défaut }
      Result := DefaultValueBCB.GetDefaultValue(FieldName);
      { Si pas la valeur par défaut }
      if Result = '' then
      begin
        FieldsList := DefaultValueBCB.GetFieldsList(cfltCatalogImport);
        if Tools.StringInList(FieldName, FieldsList) then
        begin
          SetLength(FieldsArray, 0);
          SetLength(FieldsArray, Tools.CountOccurenceString(FieldsList, ';')+1);
          Tools.SetArray(FieldsArray, FieldsList);
          case Tools.CaseFromString(FieldName, FieldsArray) of
            {BCB_NATURECAB}    0 : Result := TobParFou.GetString('GRF_TYPEARTICLE');
            {BCB_IDENTIFCAB}   1 : Result := ReferenceGA;
            {BCB_CABPRINCIPAL} 2 : Result := Tools.iif(ExistBCBPrincipal, '-', 'X');
            {BCB_CODEBARRE}    3 : Result := Barcode;
          end;
        end else
          Result := GetDefaultValueFromType;
      end;
    end;
  end;
end;

function TFRecupTarifFour.ExtractValue(FieldName, FieldType, pEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode : string) : string;
var
  iIndex     : integer;
  iFrom      : integer;
  iLen       : integer;
  CptExtract : integer;
  lstEnreg   : string;
  Value      : string;
  Separator  : string;
begin
  Result := '';
  iIndex := TslPARFOULIG.IndexOfName(FieldName);
  if iIndex > - 1 then
  begin
    if TobParFou.GetBoolean('GRF_TYPEENREG') then
    begin
      Value  := TslPARFOULIG[iIndex];
      Value  := copy(Value, pos('=', Value)+1, length(Value));
      iFrom  := StrToInt(Tools.ReadTokenSt_(Value, '|'));
      iLen   := StrToInt(Tools.ReadTokenSt_(Value, '|'));
      Result := Trim(copy(pEnreg, iFrom, iLen));
    end else
    begin
      lstEnreg  := pEnreg;
      Separator := GetSeparator;
      for CptExtract := 0 to pred(TslPARFOULIG.Count) do
      begin
        Result := Tools.ReadTokenSt_(lstEnreg, Separator);
        if CptExtract = iIndex then
          break;
      end;
    end;
    if pos('''', Result) > 0 then
    begin
      if copy(Result, length(Result), 1) = '''' then
        Result := copy(Result, 1, length(Result) -1);
      Result := ReplaceChar(Result, '''', '''''');
    end;
    if (Tools.GetTypeFieldFromStringType(FieldType) = ttfDate) and (result <> '') then
    begin
      Result := Tools.CastDateTimeForQry(StrToDateTime(result));
    end;
    if (Tools.GetTypeFieldFromStringType(FieldType) = ttfDate) and (result = '') and (FieldName='GCA_DATESUP') then
    begin
      Result := Tools.CastDateTimeForQry(iDate2099);
    end;
    if (Tools.GetTypeFieldFromStringType(FieldType) = ttfNumeric) and (pos(',', Result) > 0) then
      Result := ReplaceChar(Result, ',', '.');
    if (Result = '') and (copy(FieldName, pos('_', FieldName), Length(FieldName)) = '_LIBELLE') then
      Result := Tools.iif(ReferenceGCA <> '', Trim(copy(ReferenceGCA, 1, 18)), Trim(copy(ReferenceGA, 1, 18)));
  end else
    Result := GetDefaultValue(ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, copy(FieldName, 1, pos('_', FieldName)-1), FieldName, FieldType);
end;

procedure TFRecupTarifFour.InsertTablet(pCnx : TADOConnection);
var
  TslInsertCC      : TStringList;
  TslInsertYX      : TStringList;
  TslInsertCO      : TStringList;
  TslUpdateCC      : TStringList;
  TslUpdateYX      : TStringList;
  TslUpdateCO      : TStringList;
  TslCountByTables : TStringList;
  NbRecord         : integer;
  CptIT            : integer;
  sPrepareInsertCC : string;
  sPrepareInsertCO : string;
  sPrepareInsertYX : string;
  sPrepareUpdateCC : string;
  sPrepareUpdateCO : string;
  sPrepareUpdateYX : string;
  sPExecStart      : string;
  LastPrefix       : string;
  LastTableName    : string;
  CurrentTableName : string;
  CountValue       : string;       

  procedure ExecQryIfNeed;

    procedure ExecQry(pPreparel : string; pTslData : TStringList);
    begin
      try
        pCnx.Execute(Format('%s %s', [pPreparel, pTslData.Text]));
      except
        on E : Exception do
          ListRecap.Items.Add(Format('*** Erreur lors de la création de %s ligne(s). Erreur : %s', [IntToStr(pTslData.count), E.message]));
      end;
    end;

    procedure AddCount(pType, Prefix : string; pTslData : TstringList);
    var
      TableName : string;
      iCount    : integer;
      NewQty    : integer;
    begin
      TableName := pType + PrefixeToTable(Prefix);
      iCount    := TslCountByTables.IndexOfName(TableName);
      if iCount = -1 then
        TslCountByTables.Add(Format('%s=%s', [TableName, IntToStr(pTslData.count)]))
      else
      begin
        NewQty := StrToInt(copy(TslCountByTables[iCount], pos('=', TslCountByTables[iCount]) + 1, length(TslCountByTables[iCount]))) + pTslData.Count;
        TslCountByTables[iCount] := Format('%s=%s', [TableName, IntToStr(NewQty)]);
      end;
    end;

  begin
    try
      if TslInsertCC.count > 0 then
      begin
        ExecQry(sPrepareInsertCC, TslInsertCC);
        AddCount('i', 'CC', TslInsertCC);
      end;
      if TslUpdateCC.count > 0 then
      begin
        ExecQry(sPrepareUpdateCC, TslUpdateCC);
        AddCount('u', 'CC', TslUpdateCC);
      end;
      if TslInsertYX.count > 0 then
      begin
        ExecQry(sPrepareInsertYX, TslInsertYX);
        AddCount('i', 'YX', TslInsertYX);
      end;
      if TslUpdateYX.count > 0 then
      begin
        ExecQry(sPrepareUpdateYX, TslUpdateYX);
        AddCount('u', 'YX', TslUpdateYX);
      end;
      if TslInsertCO.count > 0 then
      begin
        ExecQry(sPrepareInsertCO, TslUpdateCO);
        AddCount('i', 'CO', TslInsertCO);
      end;
      if TslUpdateCO.count > 0 then
      begin
        ExecQry(sPrepareUpdateCO, TslUpdateCO);
        AddCount('u', 'CO', TslUpdateCO);
      end;
    finally
      TslInsertCC.Clear;
      TslInsertYX.Clear;
      TslInsertCO.Clear;
      TslUpdateCC.Clear;
      TslUpdateYX.Clear;
      TslUpdateCO.Clear;
    end;
  end;

  function GetBefore : string;
  begin
    Result := 'declare @o1 int '
            + ' set @o1 = 0'
            + ' exec sp_prepare @o1 output'
            + ' ,N''@P1 nvarchar(3), @P2 nvarchar(3), @P3 nvarchar(105), @P4 nvarchar(17), @P5 nvarchar(35)'''
            ;
  end;

  function GetAfterInsert : string;
  begin
    Result := 'VALUES(@P1,@P2,@P3,@P4,@P5)''';
  end;

  procedure AddToTsl(IsInsert : boolean; TslInsUpd : TStringList);
  var
    CptT        : integer;
    Prefix      : string;
    Value       : string;
    sPExec      : string;
    cType       : string;
    cCode       : string;
    cLabel      : string;
    cAbstract   : string;
    cFree       : string;
    CanContinue : boolean;
  begin
    for CptT := 0 to pred(TslInsUpd.count) do
    begin
      inc(NbRecord);
      if NbRecord = MaxQryExec then
      begin
        NbRecord := 0;
        ExecQryIfNeed;
      end;
      CanContinue := True;
      Prefix      := copy(TslInsUpd[CptT], 1, pos('=', TslInsUpd[CptT])-1);
      Value       := copy(TslInsUpd[CptT], pos('=', TslInsUpd[CptT]) + 1, length(TslInsUpd[CptT]));
      cType       := Tools.ReadTokenSt_(Value, '|');
      cCode       := Tools.ReadTokenSt_(Value, '|');
      cLabel      := Tools.ReadTokenSt_(Value, '|');
      cAbstract   := Tools.ReadTokenSt_(Value, '|');
      cFree       := Tools.ReadTokenSt_(Value, '|');
      if cLabel = '' then
      begin
        if not IsInsert then
          CanContinue := False
        else
          cLabel := cCode;
      end;
      if CanContinue then
      begin
        sPExec      := Format('%s ''%s'',''%s'',''%s'',''%s'',''%s'''
                             , [ sPExecStart
                               , cType
                               , cCode
                               , cLabel
                               , cAbstract
                               , cFree
                               ]);
        case Tools.CaseFromString(Prefix, ['CC', 'YX', 'CO']) of
          {CC} 0 : Tools.iif(IsInsert, TslInsertCC, TslUpdateCC).Add(sPexec);
          {YX} 1 : Tools.iif(IsInsert, TslInsertYX, TslUpdateYX).Add(sPexec);
          {CO} 2 : Tools.iif(IsInsert, TslInsertCO, TslUpdateCO).Add(sPexec);
        end;
      end;
    end;
  end;

begin
  TslInsertCC      := TStringList.Create;
  TslInsertYX      := TStringList.Create;
  TslInsertCO      := TStringList.Create;
  TslUpdateCC      := TStringList.Create;
  TslUpdateYX      := TStringList.Create;
  TslUpdateCO      := TStringList.Create;
  TslCountByTables := TStringList.Create;
  try
    sPrepareInsertCC := Format('%s,N''INSERT INTO CHOIXCOD %s', [GetBefore, GetAfterInsert]);
    sPrepareInsertYX := Format('%s,N''INSERT INTO CHOIXEXT %s', [GetBefore, GetAfterInsert]);
    sPrepareInsertCO := Format('%s,N''INSERT INTO COMMUN   %s', [GetBefore, GetAfterInsert]);
    sPrepareUpdateCC := Format('%s,N''UPDATE CHOIXCOD SET CC_LIBELLE=@P3 WHERE CC_TYPE = @P1 AND CC_CODE = @P2''', [GetBefore]);
    sPrepareUpdateYX := Format('%s,N''UPDATE CHOIXEXT SET YX_LIBELLE=@P3 WHERE YX_TYPE = @P1 AND YX_CODE = @P2''', [GetBefore]);
    sPrepareUpdateCO := Format('%s,N''UPDATE COMMUN   SET CO_LIBELLE=@P3 WHERE CO_TYPE = @P1 AND CO_CODE = @P2''', [GetBefore]);
    sPExecStart      := 'exec sp_execute @o1, ';
    LastPrefix       := '';
    NbRecord         := 1;
    AddToTsl(True , TslInsertTablet);
    AddToTsl(False, TslUpdateTablet);
    ExecQryIfNeed;
  finally
    if TslCountByTables.count > 0 then
    begin
      ListRecap.Items.Add('');
      LastTableName := '';
      for CptIT := 0 to pred(TslCountByTables.count) do
      begin
        CountValue       := TslCountByTables[CptIT];
        CurrentTableName := Copy(CountValue, 2, pos('=', CountValue)-2);
        if CurrentTableName <> LastTableName then
        begin
          LastTableName := CurrentTableName;
          ListRecap.Items.Add(Format('Table %s :', [LastTableName]));
        end;
        if Copy(CountValue, 1, 1) = 'i' then
          ListRecap.Items.Add(Format('- Création de %s lignes'   , [copy(CountValue, pos('=', CountValue)+1, length(CountValue))]))
        else
          ListRecap.Items.Add(Format('- Mise à jour de %s lignes', [copy(CountValue, pos('=', CountValue)+1, length(CountValue))]));
      end;
    end;
    FreeAndNil(TslInsertCO);
    FreeAndNil(TslInsertYX);
    FreeAndNil(TslInsertCC);
    FreeAndNil(TslUpdateCC);
    FreeAndNil(TslUpdateYX);
    FreeAndNil(TslUpdateCO);
    FreeAndNil(TslCountByTables);
  end;
end;

procedure TFRecupTarifFour.ExecuteInsertIfNeed(Prefix : string; LinesQty : integer; pCnx : TADOConnection);
begin
  if LinesQty = MaxQryExec then
  begin
    case Tools.CaseFromString(Prefix, [PrefixGCA, PrefixGA, PrefixBFT, PrefixBSF, PrefixGNE, PrefixGNL, PrefixBCB]) of
      {GCA} 0 : begin
                  InsertGCALines := 0;
                  InsertUpdateInDB(pCnx, PrefixGCA, True, TslInsertGCA);
                end;
      {GA}  1 : begin
                  InsertGALines := 0;
                  InsertUpdateInDB(pCnx, PrefixGA, True, TslInsertGA);
                  if CreateGA2 then
                    InsertUpdateInDB(pCnx, PrefixGA2, True, TslInsertGA2);
                end;
      {BFT} 2 : begin
                  InsertBFTLines := 0;
                  InsertUpdateInDB(pCnx, PrefixBFT, True, TslInsertBFT);
                end;
      {BSF} 3 : begin
                  InsertBSFLines := 0;
                  InsertUpdateInDB(pCnx, PrefixBSF, True, TslInsertBSF);
                end;
      {GNE} 4 : begin
                  InsertGNELines := 0;
                  InsertUpdateInDB(pCnx, PrefixGNE, True, TslInsertGNE);
                end;
      {GNL} 5 : begin
                  InsertGNLLines := 0;
                  InsertUpdateInDB(pCnx, PrefixGNL, True, TslInsertGNL);
                end;
      {BCB} 6 : begin
                  InsertBCBLines := 0;
                  InsertUpdateInDB(pCnx, PrefixBCB, True, TslInsertBCB);
                end;
    end;
  end;
end;

procedure TFRecupTarifFour.ExecuteUpdateIfNeed(Prefix : string; LinesQty : integer; pCnx : TADOConnection);
begin
  if LinesQty = MaxQryExec then
  begin
    case Tools.CaseFromString(Prefix, [PrefixGCA, PrefixGA]) of
      0 : {GCA}
      begin
        UpdateGCALines := 0;
        InsertUpdateInDB(pCnx, PrefixGCA, False, TslUpdateGCA);
      end;
      1 : {GA}
      begin
        UpdateGALines := 0;
        InsertUpdateInDB(pCnx, PrefixGA, False, TslUpdateGA);
        if CreateGA2 then
          InsertUpdateInDB(pCnx, PrefixGA2, False, TslUpdateGA2);
      end;
    end;
  end;
end;

function TFRecupTarifFour.RecupereCatalogue : boolean;
var
  SizeFic               : integer;
  lIndexOf              : integer;
  CptE                  : integer;
  StEnreg               : string;
  ReferenceGCA          : string;
  ReferenceGA           : string;
  FamillyTarif          : string;
  SubFamillyTarif       : string;
  Barcode               : string;
  SqlExist              : string;
  StartOn               : string;
  ProfilCode            : string;
  TstValue              : string;
  StartJnal             : string;
  TslFile               : TstringList;
  TSLCacheGCA           : TStringList;
  TSLCacheBFT           : TStringList;
  TslCacheT             : TstringList;
  TslCacheDuplicate     : TStringList;
  TSLCacheBSF           : TStringList;
  TSLCacheGNE           : TSTringList;
  TSLCacheGNL           : TSTringList;
  TSLCacheBCB           : TStringList;
  CataloguExist         : boolean;
  ItemExist             : boolean;
  FamilyExist           : boolean;
  SubFamillyExist       : boolean;
  IsPrixPose            : boolean;
  BarCodeExist          : boolean;
  SupplierInFile        : boolean;
  CanContinue           : boolean;
  Warning               : boolean;
  QryExist              : TQuery;
  TrtStart              : TDateTime;
  Mcd                   : IMCDServiceCOM;
  Cnx                   : TADOConnection;

  procedure AddWarningLineIfNeed;
  begin
    if not Warning then
      ListRecap.Items.Add('Avertissement(s) :');
    Warning := True;
  end;
  
  function GetSizeFile : integer;
  var
    Fichier   : Textfile;
    SearchRec : TSearchRec;
    FileLine  : string;
  begin
    AssignFile(Fichier, GRF_FICHIER.Text);
    Reset(Fichier);
    Readln(Fichier, FileLine);
    CloseFile(Fichier);
    FindFirst(GRF_FICHIER.Text, faAnyFile, SearchRec);
    Result := Trunc(SearchRec.Size / Length(FileLine));
    FindClose(SearchRec);
  end;

begin
  Mcd := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  TrtStart       := Now;
  SizeFic        := GetSizeFile;
  TabletExist    := False;
  Warning        := False;
  ProfilCode     := TobParFou.GetString('GRF_PROFILARTICLE');
  SupplierInfile := TobParFou.GetBoolean('GRF_MULTIFOU');
  Result         := GPFParamIsValid;
  if Result then
  begin
    StartJnal := Format('Début de la récupération : %s', [FormatDateTime('dd/mm/yyyy hh:nn:ss.zzz', Now)]);
    InitMoveProgressForm (nil, 'Traitement en cours.', 'Récupération du catalogue fournisseur.', SizeFic, False, True) ;
    try
      StartOn               := DateTimeToStr(Now);
      InsertGCALines        := 0;
      InsertGALines         := 0;
      InsertBFTLines        := 0;
      InsertBSFLines        := 0;
      InsertGNELines        := 0;
      InsertGNLLines        := 0;
      InsertBCBLines        := 0;
      CreateFamily          := (chkFamSsFam.Enabled and chkFamSsFam.Checked);
      FamilyExist           := False;
      SubFamillyExist       := False;
      ReferenceGCA          := '';
      ReferenceGA           := '';
      FamillyTarif          := '';
      SubFamillyTarif       := '';
      TslFile               := TStringList.Create;
      TslCacheT             := TstringList.Create;
      TslCacheDuplicate     := TStringList.Create;
      TSLCacheGCA           := TStringList.Create;
      TSLCacheBFT           := TStringList.Create;
      TSLCacheBSF           := TStringList.Create;
      TSLCacheGNE           := TStringList.Create;
      TSLCacheGNL           := TStringList.Create;
      TSLCacheBCB           := TStringList.Create;
      try
        TSLCacheGCA.Sorted       := True;
        TSLCacheBFT.Sorted       := True;
        TSLCacheBSF.Sorted       := True;
        TslCacheT.Sorted         := True;
        TslCacheDuplicate.Sorted := True;
        LoadCache;
        Cnx := TADOConnection.Create(application);
        try
          Cnx.ConnectionString := DBSOC.ConnectionString;
          Cnx.LoginPrompt      := false;
          Cnx.Connected        := True;
          if (CBCreationCatalogue.checked) or (CBCreationArticle.checked) then
          begin
            TslFile.Sorted               := True;
            TslInsertTablet.Sorted       := True;
            TslUpdateTablet.Sorted       := True;
            TslGAValuesToUpgraded.Sorted := True;
            CreateGA2                    := (TslPARFOULIG.IndexOfName('GA2_SOUSFAMTARART') > -1);
            CreateBCB                    := (TslPARFOULIG.IndexOfName('BCB_CODEBARRE') > -1);
            GAParametersPrices[0]        := Tools.iif(TslPARFOULIG.IndexOfName('GA_DPR')   > -1, 'X', '-');
            GAParametersPrices[1]        := Tools.iif(TslPARFOULIG.IndexOfName('GA_PAHT')  > -1, 'X', '-');
            GAParametersPrices[2]        := Tools.iif(TslPARFOULIG.IndexOfName('GA_PVHT' ) > -1, 'X', '-');
            GAParametersPrices[3]        := Tools.iif(TslPARFOULIG.IndexOfName('GA_PVTTC') > -1, 'X', '-');
            NeedValorization             := SetNeedValorization;
            TslFile.LoadFromFile(GRF_FICHIER.Text);
            for CptE := 0 to pred(TslFile.count) do
            begin
              MoveCurProgressForm(Format('Débuté le %s.%sLecture du fichier en cours (%s/%s).', [StartOn, #13#10, IntToStr(CptE), IntToSTr(TslFile.count)]));
              stEnreg := TslFile[CptE];
              if SupplierInfile then
              begin
                TheFournisseur := ExtractValue('GCA_TIERS', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
                if TheFournisseur = '' then
                  TheFournisseur := ExtractValue('GA_FOURNPRINC', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
                lIndexOf := Tools.GetIndexOnSortedTsl(TslCacheT, TheFournisseur);
                if lIndexOf = -1 then
                begin
                  TstValue := Tools.iif(SupplierIsValid(False), 'X', '-');
                  TslCacheT.Add(Format('%s=%s', [TheFournisseur, TstValue]));
                end else
                  TstValue := copy(TslCacheT[lIndexOf], pos('=', TslCacheT[lIndexOf]) + 1, 1);
              end else
                TstValue := 'X';
              CanContinue := (TstValue = 'X');
              if not CanContinue then
              begin
                AddWarningLineIfNeed;
                ListRecap.Items.Add(Format('- Erreur - Le code fournisseur est manquant - Ligne %s', [stEnreg]));
              end else
              begin
                ReferenceGCA := ExtractValue('GCA_REFERENCE', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
                ReferenceGA := '';
                if CBCreationArticle.Checked then
                begin
                  ReferenceGA  := ExtractValue('GA_ARTICLE', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
                  if ReferenceGA = '' then
                    ReferenceGA  := ExtractValue('GA_CODEARTICLE', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
                  if (ReferenceGA = '') and (ReferenceGCA <> '') then
                    ReferenceGA := CodeArticleUnique2(ReferenceGCA, '');
                  if (ReferenceGA <> '') and (copy(ReferenceGA, length(ReferenceGA), 1) <> 'X') then
                    ReferenceGA := CodeArticleUnique2(ReferenceGA, '');
                end;
                CanContinue := ((ReferenceGCA <> '') or (ReferenceGA <> ''));
                if not CanContinue then
                begin
                  AddWarningLineIfNeed;
                  ListRecap.Items.Add(Format('- Erreur - La référence catalogue ou la référence article est manquante - Ligne "%s"', [stEnreg]));
                end else
                begin
                  IsPrixPose := (ExtractValue('GA_TYPEARTICLE', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode) = 'ARP');
                  Barcode    := Tools.iif(CreateBCB, ExtractValue('BCB_CODEBARRE', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode), '');
                  { Recherche si on a déjà traité cette référence }
                  TstValue := Tools.iif(ReferenceGCA <> '', ReferenceGCA, ReferenceGA);
                  lIndexOf := Tools.GetIndexOnSortedTsl(TSLCacheGCA, TstValue);
                  if lIndexOf > -1 then
                  begin
                    AddWarningLineIfNeed;
                    if Tools.GetIndexOnSortedTsl(TslCacheDuplicate, TstValue) = -1 then
                    begin
                      TslCacheDuplicate.Add(Format('%s=X', [TstValue]));
                      ListRecap.Items.Add(Format('- Référence en doublon non traitée : %s', [TstValue]))
                    end;
                  end else
                  begin
                    TSLCacheGCA.Add(Format('%s=X', [TstValue]));
                    if CreateFamily then
                    begin
                      FamillyTarif    := ExtractValue('GA_TARIFARTICLE', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
                      SubFamillyTarif := ExtractValue('GA2_SOUSFAMTARART', '', stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode);
                    end;
                    SqlExist := GetExistSql(ReferenceGCA, ReferenceGA);
                    QryExist := OpenSql(SqlExist, True);
                    try
                      CataloguExist := Tools.iif(not QryExist.EOF, (QryExist.Fields[0].AsString <> ''), False);
                      ItemExist     := Tools.iif(not QryExist.EOF, (QryExist.Fields[1].AsString <> ''), False);
                      if CreateFamily then
                      begin
                        FamilyExist     := Tools.iif(not QryExist.EOF, (QryExist.Fields[2].AsString <> ''), False);
                        SubFamillyExist := Tools.iif(not QryExist.EOF, (QryExist.Fields[3].AsString <> ''), False);
                        if Famillytarif = '' then
                          FamillyTarif := Tools.iif(not QryExist.EOF, QryExist.Fields[4].AsString, '');
                        if SubFamillyTarif = '' then
                          SubFamillyTarif := Tools.iif(not QryExist.EOF, QryExist.Fields[5].AsString, '');
                      end;
                      if not IsPrixPose then
                        IsPrixPose := Tools.iif(not QryExist.EOF, (QryExist.Fields[6].AsString = 'ARP'), False);
                      if CreateBCB then
                      begin
                        BarCodeExist      := Tools.iif(not QryExist.EOF, (QryExist.Fields[7].AsString <> ''), False);
                        ExistBCBPrincipal := Tools.iif(not QryExist.EOF, (QryExist.Fields[8].AsString = 'X'), False);
                      end else
                      begin
                        BarCodeExist      := False;
                        ExistBCBPrincipal := False;
                      end;
                    finally
                      Ferme(QryExist);
                    end;
                    { Gestion du catalogue }
                    if (CBCreationCatalogue.checked) and (ReferenceGCA <> '') then
                    begin
                      if not CataloguExist then
                      begin
                        inc(InsertGCALines);
                        ExecuteInsertIfNeed(PrefixGCA, InsertGCALines, Cnx);
                        SetSpExec(TslInsertGCA, True, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixGCA);
                      end else
                      begin
                        inc(UpdateGCALines);
                        ExecuteUpdateIfNeed(PrefixGCA, UpdateGCALines, Cnx);
                        SetSpExec(TslUpdateGCA, False, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixGCA);
                      end;
                    end;
                    { Gestion des articles }
                    if (CBCreationArticle.checked) and (ReferenceGA <> '') then
                    begin
                      if not ItemExist then
                      begin
                        inc(InsertGALines);
                        ExecuteInsertIfNeed(PrefixGA, InsertGALines, Cnx);
                        SetSpExec(TslInsertGA , True, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixGA);
                        if CreateGA2 then
                          SetSpExec(TslInsertGA2, True, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixGA2);
                        if IsPrixPose then
                        begin
                          SetSpExec(TslInsertGNE, True, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixGNE);
                          SetSpExec(TslInsertGNL, True, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixGNL);
                        end;
                      end else
                      begin
                        inc(UpdateGALines);
                        ExecuteUpdateIfNeed(PrefixGA, UpdateGALines, Cnx);
                        SetSpExec(TslUpdateGA , False, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixGA);
                        if CreateGA2 then
                          SetSpExec(TslUpdateGA2, False, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixGA2);
                      end;
                      { Gestion des familles et sous-familles }
                      if (CreateFamily) and (FamillyTarif <> '') then
                      begin
                        { Si n'a pas trouvé les familles et sous-famille dans la table, recherche dans le cache }
                        if not FamilyExist then
                        begin
                          lIndexOf    := TSLCacheBFT.IndexOfName(Format('%s', [FamillyTarif]));
                          FamilyExist := (lIndexOf > -1);
                          if not FamilyExist then
                            TSLCacheBFT.Add(Format('%s=X', [FamillyTarif]));
                        end;
                        if not SubFamillyExist then
                        begin
                          lIndexOf        := TSLCacheBSF.IndexOfName(Format('%s_%s', [FamillyTarif, SubFamillyTarif]));
                          SubFamillyExist := (lIndexOf > -1);
                          if not SubFamillyExist then
                            TSLCacheBSF.Add(Format('%s_%s=X', [FamillyTarif, SubFamillyTarif]));
                        end;
                        if not FamilyExist then
                        begin
                          inc(InsertBFTLines);
                          ExecuteInsertIfNeed(PrefixBFT, InsertBFTLines, Cnx);
                          SetSpExec(TslInsertBFT , True, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixBFT);
                        end;
                        if (not SubFamillyExist) and (SubFamillyTarif <> '') then
                        begin
                          inc(InsertBSFLines);
                          ExecuteInsertIfNeed(PrefixBSF, InsertBSFLines, Cnx);
                          SetSpExec(TslInsertBSF , True, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixBSF);
                        end;
                      end;
                      { Gestion des codes barres }
                      if (CreateBCB) and (Barcode <> '') and (not BarCodeExist) then
                      begin
                        inc(InsertBCBLines);
                        ExecuteInsertIfNeed(PrefixBCB, InsertBCBLines, Cnx);
                        SetSpExec(TslInsertBCB , True, stEnreg, ReferenceGCA, ReferenceGA, FamillyTarif, SubFamillyTarif, Barcode, PrefixBCB);
                      end;
                    end;
                  end;
                  ReferenceGCA := '';
                  ReferenceGA  := '';
                end;
              end;
            end;
            { Il reste des données à enregistrer }
            if  (  TslInsertGCA.count    + TslUpdateGCA.count
                 + TslInsertGA.count     + TslUpdateGA.count
                 + TslInsertGA2.count    + TslUpdateGA2.count
                 + TslInsertBFT.count    + TslInsertBSF.count
                 + TslInsertGNE.count    + TslInsertGNL.count
                 + TslInsertTablet.Count + TslUpdateTablet.Count
                 + TslInsertBCB.Count
                ) > 0
            then
            begin
              if CBCreationCatalogue.checked then
              begin
                InsertUpdateInDB(Cnx, PrefixGCA, True , TslInsertGCA);
                InsertUpdateInDB(Cnx, PrefixGCA, False, TslUpdateGCA);
              end;
              if CBCreationArticle.checked then
              begin
                InsertUpdateInDB(Cnx, PrefixGA , True , TslInsertGA);
                InsertUpdateInDB(Cnx, PrefixGA , False, TslUpdateGA);
                if CreateGA2 then
                begin
                  InsertUpdateInDB(Cnx, PrefixGA2, True , TslInsertGA2);
                  InsertUpdateInDB(Cnx, PrefixGA2, False, TslUpdateGA2);
                  InsertUpdateInDB(Cnx, PrefixGNE, True, TslInsertGNE);
                  InsertUpdateInDB(Cnx, PrefixGNL, True, TslInsertGNL);
                end;
                if CreateFamily then
                begin
                  InsertUpdateInDB(Cnx, PrefixBFT, True , TslInsertBFT);
                  InsertUpdateInDB(Cnx, PrefixBSF, True , TslInsertBSF);
                end;
                if CreateBCB then
                begin
                  InsertUpdateInDB(Cnx, PrefixBCB, True , TslInsertBCB);
                end;
              end;
              if TslInsertTablet.Count + TslUpdateTablet.Count > 0 then
                InsertTablet(Cnx);
            end;
          end;
        finally
          Cnx.Close;
          Cnx.Free;
        end;
      finally
        //
        TSLCacheBCB.clear;
        TSLCacheGNL.clear;
        TSLCacheGNE.clear;
        TSLCacheBSF.clear;
        TSLCacheBFT.clear;
        TSLCacheGCA.clear;
        //
        FreeAndNil(TslFile);
        FreeAndNil(TslCacheT);
        FreeAndNil(TslCacheDuplicate);
        FreeAndNil(TSLCacheGCA);
        FreeAndNil(TSLCacheBFT);
        FreeAndNil(TSLCacheBSF);
        FreeAndNil(TSLCacheGNE);
        FreeAndNil(TSLCacheGNL);
        FreeAndNil(TSLCacheBCB);
      end;
    finally
      if InsertGCALines + UpdateGCALines + InsertGALines + UpdateGALines + InsertBFTLines + InsertBSFLines + InsertBCBLines > 0 then
      begin
        if InsertGCALines + UpdateGCALines > 0 then
        begin
          ListRecap.Items.Add('Table CATALOGU :');
          if InsertGCALines > 0 then ListRecap.Items.Add(Format('- Création de %s lignes '   , [IntToStr(InsertGCALines)]));
          if UpdateGCALines > 0 then ListRecap.Items.Add(Format('- Mise à jour de %s lignes ', [IntToStr(UpdateGCALines)]));
        end;
        if InsertGALines + UpdateGALines > 0 then
        begin
          ListRecap.Items.Add('Table ARTICLE :');
          if InsertGALines > 0 then ListRecap.Items.Add(Format('- Création de %s lignes '   , [IntToStr(InsertGALines)]));
          if UpdateGALines > 0 then ListRecap.Items.Add(Format('- Mise à jour de %s lignes ', [IntToStr(UpdateGALines)]));
        end;
        if InsertBFTLines > 0 then
        begin
          ListRecap.Items.Add('Table BTFAMILLETARIF :');
          ListRecap.Items.Add(Format('- Création de %s lignes ', [IntToStr(InsertBFTLines)]));
        end;
        if InsertBSFLines > 0 then
        begin
          ListRecap.Items.Add('Table BTSOUSFAMILLETARIF :');
          ListRecap.Items.Add(Format('- Création de %s lignes ', [IntToStr(InsertBSFLines)]));
        end;
        if InsertBCBLines > 0 then
        begin
          ListRecap.Items.Add('Table BTCODEBARRE :');
          ListRecap.Items.Add(Format('- Création de %s lignes ', [IntToStr(InsertBCBLines)]));
        end;
      end else
      begin
        ListRecap.Items.Add('Pas de lignes créées ou modifiées :');
        if CBCreationCatalogue.Checked then ListRecap.Items.Add('- table CATALOGU');
        if CBCreationArticle.Checked   then ListRecap.Items.Add('- table ARTICLE');
        if CreateFamily                then ListRecap.Items.Add('- tables BTFAMILLETARIF-BTSOUSFAMILLETARIF');
        if CreateBCB                   then ListRecap.Items.Add('- table BTCODEBARRE');
      end;
      ListRecap.Items.Add('');
      ListRecap.Items.Add(StartJnal);
      ListRecap.Items.Add(Format('Fin de la récupération : %s', [FormatDateTime('dd/mm/yyyy hh:nn:ss.zzz', Now)]));
      ListRecap.Items.Add(Format('Temps de traitement : %s', [FormatDateTime('hh:nn:ss:zz', (Now-TrtStart))]));
      MAJJnalEvent('MTF', 'OK', PTitre.Caption, ListRecap.Items.Text);
      FiniMoveProgressForm ;
    end;
  end;
end;

end.
