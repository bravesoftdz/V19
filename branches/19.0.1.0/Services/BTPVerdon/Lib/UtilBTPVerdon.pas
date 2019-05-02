unit UtilBTPVerdon;

interface

uses
  ConstServices
  , uTob
  , CommonTools
  , AdoDB
  , Classes
  , XmlIntf
  , XMLDoc
  ;

type

  T_TablesName = (  tnNone
                  , tnChantier
                  , tnDevis
                  , tnLignesBR
                  , tnTiers
                  , tnParameters
                  , tnArticles
                  , tnSalaries
                  , tnModeRegle
                  , tnReglement
                  , tnHresSalaries
                  , tnConsoStock
                  , tnIntervenants
                  );

  T_ThirdXmlFileType = (  txft_None
                        , txft_BankAccount_Code
                        , txft_BankAccountType_Code
                        , txft_Bank_Code
                        , txft_SwiftCode
                        , txft_Bic
                       );

  T_InfoTMPType = (tittNone, tittName, tittType);

  T_TiersValues = record
                    FirstExec   : boolean;
                    Count       : integer;
                    TimeOut     : integer;
                    LastSynchro : string;
                    IsActive    : boolean;
                  end;

  T_ChantierValues = record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                       IsActive    : boolean;
                     end;

  T_DevisValues = record
                    FirstExec   : boolean;
                    Count       : integer;
                    TimeOut     : integer;
                    LastSynchro : string;
                    IsActive    : boolean;
                  end;

  T_LignesBRValues = record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                       IsActive    : boolean;
                     end;

  T_IntervenantsValues = record
                           FirstExec   : boolean;
                           Count       : integer;
                           TimeOut     : integer;
                           LastSynchro : string;
                           IsActive    : boolean;
                         end;

  T_ParametresValues = record
                         FirstExec   : boolean;
                         Count       : integer;
                         TimeOut     : integer;
                         LastSynchro : string;    
                         IsActive    : boolean;
                         KeepData    : boolean;
                       end;

  T_SalariesValues =  record
                        FirstExec   : boolean;
                        Count       : integer;
                        TimeOut     : integer;
                        LastSynchro : string;
                        IsActive    : boolean;
                        KeepData    : boolean;
                      end;

  T_ArticlesValues = record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                       IsActive    : boolean;
                       KeepData    : boolean;
                     end;

  T_ModeRegleValues = record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                       IsActive    : boolean;
                       KeepData    : boolean;
                     end;

  T_ReglementValues = record
                       FirstExec   : boolean;
                       Count       : integer;
                       TimeOut     : integer;
                       LastSynchro : string;
                       IsActive    : boolean;
                       KeepData    : boolean;
                     end;

  T_HresSalariesValues = record
                           FirstExec   : boolean;
                           Count       : integer;
                           TimeOut     : integer;
                           LastSynchro : string;
                           IsActive    : boolean;
                           KeepData    : boolean;
                         end;

  T_ConsoStockValues = record
                         FirstExec   : boolean;
                         Count       : integer;
                         TimeOut     : integer;
                         LastSynchro : string;
                         IsActive    : boolean;
                         KeepData    : boolean;
                       end;

  T_FolderValues = record
                     BTPConnectionName : string;
                     BTPUserAdmin      : string;
                     BTPServer         : string;
                     BTPDataBase       : string;
                     TMPServer         : string;
                     TMPDataBase       : string;
                   end;

  TUtilBTPVerdon = class (TObject)
    class function GetTMPTableName(Tn : T_TablesName) : string;
    class function GetMsgStartEnd(Tn : T_TablesName; Start : boolean; LastSynchro : string) : string;
    class function AddLog(ServiceName : string; Tn : T_TablesName; Msg : string; LogValues : T_WSLogValues; LineLevel : integer) : string;
    class procedure AddFieldsTobAdd(Tn : T_TablesName; TobResult : TOB);
    class procedure StartLog(ServiceName : string; lTn : T_TablesName; LogValues : T_WSLogValues; LastSynchro : string);
  end;

  TImportExportTreatment = class (Tobject)
  private
    BTPArrFields           : array of string;
    TMPArrFields           : array of string;
    BTPArrAdditionalFields : array of string;
    TMPArrAdditionalFields : array of string;

    function SetFieldsArray : boolean;
    function GetTMPPrefix : string;
    function GetBTPFieldInf(BTPFieldName : string; InfType : T_InfoTMPType) : string;
    function GetTMPFieldInf(TMPFieldName : string; InfType : T_InfoTMPType) : string;
    function GetTMPIndexFieldName : string;
    function GetBTPIndexFieldName : string;
    function GetBTPLabelFieldName : string;
    function GetSqlDataExist(FieldsList, KeyValue1, KeyValue2 : string) : string;
    function GetSystemFields : string;
    function GetFieldsListFromArray(ArrData : array of string; WithType : boolean) : string;
    function GetValue(FieldNameBTP, FieldNameTMP : string; FieldType : tTypeField; TobData : TOB) : string; overload;
    function GetValue(FieldName, FieldValue : string; BTPFields : boolean=True) : string; overload;
    function GetTMPFieldsList : string;
    function GetSqlInsertValues(TobData : TOB; IsAdditional : boolean=False) : string;
    function GetTMPFieldSizeMax(FieldName : string) : integer;
    function GetSqlUpdate(TobData, TobAdd : TOB; KeyValue1, KeyValue2 : string) : string;
    function GetSqlInsert(TobData, TobAdd : TOB) : string;
    function GetSqlInsertAdditionalFields : string;
    function GetDataSearchSql(LastSynchro : string) : string;
    procedure SetLastSynchro;
    function GetFieldIndex(FieldName : string; FieldsArray : array of string) : integer;
    function IsCorrectValueCompareToType(FieldIndex : integer; Value : string; FieldsArray : array of string) : boolean;

  public
    TreatmentType : string;
    Tn            : T_TablesName;
    LogValues     : T_WSLogValues;
    FolderValues  : T_FolderValues;
    AdoQryBTP     : AdoQry;
    AdoQryTMP     : AdoQry;

    procedure AssignAdoQry(AdoQryBTP, AdoQryTMP : AdoQry; FolderValues : T_FolderValues; LogValues : T_WSLogValues);
  end;

  TExportTreatment = class (TImportExportTreatment)
  private
    function InsertUpdateData(TobData: TOB): boolean;
    procedure SetLinkedRecords(TobAdd, TobData : TOB);

  public
    LastSynchro  : string;

    function ExportTreatment(TobTable, TobAdd, TobQry: TOB): boolean;
  end;

  {$IFDEF APPSRVWITHCBP}
  TImportTreatment = class (TImportExportTreatment)
  private
    TobData            : TOB;
    TobInsertUpdate    : TOB;
    TobPayment         : TOB;
    NumDebugEvent      : integer;
    SocietyCode        : string;
    DBTSouche          : string;
    TslCacheTypeParam  : TStringList;
    TslCacheBTPData    : TStringList;
    TslCacheAffaire    : TStringList;
    TslCacheArticle    : TStringList;

    procedure AddDebugEvent(Text : string);
    procedure SetMandatoryValues(TobDataL : TOB);
    function IsLinkedDocExist(PaymentReferency, DocReferency : string) : boolean;
    function GetCompleteItemCode(ItemCode : string) : string;
    function GetParamTableName(sType : string) : string;
    function GetBTPAffaireValue(Aff1Code, FieldName : string) : string;
    function GetBTPArticleValue(Itemcode, FieldName : string) : string;
    function GetBTPTableName(Tn : T_TablesName; sType : string='') : string;
    function GetBTPWhere(TobDataL : TOB) : string;
    function GetBTPUpdateValues(TobDataL : TOB) : string;
    function GetBTPInsertValues(TobDataL : TOB) : string;
    function GetBTPDefaultValue(FieldName : string) : string;
    function GetBTPFieldsList(sType : string='') : string;
    function IsBTPParamPrefixExist(TobDataL : TOB) : boolean;
    function IsBTPRecordExist(TobDataL : TOB) : boolean;
    function IsOkPaymentData(TobDataL : TOB) : string;
    function GetUpdateXX_Traite(TobDataL : TOB) : string;
    function AddPayment(TobDataL : TOB) : boolean;
    function DoUpdateDocFromPayment : boolean;
    function TableProcess : boolean;

  public
    ParametresValues          : T_ParametresValues;
    SalariesValues            : T_SalariesValues;
    ArticlesValues            : T_ArticlesValues;
    ModeRegleValues           : T_ModeRegleValues;
    ReglementValues           : T_ReglementValues;
    HresSalariesValues        : T_HresSalariesValues;
    ConsoStockValues          : T_ConsoStockValues;
    LastSynchro               : string;
    ParametresFOFieldsList    : string;
    ParametresCCFieldsList    : string;
    ParametresYXFieldsList    : string;
    ArticlesFieldsList        : string;
    SalariesFieldsList        : string;
    ModeRegleFieldsList       : string;
    ReglementFieldsList       : string;
    HresSalariesFieldsList    : string;
    ConsoStockFieldsList      : string;
    TslParametersTypeMatching : TStringList;

    constructor create;
    destructor destroy; override;
    function ImportTreatment : boolean;

  end;
  {$ENDIF APPSRVWITHCBP}

  {$IFNDEF APPSRV}
  TExportEcr = class (TObject)
  private
    MsgCaption         : string;
    FolderCode         : string;
    CustomerCode       : string;
    PathXMLFileVTE     : string;
    PathXMLFileACH     : string;
    PathXMLFileCLI     : string;
    PathXMLFileFOU     : string;
    TempPathXMLFileVTE : string;
    TempPathXMLFileACH : string;
    TempPathXMLFileCLI : string;
    TempPathXMLFileFOU : string;
    TobEcrVen          : TOB;
    TobEcrAch          : TOB;
    TobEcrP            : TOB;
    TobCustomer        : TOB;
    TobProvider        : TOB;
    Tobcountry         : TOB;
    LineNumber         : integer;
    TslData            : TStringList;

    function GetCountryIso2(Country : string) : string;
    function LoadEcr : boolean;
    function LoadAnalytic : boolean;
    function LoadTax : boolean;
    function DoubleLineFromTaxCode(TaxCode : string) : boolean;
    function GetAmount(Amount : double; PositiveValue : boolean) : string;
    function AddEntryToFile(GLEntries : IXMLNode) : boolean;
    function AddThirdToFile(Accounts : IXMLNode; TobData : TOB) : boolean;

  public
    StartDate   : TDateTime;
    EndDate     : TDateTime;
    ForceExport : boolean;

    constructor create;
    destructor destroy; override;
    function AccountingExpTreatment : boolean;
  end;
  {$ENDIF APPSRV}

implementation

uses
  SysUtils
  , hEnt1
  , IniFiles
  , hCtrls
  , SvcMgr
  , ParamSoc
  , ed_Tools
  , Windows
  , UWinSystem
  , StrUtils
  , DateUtils
  {$IFNDEF DBXPRESS}
  , dbTables
  {$ELSE !DBXPRESS}
  , uDbxDataSet
  {$ENDIF !DBXPRESS}
  {$IFNDEF APPSRV}
  , hMsgBox
  , Controls
  , FactCpta
  , EntGC
  , UtilGC
  {$ENDIF APPSRV}
  ;

const
  DBSynchroName          = 'GLO_LSE_SYNCHRONISATION';
  LockDefaultValue       = '1';
  TraiteDefaultValue     = '0';
  DateTraiteDefaultValue = '19000101';
  PrefixCache_GAC_REF    = 'GAC_REF_'; // Préfixe du cache du règlement
  PrefixCache_GAC_JAL    = 'GAC_JAL_'; // Préfixe du cache du journal du règlement
  PrefixCache_GAC_MP     = 'GAC_MP_';  // Préfixe du cache du mode de paiement règlement
  PrefixCache_GAC_GP     = 'GAC_GP_';  // Préfixe du cache de la pièce du règlement
  PrefixCache_GAC_MNO    = 'GAC_MNO_'; // Préfixe du cache du numordre maximum du règlement (GAC_NUMORDRE)
  SocietyCodeFieldName   = 'CODESOCIETE';

{ TUtilBTPVerdon }

class function TUtilBTPVerdon.GetTMPTableName(Tn : T_TablesName): string;
begin
  case Tn of
    tnChantier     : Result := 'CHANTIER';
    tnDevis        : Result := 'DEVIS';
    tnLignesBR     : Result := 'LIGNESBR';
    tnTiers        : Result := 'TIERS';
    tnParameters   : Result := 'PARAMETRES';
    tnArticles     : Result := 'ARTICLES';
    tnSalaries     : Result := 'SALARIES';
    tnModeRegle    : Result := 'MODEREGLE';
    tnReglement    : Result := 'ECRREGLE';
    tnHresSalaries : Result := 'HEURESSALARIES';
    tnConsoStock   : Result := 'CONSOSTOCK';
    tnIntervenants : Result := 'INTERVENANTS';
  else
    Result := '';
  end;
end;

class function TUtilBTPVerdon.GetMsgStartEnd(Tn : T_TablesName; Start : boolean; LastSynchro : string) : string;
begin
  Result := Format('%s de traitement de la table %s (données créées ou modifiées depuis le %s).', [Tools.iif(Start, 'Début', 'Fin'), TUtilBTPVerdon.GetTMPTableName(Tn), LastSynchro]);
end;

class function TUtilBTPVerdon.AddLog(ServiceName : string; Tn : T_TablesName; Msg : string; LogValues : T_WSLogValues; LineLevel : integer) : string;
begin
  TServicesLog.WriteLog(ssbylLog, Msg, ServiceName, LogValues, LineLevel, True, TUtilBTPVerdon.GetTMPTableName(Tn));
end;

class procedure TUtilBTPVerdon.AddFieldsTobAdd(Tn : T_TablesName; TobResult : TOB);
begin
  case Tn of
    tnChantier :
      begin
        TobResult.AddChampSupValeur('LADR_LIBELLE'    , '');
        TobResult.AddChampSupValeur('LADR_ADRESSE1'   , '');
        TobResult.AddChampSupValeur('LADR_ADRESSE2'   , '');
        TobResult.AddChampSupValeur('LADR_ADRESSE3'   , '');
        TobResult.AddChampSupValeur('LADR_CODEPOSTAL' , '');
        TobResult.AddChampSupValeur('LADR_VILLE'      , '');
        TobResult.AddChampSupValeur('LADR_PAYS'       , '');
        TobResult.AddChampSupValeur('LADR_TYPEADRESSE', 'INT');
        TobResult.AddChampSupValeur('FADR_LIBELLE'    , '');
        TobResult.AddChampSupValeur('FADR_ADRESSE1'   , '');
        TobResult.AddChampSupValeur('FADR_ADRESSE2'   , '');
        TobResult.AddChampSupValeur('FADR_ADRESSE3'   , '');
        TobResult.AddChampSupValeur('FADR_CODEPOSTAL' , '');
        TobResult.AddChampSupValeur('FADR_VILLE'      , '');
        TobResult.AddChampSupValeur('FADR_PAYS'       , '');
        TobResult.AddChampSupValeur('FADR_TYPEADRESSE', 'AFA');
        TobResult.AddChampSupValeur('LISTEDEVIS'      , '');
        TobResult.AddChampSupValeur(SocietyCodeFieldName, GetParamSocSecur('SO_SOCIETE', ''));
      end;
  end;
end;

class procedure TUtilBTPVerdon.StartLog(ServiceName : string; lTn : T_TablesName; LogValues : T_WSLogValues; LastSynchro : string);
begin
  TUtilBTPVerdon.AddLog(ServiceName, lTn, '', LogValues, 0);
  TUtilBTPVerdon.AddLog(ServiceName, lTn, DupeString('*', 50), LogValues, 0);
  TUtilBTPVerdon.AddLog(ServiceName, lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, True, LastSynchro), LogValues, 0);
end;

{ TImportExportTreatment }

procedure TImportExportTreatment.AssignAdoQry(AdoQryBTP, AdoQryTMP : AdoQry; FolderValues : T_FolderValues; LogValues : T_WSLogValues);
begin
  AdoQryBTP.ServerName           := FolderValues.BTPServer;
  AdoQryBTP.DBName               := FolderValues.BTPDataBase;
  AdoQryBTP.PgiDB                := 'X';
  AdoQryBTP.Qry.ConnectionString := AdoQryBTP.GetConnectionString;
  AdoQryBTP.LogValues            := LogValues;
  AdoQryTMP.ServerName           := FolderValues.TMPServer;
  AdoQryTMP.DBName               := FolderValues.TMPDataBase;
  AdoQryTMP.PgiDB                := '-';
  AdoQryTMP.Qry.ConnectionString := AdoQryTMP.GetConnectionString;
  AdoQryTMP.LogValues            := LogValues;
end;

function TImportExportTreatment.SetFieldsArray : boolean;
var
  Cpt    : integer;
  ArrLen : integer;

  function GetTMPFieldType(FieldName, BTPFieldType : string) : string;
  begin
    case Tools.CaseFromString(FieldName, [  'LBR_CODEARTICLE' , 'CHA_CODE'         , 'CHA_CODECLIENT'  , 'CHA_RESPONSABLE', 'DEV_CODECHA'
                                          , 'DEV_CODECLIENT'  , 'DEV_CHARGEAFFAIRE', 'DEV_RESPONSABLE' , 'LBR_FOURNISSEUR', 'LBR_CODECHA'
                                          , 'LBR_CODEARTICLE' , 'TIE_CODETIERS'    , 'TIE_TIERSFACTURE', 'TIE_TIERSPAYEUR', 'ART_CODEARTICLE'
                                          , 'INT_CODECHA'     , 'INT_INTERVENANT'
                                         ]) of
      {0 -> 13}          0..14 : Result := 'INTEGER';
    else
      if (pos('_LOCK', FieldName) > 0) or (pos('_TRAITE', FieldName) > 0) then
        Result := 'BOOLEAN'
      else if (pos('_DATECREATION', FieldName) > 0) or (pos('_DATEMODIF', FieldName) > 0) or (pos('_DATETRAITE', FieldName) > 0) then
        Result := 'DATE'
      else if (pos('_CODESOCIETE', FieldName) > 0) then
        Result := 'VARCHAR(3)'
      else
        Result := BTPFieldType;
    end;
  end;

  procedure AddValues(BTPFieldName, TMPFieldName : string; IsAdditional : boolean=False);
  var
    FieldType : string;
  begin
    if not IsAdditional then
    begin
      if BTPFieldName <> '' then
      begin
        if BTPFieldName <> SocietyCodeFieldName then
        begin
          AdoQryBTP.FieldsList := 'DH_TYPECHAMP';
          AdoQryBTP.Request    := 'SELECT ' + AdoQryBTP.FieldsList + ' FROM DECHAMPS WHERE DH_NOMCHAMP =''' + BTPFieldName + '''';
          AdoQryBTP.SingleTableSelect;
          FieldType := AdoQryBTP.TSLResult[0];
          AdoQryBTP.Reset;
        end else
          FieldType := 'COMBO'; 
      end;
    end else
      FieldType := 'VARCHAR(100)';
    if not IsAdditional then
    begin
      if BTPFieldName <> '' then
        BTPArrFields[Cpt] := Format('%s;%s', [BTPFieldName, FieldType]);
      TMPArrFields[Cpt] := Format('%s;%s', [TMPFieldName, GetTMPFieldType(TMPFieldName, FieldType)]);
    end else
    begin
      BTPArrAdditionalFields[Cpt] := Format('%s;%s', [BTPFieldName, FieldType]);
      TMPArrAdditionalFields[Cpt] := Format('%s;%s', [TMPFieldName, FieldType]);
    end;
  end;

begin
  Result := True;
  case Tn of
    tnChantier     : ArrLen := 11;
    tnDevis        : ArrLen := 12;
    tnLignesBR     : ArrLen := 12;
    tnTiers        : ArrLen := 25;
    tnParameters   : ArrLen := 11;
    tnSalaries     : ArrLen := 15;
    tnModeRegle    : ArrLen := 10;
    tnArticles     : ArrLen := 17;
    tnReglement    : ArrLen := 20;
    tnHresSalaries : ArrLen := 21;
    tnConsoStock   : ArrLen := 15;
    tnIntervenants : ArrLen := 8;
  else
    Arrlen := 0;
  end;
  SetLength(BTPArrFields, ArrLen);
  SetLength(TMPArrFields, ArrLen);
  case Tn of
    tnChantier :
      begin
        for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
        begin
          case Cpt of
            0  : AddValues('AFF_AFFAIRE'      , 'CHA_CODECOMPLET');
            1  : AddValues('AFF_AFFAIRE1'     , 'CHA_CODE');
            2  : AddValues('AFF_LIBELLE'      , 'CHA_LIBELLE');
            3  : AddValues('AFF_DESCRIPTIF'   , 'CHA_BLOCNOTE');
            4  : AddValues('AFF_TIERS'        , 'CHA_CODECLIENT');
            5  : AddValues('AFF_DATEDEBUT'    , 'CHA_DATEDEBUT');
            6  : AddValues('AFF_DATEFIN'      , 'CHA_DATEFIN');
            7  : AddValues('AFF_STATUTAFFAIRE', 'CHA_STATUT');
            8  : AddValues('AFF_RESPONSABLE'  , 'CHA_RESPONSABLE');
            9  : AddValues('AFF_DATECREATION' , 'CHA_DATECREATION');
            10 : AddValues('AFF_DATEMODIF'    , 'CHA_DATEMODIF');
          end;
        end;
        ArrLen := 16;
        SetLength(BTPArrAdditionalFields, ArrLen);
        SetLength(TMPArrAdditionalFields, ArrLen);
        for Cpt := Low(BTPArrAdditionalFields) to High(BTPArrAdditionalFields) do
        begin
          case Cpt of
            0  : AddValues('LADR_LIBELLE'      , 'CHA_ADLLIBELLE' , True);
            1  : AddValues('LADR_ADRESSE1'     , 'CHA_ADLADRESSE1', True);
            2  : AddValues('LADR_ADRESSE2'     , 'CHA_ADLADRESSE2', True);
            3  : AddValues('LADR_ADRESSE3'     , 'CHA_ADLADRESSE3', True);
            4  : AddValues('LADR_CODEPOSTAL'   , 'CHA_ADLCP'      , True);
            5  : AddValues('LADR_VILLE'        , 'CHA_ADLVILLE'   , True);
            6  : AddValues('LADR_PAYS'         , 'CHA_ADLPAY'     , True);
            7  : AddValues('FADR_LIBELLE'      , 'CHA_ADFLIBELLE' , True);
            8  : AddValues('FADR_ADRESSE1'     , 'CHA_ADFADRESSE1', True);
            9  : AddValues('FADR_ADRESSE2'     , 'CHA_ADFADRESSE2', True);
            10 : AddValues('FADR_ADRESSE3'     , 'CHA_ADFADRESSE3', True);
            11 : AddValues('FADR_CODEPOSTAL'   , 'CHA_ADFCP'      , True);
            12 : AddValues('FADR_VILLE'        , 'CHA_ADFVILLE'   , True);
            13 : AddValues('FADR_PAYS'         , 'CHA_ADFPAY'     , True);
            14 : AddValues('LISTEDEVIS'        , 'CHA_LISTEDEVIS' , True);
            15 : AddValues(SocietyCodeFieldName, 'CHA_CODESOCIETE', True);
          end;
        end;
      end;
    tnDevis :
      begin
        for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
        begin
          case Cpt of
            0  : AddValues('GP_NUMERO'       , 'DEV_NUMDEVIS');
            1  : AddValues('GP_DATEPIECE'    , 'DEV_DATEDEVIS');
            2  : AddValues('GP_AFFAIRE1'     , 'DEV_CODECHA');
            3  : AddValues('GP_TIERS'        , 'DEV_CODECLIENT');
            4  : AddValues('GP_REFINTERNE'   , 'DEV_LIBELLE');
            5  : AddValues('GP_REPRESENTANT' , 'DEV_RESPONSABLE');
            6  : AddValues('GP_TOTALHT'      , 'DEV_MONTANTHT');
            7  : AddValues('GP_BLOCNOTE'     , 'DEV_BLOCNOTE');
            8  : AddValues('GP_LIBREPIECE1'  , 'DEV_STATUT');
            9  : AddValues('GP_SOCIETE'      , 'DEV_CODESOCIETE');
            10 : AddValues('GP_DATECREATION' , 'DEV_DATECREATION');
            11 : AddValues('GP_DATEMODIF'    , 'DEV_DATEMODIF');
          end;
        end;
      end;
    tnLignesBR :
      begin
        for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
        begin
          case Cpt of
            0  : AddValues('GL_NUMERO'       , 'LBR_NUMBR');
            1  : AddValues('GL_TIERS'        , 'LBR_FOURNISSEUR');
            2  : AddValues('GL_AFFAIRE1'     , 'LBR_CODECHA');
            3  : AddValues('GL_NUMORDRE'     , 'LBR_NUMLIGNE');
            4  : AddValues('GL_CODEARTICLE'  , 'LBR_CODEARTICLE');
            5  : AddValues('GL_LIBELLE'      , 'LBR_LIBELLE');
            6  : AddValues('GL_QTEFACT'      , 'LBR_QUANTITE');
            7  : AddValues('GL_PUHTDEV'      , 'LBR_PU');
            8  : AddValues('GL_UTILISATEUR'  , 'LBR_UTILISATEUR');
            9  : AddValues('GL_SOCIETE'      , 'LBR_CODESOCIETE');
            10 : AddValues('GL_DATECREATION' , 'LBR_DATECREATION');
            11 : AddValues('GL_DATEMODIF'    , 'LBR_DATEMODIF');
          end;
        end;
      end;
    tnTiers :
      begin
        for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
        begin
          case Cpt of
            0  : AddValues('T_AUXILIAIRE'  , 'TIE_CODETIERS');
            1  : AddValues('T_COLLECTIF'   , 'TIE_COLLECTIF');
            2  : AddValues('T_NATUREAUXI'  , 'TIE_NATUREAUXI');
            3  : AddValues('T_LIBELLE'     , 'TIE_LIBELLE');
            4  : AddValues('T_DEVISE'      , 'TIE_DEVISE');
            5  : AddValues('T_ADRESSE1'    , 'TIE_ADRESSE1');
            6  : AddValues('T_ADRESSE2'    , 'TIE_ADRESSE2');
            7  : AddValues('T_ADRESSE3'    , 'TIE_ADRESSE3');
            8  : AddValues('T_CODEPOSTAL'  , 'TIE_CP');
            9  : AddValues('T_VILLE'       , 'TIE_VILLE');
            10 : AddValues('T_PAYS'        , 'TIE_PAYS');
            11 : AddValues('T_TELEPHONE'   , 'TIE_TELEPHONE');
            12 : AddValues('T_FAX'         , 'TIE_TELEPHONE2');
            13 : AddValues('T_TELEX'       , 'TIE_TELEPHONE3');
            14 : AddValues('T_EMAIL'       , 'TIE_EMAIL');
            15 : AddValues('T_RVA'         , 'TIE_WEBURL');
            16 : AddValues('T_COMPTATIERS' , 'TIE_COMPTA');
            17 : AddValues('T_FACTURE'     , 'TIE_TIERSFACTURE');
            18 : AddValues('T_PAYEUR'      , 'TIE_TIERSPAYEUR');
            19 : AddValues('T_BLOCNOTE'    , 'TIE_BLOCNOTE');
            20 : AddValues('T_SIRET'       , 'TIE_SIRET');
            21 : AddValues('T_NIF'         , 'TIE_NUMINTRACOMM');
            22 : AddValues('T_SOCIETE'     , 'TIE_CODESOCIETE');
            23 : AddValues('T_DATEMODIF'   , 'TIE_DATEMODIF');
            24 : AddValues('T_DATECREATION', 'TIE_DATECREATION');
          end;
        end;
      end;
    tnParameters :
      for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
      begin
        case Cpt of
          0  : AddValues('CC_TYPE'   , 'PAR_TYPE');
          1  : AddValues('CC_CODE'   , 'PAR_CODE');
          2  : AddValues('CC_LIBELLE', 'PAR_LIBELLE');
          3  : AddValues('CC_ABREGE' , 'PAR_INFOCOMPL');
          4  : AddValues(''          , 'PAR_CODESOCIETE');
          5  : AddValues(''          , 'PAR_DATECREATION');
          6  : AddValues(''          , 'PAR_DATEMODIF');
          7  : AddValues(''          , 'PAR_LOCK');
          8  : AddValues(''          , 'PAR_TRAITE');
          9  : AddValues(''          , 'PAR_DATETRAITE');
          10 : AddValues(''          , 'PAR_ID');
        end;
      end;
    tnArticles :
      for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
      begin
        case Cpt of
          0  : AddValues('GA_ARTICLE'      , 'ART_CODEARTICLE');
          1  : AddValues('GA_NATUREPRES'   , 'ART_NATUREPRESTATION');
          2  : AddValues('GA_LIBELLE'      , 'ART_LIBELLE');
          3  : AddValues('GA_BLOCNOTE'     , 'ART_BLOCNOTE');
          4  : AddValues('GA_COMPTAARTICLE', 'ART_COMPTA');
          5  : AddValues('GA_FAMILLENIV1'  , 'ART_FAMILLENIV1');
          6  : AddValues('GA_FAMILLENIV2'  , 'ART_FAMILLENIV2');
          7  : AddValues('GA_FAMILLENIV3'  , 'ART_FAMILLENIV3');
          8  : AddValues('GA_LIBREART1'    , 'ART_MARQUE');
          9  : AddValues('GA_FERME'        , 'ART_ACTIF');
          10 : AddValues('GA_SOCIETE'      , 'ART_CODESOCIETE');
          11 : AddValues(''                , 'ART_DATECREATION');
          12 : AddValues(''                , 'ART_DATEMODIF');
          13 : AddValues(''                , 'ART_LOCK');
          14 : AddValues(''                , 'ART_TRAITE');
          15 : AddValues(''                , 'ART_DATETRAITE');
          16 : AddValues(''                , 'ART_ID');
        end;
      end;
    tnSalaries :
      for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
      begin
        case Cpt of
          0  : AddValues('ARS_RESSOURCE'    , 'SAL_CODESAL');
          1  : AddValues('ARS_TYPERESSOURCE', 'SAL_TYPE');
          2  : AddValues('ARS_LIBELLE'      , 'SAL_NOM');
          3  : AddValues('ARS_LIBELLE2'     , 'SAL_PRENOM');
          4  : AddValues('ARS_FONCTION1'    , 'SAL_FONCTION');
          5  : AddValues('ARS_TAUXREVIENTUN', 'SAL_PR');
          6  : AddValues('ARS_IMMAT'        , 'SAL_MATRICULE');
          7  : AddValues('ARS_LIBRERES1'    , 'SAL_BU');
          8  : AddValues(''                 , 'SAL_CODESOCIETE');
          9  : AddValues(''                 , 'SAL_DATECREATION');
          10 : AddValues(''                 , 'SAL_DATEMODIF');
          11 : AddValues(''                 , 'SAL_LOCK');
          12 : AddValues(''                 , 'SAL_TRAITE');
          13 : AddValues(''                 , 'SAL_DATETRAITE');
          14 : AddValues(''                 , 'SAL_ID');
        end;
      end;
    tnModeRegle :
      for Cpt :=  Low(BTPArrFields) to High(BTPArrFields) do
      begin
        case Cpt of
          0  : AddValues('MP_MODEPAIE' , 'MDR_CODE');
          1  : AddValues('MP_LIBELLE'  , 'MDR_LIBELLE');
          2  : AddValues('MP_CATEGORIE', 'MDR_CATEGORIE');
          3  : AddValues(''            , 'MDR_CODESOCIETE');
          4  : AddValues(''            , 'MDR_DATECREATION');
          5  : AddValues(''            , 'MDR_DATEMODIF');
          6  : AddValues(''            , 'MDR_LOCK');
          7  : AddValues(''            , 'MDR_TRAITE');
          8  : AddValues(''            , 'MDR_DATETRAITE');
          9  : AddValues(''            , 'MDR_ID');
        end;
      end;
    tnReglement :
      for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
      begin
        case Cpt of
          0  : AddValues('GAC_JALECR'      , 'EDR_JOURNAL');
          1  : AddValues('GAC_MONTANT'     , 'EDR_MONTANT');
          2  : AddValues('GAC_MONTANTDEV'  , 'EDR_MONTANTDEV');
          3  : AddValues('GAC_MODEPAIE'    , 'EDR_MODEPAIE');
          4  : AddValues('GAC_LIBELLE'     , 'EDR_LIBELLE');
          5  : AddValues('GAC_NUMCHEQUE'   , 'EDR_NUMCHEQUE');
          6  : AddValues('GAC_REFORIGINE'  , 'EDR_IDENTIFIANT');
          7  : AddValues('GAC_DATEECR'     , 'EDR_DATE');
          8  : AddValues('GAC_DATEECHEANCE', 'EDR_DATEECHEANCE');
          9  : AddValues('GAC_AUXILIAIRE'  , 'EDR_CODETIERS');
          10 : AddValues('GAC_NUMORDRE'    , 'EDR_NUMORDRE');
          11 : AddValues(''                , 'EDR_REFERPLSE');
          12 : AddValues(''                , 'EDR_DEVISE');
          13 : AddValues(''                , 'EDR_CODESOCIETE');
          14 : AddValues(''                , 'EDR_DATECREATION');
          15 : AddValues(''                , 'EDR_DATEMODIF');
          16 : AddValues(''                , 'EDR_LOCK');
          17 : AddValues(''                , 'EDR_TRAITE');
          18 : AddValues(''                , 'EDR_DATETRAITE');
          19 : AddValues(''                , 'EDR_ID');
        end;
      end;
    tnHresSalaries :
      for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
      begin
        case Cpt of
          0  : AddValues(''               , 'HSA_IDENTIFIANT');
          1  : AddValues('BCO_DATEMOUV'   , 'HSA_DATEMVT');
          2  : AddValues('BCO_AFFAIRE1'   , 'HSA_CODECHA');
          3  : AddValues('BCO_NUMERO'     , 'HSA_NUMDEVIS');
          4  : AddValues('BCO_RESSOURCE'  , 'HSA_CODESAL');
          5  : AddValues('BCO_LIBELLE'    , 'HSA_LIBELLE');
          6  : AddValues('BCO_CODEARTICLE', 'HSA_CODEPRESTATION');
          7  : AddValues('BCO_QUANTITE'   , 'HSA_NBHRS');
          8  : AddValues('BCO_DPA'        , 'HSA_DPA');
          9  : AddValues('BCO_DPR'        , 'HSA_DPR');
          10 : AddValues('BCO_PUHT'       , 'HSA_PUV');
          11 : AddValues(''               , 'HSA_BU');
          12 : AddValues(''               , 'HSA_EPURE');
          13 : AddValues(''               , 'HSA_NUMLSE');
          14 : AddValues(''               , 'HSA_CODESOCIETE');
          15 : AddValues(''               , 'HSA_DATECREATION');
          16 : AddValues(''               , 'HSA_DATEMODIF');
          17 : AddValues(''               , 'HSA_LOCK');
          18 : AddValues(''               , 'HSA_TRAITE');
          19 : AddValues(''               , 'HSA_DATETRAITE');
          20 : AddValues(''               , 'HSA_ID');
        end;
      end;
    tnConsoStock :
      for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
      begin
        case Cpt of
          0  : AddValues('BCO_DATEMOUV'   , 'COS_DATEMVT');
          1  : AddValues('BCO_AFFAIRE1'   , 'COS_CODECHA');
          2  : AddValues('BCO_CODEARTICLE', 'COS_CODEARTICLE');
          3  : AddValues('BCO_LIBELLE'    , 'COS_COMMENTAIRE');
          4  : AddValues('BCO_QUANTITE'   , 'COS_QTE');
          5  : AddValues('BCO_DPA'        , 'COS_DPA');
          6  : AddValues('BCO_DPR'        , 'COS_DPR');
          7  : AddValues('BCO_PUHT'       , 'COS_PUV');
          8  : AddValues(''               , 'COS_CODESOCIETE');
          9  : AddValues(''               , 'COS_DATECREATION');
          10 : AddValues(''               , 'COS_DATEMODIF');
          11 : AddValues(''               , 'COS_LOCK');
          12 : AddValues(''               , 'COS_TRAITE');
          13 : AddValues(''               , 'COS_DATETRAITE');
          14 : AddValues(''               , 'COS_ID');
        end;
      end;
    tnIntervenants :
      for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
      begin
        case Cpt of
          0  : AddValues('AFT_AFFAIRE'       , 'INT_CODECHACOMPLET');
          1  : AddValues('AFF_AFFAIRE1'      , 'INT_CODECHA');
          2  : AddValues('AFT_RESSOURCE'     , 'INT_INTERVENANT');
          3  : AddValues('AFT_LIENAFFTIERS'  , 'INT_ROLE');
          4  : AddValues('ARS_FERME'         , 'INT_ACTIF');
          5  : AddValues(SocietyCodeFieldName, 'INT_CODESOCIETE');
          6  : AddValues('AFT_DATECREATION'  , 'INT_DATECREATION');
          7  : AddValues('AFT_DATEMODIF'     , 'INT_DATEMODIF');
        end;
      end;
  end;
end;

function TImportExportTreatment.GetTMPPrefix : string;
begin
  case Tn of
    tnChantier     : Result := 'CHA';
    tnDevis        : Result := 'DEV';
    tnLignesBR     : Result := 'LBR';
    tnTiers        : Result := 'TIE';
    tnParameters   : Result := 'PAR';
    tnArticles     : Result := 'ART';
    tnSalaries     : Result := 'SAL';
    tnModeRegle    : Result := 'MDR';
    tnReglement    : Result := 'EDR';
    tnHresSalaries : Result := 'HSA';
    tnIntervenants : Result := 'INT';
    tnConsoStock   : Result := 'COS'; 
  else
    Result := '';
  end;
end;

function TImportExportTreatment.GetBTPFieldInf(BTPFieldName : string; InfType : T_InfoTMPType) : string;
var
  Cpt : integer;
begin
  if BTPFieldName <> '' then
  begin
    for Cpt := Low(BTPArrFields) to High(BTPArrFields) do
    begin
     if Tools.ExtractFieldName(BTPArrFields[Cpt]) = BTPFieldName then
      begin
        case InfType of
          tittName : Result := Tools.ExtractFieldName(TMPArrFields[Cpt]);
          tittType : Result := Tools.ExtractFieldType(TMPArrFields[Cpt]);
        else
          Result := '';
        end;
       break;
      end;
    end;
  end else
    Result := '';
end;

function TImportExportTreatment.GetTMPFieldInf(TMPFieldName : string; InfType : T_InfoTMPType) : string;
var
  Cpt : integer;
begin
  if TMPFieldName <> '' then
  begin
    for Cpt := Low(TMPArrFields) to High(TMPArrFields) do
    begin
     if Tools.ExtractFieldName(TMPArrFields[Cpt]) = TMPFieldName then
      begin
        case InfType of
          tittName : Result := Tools.ExtractFieldName(TMPArrFields[Cpt]);
          tittType : Result := Tools.ExtractFieldType(TMPArrFields[Cpt]);
        else
          Result := '';
        end;
       break;
      end;
    end;
  end else
    Result := '';
end;


function TImportExportTreatment.GetTMPIndexFieldName : string;
begin
  case Tn of
    tnChantier     : Result := 'CHA_CODECOMPLET';
    tnDevis        : Result := 'DEV_NUMDEVIS';
    tnLignesBR     : Result := 'LBR_NUMBR;LBR_NUMLIGNE';
    tnTiers        : Result := 'TIE_CODETIERS';
    tnParameters   : Result := 'PAR_TYPE;PAR_CODE';
    tnArticles     : Result := 'ART_CODEARTICLE';
    tnSalaries     : Result := 'SAL_CODESAL';
    tnModeRegle    : Result := 'MDR_CODE';
    tnReglement    : Result := 'EDR_IDENTIFIANT';
    tnHresSalaries : Result := 'HSA_IDENTIFIANT';
    tnIntervenants : Result := 'INT_CODECHA;INT_INTERVENANT';
    tnConsoStock   : Result := 'COS_DATEMVT;COS_CODECHA;COS_CODEARTICLE';
  else
    Result := '';
  end;
end;

function TImportExportTreatment.GetBTPIndexFieldName : string;
begin
  case Tn of
    tnChantier     : Result := 'AFF_AFFAIRE';
    tnDevis        : Result := 'GP_NUMERO';
    tnLignesBR     : Result := 'GL_NUMERO;GL_NUMORDRE';
    tnTiers        : Result := 'T_AUXILIAIRE';
    tnParameters   : Result := 'CC_TYPE;CC_CODE';
    tnArticles     : Result := 'GA_ARTICLE';
    tnSalaries     : Result := 'ARS_RESSOURCE';
    tnModeRegle    : Result := 'MP_MODEPAIE';
    tnReglement    : Result := 'GAC_REFORIGINE';
    tnHresSalaries : Result := 'BCO_NUMMOUV';
    tnIntervenants : Result := 'AFF_AFFAIRE1;AFT_RESSOURCE';
    tnConsoStock   : Result := 'BCO_DATEMOUV;BCO_NUMMOUV;BCO_INDICE';
  else
    Result := '';
  end;
end;

function TImportExportTreatment.GetBTPLabelFieldName : string;
begin
  case Tn of
    tnChantier     : Result := 'AFF_LIBELLE';
    tnDevis        : Result := 'GP_REFINTERNE';
    tnLignesBR     : Result := 'GL_LIBELLE';
    tnTiers        : Result := 'T_LIBELLE';
    tnParameters   : Result := 'CC_LIBELLE';
    tnArticles     : Result := 'GA_LIBELLE';
    tnSalaries     : Result := 'ARS_LIBELLE';
    tnModeRegle    : Result := 'MP_LIBELLE';
    tnReglement    : Result := 'GAC_LIBELLE';
    tnHresSalaries : Result := 'BCO_LIBELLE';
    tnIntervenants : Result := 'AFT_RESSOURCE';
    tnConsoStock   : Result := 'BCO_LIBELLE';
  else
    Result := '';
  end;
end;

function TImportExportTreatment.GetSqlDataExist(FieldsList, KeyValue1, KeyValue2 : string) : string;
var
  IndexFieldName : string;
  IndexFieldType : tTypeField;
  Cpt            : integer;
begin
  case Tn of
    tnLignesBR     : Result := Format('SELECT %s FROM LIGNESBR     WHERE LBR_NUMBR   = ''%s'' AND LBR_NUMLIGNE    = %s'  , [FieldsList, KeyValue1, KeyValue2]);
    tnParameters   : Result := Format('SELECT %s FROM CHOIXCOD     WHERE CC_TYPE     = ''%s'' AND CC_CODE         = %s'  , [FieldsList, KeyValue1, KeyValue2]);
    tnIntervenants : Result := Format('SELECT %s FROM INTERVENANTS WHERE INT_CODECHA = %s     AND INT_INTERVENANT = %s'  , [FieldsList, KeyValue1, KeyValue2]);
  else
    begin
      IndexFieldName := GetTMPIndexFieldName;
      IndexFieldType := ttfNone;
      for Cpt := 0 to High(TMPArrFields) do
      begin
        if Tools.ExtractFieldName(TMPArrFields[Cpt]) = IndexFieldName then
        begin
          IndexFieldType := Tools.GetTypeFieldFromStringType(Tools.ExtractFieldType(TMPArrFields[Cpt]));
          break;
        end;
      end;
      case IndexFieldType of
        ttfNumeric : Result := Format('SELECT %s FROM %s WHERE %s = %s'  , [FieldsList, TUtilBTPVerdon.GetTMPTableName(Tn), IndexFieldName, FloatToStr(StrToFloat(KeyValue1))]);
        ttfInt     : Result := Format('SELECT %s FROM %s WHERE %s = %s'  , [FieldsList, TUtilBTPVerdon.GetTMPTableName(Tn), IndexFieldName, IntToStr(StrToInt(KeyValue1))]);
      else ;
        Result := Format('SELECT %s FROM %s WHERE %s = ''%s''', [FieldsList, TUtilBTPVerdon.GetTMPTableName(Tn), IndexFieldName, KeyValue1]);
      end;
    end
  end;
end;

function TImportExportTreatment.GetSystemFields : string;
var
  Prefix : string;
begin
  Prefix := GetTMPPrefix;
  Result := Format('%s_LOCK;%s_TRAITE;%s_DATETRAITE', [Prefix, Prefix, Prefix]);
end;

function TImportExportTreatment.GetFieldsListFromArray(ArrData: array of string; WithType : boolean): string;
var
  Cpt       : integer;
  FieldName : string;
begin
  Result := '';
  for Cpt := Low(ArrData) to High(ArrData) do
  begin
    FieldName := Tools.iif(WithType, Tools.ExtractFieldName(ArrData[Cpt]), ArrData[Cpt]);
    if FieldName <> '' then
    begin
      if FieldName <> SocietyCodeFieldName then
        Result := Format('%s,%s', [Result, FieldName])
      else
        Result := Format('%s,''%s'' AS %s', [Result, GetParamSocSecur('SO_SOCIETE', ''), FieldName])
    end;
  end;
  Result := copy(Result, 2, length(Result));
end;

function TImportExportTreatment.GetValue(FieldNameBTP, FieldNameTMP : string; FieldType : tTypeField; TobData : TOB) : string;
var
  FieldSize : integer;
  Value     : variant;
begin
  Result := '';
  Value  := TobData.GetValue(FieldNameBTP);
  case FieldType of
    ttfNumeric, ttfInt :
      begin
        Result := StringReplace(Value, ',', '.', [rfReplaceAll]);
        if Result = '' then
          Result := '0';
      end;
    ttfDate :
      Result := Format('''%s''', [Tools.UsDateTime_(StrToDateTime(Value))]);
    ttfCombo, ttfText, ttfMemo :
      begin
        Result := Value;
        if Result <> '' then
        begin
          if FieldType = ttfMemo then
            Result := Tools.BlobToString_(Result);
          FieldSize := GetTMPFieldSizeMax(FieldNameTMP);
          Result    := Tools.iif(FieldSize > -1, Trim(Copy(Result, 1, FieldSize)), Result);
          if pos('''', Result) > 0 then
            Result := StringReplace(Result, '''', '''''', [rfReplaceAll]);
        end;
        Result := Format('''%s''', [Result]);
      end;
    ttfBoolean :
      begin
        if pos('_ACTIF', FieldNameTMP) > 0 then // Si champ XX_ACTIF, il faut inverser le sens du boolean (dans BTP, c'est XX_FERME)
          Result := Tools.iif(pos('-', Value) > 0, '1', '0')
        else
          Result := Tools.iif(pos('-', Value) > 0, '0', '1');
      end;
  end;
end;

function TImportExportTreatment.GetValue(FieldName, FieldValue : string; BTPFields : boolean=True) : string;
var
  FieldType : tTypeField;
  FieldSize : integer;
begin
  Result := '';
  if (FieldName <> '') and (FieldValue <> '') then
  begin
    if BTPFields then
      FieldType := Tools.GetTypeFieldFromStringType(GetBTPFieldInf(FieldName, tittType))
    else
      FieldType := Tools.GetTypeFieldFromStringType(GetTMPFieldInf(FieldName, tittType));
    case FieldType of
      ttfNumeric, ttfInt :
        begin
          Result := StringReplace(FieldValue, ',', '.', [rfReplaceAll]);
          if Result = '' then
            Result := '0';
        end;
      ttfDate :
        Result := Format('%s', [Tools.UsDateTime_(StrToDateTime(FieldValue))]);
      ttfCombo, ttfText, ttfMemo :
        begin
          Result := FieldValue;
          if Result <> '' then
          begin
            if FieldType = ttfMemo then
              Result := Tools.BlobToString_(Result);
            FieldSize := GetTMPFieldSizeMax(FieldName);
            Result    := Tools.iif(FieldSize > -1, Trim(Copy(Result, 1, FieldSize)), Result);
            if pos('''', Result) > 0 then
              Result := StringReplace(Result, '''', '''''', [rfReplaceAll]);
          end;
//          Result := Format('''%s''', [Result]);
        end;
      ttfBoolean :
        begin
          if pos('_ACTIF', FieldName) > 0 then // Si champ XX_ACTIF, il faut inverser le sens du boolean (dans BTP, c'est XX_FERME)
            Result := Tools.iif(pos('-', FieldValue) > 0, '1', '0')
          else
            Result := Tools.iif(pos('-', FieldValue) > 0, '0', '1');
        end;
    end;
  end;
end;

function TImportExportTreatment.GetTMPFieldsList : string;
var
  Cpt          : integer;
  SystemFields : string;
begin
  Result := '';
  for Cpt :=  Low(TMPArrFields) to High(TMPArrFields) do
    Result := Format('%s, %s', [Result, Tools.ExtractFieldName(TMPArrFields[Cpt])]);
  { En import, les champs système sont déj  définis }
  if TreatmentType <> 'IMPORT' then
  begin
    SystemFields := GetSystemFields;
    while SystemFields <> '' do
      Result := Format('%s, %s', [Result, Tools.ReadTokenSt_(SystemFields, ';')]);
  end;
  Result := Copy(Result, 2, length(Result));
end;

function TImportExportTreatment.GetSqlInsertValues(TobData: TOB; IsAdditional : boolean=False): string;
var
  Cpt          : integer;
  FieldNameTMP : string;
  FieldNameBTP : string;
  FieldValue   : string;
  FieldType    : tTypeField;
begin
  Result := '';
  if not IsAdditional then
  begin
    for Cpt := 0 to High(BTPArrFields) do
    begin
      FieldNameBTP := Tools.ExtractFieldName(BTPArrFields[Cpt]);
      FieldNameTMP := Tools.ExtractFieldName(TMPArrFields[Cpt]);
      FieldType    := Tools.GetTypeFieldFromStringType(Tools.ExtractFieldType(TMPArrFields[Cpt]));
      FieldValue   := GetValue(FieldNameBTP, FieldNameTMP, FieldType, TobData);
      if IsCorrectValueCompareToType(GetFieldIndex(FieldNameBTP, BTPArrFields), FieldValue, TMPArrFields) then
        Result := Format('%s, %s', [Result, FieldValue])
      else
      begin
        Result := Format('%s;%s;%s;%s', [WSCDS_ErrorMsg, FieldValue, FieldNameTMP, Tools.ExtractFieldType(TMPArrFields[Cpt])]);
        Break;
      end;
    end;
    if pos(WSCDS_ErrorMsg + ';', Result) = 0 then
      Result := Format('%s, ''%s'', ''%s'', ''%s''', [Result, LockDefaultValue, TraiteDefaultValue, DateTraiteDefaultValue]);
  end else
  begin
    for Cpt := 0 to High(BTPArrAdditionalFields) do
    begin
      FieldNameBTP := Tools.ExtractFieldName(BTPArrAdditionalFields[Cpt]);
      FieldNameTMP := Tools.ExtractFieldName(TMPArrAdditionalFields[Cpt]);
      FieldType    := Tools.GetTypeFieldFromStringType(Tools.ExtractFieldType(BTPArrAdditionalFields[Cpt]));
      FieldValue   := GetValue(FieldNameBTP, FieldNameTMP, FieldType, TobData);
      Result       := Format('%s, %s', [Result, FieldValue]);
    end;
  end;
  if pos(WSCDS_ErrorMsg + ';', Result) = 0 then
    Result := Copy(Result, 2, length(Result));
end;

function TImportExportTreatment.GetTMPFieldSizeMax(FieldName : string) : integer;
begin
  case Tools.CaseFromString(FieldName, [  'CHA_CODE'   , 'CHA_BLOCNOTE', 'CHA_ADLCP', 'CHA_ADFCP'
                                        , 'DEV_CODECHA', 'DEV_BLOCNOTE'
                                        , 'LBR_CODECHA'
                                        , 'TIE_CP'     , 'TIE_BLOCNOTE']) of
    {CHA_CODE}     0 : Result := 8;
    {CHA_BLOCNOTE} 1 : Result := 256;
    {CHA_ADLCP}    2 : Result := 5;
    {CHA_ADFCP}    3 : Result := 5;
    {DEV_CODECHA}  4 : Result := 8;
    {DEV_BLOCNOTE} 5 : Result := 256;
    {LBR_CODECHA}  6 : Result := 8;
    {TIE_CP}       7 : Result := 5;
    {TIE_BLOCNOTE} 8 : Result := 256;
  else
    Result := -1;
  end;
end;

function TImportExportTreatment.GetSqlUpdate(TobData, TobAdd : TOB; KeyValue1, KeyValue2 : string) : string;
var
  Cpt          : integer;
  FieldNameTMP : string;
  FieldNameBTP : string;
  FieldValue   : string; 
  Sql          : string;
  Prefix       : string;
  SystemFields : string;
  FieldType    : tTypeField;
  OnError      : boolean;
begin
  OnError := False;
  for Cpt := 0 to High(TMPArrFields) do
  begin
    FieldNameBTP := Tools.ExtractFieldName(BTPArrFields[Cpt]);
    FieldNameTMP := Tools.ExtractFieldName(TMPArrFields[Cpt]);
    FieldType    := Tools.GetTypeFieldFromStringType(Tools.ExtractFieldType(TMPArrFields[Cpt]));
    FieldValue   := GetValue(FieldNameBTP, FieldNameTMP, FieldType, TobData);
    if IsCorrectValueCompareToType(GetFieldIndex(FieldNameBTP, BTPArrFields), FieldValue, TMPArrFields) then
      Sql := Format('%s, %s=%s', [Sql, FieldNameTMP, FieldValue])
    else
    begin
      Result  := Format('%s;%s;%s;%s', [WSCDS_ErrorMsg, FieldValue, FieldNameTMP, Tools.ExtractFieldType(TMPArrFields[Cpt])]);
      OnError := True;
      Break;
    end;
  end;
  if not OnError then
  begin
    case Tn of
      tnChantier :
      begin
        for Cpt :=  Low(BTPArrAdditionalFields) to High(BTPArrAdditionalFields) do
        begin
          FieldNameTMP := Tools.ExtractFieldName(TMPArrAdditionalFields[Cpt]);
          FieldNameBTP := Tools.ExtractFieldName(BTPArrAdditionalFields[Cpt]);
          FieldType    := Tools.GetTypeFieldFromStringType(Tools.ExtractFieldType(TMPArrAdditionalFields[Cpt]));
          Sql          := Format('%s, %s=%s', [Sql, FieldNameTMP, GetValue(FieldNameBTP, FieldNameTMP, FieldType, TobAdd)]); //TobAdd.GetString(FieldNameBTP)]);
        end;
      end;
    end;
    Prefix       := GetTMPPrefix;
    SystemFields := GetSystemFields;
    while SystemFields <> '' do
    begin
      FieldNameTMP := Tools.ReadTokenSt_(SystemFields, ';');
      case Tools.CaseFromString(FieldNameTMP, [Prefix + '_LOCK', Prefix + '_TRAITE', Prefix + '_DATETRAITE']) of
        {LOCK}       0 : Sql := Format('%s, %s=''%s''', [Sql, FieldNameTMP, LockDefaultValue]);
        {TRAITE}     1 : Sql := Format('%s, %s=''%s''', [Sql, FieldNameTMP, TraiteDefaultValue]);
        {DATETRAITE} 2 : Sql := Format('%s, %s=''%s''', [Sql, FieldNameTMP, DateTraiteDefaultValue]);
      end;
    end;
    Sql := Copy(Sql, 2, length(Sql));
    case Tn of
      tnLignesBR     : Result := Format('UPDATE LIGNESBR     SET %s WHERE LBR_NUMBR   = ''%s'' AND LBR_NUMLIGNE    = %s', [Sql, KeyValue1, KeyValue2]);
      tnIntervenants : Result := Format('UPDATE INTERVENANTS SET %s WHERE INT_CODECHA = %s     AND INT_INTERVENANT = %s', [Sql, KeyValue1, KeyValue2]);
    else
      Result := Format('UPDATE %s SET %s WHERE %s = ''%s''', [TUtilBTPVerdon.GetTMPTableName(Tn), Sql, GetTMPIndexFieldName, KeyValue1]);
    end;
  end;
end;

function TImportExportTreatment.GetSqlInsert(TobData, TobAdd : TOB) : string;
var
  Fields    : string;
  Values    : string;
begin
  Fields := GetTMPFieldsList;
  Values := GetSqlInsertValues(TobData);
  if pos(WSCDS_ErrorMsg + ';', Values) = 0 then
  begin
    case Tn of
      tnChantier :
        begin
          Fields := Format('%s, %s', [Fields, GetSqlInsertAdditionalFields]);
          Values := Format('%s, %s', [Values, GetSqlInsertValues(TobAdd, True)]);
        end;
    end;
    Result := Format('INSERT INTO %s (%s) VALUES (%s)', [TUtilBTPVerdon.GetTMPTableName(Tn), Fields, Values]);
  end else
    Result := Values;
end;

function TImportExportTreatment.GetSqlInsertAdditionalFields : string;
var
  Cpt : integer;
begin
  Result := '';
  for Cpt :=  Low(TMPArrAdditionalFields) to High(TMPArrAdditionalFields) do
    Result := Format('%s, %s', [Result, Tools.ExtractFieldName(TMPArrAdditionalFields[Cpt])]);
  Result := Copy(Result, 2, length(Result));
end;

function TImportExportTreatment.GetDataSearchSql(LastSynchro : string) : string;
var
  BTPFieldsList : string;

  function GetLastSynchro : string;
  begin
    Result := Tools.CastDateTimeForQry(StrToDatetime(LastSynchro));
  end;

begin
  BTPFieldsList := GetFieldsListFromArray(BTPArrFields, True);
  case Tn of
    tnTiers        : Result := Format('SELECT %s FROM %s WHERE T_DATEMODIF >= ''%s'' ORDER BY T_AUXILIAIRE'                                           , [BTPFieldsList, Tools.GetTableNameFromTtn(ttnTiers)   , GetLastSynchro]);
    tnDevis        : Result := Format('SELECT %s FROM %s WHERE GP_NATUREPIECEG = ''DBT'' AND GP_DATEMODIF >= ''%s'' ORDER BY GP_NUMERO'               , [BTPFieldsList, Tools.GetTableNameFromTtn(ttnPiece)   , GetLastSynchro]);
    tnLignesBR     : Result := Format('SELECT %s FROM %s WHERE GL_NATUREPIECEG = ''BLF'' AND GL_DATEMODIF >= ''%s'' ORDER BY GL_NUMERO, GL_NUMLIGNE'  , [BTPFieldsList, Tools.GetTableNameFromTtn(ttnLigne)   , GetLastSynchro]);
    tnChantier     : Result := Format('SELECT %s FROM %s WHERE AFF_AFFAIRE0 = ''A'' AND AFF_DATEMODIF >= ''%s'' ORDER BY AFF_AFFAIRE'                 , [BTPFieldsList, Tools.GetTableNameFromTtn(ttnAffaire) , GetLastSynchro]);
    tnIntervenants : Result := Format('SELECT %s FROM %s'
                                    + ' JOIN AFFAIRE   ON AFF_AFFAIRE   = AFT_AFFAIRE'
                                    + ' JOIN RESSOURCE ON ARS_RESSOURCE = AFT_RESSOURCE'
                                    + ' WHERE AFT_DATEMODIF >= ''%s'' ORDER BY AFT_AFFAIRE'
                                      , [BTPFieldsList, Tools.GetTableNameFromTtn(ttnAffTiers), GetLastSynchro]);
    tnParameters
    , tnArticles
    , tnSalaries
    , tnModeRegle
    , tnReglement
    , tnHresSalaries
    , tnConsoStock : Result := Format('SELECT %s FROM %s WHERE %s_LOCK = 0 AND %s_TRAITE = 0', [AdoQryTMP.FieldsList, TUtilBTPVerdon.GetTMPTableName(Tn), GetTMPPrefix, GetTMPPrefix]);
  else
    Result := '';
  end;
end;

procedure TImportExportTreatment.SetLastSynchro;
var
  SettingFile : TInifile;
  IniFilePath : string;
begin
  IniFilePath := Format('%s\%s.%s', [TServicesLog.GetServicesAppDataPath(True), ServiceName_BTPVerdonIniFile, 'ini']);
  SettingFile := TIniFile.Create(IniFilePath);
  try
    SettingFile.WriteString(Format('%s_LASTSYNCHRO', [TreatmentType]), TUtilBTPVerdon.GetTMPTableName(Tn), DateTimeToStr(Now));
    SettingFile.UpdateFile;
  finally
    SettingFile.Free;
  end;
end;

function TImportExportTreatment.GetFieldIndex(FieldName : string; FieldsArray : array of string) : integer;
var
  Cpt : integer;
begin
  Result := -1;
  if FieldName <> '' then
  begin
    for Cpt := 0 to High(FieldsArray) do
    begin
      if Tools.ExtractFieldName(FieldsArray[Cpt]) = FieldName then
      begin
        Result := Cpt;
        break;
      end;
    end;
  end;
end;

function TImportExportTreatment.IsCorrectValueCompareToType(FieldIndex : integer; Value : string; FieldsArray : array of string) : boolean;
var
  FieldName : string;
  TempValue : string; 
  FieldType : tTypeField;
  lInteger  : integer;
  lDouble   : double;
  lBoolean  : boolean;
  lDate     : TDateTime;
begin
  Result := True;
  if FieldIndex > -1 then
  begin
    FieldName := FieldsArray[FieldIndex];
    if FieldName <> '' then
    begin
      FieldType := Tools.GetTypeFieldFromStringType(Tools.ExtractFieldType(FieldName));
      case FieldType of
        ttfInt     : Result := TryStrToInt(Value, lInteger);
        ttfNumeric : begin
                       TempValue := StringReplace(Value, '.', ',', [rfReplaceAll]);
                       Result    := TryStrToFloat(TempValue, lDouble);
                     end;
        ttfDate    : begin
                       TempValue := Tools.DecodeDateTimeFromQry(Value);
                       Result    := TryStrToDateTime(TempValue, lDate);
                     end;
        ttfBoolean : begin
                       TempValue := Tools.iif(pos('-', Value) > 0, '0', '1');
                       Result    := TryStrToBool(TempValue, lBoolean);
                     end;
      else
        Result := True;
      end;
    end;
  end;
end;
{ TExportTreatment }

procedure TExportTreatment.SetLinkedRecords(TobAdd, TobData : TOB);

  procedure ClearValues;
  begin
    TobAdd.SetString('LADR_LIBELLE'    , '');
    TobAdd.SetString('LADR_ADRESSE1'   , '');
    TobAdd.SetString('LADR_ADRESSE2'   , '');
    TobAdd.SetString('LADR_ADRESSE3'   , '');
    TobAdd.SetString('LADR_CODEPOSTAL' , '');
    TobAdd.SetString('LADR_VILLE'      , '');
    TobAdd.SetString('LADR_PAYS'       , '');
    TobAdd.SetString('FADR_LIBELLE'    , '');
    TobAdd.SetString('FADR_ADRESSE1'   , '');
    TobAdd.SetString('FADR_ADRESSE2'   , '');
    TobAdd.SetString('FADR_ADRESSE3'   , '');
    TobAdd.SetString('FADR_CODEPOSTAL' , '');
    TobAdd.SetString('FADR_VILLE'      , '');
    TobAdd.SetString('FADR_PAYS'       , '');
    TobAdd.SetString('LISTEDEVIS'      , '');
  end;

  procedure AddAdress;
  var
    Sql        : string;
    FieldsList : array of string;
    TobAdr     : TOB;
    TobAdrL    : TOB;
    Cpt        : integer;

    procedure AddValues(Prefix : string);
    begin
      TobAdd.SetString(Format('%sADR_LIBELLE'    , [Prefix]), TobAdrL.GetString('ADR_LIBELLE'));
      TobAdd.SetString(Format('%sADR_ADRESSE1'   , [Prefix]), TobAdrL.GetString('ADR_ADRESSE1'));
      TobAdd.SetString(Format('%sADR_ADRESSE2'   , [Prefix]), TobAdrL.GetString('ADR_ADRESSE2'));
      TobAdd.SetString(Format('%sADR_ADRESSE3'   , [Prefix]), TobAdrL.GetString('ADR_ADRESSE3'));
      TobAdd.SetString(Format('%sADR_CODEPOSTAL' , [Prefix]), TobAdrL.GetString('ADR_CODEPOSTAL'));
      TobAdd.SetString(Format('%sADR_VILLE'      , [Prefix]), TobAdrL.GetString('ADR_VILLE'));
      TobAdd.SetString(Format('%sADR_PAYS'       , [Prefix]), TobAdrL.GetString('ADR_PAYS'));
    end;

  begin
    TobAdr := TOB.Create('_ADR', nil, -1);
    try
      SetLength(FieldsList, 8);
      FieldsList[0] := 'ADR_LIBELLE';
      FieldsList[1] := 'ADR_ADRESSE1';
      FieldsList[2] := 'ADR_ADRESSE2';
      FieldsList[3] := 'ADR_ADRESSE3';
      FieldsList[4] := 'ADR_CODEPOSTAL';
      FieldsList[5] := 'ADR_VILLE';
      FieldsList[6] := 'ADR_PAYS';
      FieldsList[7] := 'ADR_TYPEADRESSE';
      Sql := Format('SELECT %s'
                  + ' FROM ADRESSES'
                  + ' WHERE ADR_REFCODE     = ''%s'''
                  + '   AND ADR_TYPEADRESSE IN (''INT'', ''AFA'')'
                  , [Trim(GetFieldsListFromArray(FieldsList, False)), TobData.GetString('AFF_AFFAIRE')]);
      if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%sSql Adress = %s', [WSCDS_DebugMsg, Sql]), LogValues, 0);
      AdoQryBTP.TSLResult.Clear;
      AdoQryBTP.FieldsList := Trim(GetFieldsListFromArray(FieldsList, False));
      AdoQryBTP.Request := Sql;
      AdoQryBTP.SingleTableSelect;
     {$IFDEF APPSRVWITHCBP}
      Tools.TStringListToTOB(AdoQryBTP.TSLResult, FieldsList, TobAdr, False);
     {$ENDIF APPSRVWITHCBP}
      AdoQryBTP.Reset;
      for Cpt := 0 to pred(TobAdr.Detail.count) do
      begin
        TobAdrL := TobAdr.Detail[Cpt];
        if TobAdrL.GetString('ADR_TYPEADRESSE') = 'INT' then
          AddValues('L')
        else if TobAdrL.GetString('ADR_TYPEADRESSE') = 'AFA' then
          AddValues('F');
      end;
    finally
      FreeAndNil(TobAdr);
    end;
  end;

  procedure AddQuotationList;
  var
    Sql       : string;
    FieldName : string;
    Value     : string;
    Cpt       : integer;
  begin
    FieldName := 'GP_NUMERO';
    Sql := Format('SELECT %s'
                + ' FROM PIECE'
                + ' WHERE GP_NATUREPIECEG = ''DBT'''
                + '       AND GP_TIERS    = ''%s'''
                + '       AND GP_AFFAIRE  = ''%s'''
                  , [FieldName, TobData.GetString('AFF_TIERS'), TobData.GetString('AFF_AFFAIRE')]);
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%sSql Quotation list = %s', [WSCDS_DebugMsg, Sql]), LogValues, 0);
    AdoQryBTP.TSLResult.Clear;
    AdoQryBTP.FieldsList := FieldName;
    AdoQryBTP.Request    := Sql;
    AdoQryBTP.SingleTableSelect;
    for Cpt := 0 to pred(AdoQryBTP.TSLResult.Count) do
      Value := Format('%s;%s', [Value, AdoQryBTP.TSLResult[Cpt]]);
    Value := Copy(Value, 2, length(Value));
    TobAdd.SetString('LISTEDEVIS' , Value);
    AdoQryBTP.Reset
  end;

begin
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%sTExportTreatment.SetLinkedRecords', [WSCDS_DebugMsg]), LogValues, 0);
  TobAdd.ClearDetail;
  case Tn of
    tnChantier :
      begin
        ClearValues;
        AddAdress;
        AddQuotationList;
      end;
  end;
end;

function TExportTreatment.InsertUpdateData(TobData: TOB): boolean;
var
  Cpt       : integer;
  UpdateQty : integer;
  InsertQty : integer;
  OtherQty  : integer;
begin
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%sStart TExportTreatment.InsertUpdateData', [WSCDS_DebugMsg]), LogValues, 0);
  Result := True;
  if (assigned(TobData)) then
  begin
    UpdateQty := TobData.GetInteger('UPDATEQTY');
    InsertQty := TobData.GetInteger('INSERTQTY');
    OtherQty  := TobData.GetInteger('OTHERQTY');
    if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%sUpdateQty=%s, InsertQty=%s, OtherQty=%s', [WSCDS_DebugMsg, IntToStr(UpdateQty), IntToStr(InsertQty), IntToStr(OtherQty)]), LogValues, 0);
    try
      for Cpt := 0 to pred(TobData.Detail.Count) do
      begin
        AdoQryTMP.RecordCount := 0;
        AdoQryTMP.Request     := TobData.Detail[Cpt].GetString('SqlQry');
        if LogValues.DebugEvents > 0 then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%sExécution de %s ', [WSCDS_DebugMsg, AdoQryTMP.Request]), LogValues, 1);
        AdoQryTMP.InsertUpdate;
        AdoQryTMP.Reset;
      end;
      if UpdateQty > 0 then
        TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%s enregistrement(s) de la table %s modifié(s)', [IntToStr(UpdateQty), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
      if InsertQty > 0 then
        TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%s enregistrement(s) de la table %s créé(s)', [IntToStr(InsertQty), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
      if OtherQty > 0 then
        TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%s enregistrement(s) de la table %s non traité(s) car verrouillé(s)', [IntToStr(OtherQty), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
      SetLastSynchro;
    except
      Result := False;
    end;
  end;
end;

function TExportTreatment.ExportTreatment(TobTable, TobAdd, TobQry: TOB): boolean;
var
  TobL          : TOB;
  Cpt           : integer;
  InsertQty     : integer;
  UpdateQty     : integer;
  OtherQty      : integer;
  FieldSize     : integer;
  SqlUnlock     : string;
  Sql           : string;
  Lock          : string;
  Treat         : string;
  KeyFieldsName : string;
  KeyField1     : string;
  KeyField2     : string;
  KeyValue1     : string;
  KeyValue2     : string;
  LabelValue    : string;
  Values        : string;
  ErrValue      : string;
  ErrName       : string;
  ErrType       : string;
  FindData      : boolean;
  CanContinue   : boolean;

  function GetSqlUnlock : string;
  var
    KeyFields : string;
    Key1      : string;
    Key2      : string;
    Exists    : boolean;
  begin
    KeyFields := GetTMPIndexFieldName;
    Key1      := Tools.ReadTokenSt_(KeyFields, ';');
    Key2      := Tools.ReadTokenSt_(KeyFields, ';');
    Exists    := (pos(Key1, SqlUnlock) > 0);
    if KeyField2 = '' then
      Result := Format('%s%s(%s = ''%s'')', [SqlUnlock, Tools.iif(Exists, ' OR ', ''), Key1, KeyValue1])
    else
      Result := Format('%s%s(%s = ''%s'' AND %s = ''%s'')', [SqlUnlock, Tools.iif(Exists, ' OR ', ''), Key1, KeyValue1, Key2, KeyValue2]);
  end;

  function GetErrorMsgFieldType(sValue, FieldName, FieldType : string) : string;
  begin
    if LogValues.LogLevel = 2 then
    begin
      TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%s - Impossible de traiter %s%s - %s, la valeur "%s" ne correspond pas au type du champ %s (%s)'
                                                                 , [  WSCDS_ErrorMsg
                                                                    , KeyValue1                                        
                                                                    , Tools.iif(KeyValue2 <> '', ' - ' + KeyValue2, '')
                                                                    , LabelValue
                                                                    , sValue
                                                                    , FieldName
                                                                    , FieldType
                                                                   ]), LogValues, 2);

    end;
  end;

begin
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%s Start TExportTreatment.TnTreatment', [WSCDS_DebugMsg]), LogValues, 0);
  Result    := True;
  InsertQty := 0;
  UpdateQty := 0;
  OtherQty  := 0;
  SetFieldsArray;
  Sql := GetDataSearchSql(LastSynchro);
  if (LogValues.DebugEvents > 0) then TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('%s TExportTreatment.TnTreatment / Sql = %s', [WSCDS_DebugMsg, Sql]), LogValues, 0);
  AdoQryBTP.TSLResult.Clear;
  AdoQryBTP.FieldsList := Trim(GetFieldsListFromArray(BTPArrFields, True));
  AdoQryBTP.Request    := Sql;
  AdoQryBTP.SingleTableSelect;
  {$IFDEF APPSRVWITHCBP}
  Tools.TStringListToTOB(AdoQryBTP.TSLResult, BTPArrFields, TobTable, True);
  {$ENDIF APPSRVWITHCBP}
  AdoQryBTP.Reset;
  Sql := '';
  if TobTable.Detail.Count > 0 then
  begin
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('Recherche des données (%s enregistrement(s) de la table %s)', [IntToStr(TobTable.Detail.Count), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
    AdoQryTMP.TSLResult.Clear;
    AdoQryTMP.FieldsList := Format('%s_LOCK,%s_TRAITE', [GetTMPPrefix, GetTMPPrefix]);
    AdoQryTMP.LogValues  := LogValues;
    AdoQryTMP.TSLResult.Clear;
    KeyFieldsName := GetTMPIndexFieldName;
    SqlUnlock     := Format('UPDATE %s SET %s_LOCK = ''0'' WHERE ', [TUtilBTPVerdon.GetTMPTableName(Tn), GetTMPPrefix]);
    FindData      := False;
    for Cpt := 0 to pred(TobTable.Detail.Count) do
    begin
      TobL          := TobTable.Detail[Cpt];
      KeyFieldsName := GetBTPIndexFieldName;
      KeyField1     := Tools.ReadTokenSt_(KeyFieldsName, ';');
      KeyField2     := Tools.ReadTokenSt_(KeyFieldsName, ';');
      KeyValue1     := TobL.GetString(KeyField1);
      FieldSize     := GetTMPFieldSizeMax(GetBTPFieldInf(KeyField1, tittName));
      KeyValue1     := Tools.iif(FieldSize > -1, Trim(Copy(KeyValue1, 1, FieldSize)), KeyValue1);
      if KeyField2 <> '' then
      begin
        KeyValue2 := TobL.GetString(KeyField2);
        FieldSize := GetTMPFieldSizeMax(GetBTPFieldInf(KeyField2, tittName));
        KeyField2 := Tools.iif(FieldSize > -1, Trim(Copy(KeyField2, 1, FieldSize)), KeyField2);
      end;
      LabelValue  := TobL.GetString(GetBTPLabelFieldName);
      CanContinue := IsCorrectValueCompareToType(GetFieldIndex(KeyField1, BTPArrFields), KeyValue1, TMPArrFields); // Test si la valeur est correcte par rapport au type
      if CanContinue then 
      begin
        if KeyField2 <> '' then
          CanContinue := True
        else
          CanContinue := IsCorrectValueCompareToType(GetFieldIndex(KeyField2, BTPArrFields), KeyValue2, TMPArrFields);
        if CanContinue then
        begin  
          SetLinkedRecords(TobAdd, TobL);
          AdoQryTMP.Request := GetSqlDataExist(AdoQryTMP.FieldsList, KeyValue1, KeyValue2); // Test si enregistrement existe
          AdoQryTMP.SingleTableSelect;
          if AdoQryTMP.RecordCount = 1 then // Update
          begin
            Values := AdoQryTMP.TSLResult[0];
            Lock   := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
            Treat  := Tools.ReadTokenSt_(Values, ToolsTobToTsl_Separator);
            if Lock = LogValues.FalseValue then
            begin
              Sql := GetSqlUpdate(TobL, TobAdd, KeyValue1, KeyValue2);
              CanContinue := (Sql <> WSCDS_ErrorMsg);
              if CanContinue then
              begin
                if LogValues.LogLevel = 2 then
                  TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('Mise à jour de %s%s - %s', [KeyValue1, Tools.iif(KeyValue2 <> '', ' - ' + KeyValue2, ''), LabelValue]), LogValues, 2);
                inc(UpdateQty);
                FindData  := True;
                SqlUnlock := GetSqlUnlock
              end else
              begin
                Tools.ReadTokenSt_(Sql, ';');
                GetErrorMsgFieldType(Tools.ReadTokenSt_(Sql, ';'), Tools.ReadTokenSt_(Sql, ';'), Tools.ReadTokenSt_(Sql, ';'));
              end;
            end else
            begin
              if LogValues.LogLevel = 2 then
                TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('Mise à jour de %s%s - %s impossible (bloqué)', [KeyValue1, Tools.iif(KeyValue2 <> '', ' - ' + KeyValue2, ''), LabelValue]), LogValues, 2);
              Inc(OtherQty);
            end;
          end else
          begin
            Sql := GetSqlInsert(TobL, TobAdd);
            CanContinue := (pos(WSCDS_ErrorMsg + ';', Sql) = 0);
            if CanContinue then
            begin
              if LogValues.LogLevel = 2 then
                TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('Création de %s%s - %s', [KeyValue1, Tools.iif(KeyValue2 <> '', ' - ' + KeyValue2, ''), LabelValue]), LogValues, 2);
              Inc(InsertQty);
              FindData  := True;
              SqlUnlock := GetSqlUnlock
            end else
            begin
              Tools.ReadTokenSt_(Sql, ';');
              ErrValue := Tools.ReadTokenSt_(Sql, ';');
              ErrName  := Tools.ReadTokenSt_(Sql, ';');
              ErrType  := Tools.ReadTokenSt_(Sql, ';');
              GetErrorMsgFieldType(ErrValue, ErrName, ErrType);
            end;
          end;
          AdoQryTmp.Reset(taqoFieldList);
          if Sql <> '' then
          begin
            TobL := TOB.Create('_QRYL', TobQry, -1);
            TobL.AddChampSupValeur('SqlQry', Sql);
            Sql := '';
          end;
        end else
          GetErrorMsgFieldType(KeyValue2, GetBTPFieldInf(KeyField2, tittName), GetBTPFieldInf(KeyField2, tittType));
      end else
        GetErrorMsgFieldType(KeyValue1, GetBTPFieldInf(KeyField1, tittName), GetBTPFieldInf(KeyField1, tittType));
    end;
    if FindData then
    begin
      TobL      := TOB.Create('_QRYL', TobQry, -1);
      TobL.AddChampSupValeur('SqlQry', SqlUnlock);
    end;
    TobQry.AddChampSupValeur('UPDATEQTY', IntToStr(UpdateQty));
    TobQry.AddChampSupValeur('INSERTQTY', IntToStr(InsertQty));
    TobQry.AddChampSupValeur('OTHERQTY', IntToStr(OtherQty));
    InsertUpdateData(TobQry);
  end else
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, Tn, Format('Aucun %s n''a été trouvé.', [TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
end;

{ TImportTreatment }

{$IFDEF APPSRVWITHCBP}
procedure TImportTreatment.AddDebugEvent(Text : string);
begin
  if (LogValues.DebugEvents > 0) then
  begin
    inc(NumDebugEvent);
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('%s%s-%s', [WSCDS_DebugMsg, IntToStr(NumDebugEvent), Text]), LogValues, 0);
  end;
end;

procedure TImportTreatment.SetMandatoryValues(TobDataL : TOB);
begin
  case Tn of
    tnReglement :
      begin
        if TobDataL.GetString('EDR_MONTANTDEV') = '' then TobDataL.SetString('EDR_MONTANTDEV', TobDataL.GetString('EDR_MONTANT'));
        if TobDataL.GetString('EDR_DEVISE')     = '' then TobDataL.SetString('EDR_DEVISE'    , GetParamSocSecur('SO_DEVISEPRINC', 'EUR'));
      end;
  end;
end;

function TImportTreatment.IsLinkedDocExist(PaymentReferency, DocReferency : string) : boolean;
var
  DocRef : string; 
  Nature : string;
  Souche : string;
  Numero : string;
  Indice : string;
  MaxNum : string;
  iIndex : integer;
begin
  Result := False;
  if DocReferency <> '' then
  begin
    iIndex := TslCacheBTPData.IndexOfName(PrefixCache_GAC_REF + PaymentReferency);
    Result := (iIndex > -1);
    if not Result then
    begin
      DocRef := DocReferency;
      Nature := Tools.ReadTokenSt_(DocRef, ';');
      Souche := Tools.ReadTokenSt_(DocRef, ';');
      Numero := Tools.ReadTokenSt_(DocRef, ';');
      Indice := Tools.ReadTokenSt_(DocRef, ';');
      AdoQryBTP.FieldsList := 'GP_NUMERO';
      AdoQryBTP.Request    := Format('SELECT %s FROM PIECE'
                                   + ' WHERE GP_NATUREPIECEG = ''%s'''
                                   + '   AND GP_SOUCHE       = ''%s'''
                                   + '   AND GP_NUMERO       = %s'
                                   + '   AND GP_INDICEG      = %s'
                                 , [AdoQryBTP.FieldsList, Nature, Souche, Numero, Indice]);
      AdoQryBTP.SingleTableSelect;
      Result := (AdoQryBTP.RecordCount = 1);
      AdoQryBTP.Reset;
      if Result then // Pièce trouvée, ajoute la pièce dans le cache et ajoute aussi le GAC_NUMORDRE le plus grand 
      begin
        TslCacheBTPData.Add(Format('%s%s=%s-%s-%s-%s', [PrefixCache_GAC_REF, PaymentReferency, Nature, Souche, Numero, Indice]));
        iIndex := TslCacheBTPData.IndexOfName(PrefixCache_GAC_MNO + PaymentReferency);
        if iIndex = -1 then
        begin
          AdoQryBTP.FieldsList := 'GAC_NUMORDRE';
          AdoQryBTP.Request    := Format('SELECT MAX(%s) FROM ACOMPTES'
                                       + ' WHERE GAC_NATUREPIECEG = ''%s'''
                                       + '   AND GAC_SOUCHE       = ''%s'''
                                       + '   AND GAC_NUMERO       = %s'
                                       + '   AND GAC_INDICEG      = %s'
                                     , [AdoQryBTP.FieldsList, Nature, Souche, Numero, Indice]);
          AdoQryBTP.SingleTableSelect;
          MaxNum := Tools.iif(AdoQryBTP.RecordCount = 1, AdoQryBTP.TSLResult[0], '0');
          MaxNum := Tools.iif(MaxNum = '', '0', MaxNum);
          AdoQryBTP.Reset;
          TslCacheBTPData.Add(Format('%s%s=%s', [PrefixCache_GAC_MNO, PaymentReferency, MaxNum]))
        end;
      end else
        TslCacheBTPData.Add(Format('%s%s=-', [PrefixCache_GAC_REF, PaymentReferency]));
    end else
      Result := (TslCacheBTPData.ValueFromIndex[iIndex] <> '-');
  end;
end;

function TImportTreatment.GetCompleteItemCode(ItemCode : string) : string;
begin
  Result := Format('%-33sX', [ItemCode]);
end;

function TImportTreatment.GetParamTableName(sType : string) : string;
var
  iIndex : integer;
begin
  iIndex := TslParametersTypeMatching.IndexOfName(sType);
  Result := Tools.iif(iIndex > -1, TslParametersTypeMatching.ValueFromIndex[iIndex], '');
end;

function TImportTreatment.GetBTPAffaireValue(Aff1Code, FieldName : string) : string;
var
  iIndex : integer;
  Value  : string;
  Aff    : string;
  Aff0   : string;
  Aff2   : string;
  Aff3   : string;
begin
  Result := '';
  Value  := '';
  iIndex := TslCacheAffaire.IndexOfName(Aff1Code);
  if iIndex = -1 then
  begin
    AdoQryBTP.FieldsList := 'AFF_AFFAIRE,AFF_AFFAIRE0,AFF_AFFAIRE2,AFF_AFFAIRE3';
    AdoQryBTP.Request    := Format('SELECT %s FROM AFFAIRE WHERE AFF_AFFAIRE1 = ''%s'''
                               , [AdoQryBTP.FieldsList, Aff1Code]);
    AdoQryBTP.SingleTableSelect;
    if AdoQryBTP.RecordCount > 0 then
    begin
      Value := AdoQryBTP.TSLResult[0];
      TslCacheAffaire.Add(Format('%s=%s', [Aff1Code, Value]));
    end else
      TslCacheAffaire.Add(Format('%s=%s', [Aff1Code, WSCDS_ErrorMsg]));
    AdoQryBTP.Reset;
  end else
    Value := copy(TslCacheAffaire[iIndex], pos('=', TslCacheAffaire[iIndex])+1, length(TslCacheAffaire[iIndex]));
  if pos(WSCDS_ErrorMsg, Value) = 0 then
  begin
    Aff  := Tools.ReadTokenSt_(Value, '^');
    Aff0 := Tools.ReadTokenSt_(Value, '^');
    Aff2 := Tools.ReadTokenSt_(Value, '^');
    Aff3 := Tools.ReadTokenSt_(Value, '^');
    case Tools.CaseFromString(FieldName, ['AFF_AFFAIRE', 'AFF_AFFAIRE0', 'AFF_AFFAIRE2', 'AFF_AFFAIRE3']) of
      {AFF_AFFAIRE}  0 : Result := Aff;
      {AFF_AFFAIRE0} 1 : Result := Aff0;
      {AFF_AFFAIRE2} 2 : Result := Aff2;
      {AFF_AFFAIRE3} 3 : Result := Aff3;
    end;
  end;
end;

function TImportTreatment.GetBTPArticleValue(Itemcode, FieldName : string) : string;
var
  iIndex       : integer;
  Value        : string;
  CompleteItem : string;
  FTax1        : string;
  FN1          : string;
  FN2          : string;
  FN3          : string;
  SalesUnit    : string;
begin
  Result := '';
  Value  := '';
  iIndex := TslCacheArticle.IndexOfName(Itemcode);
  if iIndex = -1 then
  begin
    AdoQryBTP.FieldsList := 'GA_ARTICLE,GA_FAMILLETAXE1,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_QUALIFUNITEVTE';
    AdoQryBTP.Request    := Format('SELECT %s FROM ARTICLE WHERE GA_CODEARTICLE = ''%s'''
                               , [AdoQryBTP.FieldsList, Itemcode]);
    AdoQryBTP.SingleTableSelect;
    if AdoQryBTP.RecordCount > 0 then
      Value := AdoQryBTP.TSLResult[0]
    else
      Value := WSCDS_ErrorMsg;
    TslCacheArticle.Add(Format('%s=%s', [ItemCode, Tools.iif(AdoQryBTP.RecordCount > 0, Value, WSCDS_ErrorMsg)]));
    AdoQryBTP.Reset;
  end else
    Value := copy(TslCacheArticle[iIndex], pos('=', TslCacheArticle[iIndex])+1, length(TslCacheArticle[iIndex]));
  if pos(WSCDS_ErrorMsg, Value) = 0 then
  begin
    CompleteItem := Tools.ReadTokenSt_(Value, '^');
    FTax1        := Tools.ReadTokenSt_(Value, '^');
    FN1          := Tools.ReadTokenSt_(Value, '^');
    FN2          := Tools.ReadTokenSt_(Value, '^');
    FN3          := Tools.ReadTokenSt_(Value, '^');
    SalesUnit    := Tools.ReadTokenSt_(Value, '^');
    case Tools.CaseFromString(FieldName, [  'GA_ARTICLE'       , 'GA_FAMILLETAXE1', 'GA_FAMILLENIV1', 'GA_FAMILLENIV2', 'GA_FAMILLENIV3'
                                          , 'GA_QUALIFUNITEVTE'
                                         ]) of
      {GA_ARTICLE}        0 : Result := CompleteItem;
      {GA_FAMILLETAXE1}   1 : Result := FTax1;
      {GA_FAMILLENIV1}    2 : Result := FN1;
      {GA_FAMILLENIV2}    3 : Result := FN2;
      {GA_FAMILLENIV3}    4 : Result := FN3;
      {GA_QUALIFUNITEVTE} 5 : Result := SalesUnit;
    end;
  end;
end;
  
function TImportTreatment.GetBTPTableName(Tn : T_TablesName; sType : string='') : string;
begin
  case Tn of
    tnTiers        : Result := Tools.GetTableNameFromTtn(ttnTiers);
    tnDevis        : Result := Tools.GetTableNameFromTtn(ttnPiece);
    tnLignesBR     : Result := Tools.GetTableNameFromTtn(ttnLigne);
    tnChantier     : Result := Tools.GetTableNameFromTtn(ttnAffaire);
    tnParameters   : Result := GetParamTableName(sType);
    tnArticles     : Result := Tools.GetTableNameFromTtn(ttnArticle);
    tnSalaries     : Result := Tools.GetTableNameFromTtn(ttnRessource);
    tnModeRegle    : Result := Tools.GetTableNameFromTtn(ttnModePaie);
    tnReglement    : Result := Tools.GetTableNameFromTtn(ttnAcomptes);
    tnHresSalaries : Result := Tools.GetTableNameFromTtn(ttnConsommations);
    tnConsoStock   : Result := Tools.GetTableNameFromTtn(ttnConsommations);
    tnIntervenants : Result := Tools.GetTableNameFromTtn(ttnAffTiers);
  else
    Result := '';
  end;
end;

function TImportTreatment.GetBTPWhere(TobDataL : TOB) : string;
var
  TypeParam : string;
begin
  Result := '';
  case Tn of
    tnParameters :
      begin
        TypeParam := GetParamTableName(TobDataL.GetString('PAR_TYPE'));
        if TypeParam <> '' then
        begin
          case Tools.CaseFromString(TypeParam, ['CHOIXCOD', 'CHOIXEXT', 'FONCTION']) of
           {CHOIXCOD} 0 : Result := Format('CC_TYPE = ''%s'' AND CC_CODE = ''%s''', [TobDataL.GetString('PAR_TYPE'), TobDataL.GetString('PAR_CODE')]);
           {CHOIXEXT} 1 : Result := Format('YX_TYPE = ''%s'' AND YX_CODE = ''%s''', [TobDataL.GetString('PAR_TYPE'), TobDataL.GetString('PAR_CODE')]);
           {FONCTION} 2 : Result := Format('AFO_FONCTION = ''%s'''                , [TobDataL.GetString('PAR_CODE')]);
          end;
        end;
      end;
    tnArticles     : Result := Format('GA_ARTICLE     = ''%s''', [GetCompleteItemCode(TobDataL.GetString('ART_CODEARTICLE'))]);
    tnSalaries     : Result := Format('ARS_RESSOURCE  = ''%s''', [TobDataL.GetString('SAL_CODESAL')]);
    tnModeRegle    : Result := Format('MP_MODEPAIE    = ''%s''', [TobDataL.GetString('MDR_CODE')]);
    tnReglement    : Result := Format('GAC_REFORIGINE = ''%s''', [TobDataL.GetString('EDR_IDENTIFIANT')]);
    tnHresSalaries : Result := Format('BCO_NUMMOUV    = ''%s''', [TobDataL.GetString('HSA_NUMLSE')]);
//    tnConsoStock   : Result := Format('BCO_NUMMOUV    = ''%s''', [TobDataL.GetString('HSA_NUMLSE')]);
  end;
end;

function TImportTreatment.GetBTPUpdateValues(TobDataL : TOB) : string;
var
  TypeParam : string;
begin
  Result := '';
  SetMandatoryValues(TobDataL);
  case Tn of
    tnParameters :
      begin
        TypeParam := GetParamTableName(TobDataL.GetString('PAR_TYPE'));
        if TypeParam <> '' then
        begin
          case Tools.CaseFromString(TypeParam, ['CHOIXCOD', 'CHOIXEXT', 'FONCTION']) of
            {CHOIXCOD} 0 : Result := Format('CC_LIBELLE = ''%s'', CC_ABREGE = ''%s''', [TobDataL.GetString('PAR_LIBELLE'), TobDataL.GetString('PAR_INFOCOMPL')]);
            {CHOIXEXT} 1 : Result := Format('YX_LIBELLE = ''%s'', YX_ABREGE = ''%s''', [TobDataL.GetString('PAR_LIBELLE'), TobDataL.GetString('PAR_INFOCOMPL')]);
            {FONCTION} 2 : Result := Format('AFO_LIBELLE = ''%s'''                   , [TobDataL.GetString('PAR_LIBELLE')]);
          end;
        end;
      end;
    tnArticles :
      Result := Format('  GA_NATUREPRES    = ''%s'''
                     + ', GA_LIBELLE       = ''%s'''
                     + ', GA_BLOCNOTE      = ''%s'''
                     + ', GA_COMPTAARTICLE = ''%s'''
                     + ', GA_FAMILLENIV1   = ''%s'''
                     + ', GA_FAMILLENIV2   = ''%s'''
                     + ', GA_FAMILLENIV3   = ''%s'''
                     + ', GA_LIBREART1     = ''%s'''
                     + ', GA_FERME         = ''%s'''
                     + ', GA_SOCIETE       = ''%s'''
                     , [  TobDataL.GetString('ART_NATUREPRESTATION')
                        , TobDataL.GetString('ART_LIBELLE')
                        , TobDataL.GetString('ART_BLOCNOTE')
                        , TobDataL.GetString('ART_COMPTA')
                        , TobDataL.GetString('ART_FAMILLENIV1')
                        , TobDataL.GetString('ART_FAMILLENIV2')
                        , TobDataL.GetString('ART_FAMILLENIV3')
                        , TobDataL.GetString('ART_MARQUE')
                        , TobDataL.GetString('ART_ACTIF')
                        , TobDataL.GetString('ART_CODESOCIETE')
                       ]);
    tnSalaries :
      Result := Format('  ARS_LIBELLE       = ''%s'''
                     + ', ARS_LIBELLE2      = ''%s'''
                     + ', ARS_FONCTION1     = ''%s'''
                     + ', ARS_TAUXREVIENTUN = %s'
                     + ', ARS_IMMAT         = ''%s'''
                     + ', ARS_LIBRERES1     = ''%s'''
//                     + ', ARS_UTILASSOCIE   = ''%s'''
                     + ', ARS_DATEMODIF     = ''%s'''
                     , [  TobDataL.GetString('SAL_NOM')
                        , TobDataL.GetString('SAL_PRENOM')
                        , TobDataL.GetString('SAL_FONCTION')
                        , FloatToStr(Valeur(TobDataL.GetString('SAL_PR')))
                        , TobDataL.GetString('SAL_MATRICULE')
                        , TobDataL.GetString('SAL_BU')
//                        , TobDataL.GetString('SAL_CODEUTILISATEUR')
                        , Tools.CastDateTimeForQry(TobDataL.GetDateTime('SAL_DATEMODIF'))
                       ]);
    tnModeRegle :
      Result := Format('  MP_LIBELLE       = ''%s'''
                     + ', MP_CATEGORIE     = ''%s'''
                     , [  TobDataL.GetString('MDR_LIBELLE')
                        , TobDataL.GetString('MDR_CATEGORIE')
                       ]);
    tnReglement :
      Result    := Format('  GAC_JALECR       = ''%s'''
                        + ', GAC_MONTANT      = %s'
                        + ', GAC_MONTANTDEV   = %s'
                        + ', GAC_MODEPAIE     = ''%s'''
                        + ', GAC_LIBELLE      = ''%s'''
                        + ', GAC_NUMCHEQUE    = ''%s'''
                        + ', GAC_DATEECR      = ''%s'''
                        + ', GAC_DATEECHEANCE = ''%s'''
                        , [  TobDataL.GetString('EDR_JOURNAL')
                           , FloatToStr(Valeur(TobDataL.GetString('EDR_MONTANT')))
                           , FloatToStr(Valeur(TobDataL.GetString('EDR_MONTANTDEV')))
                           , TobDataL.GetString('EDR_MODEPAIE')
                           , TobDataL.GetString('EDR_LIBELLE')
                           , TobDataL.GetString('EDR_NUMCHEQUE')
                           , Tools.CastDateTimeForQry(TobDataL.GetDateTime('EDR_DATE'))
                           , Tools.CastDateTimeForQry(TobDataL.GetDateTime('EDR_DATEECHEANCE'))
                          ]);
    tnHresSalaries :
      Result := Format('  BCO_DATEMOUV    = ''%s'''
                     + ', BCO_AFFAIRE1    = ''%s'''
                     + ', BCO_RESSOURCE   = ''%s'''
                     + ', BCO_LIBELLE     = ''%s'''
                     + ', BCO_CODEARTICLE = ''%s'''
                     + ', BCO_QUANTITE    = %s'
                     + ', BCO_TYPEHEURE   = ''%s'''
                     + ', BCO_DPA         = %s'
                     + ', BCO_DPR         = %s'
                     + ', BCO_PUHT        = %s'
                     , [  Tools.CastDateTimeForQry(TobDataL.GetDateTime('HSA_DATEMVT'))
                        , TobDataL.GetString('HSA_CODECHA')
                        , TobDataL.GetString('HSA_CODESAL')
                        , TobDataL.GetString('HSA_LIBELLE')
                        , TobDataL.GetString('HSA_CODEPRESTATION')
                        , FloatToStr(Valeur(TobDataL.GetString('HSA_NBHRS')))
                        , TobDataL.GetString('HSA_TYPEHRS')
                        , FloatToStr(Valeur(TobDataL.GetString('HSA_DPA')))
                        , FloatToStr(Valeur(TobDataL.GetString('HSA_DPR')))
                        , FloatToStr(Valeur(TobDataL.GetString('HSA_PUV')))
                       ]);
  end;
end;

function TImportTreatment.GetBTPInsertValues(TobDataL : TOB) : string;
var
  lFieldsList : string;
  FieldName   : string;
  FieldValue  : string;
  TypeParam   : string;
  Nature      : string;
  Souche      : string;
  Numero      : string;
  Indice      : string;
  MaxNum      : string;
  AFFCode     : string;
  iIndex      : integer;
  MvtNumber   : double;
  IsHSA       : boolean;
  OnError     : boolean;

  function GetDoubleValue(Value : double) : string;
  begin
    Result := FormatFloat('#0.00', Value);
    Result := StringReplace(Result, ',', '.', [rfReplaceAll]);
  end;

  function GetBcoQteFromHSA(Value : double) : double;
  begin
    Result := Arrondi(Tools.ConvertMinuteTo100(TobDataL.GetInteger('HSA_NBHRS')), 2);
  end;

begin
  Result := '';
  SetMandatoryValues(TobDataL);
  case Tn of
    tnParameters :
    begin
      TypeParam := GetParamTableName(TobDataL.GetString('PAR_TYPE'));
      if TypeParam <> '' then
      begin
        lFieldsList := GetBTPFieldsList(TobDataL.GetString('PAR_TYPE'));
        while lFieldsList <> '' do
        begin
          FieldName := Tools.ReadTokenSt_(lFieldsList, ',');
          case Tools.CaseFromString(TypeParam, ['CHOIXCOD', 'CHOIXEXT', 'FONCTION']) of
            {CHOIXCOD, CHOIXEXT} 0..1 :
            begin
              case Tools.CaseFromString(FieldName, ['CC_TYPE', 'YX_TYPE', 'CC_CODE', 'YX_CODE', 'CC_LIBELLE', 'YX_LIBELLE', 'CC_ABREGE', 'YX_ABREGE']) of
                {xx_TYPE}    0..1 : FieldValue := Format('''%s''', [TobDataL.GetString('PAR_TYPE')]);
                {xx_CODE}    2..3 : FieldValue := Format('''%s''', [TobDataL.GetString('PAR_CODE')]);
                {xx_LIBELLE} 4..5 : FieldValue := Format('''%s''', [GetValue('PAR_LIBELLE'  , TobDataL.GetString('PAR_LIBELLE')  , False)]);
                {xx_ABREGE}  6..7 : FieldValue := Format('''%s''', [GetValue('PAR_INFOCOMPL', TobDataL.GetString('PAR_INFOCOMPL'), False)]);
              else
                FieldValue := '''''';
              end;
            end;
            {FONCTION} 2 :
            begin
              case Tools.CaseFromString(FieldName, ['AFO_FONCTION', 'AFO_LIBELLE']) of
                {AFO_FONCTION} 0 : FieldValue := Format('''%s''', [TobDataL.GetString('PAR_CODE')]);
                {AFO_LIBELLE}  1 : FieldValue := Format('''%s''', [copy(GetValue('PAR_LIBELLE', TobDataL.GetString('PAR_LIBELLE'), False), 1, 35)]);
              else
                FieldValue := GetBTPDefaultValue(FieldName);
              end;
            end;
          end;
          Result := Format('%s, %s', [Result, FieldValue]);
        end;
      end;
    end;
    tnArticles :
    begin
      TypeParam   := 'MAR';
      lFieldsList := ArticlesFieldsList;
      while lFieldsList <> '' do
      begin
        FieldName := Tools.ReadTokenSt_(lFieldsList, ',');
        case Tools.CaseFromString(FieldName, [  'GA_ARTICLE'        , 'GA_TYPEARTICLE'    , 'GA_CODEARTICLE'     , 'GA_NATUREPRES'     , 'GA_LIBELLE'
                                              , 'GA_BLOCNOTE'       , 'GA_COMPTAARTICLE'  , 'GA_DOMAINE'         , 'GA_FAMILLENIV1'    , 'GA_FAMILLENIV2'
                                              , 'GA_FAMILLENIV3'    , 'GA_LIBREART1'      , 'GA_FERME'           , 'GA_SOCIETE'        , 'GA_STATUTART'
                                              , 'GA_CALCPRIXHT'     , 'GA_CALCPRIXPR'     , 'GA_ACTIVITEPREPRISE', 'GA_CREERPAR'       , 'GA_UTILISATEUR'
                                              , 'GA_CREATEUR'       , 'GA_FAMILLETAXE1'   , 'GA_INVISIBLEWEB'    , 'GA_REMISEPIED'     , 'GA_ESCOMPTABLE'
                                              , 'GA_COMMISSIONNABLE', 'GA_DATESUPPRESSION', 'GA_DATECREATION'    , 'GA_DATEMODIF'      , 'GA_PRIXPOURQTE'
                                              , 'GA_QUALIFMARGE'    , 'GA_GEREANAL'       , 'GA_REMISELIGNE'     , 'GA_ACTIVITEREPRISE', 'GA_CALCPRIXTTC'
                                              , 'GA_ACTIVITEEFFECT' , 'GA_QUALIFUNITESTO' , 'GA_QUALIFUNITEVTE'  , 'GA_QUALIFUNITEACT'
                                             ]) of
          {GA_ARTICLE}          0  : FieldValue := Format('''%s''', [GetCompleteItemCode(TobDataL.GetString('ART_CODEARTICLE'))]);
          {GA_TYPEARTICLE}      1  : FieldValue := Format('''%s''', [TypeParam]);
          {GA_CODEARTICLE}      2  : FieldValue := Format('''%s''', [TobDataL.GetString('ART_CODEARTICLE')]);
          {GA_NATUREPRES}       3  : FieldValue := Format('''%s''', [TobDataL.GetString('ART_NATUREPRESTATION')]);
          {GA_LIBELLE}          4  : FieldValue := Format('''%s''', [GetValue('ART_LIBELLE' , TobDataL.GetString('ART_LIBELLE') , False)]);
          {GA_BLOCNOTE}         5  : FieldValue := Format('''%s''', [GetValue('ART_BLOCNOTE', TobDataL.GetString('ART_BLOCNOTE'), False)]);
          {GA_COMPTAARTICLE}    6  : FieldValue := Format('''%s''', [TobDataL.GetString('ART_COMPTA')]);
          {GA_DOMAINE}          7  : FieldValue := Format('''%s''', [TobDataL.GetString('ART_DOMAINE')]);
          {GA_FAMILLENIV1}      8  : FieldValue := Format('''%s''', [TobDataL.GetString('ART_FAMILLENIV1')]);
          {GA_FAMILLENIV2}      9  : FieldValue := Format('''%s''', [TobDataL.GetString('ART_FAMILLENIV2')]);
          {GA_FAMILLENIV3}      10 : FieldValue := Format('''%s''', [TobDataL.GetString('ART_FAMILLENIV3')]);
          {GA_LIBREART1}        11 : FieldValue := Format('''%s''', [TobDataL.GetString('ART_MARQUE')]);
          {GA_FERME}            12 : FieldValue := Format('''%s''', [Tools.iif(TobDataL.GetString('ART_ACTIF') = '1', 'X', '-')]);
          {GA_SOCIETE}          13 : FieldValue := Format('''%s''', [TobDataL.GetString('ART_CODESOCIETE')]);
          {GA_STATUTART}        14 : FieldValue := '''UNI''';
          {GA_CALCPRIXHT}       15 : FieldValue := '''AUC''';
          {GA_CALCPRIXPR}       16 : FieldValue := '''AUC''';
          {GA_ACTIVITEPREPRISE} 17 : FieldValue := '''F''';
          {GA_CREERPAR}         18 : FieldValue := '''IMP''';
          {GA_UTILISATEUR}      19 : FieldValue := Format('''%s''', [FolderValues.BTPUserAdmin]);
          {GA_CREATEUR}         20 : FieldValue := Format('''%s''', [FolderValues.BTPUserAdmin]);
          {GA_FAMILLETAXE1}     21 : FieldValue := Format('''%s''', [GetParamSocSecur('SO_GCFAMILLETAXE1', '')]);
          {GA_INVISIBLEWEB}     22 : FieldValue := '''X''';
          {GA_REMISEPIED}       23 : FieldValue := '''X''';
          {GA_ESCOMPTABLE}      24 : FieldValue := '''X''';
          {GA_COMMISSIONNABLE}  25 : FieldValue := '''X''';
          {GA_DATESUPPRESSION}  26 : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(IDate2099)]);
          {GA_DATECREATION}     27 : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(TobDataL.GetDateTime('ART_DATECREATION'))]);
          {GA_DATEMODIF}        28 : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(TobDataL.GetDateTime('ART_DATEMODIF'))]);
          {GA_PRIXPOURQTE}      29 : FieldValue := '1';
          {GA_QUALIFMARGE}      30 : FieldValue := '''CO''';
          {GA_GEREANAL}         31 : FieldValue := '''X''';
          {GA_REMISELIGNE}      32 : FieldValue := '''X''';
          {GA_ACTIVITEREPRISE}  33 : FieldValue := '''F''';
          {GA_CALCPRIXTTC}      34 : FieldValue := '''AUC''';
          {GA_ACTIVITEEFFECT}   35 : FieldValue := Format('''%s''', [Tools.iif(TypeParam = 'PRE', 'X', '-')]);
          {GA_QUALIFUNITESTO}   36 : FieldValue := Format('''%s''', [Tools.iif(TypeParam = 'PRE', GetParamSocSecur('SO_AFMESUREACTIVITE', ''), '')]);
          {GA_QUALIFUNITEVTE}   37 : FieldValue := Format('''%s''', [Tools.iif(TypeParam = 'PRE', GetParamSocSecur('SO_AFMESUREACTIVITE', ''), '')]);
          {GA_QUALIFUNITEACT}   38 : FieldValue := Format('''%s''', [Tools.iif(TypeParam = 'PRE', GetParamSocSecur('SO_AFMESUREACTIVITE', ''), '')]);
        else
          FieldValue := GetBTPDefaultValue(FieldName);
        end;
        Result := Format('%s, %s', [Result, FieldValue]);
      end;
    end;
    tnSalaries :
    begin
      lFieldsList := SalariesFieldsList;
      while lFieldsList <> '' do
      begin
        FieldName := Tools.ReadTokenSt_(lFieldsList, ',');
        case Tools.CaseFromString(FieldName, [  'ARS_RESSOURCE'    , 'ARS_TYPERESSOURCE', 'ARS_LIBELLE'  , 'ARS_LIBELLE2'   , 'ARS_FONCTION1'
                                              , 'ARS_TAUXREVIENTUN', 'ARS_IMMAT'        , 'ARS_LIBRERES1', 'ARS_UTILASSOCIE', 'ARS_DATEMODIF'
                                              , 'ARS_DATECREATION' , 'ARS_SOCIETE'
                                             ]) of
          {ARS_RESSOURCE}     0  : FieldValue := Format('''%s''', [TobDataL.GetString('SAL_CODESAL')]);
          {ARS_TYPERESSOURCE} 1  : FieldValue := Format('''%s''', [TobDataL.GetString('SAL_TYPE')]);
          {ARS_LIBELLE}       2  : FieldValue := Format('''%s''', [GetValue('SAL_NOM', TobDataL.GetString('SAL_NOM'), False)]);
          {ARS_LIBELLE2}      3  : FieldValue := Format('''%s''', [GetValue('SAL_PRENOM', TobDataL.GetString('SAL_PRENOM'), False)]);
          {ARS_FONCTION1}     4  : FieldValue := Format('''%s''', [TobDataL.GetString('SAL_FONCTION')]);
          {ARS_TAUXREVIENTUN} 5  : FieldValue := Format('%s'    , [FloatToStr(Valeur(TobDataL.GetString('SAL_PR')))]);
          {ARS_IMMAT}         6  : FieldValue := Format('''%s''', [TobDataL.GetString('SAL_MATRICULE')]);
          {ARS_LIBRERES1}     7  : FieldValue := Format('''%s''', [TobDataL.GetString('SAL_BU')]);
          {ARS_UTILASSOCIE}   8  : FieldValue := Format('''%s''', [TobDataL.GetString('SAL_CODEUTILISATEUR')]);
          {ARS_DATEMODIF}     9  : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(NowH)]);
          {ARS_DATECREATION}  10 : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(NowH)]);
          {ARS_SOCIETE}       11 : FieldValue := Format('''%s''', [SocietyCode]);
        else
          FieldValue := GetBTPDefaultValue(FieldName);
        end;
        Result := Format('%s, %s', [Result, FieldValue]);
      end;
    end;
    tnModeRegle :
    begin
      lFieldsList := ModeRegleFieldsList;
      while lFieldsList <> '' do
      begin
        FieldName := Tools.ReadTokenSt_(lFieldsList, ',');
        case Tools.CaseFromString(FieldName, ['MP_MODEPAIE', 'MP_LIBELLE', 'MP_CATEGORIE']) of
          {MP_MODEPAIE}  0  : FieldValue := Format('''%s''', [TobDataL.GetString('MDR_CODE')]);
          {MP_LIBELLE}   1  : FieldValue := Format('''%s''', [GetValue('MDR_LIBELLE', TobDataL.GetString('MDR_LIBELLE'), False)]);
          {MP_CATEGORIE} 2  : FieldValue := Format('''%s''', [TobDataL.GetString('MDR_CATEGORIE')]);
        else
          FieldValue := GetBTPDefaultValue(FieldName);
        end;
        Result := Format('%s, %s', [Result, FieldValue]);
      end;
    end;
    tnReglement :
    begin
      iIndex := TslCacheBTPData.IndexOfName(PrefixCache_GAC_REF + TobDataL.GetString('EDR_IDENTIFIANT'));
      if iIndex > -1 then
      begin
        lFieldsList := ReglementFieldsList;
        TypeParam   := TslCacheBTPData.ValueFromIndex[iIndex];
        Nature      := Tools.ReadTokenSt_(TypeParam, '-');
        Souche      := Tools.ReadTokenSt_(TypeParam, '-');
        Numero      := Tools.ReadTokenSt_(TypeParam, '-');
        Indice      := Tools.ReadTokenSt_(TypeParam, '-');
        iIndex      := TslCacheBTPData.IndexOfName(PrefixCache_GAC_MNO + TobDataL.GetString('EDR_IDENTIFIANT'));
        MaxNum      := IntToStr(StrToInt(TslCacheBTPData.ValueFromIndex[iIndex]) + 1);
        TslCacheBTPData[iIndex] := Format('%s%s=%s', [PrefixCache_GAC_MNO, TobDataL.GetString('EDR_IDENTIFIANT'), MaxNum]);
        while lFieldsList <> '' do
        begin
          FieldName := Tools.ReadTokenSt_(lFieldsList, ',');
          case Tools.CaseFromString(FieldName, [  'GAC_NATUREPIECEG', 'GAC_SOUCHE'    , 'GAC_NUMERO'      , 'GAC_INDICEG'   , 'GAC_JALECR'
                                                , 'GAC_MONTANT'     , 'GAC_MONTANTDEV', 'GAC_MODEPAIE'    , 'GAC_LIBELLE'   , 'GAC_NUMCHEQUE'
                                                , 'GAC_REFORIGINE'  , 'GAC_DATEECR'   , 'GAC_DATEECHEANCE', 'GAC_AUXILIAIRE', 'GAC_NUMORDRE'
                                                , 'GAC_DEVISE'
                                               ]) of
            {GAC_NATUREPIECEG} 0  : FieldValue := Format('''%s''', [Nature]);
            {GAC_SOUCHE}       1  : FieldValue := Format('''%s''', [Souche]);
            {GAC_NUMERO}       2  : FieldValue := Format('''%s''', [Numero]);
            {GAC_INDICEG}      3  : FieldValue := Format('''%s''', [Indice]);
            {GAC_JALECR}       4  : FieldValue := Format('''%s''', [TobDataL.GetString('EDR_JOURNAL')]);
            {GAC_MONTANT}      5  : FieldValue := Format('%s'    , [FloatToStr(Valeur(TobDataL.GetString('EDR_MONTANT')))]);
            {GAC_MONTANTDEV}   6  : FieldValue := Format('%s'    , [FloatToStr(Valeur(TobDataL.GetString('EDR_MONTANTDEV')))]);
            {GAC_MODEPAIE}     7  : FieldValue := Format('''%s''', [TobDataL.GetString('EDR_MODEPAIE')]);
            {GAC_LIBELLE}      8  : FieldValue := Format('''%s''', [GetValue('EDR_LIBELLE', TobDataL.GetString('EDR_LIBELLE'), False)]);
            {GAC_NUMCHEQUE}    9  : FieldValue := Format('''%s''', [TobDataL.GetString('EDR_NUMCHEQUE')]);
            {GAC_REFORIGINE}   10 : FieldValue := Format('''%s''', [TobDataL.GetString('EDR_IDENTIFIANT')]);
            {GAC_DATEECR}      11 : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(TobDataL.GetDateTime('EDR_DATE'))]);
            {GAC_DATEECHEANCE} 12 : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(TobDataL.GetDateTime('EDR_DATEECHEANCE'))]);
            {GAC_AUXILIAIRE}   13 : FieldValue := Format('''%s''', [TobDataL.GetString('EDR_CODETIERS')]);
            {GAC_NUMORDRE}     14 : FieldValue := Format('%s'    , [FloatToStr(Valeur(MaxNum))]);
            {GAC_DEVISE}       15 : FieldValue := Format('''%s''', [TobDataL.GetString('EDR_DEVISE')]);
          else
            FieldValue := GetBTPDefaultValue(FieldName);
          end;
          Result := Format('%s, %s', [Result, FieldValue]);
        end;
      end;
    end;
    tnHresSalaries, tnConsoStock :
    begin
      if Tools.GetNumUniqueConso(MvtNumber, AdoQryBTP.ServerName, AdoQryBTP.DBName) = GncOk then
      begin
        IsHSA       := (Tn = tnHresSalaries);
        lFieldsList := ConsoStockFieldsList;
        while lFieldsList <> '' do
        begin
          FieldName := Tools.ReadTokenSt_(lFieldsList, ',');
          case Tools.CaseFromString(FieldName, [  'BCO_DATEMOUV'    , 'BCO_AFFAIRE1'   , 'BCO_CODEARTICLE', 'BCO_LIBELLE'      , 'BCO_QUANTITE'
                                                , 'BCO_DPA'         , 'BCO_DPR'        , 'BCO_PUHT'       , 'BCO_NUMMOUV'      , 'BCO_QTEVENTE'
                                                , 'BCO_AFFAIRE'     , 'BCO_AFFAIRE0'   , 'BCO_AFFAIRE2'   , 'BCO_AFFAIRE3'     , 'BCO_MOIS'
                                                , 'BCO_SEMAINE'     , 'BCO_NATUREMOUV' , 'BCO_ARTICLE'    , 'BCO_QUALIFQTEMOUV', 'BCO_FACTURABLE'
                                                , 'BCO_FAMILLETAXE1', 'BCO_DATETRAVAUX', 'BCO_FAMILLENIV1', 'BCO_FAMILLENIV2'  , 'BCO_FAMILLENIV3'
                                                , 'BCO_TRAITEVENTE' , 'BCO_MONTANTACH' , 'BCO_MONTANTPR'  , 'BCO_MONTANTHT'    , 'BCO_RESSOURCE'
                                                , 'BCO_NATUREPIECEG', 'BCO_SOUCHE'     , 'BCO_NUMERO'     , 'BCO_INDICEG'      , 'BCO_FROMDEVIS'
                                               ]) of
            {BCO_DATEMOUV}      0  : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(TobDataL.GetDateTime(Tools.iif(IsHSA, 'HSA_DATEMVT', 'COS_DATEMVT')))]);
            {BCO_AFFAIRE1}      1  : FieldValue := Format('''%s''', [TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODECHA', 'COS_CODECHA'))]);
            {BCO_CODEARTICLE}   2  : FieldValue := Format('''%s''', [TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE'))]);
            {BCO_LIBELLE}       3  : FieldValue := Format('''%s''', [GetValue(Tools.iif(IsHSA, 'HSA_LIBELLE', 'COS_COMMENTAIRE'), Copy(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_LIBELLE', 'COS_COMMENTAIRE')), 1, 35), False)]);
            {BCO_QUANTITE}      4  : FieldValue := Format('%s'    , [Tools.iif(IsHSA, GetDoubleValue(GetBcoQteFromHSA(TobDataL.GetInteger('HSA_NBHRS'))), GetDoubleValue(TobDataL.GetDouble('COS_QTE')))]);
            {BCO_DPA}           5  : FieldValue := Format('%s'    , [GetDoubleValue(TobDataL.GetDouble(Tools.iif(IsHSA, 'HSA_DPA', 'COS_DPA')))]);
            {BCO_DPR}           6  : FieldValue := Format('%s'    , [GetDoubleValue(TobDataL.GetDouble(Tools.iif(IsHSA, 'HSA_DPR', 'COS_DPR')))]);
            {BCO_PUHT}          7  : FieldValue := Format('%s'    , [GetDoubleValue(TobDataL.GetDouble(Tools.iif(IsHSA, 'HSA_PUV', 'COS_PUV')))]);
            {BCO_NUMMOUV}       8  : FieldValue := Format('%s'    , [FloatToStr(MvtNumber)]);
            {BCO_QTEVENTE}      9  : FieldValue := Format('%s'    , [GetDoubleValue(TobDataL.GetDouble(Tools.iif(IsHSA, '', 'COS_QTE')))]);
            {BCO_AFFAIRE}       10 : FieldValue := Format('''%s''', [GetBTPAffaireValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODECHA', 'COS_CODECHA')), 'AFF_AFFAIRE')]);
            {BCO_AFFAIRE0}      11 : FieldValue := Format('''%s''', [GetBTPAffaireValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODECHA', 'COS_CODECHA')), 'AFF_AFFAIRE0')]);
            {BCO_AFFAIRE2}      12 : FieldValue := Format('''%s''', [GetBTPAffaireValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODECHA', 'COS_CODECHA')), 'AFF_AFFAIRE2')]);
            {BCO_AFFAIRE3}      13 : FieldValue := Format('''%s''', [GetBTPAffaireValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODECHA', 'COS_CODECHA')), 'AFF_AFFAIRE3')]);
            {BCO_MOIS}          14 : FieldValue := Format('%s'    , [IntToStr(MonthOfTheYear(StrToDate(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_DATEMVT', 'COS_DATEMVT')))))]);
            {BCO_SEMAINE}       15 : FieldValue := Format('%s'    , [IntToStr(WeekOfTheYear(StrToDate(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_DATEMVT', 'COS_DATEMVT')))))]);
            {BCO_NATUREMOUV}    16 : FieldValue := Format('''%s''', [Tools.iif(IsHSA, 'MO', 'FOU')]);
            {BCO_ARTICLE}       17 : FieldValue := Format('''%s''', [GetBTPArticleValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE')), 'GA_ARTICLE')]);
            {BCO_QUALIFQTEMOUV} 18 : FieldValue := Format('''%s''', [GetBTPArticleValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE')), 'GA_QUALIFUNITEVTE')]);
            {BCO_FACTURABLE}    19 : FieldValue := Format('''%s''', ['N']);
            {BCO_FAMILLETAXE1}  20 : FieldValue := Format('''%s''', [GetBTPArticleValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE')), 'GA_FAMILLETAXE1')]);
            {BCO_DATETRAVAUX}   21 : FieldValue := Format('''%s''', [Tools.CastDateTimeForQry(TobDataL.GetDateTime(Tools.iif(IsHSA, 'HSA_DATEMVT', 'COS_DATEMVT')))]);
            {BCO_FAMILLENIV1}   22 : FieldValue := Format('''%s''', [GetBTPArticleValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE')), 'GA_FAMILLENIV1')]);
            {BCO_FAMILLENIV2}   23 : FieldValue := Format('''%s''', [GetBTPArticleValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE')), 'GA_FAMILLENIV2')]);
            {BCO_FAMILLENIV3}   24 : FieldValue := Format('''%s''', [GetBTPArticleValue(TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE')), 'GA_FAMILLENIV3')]);
            {BCO_TRAITEVENTE}   25 : FieldValue := Format('''%s''', ['-']);                            
            {BCO_MONTANTACH}    26 : FieldValue := Format('%s'    , [GetDoubleValue(Tools.iif(IsHSA, GetBcoQteFromHSA(TobDataL.GetInteger('HSA_NBHRS')), TobDataL.GetDouble('COS_QTE')) * TobDataL.GetDouble(Tools.iif(IsHSA, 'HSA_DPA', 'COS_DPA')))]);
            {BCO_MONTANTPR}     27 : FieldValue := Format('%s'    , [GetDoubleValue(Tools.iif(IsHSA, GetBcoQteFromHSA(TobDataL.GetInteger('HSA_NBHRS')), TobDataL.GetDouble('COS_QTE')) * TobDataL.GetDouble(Tools.iif(IsHSA, 'HSA_DPR', 'COS_DPR')))]);
            {BCO_MONTANTHT}     28 : FieldValue := Format('%s'    , [GetDoubleValue(Tools.iif(IsHSA, GetBcoQteFromHSA(TobDataL.GetInteger('HSA_NBHRS')), TobDataL.GetDouble('COS_QTE')) * TobDataL.GetDouble(Tools.iif(IsHSA, 'HSA_PUV', 'COS_PUV')))]);
            {BCO_RESSOURCE}     29 : FieldValue := Format('''%s''', [Tools.iif(IsHSA, TobDataL.GetString('HSA_CODESAL'), '')]);
            {BCO_NATUREPIECEG}  30 : FieldValue := Format('''%s''', [Tools.iif(IsHSA, Tools.iif(TobDataL.GetString('HSA_NUMDEVIS') <> '0', 'DBT', '')    , '')]);
            {BCO_SOUCHE}        31 : FieldValue := Format('''%s''', [Tools.iif(IsHSA, Tools.iif(TobDataL.GetString('HSA_NUMDEVIS') <> '0', DBTSouche, ''), '')]);
            {BCO_NUMERO}        32 : FieldValue := Format('''%s''', [Tools.iif(IsHSA, TobDataL.GetString('HSA_NUMDEVIS'), '')]);
            {BCO_INDICEG}       33 : FieldValue := Format('''%s''', [Tools.iif(IsHSA, '0', '')]);
            {BCO_FROMDEVIS}     34 : FieldValue := Format('''%s''', [Tools.iif(IsHSA, Tools.iif(TobDataL.GetString('HSA_NUMDEVIS') <> '0', 'X', '-'), '')]);
          else
            FieldValue := GetBTPDefaultValue(FieldName);
          end;
          { Cas du chantier ACDC, test s'il existe }
          if     (IsHSA)
//             and (FieldName = 'BCO_AFFAIRE1')
             and (TobDataL.GetString('HSA_CODECHA') = WSCDS_CodeAFFAtt)
          then
          begin
            AFFCode := GetBTPAffaireValue(WSCDS_CodeAFFAtt, 'AFF_AFFAIRE');
            if (AFFCode = '') or (AFFCode = WSCDS_ErrorMsg) then
            begin
              Result := Format('%s=HSA_CODECHA;%s (HSA_NUMDEVIS=%s, HSA_BU=%s)', [WSCDS_ErrorMsg, WSCDS_CodeAFFAtt, TobDataL.GetString('HSA_NUMDEVIS'), TobDataL.GetString('HSA_BU')]);
              break;
            end;
          end;
          if (FieldValue = '''''') then
          begin
            { Gestion des erreurs }
            OnError := (FieldName = 'BCO_ARTICLE');
            if not OnError then
              OnError := ((not IsHSA) and (FieldName = 'BCO_AFFAIRE'));
            if     (not OnError)
               and (IsHSA)
               and (TobDataL.GetString('HSA_CODECHA') = '0')
               and (TobDataL.GetString('HSA_NUMDEVIS') = '0')
               and (TobDataL.GetString('HSA_BU')       = '')
            then
              OnError := True;
            if OnError then
            begin
              if FieldName = 'BCO_AFFAIRE' then
                Result := Format('%s=%s;%s', [WSCDS_ErrorMsg, Tools.iif(IsHSA, 'HSA_CODECHA', 'COS_CODECHA')           , TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODECHA', 'COS_CODECHA'))])
              else if FieldName = 'BCO_ARTICLE' then
                Result := Format('%s=%s;%s', [WSCDS_ErrorMsg, Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE'), TobDataL.GetString(Tools.iif(IsHSA, 'HSA_CODEPRESTATION', 'COS_CODEARTICLE'))])
              ;
              break;
            end;
          end;
          Result := Format('%s, %s', [Result, FieldValue]);
        end;
      end;
    end;
  end;
  if copy(Result , 1, length(WSCDS_ErrorMsg)) <> WSCDS_ErrorMsg then
    Result := Copy(Result, 2, length(Result));
end;

function TImportTreatment.GetBTPDefaultValue(FieldName : string) : string;
var
  FieldType : tTypeField;
begin
  FieldType  := Tools.GetFieldType(FieldName, AdoQryBTP.ServerName, AdoQryBTP.DBName);
  case FieldType of
    ttfDate : Result := Format('''%s''', [Tools.CastDateTimeForQry(IDate1900)]);
  else
    Result := Tools.GetDefaultValueFromtTypeField(FieldType);
  end;
end;

function TImportTreatment.GetBTPFieldsList(sType : string='') : string;
var
  TypeParam : string;
begin
  Result := '';
  case Tn of
    tnParameters :
      begin
        TypeParam := GetParamTableName(sType);
        if TypeParam <> '' then
        begin
          case Tools.CaseFromString(TypeParam, ['CHOIXCOD', 'CHOIXEXT', 'FONCTION']) of
            {CHOIXCOD} 0 : Result := ParametresCCFieldsList;
            {CHOIXEXT} 1 : Result := ParametresYXFieldsList;
            {FONCTION} 2 : Result := ParametresFOFieldsList;
          end;
        end;
      end;
    tnArticles     : Result := ArticlesFieldsList;  
    tnSalaries     : Result := SalariesFieldsList;
    tnModeRegle    : Result := ModeRegleFieldsList;
    tnReglement    : Result := ReglementFieldsList;
    tnHresSalaries : Result := HresSalariesFieldsList;
    tnConsoStock   : Result := ConsoStockFieldsList; 
  end;
end;

function TImportTreatment.IsBTPParamPrefixExist(TobDataL : TOB) : boolean;
var
  iIndex : integer;
begin
  Result := True;
  case Tn of
    tnParameters :
      begin
        iIndex := TslCacheTypeParam.IndexOfName(TobDataL.GetString('PAR_TYPE'));
        if iIndex = -1 then
        begin
          AdoQryBTP.FieldsList := 'DO_COMBO';
          AdoQryBTP.Request    := Format('SELECT %s FROM DECOMBOS WHERE DO_TYPE = ''%s''', [AdoQryBTP.FieldsList, TobDataL.GetString('PAR_TYPE')]);
          AdoQryBTP.SingleTableSelect;
          Result := (AdoQryBTP.RecordCount = 1);
          if Result then
            TslCacheTypeParam.Add(Format('%s=X', [TobDataL.GetString('PAR_TYPE')]));
          AdoQryBTP.Reset;
        end;
      end;
  end;
end;

function TImportTreatment.IsBTPRecordExist(TobDataL : TOB) : boolean;
var
  TypeParam : string;
begin
  case Tn of
    tnParameters :
      begin
        TypeParam := GetParamTableName(TobDataL.GetString('PAR_TYPE'));
        case Tools.CaseFromString(TypeParam, ['CHOIXCOD', 'CHOIXEXT', 'FONCTION']) of
          {CHOIXCOD} 0 : AdoQryBTP.FieldsList := 'CC_TYPE,CC_CODE';
          {CHOIXEXT} 1 : AdoQryBTP.FieldsList := 'YX_TYPE,YX_CODE';
          {FONCTION} 2 : AdoQryBTP.FieldsList := 'AFO_FONCTION';
        else
          AdoQryBTP.FieldsList := '';
        end;
      end;
    tnArticles     : AdoQryBTP.FieldsList := 'GA_ARTICLE';
    tnSalaries     : AdoQryBTP.FieldsList := 'ARS_RESSOURCE';
    tnModeRegle    : AdoQryBTP.FieldsList := 'MP_MODEPAIE';
    tnReglement    : AdoQryBTP.FieldsList := 'GAC_REFORIGINE';
    tnHresSalaries : AdoQryBTP.FieldsList := 'BCO_NUMMOUV';
  end;
  if AdoQryBTP.FieldsList <> '' then
  begin
    AdoQryBTP.Request := Format('SELECT %s FROM %s WHERE %s'
                             , [  AdoQryBTP.FieldsList
                                , GetBTPTableName(Tn, Tools.iif(Tn=tnParameters, TobDataL.GetString('PAR_TYPE'), ''))
                                , GetBTPWhere(TobDataL)
                               ]);
    AdoQryBTP.SingleTableSelect;
    Result := (AdoQryBTP.RecordCount = 1);
    AdoQryBTP.Reset;
  end else
    Result := False;
end;
  
function TImportTreatment.IsOkPaymentData(TobDataL : TOB) : string;
var
  iIndex     : integer;
  FieldValue : string;
  Error      : boolean;

  function DataExist(TMPFieldName, BTPFieldName, PrefixCache, TxtMsg : string; ltTableName : tTableName) : string;
  begin
    FieldValue := TobDataL.GetString(TMPFieldName);
    iIndex     := TslCacheBTPData.IndexOfName(PrefixCache + FieldValue);
    if iIndex = -1  then
    begin
      AdoQryBTP.FieldsList := BTPFieldName;
      AdoQryBTP.Request    := Format('SELECT %s FROM %s WHERE %s = ''%s''', [AdoQryBTP.FieldsList, Tools.GetTableNameFromTtn(ltTableName), AdoQryBTP.FieldsList, FieldValue]);
      AdoQryBTP.SingleTableSelect;
      Error := (AdoQryBTP.RecordCount = 0);
      TslCacheBTPData.Add(Format('%s=%s', [PrefixCache + FieldValue, Tools.iif(Error, '-', 'X')]));
      AdoQryBTP.Reset;
    end else
      Error := (TslCacheBTPData.ValueFromIndex[iIndex] = '-');
    if Error then
      Result := Format('%s %s n''existe pas.', [TxtMsg, FieldValue])
    else
      Result := '';
  end;

begin
  Result := '';
  Result := DataExist('EDR_JOURNAL', 'J_JOURNAL', PrefixCache_GAC_JAL, 'Le journal', ttnJournal);                                                           // Test le code journal
  Result := Result + Tools.iif(Result = '', '', ' - ') + DataExist('EDR_MODEPAIE', 'MP_MODEPAIE', PrefixCache_GAC_JAL, 'Le mode de paiement', ttnModePaie); // Test le mode de paiement
  if TobDataL.GetString('EDR_DEVISE') <> '' then
    Result := Result + Tools.iif(Result = '', '', ' - ') + DataExist('EDR_DEVISE', 'D_DEVISE', PrefixCache_GAC_JAL, 'La devise', ttnDevise);                // Test la devise
end;

function TImportTreatment.GetUpdateXX_Traite(TobDataL : TOB) : string;

  function GetSql(IsUpdate : boolean; Prefix : string) : string;
  var
    IdFieldName : string;
    Where       : string;
  begin
    IdFieldName := Format('%s_ID', [Prefix]);
    Where       := Format('%s = %s', [IdFieldName, TobDataL.GetString(IdFieldName)]);
    if IsUpdate then
      Result := Format('UPDATE %s SET %s_TRAITE = 1, %s_DATETRAITE = ''%s'' WHERE %s', [TUtilBTPVerdon.GetTMPTableName(Tn), Prefix, Prefix, Tools.CastDateTimeForQry(NowH), Where])
    else
      Result := Format('DELETE FROM %s WHERE %s', [TUtilBTPVerdon.GetTMPTableName(Tn), Where]);
  end;

begin
  case Tn of
    tnParameters   : Result := GetSql(ParametresValues.KeepData  , 'PAR');
    tnArticles     : Result := GetSql(ArticlesValues.KeepData    , 'ART');
    tnSalaries     : Result := GetSql(SalariesValues.KeepData    , 'SAL');
    tnModeRegle    : Result := GetSql(ModeRegleValues.KeepData   , 'MDR');
    tnReglement    : Result := GetSql(ReglementValues.KeepData   , 'EDR');
    tnHresSalaries : Result := GetSql(HresSalariesValues.KeepData, 'HSA');
    tnConsoStock   : Result := GetSql(ConsoStockValues.KeepData  , 'COS');
  else
    Result := '';
  end;
end;

function TImportTreatment.AddPayment(TobDataL : TOB) : boolean;
var
  iIndex      : integer;
  DocRef      : string;
  TobPaymentL : TOB;
begin
  Result := True;
  iIndex := TslCacheBTPData.IndexOfName(PrefixCache_GAC_REF + TobDataL.GetString('EDR_IDENTIFIANT'));
  if iIndex > -1 then
  begin
    DocRef      := TslCacheBTPData.ValueFromIndex[iIndex];
    TobPaymentL := TobPayment.FindFirst(['DOCREF'], [DocRef], true);
    if not assigned(TobPaymentL) then
    begin
      TobPaymentL := TOB.Create('', TobPayment, -1);
      TobPaymentL.AddChampSupValeur('DOCREF'   , DocRef);
      TobPaymentL.AddChampSupValeur('AMOUNT'   , Valeur(TobDataL.GetString('EDR_MONTANT')));
      TobPaymentL.AddChampSupValeur('CURAMOUNT', Valeur(TobDataL.GetString('EDR_MONTANTDEV')));
    end else
    begin
      TobPaymentL.SetDouble('AMOUNT'   , (TobPaymentL.GetDouble('AMOUNT')    + Valeur(TobDataL.GetString('EDR_MONTANT'))));
      TobPaymentL.SetDouble('CURAMOUNT', (TobPaymentL.GetDouble('CURAMOUNT') + Valeur(TobDataL.GetString('EDR_MONTANTDEV'))));
    end;
  end;
end;

function TImportTreatment.DoUpdateDocFromPayment : boolean;
var
  Cpt         : integer;
  Cpt1        : integer;
  DecQty      : integer;
  TobPaymentL : TOB;
  TobTerms    : TOB;
  TobTermsL   : TOB;
  DocRef      : string;
  Nature      : string;
  Souche      : string;
  Numero      : string;
  Indice      : string;
  Data        : string;
  DocumentMtC : double;
  PaymentMt   : double;
  PaymentMtC  : double;
  TermsMt     : double;
  TermsMtC    : double;
  RgMt        : double;
  CollectRgMt : double;
  Coefficient : double;
  FieldsList  : array of string;

  function GetWhere(Prefix : string) : string;
  begin
    Result := Format('     %s_NATUREPIECEG = ''%s'''
                   + ' AND %s_SOUCHE       = ''%s'''
                   + ' AND %s_NUMERO       = %s'
                   + ' AND %s_INDICEG      = %s'
                     , [Prefix, Nature, Prefix, Souche, Prefix, Numero, Prefix, Indice]);
  end;

begin
  Result := True;
  DecQty := StrToInt(Tools.GetParamSocSecur_('SO_DECVALEUR', '2', AdoQryBTP.ServerName, AdoQryBTP.DBName));
  for Cpt := 0 to pred(TobPayment.detail.count) do
  begin
    TobPaymentL := TobPayment.detail[Cpt];
    DocRef      := TobPaymentL.GetString('DOCREF');
    Nature      := Tools.ReadTokenSt_(DocRef, '-');
    Souche      := Tools.ReadTokenSt_(DocRef, '-');
    Numero      := Tools.ReadTokenSt_(DocRef, '-');
    Indice      := Tools.ReadTokenSt_(DocRef, '-');
    DocumentMtC := 0;
    PaymentMt   := 0;
    PaymentMtC  := 0;
    RgMt        := 0;
    CollectRgMt := 0;
    { Recherche le montant des acomptes }
    try
      AdoQryBTP.FieldsList := 'GAC_MONTANT, GAC_MONTANTDEV';
      AdoQryBTP.Request    := Format('SELECT %s FROM ACOMPTES WHERE %s', [AdoQryBTP.FieldsList, GetWhere('GAC')]);
      AdoQryBTP.SingleTableSelect;
      Result := (AdoQryBTP.RecordCount > 0);
      if Result then
      begin
        for Cpt1 := 0 to pred(AdoQryBTP.TSLResult.Count) do
        begin
          Data       := AdoQryBTP.TSLResult[Cpt1];
          PaymentMt  := PaymentMt  + Valeur(Tools.ReadTokenSt_(Data, '^'));
          PaymentMtC := PaymentMtC + Valeur(Tools.ReadTokenSt_(Data, '^'));
        end;
      end;
    finally
      AdoQryBTP.Reset;
    end;
    { Mise à jour du montant acompte de la pièce }
    if Result then
    begin
      try
        AdoQryBTP.Request := Format('UPDATE PIECE SET GP_ACOMPTE = %s, GP_ACOMPTEDEV = %s WHERE %s', [FloatToStr(PaymentMt), FloatToStr(PaymentMtC), GetWhere('GP')]);
        AdoQryBTP.InsertUpdate;
        Result := (AdoQryBTP.RecordCount = 1);
      finally
        AdoQryBTP.Reset;
      end;
    end;
    { Recherche le TTC de la pièce }
    if Result then
    begin
      try
        AdoQryBTP.FieldsList := 'GP_TOTALTTCDEV';
        AdoQryBTP.Request    := Format('SELECT %s FROM PIECE WHERE %s', [AdoQryBTP.FieldsList, GetWhere('GP')]);
        AdoQryBTP.SingleTableSelect;
        Result := (AdoQryBTP.RecordCount > 0);
        if Result then
          DocumentMtC := Valeur(AdoQryBTP.TSLResult[0]);
      finally
        AdoQryBTP.Reset;
      end;
    end;
    { Recalcul le montant des échéances }
    if Result then
    begin
      try
        AdoQryBTP.FieldsList := 'GPE_NATUREPIECEG,GPE_SOUCHE,GPE_NUMERO,GPE_INDICEG,GPE_NUMECHE,GPE_MONTANTECHE,GPE_MONTANTDEV';
        AdoQryBTP.Request    := Format('SELECT %s FROM PIEDECHE WHERE %s', [AdoQryBTP.FieldsList, GetWhere('GPE')]);
        AdoQryBTP.SingleTableSelect;
        if AdoQryBTP.RecordCount > 0 then
        begin
          TobTerms := TOB.Create('_TERMS', nil, -1);
          try
            SetLength(FieldsList, 7);
            FieldsList[0] := 'GPE_NATUREPIECEG';
            FieldsList[1] := 'GPE_SOUCHE';
            FieldsList[2] := 'GPE_NUMERO';
            FieldsList[3] := 'GPE_INDICEG';
            FieldsList[4] := 'GPE_NUMECHE';
            FieldsList[5] := 'GPE_MONTANTECHE';
            FieldsList[6] := 'GPE_MONTANTDEV';
            Tools.TStringListToTOB(AdoQryBTP.TSLResult, FieldsList, TobTerms, False);
            { Calcul somme des échéances }
            TermsMt  := 0;
            TermsMtC := 0;
            for Cpt1 := 0 to pred(TobTerms.detail.count) do
            begin
              TobTermsL := TobTerms.Detail[Cpt1];
              TermsMt  := TermsMt  + TobTermsL.GetDouble('GPE_MONTANTECHE');
              TermsMtC := TermsMtC + TobTermsL.GetDouble('GPE_MONTANTDEV');
            end;
            Coefficient := (DocumentMtC - PaymentMtC) / TermsMtC;
            if Arrondi(Coefficient - 1, 9) <> 0 then
            begin
              for Cpt1 := 0 to pred(TobTerms.detail.count) do
              begin
                TobTermsL := TobTerms.Detail[Cpt1];
                TermsMt  := Arrondi(TobTermsL.GetDouble('GPE_MONTANTECHE') * Coefficient, DecQty);
                TermsMtC := Arrondi(TobTermsL.GetDouble('GPE_MONTANTDEV')  * Coefficient, DecQty);
                try
                  AdoQryBTP.Request := Format('UPDATE PIEDECHE SET GPE_MONTANTECHE = %s, GPE_MONTANTDEV = %s WHERE %s AND GPE_NUMECHE = %s'
                                           , [  StringReplace(FloatToStr(TermsMt) , ',', '.', [rfReplaceAll])
                                              , StringReplace(FloatToStr(TermsMtC), ',', '.', [rfReplaceAll])
                                              , GetWhere('GPE')
                                              , TobTermsL.GetString('GPE_NUMECHE')]);
                  AdoQryBTP.InsertUpdate;
                  Result := (AdoQryBTP.RecordCount = 1);
                finally
                  AdoQryBTP.Reset;                              
                end;
              end;
            end;
          finally
            FreeAndNil(TobTerms);
          end;
        end;
      finally
        AdoQryBTP.Reset;
      end;
    end;
  end;
end;
  
function TImportTreatment.TableProcess : boolean;
var
  Cpt         : integer;
  Cpt1        : integer;
  Qty         : integer;
  TobDataL    : TOB;
  TableName   : string;
  ErrorMsg    : string;
  FieldName   : string;
  FieldValue  : string;
  CanContinue : boolean;
  Exist       : boolean;

  function GetArrValue(Pos : integer) : string;
  begin
    Result := '';
    case Tn of
      tnParameters :
      begin
        case Pos of
          1 : Result := TobDataL.GetString('PAR_TYPE');
          2 : Result := TobDataL.GetString('PAR_CODE');
          3 : Result := TobDataL.GetString('PAR_LIBELLE');
        end;
      end;
      tnArticles :
      begin
        case Pos of
          1 : Result := TobDataL.GetString('ART_CODEARTICLE');
          2 : Result := TobDataL.GetString('ART_LIBELLE');
          3 : Result := TobDataL.GetString('ART_NATUREPRESTATION');
        end;
      end;
      tnSalaries :
      begin
        case Pos of
          1 : Result := TobDataL.GetString('SAL_CODESAL');
          2 : Result := TobDataL.GetString('SAL_NOM');
          3 : Result := TobDataL.GetString('SAL_PRENOM');
        end;
      end;
      tnModeRegle :
      begin
        case Pos of
          1 : Result := TobDataL.GetString('MDR_CODE');
          2 : Result := TobDataL.GetString('MDR_LIBELLE');
          3 : Result := TobDataL.GetString('MDR_CATEGORIE');
        end;
      end;
      tnReglement :
      begin
        case Pos of
          1 : Result := TobDataL.GetString('EDR_IDENTIFIANT');
          2 : Result := TobDataL.GetString('EDR_CODETIERS');
          3 : Result := TobDataL.GetString('EDR_LIBELLE');
        end;
      end;
      tnHresSalaries :
      begin
        case Pos of
          1 : Result := Format('HSA_IDENTIFIANT=%s', [TobDataL.GetString('HSA_IDENTIFIANT')]);
          2 : begin
                if      TobDataL.GetString('HSA_CODECHA')  <> '0' then Result := Format('HSA_CODECHA=%s' , [TobDataL.GetString('HSA_CODECHA')])
                else if TobDataL.GetString('HSA_NUMDEVIS') <> '0' then Result := Format('HSA_NUMDEVIS=%s', [TobDataL.GetString('HSA_NUMDEVIS')])
                else if TobDataL.GetString('HSA_BU')       <> ''  then Result := Format('HSA_BU=%s'      , [TobDataL.GetString('HSA_BU')])
                else                                                       Result := '';
              end;
          3 : Result := Format('HSA_CODESAL=%s', [TobDataL.GetString('HSA_CODESAL')]);
        end;
      end;
      tnConsoStock :
      begin
        case Pos of
          1 : Result := TobDataL.GetString('COS_DATEMVT');
          2 : Result := TobDataL.GetString('COS_CODECHA');
          3 : Result := TobDataL.GetString('COS_CODEARTICLE');
        end;
      end;
    end;
  end;

begin
  Result        := True;
  NumDebugEvent := 0;
  Qty           := 0;
  SocietyCode   := GetParamSocSecur('SO_SOCIETE', '');
  AdoQryBTP.FieldsList := 'GPP_SOUCHE';
  AdoQryBTP.Request    := 'SELECT ' + AdoQryBTP.FieldsList + ' FROM PARPIECE WHERE GPP_NATUREPIECEG = ''DBT''';
  AdoQryBTP.SingleTableSelect;
  DBTSouche := AdoQryBTP.TSLResult[0];
  AdoQryBTP.Reset;
  TUtilBTPVerdon.StartLog(ServiceName_BTPVerdonImp, Tn, LogValues, ParametresValues.LastSynchro);
  TobInsertUpdate.ClearDetail;
  TobData.ClearDetail;
  AdoQryTMP.FieldsList := GetTMPFieldsList;
  AdoQryTMP.Request    := GetDataSearchSql(LastSynchro);
  AddDebugEvent(Format('Sql = %s', [AdoQryTMP.Request]));
  AdoQryTMP.SingleTableSelect;
  AdoQryTMP.FieldsList := '';
  AddDebugEvent(Format('AdoQryTMP.RecordCount = %s', [IntToStr(AdoQryTMP.RecordCount)]));
  Tools.TStringListToTOB(AdoQryTMP.TSLResult, TMPArrFields, TobData, True);
  AdoQryTMP.Reset;
  AddDebugEvent(Format('TobData.Detail.Count = %s', [IntToStr(TobData.Detail.Count)]));
  if TobData.Detail.Count > 0 then
  begin
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('Recherche des données (%s enregistrement(s) de la table %s)', [IntToStr(TobData.Detail.Count), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
    for Cpt := 0 to pred(TobData.Detail.count) do
    begin
      TobDataL  := TobData.Detail[Cpt];
      TableName := GetBTPTableName(Tn, Tools.iif(Tn = tnParameters, TobDataL.GetString('PAR_TYPE'), ''));
      if TableName <> '' then
      begin
        if Tn = tnParameters then // Si Paramètres, test si le préfixe existe dans CHOIXCOD ou CHOIXEXT
          CanContinue := IsBTPParamPrefixExist(TobDataL)
        else
          CanContinue := True;
        if CanContinue then
        begin
          if Tn = tnReglement then
            CanContinue := IsLinkedDocExist(TobDataL.GetString('EDR_IDENTIFIANT'), TobDataL.GetString('EDR_REFERPLSE')) // Pour les règlements, il faut tester si la pièce associée existe
          else
            CanContinue := True;
          if CanContinue then
          begin
            if Tn = tnReglement then
              ErrorMsg  := IsOkPaymentData(TobDataL)
            else
              ErrorMsg := '';
            if ErrorMsg = '' then
            begin
              inc(Qty);
              { Transforme les données en valeurs correctes pour insert ou update }
(*
              for Cpt1 := 0 to pred(TobDataL.NombreChampSup) do
              begin
                FieldName  := TobDataL.GetNomChamp(1000+Cpt1);
                FieldValue := GetValue(FieldName, TobDataL.GetValeur(1000+Cpt1), False);
                TobDataL.SetString(FieldName, FieldValue);
              end;
*)
              { Affectle chantier ACDC si chantier est vide est devis ou bu renseigné, sinon, laisse vide pour gérer une erreur }
              if     (Tn = tnHresSalaries)
                 and (TobDataL.GetString('HSA_CODECHA') = '0')
                 and (   (TobDataL.GetString('HSA_NUMDEVIS') <> '0')
                      or (TobDataL.GetString('HSA_BU')       <> ''))
              then
                TobDataL.SetString('HSA_CODECHA', WSCDS_CodeAFFAtt);
              Exist := IsBTPRecordExist(TobDataL);
             if Exist then
                AdoQryBTP.Request := Format('UPDATE %s SET %s WHERE %s', [TableName, GetBTPUpdateValues(TobDataL), GetBTPWhere(TobDataL)])
              else
                AdoQryBTP.Request := Format('INSERT INTO %s (%s) VALUES (%s)', [TableName, GetBTPFieldsList(Tools.iif(Tn = tnParameters, TobDataL.GetString('PAR_TYPE'), '')), GetBTPInsertValues(TobDataL)]);
              { Il y a une erreur }
              if pos(WSCDS_ErrorMsg, AdoQryBTP.Request) > 0 then
              begin
                ErrorMsg := copy(AdoQryBTP.Request, pos('=', AdoQryBTP.Request)+1, length(AdoQryBTP.Request));
                ErrorMsg := copy(ErrorMsg, 1, length(ErrorMsg)-1);
                TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('%s - La donnée %s=%s est inexistante.', [WSCDS_ErrorMsg, Tools.ReadTokenSt_(ErrorMsg, ';'), Tools.ReadTokenSt_(ErrorMsg, ';')]), LogValues, 2);
              end else
              begin
                if LogValues.LogLevel = 2 then
                  TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('%s de %s - %s - %s', [Tools.iif(Exist, 'Mise à jour', 'Création'), GetArrValue(1), GetArrValue(2), GetArrValue(3)]), LogValues, 2);
                AddDebugEvent(Format('AdoQryBTP.Request = %s', [AdoQryBTP.Request]));
                if AdoQryBTP.Request <> ''  then
                begin
                  AdoQryBTP.InsertUpdate;
                  Result := (AdoQryBTP.RecordCount = 1);
                  AdoQryBTP.Reset;
                  if (Result) and (Tn = tnReglement) then // Si règlement, il faut faire ajouter en cache le cumul des règlement pour une même pièce
                    AddPayment(TobDataL);
                  if Result then
                  begin
                    AdoQryTMP.Request := GetUpdateXX_Traite(TobDataL);
                    AddDebugEvent(Format('AdoQryTMP.Request = %s', [AdoQryTMP.Request]));
                    AdoQryTMP.InsertUpdate;
                    Result := (AdoQryTMP.RecordCount = 1);
                    AdoQryTMP.Reset;
                  end;
                  if not Result then
                    Break;
                end;
              end;
            end else
              TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('%s - %s', [WSCDS_ErrorMsg, ErrorMsg]), LogValues, 1);
          end else
            TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('%s - La pièce référence %s associée au règlement %s n''existe pas.'
                                                                       , [  WSCDS_ErrorMsg
                                                                          , TobDataL.GetString('EDR_REFERPLSE')
                                                                          , TobDataL.GetString('EDR_IDENTIFIANT')
                                                                         ]), LogValues, 1);
        end else
          TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('%s - Le préfixe %s (code %s) n''existe pas dans la table %s.'
                                                                     , [  WSCDS_ErrorMsg
                                                                        , TobDataL.GetString('PAR_TYPE')
                                                                        , TobDataL.GetString('PAR_CODE')
                                                                        , TableName
                                                                       ]), LogValues, 1);
      end else
        TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('%s - Le préfixe %s (code %s) n''est pas présent dans la section "IMPORT_PARAMETERSTYPEMATCHING" du fichier de configuration %s.ini.'
                                                                   , [  WSCDS_ErrorMsg
                                                                      , TobDataL.GetString('PAR_TYPE')
                                                                      , TobDataL.GetString('PAR_CODE')
                                                                      , ServiceName_BTPVerdonIniFile
                                                                     ]), LogValues, 1);
    end;
    if (Result) and (Tn = tnReglement) then
      Result := DoUpdateDocFromPayment;
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('%s enregistrement(s) de la table %s traité(s).)', [IntToStr(Qty), TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
  end else
    TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, Format('Aucun %s n''a été trouvé.', [TUtilBTPVerdon.GetTMPTableName(Tn)]), LogValues, 1);
  TobData.ClearDetail;
  if Result then
    SetLastSynchro;
  TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonImp, Tn, TUtilBTPVerdon.GetMsgStartEnd(Tn, False, ParametresValues.LastSynchro), LogValues, 0);
end;

constructor TImportTreatment.create;
begin
  TslParametersTypeMatching := TStringList.Create;
  TslCacheTypeParam         := TStringList.Create;
  TslCacheBTPData           := TStringList.Create;
  TslCacheAffaire           := TStringList.Create;
  TslCacheArticle           := TStringList.Create;
end;

destructor TImportTreatment.destroy;
begin
  inherited;
  FreeAndNil(TslParametersTypeMatching);
  FreeAndNil(TslCacheTypeParam);
  FreeAndNil(TslCacheBTPData);
  FreeAndNil(TslCacheAffaire);
  FreeAndNil(TslCacheArticle);
  AdoQryBTP.Qry := nil;
  AdoQryTMP.Qry := nil;
end;

function TImportTreatment.ImportTreatment : boolean;
begin
  SetFieldsArray;
  TobData := TOB.Create('_DATA', nil, -1);
  try
    TobInsertUpdate := TOB.Create('_SQL', nil, -1);
    try
      TobPayment := TOB.create('_PAYMENT', nil, -1);
      try
        Result := TableProcess;
      finally
        FreeAndNil(TobPayment);
      end;
    finally
      FreeAndNil(TobInsertUpdate);
    end;
  finally
    FreeAndNil(TobData);
  end;
end;
{$ENDIF APPSRVWITHCBP}

{ TExportEcr }
{$IFNDEF APPSRV}
constructor TExportEcr.create;
begin
  TobEcrVen   := TOB.Create('_ECRITURE', nil, -1);
  TobEcrAch   := TOB.Create('_ECRITURE', nil, -1);
  TobEcrP     := TOB.Create('_ECRITURE', nil, -1);
  TobCustomer := TOB.Create('_CLIENT', nil, -1);
  TobProvider := TOB.Create('_FOURNISSEUR', nil, -1);
  Tobcountry  := TOB.Create('_PAYS', nil, -1);
  TslData     := TStringList.Create;
end;

destructor TExportEcr.destroy;
begin
  inherited;
  FreeAndNil(TobEcrP);
  FreeAndNil(TobEcrVen);
  FreeAndNil(TobEcrAch);
  FreeAndNil(TobCustomer);
  FreeAndNil(TobProvider);
  FreeAndNil(Tobcountry);
  FreeAndNil(TslData);
end;

function TExportEcr.GetCountryIso2(Country : string) : string;
var
  TobL : TOB;
begin
  Result := '';
  if Country <> '' then
  begin
    TobL := TobCountry.FindFirst(['PY_PAYS'], [Country], True);
    if assigned(TobL) then
      Result := TobL.GetString('PY_CODEISO2');
  end;
end;

function TExportEcr.LoadEcr : boolean;
var
  TobEcr : TOB;
  Sql    : string;

  procedure SearchThird(Flow, AuxiliaryCode : string);
  var
    TobT  : TOB;
    TobL  : TOB;
    TobL1 : TOB;
    Sql   : string;
    Qry   : TQuery;
  begin
    if not assigned(Tools.iif(Flow = 'VTE', TobCustomer, TobProvider).FindFirst(['T_AUXILIAIRE'], [AuxiliaryCode], True)) then
    begin
      TobT := TOB.Create('TIERS', Tools.iif(Flow = 'VTE', TobCustomer, TobProvider), -1);
      Sql  := Format('SELECT TIERS.*'
                   + '     , YTC_TEXTELIBRE1'
                   + '     , US_UTILISATEUR' 
                   + ' FROM TIERS'
                   + ' LEFT JOIN TIERSCOMPL ON YTC_AUXILIAIRE = T_AUXILIAIRE'
                   + ' LEFT JOIN COMMERCIAL ON GCL_COMMERCIAL = T_REPRESENTANT'
                   + ' LEFT JOIN UTILISAT   ON US_UTILISATEUR = GCL_UTILASSOCIE'
                   + ' WHERE T_AUXILIAIRE = "%s"', [AuxiliaryCode]);
      Qry  := OpenSql(Sql, True);
      try
        if not Qry.Eof then
        begin
          TobT.SelectDB('', Qry);
          TslData.Add(Format('- %s : %s - %s', [Tools.iif(TobT.GetString('T_NATUREAUXI') = 'CLI', 'Client', 'Fournisseur'), TobT.GetString('T_TIERS'), TobT.GetString('T_LIBELLE')]));
          { Recherche l'adresse de facturation }
          TobL := TOB.Create('_ADRESSES', TobT, -1);
          TobL.AddChampSupValeur('TYPE', 'ADRESSES');
          Sql  := Format('SELECT TOP 1 * FROM ADRESSES WHERE ADR_REFCODE = "%s" AND ADR_FACT = "X"', [TobT.GetString('T_TIERS')]);
          TobL.LoadDetailFromSQL(Sql);
          if TobL.detail.count = 0 then // Pas d'adresse de facturation, prend l'adresse de la fiche tiers
          begin
            TobL1 := TOB.Create('ADRESSES', TobL, -1);
            TobL1.InitValeurs;
            TobL1.SetString('ADR_ADRESSE1'  , TobT.GetString('T_ADRESSE1'));
            TobL1.SetString('ADR_ADRESSE2'  , TobT.GetString('T_ADRESSE2'));
            TobL1.SetString('ADR_ADRESSE3'  , TobT.GetString('T_ADRESSE3'));
            TobL1.SetString('ADR_CODEPOSTAL', TobT.GetString('T_CODEPOSTAL'));
            TobL1.SetString('ADR_VILLE'     , TobT.GetString('T_VILLE'));
            TobL1.SetString('ADR_REGION'    , TobT.GetString('T_REGION'));
            TobL1.SetString('ADR_PAYS'      , TobT.GetString('T_PAYS'));
          end;
          { Recherche les contacts }
          TobL := TOB.Create('_CONTACT', TobT, -1);
          TobL.AddChampSupValeur('TYPE', 'CONTACT');
          Sql  := Format('SELECT * FROM CONTACT WHERE C_TIERS = "%s" ORDER BY C_PRINCIPAL DESC', [TobT.GetString('T_TIERS')]);
          TobL.LoadDetailFromSQL(Sql);
          { Recherche les RIB }
          TobL := TOB.Create('_RIB', TobT, -1);
          TobL.AddChampSupValeur('TYPE', 'RIB');
          Sql  := Format('SELECT * FROM RIB WHERE R_AUXILIAIRE = "%s" ORDER BY R_PRINCIPAL DESC', [AuxiliaryCode]);
          TobL.LoadDetailFromSQL(Sql);
        end;
      finally
        Ferme(Qry);
      end;
    end;
  end;

begin
  TobEcr := TOB.Create('_ECR', nil, -1);
  try
    Sql := Format('SELECT J_NATUREJAL'
                + '     , G_LIBELLE'
                + '     , ECRITURE.*'
                + ' FROM ECRITURE'
                + ' JOIN JOURNAL  ON J_JOURNAL = E_JOURNAL AND J_NATUREJAL IN ("VTE", "ACH")'
                + ' JOIN GENERAUX ON G_GENERAL = E_GENERAL'
                + ' WHERE E_DATECOMPTABLE BETWEEN "%s" AND "%s"'
                + Tools.iif(ForceExport, '', ' AND E_EXPORTE = "---"')
                + ' ORDER BY E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE'
                  , [Tools.UsDateTime_(StartDate), Tools.UsDateTime_(EndDate)]);
    TobEcr.LoadDetailFromSQL(Sql);
    Result := (TobEcr.Detail.Count > 0);
    if Result then
    begin
      { Eclate Vente et Achat}
      repeat
        if TobEcr.Detail[0].GetString('E_TYPEMVT') = 'TTC' then
          SearchThird(TobEcr.Detail[0].GetString('J_NATUREJAL'), TobEcr.Detail[0].GetString('E_AUXILIAIRE'));
        if TobEcr.Detail[0].GetString('J_NATUREJAL') = 'VTE' then
          TobEcr.Detail[0].ChangeParent(TobEcrVen, -1)
        else
          TobEcr.Detail[0].ChangeParent(TobEcrAch, -1);
      until TobEcr.Detail.count = 0
    end;
  finally
    FreeAndNil(TobEcr);
  end;
end;

function TExportEcr.LoadAnalytic : boolean;
var
  CptA : integer;
begin
  Result := True;
  if TobEcrP.Detail.Count > 0 then
  begin
    for CptA := 0 to pred(TobEcrP.Detail.Count) do
      LoadAnalyticOnTobEntry(TobEcrP.Detail[CptA]);
  end;
end;

function TExportEcr.LoadTax : boolean;
var
  CptT    : integer;
  TobL    : TOB;
  Sql     : string;
  AccRef  : string;
  DocType : string;
  DocSub  : string;
  DocNum  : integer;
  DocInd  : integer;
  Qry     : TQuery;
begin
  Result := True;
  if TobEcrP.Detail.Count > 0 then
  begin
    for CptT := 0 to pred(TobEcrP.Detail.Count) do
    begin
      TobL := TobEcrP.detail[CptT];
      if TobL.GetString('E_TYPEMVT') = 'HT' then
      begin
        AccRef  := TobL.GetString('E_REFGESCOM');
        DocType := Tools.ReadTokenSt_(AccRef, ';');
        DocSub  := Tools.ReadTokenSt_(AccRef, ';');
        Tools.ReadTokenSt_(AccRef, ';');
        DocNum  := StrToInt(Tools.ReadTokenSt_(AccRef, ';'));
        DocInd  := StrToInt(Tools.ReadTokenSt_(AccRef, ';'));
        Sql     := Format(  'SELECT BP4_BASEHT, BP4_MONTANTTAXE, BP4_CODETAXE, BP4_CPTTVA FROM BPIECEREPTVA'
                          + ' WHERE BP4_NATUREPIECEG = "%s"'
                          + '   AND BP4_SOUCHE       = "%s"'
                          + '   AND BP4_NUMERO       = %s'
                          + '   AND BP4_INDICEG      = %s'
                          + '   AND BP4_GENERAL      = "%s"'
                          + '   AND BP4_CODETAXE     = "%s"'
                          , [  DocType
                             , DocSub
                             , IntToStr(DocNum)
                             , IntToStr(DocInd)
                             , TobL.GetString('E_GENERAL')
                             , TobL.GetString('E_TVA')
                            ]
                         );
        Qry := OpenSql(Sql, True);
        try
          TobL.AddChampSupValeur('BP4_BASEHT'     , Qry.FindField('BP4_BASEHT').AsString);
          TobL.AddChampSupValeur('BP4_MONTANTTAXE', Qry.FindField('BP4_MONTANTTAXE').AsString);
          TobL.AddChampSupValeur('BP4_CODETAXE'   , Qry.FindField('BP4_CODETAXE').AsString);
          TobL.AddChampSupValeur('BP4_CPTTVA'     , Qry.FindField('BP4_CPTTVA').AsString);
        finally
          Qry.Free;
        end;
      end;
    end;
  end;
end;

function TExportEcr.DoubleLineFromTaxCode(TaxCode : string) : boolean;
begin
  case Tools.CaseFromString(TaxCode, [  'A'
                                      , 'A1', 'A3', 'A5'
                                      , 'B1', 'B3', 'B5'
                                      , 'F1', 'F3', 'F5'
                                      , 'G1', 'G3', 'G5'
                                     ]) of
    0..12 : Result := True;
  else
    Result := False;
  end;
end;


function TExportEcr.GetAmount(Amount : double; PositiveValue : boolean) : string;
var
  lAmount : double;
begin
  lAmount := Arrondi(Amount, 2) * Tools.iif(PositiveValue, 1, -1);
  Result  := FormatFloat('0.00', lAmount);
  Result := StringReplace(Result, ',', '.', [rfReplaceAll]);
end;

function TExportEcr.AddEntryToFile(GLEntries : IXMLNode) : boolean;
var
  Cpt        : integer;
  TobL       : TOB;
  AccDate    : string;
  ThirdType  : string;
  ThirdCode  : string;
  Payment    : string;
  IsInvoice  : boolean;
  IsSales    : boolean;
  GLEntry    : IXMLNode;
  SubLevel   : IXMLNode;
  SubLevel1  : IXMLNode;
  SubLevel2  : IXMLNode;

  function GetYourRef(Referency : string) : string;
  begin
    Result := Format('%s;%s', [Tools.ReadTokenSt_(Referency, ';'), Tools.ReadTokenSt_(Referency, ';')]);          // Nature + Souche
    Tools.ReadTokenSt_(Referency, ';');                                                                           // Enlève la date
    Result := Format(';%s;%s', [Result, Tools.ReadTokenSt_(Referency, ';'), Tools.ReadTokenSt_(Referency, ';')]); // Numéro + Indice
  end;

  procedure AddLine(Account, DebitAmount, CreditAmount : string);
  begin
    SubLevel := GLEntry.AddChild('FinEntryLine');              SubLevel.Attributes['number']  := IntToStr(LineNumber);
                                                               SubLevel.Attributes['type']    := 'N';
                                                               SubLevel.Attributes['subtype'] := Tools.iif(IsSales, 'K', 'T');
      SubLevel1 := SubLevel.AddChild('Date');                  SubLevel1.Text                 := AccDate;
      SubLevel1 := SubLevel.AddChild('GLAccount');             SubLevel1.Attributes['code']   := Account;
      SubLevel1 := SubLevel.AddChild('Description');           SubLevel1.Text                 := TobL.GetString('G_LIBELLE');
      SubLevel1 := SubLevel.AddChild('Costcenter');            SubLevel1.Attributes['code']   := '';
      SubLevel1 := SubLevel.AddChild('Costunit');              SubLevel1.Attributes['code']   := '';
      SubLevel1 := SubLevel.AddChild(ThirdType);               SubLevel1.Attributes['code']   := ThirdCode;
                                                               SubLevel1.Attributes['number'] := ThirdCode;
      SubLevel1 := SubLevel.AddChild('Resource');              SubLevel1.Attributes['number'] := TobL.GetString('E_UTILISATEUR');
      SubLevel1 := SubLevel.AddChild('Amount');
        Sublevel2 := SubLevel1.AddChild('Currency');           SubLevel2.Attributes['code']   := TobL.GetString('E_DEVISE');
        Sublevel2 := SubLevel1.AddChild('Debit');              SubLevel2.Text                 := DebitAmount;
        Sublevel2 := SubLevel1.AddChild('Credit');             SubLevel2.Text                 := CreditAmount;
        Sublevel2 := SubLevel1.AddChild('VAT');                SubLevel2.Attributes['code']   := TobL.GetString('E_TVA');
      SubLevel1 := SubLevel.AddChild('VATTransaction');        SubLevel1.Attributes['code']   := TobL.GetString('E_TVA');
        Sublevel2 := SubLevel1.AddChild('VATAmount');          SubLevel2.Text                 := GetAmount(TobL.GetDouble('BP4_MONTANTTAXE'), True);
        Sublevel2 := SubLevel1.AddChild('VATBaseAmount');      SubLevel2.Text                 := GetAmount(TobL.GetDouble('BP4_BASEHT'), True);
      SubLevel1 := SubLevel.AddChild('Payment');
        SubLevel2 := SubLevel1.AddChild('PaymentCondition');   SubLevel2.Attributes['code']   := Payment;
        SubLevel2 := SubLevel1.AddChild('Reference');          SubLevel2.Text                 := TobL.GetString('E_REFINTERNE');
        SubLevel2 := SubLevel1.AddChild('CSSDDate1');          SubLevel2.Text                 := AccDate;
        SubLevel2 := SubLevel1.AddChild('CSSDDate2');          SubLevel2.Text                 := AccDate;
      if IsSales then
      begin
        SubLevel1 := SubLevel.AddChild('Delivery');
          SubLevel2 := SubLevel1.AddChild('Date');               SubLevel2.Text               := AccDate;
      end;
      SubLevel1 := SubLevel.AddChild('FinReferences');         SubLevel1.Attributes['TransactionOrigin']   := 'N';
        SubLevel2 := SubLevel1.AddChild('UniquePostingNumber');SubLevel2.Text                 := '0';
        SubLevel2 := SubLevel1.AddChild('YourRef');            SubLevel2.Text                 := GetYourRef(TobL.GetString('E_REFGESCOM'));
        SubLevel2 := SubLevel1.AddChild('DocumentDate');       SubLevel2.Text                 := AccDate;
  end;

begin
  Result    := True;
  TobL      := TobEcrP.FindFirst(['E_TYPEMVT'], ['TTC'], True);
  IsInvoice := ((TobL.GetString('E_NATUREPIECE') = 'FC') or (TobL.GetString('E_NATUREPIECE') = 'FF'));
  IsSales   := (TobL.GetString('J_NATUREJAL') = 'VTE');
  Payment   := TobL.GetString('E_MODEPAIE');
  ThirdCode := TobL.GetString('E_AUXILIAIRE');
  AccDate   := FormatDateTime('yyyy-mm-dd', TobL.GetDateTime('E_DATECOMPTABLE'));
  ThirdType := Tools.iif(IsSales, 'Debtor', 'Creditor');
  GLEntry := GLEntries.AddChild('GLEntry');     GLEntry.Attributes['entry']  := TobL.GetString('E_NUMEROPIECE');
                                                GLEntry.Attributes['status'] := 'E';
  SubLevel := GLEntry.AddChild('Division');     SubLevel.Attributes['code']  := FolderCode;
  SubLevel := GLEntry.AddChild('Description');  SubLevel.Text                := TobL.GetString('E_LIBELLE');
  SubLevel := GLEntry.AddChild('Date');         SubLevel.Text                := AccDate;
  SubLevel := GLEntry.AddChild('DocumentDate'); SubLevel.Text                := AccDate;
  SubLevel := GLEntry.AddChild('Journal');      SubLevel.Attributes['code']  := TobL.GetString('E_JOURNAL');
  SubLevel := GLEntry.AddChild('Amount');
    SubLevel1 := SubLevel.AddChild('Currency'); SubLevel1.Attributes['code'] := TobL.GetString('E_DEVISE');
    SubLevel1 := SubLevel.AddChild('Value');    SubLevel1.Text               := GetAmount(TobL.GetDouble('E_DEBIT') + TobL.GetDouble('E_CREDIT'), IsInvoice);
  for Cpt := 0 to pred(TobEcrP.Detail.Count) do
  begin
    TobL := TobEcrP.Detail[Cpt];
    if TobL.GetString('E_TYPEMVT') = 'HT' then 
    begin
      Inc(LineNumber);
      AddLine(TobL.GetString('E_GENERAL'),GetAmount(TobL.GetDouble('E_DEBIT') , True), GetAmount(TobL.GetDouble('E_CREDIT'), True));
      if DoubleLineFromTaxCode(TobL.GetString('E_TVA')) then
      begin
        Inc(LineNumber);
        AddLine(TobL.GetString('BP4_CPTTVA'),GetAmount(0 , True), GetAmount(0, True));
      end;
    end;
  end;
end;

function TExportEcr.AddThirdToFile(Accounts : IXMLNode; TobData : TOB) : boolean;
var
  IsCustomer  : boolean;
  Account     : IXMLNode;
  SubLevel    : IXMLNode;
  SubLevel1   : IXMLNode;
  SubLevel2   : IXMLNode;
  SubLevel3   : IXMLNode;
  SubLevel4   : IXMLNode;
  TobAddress  : TOB;
  TobContact  : TOB;
  TobBank     : TOB;
  TobBankL    : TOB;
  TobContactL : TOB;
  TobAddressL : TOB;
  Cpt         : integer;
  ThirdMemo   : string;
  ThirdType   : string;


  function GetBankInformation(Country : string; tType : T_ThirdXmlFileType) : string;
  var
    BICCode : string;
  begin
    BICCode := Tools.iif(TobBankL.GetString('R_CODEBIC') <> '', TobBankL.GetString('R_CODEBIC'), '');
    if GetCountryIso2(Country) = 'BE' then
    begin
      case tType of
        txft_BankAccount_Code     : Result := StringReplace(TobBankL.GetString('R_NUMEROCOMPTE'), ' ', '', [rfReplaceAll]);
        txft_BankAccountType_Code : Result := GetCountryIso2(TobBankL.GetString('R_PAYS'));
        txft_Bank_Code            : Result := ''; 
        txft_Bic                  : Result := BICCode;
        txft_SwiftCode            : Result := Tools.iif(BICCode <> '', Format('%sXXX', [TobBankL.GetString('R_CODEBIC')]), '');
      end;
    end else
    if TobBankL.GetString('R_CODEIBAN') = '' then
    begin
      case tType of
        txft_BankAccount_Code     : Result := TobBankL.GetString('R_NUMEROCOMPTE');
        txft_BankAccountType_Code : Result := 'DEF';
        txft_Bank_Code            : Result := '';
        txft_Bic                  : Result := BICCode;
        txft_SwiftCode            : Result := BICCode;
      end;
    end else
      case tType of
        txft_BankAccount_Code     : Result := TobBankL.GetString('R_CODEIBAN');
        txft_BankAccountType_Code : Result := 'IBA';
        txft_Bank_Code            : Result := '';
        txft_Bic                  : Result := BICCode;
        txft_SwiftCode            : Result := '';
      end;
  end;

  function GetTaxSytem : string;
  begin
    { L=Assujeti, E=Exempté, N=Non assujeti }
    case Tools.CaseFromString(TobData.GetString('T_REGIMETVA'), ['COR', 'DTM', 'EXO', 'EXP', 'FRA', 'INT']) of
      {COR} 0 : Result := 'N';
      {DTM} 1 : Result := 'N';
      {EXO} 2 : Result := 'E';
      {EXP} 3 : Result := 'E';
      {FRA} 4 : Result := 'L';
      {INT} 5 : Result := 'L';
    else
      Result := 'L';
    end;
  end;

  procedure AddContact(Exist : boolean);
  const
    DefaultValue = 'XXX';
  begin
    SubLevel1 := SubLevel.AddChild('Contact');             SubLevel1.Attributes['default']    := Tools.iif(Exist, Tools.iif(TobContactL.GetBoolean('C_PRINCIPAL'), '1', '0'), '1');
                                                           SubLevel1.Attributes['gender']     := Tools.iif(Exist, Tools.iif(TobContactL.GetString('C_SEXE') = 'M', 'M', Tools.iif(TobContactL.GetString('C_SEXE') = 'F', 'V', 'O')), 'M'); 
                                                           SubLevel1.Attributes['status']     := 'A';
      SubLevel2 := SubLevel1.AddChild('LastName');         SubLevel2.Text                     := Tools.iif(Exist, TobContactL.GetString('C_NOM')   , DefaultValue);
      SubLevel2 := SubLevel1.AddChild('FirstName');        SubLevel2.Text                     := Tools.iif(Exist, TobContactL.GetString('C_PRENOM'), DefaultValue);
      SubLevel2 := SubLevel1.AddChild('MiddleName');       SubLevel2.Text                     := Tools.iif(Exist, '', DefaultValue);
      SubLevel2 := SubLevel1.AddChild('Initials');         SubLevel2.Text                     := Tools.iif(Exist, '', DefaultValue);
      SubLevel2 := SubLevel1.AddChild('Title');            SubLevel2.Attributes['code']       := Tools.iif(Exist, TobContactL.GetString('C_CIVILITE'), DefaultValue);
      SubLevel2 := SubLevel1.AddChild('Addresses');
        SubLevel3 := SubLevel2.AddChild('Address');        SubLevel3.Attributes['type']       := 'V';
                                                           SubLevel3.Attributes['desc']       := '';
          SubLevel4 := SubLevel3.AddChild('AddressLine1'); SubLevel4.Text                     := TobAddressL.GetString('ADR_ADRESSE1');
          SubLevel4 := SubLevel3.AddChild('AddressLine2'); SubLevel4.Text                     := TobAddressL.GetString('ADR_ADRESSE2');
          SubLevel4 := SubLevel3.AddChild('AddressLine3'); SubLevel4.Text                     := TobAddressL.GetString('ADR_ADRESSE3');
          SubLevel4 := SubLevel3.AddChild('PostalCode');   SubLevel4.Text                     := TobAddressL.GetString('ADR_CODEPOSTAL');
          SubLevel4 := SubLevel3.AddChild('City');         SubLevel4.Text                     := TobAddressL.GetString('ADR_VILLE');
          SubLevel4 := SubLevel3.AddChild('State');        SubLevel4.Attributes['code']       := TobAddressL.GetString('ADR_REGION');
          SubLevel4 := SubLevel3.AddChild('Country');      SubLevel4.Attributes['code']       := GetCountryIso2(TobAddressL.GetString('ADR_PAYS'));
          SubLevel4 := SubLevel3.AddChild('Phone');        SubLevel4.Text                     := Tools.iif(Exist, TobContactL.GetString('C_TELEPHONE'), DefaultValue);
          SubLevel4 := SubLevel3.AddChild('Fax');          SubLevel4.Text                     := Tools.iif(Exist, TobContactL.GetString('C_FAX')      , DefaultValue);
      SubLevel2 := SubLevel1.AddChild('Manager');          SubLevel2.Attributes['number']     := TobData.GetString('US_UTILISATEUR');
      SubLevel2 := SubLevel1.AddChild('Language');         SubLevel2.Attributes['code']       := RechDom('TTLANGUE', TobData.GetString('T_LANGUE'), True);
      SubLevel2 := SubLevel1.AddChild('JobTitle');         SubLevel2.Attributes['code']       := Tools.iif(Exist, TobContactL.GetString('C_SERVICECODE'), DefaultValue);
      SubLevel2 := SubLevel1.AddChild('JobDescription');   SubLevel2.Text                     := Tools.iif(Exist, TobContactL.GetString('C_SERVICE')    , DefaultValue);
      SubLevel2 := SubLevel1.AddChild('Phone');            SubLevel2.Text                     := Tools.iif(Exist, TobContactL.GetString('C_TELEPHONE')  , DefaultValue);
      SubLevel2 := SubLevel1.AddChild('Fax');              SubLevel2.Text                     := Tools.iif(Exist, TobContactL.GetString('C_FAX')        , DefaultValue);
      SubLevel2 := SubLevel1.AddChild('Mobile');           SubLevel2.Text                     := Tools.iif(Exist, TobContactL.GetString('C_TELEX')      , DefaultValue);
      SubLevel2 := SubLevel1.AddChild('Email');            SubLevel2.Text                     := Tools.iif(Exist, TobContactL.GetString('C_RVA')        , DefaultValue);
  end;

begin
  Result     := True;
  IsCustomer := (TobData.GetString('T_NATUREAUXI') = 'CLI');
  TobAddress := TobData.FindFirst(['TYPE'], ['ADRESSES'], True);
  TobContact := TobData.FindFirst(['TYPE'], ['CONTACT'] , True);
  TobBank    := TobData.FindFirst(['TYPE'], ['RIB']     , True);
  ThirdType  := Tools.iif(IsCustomer, 'Debtor', 'Creditor');
  if TobData.GetString('T_BLOCNOTE') <> '' then
  begin
    ThirdMemo := Tools.BlobToString_(TobData.GetString('T_BLOCNOTE'));
    ThirdMemo := StringReplace(ThirdMemo, '~~', #13#10, [rfReplaceAll, rfIgnoreCase]);
  end;

  Account  := Accounts.AddChild('Account'); Account.Attributes['code']    := TobData.GetString('T_AUXILIAIRE');
                                            Account.Attributes['status']  := Tools.iif(TobData.GetBoolean('T_FERME'), 'P', 'A');
                                            Account.Attributes['type']    := Tools.iif(IsCustomer, 'C', 'S');
  SubLevel := Account.AddChild('Name');     SubLevel.Text                 := TobData.GetString('T_LIBELLE');
  SubLevel := Account.AddChild('Phone');    SubLevel.Text                 := TobData.GetString('T_TELEPHONE');
  SubLevel := Account.AddChild('Fax');      SubLevel.Text                 := TobData.GetString('T_FAX');
  SubLevel := Account.AddChild('Email');    SubLevel.Text                 := TobData.GetString('T_EMAIL');
  SubLevel := Account.AddChild('HomePage'); SubLevel.Text                 := TobData.GetString('T_RVA');
  SubLevel := Account.AddChild('Contacts');
  if assigned(TobContact) then
  begin
    TobAddressL := TobAddress.detail[0];
    if TobContact.Detail.Count > 0 then
    begin
      for Cpt := 0 to pred(TobContact.Detail.Count) do
      begin
        TobContactL := TobContact.Detail[Cpt];
        AddContact(True);
      end;
    end else
    begin
      TobContactL := TOB.Create('_NONE', nil, -1);
      try
        AddContact(False);
      finally
        FreeAndNil(TobContactL);
      end;
    end;
  end;
  SubLevel := Account.AddChild('Note');     SubLevel.Text := ThirdMemo;
  SubLevel := Account.AddChild(ThirdType);  SubLevel.Attributes['number']   := TobData.GetString('T_AUXILIAIRE');
                                             SubLevel.Attributes['code']     := TobData.GetString('T_AUXILIAIRE');
    SubLevel1 := SubLevel.AddChild('Currency'); SubLevel1.Attributes['code'] := TobData.GetString('T_DEVISE');
    SubLevel1 := SubLevel.AddChild('BankAccounts');
    if (assigned(TobBank)) and (TobBank.Detail.Count > 0) then
    begin
      for Cpt := 0 to pred(TobBank.Detail.Count) do
      begin
        TobBankL := TobBank.Detail[Cpt];
        SubLevel2 := SubLevel1.AddChild('BankAccount');     SubLeveL2.Attributes['code']    := GetBankInformation(TobBankL.GetString('R_PAYS'), txft_BankAccount_Code);
                                                            SubLevel2.Attributes['default'] := Tools.iif(TobBankL.GetBoolean('R_PRINCIPAL'), '1', '0');
        SubLevel2 := SubLevel1.AddChild('BankAccountType'); SubLevel2.Attributes['code']    := GetBankInformation(TobBankL.GetString('R_PAYS'), txft_BankAccountType_Code);
        SubLevel2 := SubLevel1.AddChild('Bank');            SubLevel2.Attributes['code']    := GetBankInformation(TobBankL.GetString('R_PAYS'), txft_Bank_Code);
          SubLevel3 := SubLevel2.AddChild('Name');          SubLevel3.Text                  := TobBankL.GetString('R_DOMICILIATION');
          SubLevel3 := SubLevel2.AddChild('Country');       SubLevel3.Attributes['code']    := GetCountryIso2(TobBankL.GetString('R_PAYS'));
          if TobBankL.GetString('R_CODEIBAN') <> '' then
            SubLevel4 := SubLevel2.AddChild('IBAN');        SubLevel4.Text                  := TobBankL.GetString('R_CODEIBAN');
          SubLevel4 := SubLevel2.AddChild('BIC');           SubLevel4.Text                  := GetBankInformation(TobBankL.GetString('R_PAYS'), txft_Bic);
          if    (GetCountryIso2(TobBankL.GetString('R_PAYS')) = 'BE')
             or (TobBankL.GetString('R_CODEIBAN') = '') then
            SubLevel4 := SubLevel2.AddChild('SwiftCode');   SubLevel4.Text                  := GetBankInformation(TobBankL.GetString('R_PAYS'), txft_SwiftCode);
        SubLevel2 := SubLevel1.AddChild('Address');
          SubLevel3 := SubLevel2.AddChild('City');          SubLevel3.Text                  := TobBankL.GetString('R_VILLE');
          SubLevel3 := SubLevel2.AddChild('Country');       SubLevel3.Attributes['code']    := GetCountryIso2(TobBankL.GetString('R_PAYS'));
      end;
    end;
    SubLevel1 := SubLevel.AddChild('GLOffset');         SubLevel1.Attributes['code']   := Tools.iif(IsCustomer, VH_GC.GCCpteHTVTE, VH_GC.GCCpteHTACH);
    SubLevel1 := SubLevel.AddChild('GLCentralization'); SubLevel1.Attributes['code']   := TobData.GetString('T_COLLECTIF');
    SubLevel2 := SubLevel.AddChild(Tools.iif(IsCustomer, 'CreditLine', 'ExternalCode')); SubLevel2.Text := Tools.iif(IsCustomer, GetAmount(TobData.GetDouble('T_CREDITACCORDE'), True), TobData.GetString('YTC_TEXTELIBRE1'));
  SubLevel := Account.AddChild('VATNumber');           SubLevel.Text                  := TobData.GetString('T_NIF');
  SubLevel := Account.AddChild('VATLiability');        SubLevel.Text                  := GetTaxSytem;
  SubLevel := Account.AddChild('PaymentCondition');    SubLevel.Attributes['code']   := TobData.GetString('T_MODEREGLE');
  SubLevel := Account.AddChild('CompanySize');         SubLevel.Attributes['code']   := 'UNKNOWN';
  SubLevel := Account.AddChild('CompanyOrigin');       SubLevel.Attributes['code']   := 'P';
  SubLevel := Account.AddChild('CompanyRating');       SubLevel.Attributes['code']   := '7';
  SubLevel := Account.AddChild('Sector');              SubLevel.Attributes['code']   := 'UNKNOWN';
  SubLevel := Account.AddChild('AccountCategory');     SubLevel.Attributes['code']   := 'ME';
  SubLevel := Account.AddChild('DunsNumber');          SubLevel.Text                  := CustomerCode;
end;

function TExportEcr.AccountingExpTreatment: boolean;
var
  MsgConfirm     : string;
  MsgDelete      : string;
  PathXMLFile    : string;
  PathXMLTmpFile : string;
  Msg            : string;
  TobEcrL        : TOB;
  WaitMsg        : WaitingMessage;
  CanContinue    : boolean;
  WithDelete     : boolean;

  procedure AddLineToTob;
  var
    TobEcrPL   : TOB;
  begin
    TobEcrPL := TOB.Create('_LINE', TobEcrP, -1);
    TobEcrPL.Dupliquer(TobEcrL, True, True);
  end;

  procedure EntryTreatment(Flow : string);
  var
    TobData    : TOB;
    CurrentPce : integer;
    XmlDoc     : IXMLDocument;
    RootNode   : IXMLNode;
    GLEntries  : IXMLNode;
    Cpt        : integer;

    procedure AddDocToFile;
    begin
      if TobEcrP.Detail.Count > 0 then
      begin
        LoadAnalytic;
        LoadTax;
        AddEntryToFile(GLEntries);
        TobEcrP.ClearDetail;
      end;
    end;

  begin
    TobData := Tools.iif(Flow = 'VTE', TobEcrVen, TobEcrAch);
    CurrentPce := 0;
    LineNumber := 0;
    XmlDoc := NewXMLDocument();
    try
      XmlDoc.Options := [doNodeAutoIndent];
      RootNode := XmlDoc.AddChild('eExact');
        RootNode.Attributes['xmlns:xsi']                     := 'http://www.w3.org/2001/XMLSchema-instance';
        RootNode.Attributes['xsi:noNamespaceSchemaLocation'] := 'eExact-Schema.xsd';
      GLEntries := RootNode.AddChild('GLEntries');
      for Cpt := 0 to pred(TobData.Detail.Count) do
      begin
        TobEcrL := TobData.Detail[Cpt];
        if CurrentPce <> TobEcrL.GetInteger('E_NUMEROPIECE') then
        begin
          AddDocToFile;
          CurrentPce := TobEcrL.GetInteger('E_NUMEROPIECE');
          AddLineToTob;
        end else
          AddLineToTob;
      end;
      AddDocToFile;
      TobEcrP.ClearDetail;
      XmlDoc.SaveToFile(Tools.iif(Flow = 'VTE', TempPathXMLFileVTE, TempPathXMLFileACH));
    finally
      XmlDoc := nil;
    end;
  end;

  function ThirdTreatment(Flow : string) : boolean;
  var
    TobData    : TOB;
    XmlDoc     : IXMLDocument;
    RootNode   : IXMLNode;
    Accounts   : IXMLNode;
    Cpt        : integer;
  begin
    Result  := True;
    TobData := Tools.iif(Flow = 'VTE', TobCustomer, TobProvider);
    XmlDoc := NewXMLDocument();
    try
      XmlDoc.Options := [doNodeAutoIndent];
      RootNode := XmlDoc.AddChild('eExact');
        RootNode.Attributes['xmlns:xsi']                     := 'http://www.w3.org/2001/XMLSchema-instance';
        RootNode.Attributes['xsi:noNamespaceSchemaLocation'] := 'eExact-Schema.xsd';
      Accounts := RootNode.AddChild('Accounts');
      for Cpt := 0 to pred(TobData.Detail.Count) do
        AddThirdToFile(Accounts, TobData.Detail[Cpt]);
      TobEcrP.ClearDetail;
      XmlDoc.SaveToFile(Tools.iif(Flow = 'VTE', TempPathXMLFileCLI, TempPathXMLFileFOU));
    finally
      XmlDoc := nil;
    end;
  end;

  function CheckExportEntry(TobE : TOB) : boolean;
  var
    Cpt        : integer;
    TobEL      : TOB;
    CurrentIdx : string;
    Sql        : string;

    function GetKey : string;
    begin
      Result := TobEL.GetString('E_ENTITY') + ';' + TobEL.GetString('E_EXERCICE') + ';' + TobEL.GetString('E_JOURNAL') + ';' + TobEL.GetString('E_NUMEROPIECE');
    end;

  begin
    Result     := True;
    CurrentIdx := '';
    for Cpt := 0 to pred(TobE.Detail.Count) do
    begin
      TobEL := TobE.Detail[Cpt];
      if CurrentIdx <> GetKey then
      begin
        TslData.Add(Format(' . %s - %s - %s - %s', [TobEL.GetString('E_NATUREPIECE'), TobEL.GetString('E_EXERCICE'), TobEL.GetString('E_JOURNAL'), TobEL.GetString('E_NUMEROPIECE')]));
        CurrentIdx := GetKey;
        Sql        := Format('%s OR (E_ENTITY = %s AND E_EXERCICE = "%s" AND E_JOURNAL = "%s" AND E_NUMEROPIECE = %s)'
                             , [  Sql
                                , TobEL.GetString('E_ENTITY')
                                , TobEL.GetString('E_EXERCICE')
                                , TobEL.GetString('E_JOURNAL')
                                , TobEL.GetString('E_NUMEROPIECE')
                               ]);
      end;
    end;
    if Sql <> '' then
    begin
      Sql    := Format('UPDATE ECRITURE SET E_EXPORTE = "X" WHERE %s', [Copy(Sql, 4, length(Sql))]);
      Result := (ExecuteSql(Sql) = TobE.Detail.Count);
    end;
  end;

  function GetFileName(Suffix : string; WithExtension : boolean=true) : string;
  begin
    Result := Format('%s_%s%s', [FormatDateTime('yyyymmdd', NowH), Suffix, Tools.iif(WithExtension, '.xml', '')]);
  end;

begin
  Result         := True;
  MsgCaption     := 'Export écritures';
  PathXMLFile    := IncludeTrailingBackslash(GetParamSocSecur('SO_EXPXMLDIR', ''));
  PathXMLTmpFile := IncludeTrailingBackslash(GetEnvironmentVariable('TEMP'));
  PathXMLFileVTE := PathXMLFile + GetFileName('EcrVTE');
  PathXMLFileACH := PathXMLFile + GetFileName('EcrACH');
  PathXMLFileCLI := PathXMLFile + GetFileName('Clients');
  PathXMLFileFOU := PathXMLFile + GetFileName('Fournisseurs');
  FolderCode     := GetParamSocSecur('SO_EXPXMLDOSSIERCPTA', '');
  CustomerCode   := GetParamSocSecur('SO_EXPXMLCLIENTCPTA', '');
  MsgConfirm     := '';
  CanContinue    := ((PathXMLFile <> '') and (FolderCode <> '') and (CustomerCode <> ''));
  if not CanContinue then
  begin
         if PathXMLFile  = '' then PGIError('Export impossible, le répertoire de stockage des fichiers n''est pas renseigné.', MsgCaption)
    else if FolderCode   = '' then PGIError('Export impossible, le numéro de dossier EXACT n''est pas renseigné.', MsgCaption)
    else if CustomerCode = '' then PGIError('Export impossible, le code client EXACT n''est pas renseigné.', MsgCaption)
  end else
  begin
    WithDelete := ((FileExists(PathXMLFileVTE)) or (FileExists(PathXMLFileACH)) or (FileExists(PathXMLFileCLI)) or (FileExists(PathXMLFileFOU)));
    MsgDelete  := Tools.iif(WithDelete, Format('ATTENTION, un ou plusieurs fichiers préfixés de "%s" existent déjà dans "%s". En continuant, ces derniers seront écrasés.%s', [GetFileName('', False), PathXMLFile, #13#10]), '');
    MsgConfirm := Format('Ce traitement exporte toutes les écritures%s non exportées du %s au %s et bloque les pièces associées.#13#10%s Voulez-vous continuer ?'
                         , [  Tools.iif(ForceExport, ' déjà exportées et', '')
                            , DateToStr(StartDate)
                            , DateToStr(EndDate)
                            , MsgDelete
                           ]);
    if PGIAsk(MsgConfirm, MsgCaption) = mrYes then
    begin
      WaitMsg := WaitingMessage.Create('Export des écritures.', 'Traitement en cours ...');
      try
        if WithDelete then
        begin
          TslData.Add(Format('Demande de suppression des fichiers préfixése de "%s" existants dans "%s".', [GetFileName('', False), PathXMLFile]));
          TslData.Add('');
          DeleteFile(PAnsiChar(PathXMLFileVTE));
          DeleteFile(PAnsiChar(PathXMLFileACH));
          DeleteFile(PAnsiChar(PathXMLFileCLI));
          DeleteFile(PAnsiChar(PathXMLFileFOU));
        end;
        TobCountry.LoadDetailFromSQL('SELECT PY_PAYS, PY_CODEISO2 FROM PAYS');
        TempPathXMLFileVTE := PathXMLTmpFile + GetFileName('EcrVTE');
        TempPathXMLFileACH := PathXMLTmpFile + GetFileName('EcrACH');
        TempPathXMLFileCLI := PathXMLTmpFile + GetFileName('Clients');
        TempPathXMLFileFOU := PathXMLTmpFile + GetFileName('Fournisseurs');
        { Supprime les fichiers du répertoire temporaire si existent }
        DeleteFile(PAnsiChar(TempPathXMLFileVTE));
        DeleteFile(PAnsiChar(TempPathXMLFileACH));
        DeleteFile(PAnsiChar(TempPathXMLFileCLI));
        DeleteFile(PAnsiChar(TempPathXMLFileFOU));
        TslData.Add(Format('Exports effectués : ', []));
        TslData.Add('');
        if LoadEcr then
        begin
          TslData.Add('- Pièces comptable (Nature - Exercice - Journal - Numéro de pièce) :');
          BeginTrans;
          try
            if TobEcrVen.Detail.Count   > 0 then
            begin
              try
                EntryTreatment('VTE');
              finally
                Result := CheckExportEntry(TobEcrVen);
                if not Result then
                  PGIError('Erreur lors de la mise à jour des écritures de ventes.', MsgCaption);
              end;
            end;
            if (Result) and (TobEcrAch.Detail.Count   > 0) then
            begin
              try
                EntryTreatment('ACH');
              finally
                Result := CheckExportEntry(TobEcrAch);
                if not Result then
                  PGIError('Erreur lors de la mise à jour des écritures de ventes.', MsgCaption);
              end;
            end;
            if (Result) and (TobCustomer.Detail.Count > 0) then Result := ThirdTreatment('VTE');
            if (Result) and (TobProvider.Detail.Count > 0) then Result := ThirdTreatment('ACH');
          finally
            if Result then
              CommitTrans
            else
              Rollback;
          end;
          if Result then
          begin
            MoveFile(PAnsiChar(TempPathXMLFileVTE), PAnsiChar(PathXMLFileVTE));
            MoveFile(PAnsiChar(TempPathXMLFileACH), PAnsiChar(PathXMLFileACH));
            MoveFile(PAnsiChar(TempPathXMLFileCLI), PAnsiChar(PathXMLFileCLI));
            MoveFile(PAnsiChar(TempPathXMLFileFOU), PAnsiChar(PathXMLFileFOU));
            SetParamSoc('SO_EXPXMLDE', StartDate);
            SetParamSoc('SO_EXPXMLA' , EndDate);
          end;
          Msg := Format('%s des écritures %s de ventes et achats du %s au %s.%s%s'
                        , [  Tools.iif(Result, 'Export', 'Tentative d''export')
                           , Tools.iif(ForceExport, 'déjà exportées et non exportées', 'non exportées')
                           , DateTimeToStr(StartDate)
                           , DateTimeToStr(EndDate)
                           , #13#10
                           , TslData.Text
                          ]);
          MAJJnalEvent('EPC', Tools.iif(Result, 'OK', 'ERR'), MsgCaption, Msg);
          PGIBox(Format('Traitement terminé avec %s.', [Tools.iif(Result, 'succés', 'des erreurs')]), MsgCaption);
        end else
          PGIBox('Il n''y a pas d''écritures à exporter sur la période demandée.', MsgCaption);
      finally
        WaitMsg.Free
      end;
    end;
  end;
end;
{$ENDIF APPSRV}

end.
