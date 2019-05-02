unit uExecuteService;

interface

uses
  Windows
  , Classes
  , CommonTools
  , UConnectWSConst
  , ConstServices
  ;

type

  TSvcSyncBTPY2Execute = class(TObject)
  private
    TSlConnectionValues            : TStringList;
    TSlValues                      : TStringList;                                                         
    TSlFilter                      : TStringList;
    TSlIndice                      : TStringList;
    TSlCacheThirdBTP               : TStringList;
    TSlCacheSectionBTP             : TStringList;
    TSlCacheAcountBTP              : TStringList;
    TSlCachePaymentBTP             : TStringList;
    TSlCacheCorrespBTP             : TStringList;
    TSlCacheCurrencyBTP            : TStringList;
    TSlCacheCountryBTP             : TStringList;
    TSlCacheRecoveryBTP            : TStringList;
    TSlCacheCommonBTP              : TStringList;
    TSlCacheChoixCodBTP            : TStringList;
    TSlCacheChoixExtBTP            : TStringList;
    TSlCacheJournalBTP             : TStringList;
    TSlCacheBankIdBTP              : TStringList;
    TSlCacheChangeRateBTP          : TStringList;
    TSlCacheFiscalYearBTP          : TStringList;
    TSlcacheSocietyParamBTP        : TStringList;
    TSlcacheEstablishmentBTP       : TStringList;
    TSlcachePaymentModeBTP         : TStringList;
    TSlcacheZipCodeBTP             : TStringList;
    TSlcacheContactBTP             : TStringList;
    TSlCacheAccForPaymentBTP       : TStringList;
    TSLCacheDocPaymentBTP          : TStringList;
    TSLCacheTaxRateBTP             : TStringList;
    TSlCacheThirdY2                : TStringList;
    TSlCacheSectionY2              : TStringList;
    TSlCacheAcountY2               : TStringList;
    TSlCachePaymentY2              : TStringList;
    TSlCacheCorrespY2              : TStringList;
    TSlCacheCurrencyY2             : TStringList;
    TSlCacheCountryY2              : TStringList;
    TSlCacheRecoveryY2             : TStringList;
    TSlCacheCommonY2               : TStringList;
    TSlCacheChoixCodY2             : TStringList;
    TSlCacheChoixExtY2             : TStringList;
    TSlCacheJournalY2              : TStringList;
    TSlCacheBankIdY2               : TStringList;
    TSlCacheChangeRateY2           : TStringList;
    TSlCacheFiscalYearY2           : TStringList;
    TSlcacheSocietyParamY2         : TStringList;
    TSlcacheEstablishmentY2        : TStringList;
    TSlcachePaymentModeY2          : TStringList;
    TSlcacheZipCodeY2              : TStringList;
    TSlcacheContactY2              : TStringList;
    TSlCacheAccForPaymentY2        : TStringList;
    TSLCacheTaxRateY2              : TStringList;
    TSlCacheWSFields               : TStringList;
    TSlCacheSendAccParam           : TStringList;
    TSlCacheGetY2Data              : TStringList;
    TSLCacheUpdateFrequencySetting : TStringList;
    TSLUpdateInsertData            : TStringList;
    TSLTRAFileQty                  : TStringList;
    AdoQryY2                       : AdoQry;
    AdoQryBTP                      : AdoQry;
    BTPValues                      : T_WSBTPValues;
    Y2Values                       : T_WSY2Values;
    NewAna                         : boolean;

    procedure ClearValuesConnection;
    procedure SetLastSyncIniFile;
    procedure SetFilterFromDSType(DSType: T_WSDataService; TSl: TStringList);
    function GetInfoFromDSType(InfoType: T_WSInfoFromDSType; DSType: T_WSDataService; FieldName: string = ''): string;
    function GetSingleValueFromTSl(FieldName: string; TslOrig : TStringList): string;
    function GetMultipleValueFromTSl(FieldsName : string; TslOrig : TStringList) : string;
    function GetTableNameFromDSType(DSType: T_WSDataService): string;
    function GetDSTypeFromTableName(TableName : string) : T_WSDataService;
    function GetKeyValues(DSType: T_WSDataService; lTSlValues: TStringList): string;
    function GetIndiceY2DataCache(DSType: T_WSDataService; TSLValue: TStringList): string;
    function AddQuotes(DSType: T_WSDataService; FieldName, FieldValue: string): string;
    function ValueAnalysis(LineValue, FieldName : string; FieldType : tTypeField) : string;
    function GetY2Data(DSType: T_WSDataService) : boolean;
    function GetPayment(DSType: T_WSDataService) : boolean;
    function GetData: boolean;
    function SendData: boolean;
    function ReadSettings : boolean;
    function CreateMemoryCache : boolean;
    procedure AddWindowsLog(Start: boolean);
    procedure LoadEcrFromBE0(AdoQryEcr, AdoQryAna: AdoQry; BE0Values: string);
    procedure SetSendY2TSl(AdoQryEcr, AdoQryAna: AdoQry; TSlResult: TStringList);
    procedure SearchRelatedParameters(TSlEntry: TStringList);
    procedure SearchOthersParameters;
    function GetFieldsList(lTSl: TStringList): string;
    function SetDataToSend(FieldsList, FieldsValues: string): string;
    function AddUpdateValues(AdoQryParam: AdoQry; DSType: T_WSDataService; FieldsList: string; KeyValue1, KeyValue2: string; IsRelatedParameters: boolean) : boolean;
    function FieldExistInTSl(lTSl: TStringList; FieldName: string): boolean;
    function GetTSlFromDSType(DSType: T_WSDataService; IsBTP: boolean): TStringList;
    function GetFieldTypeFromCache(DSType: T_WSDataService; FieldName: string): tTypeField;
    function SendY2Settings: boolean;
    function SendAccountingEntries(TslEntries : TStringList; var ErrorMsg : string) : integer;
    function AddData(wsAction: T_WSAction; DSType: T_WSDataService; lTSlValues: TStringList): boolean;
    procedure ExtractIndice(Cpt: integer; TSlOrig, TSlResult: TStringList);
    function CanExportImportTableFromFrenquency(DSType: T_WSDataService) : boolean;

  public
    AppFilePath   : string;
    IniFilePath   : string;
    LogFilePath   : string;
    SecondTimeout : integer;
    LogValues     : T_WSLogValues;

    procedure CreateObjects;
    procedure FreeObjects;
    function ServiceExecute: Boolean;
    function InitApplication : boolean;
  end;

implementation

uses
  Registry
  , Graphics
  , Controls
  , Dialogs
  , SysUtils
  , Messages
  , uWSDataService
  , SvcMgr
  , StrUtils
  , UConnectWSCEGID
  , IniFiles
  , DateUtils
  , TRAFileUtil
  , uLog
  , AdoDB
  {$IF not defined(APPSRV)}
  , ParamSoc
  {$IFEND (APPSRV)}
  ;

const
  IniFileLastSynchro = 'LastSynchro';
  IniFileBTPUser     = 'BTPUser';
  IniFileBTPServer   = 'BTPServer';
  IniFileBTPFolder   = 'BTPFolder';
  IniFileY2Server    = 'Y2Server';
  IniFileY2Folder    = 'Y2Folder';

function TSvcSyncBTPY2Execute.CreateMemoryCache : boolean;
var
  lTslFieldsList: TStringList;

  function SetBTPDefaultValue(DSType: T_WSDataService) : boolean;
  var
    CptCache     : integer;
    lTsl         : TStringList;
    Prefix       : string;
    Value        : string;
    FieldName    : string;
    FieldType    : string;
    DefaultValue : string;
  begin
    Result := True;
    lTsl   := GetTSlFromDSType(DSType, True);
    Prefix := TGetFromDSType.dstPrefix(DSType);
    AdoQryBTP.TSLResult.Clear;
    AdoQryBTP.FieldsList := 'DH_NOMCHAMP,DH_TYPECHAMP';
    AdoQryBTP.Request    := Format('SELECT %s FROM DECHAMPS WHERE DH_PREFIXE = ''%s'' ORDER BY DH_NOMCHAMP', [AdoQryBTP.FieldsList, Prefix]);
    try
      AdoQryBTP.SingleTableSelect;
      if AdoQryBTP.TSLResult.Count > 0 then
      begin
        for CptCache := 0 to pred(AdoQryBTP.TSLResult.Count) do
        begin
          Value     := AdoQryBTP.TSLResult[CptCache];
          FieldName := Tools.ReadTokenSt_(Value, ToolsTobToTsl_Separator);
          FieldType := Value;
          if Tools.GetTypeFieldFromStringType(FieldType) <> ttfMemo then // Exclusion des memo
          begin
            case Tools.GetTypeFieldFromStringType(Value) of
              ttfNumeric : DefaultValue := '0';
              ttfInt     : DefaultValue := '0';
              ttfBoolean : DefaultValue := '''-''';
              ttfDate    : DefaultValue := '''' + DateToStr(2) + '''';
              ttfMemo    : DefaultValue := '''''';
              ttfCombo   : DefaultValue := '''''';
              ttfText    : DefaultValue := '''''';
            end;
            lTsl.Add(FieldName + ToolsTobToTsl_Separator + FieldType + ToolsTobToTsl_Separator + DefaultValue);
          end;
        end;
      end;
      AdoQryBTP.Reset;
    except
      if LogValues.DebugEvents > 0 then
        TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.CreateMemoryCache/SetBTPDefaultValue. Erreur sur ex�cution de %s / %s / %s'
                                        , [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryBTP.ServerName, AdoQryBTP.DBName, AdoQryBTP.Request]), ServiceName_BTPY2, LogValues, 0);
      Result := False;
      exit;
    end;
  end;

  procedure SetY2FieldsList(DSType: T_WSDataService);
  var
    lTsl   : TStringList;
    Cpt    : Integer;
    Prefix : string;
  begin
    lTsl := GetTSlFromDSType(DSType, False);
    for Cpt := 0 to pred(lTslFieldsList.Count) do
    begin
      Prefix := Copy(lTslFieldsList[Cpt], pos('=', lTslFieldsList[Cpt]) + 1, length(lTslFieldsList[Cpt]));
      Prefix := copy(Prefix, 1, pos('_', Prefix));
      if Prefix = TGetFromDSType.dstPrefix(DSType) + '_' then
        lTsl.Add(copy(lTslFieldsList[Cpt], pos('=', lTslFieldsList[Cpt]) + 1, length(lTslFieldsList[Cpt])) + ToolsTobToTsl_Separator);
    end;
  end;

  procedure SetViewFields(DSType: T_WSDataService);
  begin
    TSlCacheWSFields.Add(TGetFromDSType.dstWSName(DSType) + '=' + TGetFromDSType.dstFiedsList(DSType));
  end;

begin
  Result := True;
  { Liste des champs BTP }
  SetBTPDefaultValue(wsdsThird);
  SetBTPDefaultValue(wsdsAnalyticalSection);
  SetBTPDefaultValue(wsdsAccount);
  SetBTPDefaultValue(wsdsPaymenChoice);
  SetBTPDefaultValue(wsdsCorrespondence);
  SetBTPDefaultValue(wsdsCurrency);
  SetBTPDefaultValue(wsdsCountry);
  SetBTPDefaultValue(wsdsRecovery);
  SetBTPDefaultValue(wsdsCommon);
  SetBTPDefaultValue(wsdsChoixCod);
  SetBTPDefaultValue(wsdsChoixExt);
  SetBTPDefaultValue(wsdsJournal);
  SetBTPDefaultValue(wsdsBankId);
  SetBTPDefaultValue(wsdsChangeRate);
  SetBTPDefaultValue(wsdsFiscalYear);
  SetBTPDefaultValue(wsdsSocietyParameters);
  SetBTPDefaultValue(wsdsEstablishment);
  SetBTPDefaultValue(wsdsPaymentMode);
  SetBTPDefaultValue(wsdsZipCode);
  SetBTPDefaultValue(wsdsContact);
  SetBTPDefaultValue(wsdsAccForPayment);
//  SetBTPDefaultValue(wsdsAccForPaymentOther);
  SetBTPDefaultValue(wsdsDocPayment);
  SetBTPDefaultValue(wsdsTaxRate);
  { Liste des champs par vue }
  SetViewFields(wsdsThird);
  SetViewFields(wsdsAnalyticalSection);
  SetViewFields(wsdsAccount);
  SetViewFields(wsdsJournal);
  SetViewFields(wsdsBankId);
  SetViewFields(wsdsChoixCod);
  SetViewFields(wsdsCommon);
  SetViewFields(wsdsChoixExt);
  SetViewFields(wsdsRecovery);
  SetViewFields(wsdsCountry);
  SetViewFields(wsdsCurrency);
  SetViewFields(wsdsCorrespondence);
  SetViewFields(wsdsPaymenChoice);
  SetViewFields(wsdsChangeRate);
  SetViewFields(wsdsFiscalYear);
  SetViewFields(wsdsSocietyParameters);
  SetViewFields(wsdsEstablishment);
  SetViewFields(wsdsPaymentMode);
  SetViewFields(wsdsZipCode);
  SetViewFields(wsdsContact);
  SetViewFields(wsdsFieldsList);
  SetViewFields(wsdsAccForPayment);
//  SetViewFields(wsdsAccForPaymentOther);
  SetViewFields(wsdsTaxRate);
  { Liste des champs Y2 }
  lTslFieldsList := TStringList.Create;
  try
    try
      if LogValues.DebugEvents = 2 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.CreateMemoryCache - Before GetData', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
      Result := (TReadWSDataService.GetData(wsdsFieldsList, AdoQryY2.ServerName, AdoQryY2.DBName, lTslFieldsList, TSlCacheWSFields, LogValues) = WSCDS_GetDataOk);
      if LogValues.DebugEvents = 2 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.CreateMemoryCache - After GetData', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
      if Result then
      begin
        if LogValues.DebugEvents = 2 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.CreateMemoryCache - Start SetY2FieldsList', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
        SetY2FieldsList(wsdsThird);
        SetY2FieldsList(wsdsAnalyticalSection);
        SetY2FieldsList(wsdsAccount);
        SetY2FieldsList(wsdsJournal);
        SetY2FieldsList(wsdsBankId);
        SetY2FieldsList(wsdsChoixCod);
        SetY2FieldsList(wsdsCommon);
        SetY2FieldsList(wsdsChoixExt);
        SetY2FieldsList(wsdsRecovery);
        SetY2FieldsList(wsdsCountry);
        SetY2FieldsList(wsdsCurrency);
        SetY2FieldsList(wsdsCorrespondence);
        SetY2FieldsList(wsdsPaymenChoice);
        SetY2FieldsList(wsdsChangeRate);
        SetY2FieldsList(wsdsFiscalYear);
        SetY2FieldsList(wsdsSocietyParameters);
        SetY2FieldsList(wsdsEstablishment);
        SetY2FieldsList(wsdsPaymentMode);
        SetY2FieldsList(wsdsZipCode);
        SetY2FieldsList(wsdsContact);
        SetY2FieldsList(wsdsAccForPayment);
//        SetY2FieldsList(wsdsAccForPaymentOther);
        SetY2FieldsList(wsdsTaxRate);
        if LogValues.DebugEvents = 2 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.CreateMemoryCache - End SetY2FieldsList', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
      end;
    except
      TServicesLog.WriteLog(ssbylLog, Format('%s : Cr�ation du cache (TSvcSyncBTPY2Execute.CreateMemoryCache).', [WSCDS_ErrorMsg]), ServiceName_BTPY2, LogValues, 0);
    end;
  finally
    FreeAndNil(lTslFieldsList);
  end;
end;

procedure TSvcSyncBTPY2Execute.AddWindowsLog(Start: boolean);
var
  LogText        : string;
  ConnectionLine : string;
  Cpt            : Integer;
begin
  if Start then
  begin
    LogText := Format('Initialisation de  %s.', [ServiceName_BTPY2])
             + Format('%s> D�clenchement toutes les : %s secondes.', [#13#10, IntToStr(SecondTimeout)])
             + Format('%s> Nombre de connexions : %s'          , [#13#10, IntToStr(TSlConnectionValues.Count)]);
    for Cpt := 0 to pred(TSlConnectionValues.Count) do
    begin
      ConnectionLine := TSlConnectionValues[Cpt];
      LogText := LogText
               + Format('%s   %s :'                     , [#13#10, Tools.ReadTokenSt_(ConnectionLine, '=')])
               + Format('%s   . Compte utilisateur : %s', [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
               + Format('%s   . BTP-Serveur : %s'       , [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
               + Format('%s   . BTP-Dossier : %s'       , [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')]);
      Tools.ReadTokenSt_(ConnectionLine, ';'); // LastSynchro qui n'apparait pas dans le log
      LogText := Logtext + Format('%s   . Y2-Serveur : %s'        , [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
                         + Format('%s   . Y2-Dossier : %s'       , [#13#10, Tools.ReadTokenSt_(ConnectionLine, ';')])
    end;
  end else
  begin

  end;
  TServicesLog.WriteLog(ssbylWindows, LogText, ServiceName_BTPY2, LogValues, 0);
end;

procedure TSvcSyncBTPY2Execute.LoadEcrFromBE0(AdoQryEcr, AdoQryAna: AdoQry; BE0Values: string);
var
  Cpt: integer;
  lBE0Values  : string;
  BE0Entity   : string;
  BE0Exercice : string;
  BE0Journal  : string;
  BE0NumPce   : string;

  function GetValueFromFieldName(FieldName : string) : string;
  var
    FieldIndex  : integer;
  begin
    FieldIndex := Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, FieldName);
    Result     := Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], FieldIndex, ToolsTobToTsl_Separator);
  end;

begin
  lBE0Values  := BE0Values;
  BE0Entity   := Tools.ReadTokenSt_(lBE0Values, ToolsTobToTsl_Separator);
  BE0Exercice := Tools.ReadTokenSt_(lBE0Values, ToolsTobToTsl_Separator);
  BE0Journal  := Tools.ReadTokenSt_(lBE0Values, ToolsTobToTsl_Separator);
  BE0NumPce   := Tools.ReadTokenSt_(lBE0Values, ToolsTobToTsl_Separator);
  AdoQryEcr.Request := Format('SELECT %s'
                            + ' FROM ECRITURE'
                            + ' LEFT JOIN RIB ON R_AUXILIAIRE = E_AUXILIAIRE AND R_PRINCIPAL = ''X'''
                            + ' WHERE E_ENTITY      = ''%s'''
                            + '   AND E_EXERCICE    = ''%s'''
                            + '   AND E_JOURNAL     = ''%s'''
                            + '   AND E_NUMEROPIECE = %s'
                           , [  AdoQryEcr.FieldsList
                              , BE0Entity
                              , BE0Exercice
                              , BE0Journal
                              , BE0NumPce
                             ]);
  AdoQryEcr.LogValues := LogValues;
  try
    AdoQryEcr.SingleTableSelect;
    { Recherche l'IBAN si n�cessaire et charge l'�ventuelle analytique }
    for Cpt := 0 to pred(AdoQryEcr.TSLResult.Count) do
    begin
      if GetValueFromFieldName('E_ANA') = 'X' then
      begin
        AdoQryAna.TSLResult.Clear;
        AdoQryAna.Request := Format('SELECT %s FROM ANALYTIQ WHERE Y_ENTITY = ''%s'' AND Y_EXERCICE = ''%s'' AND Y_JOURNAL = ''%s'' AND Y_NUMEROPIECE = %s'
                                    , [AdoQryAna.FieldsList
                                       , Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_ENTITY')  , ToolsTobToTsl_Separator)
                                       , Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_EXERCICE'), ToolsTobToTsl_Separator)
                                       , Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_JOURNAL') , ToolsTobToTsl_Separator)
                                       , Tools.GetStValueFromTSl(AdoQryEcr.TSLResult[Cpt], Tools.GetTSlIndexFromFieldName(AdoQryEcr.FieldsList, 'E_NUMEROPIECE'), ToolsTobToTsl_Separator)
                                      ]);
        AdoQryAna.LogValues   := LogValues;
        AdoQryAna.SingleTableSelect;
        Break;
      end;
    end;
  except
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.LoadEcrFromBE0. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryEcr.ServerName, AdoQryEcr.DBName, AdoQryEcr.Request]), ServiceName_BTPY2, LogValues, 0);
    exit;
  end;
end;

procedure TSvcSyncBTPY2Execute.SetSendY2TSl(AdoQryEcr, AdoQryAna: AdoQry; TSlResult: TStringList);
var
  CptE            : integer;
  CptY            : integer;
  AxisIndex       : integer;
  LineNumberIndex : integer;
  LineValue       : string;
  FieldName       : string;
  FieldValue      : string;
  LineNumber      : string;
  AxisNumber      : string;
  WithAna         : boolean;
  AddAxis         : boolean;

  function GetLevel(Level: integer): string;
  begin
    Result := ToolsTobToTsl_LevelName + IntToStr(Level);
  end;

  function GetLineValue(FieldsList, FieldsValue: string): string;
  begin
    Result := '';
    while FieldsList <> '' do
    begin
      FieldName  := Tools.ReadTokenSt_(FieldsList, ',');
      FieldValue := Tools.ReadTokenSt_(FieldsValue, ToolsTobToTsl_Separator);
      Result     := Result + ToolsTobToTsl_Separator + FieldName + '=' + FieldValue;
    end;
  end;

begin
  WithAna := (AdoQryAna.TSLResult.Count > 0);
  if WithAna then
  begin
    LineNumberIndex := Tools.GetTSlIndexFromFieldName(AdoQryAna.FieldsList, 'Y_NUMLIGNE');
    AxisIndex       := Tools.GetTSlIndexFromFieldName(AdoQryAna.FieldsList, 'Y_AXE');
  end else
  begin
    LineNumberIndex := -1;
    AxisIndex       := -1;
  end;
  TSlResult.Add(GetLevel(1) + '=COMPTABILITE');
  for CptE := 0 to pred(AdoQryEcr.TSLResult.count) do
  begin
    LineValue := GetLineValue(AdoQryEcr.FieldsList, AdoQryEcr.TSLResult[CptE]);
    TSlResult.Add(GetLevel(2) + '=ECRITURE' + LineValue);
    { La ligne est ventilable, ajout de l'analytique }
    if (WithAna) and (Tools.GetStValueFromTSl(TSlResult[pred(TSlResult.Count)], 'E_ANA') = 'X') then
    begin
      AddAxis := False;
      LineNumber := Tools.GetStValueFromTSl(TSlResult[pred(TSlResult.Count)], 'E_NUMLIGNE');
      for CptY := 0 to pred(AdoQryAna.TSLResult.count) do
      begin
        if Tools.GetStValueFromTSl(AdoQryAna.TSLResult[CptY], LineNumberIndex, ToolsTobToTsl_Separator) = LineNumber then
        begin
          AxisNumber := Tools.GetStValueFromTSl(AdoQryAna.TSLResult[CptY], AxisIndex, ToolsTobToTsl_Separator);
          if not AddAxis then
          begin
            AddAxis := True;
            TSlResult.Add(GetLevel(3) + '=' + AxisNumber);
          end;
          LineValue := GetLineValue(AdoQryAna.FieldsList, AdoQryAna.TSLResult[CptY]);
          TSlResult.Add(GetLevel(4) + '=ANALYTIQ' + LineValue);
        end;
      end;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.GetFieldsList(lTSl: TStringList): string;
var
  Cpt: integer;
begin
  Result := '';
  for Cpt := 0 to pred(lTSl.Count) do
    Result := Result + ',' + Copy(lTSl[Cpt], 1, Pos(ToolsTobToTsl_Separator, lTSl[Cpt]) - 1);
  Result := Copy(Result, 2, Length(Result));
end;

function TSvcSyncBTPY2Execute.SetDataToSend(FieldsList, FieldsValues: string): string;
var
  lFieldsList   : string;
  lFieldsValues : string;
begin
  Result := '';
  if FieldsList <> '' then
  begin
    lFieldsList := FieldsList;
    lFieldsValues := FieldsValues;
    while lFieldsList <> '' do
      Result := Result + ToolsTobToTsl_Separator + Tools.ReadTokenSt_(lFieldsList, ',') + '=' + Tools.ReadTokenSt_(lFieldsValues, ToolsTobToTsl_Separator);
    Result := Copy(Result, 2, Length(Result));
  end else
    Result := '';
end;

function TSvcSyncBTPY2Execute.AddUpdateValues(AdoQryParam: AdoQry; DSType: T_WSDataService; FieldsList: string; KeyValue1, KeyValue2: string; IsRelatedParameters: boolean) : boolean;
var
  sIndex         : string;
  Y2FieldValue   : string;
  TmpData        : string;
  TmpFieldName   : string;
  TmpFieldValue  : string;
  BtpSendData    : string;
  Y2SendData     : string;
  CommonSendData : string;
  AddData        : string;
  Y2FieldsList   : string;
  TableName      : string;
  Y2StartPosData : integer;
  Cpt            : integer;
  RecordCount    : integer;
  SameValue      : boolean;

  function GetRequest(lAdoQry: AdoQry): string;
  var
    FormattedLastSynchro: string;
  begin
    FormattedLastSynchro := FormatDateTime('yyyymmdd hh:nn:ss', StrToDateTime(BTPValues.LastSynchro));
    case DSType of
      wsdsThird             : Result := Format('SELECT %s FROM TIERS WHERE T_AUXILIAIRE = ''%s'' AND T_DATEMODIF > ''%s'''                  , [lAdoQry.FieldsList, KeyValue1, FormattedLastSynchro]);
      wsdsBankId            : Result := Format('SELECT %s FROM RIB WHERE R_AUXILIAIRE = ''%s'' AND R_DATEMODIF > ''%s'''                    , [lAdoQry.FieldsList, KeyValue1, FormattedLastSynchro]);
      wsdsAnalyticalSection : Result := Format('SELECT %s FROM SECTION WHERE S_AXE = ''%s'' AND S_SECTION = ''%s'' AND S_DATEMODIF > ''%s''', [lAdoQry.FieldsList, KeyValue1, KeyValue2, FormattedLastSynchro]);
      wsdsContact           : Result := Format('SELECT %s FROM CONTACT WHERE C_AUXILIAIRE = ''%s'' AND C_DATEMODIF > ''%s'''                , [lAdoQry.FieldsList, KeyValue1, FormattedLastSynchro]);
      wsdsChoixCod          : Result := Format('SELECT %s FROM CHOIXCOD WHERE CC_TYPE IN (%s) ORDER BY CC_TYPE, CC_CODE'                    , [lAdoQry.FieldsList, KeyValue1]);
      wsdsChoixExt          : Result := Format('SELECT %s FROM CHOIXEXT WHERE YX_TYPE IN (%s) ORDER BY YX_TYPE, YX_CODE'                    , [lAdoQry.FieldsList, KeyValue1]);
      wsdsCurrency          : Result := Format('SELECT %s FROM DEVISE ORDER BY D_DEVISE'                                                    , [lAdoQry.FieldsList]);
      wsdsPaymenChoice      : Result := Format('SELECT %s FROM MODEREGL ORDER BY MR_MODEREGLE'                                              , [lAdoQry.FieldsList]);
      wsdsChangeRate        : Result := Format('SELECT %s FROM CHANCELL ORDER BY H_DEVISE, H_DATECOURS'                                     , [lAdoQry.FieldsList]);
      wsdsTaxRate           : Result := Format('SELECT %s FROM TXCPTTVA ORDER BY TV_TVAOUTPF, TV_CODETAUX, TV_REGIME'                       , [lAdoQry.FieldsList]);
    end;
  end;

  function GetY2DataFromCache(FieldsList, FieldsValue: string): string;
  var
    FieldsKey    : string;
    KeyField1    : string;
    KeyField2    : string;
    KeyValue1    : string;
    KeyValue2    : string;
    FieldName    : string;
    FieldValue   : string;
    IndexName    : string;
    lFieldsList  : string;
    lFieldsValue : string;
    CacheIndex   : integer;
  begin
    FieldsKey    := GetInfoFromDSType(wsidFieldsKey, DSType);
    KeyField1    := Tools.ReadTokenSt_(FieldsKey, ';');
    KeyField2    := Tools.ReadTokenSt_(FieldsKey, ';');
    lFieldsList  := FieldsList;
    lFieldsValue := FieldsValue;
    while lFieldsList <> '' do
    begin
      FieldName  := Tools.ReadTokenSt_(lFieldsList, ',');
      FieldValue := Tools.ReadTokenSt_(lFieldsValue, ToolsTobToTsl_Separator);
      case Tools.CaseFromString(FieldName, [KeyField1, KeyField2]) of
        {KeyField1} 0 : KeyValue1 := FieldValue;
        {KeyField2} 1 : KeyValue2 := FieldValue;
      end;
      if (KeyValue1 <> '') and (KeyValue2 <> '') then
        Break;
    end;
    IndexName  := Format('%s_%s%s', [GetInfoFromDSType(wsidTableName, DSType), KeyValue1, KeyValue2]);
    CacheIndex := TSlCacheGetY2Data.IndexOfName(IndexName);
    if CacheIndex > -1 then
      Result := copy(TSlCacheGetY2Data[CacheIndex], pos('=', TSlCacheGetY2Data[CacheIndex]) + 1, length(TSlCacheGetY2Data[CacheIndex]))
    else
      Result := '';
  end;

begin
  Result      := False;
  TableName   := GetInfoFromDSType(wsidTableName, DSType);
  RecordCount := 0;
  sIndex      := Tools.iif(IsRelatedParameters, Format('%s_%s_%s', [GetInfoFromDSType(wsidTableName, DSType), KeyValue1, KeyValue2]), GetInfoFromDSType(wsidTableName, DSType));
  if (Tools.CanInsertedInTable(TableName {$IFDEF APPSRV}, AdoQryBTP.ServerName, AdoQryBTP.DBName{$ENDIF !APPSRV})) // On a le droit d'exporter la table
    and (CanExportImportTableFromFrenquency(DSType))                                                               // On peut exporter par rapport � la fr�quence
    and ((TSlCacheSendAccParam.IndexOfName(sIndex) = -1)                                                           // L'enregistrement en cours n'a pas d�j� �t� ajout�
    and (((IsRelatedParameters) and (KeyValue1 <> '')) or (not IsRelatedParameters)))                              // Param�tres associ�s et Key <> '' ou autres param�tres                               
  then
  begin
    { Charge les enregistrements depuis la base BTP }
    AdoQryParam.TSLResult.Clear;
    AdoQryParam.FieldsList  := FieldsList;
    AdoQryParam.Request     := GetRequest(AdoQryParam);
    AdoQryParam.LogValues   := LogValues;
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.AddUpdateValues - AdoQryParam.Request : %s', [WSCDS_DebugMsg, AdoQryParam.Request]), ServiceName_BTPY2, LogValues, 0);
    try
      AdoQryParam.SingleTableSelect;
      if AdoQryParam.RecordCount > 0 then
      begin
        Y2FieldsList := GetFieldsList(GetTSlFromDSType(DSType, False));
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.AddUpdateValues - Y2FieldsList : %s', [WSCDS_DebugMsg, Y2FieldsList]), ServiceName_BTPY2, LogValues, 0);

        { Boucle sur les enregistrements BTP trouv�s et tests s'ils existent dans les datas provenant d'Y2}
        for Cpt := 0 to pred(AdoQryParam.TSLResult.count) do
        begin
          CommonSendData := '';
          BtpSendData    := SetDataToSend(AdoQryParam.FieldsList, AdoQryParam.TSLResult[Cpt]);
          Y2SendData     := GetY2DataFromCache(AdoQryParam.FieldsList, AdoQryParam.TSLResult[Cpt]);
          if LogValues.DebugEvents = 2 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.AddUpdateValues - BtpSendData : %s - Y2SendData : %s', [WSCDS_DebugMsg, BtpSendData, Y2SendData]), ServiceName_BTPY2, LogValues, 0);
          if BtpSendData <> Y2SendData then // Les donn�es ou colonnes sont diff�rentes, calcul des champs communs (ainsi que les donn�es) et on renvoie
          begin
            SameValue := (Y2SendData <> '');
            while BtpSendData <> '' do
            begin
              TmpData       := Tools.ReadTokenSt_(BtpSendData, ToolsTobToTsl_Separator);
              TmpFieldName  := Copy(TmpData, 1, Pos('=', TmpData) - 1);
              TmpFieldValue := Copy(TmpData, Pos('=', TmpData) + 1, Length(TmpData));
              AddData       := TmpFieldName + '=' + TmpFieldValue;
              if Pos(',' + TmpFieldName + ',', ',' + Y2FieldsList + ',') > 0 then
              begin
                CommonSendData := CommonSendData + ToolsTobToTsl_Separator + AddData;
                if (Y2SendData <> '') and (SameValue) then // S'il existe des donn�es issues de Y2, test si la valeur est la m�me
                begin
                  Y2StartPosData := Pos(TmpFieldName + '=', Y2SendData) + length(TmpFieldName) + 1;
                  while Y2SendData[Y2StartPosData] <> ToolsTobToTsl_Separator do
                  begin
                    Y2FieldValue := Y2FieldValue + Copy(Y2SendData, Y2StartPosData, 1);
                    Inc(Y2StartPosData);
                  end;
                  SameValue := (TmpFieldValue = Y2FieldValue);
                  Y2FieldValue := '';
                end;
              end;
            end;
            { Ajouter uniquement si pas les m�mes valeur dans les champs communs }
            if not SameValue then
            begin
              inc(RecordCount);
              Result         := True;
              CommonSendData := Copy(CommonSendData, 2, Length(CommonSendData));
              TSlCacheSendAccParam.Add(sIndex + '=' + CommonSendData);
            end;
          end;
        end;
        if (Result) and (not IsRelatedParameters) then
          TServicesLog.WriteLog(ssbylLog, Format('%s : %s enregistrement(s) � envoyer.', [TableName, IntToStr(RecordCount)]), ServiceName_BTPY2, LogValues, 4);
      end else
        TSlCacheSendAccParam.Add(sIndex + '=' + WSCDS_EmptyValue);
      AdoQryParam.Reset;
    except
      if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.AddUpdateValues. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryParam.ServerName, AdoQryParam.DBName, AdoQryParam.Request]), ServiceName_BTPY2, LogValues, 0);
      Result := False;
      exit;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.FieldExistInTSl(lTSl: TStringList; FieldName: string): boolean;
var
  Cpt: integer;
begin
  Result := False;
  if (lTSl.Count > 0) and (FieldName <> '') then
  begin
    for Cpt := 0 to pred(lTSl.Count) do
    begin
      Result := (Copy(lTSl[Cpt], 1, Pos(ToolsTobToTsl_Separator, lTSl[Cpt]) - 1) = FieldName);
      if Result then
        Break;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.GetTSlFromDSType(DSType: T_WSDataService; IsBTP: boolean): TStringList;
begin
  case DSType of
    wsdsThird              : Result := Tools.iif(IsBTP, TSlCacheThirdBTP        , TSlCacheThirdY2);
    wsdsAnalyticalSection  : Result := Tools.iif(IsBTP, TSlCacheSectionBTP      , TSlCacheSectionY2);
    wsdsAccount            : Result := Tools.iif(IsBTP, TSlCacheAcountBTP       , TSlCacheAcountY2);
    wsdsJournal            : Result := Tools.iif(IsBTP, TSlCacheJournalBTP      , TSlCacheJournalY2);
    wsdsBankId             : Result := Tools.iif(IsBTP, TSlCacheBankIdBTP       , TSlCacheBankIdY2);
    wsdsChoixCod           : Result := Tools.iif(IsBTP, TSlCacheChoixCodBTP     , TSlCacheChoixCodY2);
    wsdsCommon             : Result := Tools.iif(IsBTP, TSlCacheCommonBTP       , TSlCacheCommonY2);
    wsdsChoixExt           : Result := Tools.iif(IsBTP, TSlCacheChoixExtBTP     , TSlCacheChoixExtY2);
    wsdsRecovery           : Result := Tools.iif(IsBTP, TSlCacheRecoveryBTP     , TSlCacheRecoveryY2);
    wsdsCountry            : Result := Tools.iif(IsBTP, TSlCacheCountryBTP      , TSlCacheCountryY2);
    wsdsCurrency           : Result := Tools.iif(IsBTP, TSlCacheCurrencyBTP     , TSlCacheCurrencyY2);
    wsdsCorrespondence     : Result := Tools.iif(IsBTP, TSlCacheCorrespBTP      , TSlCacheCorrespY2);
    wsdsPaymenChoice       : Result := Tools.iif(IsBTP, TSlCachePaymentBTP      , TSlCachePaymentY2);
    wsdsChangeRate         : Result := Tools.iif(IsBTP, TSlCacheChangeRateBTP   , TSlCacheChangeRateY2);
    wsdsFiscalYear         : Result := Tools.iif(IsBTP, TSlCacheFiscalYearBTP   , TSlCacheFiscalYearY2);
    wsdsSocietyParameters  : Result := Tools.iif(IsBTP, TSlcacheSocietyParamBTP , TSlcacheSocietyParamY2);
    wsdsEstablishment      : Result := Tools.iif(IsBTP, TSlcacheEstablishmentBTP, TSlcacheEstablishmentY2);
    wsdsPaymentMode        : Result := Tools.iif(IsBTP, TSlcachePaymentModeBTP  , TSlcachePaymentModeY2);
    wsdsZipCode            : Result := Tools.iif(IsBTP, TSlcacheZipCodeBTP      , TSlcacheZipCodeY2);
    wsdsContact            : Result := Tools.iif(IsBTP, TSlcacheContactBTP      , TSlcacheContactY2);
    wsdsAccForPayment      : Result := Tools.iif(IsBTP, TSlCacheAccForPaymentBTP, TSlCacheAccForPaymentY2);
    wsdsAccForPaymentOther : Result := Tools.iif(IsBTP, TSlCacheAccForPaymentBTP, TSlCacheAccForPaymentY2);
    wsdsDocPayment         : Result := Tools.iif(IsBTP, TSlCacheDocPaymentBTP   , nil);
    wsdsTaxRate            : Result := Tools.iif(IsBTP, TSLCacheTaxRateBTP      , TSLCacheTaxRateY2);
  else
    Result := nil;
  end;
end;

function TSvcSyncBTPY2Execute.GetFieldTypeFromCache(DSType: T_WSDataService; FieldName: string): tTypeField;
var
  lTsl      : TStringList;
  Cpt       : integer;
  LineValue : string;
  FieldType : string;
begin
  Result := ttfNone;
  lTsl   := GetTSlFromDSType(DSType, True);
  for Cpt := 0 to pred(lTsl.Count) do
  begin
    LineValue := lTsl[Cpt];
    if Copy(LineValue, 1, Pos(ToolsTobToTsl_Separator, LineValue) - 1) = FieldName then
    begin
      Tools.ReadTokenSt_(LineValue, ToolsTobToTsl_Separator);
      FieldType := Tools.ReadTokenSt_(LineValue, ToolsTobToTsl_Separator);
      Result    := Tools.GetTypeFieldFromStringType(FieldType);
      Break;
    end;
  end;
end;

(* Le principe de g�n�ration du fichier TRA :
   - Pour chaque ligne de TSlCacheSendAccParam qui doit �tre envoy�e
     . appel de la cr�ation de la ligne TRA par rapport � la table � traiter
     . ajout de cette ligne dans un TSTringList au format TRA
   - Ecriture sur disque de ce fichier
   - Test de la taille (doit �tre inf�rieure � 4 Mo)
   - Si sup�rieur, d�coupe du fichier en x fichiers
   - Compression du/des fichiers
   - Envoie du/des fichiers
*)
function TSvcSyncBTPY2Execute.SendY2Settings: boolean;
var
  Cpt             : integer;
  LineValue       : string;
  TableName       : string;
  traLine         : string;
  TempPath        : string;
  TmpFileName     : string;
  PathFileName    : string;
  PathFileNameZip : string;
  traRCode        : T_TraRecordCode;
  TslTra          : TStringList;
  SendEntry       : TSendEntryY2;
  OnError         : Boolean;
  TraGeneration   : FileTraGeneration;

  function GetAdditionalDataFromThird(DSType: T_WSDataService): string;
  var
    Auxiliary     : string;
    AddValue      : string;
    IndexValue    : string;
    MainFieldName : string;
    Index         : Integer;
    Cpt           : Integer;
    IsMain        : Boolean;
  begin
    Result := '';
    Auxiliary := Tools.GetStValueFromTSl(LineValue, 'T_AUXILIAIRE');
    case DSType of
      wsdsBankId :
        begin
          IndexValue    := Format('RIB_%s_', [Auxiliary]);
          MainFieldName := 'R_PRINCIPAL';
        end;
      wsdsContact:
        begin
          IndexValue    := Format('CONTACT_%s_', [Auxiliary]);
          MainFieldName := 'C_PRINCIPAL';
        end;
    end;
    Index := TSlCacheSendAccParam.IndexOfName(IndexValue);
    if Index > -1 then
    begin
      AddValue := TSlCacheSendAccParam[Index];
      IsMain   := (Tools.GetStValueFromTSl(TSlCacheSendAccParam[Index], MainFieldName) = 'X');
      if IsMain then
        Result := TSlCacheSendAccParam[Index]
      else
      begin
        for Cpt := Index to pred(TSlCacheSendAccParam.count) do
        begin
          AddValue := TSlCacheSendAccParam[Index];
          if (Copy(AddValue, 1, Pos('=', AddValue) - 1) = IndexValue) and (Tools.GetStValueFromTSl(TSlCacheSendAccParam[Index], MainFieldName) = 'X') then
          begin
            Result := TSlCacheSendAccParam[Index];
            Exit;
          end;
        end;
      end;
    end;
  end;

begin
  if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.SendY2Settings', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
  Result := True;
  if TSlCacheSendAccParam.count > 0 then
  begin
    TraGeneration := FileTraGeneration.Create;
    try
      TServicesLog.WriteLog(ssbylLog, 'Pr�paration du fichier contenant le param�trage comptable.', ServiceName_BTPY2, LogValues, 3);
      TraGeneration.NewAna    := NewAna;
      TraGeneration.LogValues := LogValues;
      { G�n�ration du TRA }
      TslTra := TStringList.Create;
      try
        traLine := TraGeneration.GetTraFirstLine(trapS5, trafoCLI, trafcJRL, traffETE, BTPValues.UserAdmin);
        if traLine <> '' then
          TslTra.Add(traLine);
        for Cpt := 0 to pred(TSlCacheSendAccParam.count) do
        begin
          LineValue := TSlCacheSendAccParam[Cpt];
          if LogValues.DebugEvents = 2 then TServicesLog.WriteLog(ssbylLog, Format('%s - Before TraGeneration.GetTraLine - Table name : %s - LineValue : %s', [WSCDS_DebugMsg, TableName, LineValue]), ServiceName_BTPY2, LogValues, 0);
          if pos(WSCDS_EmptyValue, LineValue) = 0 then
          begin
            TableName := Copy(LineValue, 1, Pos('=', LineValue) - 1);
            if pos('_', TableName) > 0 then
              TableName := Copy(LineValue, 1, Pos('_', LineValue) - 1);
            traRCode := TraGeneration.GetTraRecordCodeFromTableName(TableName);
            case traRCode of
              {TIERS}             trarcThird:
                LineValue := LineValue + '#�**�#' + GetAdditionalDataFromThird(wsdsBankId) + '#�**�#' + GetAdditionalDataFromThird(wsdsContact);
            end;
            traLine := TraGeneration.GetTraLine(traRCode, LineValue);
            if traLine <> '' then
              TslTra.Add(traLine);
          end;
        end;
      finally
        TempPath    := GetEnvironmentVariable('TEMP');
        TmpFileName := Format('%s%s%s%s%s%s.TRA', [FormatDateTime('yyyy', Now), FormatDateTime('mm', Now), FormatDateTime('dd', Now), FormatDateTime('hh', Now), FormatDateTime('nn', Now), FormatDateTime('zzz', Now)]);
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - Before write %s\%s', [WSCDS_DebugMsg, TempPath, TmpFileName]), ServiceName_BTPY2, LogValues, 0);
        EcritLog(TempPath, TslTra.Text, TmpFileName);
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - After write %s\%s', [WSCDS_DebugMsg, TempPath, TmpFileName]), ServiceName_BTPY2, LogValues, 0);
        FreeAndNil(TslTra);
      end;
      PathFileName    := IncludeTrailingPathDelimiter(TempPath) + TmpFileName;
      PathFileNameZip := Tools.CompressFile(PathFileName);
      TSLTRAFileQty.Clear;
      Tools.FileCut(PathFileNameZip, Tools.GetKoFromMo(3), TSLTRAFileQty);
      if TGetParamWSCEGID.ConnectToY2 then
      begin
        SendEntry := TSendEntryY2.Create;
        try
          TServicesLog.WriteLog(ssbylLog, 'Envoi du param�trage comptable.', ServiceName_BTPY2, LogValues, 3);
          SendEntry.ServerName := BTPValues.Server;
          SendEntry.DBName     := BTPValues.DataBase;
          Result := SendEntry.SendAccountingParameters(TSLTRAFileQty, LogValues);
          try
          finally
            for Cpt := 0 to pred(SendEntry.TslResult.Count) do
            begin
              OnError := (pos(WSCDS_ErrorMsg, SendEntry.TslResult[Cpt]) > 0);
              TServicesLog.WriteLog(ssbylLog, SendEntry.TslResult[Cpt], ServiceName_BTPY2, LogValues, Tools.iif(OnError, 0, 4), False);
            end;
          end;
        finally
          SendEntry.Free;
        end;
      end else
        TServicesLog.WriteLog(ssbylLog, Format('%s : Les �l�ments de connexion Y2 sont incorrects.', [WSCDS_ErrorMsg]), ServiceName_BTPY2, LogValues, 0);
      DeleteFile(PathFileName);
    finally
      TraGeneration.Free;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.SendAccountingEntries(TslEntries : TStringList; var ErrorMsg : string) : integer;
var
  SendEntry : TSendEntryY2;
  DocInfo   : T_WSDocumentInf;
begin
  SendEntry := TSendEntryY2.Create;
  try
    SendEntry.ServerName := BTPValues.Server;
    SendEntry.DBName     := BTPValues.DataBase;
    DocInfo.dFromDoc     := False;
    ErrorMsg             := '';
    Result               := SendEntry.SendEntryCEGID(wsetDocument, TslEntries, DocInfo, ErrorMsg, LogValues);
  finally
    SendEntry.Free;
  end;
end;

procedure TSvcSyncBTPY2Execute.SearchRelatedParameters(TSlEntry: TStringList);
var
  AdoQryParam : AdoQry;
  Cpt         : integer;
  Values      : string;
  KeyValue1   : string;
  KeyValue2   : string;
begin
  AdoQryParam := AdoQry.Create;
  try
    AdoQryParam.ServerName := AdoQryBTP.ServerName;
    AdoQryParam.DBName     := AdoQryBTP.DBName;
    AdoQryParam.Qry        := TADOQuery.create(nil);
    for Cpt := 0 to pred(TSlEntry.count) do
    begin
      Values := TSlEntry[Cpt];
      if Copy(Values, 1, Length(ToolsTobToTsl_LevelName) + 1) = ToolsTobToTsl_LevelName + '2' then
      begin
        KeyValue1 := Tools.GetStValueFromTSl(Values, 'E_AUXILIAIRE');
        KeyValue2 := '';
        AddUpdateValues(AdoQryParam, wsdsThird  , GetFieldsList(TSlCacheThirdBTP)  , KeyValue1, KeyValue2, True);
        AddUpdateValues(AdoQryParam, wsdsBankId , GetFieldsList(TSlCacheBankIdBTP) , KeyValue1, KeyValue2, True);
        AddUpdateValues(AdoQryParam, wsdsContact, GetFieldsList(TSlcacheContactBTP), KeyValue1, KeyValue2, True);
      end
      else if Copy(Values, 1, Length(ToolsTobToTsl_LevelName) + 1) = ToolsTobToTsl_LevelName + '4' then // Export de la section de l'analytique courante
      begin
        KeyValue1 := Tools.GetStValueFromTSl(Values, 'Y_AXE');
        KeyValue2 := Tools.GetStValueFromTSl(Values, 'Y_SECTION');
        AddUpdateValues(AdoQryParam, wsdsAnalyticalSection, GetFieldsList(TSlCacheSectionBTP), KeyValue1, KeyValue2, True);
      end;
    end;
  finally
    AdoQryParam.Qry.Free;
    AdoQryParam.Free;
  end;
end;

procedure TSvcSyncBTPY2Execute.SearchOthersParameters;
var
  AdoQryParam : AdoQry;

  function GetKeyValue(DSType: T_WSDataService): string;
  var
    MultipleValues : string;
  begin
    case DSType of
      wsdsChoixCod, wsdsChoixExt:
        begin
          MultipleValues := TGetFromDSType.ExtractType(DSType);
          while MultipleValues <> '' do
            Result := Result + ',''' + Tools.ReadTokenSt_(MultipleValues, ';') + '''';
          Result := copy(Result, 2, Length(Result));
        end;
    else
      Result := '';
    end;
  end;

begin
  if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.SearchOthersParameters', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
  AdoQryParam := AdoQry.Create;
  try
    AdoQryParam.ServerName := AdoQryBTP.ServerName;
    AdoQryParam.DBName     := AdoQryBTP.DBName;
    AdoQryParam.Qry        := TADOQuery.create(nil);
    TServicesLog.WriteLog(ssbylLog, 'Tables de param�trage.', ServiceName_BTPY2, LogValues, 3);
    AddUpdateValues(AdoQryParam, wsdsChoixCod    , GetFieldsList(TSlCacheChoixCodBTP)  , GetKeyValue(wsdsChoixCod)    , '', False); // ChoixCod
    AddUpdateValues(AdoQryParam, wsdsChoixExt    , GetFieldsList(TSlCacheChoixExtBTP)  , GetKeyValue(wsdsChoixExt)    , '', False); // ChoixExt
    AddUpdateValues(AdoQryParam, wsdsCurrency    , GetFieldsList(TSlCacheCurrencyBTP)  , GetKeyValue(wsdsCurrency)    , '', False); // Devise
    AddUpdateValues(AdoQryParam, wsdsChangeRate  , GetFieldsList(TSlCacheChangeRateBTP), GetKeyValue(wsdsChangeRate)  , '', False); // Taux de change
    AddUpdateValues(AdoQryParam, wsdsPaymenChoice, GetFieldsList(TSlCachePaymentBTP)   , GetKeyValue(wsdsPaymenChoice), '', False); // Mode de r�glement
  finally
    AdoQryParam.Qry.Free;
    AdoQryParam.Free;
  end;
end;

procedure TSvcSyncBTPY2Execute.SetFilterFromDSType(DSType: T_WSDataService; TSl: TStringList);
var
  sCode     : string;
  Prefix    : string;
  TableName : string;
  First     : Boolean;
begin
  TableName := GetTableNameFromDSType(DSType);
  sCode     := TGetFromDSType.ExtractType(DSType);
  case DSType of
    wsdsThird:
      begin
        First := True;
        TSlFilter.Add(';;;(');
        while sCode <> '' do
        begin
          TSl.Add(Tools.iif(First, '', 'OR') + ';T_NATUREAUXI;=;' + Tools.ReadTokenSt_(sCode, ';'));
          First := False;
        end;
        TSlFilter.Add(';;;)');
        TSl.Add('AND;T_DATEMODIF;>=;' + BTPValues.LastSynchro);
      end;
    wsdsChoixCod, wsdsCommon, wsdsChoixExt:
      begin
        Prefix := TGetFromDSType.dstPrefix(DSType);
        First  := True;
        while sCode <> '' do
        begin
          TSl.Add(Tools.iif(First, '', 'OR') + ';' + Prefix + '_TYPE;=;' + Tools.ReadTokenSt_(sCode, ';'));
          First := False;
        end;
      end;
    wsdsRecovery:
      begin
        First := True;
        while sCode <> '' do
        begin
          TSl.Add(Tools.iif(First, '', 'OR') + ';RR_TYPERELANCE;=;' + Tools.ReadTokenSt_(sCode, ';'));
          First := False;
        end;
      end;
    wsdsContact:
      begin
        TSl.Add(';C_DATEMODIF;>=;' + BTPValues.LastSynchro);
        TSl.Add('AND;C_NATUREAUXI;=;CLI');
      end;
    wsdsAccForPayment, wsdsAccForPaymentOther :
      begin
        { Etat du lettrage }
        sCode := Tools.iif(DSType = wsdsAccForPayment, 'TL;PL', 'RI;AL');
        First := True;
        TSlFilter.Add(';;;(');
        while sCode <> '' do
        begin
          TSl.Add(Tools.iif(First, '', 'OR') + ';E_ETATLETTRAGE;=;' + Tools.ReadTokenSt_(sCode, ';'));
          First := False;
        end;
        TSlFilter.Add(';;;)');
        { Uniquement r�glement }
        if DSType = wsdsAccForPaymentOther then
        begin
          TSlFilter.Add('AND;;;');
          sCode := 'RC;RF';
          First := True;
          TSlFilter.Add(';;;(');
          while sCode <> '' do
          begin
            TSl.Add(Tools.iif(First, '', 'OR') + ';E_NATUREPIECE;=;' + Tools.ReadTokenSt_(sCode, ';'));
            First := False;
          end;
          TSlFilter.Add(';;;)');
        end;
        TSl.Add('AND;E_QUALIFPIECE;=;N;');
        TSl.Add('AND;E_DATEMODIF;>=;' + BTPValues.LastSynchro);
      end;
    wsdsAnalyticalSection : TSl.Add(';CSP_DATEMODIF;>=;' + BTPValues.LastSynchro);
    wsdsAccount           : TSl.Add(';G_DATEMODIF;>=;'   + BTPValues.LastSynchro);
    wsdsJournal           : TSl.Add(';J_DATEMODIF;>=;'   + BTPValues.LastSynchro);
    wsdsBankId            : TSl.Add(';R_DATEMODIF;>=;'   + BTPValues.LastSynchro);
    wsdsChangeRate        : TSl.Add(';H_DATECOURS;>=;'   + BTPValues.LastSynchro);
    wsdsEstablishment     : TSl.Add(';ET_DATEMODIF;>=;'  + BTPValues.LastSynchro);
    wsdsCountry           : TSl.Add(WSCDS_EmptyValue);
    wsdsCurrency          : TSl.Add(WSCDS_EmptyValue);
    wsdsCorrespondence    : TSl.Add(WSCDS_EmptyValue);
    wsdsPaymenChoice      : TSl.Add(WSCDS_EmptyValue);
    wsdsFiscalYear        : TSl.Add(WSCDS_EmptyValue);
    wsdsSocietyParameters : TSl.Add(WSCDS_EmptyValue);
    wsdsPaymentMode       : TSl.Add(WSCDS_EmptyValue);
    wsdsZipCode           : TSl.Add(WSCDS_EmptyValue);
    wsdsTaxRate           : TSl.Add(WSCDS_EmptyValue);
  end;
end;

function TSvcSyncBTPY2Execute.GetInfoFromDSType(InfoType: T_WSInfoFromDSType; DSType: T_WSDataService; FieldName: string = ''): string;
var
  TableName : string;
  Prefix    : string;
  Value     : string;
  Value1    : string;
  Value2    : string;
  lDate     : TDateTime;

  function GetValueFrom(lTSlIndice: TStringList; FieldName: string): string;
  var
    iCpt    : integer;
    lValues : string;
  begin
    Result := '';
    for iCpt := 0 to Pred(lTSlIndice.Count) do
    begin
      lValues := lTSlIndice[iCpt];
      if Copy(lValues, 1, Pos('=', lValues) - 1) = FieldName then
      begin
        Result := Copy(lValues, Pos('=', lValues) + 1, Length(lValues));
        Break;
      end;
    end;
  end;

begin
  TableName := GetTableNameFromDSType(DSType);
  case DSType of
    wsdsThird:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'T_AUXILIAIRE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';T_NATUREAUXI;T_AUXILIAIRE;T_TIERS;') > 0));
        wsidFieldsList    : Result := 'T_AUXILIAIRE';
        wsidRequest       : Result := Format('SELECT %s FROM %s WHERE T_AUXILIAIRE = ''%s''', [AdoQryBTP.FieldsList, TableName, GetValueFrom(TSlIndice, AdoQryBTP.FieldsList)]);
      else
        Result := '';
      end;
    end;
    wsdsAnalyticalSection:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'S_AXE;S_SECTION';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';S_AXE;S_SECTION;') > 0));
        wsidFieldsList    : Result := 'S_SECTION';
        wsidRequest:
          begin
            Value  := GetValueFrom(TSlIndice, 'S_AXE');
            Value1 := GetValueFrom(TSlIndice, 'S_SECTION');
            Result := Format('SELECT S_SECTION FROM %s WHERE S_AXE = ''%s'' AND S_SECTION = ''%s''', [TableName, Value, Value1]);
          end;
      else
        Result := '';
      end;
    end;
    wsdsAccount:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'G_GENERAL';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';G_GENERAL;') > 0));
        wsidFieldsList    : Result := 'G_GENERAL';
        wsidRequest       : Result := Format('SELECT %s FROM %s WHERE G_GENERAL = ''%s''', [AdoQryBTP.FieldsList, TableName, GetValueFrom(TSlIndice, AdoQryBTP.FieldsList)]);
      else
        Result := '';
      end;
    end;
    wsdsJournal:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'J_JOURNAL';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';J_JOURNAL;') > 0));
        wsidFieldsList    : Result := 'J_JOURNAL';
        wsidRequest       : Result := Format('SELECT J_JOURNAL FROM %s WHERE J_JOURNAL = ''%s''', [TableName, GetValueFrom(TSlIndice, 'J_JOURNAL')]);
      else
        Result := '';
      end;
    end;
    wsdsBankId :
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'R_AUXILIAIRE;R_NUMERORIB';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';R_AUXILIAIRE;R_NUMERORIB;') > 0));
        wsidFieldsList    : Result := 'R_AUXILIAIRE';
        wsidRequest:
          begin
            Value  := GetValueFrom(TSlIndice, 'R_AUXILIAIRE');
            Value1 := GetValueFrom(TSlIndice, 'R_NUMERORIB');
            Result := Format('SELECT R_AUXILIAIRE FROM %s WHERE R_AUXILIAIRE = ''%s'' AND R_NUMERORIB = ''%s''', [TableName, Value, Value1]);
          end;
      else
        Result := '';
      end;
    end;
    wsdsChoixCod, wsdsCommon, wsdsChoixExt:
    begin
      Prefix := TGetFromDSType.dstPrefix(DSType);
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := Prefix + '_TYPE;' + Prefix + '_CODE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';' + Prefix + '_TYPE;' + Prefix + '_CODE;') > 0));
        wsidFieldsList    : Result := Prefix + '_TYPE';
        wsidRequest:
          begin
            Value  := GetValueFrom(TSlIndice, Prefix + '_TYPE');
            Value1 := GetValueFrom(TSlIndice, Prefix + '_CODE');
            Result := Format('SELECT ' + Prefix + '_TYPE FROM %s WHERE ' + Prefix + '_TYPE = ''%s'' AND ' + Prefix + '_CODE = ''%s''', [TableName, Value, Value1]);
          end;
      else
        Result := '';
      end;
    end;
    wsdsRecovery:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'RR_TYPERELANCE;RR_FAMILLERELANCE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';RR_TYPERELANCE;RR_FAMILLERELANCE;') > 0));
        wsidFieldsList    : Result := 'RR_TYPERELANCE';
        wsidRequest:
          begin
            Value  := GetValueFrom(TSlIndice, 'RR_TYPERELANCE');
            Value1 := GetValueFrom(TSlIndice, 'RR_FAMILLERELANCE');
            Result := Format('SELECT RR_TYPERELANCE FROM %s WHERE RR_TYPERELANCE = ''%s'' AND RR_FAMILLERELANCE = ''%s''', [TableName, Value, Value1]);
          end;
      else
        Result := '';
      end;
    end;
    wsdsCountry:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'PY_PAYS';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';PY_PAYS;') > 0));
        wsidFieldsList    : Result := 'PY_PAYS';
        wsidRequest       : Result := Format('SELECT PY_PAYS FROM %s WHERE PY_PAYS = ''%s''', [TableName, GetValueFrom(TSlIndice, 'PY_PAYS')]);
      else
        Result := '';
      end;
    end;
    wsdsCurrency:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'D_DEVISE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';D_DEVISE;') > 0));
        wsidFieldsList    : Result := 'D_DEVISE';
        wsidRequest       : Result := Format('SELECT D_DEVISE FROM %s WHERE D_DEVISE = ''%s''', [TableName, GetValueFrom(TSlIndice, 'D_DEVISE')]);
      else
        Result := '';
      end;
    end;
    wsdsCorrespondence:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'CR_TYPE;CR_CORRESP';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';CR_TYPE;CR_CORRESP;') > 0));
        wsidFieldsList    : Result := 'CR_CORRESP';
        wsidRequest:
          begin
            Value  := GetValueFrom(TSlIndice, 'CR_TYPE');
            Value1 := GetValueFrom(TSlIndice, 'CR_CORRESP');
            Result := Format('SELECT CR_CORRESP FROM %s WHERE CR_TYPE = ''%s'' AND CR_CORRESP = ''%s''', [TableName, Value, Value1]);
          end;
      else
        Result := '';
      end;
    end;
    wsdsPaymenChoice:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'MR_MODEREGLE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';MR_MODEREGLE;') > 0));
        wsidFieldsList    : Result := 'MR_MODEREGLE';
        wsidRequest       : Result := Format('SELECT MR_MODEREGLE FROM %s WHERE MR_MODEREGLE = ''%s''', [TableName, GetValueFrom(TSlIndice, 'MR_MODEREGLE')]);
      else
        Result := '';
      end;
    end;
    wsdsChangeRate:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'H_DEVISE;H_DATECOURS';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';H_DEVISE;H_DATECOURS;') > 0));
        wsidFieldsList    : Result := 'H_DEVISE';
        wsidRequest:
          begin
            Value  := GetValueFrom(TSlIndice, 'H_DEVISE');
            Value1 := Tools.SetStrDateTimeFromStrUTCDateTime(GetValueFrom(TSlIndice, 'H_DATECOURS'));
            lDate  := IncDay(StrToDateTime(Value1));
            Value2 := DateTimeToStr(lDate);
            Value1 := Tools.CastDateTimeForQry(StrToDateTime(Value1));
            Value2 := Tools.CastDateTimeForQry(StrToDateTime(Value2));
            Result := Format('SELECT H_DEVISE FROM %s WHERE H_DEVISE = ''%s'' AND H_DATECOURS >= ''%s'' AND H_DATECOURS < ''%s''', [TableName, Value, Value1, Value2]);
          end;
      else
        Result := '';
      end;
    end;
    wsdsFiscalYear:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'EX_ENTITY;EX_EXERCICE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';EX_ENTITY;EX_EXERCICE;') > 0));
        wsidFieldsList    : Result := 'EX_EXERCICE';
        wsidRequest:
        begin
          Value  := GetValueFrom(TSlIndice, 'EX_ENTITY');
          Value1 := GetValueFrom(TSlIndice, 'EX_EXERCICE');
          Result := Format('SELECT EX_EXERCICE FROM %s WHERE EX_ENTITY = ''%s'' AND EX_EXERCICE = ''%s''', [TableName, Value, Value1]);
        end;
      else
        Result := '';
      end;
    end;
    wsdsSocietyParameters:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'SOC_NOM';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';SOC_NOM;') > 0));
        wsidFieldsList    : Result := 'SOC_NOM';
        wsidRequest       : Result := Format('SELECT SOC_NOM FROM %s WHERE SOC_NOM = ''%s''', [TableName, GetValueFrom(TSlIndice, 'SOC_NOM')]);
      else
        Result := '';
      end;
    end;
    wsdsEstablishment:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'ET_ETABLISSEMENT';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';ET_ETABLISSEMENT;') > 0));
        wsidFieldsList    : Result := 'ET_ETABLISSEMENT';
        wsidRequest       : Result := Format('SELECT ET_ETABLISSEMENT FROM %s WHERE ET_ETABLISSEMENT = ''%s''', [TableName, GetValueFrom(TSlIndice, 'ET_ETABLISSEMENT')]);
      else
        Result := '';
      end;
    end;
    wsdsPaymentMode:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'MP_MODEPAIE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';MP_MODEPAIE;') > 0));
        wsidFieldsList    : Result := 'MP_MODEPAIE';
        wsidRequest       : Result := Format('SELECT MP_MODEPAIE FROM %s WHERE MP_MODEPAIE = ''%s''', [TableName, GetValueFrom(TSlIndice, 'MP_MODEPAIE')]);
      else
        Result := '';
      end;
    end;
    wsdsZipCode:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'O_CODEPOSTAL;O_VILLE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';O_CODEPOSTAL;O_VILLE;') > 0));
        wsidFieldsList    : Result := 'O_CODEPOSTAL';
        wsidRequest:
          begin
            Value  := GetValueFrom(TSlIndice, 'O_CODEPOSTAL');
            Value1 := GetValueFrom(TSlIndice, 'O_VILLE');
            if Pos('''', Value1) > 0 then
              Value1 := StringReplace(Value1, '''', '''''', [rfReplaceAll]);
            Result := Format('SELECT O_CODEPOSTAL FROM %s WHERE O_CODEPOSTAL = ''%s'' AND O_VILLE = ''%s''', [TableName, Value, Value1]);
          end;
      else
        Result := '';
      end;
    end;
    wsdsContact:
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'C_TYPECONTACT;C_AUXILIAIRE;C_NUMEROCONTACT';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';C_TYPECONTACT;C_AUXILIAIRE;C_NUMEROCONTACT;') > 0));
        wsidFieldsList    : Result := 'C_NUMEROCONTACT';
        wsidRequest:
          begin
            Value  := GetValueFrom(TSlIndice, 'C_TYPECONTACT');
            Value1 := GetValueFrom(TSlIndice, 'C_AUXILIAIRE');
            Value2 := GetValueFrom(TSlIndice, 'C_NUMEROCONTACT');
            Result := Format('SELECT C_NUMEROCONTACT FROM %s WHERE C_TYPECONTACT = ''%s'' AND C_AUXILIAIRE = ''%s'' AND C_NUMEROCONTACT = ''%s''', [TableName, Value, Value1, Value2]);
          end;
      else
        Result := '';
      end;
    end;
    wsdsAccForPayment :
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'E_ENTITY;E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE;E_QUALIFPIECE';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';E_ENTITY;E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE;E_QUALIFPIECE;') > 0));
        wsidFieldsList    : Result := 'E_NUMEROPIECE';
        wsidRequest       : Result := '';
      else
        Result := '';
      end;
    end;
    wsdsTaxRate :
    begin
      case InfoType of
        wsidTableName     : Result := TableName;
        wsidFieldsKey     : Result := 'TV_TVAOUTPF;TV_CODETAUX;TV_REGIME';
        wsidExcludeFields : Result := BoolToStr((pos(';' + FieldName + ';', ';TV_TVAOUTPF;TV_CODETAUX;TV_REGIME;') > 0));
        wsidFieldsList    : Result := 'TV_CODETAUX';
        wsidRequest:
        begin
          Value  := GetValueFrom(TSlIndice, 'TV_TVAOUTPF');
          Value1 := GetValueFrom(TSlIndice, 'TV_CODETAUX');
          Value2 := GetValueFrom(TSlIndice, 'TV_REGIME');
          Result := Format('SELECT TV_CODETAUX FROM %s WHERE TV_TVAOUTPF = ''%s'' AND TV_CODETAUX = ''%s'' AND TV_REGIME = ''%s''', [TableName, Value, Value1, Value2]);
        end;
      else
        Result := '';
      end;
    end;
  else
    Result := '';
  end;
end;

function TSvcSyncBTPY2Execute.GetSingleValueFromTSl(FieldName: string; TslOrig : TStringList): string;
var
  iCpt    : integer;
  lValues : string;
begin
  Result := '';
  for iCpt := 0 to Pred(TslOrig.count) do
  begin
    lValues := TslOrig[iCpt];
    if Copy(lValues, 1, Pos('=', lValues) - 1) = FieldName then
    begin
      Result := Copy(lValues, Pos('=', lValues) + 1, Length(lValues));
      Break;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.GetMultipleValueFromTSl(FieldsName : string; TslOrig : TStringList): string;
var
  iCpt           : integer;
  iCptArr        : integer;
  arrLength      : integer;
  FindQty        : integer;
  lValues        : string;
  lFieldsName    : string;
  lFieldName     : string;
  lFieldValue    : string;
  arrFieldsName  : array of string;
  arrFieldsValue : array of string;
begin
  if pos(';', FieldsName) > 0 then
  begin
    Result := '';
    lFieldsName := FieldsName;
    arrLength   := 0;
    FindQty     := 0;
    for iCpt := 1 to length(lFieldsName)  do
    begin
      if FieldsName[iCpt] = ';' then
       inc(arrLength);
    end;
    inc(arrLength);
    SetLength(arrFieldsName , arrLength);
    SetLength(arrFieldsValue, arrLength);
    arrLength := 0;
    while lFieldsName <> '' do
    begin
      arrFieldsName[arrLength] := Tools.ReadTokenSt_(lFieldsName, ';');
      inc(arrLength);
    end;
    for iCpt := 0 to Pred(TslOrig.count) do
    begin
      lValues     := TslOrig[iCpt];
      lFieldName  := Copy(lValues, 1, Pos('=', lValues) - 1);
      lFieldValue := Copy(lValues, Pos('=', lValues) + 1, Length(lValues));
      for iCptArr := Low(arrFieldsName) to High(arrFieldsName) do
      begin
        if lFieldName = arrFieldsName[iCptArr] then
        begin
          inc(FindQty);
          arrFieldsValue[iCptArr] := lFieldValue;
        end;
      end;
      if FindQty = High(arrFieldsName) + 1 then
        break;
    end;
    for iCpt := 0 to High(arrFieldsValue) do
      Result := Result + ToolsTobToTsl_Separator + arrFieldsValue[iCpt];
    Result := copy(Result, 2, length(Result));
  end else
    Result := GetSingleValueFromTSl(FieldsName, TslOrig);
end;

function TSvcSyncBTPY2Execute.GetTableNameFromDSType(DSType: T_WSDataService): string;
begin
  case DSType of
    wsdsThird             : Result := Tools.GetTableNameFromTtn(ttnTiers);
    wsdsAnalyticalSection : Result := Tools.GetTableNameFromTtn(ttnSection);
    wsdsAccount           : Result := Tools.GetTableNameFromTtn(ttnGeneraux);
    wsdsJournal           : Result := Tools.GetTableNameFromTtn(ttnJournal);
    wsdsBankId            : Result := Tools.GetTableNameFromTtn(ttnRib);
    wsdsChoixCod          : Result := Tools.GetTableNameFromTtn(ttnChoixCod);
    wsdsCommon            : Result := Tools.GetTableNameFromTtn(ttnCommun);
    wsdsChoixExt          : Result := Tools.GetTableNameFromTtn(ttnChoixExt);
    wsdsRecovery          : Result := Tools.GetTableNameFromTtn(ttnRelance);
    wsdsCountry           : Result := Tools.GetTableNameFromTtn(ttnPays);
    wsdsCurrency          : Result := Tools.GetTableNameFromTtn(ttnDevise);
    wsdsChangeRate        : Result := Tools.GetTableNameFromTtn(ttnChancell);
    wsdsCorrespondence    : Result := Tools.GetTableNameFromTtn(ttnCorresp);
    wsdsPaymenChoice      : Result := Tools.GetTableNameFromTtn(ttnModeRegl);
    wsdsFiscalYear        : Result := Tools.GetTableNameFromTtn(ttnExercice);
    wsdsSocietyParameters : Result := Tools.GetTableNameFromTtn(ttnParamSoc);
    wsdsEstablishment     : Result := Tools.GetTableNameFromTtn(ttnEtabliss);
    wsdsPaymentMode       : Result := Tools.GetTableNameFromTtn(ttnModePaie);
    wsdsZipCode           : Result := Tools.GetTableNameFromTtn(ttnCodePostaux);
    wsdsContact           : Result := Tools.GetTableNameFromTtn(ttnContact);
    wsdsAccForPayment     : Result := Tools.GetTableNameFromTtn(ttnEcriture);
    wsdsDocPayment        : Result := Tools.GetTableNameFromTtn(ttnAcomptes);
    wsdsTaxRate           : Result := Tools.GetTableNameFromTtn(ttnTxCptTva);         
  else
    Result := '';
  end;
end;

function TSvcSyncBTPY2Execute.GetDSTypeFromTableName(TableName : string) : T_WSDataService;
begin
  case Tools.CaseFromString(TableName, [  Tools.GetTableNameFromTtn(ttnTiers)      , Tools.GetTableNameFromTtn(ttnSection) , Tools.GetTableNameFromTtn(ttnGeneraux)
                                        , Tools.GetTableNameFromTtn(ttnJournal)    , Tools.GetTableNameFromTtn(ttnRib)     , Tools.GetTableNameFromTtn(ttnChoixCod)
                                        , Tools.GetTableNameFromTtn(ttnCommun)     , Tools.GetTableNameFromTtn(ttnChoixExt), Tools.GetTableNameFromTtn(ttnRelance)
                                        , Tools.GetTableNameFromTtn(ttnPays)       , Tools.GetTableNameFromTtn(ttnDevise)  , Tools.GetTableNameFromTtn(ttnChancell)
                                        , Tools.GetTableNameFromTtn(ttnCorresp)    , Tools.GetTableNameFromTtn(ttnModeRegl), Tools.GetTableNameFromTtn(ttnExercice)
                                        , Tools.GetTableNameFromTtn(ttnParamSoc)   , Tools.GetTableNameFromTtn(ttnEtabliss), Tools.GetTableNameFromTtn(ttnModePaie)
                                        , Tools.GetTableNameFromTtn(ttnCodePostaux), Tools.GetTableNameFromTtn(ttnContact), Tools.GetTableNameFromTtn(ttnEcriture)
                                        , Tools.GetTableNameFromTtn(ttnAcomptes)   , Tools.GetTableNameFromTtn(ttnTxCptTva)
                                       ]) of
    {ttnTiers}              0  : Result := wsdsThird;
    {ttnSection}            1  : Result := wsdsAnalyticalSection;
    {ttnGeneraux}           2  : Result := wsdsAccount;
    {wsdsJournal}           3  : Result := wsdsJournal;
    {wsdsBankId}            4  : Result := wsdsBankId;
    {wsdsChoixCod}          5  : Result := wsdsChoixCod;
    {wsdsCommon}            6  : Result := wsdsCommon;
    {wsdsChoixExt}          7  : Result := wsdsChoixExt;
    {wsdsRecovery}          8  : Result := wsdsRecovery;
    {wsdsCountry}           9  : Result := wsdsCountry;
    {wsdsCurrency}          10 : Result := wsdsCurrency;
    {wsdsChangeRate}        11 : Result := wsdsChangeRate;
    {wsdsCorrespondence}    12 : Result := wsdsCorrespondence;
    {wsdsPaymenChoice}      13 : Result := wsdsPaymenChoice;
    {wsdsFiscalYear}        14 : Result := wsdsFiscalYear;
    {wsdsSocietyParameters} 15 : Result := wsdsSocietyParameters;
    {wsdsEstablishment}     16 : Result := wsdsEstablishment;
    {wsdsPaymentMode}       17 : Result := wsdsPaymentMode;
    {wsdsZipCode}           18 : Result := wsdsZipCode;
    {wsdsContact}           19 : Result := wsdsContact;
    {wsdsAccForPayment}     20 : Result := wsdsAccForPayment;
    {wsdsDocPayment}        21 : Result := wsdsDocPayment;
    {wsdsTaxRate}           22 : Result := wsdsTaxRate;
  else
    Result := wsdsNone;
  end;
end;

procedure TSvcSyncBTPY2Execute.ExtractIndice(Cpt: integer; TSlOrig, TSlResult: TStringList);
var
  sIndex      : string;
  Value       : string;
  Pos         : integer;
  iCpt        : integer;
  IndexChange : Boolean;
begin
  sIndex := WSCDS_IndiceField + IntToStr(Cpt) + '#';
  Pos    := TSlOrig.IndexOfName(sIndex);
  if Pos > -1 then
  begin
    sIndex := TSlOrig[Pos];
    for iCpt := Pos to Pred(TSlOrig.Count) do
    begin
      Value       := TSlOrig[iCpt];
      IndexChange := ((Copy(Value, 1, 7) = WSCDS_IndiceField) and (Value <> sIndex));
      if not IndexChange then
      begin
        { Sp�cif pour SECTION, pr�fixe Y2=CSP_, pr�fixe BTP=S_ }
        if Copy(Value, 1, 4) = 'CSP_' then
          Value := StringReplace(Value, 'CSP_', 'S_', [rfReplaceAll]);
        TSlResult.Add(Value);
      end else
        Break;
    end;
  end;
end;
//{$IFEND !APPSRV}

function TSvcSyncBTPY2Execute.CanExportImportTableFromFrenquency(DSType: T_WSDataService) : boolean;
var
  Index       : Integer;
  LastSynchro : TDate;
  Action      : string;
  TableName   : string;
begin
  if (DSType = wsdsAccForPayment) or (DSType = wsdsAccForPaymentOther) then
    Index  := TSLCacheUpdateFrequencySetting.IndexOfName(TGetFromDSType.dstTableName(DSType))
  else
    Index  := TSLCacheUpdateFrequencySetting.IndexOfName(TGetFromDSType.dstTableName(DSType));
  Result := (Index = -1);
  if not Result then
  begin
    LastSynchro := Int(StrToDateTime(BTPValues.LastSynchro));
    TableName   := Copy(TSLCacheUpdateFrequencySetting[Index], 1, Pos('=', TSLCacheUpdateFrequencySetting[Index]) - 1);
    Action      := Trim(Copy(TSLCacheUpdateFrequencySetting[Index], Pos('=', TSLCacheUpdateFrequencySetting[Index]) + 1, Length(TSLCacheUpdateFrequencySetting[Index])));
    case Tools.CaseFromString(UpperCase(Action), ['DAILY', 'MONTHLY', 'ANNUAL', 'ONCE', 'EVERYTIME', 'NEVER']) of
      {DAILY}     0: Result := (LastSynchro < Now);
      {MONTHLY}   1: Result := (MonthOf(LastSynchro) < MonthOf(Now));
      {ANNUAL}    2: Result := (YearOf(LastSynchro) < YearOf(Now));
      {ONCE}      3: Result := (LastSynchro <= 2);
      {EVERYTIME} 4: Result := True;
      {NEVER}     5: Result := False;
    else
      begin
        TServicesLog.WriteLog(ssbylLog, Format('ATTENTION, la fr�quence "%s" associ�e � la table "%s" du fichier de configuration est inconnue. Le traitement n''a pas �t� effectu� pour cette derni�re.'
                                              , [Action, TableName]), ServiceName_BTPY2, LogValues, 3);
        Result := False;
      end;
    end;
  end;
end;

function TSvcSyncBTPY2Execute.AddData(wsAction: T_WSAction; DSType: T_WSDataService; lTSlValues: TStringList): boolean;
var
  CptUI             : integer;
  KeyValue          : string;
  KeyValuel         : string;
  FieldName         : string;
  FieldValue        : string;
  Sql               : string;
  Where             : string;
  InsertedFields    : string;
  InsertedValues    : string;
  CacheLine         : string;
  TSlInsertedFields : TStringList;
  FieldType         : tTypeField;
begin
  Result         := True;
  KeyValue       := GetKeyValues(DSType, lTSlValues);
  InsertedFields := '';
  InsertedValues := '';
  if LogValues.LogLevel >= 2 then
    TServicesLog.WriteLog(ssbylLog, Format('%s de "%s"', [Tools.iif(wsAction = wsacUpdate, 'Modification', 'Cr�ation'), KeyValue]), ServiceName_BTPY2, LogValues, 4);
  for CptUI := 0 to pred(lTSlValues.Count) do
  begin
    if copy(lTSlValues[CptUI], 1, 7) <> WSCDS_IndiceField then
    begin                                                                                            
      FieldName := copy(lTSlValues[CptUI], 1, Pos('=', lTSlValues[CptUI]) - 1);
      FieldType := GetFieldTypeFromCache(DSType, FieldName);
      if (FieldExistInTSl(GetTSlFromDSType(DSType, True), FieldName))             // Le champ existe dans BTP
        and ((wsAction = wsacInsert)                                              // On est en Insert
             or ((wsAction = wsacUpdate)                                          // On est en Update
                 and (GetInfoFromDSType(wsidExcludeFields, DSType, FieldName) = '0')) //  et le champs ne fait pas parti des champs � exclure
            )
      then
      begin
        FieldValue := ValueAnalysis(lTSlValues[CptUI], FieldName, FieldType);
        FieldValue := AddQuotes(DSType, FieldName, FieldValue);
        Sql := Sql + Format(', %s=%s', [FieldName, FieldValue]);
        InsertedFields := InsertedFields + ', ' + FieldName;
        InsertedValues := InsertedValues + ', ' + FieldValue;
      end;
    end;
  end;
  if wsAction = wsacInsert then  // Ajouter tous les autres champs
  begin
    TSlInsertedFields := GetTSlFromDSType(DSType, True);
    if Assigned(TSlInsertedFields) then
    begin
      for CptUI := 0 to pred(TSlInsertedFields.Count) do
      begin
        CacheLine := TSlInsertedFields[CptUI];
        FieldName := Tools.ReadTokenSt_(CacheLine, ToolsTobToTsl_Separator);
        Tools.ReadTokenSt_(CacheLine, ToolsTobToTsl_Separator);
        FieldValue := Tools.ReadTokenSt_(CacheLine, ToolsTobToTsl_Separator);
        if Pos(FieldName, InsertedFields) = 0 then
        begin
          InsertedFields := InsertedFields + ', ' + FieldName;
          InsertedValues := InsertedValues + ', ' + FieldValue;
        end;
      end;
    end;
  end;
  Sql            := Copy(Sql, 2, Length(Sql));
  InsertedFields := Copy(InsertedFields, 2, Length(InsertedFields));
  InsertedValues := Copy(InsertedValues, 2, Length(InsertedValues));
  if wsAction = wsacUpdate then
  begin
    KeyValuel := KeyValue;
    while KeyValuel <> '' do
    begin
      FieldName := Tools.ReadTokenSt_(KeyValuel, '-');
      FieldType := GetFieldTypeFromCache(DSType, copy(FieldName, 1, pos('=', FieldName)- 1));
      if FieldType = ttfDate then
      begin
        FieldValue := copy(FieldName, pos('=', FieldName)+ 2, 10); //length(FieldName) -1);
        FieldValue := Format('''%s''', [Tools.CastDateForQry(StrToDate(FieldValue))]);
        FieldName  := copy(FieldName, 1, pos('=', FieldName)) + FieldValue;
      end;
      Where := Where + ' AND ' + FieldName; //Tools.ReadTokenSt_(KeyValuel, '-');
    end;
    Where := Copy(Where, 5, length(Where));
    Sql := Format('UPDATE %s SET %s WHERE %s', [GetInfoFromDSType(wsidTableName, DSType), Sql, Where]);
  end else
    Sql := Format('INSERT INTO %s (%s) VALUES(%s)', [GetInfoFromDSType(wsidTableName, DSType), InsertedFields, InsertedValues]);
  TSLUpdateInsertData.Add(Sql);
end;

function TSvcSyncBTPY2Execute.GetKeyValues(DSType: T_WSDataService; lTSlValues: TStringList): string;
var
  FieldsKey  : string;
  Value      : string;
  FieldName  : string;
  FieldValue : string;
  sDate      : string;
  sHour      : string;
  Posit      : integer;
  iDateTime  : TDateTime;
begin
  Result    := '';
  FieldsKey := GetInfoFromDSType(wsidFieldsKey, DSType);
  while FieldsKey <> '' do
  begin
    Posit := lTSlValues.IndexOfName(Tools.ReadTokenSt_(FieldsKey, ';'));
    if Posit > -1 then
    begin
      FieldName  := copy(lTSlValues[Posit], 1, Pos('=', lTSlValues[Posit]) - 1);
      FieldValue := copy(lTSlValues[Posit], Pos('=', lTSlValues[Posit]) + 1, length(lTSlValues[Posit]));
      case GetFieldTypeFromCache(DSType, FieldName) of
        ttfCombo : FieldValue := Trim(FieldValue);
        ttfDate  : begin
                     FieldValue := Tools.SetStrDateTimeFromStrUTCDateTime(FieldValue);
                     iDateTime  := StrToDateTime(FieldValue);
                     if iDateTime - trunc(iDateTime) > 0 then // Il y a des heures
                     begin
                       sDate := copy(FieldValue, 1, pos(' ', FieldValue) -1);
                       sHour := copy(FieldValue, pos(' ', FieldValue), length(FieldValue));
                       FieldValue := Tools.CastDateForQry(StrToDate(sDate)) + sHour;
                     end;
                   end;
      end;
      FieldValue := AddQuotes(DSType, FieldName, FieldValue);
      Value      := Format('-%s=%s', [FieldName, FieldValue]);
      Result     := Result + Value;
    end;
  end;
  Result := Copy(Result, 2, Length(Result));
end;

function TSvcSyncBTPY2Execute.GetIndiceY2DataCache(DSType: T_WSDataService; TSLValue: TStringList): string;
var
  TableName : string;
  FieldsKey : string;
  FieldKey  : string;
  Index     : Integer;
begin
  TableName := GetInfoFromDSType(wsidTableName, DSType);
  FieldsKey := GetInfoFromDSType(wsidFieldsKey, DSType);
  Result := TableName + '_';
  while FieldsKey <> '' do
  begin
    FieldKey := Tools.ReadTokenSt_(FieldsKey, ';');
    Index := TSLValue.IndexOfName(FieldKey);
    if Index > -1 then
      Result := Result + Copy(TSLValue[Index], Pos('=', TSLValue[Index]) + 1, Length(TSLValue[Index]));
  end;
end;

function TSvcSyncBTPY2Execute.AddQuotes(DSType: T_WSDataService; FieldName, FieldValue: string): string;
begin
  if Pos('''', FieldValue) > 0 then
    FieldValue := StringReplace(FieldValue, '''', '''''', [rfReplaceAll]);
  case GetFieldTypeFromCache(DSType, FieldName) of
    ttfBoolean, ttfCombo, ttfMemo, ttfText, ttfDate : Result := Format('''%s''', [FieldValue]);
  else
    Result := FieldValue;
  end;
end;

function TSvcSyncBTPY2Execute.ValueAnalysis(LineValue, FieldName : string; FieldType : tTypeField) : string;
begin
  if Pos('_DATEMODIF', FieldName) > 0 then
    Result := DateTimeToStr(Now)
  else
  begin
    Result := copy(LineValue, Pos('=', LineValue) + 1, length(LineValue));
    case FieldType of
      ttfDate    : Result := Tools.SetStrDateTimeFromStrUTCDateTime(Result);
      ttfCombo   : Result := Trim(Result);
      ttfNumeric :
        begin
          Result := StringReplace(Result, ',', '.', [rfReplaceAll]);
          if Result = '' then
            Result := '0';
        end;
    end;
  end;
  if FieldType = ttfDate then
    Result := FormatDateTime('yyyymmdd hh:nn:ss', StrToDateTime(Result));
end;
  
function TSvcSyncBTPY2Execute.GetY2Data(DSType: T_WSDataService): boolean;
var
  Cpt             : integer;
  Index           : integer;
  ModifyQty       : integer;
  InsertQty       : integer;
  TableName       : string;
  GetDataState    : string;
  LogMsgTableName : string;
  TslCacheInsert  : TStringList;

  procedure SetCacheY2Data;
  var
    IndexValue  : string;
    RecordValue : string;
    Index       : integer;
  begin
    if TSlIndice.Count > 0 then
    begin
      IndexValue  := GetIndiceY2DataCache(DSType, TSlIndice) + '=';
      RecordValue := '';
      for Index := 0 to pred(TSlIndice.Count) do
      begin
        if Copy(TSlIndice[Index], 1, 7) <> WSCDS_IndiceField then
          RecordValue := RecordValue + ToolsTobToTsl_Separator + TSlIndice[Index];
      end;
      RecordValue := Copy(RecordValue, 2, length(RecordValue));
      TSlCacheGetY2Data.Add(IndexValue + RecordValue);
    end;
  end;

begin
  try
    TableName := GetInfoFromDSType(wsidTableName, DSType);
    ModifyQty := 0;
    InsertQty := 0;
    LogMsgTableName := Format('%s=%s', [DateTimeToStr(Now), TableName]);
    SetFilterFromDSType(DSType, TSlFilter);
    try
      GetDataState := TReadWSDataService.GetData(DSType, Y2Values.Server, Y2Values.DataBase, TSlValues, TSlCacheWSFields, LogValues, TSlFilter, '', GetTSlFromDSType(DSType, True));
      Result       := (GetDataState = WSCDS_GetDataOk);
      if Result then
      begin
        if TSlValues.Count > 0 then
        begin
          TslCacheInsert := TStringList.Create;
          try
            TslCacheInsert.Sorted := True;
            for Cpt := 0 to Pred(TSlValues.Count) do
            begin
              if Copy(TSlValues[Cpt], 1, 7) = WSCDS_IndiceField then
              begin
                Index := StrToInt(Copy(TSlValues[Cpt], Pos('=', TSlValues[Cpt]) + 1, Length(TSlValues[Cpt])));
                TSlIndice.Clear;
                ExtractIndice(Index, TSlValues, TSlIndice);
                { Met la valeur trouv�e en cache pour comparer si les valeurs ont chang�es lors du Send vers Y2 }
                SetCacheY2Data;         
                { Test si existe d�j� }
                AdoQryBTP.TSLResult.Clear;
                AdoQryBTP.FieldsList := GetInfoFromDSType(wsidFieldsList, DSType);
                AdoQryBTP.Request    := GetInfoFromDSType(wsidRequest, DSType);
                try
                  AdoQryBTP.SingleTableSelect;
                  if AdoQryBTP.RecordCount > 0 then
                  begin
                    Inc(ModifyQty);
                    Result := AddData(wsacUpdate, DSType, TSlIndice);
                  end else
                  begin
                    if DSType <> wsdsSocietyParameters then // Il ne faut pas cr�er les nouveaux PSoc de la compta
                    begin
                      if Tools.GetIndexOnSortedTsl(TslCacheInsert, AdoQryBTP.Request, True) = -1 then
                      begin
                        TslCacheInsert.Add(AdoQryBTP.Request);
                        Inc(InsertQty);
                        Result := AddData(wsacInsert, DSType, TSlIndice);
                      end;
                    end;
                  end;
                  AdoQryBTP.Reset;
                except
                  if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.GetY2Data. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryBTP.ServerName, AdoQryBTP.DBName, AdoQryBTP.Request]), ServiceName_BTPY2, LogValues, 0);
                  Result := False;
                  exit;
                end;
              end;
            end;
          finally
            TslCacheInsert.Free;
          end;
        end;
        if (ModifyQty + InsertQty) > 0 then
        begin
          if ModifyQty > 0 then
            LogMsgTableName := Format('%s (%s enregistrement(s) � modifier.', [LogMsgTableName, IntToStr(ModifyQty)]);
          if InsertQty > 0 then
            LogMsgTableName := Format('%s %s%s enregistrement(s) � cr�er.', [LogMsgTableName, Tools.iif(ModifyQty > 0, ' ', '('), IntToStr(InsertQty)]);
          LogMsgTableName := Format('%s)', [LogMsgTableName]);
          TServicesLog.WriteLog(ssbylLog, LogMsgTableName, ServiceName_BTPY2, LogValues, 3, False);
        end;
      end else
      begin
        { Si pas de log m�tier, �crit l'erreur dans le log windows }
        if LogValues.LogLevel > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s : %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, GetDataState]), ServiceName_BTPY2, LogValues, 0);
      end;
    except
      if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s : R�cup�ration des donn�es de %s (TSvcSyncBTPY2Execute.GetY2Data).', [WSCDS_DebugMsg, WSCDS_ErrorMsg, TableName]), ServiceName_BTPY2, LogValues, 0);
      Result := False;
    end;
  finally
    TSlFilter.Clear;
    TSlValues.Clear;
  end;
end;

function TSvcSyncBTPY2Execute.GetPayment(DSType: T_WSDataService) : boolean;
var
  GetDataState       : string;
  AccJrl             : string;
  AccPaymentMode     : string;
  AccDocType         : string;
  AccChkNumber       : string;
  AccLabel           : string;
  AccDocQualif       : string;
  AccAuxiliary       : string;
  DocValues          : string;
  BE0Type            : string;
  BE0Values          : string;
  BE0Stump           : string;
  AccNumber          : integer;
  AccLineNumber      : Integer;
  BE0Index           : integer;
  BE0Number          : integer;
  PaymentQty         : integer;
  DocumentQty        : integer;
  AccDeb             : double;
  AccDebDev          : double;
  AccCred            : double;
  AccCredDev         : double;
  AccAmount          : double;
  AccAmountDev       : double;
  DocAmount          : double;
  DocAmountDev       : double;
  AccDate            : TDateTime;
  AccDeadlineDate    : TDateTime;
  TslPayments        : TStringList;
  TslDocuments       : TStringList;
  TslDocumentsDet    : TStringList;
  TslInsertList      : TStringList;
  TslCacheThird      : TStringList;
  TslCacheFiscalYear : TStringList;

  procedure CopyIndiceToTsl(lTsl : TStringList);
  var
    Cptc : integer;
  begin
    for CptC := 0 to pred(TSlIndice.count) do
      lTsl.Add(TSlIndice[CptC]);
  end;

  procedure SetDocumentTsl(Third, Lettering : string);
  var
    CptD   : integer;
    IndexL : integer;
  begin
    for CptD := 0 to Pred(TslDocuments.Count) do
    begin
      if Copy(TslDocuments[CptD], 1, 7) = WSCDS_IndiceField then
      begin
        IndexL := StrToInt(Copy(TslDocuments[CptD], Pos('=', TslDocuments[CptD]) + 1, Length(TslDocuments[CptD])));
        TslDocumentsDet.Clear;
        ExtractIndice(IndexL, TslDocuments, TslDocumentsDet);
        if (GetSingleValueFromTSl('E_AUXILIAIRE', TslDocumentsDet) = Third) and (GetSingleValueFromTSl('E_LETTRAGE', TslDocumentsDet) = Lettering) then
          break;
      end;
    end;
  end;

  function GetThirdLabel(Auxiliary : string) : string;
  var
    lIndex  : integer;
    AdoQryT : AdoQry;
  begin
    if Auxiliary <> '' then
    begin
      lIndex := TslCacheThird.IndexOfName(Auxiliary);
      if lIndex = -1 then
      begin
        AdoQryT := AdoQry.Create;
        try
          AdoQryT.ServerName  := BTPValues.Server;
          AdoQryT.DBName      := BTPValues.DataBase;
          AdoQryT.Qry         := TADOQuery.create(nil);
          AdoQryT.FieldsList  := 'T_LIBELLE';
          AdoQryT.Request     := Format('SELECT %s FROM TIERS WHERE T_AUXILIAIRE = ''%s''', [AdoQryT.FieldsList, Auxiliary]);
          AdoQryT.LogValues   := LogValues;
          try
            AdoQryT.SingleTableSelect;
            if AdoQryT.TSLResult.count > 0 then
              TslCacheThird.Add(Format('%s=%s', [Auxiliary, AdoQryT.TSLResult[0]]));
            Result := AdoQryT.TSLResult[0];
          except
            if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.GetPayment/GetThirdLabel. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryT.ServerName, AdoQryT.DBName, AdoQryT.Request]), ServiceName_BTPY2, LogValues, 0);
            exit;
          end;
        finally
          AdoQryT.Qry.Free;
          AdoQryT.free;
        end;
      end else
        Result := copy(TslCacheThird[lIndex], pos('=', TslCacheThird[lIndex]) + 1, length(TslCacheThird[lIndex]));
    end else
      Result := '';
  end;

  function AddInsertToTsl : boolean;
  var
    IsPayment           : string;
    sIndex              : string;
    SqlInsert           : string;
    SqlUpdate           : string;
    Amount              : string;
    AmountDev           : string;
    ThirdLabel          : string;
    FiscalYearCode      : string;
    lAccDate            : string;
    FiscalYearStartDate : string;
    FiscalYearEndDate   : string;
    OrderNum            : integer;
    StartScan           : integer;
    Cpt                 : integer;
    FiscalYearIndex     : integer;
    AdoQryE             : AdoQry;
    AlreadyExist        : boolean;

    procedure SetFiscalYearValues(LineValue : string; EmptyValue : boolean=False);
    var
      FiscalYearValue : string;
    begin
      FiscalYearValue     := LineValue;
      FiscalYearCode      := Tools.iif(not EmptyValue, Tools.ReadTokenSt_(FiscalYearValue, ToolsTobToTsl_Separator), '');
      FiscalYearStartDate := Tools.iif(not EmptyValue, Tools.ReadTokenSt_(FiscalYearValue, ToolsTobToTsl_Separator), '');
      FiscalYearEndDate   := Tools.iif(not EmptyValue, Tools.ReadTokenSt_(FiscalYearValue, ToolsTobToTsl_Separator), '');
    end;

  begin
    Result     := True;
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.GetPayment/AddInsertToTsl.', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
    IsPayment  := Tools.iif(copy(AccDocType, 1, 1) = 'R', 'X', '-');
    sIndex     := AccAuxiliary + '-' + BE0Type + '-' + BE0Stump + '-' + IntToSTr(BE0Number) + '-' + IntToSTr(BE0Index);
    OrderNum   := 0;
    ThirdLabel := GetThirdLabel(AccAuxiliary);
    { Si le montant du rgt est inf�rieur � celui de la pi�ce, il faut prendre celui dui rgt, sinon, l'inverse }
    if AccAmount <= DocAmount then
    begin
      Amount    := FloatToStr(AccAmount);
      AmountDev := FloatToStr(AccAmountDev);
    end else
    begin
      Amount    := FloatToStr(DocAmount);
      AmountDev := FloatToStr(DocAmountDev);
    end;
    Amount    := StringReplace(Amount   , ',', '.', [rfReplaceAll]);
    AmountDev := StringReplace(AmountDev, ',', '.', [rfReplaceAll]);
    lAccDate  := Tools.CastDateTimeForQry(AccDate);
    { Calcul de GAC_NUMORDRE }
    StartScan := TslInsertList.IndexOfName(sIndex);
    if StartScan > -1 then
    begin
      for Cpt := StartScan to pred(TslInsertList.Count) do
      begin
        if copy(TslInsertList[Cpt], 1, pos('=', TslInsertList[Cpt])-1) = sIndex then
          inc(OrderNum);
      end;
    end;
    { TEST SI LA LIGNE EST DEJA IMPORTEE }
    { Recherche exercice fiscal dans le cache }
    FiscalYearIndex := -1;
    for Cpt := 0 to pred(TslCacheFiscalYear.count) do
    begin
      SetFiscalYearValues(TslCacheFiscalYear[Cpt]);
      if (AccDate >= StrToDateTime(FiscalYearStartDate)) and (AccDate <= StrToDateTime(FiscalYearEndDate)) then
      begin
        FiscalYearIndex := Cpt;
        break;
      end;
    end;
    AdoQryE := AdoQry.Create;
    try
      AdoQryE.ServerName := BTPValues.Server;
      AdoQryE.DBName     := BTPValues.DataBase;
      AdoQryE.Qry        := TADOQuery.create(nil);
      { Exercice non trouv� dans le cache, ajoute }
      if FiscalYearIndex = -1 then
      begin
        SetFiscalYearValues('', True);
        AdoQryE.FieldsList := 'EX_EXERCICE,EX_DATEDEBUT,EX_DATEFIN';
        AdoQryE.Request    := Format('SELECT %s'
                                   + ' FROM EXERCICE'
                                   + ' WHERE EX_DATEDEBUT <= ''%s'''
                                   + '   AND EX_DATEFIN   >= ''%s'''
                                  , [AdoQryE.FieldsList, lAccDate, lAccDate]);
        AdoQryE.LogValues   := LogValues;
        try
          AdoQryE.SingleTableSelect;
          if AdoQryE.RecordCount = 1 then
          begin
            TslCacheFiscalYear.Add(AdoQryE.TSLResult[0]);
            SetFiscalYearValues(AdoQryE.TSLResult[0]);
          end;
        except
          if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.GetPayment/AddInsertToTsl. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryE.ServerName, AdoQryE.DBName, AdoQryE.Request]), ServiceName_BTPY2, LogValues, 0);
          Result := False;
          exit;
        end;
        AdoQryE.Reset;
      end;
      { Test si l'enregistrement a d�j� �t� import� }
      AdoQryE.TSLResult.Clear;
      AdoQryE.RecordCount := 0;
      AdoQryE.FieldsList  := 'GAC_NUMECR';
      AdoQryE.Request     := Format('SELECT %s'
                                  + ' FROM ACOMPTES'
                                  + ' WHERE GAC_NATUREPIECEG = ''%s'''
                                  + '   AND GAC_SOUCHE       = ''%s'''
                                  + '   AND GAC_NUMERO       = %s'
                                  + '   AND GAC_INDICEG      = %s'
                                  + '   AND GAC_JALECR       = ''%s'''
                                  + '   AND GAC_NUMECR       = %s'
                                  + '   AND GAC_NUMLECR      = %s'
                                  + '   AND GAC_DATEECR     >= ''%s'''
                                  + '   AND GAC_DATEECR     <= ''%s'''
                                 , [AdoQryE.FieldsList
                                    , BE0Type
                                    , BE0Stump
                                    , IntToSTr(BE0Number)
                                    , IntToSTr(BE0Index)
                                    , AccJrl
                                    , IntToStr(AccNumber)
                                    , IntToStr(AccLineNumber)
                                    , Tools.CastDateTimeForQry(StrToDateTime(FiscalYearStartDate))
                                    , Tools.CastDateTimeForQry(StrToDateTime(FiscalYearEndDate))
                                   ]);
      AdoQryE.LogValues   := LogValues;
      try
        AdoQryE.SingleTableSelect;
        AlreadyExist := (AdoQryE.RecordCount = 1);
      except
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.GetPayment/AddInsertToTsl. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryE.ServerName, AdoQryE.DBName, AdoQryE.Request]), ServiceName_BTPY2, LogValues, 0);
          Result := False;
        exit;
      end;
    finally
      AdoQryE.Qry.Free;
      AdoQryE.free;
    end;
    { Jamais import�, pr�pare l'INSERT }
    if not AlreadyExist then
    begin
      SqlInsert := 'INSERT INTO ACOMPTES'
                 + ' (GAC_NATUREPIECEG, GAC_SOUCHE    , GAC_NUMERO     , GAC_INDICEG      , GAC_JALECR         , GAC_NUMECR   , GAC_MONTANT   , GAC_MONTANTDEV, GAC_MODEPAIE'
                 + ' ,GAC_ISREGLEMENT , GAC_CBINTERNET, GAC_CBLIBELLE  , GAC_DATEEXPIRE   , GAC_TYPECARTE      , GAC_CBNUMCTRL, GAC_CBNUMAUTOR, GAC_NUMCHEQUE , GAC_LIBELLE'
                 + ' ,GAC_QUALIFPIECE , GAC_AUXILIAIRE, GAC_MONTANTTRA , GAC_MONTANTTRADEV, GAC_PIECEPRECEDENTE, GAC_NUMORDRE , GAC_CPTADIFF  , GAC_DATEECR   , GAC_DATEECHEANCE'
                 + ' ,GAC_FOURNISSEUR , GAC_ORIGINE   , GAC_NUMLECR)'
                 + 'VALUES'
                 + ' ('
                 + ''''  + BE0Type  + ''''                                           // GAC_NATUREPIECEG
                 + ','   + '''' + BE0Stump + ''''                                    // GAC_SOUCHE
                 + ','   + IntToSTr(BE0Number)                                       // GAC_NUMERO
                 + ','   + IntToSTr(BE0Index)                                        // GAC_INDICEG
                 + ','   + ''''  + AccJrl + ''''                                     // GAC_JALECR
                 + ','   + IntToStr(AccNumber)                                       // GAC_NUMECR
                 + ','   + Amount                                                    // GAC_MONTANT
                 + ','   + AmountDev                                                 // GAC_MONTANTDEV
                 + ','   + '''' + AccPaymentMode + ''''                              // GAC_MODEPAIE
                 + ','   + '''' + IsPayment + ''''                                   // GAC_ISREGLEMENT
                 + ','   + ''''''                                                    // GAC_CBINTERNET
                 + ','   + '''' + ThirdLabel + ''''                                  // GAC_CBLIBELLE
                 + ','   + ''''''                                                    // GAC_DATEEXPIRE
                 + ','   + ''''''                                                    // GAC_TYPECARTE
                 + ','   + ''''''                                                    // GAC_CBNUMCTRL
                 + ','   + ''''''                                                    // GAC_CBNUMAUTOR
                 + ','   + '''' + AccChkNumber + ''''                                // GAC_NUMCHEQUE
                 + ','   + '''' + AccLabel + ''''                                    // GAC_LIBELLE
                 + ','   + '''' + AccDocQualif + ''''                                // GAC_QUALIFPIECE
                 + ','   + '''' + AccAuxiliary + ''''                                // GAC_AUXILIAIRE
                 + ','   + Amount                                                    // GAC_MONTANTTRA
                 + ','   + AmountDev                                                 // GAC_MONTANTTRADEV
                 + ','   + ''''''                                                    // GAC_PIECEPRECEDENTE
                 + ','   + IntToStr(OrderNum)                                        // GAC_NUMORDRE
                 + ','   + '''-'''                                                   // GAC_CPTADIFF
                 + ','   + '''' + lAccDate + ''''                                    // GAC_DATEECR
                 + ','   + '''' + Tools.CastDateTimeForQry(AccDeadlineDate) + ''''   // GAC_DATEECHEANCE
                 + ','   + ''''''                                                    // GAC_FOURNISSEUR
                 + ','   + '''Y2'''                                                  // GAC_ORIGINE
                 + ','   + IntToStr(AccLineNumber)                                   // GAC_NUMLECR
                 + ' )'
                 ;
      TslInsertList.Add(sIndex + '=' + SqlInsert);
      SqlUpdate := Format('UPDATE PIECE SET GP_ACOMPTEDEV = GP_ACOMPTEDEV + %s, GP_ACOMPTE = GP_ACOMPTE + %s WHERE GP_NATUREPIECEG = %s AND GP_SOUCHE = %s AND GP_NUMERO = %s AND GP_INDICEG = %s'
                       , [  AmountDev
                          , Amount
                          , '''' + BE0Type  + ''''
                          , '''' + BE0Stump  + ''''
                          , IntToSTr(BE0Number)
                          , IntToSTr(BE0Index)
                         ]);
      TslInsertList.Add(sIndex + '=' + SqlUpdate);
    end;
  end;

  function AddInsertToTslNL(DocEntity, DocFiscalYear, DocJrl, DocNumber : string) : boolean;
  begin


  end;

  procedure SeparatePaymentDocument;
  var
    Cpt     : integer;
    Index   : integer;
    DocType : string;
  begin
    { S�pare les r�glements des pi�ces }
    for Cpt := 0 to Pred(TSlValues.Count) do
    begin
      if Copy(TSlValues[Cpt], 1, 7) = WSCDS_IndiceField then
      begin
        Index := StrToInt(Copy(TSlValues[Cpt], Pos('=', TSlValues[Cpt]) + 1, Length(TSlValues[Cpt])));
        TSlIndice.Clear;
        ExtractIndice(Index, TSlValues, TSlIndice);
        DocType := Trim(GetSingleValueFromTSl('E_NATUREPIECE', TSlIndice));
        if pos(';' + DocType + ';', ';RC;RF;OC;OF;') > 0 then
        begin
          inc(PaymentQty);
          CopyIndiceToTsl(TslPayments)
        end else
        begin
          inc(DocumentQty);
          CopyIndiceToTsl(TslDocuments);
        end;
      end;
    end;
    TSlIndice.Clear;
  end;

  function Treatment : boolean;
  var
    Cpt           : integer;
    Index         : integer;
    DocNumber     : integer;
    AccValues     : string;
    logIndex      : string;
    logAuxiliary  : string;
    logDocType    : string;
    logDocNumber  : string;
    MSgErr        : string;
    DocEntity     : string;
    DocFiscalYear : string;
    DocJrl        : string;
    AdoQryL       : AdoQry;
  begin

//exit;

    Result      := True;
    PaymentQty  := 0;
    DocumentQty := 0;
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.GetPayment : S�pare les r�glements des pi�ces.', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
    SeparatePaymentDocument;
    { Pour chaque r�glement, recherche la pi�ce associ�e (auxiliaire et code lettrage identique) pour r�cup�rer E_COUVERTURE et E_COUVERTUREDEV
      et les champs permettant de retrouver l'enregistrement correspondant dans BTPECRITURE }
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.GetPayment : Recherche des pi�ces associ�es aux r�glements.', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
    for Cpt := 0 to Pred(TslPayments.Count) do
    begin
      if Copy(TslPayments[Cpt], 1, 7) = WSCDS_IndiceField then
      begin
        Index := StrToInt(Copy(TslPayments[Cpt], Pos('=', TslPayments[Cpt]) + 1, Length(TslPayments[Cpt])));
        TSlIndice.Clear;
        ExtractIndice(Index, TslPayments, TSlIndice);
        SetDocumentTsl(GetSingleValueFromTSl('E_AUXILIAIRE', TSlIndice), GetSingleValueFromTSl('E_LETTRAGE', TSlIndice));
        if TslDocumentsDet.count > 0 then
        begin
          AccValues := GetMultipleValueFromTSl('E_JOURNAL'
                                               + ';E_NUMEROPIECE'
                                               + ';E_MODEPAIE'
                                               + ';E_NATUREPIECE'
                                               + ';E_NUMTRAITECHQ'
                                               + ';E_LIBELLE'
                                               + ';E_QUALIFPIECE'
                                               + ';E_AUXILIAIRE'
                                               + ';E_DATECOMPTABLE'
                                               + ';E_DATEECHEANCE'
                                               + ';E_DEBIT'
                                               + ';E_DEBITDEV'
                                               + ';E_CREDIT'
                                               + ';E_CREDITDEV'
                                               + ';E_NUMLIGNE'
                                               , TSlIndice);
          AccJrl          := trim(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                          // E_JOURNAL
          AccNumber       := StrToInt(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                      // E_NUMEROPIECE
          AccPaymentMode  := trim(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                          // E_MODEPAIE
          AccDocType      := trim(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                          // E_NATUREPIECE
          AccChkNumber    := Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator);                                                // E_NUMTRAITECHQ
          AccLabel        := Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator);                                                // E_LIBELLE
          AccDocQualif    := trim(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                          // E_QUALIFPIECE
          AccAuxiliary    := Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator);                                                // E_AUXILIAIRE
          AccDate         := StrToDate(Tools.SetStrUTCDateTimeToDateTime(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator)));  // E_DATECOMPTABLE
          AccDeadlineDate := StrToDate(Tools.SetStrUTCDateTimeToDateTime(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator)));  // E_DATEECHEANCE
          AccDeb          := StrToFloat(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                    // E_DEBIT
          AccDebDev       := StrToFloat(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                    // E_DEBITDEV
          AccCred         := StrToFloat(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                    // E_CREDIT
          AccCredDev      := StrToFloat(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                    // E_CREDITDEV
          AccLineNumber   := StrToInt(Tools.ReadTokenSt_(AccValues, ToolsTobToTsl_Separator));                                      // E_NUMLIGNE
          AccAmount       := AccDeb    + AccCred;                                                                                   // Montant de l'acompte
          AccAmountDev    := AccDebDev + AccCredDev;                                                                                // Montant de l'acompte en devise
          DocValues       := GetMultipleValueFromTSl('E_ENTITY;E_EXERCICE;E_JOURNAL;E_NUMEROPIECE;E_COUVERTURE;E_COUVERTUREDEV', TslDocumentsDet);
          DocEntity     := Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator);
          DocFiscalYear := trim(Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator));
          DocJrl        := Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator);
          DocNumber     := StrToInt(Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator));
          DocAmount     := StrToFloat(Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator));
          DocAmountDev  := StrToFloat(Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator)); 
          AdoQryL       := AdoQry.Create;
          try
            AdoQryL.ServerName := BTPValues.Server;
            AdoQryL.DBName     := BTPValues.DataBase;
            AdoQryL.Qry        := TADOQuery.create(nil);
            AdoQryL.FieldsList := 'BE0_NATUREPIECEG,BE0_SOUCHE,BE0_NUMERO,BE0_INDICEG';
            AdoQryL.Request    := Format('SELECT %s FROM BTPECRITURE WHERE BE0_ENTITY = %s AND BE0_JOURNAL = ''%s'' AND BE0_EXERCICE = ''%s'' AND BE0_REFERENCEY2 = %s AND BE0_ORIGINE = ''BTP'''
                                      , [AdoQryL.FieldsList
                                         , DocEntity
                                         , DocJrl
                                         , DocFiscalYear
                                         , IntToStr(DocNumber)]);
            AdoQryL.LogValues   := LogValues;
            try
              MSgErr := AdoQryL.SingleTableSelect;
              if AdoQryL.RecordCount > 0 then
              begin
                BE0Values := AdoQryL.TSLResult[0];
                BE0Type   := Tools.ReadTokenSt_(BE0Values, ToolsTobToTsl_Separator);
                BE0Stump  := Tools.ReadTokenSt_(BE0Values, ToolsTobToTsl_Separator);
                BE0Number := StrToint(Tools.ReadTokenSt_(BE0Values, ToolsTobToTsl_Separator));
                BE0Index  := StrToint(Tools.ReadTokenSt_(BE0Values, ToolsTobToTsl_Separator));
                Result    := AddInsertToTsl;
              end;
              if not Result then
                exit;
            except
              if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.GetPayment/Treatment. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryL.ServerName, AdoQryL.DBName, AdoQryL.Request]), ServiceName_BTPY2, LogValues, 0);
              Result := False;
              exit;
            end;
          finally
            AdoQryL.Qry.Free;
            AdoQryL.free;
          end;
        end;
      end;
    end;
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.GetPayment/Treatment #01 : Cr�ation des r�glements et mise � jour des pi�ces.', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
    if TslInsertList.count > 0 then
    begin
      TServicesLog.WriteLog(ssbylLog, Format('%s r�glement(s) trouv�(s).', [IntToStr(PaymentQty)]), ServiceName_BTPY2, LogValues, 3);
      for Cpt := 0 to pred(TslInsertList.count) do
      begin
        if pos(' ACOMPTES ', TslInsertList[Cpt]) > 0 then
        begin
          logIndex    := copy(TslInsertList[Cpt], 1, pos('=', TslInsertList[Cpt]) - 1);
          logAuxiliary := Tools.ReadTokenSt_(logIndex, '-');
          logDocType   := Tools.ReadTokenSt_(logIndex, '-');
          Tools.ReadTokenSt_(logIndex, '-');
          logDocNumber := Tools.ReadTokenSt_(logIndex, '-');
          TServicesLog.WriteLog(ssbylLog, Format('Auxiliaire %s - Pi�ce %s n� %s', [logAuxiliary, logDocType, logDocNumber]), ServiceName_BTPY2, LogValues, 4);
        end;
        AdoQryBTP.Request := copy(TslInsertList[Cpt], pos('=', TslInsertList[Cpt]) + 1, length(TslInsertList[Cpt]));
        try
          AdoQryBTP.InsertUpdate;
          Result := (AdoQryBTP.RecordCount = 1);
        except
          if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.GetPayment/Treatment #02. Erreur sur ex�cution de %s / %s / %s'
                                                                            , [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryBTP.ServerName, AdoQryBTP.DBName, AdoQryBTP.Request]), ServiceName_BTPY2, LogValues, 0);
          Result := False;
          exit;
        end;
        AdoQryBTP.Reset;
        if not Result then
          Break;
      end;
    end;
  end;

  function TreatmentNL : boolean;
  var
    Cpt           : integer;
    CptIndice     : integer;
    DocNumber     : integer;
    DocLineNumber : integer;
    DocEchNumber  : integer;
    Index         : integer;
    AdoQryL       : AdoQry;
    MsgErr        : string;
    DocValues     : string;
    DocEntity     : string;
    DocFiscalYear : string;
    DocJrl        : string;
    DocDate       : string;
    DocIndexPrev  : string;
    DocIndexCurr  : string;
    DocQualif     : string;
    Value         : string;
    FieldsList    : string;
    TslCacheRgtNL : TSTringlist;
    TslCurrRgtNL  : TSTringlist;

    function IsAlreadyImported(DocIndex : string) : boolean;
    begin
      if Tools.GetIndexOnSortedTsl(TslCacheRgtNL, DocIndex) = -1 then
      begin
        TslCacheRgtNL.Add(Format('%s=X', [DocIndex]));
        AdoQryL := AdoQry.Create;
        try
          AdoQryL.ServerName := BTPValues.Server;
          AdoQryL.DBName     := BTPValues.DataBase;
          AdoQryL.Qry        := TADOQuery.create(nil);
          AdoQryL.FieldsList := 'BE0_NUMEROPIECE';
          AdoQryL.Request    := Format('SELECT %s FROM BTPECRITURE WHERE BE0_ENTITY = %s AND BE0_JOURNAL = ''%s'' AND BE0_EXERCICE = ''%s'' AND BE0_REFERENCEY2 = %s AND BE0_ORIGINE = ''Y2'''
                                    , [AdoQryL.FieldsList
                                       , DocEntity
                                       , DocJrl
                                       , DocFiscalYear
                                       , IntToStr(DocNumber)]);
          AdoQryL.LogValues := LogValues;
          try
            MSgErr := AdoQryL.SingleTableSelect;
            Result := (AdoQryL.RecordCount > 0);
            if not Result then
              AddInsertToTslNL(DocEntity, DocFiscalYear, DocJrl, IntToStr(DocNumber))
            else
            begin
              if LogValues.DebugEvents = 2 then TServicesLog.WriteLog(ssbylLog, Format('Le r�glement "journal %s - n� pi�ce %s - du %s"  a d�j� �t� r�cup�r� sous le num�ro %s.'
                                                                                    , [  DocJrl
                                                                                       , IntToStr(DocNumber)
                                                                                       , DocDate
                                                                                       , AdoQryL.TSLResult[0]
                                                                                      ])
                                                                      , ServiceName_BTPY2, LogValues, 3);

            end;
          except
            if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.GetPayment/Treatment. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryL.ServerName, AdoQryL.DBName, AdoQryL.Request]), ServiceName_BTPY2, LogValues, 0);
            Result := False;
            exit;
          end;
        finally
          AdoQryL.Qry.Free;
          AdoQryL.free;
        end;
      end else
        Result := True;
    end;

    function GetFieldsList : string;
    var
      CptFL : integer;
      Value : string;
    begin
      if Fieldslist = '' then
      begin
        for CptFL := 0 to pred(TslCurrRgtNL.count) do
        begin
          Value  := TslPayments[Cpt];
          if Copy(Value, 1, 7) <> WSCDS_IndiceField then
            FieldsList := Format('%s, %s', [FieldsList, copy(Value, 1, pos('=', Value) - 1)]);
        end;
        FieldsList := copy(FieldsList, 2, length(FieldsList));
      end;
      Result := FieldsList;
    end;

    function AccountInsert : boolean;
    var
      CptI      : integer;
      SqlInsert : string;
      Value     : string;
      TslInsert : TStringList;
    begin
      Result    := True;
      TslInsert := TStringList.Create;
      try
        SqlInsert := ''; //Format('INSERT INTO ECRITURE (%s) VALUES', []);
        for CptI := 0 to pred(TslCurrRgtNL.count) do
        begin
          Value := TslPayments[Cpt];
          if Copy(Value, 1, 7) <> WSCDS_IndiceField then
            SqlInsert := Format('%s, ''%s''', [SqlInsert, copy(Value, pos('=', Value) + 1 , length(Value))]);
        end;
        SqlInsert := copy(SqlInsert, 2, length(SqlInsert));
        SqlInsert := Format('INSERT INTO ECRITURE (%s) VALUES (%s)', [GetFieldsList, SqlInsert]);
      finally
        FreeAndNil(TslInsert);
      end;
    end;

  begin
    Result     := True;
    FieldsList := '';
    TslInsertList.Clear;
    TslCacheRgtNL := TStringList.Create;
    try
      TslCacheRgtNL.Sorted := True;
      TslCurrRgtNL := TStringList.Create;
      try
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.GetPayment/Tr : S�pare les r�glements NL des pi�ces.', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
        SeparatePaymentDocument;
        TSlIndice.Clear;
//        TslPayments.Sort;
        DocIndexPrev := 'None';
        DocIndexCurr := '';
        for Cpt := 0 to Pred(TslPayments.Count) do
        begin
          if Copy(TslPayments[Cpt], 1, 7) = WSCDS_IndiceField then
          begin
            Index := StrToInt(Copy(TslPayments[Cpt], Pos('=', TslPayments[Cpt]) + 1, Length(TslPayments[Cpt])));
            TSlIndice.Clear;
            ExtractIndice(Index, TslPayments, TSlIndice);
            DocValues     := GetMultipleValueFromTSl('E_ENTITY;E_EXERCICE;E_JOURNAL;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE;E_QUALIFPIECE;E_DATECOMPTABLE', TSlIndice);
            DocEntity     := Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator);
            DocFiscalYear := trim(Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator));
            DocJrl        := Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator);
            DocNumber     := StrToInt(Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator));
            DocLineNumber := StrToInt(Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator));
            DocEchNumber  := StrToInt(Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator));
            DocQualif     := Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator);
            DocDate       := Tools.ReadTokenSt_(DocValues, ToolsTobToTsl_Separator);
            DocIndexCurr  := DocEntity + DocFiscalYear + DocJrl + IntToStr(DocNumber);
            if not IsAlreadyImported(DocIndexCurr) then
            begin
              if DocIndexPrev = 'None' then
                DocIndexPrev := DocIndexCurr;
              { Nv r�glement }
              if DocIndexCurr <> DocIndexPrev then
              begin
                AccountInsert;
                TslCurrRgtNL.Clear;
                DocIndexPrev := DocIndexCurr;
              end else
              begin
                for CptIndice := 0 to pred(TSlIndice.count) do
                  TslCurrRgtNL.Add(TSlIndice[CptIndice]);
              end;
            end;
          end;
        end;
      finally
        FreeAndNil(TslCurrRgtNL);
      end;
    finally
      FreeAndNil(TslCacheRgtNL);
    end;
  end;

begin
  Result := True;
  if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.GetPayment', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
  if Tools.GetParamSocSecur_('SO_BTGETREGLCPTA', '-' {$IFDEF APPSRV}, BTPValues.Server, BTPValues.DataBase{$ENDIF APPSRV}) = 'X' then
  begin
    TServicesLog.WriteLog(ssbylLog, Format('Recherche des r�glements %s.', [Tools.iif(DSType = wsdsAccForPayment, 'lettr�s', 'non lettr�s')]), ServiceName_BTPY2, LogValues, 3);                   
    TSlFilter.Clear;
    SetFilterFromDSType(DSType, TSlFilter);
    try
      GetDataState := TReadWSDataService.GetData(DSType, Y2Values.Server, Y2Values.DataBase, TSlValues, TSlCacheWSFields, LogValues, TSlFilter);
      Result       := (GetDataState = WSCDS_GetDataOk);
      if Result then
      begin
        if TSlValues.Count > 0 then
        begin
          TslPayments := TStringList.Create;
          try
            TslDocuments := TStringList.Create;
            try
              TslDocumentsDet := TStringList.Create;
              try
                TslInsertList := TStringList.Create;
                try
                  TslCacheThird := TStringList.Create;
                  try
                    TslCacheFiscalYear := TStringList.Create;
                    try
                      if DSType = wsdsAccForPayment then
                        Result := Treatment
                      else
                        Result := TreatmentNL;
                    finally
                      FreeAndNil(TslCacheFiscalYear);
                    end;
                  finally
                    FreeAndNil(TslCacheThird);
                  end;
                finally
                  FreeAndNil(TslInsertList);
                end;
              finally
                FreeAndNil(TslDocumentsDet);
              end;
            finally
              FreeAndNil(TslDocuments);
            end;
          finally
            FreeAndNil(TslPayments);
          end;
        end;
      end;
    except
      TServicesLog.WriteLog(ssbylLog, Format('%s : Recherche des r�glements (TSvcSyncBTPY2Execute.GetPayment). Erreur : %s', [WSCDS_ErrorMsg, GetDataState]), ServiceName_BTPY2, LogValues, 0);
    end;
  end else
    Result := True;
end;

procedure TSvcSyncBTPY2Execute.CreateObjects;
begin
  TSlConnectionValues            := TstringList.Create;
  TSlValues                      := TstringList.Create;
  TSlFilter                      := TStringList.Create;
  TSlIndice                      := TStringList.Create;
  TSlCacheThirdBTP               := TStringList.Create;
  TSlCacheSectionBTP             := TStringList.Create;
  TSlCacheAcountBTP              := TStringList.Create;
  TSlCachePaymentBTP             := TStringList.Create;
  TSlCacheCorrespBTP             := TStringList.Create;
  TSlCacheCurrencyBTP            := TStringList.Create;
  TSlCacheCountryBTP             := TStringList.Create;
  TSlCacheRecoveryBTP            := TStringList.Create;
  TSlCacheCommonBTP              := TStringList.Create;
  TSlCacheChoixCodBTP            := TStringList.Create;
  TSlCacheChoixExtBTP            := TStringList.Create;
  TSlCacheJournalBTP             := TStringList.Create;
  TSlCacheBankIdBTP              := TStringList.Create;
  TSlCacheChangeRateBTP          := TStringList.Create;
  TSlCacheFiscalYearBTP          := TStringList.Create;
  TSlcacheSocietyParamBTP        := TStringList.Create;
  TSlcacheEstablishmentBTP       := TStringList.Create;
  TSlcachePaymentModeBTP         := TStringList.Create;
  TSlcacheZipCodeBTP             := TStringList.Create;
  TSlcacheContactBTP             := TStringList.Create;
  TSlCacheAccForPaymentBTP       := TStringList.Create;
  TSLCacheDocPaymentBTP          := TStringList.Create;
  TSLCacheTaxRateBTP             := TStringList.Create;
  TSlCacheThirdY2                := TStringList.Create;
  TSlCacheSectionY2              := TStringList.Create;
  TSlCacheAcountY2               := TStringList.Create;
  TSlCachePaymentY2              := TStringList.Create;
  TSlCacheCorrespY2              := TStringList.Create;
  TSlCacheCurrencyY2             := TStringList.Create;
  TSlCacheCountryY2              := TStringList.Create;
  TSlCacheRecoveryY2             := TStringList.Create;
  TSlCacheCommonY2               := TStringList.Create;
  TSlCacheChoixCodY2             := TStringList.Create;
  TSlCacheChoixExtY2             := TStringList.Create;
  TSlCacheJournalY2              := TStringList.Create;
  TSlCacheBankIdY2               := TStringList.Create;
  TSlCacheChangeRateY2           := TStringList.Create;
  TSlCacheFiscalYearY2           := TStringList.Create;
  TSlcacheSocietyParamY2         := TStringList.Create;
  TSlcacheEstablishmentY2        := TStringList.Create;
  TSlcachePaymentModeY2          := TStringList.Create;
  TSlcacheZipCodeY2              := TStringList.Create;
  TSlcacheContactY2              := TStringList.Create;
  TSlCacheAccForPaymentY2        := TStringList.Create;
  TSLCacheTaxRateY2              := TStringList.Create;
  TSlCacheWSFields               := TStringList.Create;
  TSlCacheSendAccParam           := TStringList.Create;
  TSlCacheGetY2Data              := TStringList.Create;
  TSLCacheUpdateFrequencySetting := TStringList.Create;
  TSLUpdateInsertData            := TStringList.Create;
  TSLTRAFileQty                  := TStringList.Create;
  AdoQryBTP                      := AdoQry.Create;
  AdoQryY2                       := AdoQry.Create;
end;

procedure TSvcSyncBTPY2Execute.FreeObjects;
begin
  FreeAndNil(TSlConnectionValues);
  FreeAndNil(TSlValues);
  FreeAndNil(TSlFilter);
  FreeAndNil(TSlIndice);
  FreeAndNil(TSlCacheThirdBTP);
  FreeAndNil(TSlCacheSectionBTP);
  FreeAndNil(TSlCacheAcountBTP);
  FreeAndNil(TSlCachePaymentBTP);
  FreeAndNil(TSlCacheCorrespBTP);
  FreeAndNil(TSlCacheCurrencyBTP);
  FreeAndNil(TSlCacheCountryBTP);
  FreeAndNil(TSlCacheRecoveryBTP);
  FreeAndNil(TSlCacheCommonBTP);
  FreeAndNil(TSlCacheChoixCodBTP);
  FreeAndNil(TSlCacheChoixExtBTP);
  FreeAndNil(TSlCacheJournalBTP);
  FreeAndNil(TSlCacheBankIdBTP);
  FreeAndNil(TSlCacheChangeRateBTP);
  FreeAndNil(TSlCacheFiscalYearBTP);
  FreeAndNil(TSlcacheSocietyParamBTP);
  FreeAndNil(TSlcacheEstablishmentBTP);
  FreeAndNil(TSlcachePaymentModeBTP);
  FreeAndNil(TSlcacheZipCodeBTP);
  FreeAndNil(TSlcacheContactBTP);
  FreeAndNil(TSlCacheAccForPaymentBTP);
  FreeAndNil(TSLCacheDocPaymentBTP);
  FreeAndNil(TSLCacheTaxRateBTP);
  FreeAndNil(TSlCacheThirdY2);
  FreeAndNil(TSlCacheSectionY2);
  FreeAndNil(TSlCacheAcountY2);
  FreeAndNil(TSlCachePaymentY2);
  FreeAndNil(TSlCacheCorrespY2);
  FreeAndNil(TSlCacheCurrencyY2);
  FreeAndNil(TSlCacheCountryY2);
  FreeAndNil(TSlCacheRecoveryY2);
  FreeAndNil(TSlCacheCommonY2);
  FreeAndNil(TSlCacheChoixCodY2);
  FreeAndNil(TSlCacheChoixExtY2);
  FreeAndNil(TSlCacheJournalY2);
  FreeAndNil(TSlCacheBankIdY2);
  FreeAndNil(TSlCacheChangeRateY2);
  FreeAndNil(TSlCacheFiscalYearY2);
  FreeAndNil(TSlcacheSocietyParamY2);
  FreeAndNil(TSlcacheEstablishmentY2);
  FreeAndNil(TSlcachePaymentModeY2);
  FreeAndNil(TSlcacheZipCodeY2);
  FreeAndNil(TSlcacheContactY2);
  FreeAndNil(TSlCacheAccForPaymentY2);
  FreeAndNil(TSLCacheTaxRateY2);
  FreeAndNil(TSlCacheWSFields);
  FreeAndNil(TSlCacheSendAccParam);
  FreeAndNil(TSlCacheGetY2Data);
  FreeAndNil(TSLCacheUpdateFrequencySetting);
  FreeAndNil(TSLUpdateInsertData);
  FreeAndNil(TSLTRAFileQty);
  AdoQryBTP.Free;
  AdoQryY2.Free;
end;

procedure TSvcSyncBTPY2Execute.ClearValuesConnection;
begin
  TSlCacheSendAccParam.Clear;
  TSlCacheGetY2Data.Clear;
  BTPValues.ConnectionName := '';
  BTPValues.UserAdmin      := '';
  BTPValues.Server         := '';
  BTPValues.DataBase       := '';
  BTPValues.LastSynchro    := '';
  Y2Values.ConnectionName  := '';
  Y2Values.Server          := '';
  Y2Values.DataBase        := '';
end;

procedure TSvcSyncBTPY2Execute.SetLastSyncIniFile;
var
  SettingFile  : TInifile;
  LastSyncDate : string;
  Values       : string;
  Folder       : string;
  UserAdmin    : string;
  BTPServer    : string;
  BTPDataBase  : string;
  LastSynchro  : string;
  Y2Server     : string;
  Y2DataBase   : string;
  Cpt          : integer;
begin
  if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - Start TSvcSyncBTPY2Execute.SetLastSyncIniFile', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
  SettingFile := TIniFile.Create(IniFilePath);
  try
    LastSyncDate := DateTimeToStr(Now);
    try
      SettingFile.WriteString(BTPValues.ConnectionName, IniFileLastSynchro, LastSyncDate);
      SettingFile.UpdateFile;
    except
      if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - WriteString TSvcSyncBTPY2Execute.SetLastSyncIniFile', [WSCDS_DebugMsg, WSCDS_ErrorMsg]), ServiceName_BTPY2, LogValues, 0);
    end;
  finally
    { Il faut recalculer la date synchro }
    for Cpt := 0 to pred(TSlConnectionValues.Count) do
    begin
      Values := TSlConnectionValues[Cpt];
      if copy(Values, 1, 6) = 'FOLDER' then
      begin
        Folder      := copy(Values, 1, pos('=', Values)-1);
        Values      := copy(Values, pos('=', Values)+1, length(Values));
        UserAdmin   := Tools.ReadTokenSt_(Values, ';');
        BTPServer   := Tools.ReadTokenSt_(Values, ';');
        BTPDataBase := Tools.ReadTokenSt_(Values, ';');
        LastSynchro := LastSyncDate;
        Tools.ReadTokenSt_(Values, ';');
        Y2Server    := Tools.ReadTokenSt_(Values, ';');
        Y2DataBase  := Tools.ReadTokenSt_(Values, ';');
        TSlConnectionValues[Cpt] := Format('%s=%s;%s;%s;%s;%s;%s', [Folder, UserAdmin, BTPServer, BTPDataBase, LastSynchro, Y2Server, Y2DataBase]);
      end;
    end;
    SettingFile.Free;
  end;
end;

function TSvcSyncBTPY2Execute.ReadSettings : boolean;
var
  SettingFile  : TInifile;
  Section      : string;
  SectionExist : boolean;
  Cpt          : integer;

  procedure AddFrequencySetting(TableName: string);
  begin
    TSLCacheUpdateFrequencySetting.Add(Format('%s=%s', [TableName, SettingFile.ReadString(Section, TableName, 'EveryTime')]));
  end;

begin
  Result := True;
  ClearValuesConnection;
  SettingFile := TIniFile.Create(IniFilePath);
  try
    Section := 'GLOBALSETTINGS';
    SecondTimeout                  := SettingFile.ReadInteger(Section, 'SecondTimeout', 0);
    LogValues.LogLevel             := SettingFile.ReadInteger(Section, 'LogLevel', 0);
    LogValues.LogMoMaxSize         := SettingFile.ReadInteger(Section, 'LogMoMaxSize', 0);
    LogValues.DebugEvents          := SettingFile.ReadInteger(Section, 'DebugEvents', 0);
    LogValues.OneLogPerDay         := (SettingFile.ReadInteger(Section, 'OneLogPerDay', 0) = 1);
    LogValues.LogMaxQty            := 0;
    LogValues.LogPath              := LogFilePath;
    LogValues.Download             := (SettingFile.ReadInteger(Section, 'Download', 0) = 1);
    LogValues.Upload               := (SettingFile.ReadInteger(Section, 'Upload', 0) = 1);
    LogValues.DebugFilesDirectory  := SettingFile.ReadString(Section, 'DebugFilesDirectory'    , '');
    LogValues.ExecutionPeriodDays  := SettingFile.ReadString(Section, 'ExecutionPeriodDays'    , '');
    LogValues.ExecutionPeriodStart := SettingFile.ReadString(Section, 'ExecutionPeriodStart'    , '');
    LogValues.ExecutionPeriodEnd   := SettingFile.ReadString(Section, 'ExecutionPeriodEnd'    , '');
    Section := 'UPDATEFREQUENCY';
    AddFrequencySetting('CHANCELL');
    AddFrequencySetting('CHOIXCOD');
    AddFrequencySetting('CHOIXEXT');
    AddFrequencySetting('CODEPOST');
    AddFrequencySetting('COMMUN');
    AddFrequencySetting('CORRESP');
    AddFrequencySetting('DEVISE');
    AddFrequencySetting('ETABLISS');
    AddFrequencySetting('EXERCICE');
    AddFrequencySetting('GENERAUX');
    AddFrequencySetting('JOURNAL');
    AddFrequencySetting('MODEPAIE');
    AddFrequencySetting('MODEREGL');
    AddFrequencySetting('PARAMSOC');
    AddFrequencySetting('PAYS');
    AddFrequencySetting('RELANCE');
    AddFrequencySetting('TXCPTTVA');
    AddFrequencySetting('ECRITURE');
    SectionExist := True;
    Cpt          := 0;
    while SectionExist do
    begin
      Inc(Cpt);
      Section      := 'FOLDER' + IntToStr(Cpt);
      SectionExist := (SettingFile.ReadString(Section, 'BTPUser', '') <> '');
      if SectionExist then
      begin
        BTPValues.UserAdmin   := SettingFile.ReadString(Section, IniFileBTPUser    , '');
        BTPValues.Server      := SettingFile.ReadString(Section, IniFileBTPServer  , '');
        BTPValues.DataBase    := SettingFile.ReadString(Section, IniFileBTPFolder, '');
        BTPValues.LastSynchro := SettingFile.ReadString(Section, IniFileLastSynchro, '');
        Y2Values.Server       := SettingFile.ReadString(Section, IniFileY2Server   , '');
        Y2Values.DataBase     := SettingFile.ReadString(Section, IniFileY2Folder , '');
        if BTPValues.UserAdmin <> '' then
        begin
          AdoQryBTP.ServerName  := Tools.iif(AdoQryBTP.ServerName = '', BTPValues.Server, AdoQryBTP.ServerName);
          AdoQryBTP.DBName      := Tools.iif(AdoQryBTP.DBName = '', BTPValues.DataBase, AdoQryBTP.DBName);
          AdoQryBTP.LogValues   := LogValues;
          AdoQryBTP.Qry         := TADOQuery.Create(nil);
          AdoQryY2.ServerName   := Tools.iif(AdoQryY2.ServerName = '', Y2Values.Server, AdoQryY2.ServerName);
          AdoQryY2.DBName       := Tools.iif(AdoQryY2.DBName = '', Y2Values.DataBase, AdoQryY2.DBName);
          AdoQryY2.LogValues    := LogValues;
          AdoQryY2.Qry          := TADOQuery.Create(nil);
          TSlConnectionValues.Add(Format('%s=%s;%s;%s;%s;%s;%s', [Section, BTPValues.UserAdmin, BTPValues.Server, BTPValues.DataBase, BTPValues.LastSynchro, Y2Values.Server, Y2Values.DataBase]));
        end;
      end;
    end;
  finally
    SettingFile.Free;
  end;
  ClearValuesConnection;
end;

function TSvcSyncBTPY2Execute.GetData: boolean;

  function GetPrefixSECTION(IsBTP: boolean): string;
  begin
    Result := Tools.iif(IsBTP, 'S', 'CSP');
  end;

  function GetTotalQty: integer;
  var
    CptQty : integer;
    Value  : string;
  begin
    Result := 0;
    for CptQty := pred(TSlValues.Count) downto 0 do
    begin
      Value := TSlValues[CptQty];
      if Copy(Value, 1, 1) = '#' then
      begin
        Result := StrToInt(Copy(Value, Pos('=', Value) + 1, Length(Value))) + 1;
        Break;
      end;
    end;
  end;

  function GetFromDSType(DSType: T_WSDataService): boolean;
  begin
    if CanExportImportTableFromFrenquency(DSType) then
  	begin
      if (DSType = wsdsAccForPayment) or (DSType = wsdsAccForPaymentOther) then
      begin
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.GetPayment%s/GetFromDSType ', [WSCDS_DebugMsg, Tools.iif(DSType = wsdsAccForPayment, '(lettr�s)', '(non lettr�s)')]), ServiceName_BTPY2, LogValues, 0);
        Result := GetPayment(DSType);
      end else
      begin
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.GetData/GetFromDSType - %s', [WSCDS_DebugMsg, GetTableNameFromDSType(DsType)]), ServiceName_BTPY2, LogValues, 0);
        Result := GetY2Data(DSType);
      end;
    end else
      Result := True;
  end;

  function InsertOrUpdateData: Boolean;
  var
    Cpt    : integer;
    UpdQty : integer;
    InsQty : integer;
  begin
    TServicesLog.WriteLog(ssbylLog, 'Mise � jour des donn�es cr��es ou modifi�es depuis le ' + BTPValues.LastSynchro, ServiceName_BTPY2, LogValues, 3);
    Result := True;
    UpdQty := 0;
    InsQty := 0;
    AdoQryBTP.TSLResult.Clear;
    for Cpt := 0 to pred(TSLUpdateInsertData.Count) do
    begin
      AdoQryBTP.Request := TSLUpdateInsertData[Cpt];
      if Copy(AdoQryBTP.Request, 1, 6) = 'UPDATE' then
        Inc(UpdQty)
      else
        Inc(InsQty);
      try
        AdoQryBTP.InsertUpdate;
        Result := (AdoQryBTP.RecordCount = 1);
        AdoQryBTP.Reset;
        if not Result then
          Break;
      except
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.GetData/InsertOrUpdateData. Erreur sur ex�cution de %s / %s / %s', [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryBTP.ServerName, AdoQryBTP.DBName, AdoQryBTP.Request]), ServiceName_BTPY2, LogValues, 0);
        Result := False;
        exit;
      end;
    end;
    if Result then
    begin
      if UpdQty > 0 then
        TServicesLog.WriteLog(ssbylLog, Format('Modification de %s enregistrement(s)', [IntToStr(UpdQty)]), ServiceName_BTPY2, LogValues, 4);
      if InsQty > 0 then
        TServicesLog.WriteLog(ssbylLog, Format('Cr�ation de %s enregistrement(s)'    , [IntToStr(InsQty)]), ServiceName_BTPY2, LogValues, 4);
    end;
  end;

begin
  TServicesLog.WriteLog(ssbylLog, 'R�cup�ration des donn�es cr��es ou modifi�es depuis le ' + BTPValues.LastSynchro, ServiceName_BTPY2, LogValues, 2);
  Result := True;
  TSlCacheGetY2Data.Clear;
  TSLUpdateInsertData.Clear;
  if Result then Result := GetFromDSType(wsdsCommon);
  if Result then Result := GetFromDSType(wsdsChoixCod);
  if Result then Result := GetFromDSType(wsdsChoixExt);
  if Result then Result := GetFromDSType(wsdsEstablishment);
  if Result then Result := GetFromDSType(wsdsSocietyParameters);
  if Result then Result := GetFromDSType(wsdsPaymentMode);
  if Result then Result := GetFromDSType(wsdsZipCode);
  if Result then Result := GetFromDSType(wsdsRecovery);
  if Result then Result := GetFromDSType(wsdsCountry);
  if Result then Result := GetFromDSType(wsdsCurrency);
  if Result then Result := GetFromDSType(wsdsCorrespondence);
  if Result then Result := GetFromDSType(wsdsFiscalYear);
  if Result then Result := GetFromDSType(wsdsJournal);
  if Result then Result := GetFromDSType(wsdsAccount);
  if Result then Result := GetFromDSType(wsdsPaymenChoice);
  if Result then Result := GetFromDSType(wsdsThird);
  if Result then Result := GetFromDSType(wsdsAnalyticalSection);
  if Result then Result := GetFromDSType(wsdsBankId);
  if Result then Result := GetFromDSType(wsdsChangeRate);
  if Result then Result := GetFromDSType(wsdsContact);
  if Result then Result := GetFromDSType(wsdsTaxRate);
  if Result then Result := InsertOrUpdateData;
  if Result then Result := GetFromDSType(wsdsAccForPayment);
//  if Result then Result := GetFromDSType(wsdsAccForPaymentOther);
  if Result then
    TServicesLog.WriteLog(ssbylLog, 'Fin de la r�cup�ration des donn�es cr��es ou modifi�es depuis le ' + BTPValues.LastSynchro, ServiceName_BTPY2, LogValues, 2)
  else
    TServicesLog.WriteLog(ssbylLog, Format('%s - Interruption de la r�cup�ration des donn�es cr��es ou modifi�es depuis le %s', [WSCDS_ErrorMsg, BTPValues.LastSynchro]), ServiceName_BTPY2, LogValues, 0);
end;

function TSvcSyncBTPY2Execute.SendData: boolean;
var
  Cpt             : integer;
  Cpt1            : integer;
  Qty             : integer;
  AdoQryEcr       : AdoQry;
  AdoQryAna       : AdoQry;
  TSlAccEntries   : TStringList;
  TSlSendEntries  : TStringList;
  LogValue        : string;
  LogValueExo     : string;
  LogValueAux     : string;
  LogValueJou     : string;
  LogValueNum     : string;
  LogValueDte     : string;

  function SendEntriesToY2 : boolean;
  var
    Y2EntriesNumber : integer;
    Msg             : string;
    ErrorMSg        : string;  
  begin
    Result := True;
    if TSlAccEntries.count > 0 then
    begin
      inc(Cpt1);
      TServicesLog.WriteLog(ssbylLog, Format('Pi�ce %s : ', [IntToStr(Cpt1)]), ServiceName_BTPY2, LogValues, 4);
      if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - Before SendAccountingEntries', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
      try
        Y2EntriesNumber := SendAccountingEntries(TSlAccEntries, ErrorMsg);
        try
          Result := (Y2EntriesNumber >= 0);
        finally
          if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - After SendAccountingEntries', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
          case Y2EntriesNumber of
            -1 : Msg := 'Erreur lors de la mise � jour ou la cr�ation dans la table d''historisation.';
            0  : begin
                  Msg := 'Erreur lors de l''envoi';
                  if ErrorMsg <> '' then
                    Msg := Format('%s (%s)', [Msg, ErrorMsg]);
                 end;
          else
            Msg := Format('Envoy�e avec succ�s et enregistr�e sous le n� %s.', [IntToStr(Y2EntriesNumber)]);
          end;
          TServicesLog.WriteLog(ssbylLog, Msg, ServiceName_BTPY2, LogValues, 5);
        end;
      except
        if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - Error on SendAccountingEntries', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
      end;
    end;
    TSlAccEntries.clear;
  end;

begin
  if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.SendData', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
  TServicesLog.WriteLog(ssbylLog, 'Envoi des donn�es cr��es ou modifi�es depuis le ' + BTPValues.LastSynchro, ServiceName_BTPY2, LogValues, 2);
  Result := True;
  { Recherche s'il existe des �critures non envoy�es }
  AdoQryBTP.TSLResult.Clear;
  AdoQryBTP.FieldsList := 'BE0_ENTITY,BE0_EXERCICE,BE0_JOURNAL,BE0_NUMEROPIECE';
  AdoQryBTP.Request    := Format('SELECT %s FROM BTPECRITURE WHERE BE0_REFERENCEY2 = 0', [AdoQryBTP.FieldsList]);
  try
    AdoQryBTP.SingleTableSelect;
    if AdoQryBTP.RecordCount > 0 then
    begin
      TServicesLog.WriteLog(ssbylLog, Format('%s pi�ce(s) comptable trouv�e(s).', [IntToStr(AdoQryBTP.RecordCount)]), ServiceName_BTPY2, LogValues, 3);
      AdoQryEcr := AdoQry.Create;
      try
        AdoQryAna := AdoQry.Create;
        try
          TSlAccEntries := TStringList.create;
          try
            TSlSendEntries := TStringList.create;
            try
              AdoQryEcr.ServerName := AdoQryBTP.ServerName;
              AdoQryEcr.DBName     := AdoQryBTP.DBName;
              AdoQryEcr.Qry        := TADOQuery.Create(nil);
              AdoQryEcr.FieldsList := Tools.GetFieldsListFromPrefix('E' {$IF defined(APPSRV)}, AdoQryBTP.ServerName, AdoQryBTP.DBName{$IFEND !APPSRV}) + ',R_CODEIBAN';
              AdoQryAna.ServerName := AdoQryBTP.ServerName;
              AdoQryAna.DBName     := AdoQryBTP.DBName;
              AdoQryAna.Qry        := TADOQuery.Create(nil);
              AdoQryAna.FieldsList := Tools.GetFieldsListFromPrefix('Y' {$IF defined(APPSRV)}, AdoQryBTP.ServerName, AdoQryBTP.DBName{$IFEND !APPSRV});
              Qty                  := 0;
              if AdoQryBTP.RecordCount > 0 then
              begin
                if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.SearchRelatedParameters', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
                for Cpt := 0 to pred(AdoQryBTP.TSLResult.Count) do
                begin
                  AdoQryEcr.TSLResult.Clear;
                  LoadEcrFromBE0(AdoQryEcr, AdoQryAna, AdoQryBTP.TSLResult[Cpt]); // Charge les �critures et analytique
                  if AdoQryEcr.RecordCount > 0 then
                  begin
                    inc(Qty);
                    { Si log, recherche les donn�es � ajouter dans le log }
                    if LogValues.LogLevel > 0 then
                    begin
                      LogValue    := SetDataToSend(AdoQryEcr.FieldsList, AdoQryEcr.tslresult[0]);
                      LogValueExo := Tools.GetStValueFromTSl(LogValue, 'E_EXERCICE');
                      LogValueAux := Tools.GetStValueFromTSl(LogValue, 'E_AUXILIAIRE');
                      LogValueJou := Tools.GetStValueFromTSl(LogValue, 'E_JOURNAL');
                      LogValueNum := Tools.GetStValueFromTSl(LogValue, 'E_NUMEROPIECE');
                      LogValueDte := Tools.GetStValueFromTSl(LogValue, 'E_DATECOMPTABLE');
                      TServicesLog.WriteLog(ssbylLog, Format('Pi�ce %s : Exercice "%s" - Auxiliaire "%s" - Journal "%s" - Date "%s" - Num�ro "%s"'
                                                          , [IntToStr(Qty) ,LogValueExo, LogValueAux, LogValueJou, LogValueDte, LogValueNum]), ServiceName_BTPY2, LogValues, 4);
                    end;
                    TSlAccEntries.Clear;
                    SetSendY2TSl(AdoQryEcr, AdoQryAna, TSlAccEntries);
                    SearchRelatedParameters(TSlAccEntries);
                    TSlSendEntries.Add('###NEWECR');
                    for Cpt1 := 0 to pred(TSlAccEntries.count) do
                      TSlSendEntries.Add(TSlAccEntries[Cpt1]);
                  end;
                end;
                SearchOthersParameters;   // Pr�paration autres param�tres
                Result := SendY2Settings; // Envoi du param�trage
                if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.SendEntriesToY2', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
                if Result then // Envoi des �critures
                begin
                  TServicesLog.WriteLog(ssbylLog, Format('Envoi des pi�ces comptables (%s pi�ce(s)).', [IntToStr(AdoQryBTP.RecordCount)]), ServiceName_BTPY2, LogValues, 3);
                  TSlAccEntries.clear;
                  Cpt1 := 0;
                  for Cpt := 0 to pred(TSlSendEntries.count) do
                  begin
                    if TSlSendEntries[Cpt] = '###NEWECR' then
                      Result := SendEntriesToY2
                    else
                      TSlAccEntries.Add(TSlSendEntries[Cpt]);
                  end;
                  if Result then
                    Result := SendEntriesToY2;
                end;
              end;
            finally
              FreeAndNil(TSlSendEntries);
            end;
          finally
            FreeAndNil(TSlAccEntries);
          end;
        finally
          AdoQryAna.Qry.Free;
          AdoQryAna.free;
        end;
      finally
        AdoQryEcr.Qry.Free;
        AdoQryEcr.free;
      end;
    end;
    AdoQryBTP.Reset;
  except
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.SendData. Erreur sur ex�cution de %s / %s / %s'
                                                     , [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryBTP.ServerName, AdoQryBTP.DBName, AdoQryBTP.Request]), ServiceName_BTPY2, LogValues, 0);
    Result := False;
    exit;
  end;
  TServicesLog.WriteLog(ssbylLog, 'Fin de l''envoi des donn�es cr��es ou modifi�es depuis le ' + BTPValues.LastSynchro, ServiceName_BTPY2, LogValues, 2);
end;

function TSvcSyncBTPY2Execute.ServiceExecute: Boolean;
var
  Cpt        : integer;
 {$IF defined(APPSRV)}
  AdoQryPAna : AdoQry;
 {$IFEND (APPSRV)}

  procedure InitConnectionData(Values: string);
  begin
    BTPValues.ConnectionName := Tools.ReadTokenSt_(Values, '=');
    BTPValues.UserAdmin      := Tools.ReadTokenSt_(Values, ';');
    BTPValues.Server         := Tools.ReadTokenSt_(Values, ';');
    BTPValues.DataBase       := Tools.ReadTokenSt_(Values, ';');
    BTPValues.LastSynchro    := Tools.ReadTokenSt_(Values, ';');
    Y2Values.ConnectionName  := BTPValues.ConnectionName;
    Y2Values.Server          := Tools.ReadTokenSt_(Values, ';');
    Y2Values.DataBase        := Tools.ReadTokenSt_(Values, ';');
    AdoQryY2.ServerName      := Y2Values.Server;
    AdoQryY2.DBName          := Y2Values.DataBase;
    AdoQryBTP.ServerName     := BTPValues.Server;
    AdoQryBTP.DBName         := BTPValues.DataBase;
  end;

begin
  Result := True;
  {$IFDEF TSTSRV}
  if not TServicesLog.CanExecuteFromPeriod(LogValues, ServiceName_BTPY2, True) then exit;
  {$ENDIF TSTSRV}
  if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.ServiceExecute', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
  LogFilePath := TServicesLog.CreateLog(LogValues, ServiceName_BTPY2); //LogsManagement;
  TServicesLog.WriteLog(ssbylLog, '', ServiceName_BTPY2, LogValues, 0, False);
  TServicesLog.WriteLog(ssbylLog, DupeString('*', 50), ServiceName_BTPY2, LogValues, 0);
  TServicesLog.WriteLog(ssbylLog, 'D�but d''ex�cution du service.', ServiceName_BTPY2, LogValues, 0);
  try
    for Cpt := 0 to pred(TSlConnectionValues.Count) do
    begin
      InitConnectionData(TSlConnectionValues[Cpt]);
      TServicesLog.WriteLog(ssbylLog, Format('Traitement de la connexion %s', [BTPValues.ConnectionName]), ServiceName_BTPY2, LogValues, 1);
      try
       {$IF defined(APPSRV)}
        AdoQryPAna := AdoQry.Create;
        try
          { Test les paramsocs de la base BTP }
          AdoQryPAna.ServerName  := AdoQryBTP.ServerName;
          AdoQryPAna.DBName      := AdoQryBTP.DBName;
          AdoQryPAna.FieldsList  := 'SOC_DATA';
          AdoQryPAna.Request     := Format('SELECT %s FROM PARAMSOC WHERE SOC_NOM IN (''%s'', ''%s'', ''%s'') ORDER BY SOC_NOM DESC', [AdoQryPAna.FieldsList, WSCDS_SocServer, WSCDS_SocNumPort, WSCDS_SocCegidDos]);
          AdoQryPAna.LogValues   := LogValues;
          AdoQryPAna.Qry         := TADOQuery.Create(nil);
          try
            AdoQryPAna.SingleTableSelect;
            if AdoQryPAna.RecordCount > 0 then
              Result := ((AdoQryPAna.TSLResult[0] <> '') and (AdoQryPAna.TSLResult[1] <> '') and (AdoQryPAna.TSLResult[2] <> ''))
            else
              Result := False;
          except
            if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - %s - TSvcSyncBTPY2Execute.ServiceExecute. Erreur sur ex�cution de %s / %s / %s'
                                                             , [WSCDS_DebugMsg, WSCDS_ErrorMsg, AdoQryPAna.ServerName, AdoQryPAna.DBName, AdoQryPAna.Request]), ServiceName_BTPY2, LogValues, 0);
            Result := False;
            exit;
          end;
        finally
           AdoQryPAna.Qry.Free;
           AdoQryPAna.Free;
        end;
        { Analytique }
        if Result then
          NewAna := (Tools.GetParamSocSecur_('SO_NEWANALYTIQUE', '-', AdoQryBTP.ServerName, AdoQryBTP.DBName) = 'X');
        {$ELSE !APPSRV}
        Result := ((GetParamSocSecur(WSCDS_SocServer, '') <> '') and (GetParamSocSecur(WSCDS_SocNumPort, '') <> '') and (GetParamSocSecur(WSCDS_SocCegidDos, '') <> ''));
        NewAna := GetParamSocSecur('SO_NEWANALYTIQUE', False);
        {$IFEND !APPSRV}
        if Result then
        begin
          if LogValues.Download then
            Result := GetData;
          if (Result) and (LogValues.Upload) then
            Result := SendData;
        end else
          TServicesLog.WriteLog(ssbylLog, Format('%s : Les �l�ments de connexion Y2 sont incorrects.', [WSCDS_ErrorMsg]), ServiceName_BTPY2, LogValues, 0);
      finally
        if Result then
          SetLastSyncIniFile;
        ClearValuesConnection;
      end;
    end;
  finally
    TServicesLog.WriteLog(ssbylLog, 'Fin d''ex�cution du service.', ServiceName_BTPY2, LogValues, 0);
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - End TSvcSyncBTPY2Execute.ServiceExecute', [WSCDS_DebugMsg]), ServiceName_BTPY2, LogValues, 0);
  end;
end;

function TSvcSyncBTPY2Execute.InitApplication : boolean;
begin
  Result := ReadSettings;
  if Result then
    Result := CreateMemoryCache;
  if Result then
    AddWindowsLog(True);
end;

end.

