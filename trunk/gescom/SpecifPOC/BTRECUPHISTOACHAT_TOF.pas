{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/07/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECUPHISTOACHAT ()
Mots clefs ... : TOF;BTRECUPHISTOACHAT
*****************************************************************}
Unit BTRECUPHISTOACHAT_TOF ;

Interface

Uses
  StdCtrls
  , Controls
  , Classes
  {$IFNDEF EAGLCLIENT}
  , db
  , uDbxDataSet
  , mul
  , Fe_Main
  {$ELSE EAGLCLIENT}
  , eMul
  , uTob
  {$ENDIF EAGLCLIENT}
  , forms
  , sysutils
  , ComCtrls
  , HCtrls
  , HEnt1
  , HMsgBox
  , UTOF
  ;

function BTLanceFiche_POCRecupHistoAchat(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_BTRECUPHISTOACHAT = Class (TOF)
  private
    FicImport : THEdit;

    procedure bValider_OnClick(Sender : TObject);
    procedure JnalEvent_OnClick(Sender : TObject);
    function DataIntegration : boolean;

  public
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

uses
  HTB97
  , UtilGC
  , UtilXlsBTP
  , CommonTools
  , GerePiece
  , Utob
  , UConnectWSConst
  , ParamSoc
  , EntGC
  , Ent1
  , ed_Tools
  , FormsName
  , DateUtils
  ;

function BTLanceFiche_POCRecupHistoAchat(Nat, Cod : String ; Range,Lequel,Argument : string) : string;
begin
  if (Nat = '') and (Cod = '') then
    Result := ''
  else
  begin
    V_PGI.ZoomOle := True;
    Result := AGLLanceFiche(Nat, Cod, Range, Lequel, Argument);
    V_PGI.ZoomOle := False;
  end;
end;

procedure TOF_BTRECUPHISTOACHAT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPHISTOACHAT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPHISTOACHAT.OnArgument (S : String ) ;
begin
  Inherited ;
  FicImport := THEdit(GetControl('FICIMPORT'));
  TToolbarButton97(GetControl('BVALIDER')).OnClick  := bValider_OnClick;
  TToolbarButton97(GetControl('JNALEVENT')).OnClick := JnalEvent_OnClick;
end ;

procedure TOF_BTRECUPHISTOACHAT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPHISTOACHAT.bValider_OnClick(Sender: TObject);
var
  PathFile    : string;
  CanContinue : boolean;
  WinExcel    : OleVariant;
  NewInst     : Boolean;
begin
  PathFile    := FicImport.Text;
  CanContinue := False;
  if PathFile = '' then
    LastError := 1
  else if not FileExists(PathFile) then
    LastError := 2
  else if not OpenExcel(true, WinExcel, NewInst) then
  begin
    LastError := 3;
  end else
    CanContinue := (PGIAsk(Format('Veuillez confirmer la récupération de l''historique d''achat depuis %s.', [PathFile])) = mrYes);
  ExcelClose(WinExcel);
  if CanContinue then
    DataIntegration
  else
  begin
    TForm(Ecran).ModalResult := 0;
    if LastError > 0 then
    begin
      case LastError of
        1 : LastErrorMsg := 'Veuillez sélectionner un fichier.';
        2 : LastErrorMsg := Format('Le fichier %s n''existe pas.', [PathFile]);
        3 : LastErrorMsg := 'Erreur lors de l''ouverture d''Excel.';
      end;
      PGIError(LastErrorMsg, Ecran.Caption);
      Exit;
    end;
  end;
end;

procedure TOF_BTRECUPHISTOACHAT.JnalEvent_OnClick(Sender : TObject);
begin
  OpenForm.JnalEvent('TYPEEVENT=RDD;LABEL=' + Ecran.Caption);
end;

function TOF_BTRECUPHISTOACHAT.DataIntegration: boolean;
var
  LastDocNumber        : string;
  LastDocNumberOnError : string;
  ValAff3              : string;
  ValIntRef            : string;
  ValThird             : string;
  ValExtRef            : string;
  ValFLevel2           : string;
  ValRefBSV            : string;
  ValTemp              : string;
  Msg                  : string;
  PathFileName         : string;              
  NewPathFileName      : string;              
  ValDate              : TDateTime;
  StartDate            : TDateTime;
  EndDate              : TDateTime;
  ValUPrice            : double;
  ValQty               : double;
  LinesQty             : integer;
  Cpt                  : integer;   
  LinesCpt             : integer;
  NumError             : integer;
  DocQty               : integer;
  TobCache             : TOB;
  TobGP                : TOB;
  TobGL                : TOB;
  TslReport            : TStringList;
  WinExcel             : OleVariant;
  WorkTab              : Variant;
  CurrentTab           : Variant;
  NewInst              : Boolean;
const
  TabName   = 'Récapitulatif FA';
  CellQty   = 10;
  cnAff1    = 1;
  cnIntRef  = 2;
  cnThird   = 3;
  cnExtRef  = 4;
  cnDate    = 6;
  cnQty     = 7;
  cnUPrice  = 8;
  cnFLevel2 = 9;
  cnBsvREF  = 15;
  TyThird   = 'THIRD';
  tyCase    = 'CASE';
  tyItem    = 'ITEM';
  tyFamily  = 'FAMILY';

  function IsThirdExist(ThirdCode : string) : integer;
  var
    TobT : TOB;
    Sql  : string;
    Qry  : TQuery;
  begin
    Result := 0;
    TobT := TobCache.FindFirst(['TYPECACHE', 'KEY'], [TyThird, ThirdCode], True);
    if not assigned(TobT) then
    begin
      Sql := 'SELECT T_TIERS, YTC_TEXTELIBRE1'
           + ' FROM TIERS'
           + ' JOIN TIERSCOMPL ON YTC_TIERS = T_TIERS'
           + ' WHERE T_TIERS = "' + ThirdCode + '"';
      Qry := OpenSQL(Sql, True);
      try
        if Qry.Eof then
          Result := 1;
        TobT := TOB.Create('_VALUES', TobCache, -1);
        TobT.AddChampSupValeur('TYPECACHE', TyThird);
        TobT.AddChampSupValeur('FINDIT'   , Tools.iif(Result = 0, 'X', '-'));
        TobT.AddChampSupValeur('KEY'      , ThirdCode);
        TobT.AddChampSupValeur('FIELD1'   , Tools.iif(Result = 0, Qry.FindField('YTC_TEXTELIBRE1').AsString , WSCDS_EmptyValue));
        TobT.AddChampSupValeur('FIELD2'   , Tools.iif(Result = 0, Qry.FindField('T_TIERS').AsString         , ThirdCode));
        TobT.AddChampSupValeur('FIELD3'   , '');
        TobT.AddChampSupValeur('FIELD4'   , '');
        TobT.AddChampSupValeur('FIELD5'   , '');
        if (Result = 0) and (Qry.FindField('YTC_TEXTELIBRE1').AsString = '') then
          Result := 2;
      finally
        Ferme(Qry);
      end;
    end else
    begin
      if (Tobt.GetSTring('FIELD1') = WSCDS_EmptyValue) or (Tobt.GetSTring('FIELD1') = '') then // YTC_TEXTELIBRE1 est vide
        Result := 2
      else if TobT.GetBoolean('FINDIT') then              // Le tiers existe
        Result := 0
      else                                                // Le tiers n'existe pas
        Result := 1;
    end;
  end;

  function IsItemExist(ItemCode : string) : boolean;
  var
    TobT : TOB;
    Sql  : string;
    Qry  : TQuery;
  begin
    TobT := TobCache.FindFirst(['TYPECACHE', 'KEY'], [tyItem, ItemCode], True);
    if not assigned(TobT) then
    begin
      Sql := 'SELECT GA_ARTICLE, GA_CODEARTICLE'
           + ' FROM ARTICLE'
           + ' WHERE GA_CODEARTICLE = "' + ItemCode + '"';
      Qry := OpenSQL(Sql, True);
      try
        Result := (not Qry.Eof);
        TobT := TOB.Create('_VALUES', TobCache, -1);
        TobT.AddChampSupValeur('TYPECACHE', tyItem);
        TobT.AddChampSupValeur('KEY'      , ItemCode);
        TobT.AddChampSupValeur('FINDIT'   , Tools.iif(Result, 'X', '-'));
        TobT.AddChampSupValeur('FIELD1'   , Tools.iif(Result, Qry.FindField('GA_ARTICLE').AsString    , WSCDS_EmptyValue));
        TobT.AddChampSupValeur('FIELD2'   , Tools.iif(Result, Qry.FindField('GA_CODEARTICLE').AsString, ItemCode));
        TobT.AddChampSupValeur('FIELD3'   , '');
        TobT.AddChampSupValeur('FIELD4'   , '');
        TobT.AddChampSupValeur('FIELD5'   , '');
      finally
        Ferme(Qry);
      end;
    end else
      Result := TobT.GetBoolean('FINDIT');
  end;

  function IsCaseExist(CasePart1 : string) : boolean;
  var
    TobT : TOB;
    Sql    : string;
    Qry    : TQuery;
  begin
    TobT := TobCache.FindFirst(['TYPECACHE', 'KEY'], [tyCase, CasePart1], True);
    if not assigned(TobT) then
    begin
      Sql := 'SELECT AFF_AFFAIRE, AFF_AFFAIRE0, AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3'
           + ' FROM AFFAIRE'
           + ' WHERE AFF_AFFAIRE1 = "' + CasePart1 + '"';
      Qry := OpenSQL(Sql, True);
      try
        Result := (not Qry.Eof);
        TobT := TOB.Create('_VALUES', TobCache, -1);
        TobT.AddChampSupValeur('TYPECACHE', tyCase);
        TobT.AddChampSupValeur('FINDIT'   , Tools.iif(Result, 'X', '-'));
        TobT.AddChampSupValeur('KEY'      , CasePart1);
        TobT.AddChampSupValeur('FIELD1'   , Tools.iif(Result, Qry.FindField('AFF_AFFAIRE').AsString , WSCDS_EmptyValue));
        TobT.AddChampSupValeur('FIELD2'   , Tools.iif(Result, Qry.FindField('AFF_AFFAIRE0').AsString, WSCDS_EmptyValue));
        TobT.AddChampSupValeur('FIELD3'   , Tools.iif(Result, Qry.FindField('AFF_AFFAIRE1').AsString, CasePart1));
        TobT.AddChampSupValeur('FIELD4'   , Tools.iif(Result, Qry.FindField('AFF_AFFAIRE2').AsString, WSCDS_EmptyValue));
        TobT.AddChampSupValeur('FIELD5'   , Tools.iif(Result, Qry.FindField('AFF_AFFAIRE3').AsString, WSCDS_EmptyValue));
      finally
        Ferme(Qry);
      end;
    end else
      Result := (TobT.GetBoolean('FINDIT'));
  end;

  function IsFamilyExist(FamilyCode : string) : boolean;
  var
    TobT : TOB;
    Sql  : string;
    Qry  : TQuery;
  begin
    TobT := TobCache.FindFirst(['TYPECACHE', 'KEY'], [tyFamily, FamilyCode], True);
    if not assigned(TobT) then
    begin
      Sql := 'SELECT CC_CODE, CC_LIBRE'
           + ' FROM CHOIXCOD'
           + ' WHERE CC_TYPE = "FN2"'
           + '   AND CC_CODE = "' + FamilyCode + '"';
      Qry := OpenSQL(Sql, True);
      try
        Result := (not Qry.Eof);
        TobT := TOB.Create('_VALUES', TobCache, -1);
        TobT.AddChampSupValeur('TYPECACHE', tyFamily);
        TobT.AddChampSupValeur('FINDIT'   , Tools.iif(Result, 'X', '-'));
        TobT.AddChampSupValeur('KEY'      , FamilyCode);
        TobT.AddChampSupValeur('FIELD1'   , Tools.iif(Result, Qry.FindField('CC_LIBRE').AsString, WSCDS_EmptyValue));
        TobT.AddChampSupValeur('FIELD2'   , Tools.iif(Result, Qry.FindField('CC_CODE').AsString , FamilyCode));
        TobT.AddChampSupValeur('FIELD3'   , '');
        TobT.AddChampSupValeur('FIELD4'   , '');
        TobT.AddChampSupValeur('FIELD5'   , '');
      finally
        Ferme(Qry);
      end;
    end else
      Result := TobT.GetBoolean('FINDIT');
  end;

  function GetValueFromCache(CacheType : string; KeyValue : string; FieldName : string='FIELD1') : string;
  var
    TobL : TOB;
  begin
    TobL := TobCache.FindFirst(['TYPECACHE', 'KEY'], [CacheType, KeyValue], True);
    if assigned(TobL) then
      Result := TobL.GetString(FieldName)
    else
      Result := '';
  end;

  procedure EmptyTobGP;
  begin
    TobGP.ClearDetail;
    TobGP.SetString('NATUREPIECEG' , '');
    TobGP.SetString('TIERS'        , '');
    TobGP.SetString('ETABLISSEMENT', '');
    TobGP.SetString('DOMAINE'      , '');
    TobGP.SetString('DEPOT'        , '');
    TobGP.SetString('REFINTERNE'   , '');
    TobGP.SetString('REFEXTERNE'   , '');
    TobGP.SetString('AFFAIRE'      , '');
    TobGP.SetString('AFFAIRE0'     , '');
    TobGP.SetString('AFFAIRE1'     , '');
    TobGP.SetString('AFFAIRE2'     , '');
    TobGP.SetString('AFFAIRE3'     , '');
    TobGP.SetString('BSVREF'       , '');
  end;

  function IsLineValuesOk : integer; 
  var
    ItemCode : string;
  begin
    Result := 0;
    if ValThird = '' then // Le tiers est vide
      Result := 7;
    if Result = 0 then
      Result := IsThirdExist(ValThird);    // 1 : Tiers inexistant, 2 : Pas d'article par défaut dans le tiers
    if Result = 0 then
    begin
      if not IsCaseExist(ValAff3) then   // 3 : Affaire inexistante
        Result := 3;
    end;
    if Result = 0 then                   // 4 : Article inexistant
    begin
      ItemCode := GetValueFromCache(TyThird, ValThird);
      if not IsItemExist(ItemCode) then
        Result := 4;
    end;
    if Result = 0 then                   // 5 : Date non renseignée ou 01/01/1900
    begin
      if ValDate = IDate1900 then
        Result := 5;
    end;
    if Result = 0 then                   // 6 : La famille niveau 2 n'existe pas
    begin
      if not IsFamilyExist(ValFLevel2) then
        Result := 6;
    end;
  end;

  procedure AddLine;
  begin
    TobGL := TOB.Create('_DATAL', TobGP, -1);
    TobGL.AddChampSupValeur('TYPELIGNE'    , 'ART');
    TobGL.AddChampSupValeur('CODEARTICLE'  , GetValueFromCache(TyThird, ValThird));
    TobGL.AddChampSupValeur('ARTICLE'      , GetValueFromCache(tyItem, TobGL.GetString('CODEARTICLE')));
    TobGL.AddChampSupValeur('QTEFACT'      , ValQty);
    TobGL.AddChampSupValeur('QUALIFQTE'    , '');
    TobGL.AddChampSupValeur('DATEPIECE'    , ValDate);
    TobGL.AddChampSupValeur('DATELIVRAISON', ValDate);
    TobGL.AddChampSupValeur('DEPOT'        , '');
    TobGL.AddChampSupValeur('FAMILLENIV1'  , GetValueFromCache(tyFamily, ValFLevel2));
    TobGL.AddChampSupValeur('FAMILLENIV2'  , ValFLevel2);
    TobGL.AddChampSupValeur('PUHTDEV'      , ValUPrice);
    TobGL.AddChampSupValeur('DEPOT'        , VH_GC.GCDepotDefaut);
    TobGL.AddChampSupValeur('AVECPRIX'     , 'X');
    TobGL.AddChampSupValeur('DOMAINE'      , '');
  end;

  procedure AddDoc;
  begin
    if TobGP.Detail.count > 0 then // Changement de pièce et il existe des lignes, création de la pièce et vide la tob
    begin
      if CreatePieceFromTob(TobGP) then
        TslReport.Add(Format('Pièce %s correctement intégrée (%s n° %s).', [TobGP.GetString('REFINTERNE'), TobGP.GetString('GP_NATUREPIECEG'), TobGP.GetString('GP_NUMERO')]))
      else
        TslReport.Add(Format('Erreur lors de la création de la pièce %s.', [TobGP.GetString('REFINTERNE')]));
      EmptyTobGP;
    end;
    LastDocNumber := ValIntRef;
    TobGP.SetString('NATUREPIECEG' , 'FF');
    TobGP.SetDateTime('DATEPIECE'  , ValDate);
    TobGP.SetString('TIERS'        , ValThird);
    TobGP.SetString('ETABLISSEMENT', VH^.EtablisDefaut);
    TobGP.SetString('DOMAINE'      , '');
    TobGP.SetString('DEPOT'        , VH_GC.GCDepotDefaut);
    TobGP.SetString('REFINTERNE'   , ValIntRef);
    TobGP.SetString('REFEXTERNE'   , ValExtRef);
    TobGP.SetString('BSVREF'       , ValRefBSV);
    TobGP.SetString('AFFAIRE'      , GetValueFromCache(tyCase, ValAff3, 'FIELD1'));
    TobGP.SetString('AFFAIRE0'     , GetValueFromCache(tyCase, ValAff3, 'FIELD2'));
    TobGP.SetString('AFFAIRE1'     , GetValueFromCache(tyCase, ValAff3, 'FIELD3'));
    TobGP.SetString('AFFAIRE2'     , GetValueFromCache(tyCase, ValAff3, 'FIELD4'));
    TobGP.SetString('AFFAIRE3'     , GetValueFromCache(tyCase, ValAff3, 'FIELD5'));
    TobGP.SetString('HORSCOMPTA'   , 'X');
    AddLine;
  end;

  procedure SetTempValues;
  begin
    ValAff3    := GetExcelText(CurrentTab, LinesCpt, cnAff1);             // GP_AFFAIRE3
    ValIntRef  := GetExcelText(CurrentTab, LinesCpt, cnIntRef);           // GP_REFINTERNE
    ValThird   := GetExcelText(CurrentTab, LinesCpt, cnThird);            // GP_TIERS
    ValExtRef  := GetExcelText(CurrentTab, LinesCpt, cnExtRef);           // GP_REFEXTERNE
    ValRefBSV  := GetExcelText(CurrentTab, LinesCpt, cnBsvREF);           // GP_BSVREF
    ValFLevel2 := GetExcelText(CurrentTab, LinesCpt, cnFLevel2);          // GL_FAMILLENIV2
    ValTemp    := GetExcelText(CurrentTab, LinesCpt, cnQty);              // GL_QTEFACT
    ValTemp    := Tools.iif(ValTemp = '', '1', ValTemp);
    ValQty     := StrToFloat(ValTemp);
    ValTemp    := GetExcelText(CurrentTab, LinesCpt, cnDate);             // xx_DATEPIECE
    ValTemp    := Tools.iif(ValTemp = '', DateToStr(IDate1900), ValTemp);
    ValDate    := StrToDate(ValTemp);
    ValTemp    := GetExcelText(CurrentTab, LinesCpt, cnUPrice);           // GL_PUHTDEV
    ValTemp    := Tools.iif(pos(' €', ValTemp) > 0, copy(ValTemp, 1, pos(' €', ValTemp)-1), ValTemp);
    ValUPrice  := Tools.iif(ValTemp = '', 0, Valeur(ValTemp));
  end;

  procedure AddGPSuppFields;
  begin
    TobGP.AddChampSup('NATUREPIECEG' , False);
    TobGP.AddChampSup('DATEPIECE'    , False);
    TobGP.AddChampSup('TIERS'        , False);
    TobGP.AddChampSup('ETABLISSEMENT', False);
    TobGP.AddChampSup('DOMAINE'      , False);
    TobGP.AddChampSup('DEPOT'        , False);
    TobGP.AddChampSup('REFINTERNE'   , False);
    TobGP.AddChampSup('REFEXTERNE'   , False);
    TobGP.AddChampSup('AFFAIRE'      , False);
    TobGP.AddChampSup('AFFAIRE0'     , False);
    TobGP.AddChampSup('AFFAIRE1'     , False);
    TobGP.AddChampSup('AFFAIRE2'     , False);
    TobGP.AddChampSup('AFFAIRE3'     , False);
    TobGP.AddChampSup('HORSCOMPTA'   , False);
    TobGP.AddChampSup('DOMAINE'      , False);
    TobGP.AddChampSup('BSVREF'       , False);
  end;

begin
  StartDate := Now;
  Result    := True;
  TslReport := TStringList.Create;
  try
    TobCache := TOB.Create('_CACHEAFF', nil, -1);
    try
      TobGP := TOB.Create('_DATA', nil, -1);
      try
        AddGPSuppFields;
        OpenExcel(True, WinExcel, NewInst);
        DocQty := 0;
        try
          LastDocNumber        := '';
          LastDocNumberOnError := '';
          Cpt                  := 0;
          LinesQty             := 0;
          LinesCpt             := 6;
          WorkTab              := OpenWorkBook(FicImport.Text, WinExcel);
          CurrentTab           := SelectSheet(WorkTab, TabName);
          InitMoveProgressForm(nil, 'Ouverture du fichier en cours.', Ecran.Caption, LinesQty, False, True);
          try
            while GetExcelText(CurrentTab, LinesCpt, cnIntRef) <> '' do
            begin
              inc(LinesQty); // Nbre de ligne total
              if LastDocNumber <> GetExcelText(CurrentTab, LinesCpt, cnIntRef) then // Nbre de pièce total
              begin
                inc(DocQty);
                LastDocNumber := GetExcelText(CurrentTab, LinesCpt, cnIntRef);
              end;
              inc(LinesCpt);
            end;
          finally
            LinesCpt := 6;
            FiniMoveProgressForm;
          end;
          InitMoveProgressForm(nil, 'Traitement en cours.', Ecran.Caption, LinesQty, False, True);
          try
            LastDocNumber := '';
            while GetExcelText(CurrentTab, LinesCpt, cnIntRef) <> '' do
            begin
              inc(Cpt);
              MoveCurProgressForm(Format('%s/%s', [IntToStr(Cpt), IntToStr(LinesQty)]));
              SetTempValues;
              NumError := IsLineValuesOk;
              if (NumError = 0) and (ValIntRef <> LastDocNumberOnError) then
              begin
                if ValIntRef <> LastDocNumber then
                  AddDoc
                else
                  AddLine;
              end else
              begin
                if ValIntRef <> LastDocNumber then // La ligne en erreur est peut être différente de la pièce précédente
                  AddDoc;
                LastDocNumberOnError := ValIntRef;
                EmptyTobGP;
                Msg := Format('Pièce %s non intégrée.', [LastDocNumberOnError]);
                case NumError of
                  0 : TslReport.Add(Msg);
                  1 : TslReport.Add(Format('%s Le tiers ou les données complémentaires de %s n''existe pas.', [Msg, ValThird]));
                  2 : TslReport.Add(Format('%s Le tiers %s n''a pas d''article générique.'                  , [Msg, ValThird]));
                  3 : TslReport.Add(Format('%s Le numéro de chantier %s est inexistant.'                    , [Msg, ValAff3]));
                  4 : TslReport.Add(Format('%s L''article générique du tiers %s est inexistant.'            , [Msg, ValThird]));
                  5 : TslReport.Add(Format('%s La date de pièce %s n''est pas valide.'                      , [Msg, DateToStr(ValDate)]));
                  6 : TslReport.Add(Format('%s La famille %s est inexistante.'                              , [Msg, ValFLevel2]));
                  7 : TslReport.Add(Format('%s Le tiers est vide.'                                          , [Msg, ValThird]));
                end;
              end;
              inc(LinesCpt);
            end;
            AddDoc;
          finally
            FiniMoveProgressForm;
          end;
        finally
          ExcelClose(WinExcel);
          PathFileName    := FicImport.Text;
          NewPathFileName := ExtractFilePath(PathFileName) + 'IntégréLe_' + FormatDateTime('yyymmdd_hhmm', Now) + '_' + ExtractFileName(PathFileName);
          RenameFile(PathFileName, NewPathFileName);
          EndDate := Now;
          MAJJnalEvent('RDD', 'OK', Ecran.Caption, Format('Début du traitement : %s%sFin du traitement : %s%sDepuis le fichier %s%sRenommé en %s%sNombre de pièces : %s%s%s%s'
                                                          , [DatetimeToStr(StartDate), #13#10, DatetimeToStr(EndDate), #13#10, PathFileName, #13#10, NewPathFileName, #13#10, IntToStr(DocQty), #13#10, #13#10, TslReport.Text]));
          PGIBox('Traitement teminé.');
          FicImport.Text := '';
          TForm(Ecran).ModalResult := 0;
        end;
      finally
        FreeAndNil(TobGP);
      end;
    finally
      FreeAndNil(TobCache);
    end;
  finally
    FreeAndNil(TslReport);
  end;
end;

Initialization
  registerclasses ( [ TOF_BTRECUPHISTOACHAT ] ) ; 
end.

