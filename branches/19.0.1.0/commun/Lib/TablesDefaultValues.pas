(*
  Les méthodes GetFieldsList renvoient une liste de champs par rapport à un contexte défini dans le type ContextFieldsListType
  Les méthodes GetDefaultValue renvoient la valeur par défaut pour chaque champ défini dans le contexte icfltDefault   
*)
unit TablesDefaultValues;

interface

type

  ContextFieldsListType = (  cflttNone
                           , cfltProfile        // Champs définis dans les profils article
                           , cfltDefault        // Champs par défaut lors de la création d'un enregistrement
                           , cfltCatalogImport  // Champs pour l'import catalogue
                           );

  DefaultValuesGA = class
    class function GetFieldsList(cflType : ContextFieldsListType) : string;
    class function GetDefaultValue(FieldName : string) : string;
  end;

  DefaultValuesGCA = class
    class function GetFieldsList(cflType : ContextFieldsListType) : string;
    class function GetDefaultValue(FieldName : string) : string;
  end;

  DefaultValuesBFT = class
    class function GetFieldsList(cflType : ContextFieldsListType) : string;
  end;

  DefaultValuesBSF = class
    class function GetFieldsList(cflType : ContextFieldsListType) : string;
  end;

  DefaultValueGNE = class
    class function GetFieldsList(cflType: ContextFieldsListType) : string;
    class function GetDefaultValue(FieldName : string) : string;
  end;

  DefaultValueGNL = class
    class function GetFieldsList(cflType : ContextFieldsListType) : string;
    class function GetDefaultValue(FieldName : string) : string;
  end;

  DefaultValueBCB = class
    class function GetFieldsList(cflType : ContextFieldsListType) : string;
    class function GetDefaultValue(FieldName : string) : string;
  end;

implementation

uses
  CommonTools
  , Hent1
  , HeureUtil
  , SysUtils
  ;

{ DefaultValues }

class function DefaultValuesGA.GetFieldsList(cflType : ContextFieldsListType) : string;
begin
  case cflType of
    cfltProfile :
      Result := 'GA_FAMILLENIV1'
              + ';GA_FAMILLENIV2'
              + ';GA_FAMILLENIV3'
              + ';GA_COMPTAARTICLE'
              + ';GA_TENUESTOCK'
              + ';GA_CALCPRIXPR'
              + ';GA_COEFFG'
              + ';GA_DPRAUTO'
              + ';GA_LOT'
              + ';GA_NUMEROSERIE'
              + ';GA_CONTREMARQUE'
              + ';GA_REMISEPIED'
              + ';GA_REMISELIGNE'
              + ';GA_ESCOMPTABLE'
              + ';GA_FAMILLETAXE1'
              + ';GA_COMMISSIONNABLE'
              + ';GA_CALCPRIXHT'
              + ';GA_CALCPRIXTTC'
              + ';GA_COEFCALCHT'
              + ';GA_COEFCALCTTC'
              + ';GA_CALCAUTOHT'
              + ';GA_CALCAUTOTTC'
              + ';GA_TARIFARTICLE'
              + ';GA2_SOUSFAMTARART'
              + ';GA_PAYSORIGINE'
              + ';GA_ARRONDIPRIX'
              + ';GA_ARRONDIPRIXTTC'
              + ';GA_PRIXUNIQUE'
              + ';GA_COLLECTION'
              + ';GA_FOURNPRINC'
              + ';GA_LIBREART1'
              + ';GA_LIBREART2'
              + ';GA_LIBREART3'
              + ';GA_LIBREART4'
              + ';GA_LIBREART5'
              + ';GA_LIBREART6'
              + ';GA_LIBREART7'
              + ';GA_LIBREART8'
              + ';GA_LIBREART9'
              + ';GA_LIBREARTA'
              + ';GA_VALLIBRE1'
              + ';GA_VALLIBRE2'
              + ';GA_VALLIBRE3'
              + ';GA_DATELIBRE1'
              + ';GA_DATELIBRE2'
              + ';GA_DATELIBRE3'
              + ';GA_CHARLIBRE1'
              + ';GA_CHARLIBRE2'
              + ';GA_CHARLIBRE3'
              + ';GA_BOOLLIBRE1'
              + ';GA_BOOLLIBRE2'
              + ';GA_BOOLLIBRE3'
              ;
    cfltDefault :
      Result := 'GA_FERME'
              + ';GA_UTILISATEUR'
              + ';GA_CREATEUR'
              + ';GA_SOCIETE'
              + ';GA_STATUTART'
              + ';GA_DATECREATION'
              + ';GA_DATEMODIF'
              + ';GA_QUALIFMARGE'
              + ';GA_GEREANAL'
              + ';GA_INVISIBLEWEB'
              + ';GA_ACTIVITEREPRISE'
              + ';GA_REMISELIGNE'
              + ';GA_REMISEPIED'
              ;
    cfltCatalogImport :
      Result := 'GA_ARTICLE'
              + ';GA2_ARTICLE'
              + ';GA_CODEARTICLE'
              + ';GA2_CODEARTICLE'
              + ';GA_TYPEARTICLE'
              + ';GA_LIBELLE'
              + ';GA_COMPTAARTICLE'
              + ';GA_DATESUPPRESSION'
              + ';GA_TENUESTOCK'
              + ';GA_FAMILLETAXE1'
              + ';GA_REMISEPIED'
              + ';GA_ESCOMPTABLE'
              + ';GA_FOURNPRINC'
              + ';GA_CREERPAR'
              ;
  else
    Result := '';
  end;
end;

class function DefaultValuesGA.GetDefaultValue(FieldName: string) : string;
var
  FieldsList   : string;
  DefaultArray : array of string;
begin
  Result     := '';
  FieldsList := DefaultValuesGA.GetFieldsList(cfltDefault);
  if Tools.StringInList(FieldName, FieldsList) then
  begin
    SetLength(DefaultArray, Tools.CountOccurenceString(FieldsList, ';')+1);
    Tools.SetArray(DefaultArray, FieldsList);
    case Tools.CaseFromString(FieldName, DefaultArray) of
      {GA_FERME}                     0     : Result := '-';
      {GA_UTILISATEUR,GA_CREATEUR}   1,2   : Result := V_PGI.User;
      {GA_SOCIETE}                   3     : Result := V_PGI.CodeSociete;
      {GA_STATUTART}                 4     : Result := 'UNI';
      {GA_DATECREATION,GA_DATEMODIF} 5,6 : Result := Tools.CastDateTimeForQry(CurrentDate);
      {GA_QUALIFMARGE}               7     : Result := 'CO';
      {GA_GEREANAL}                  8     : Result := 'X';
      {GA_INVISIBLEWEB}              9     : Result := 'X';
      {GA_ACTIVITEREPRISE}           10    : Result := 'F';
      {GA_REMISELIGNE}               11    : Result := 'X';
      {GA_REMISEPIED}                12    : Result := 'X';
    end;
  end;
end;

{ DefaultValuesGCA }

class function DefaultValuesGCA.GetFieldsList(cflType : ContextFieldsListType) : string;
begin                   
  case cflType of
    cfltDefault :
      Result := 'GCA_DATEREFERENCE'
              + ';GCA_DATECREATION'
              + ';GCA_DATEMODIF'
              + ';GCA_DATEDEB'
              + ';GCA_DATEFIN'
              + ';GCA_CREATEUR'
              + ';GCA_UTILISATEUR'
              + ';GCA_DATELIBRE1'
              + ';GCA_DATELIBRE2'
              + ';GCA_DATELIBRE3'
              + ';GCA_CREERPAR'
              + ';GCA_DATESUP'
              ;
    cfltCatalogImport :
      Result := 'GCA_REFERENCE'
              + ';GCA_TIERS'
              + ';GCA_DATESUP'
              + ';GCA_ARTICLE'
              ;
  else
    Result := '';
  end;
end;

class function DefaultValuesGCA.GetDefaultValue(FieldName: string): string;
var
  FieldsList   : string;
  DefaultArray : array of string;
begin
  Result     := '';
  FieldsList := DefaultValuesGCA.GetFieldsList(cfltDefault);
  if Tools.StringInList(FieldName, FieldsList) then
  begin
    SetLength(DefaultArray, Tools.CountOccurenceString(FieldsList, ';')+1);
    Tools.SetArray(DefaultArray, FieldsList);
    case Tools.CaseFromString(FieldName, DefaultArray) of
      {GCA_DATEREFERENCE,CREATION,MODIF} 0..2 : Result := Tools.CastDateTimeForQry(CurrentDate);
      {GCA_DATEDEB,GCA_DATEFIN}          3,4  : Result := Tools.CastDateTimeForQry(IDate1900);
      {GCA_CREATEUR,GCA_UTILISATEUR}     5,6  : Result := V_PGI.User;
      {GCA_DATELIBRE1,2,3}               7..9 : Result := Tools.CastDateTimeForQry(IDate1900);
      {GCA_CREERPAR}                     10   : Result := 'IMP';
      {GCA_DATESUP}                      11   : Result := Tools.CastDateTimeForQry(IDate2099);
    end;
  end;
end;

{ DefaultValuesBFT }

class function DefaultValuesBFT.GetFieldsList(cflType : ContextFieldsListType) : string;
begin
  case cflType of
    cfltCatalogImport :
      Result := 'BFT_FAMILLETARIF'
              + ';BFT_LIBELLE'
              ;
  else
    Result := '';
  end;
end;

{ DefaultValuesBSF }

class function DefaultValuesBSF.GetFieldsList(cflType : ContextFieldsListType) : string;
begin
  case cflType of
    cfltCatalogImport :
      Result := 'BSF_FAMILLETARIF'
              + ';BSF_SOUSFAMTARART'
              + ';BSF_LIBELLE'
              ;
  else
    Result := '';
  end;

end;

{ DefaultValueGNE }

class function DefaultValueGNE.GetFieldsList(cflType: ContextFieldsListType) : string;
begin
  case cflType of
    cfltDefault :
      Result := 'GNE_QTEDUDETAIL'
              + ';GNE_DATECREATION'
              + ';GNE_DATEMODIF'
              ;
    cfltCatalogImport :
      Result := 'GNE_NOMENCLATURE'
              + ';GNE_LIBELLE'
              + ';GNE_ARTICLE'
              ;
  else
    Result := '';
  end;
end;

class function DefaultValueGNE.GetDefaultValue(FieldName: string): string;
var
  FieldsList   : string;
  DefaultArray : array of string;
begin
  Result     := '';
  FieldsList := DefaultValueGNE.GetFieldsList(cfltDefault);
  if Tools.StringInList(FieldName, FieldsList) then
  begin
    SetLength(DefaultArray, Tools.CountOccurenceString(FieldsList, ';')+1);
    Tools.SetArray(DefaultArray, DefaultValueGNE.GetFieldsList(cfltDefault));
    case Tools.CaseFromString(FieldName, DefaultArray) of
      {GNE_QTEDUDETAIL}        0   : Result := '1';
      {GNE_DATECREATION,MODIF} 1,2 : Result := Tools.CastDateTimeForQry(CurrentDate);
    end;
  end;
end;

{ DefaultValueGNL }

class function DefaultValueGNL.GetFieldsList(cflType: ContextFieldsListType): string;
begin
  case cflType of
    cfltDefault :
      Result := 'GNL_JOKER'
              + ';GNL_DATECREATION'
              + ';GNL_DATEMODIF'
              + ';GNL_CREATEUR'
              + ';GNL_UTILISATEUR'
              ;
    cfltCatalogImport :
      Result := 'GNL_NOMENCLATURE'
              + ';GNL_NUMLIGNE'
              + ';GNL_LIBELLE'
              + ';GNL_CODEARTICLE'
              + ';GNL_ARTICLE'
              + ';GNL_QTE'
              ;
  else
    Result := '';
  end;
end;

class function DefaultValueGNL.GetDefaultValue(FieldName: string): string;
var
  FieldsList   : string;
  DefaultArray : array of string;
begin
  Result     := '';
  FieldsList := DefaultValueGNL.GetFieldsList(cfltDefault);
  if Tools.StringInList(FieldName, FieldsList) then
  begin
    SetLength(DefaultArray, Tools.CountOccurenceString(FieldsList, ';')+1);
    Tools.SetArray(DefaultArray, DefaultValueGNL.GetFieldsList(cfltDefault));
    case Tools.CaseFromString(FieldName, DefaultArray) of
      {GNL_JOKER}                      0 :   Result := 'N';
      {GNL_DATECREATION,GNL_DATEMODIF} 1,2 : Result := Tools.CastDateTimeForQry(CurrentDate);
      {GNL_CREATEUR,GNL_UTILISATEUR}   3,4 : Result := V_PGI.User;
    end;
  end;
end;

{ DefaultValueBCB }

class function DefaultValueBCB.GetFieldsList(cflType: ContextFieldsListType): string;
begin
  case cflType of
    cfltDefault :
      Result := 'BCB_QUALIFCODEBARRE'
              ;
    cfltCatalogImport :
      Result := 'BCB_NATURECAB'
              + ';BCB_IDENTIFCAB'
              + ';BCB_CABPRINCIPAL'
              + ';BCB_CODEBARRE'
              ;
  else
    Result := '';
  end;
end;

class function DefaultValueBCB.GetDefaultValue(FieldName: string): string;
var
  FieldsList   : string;
  DefaultArray : array of string;
begin
  Result     := '';
  FieldsList := DefaultValueBCB.GetFieldsList(cfltDefault);
  if Tools.StringInList(FieldName, FieldsList) then
  begin
    SetLength(DefaultArray, Tools.CountOccurenceString(FieldsList, ';')+1);
    Tools.SetArray(DefaultArray, DefaultValueBCB.GetFieldsList(cfltDefault));
    case Tools.CaseFromString(FieldName, DefaultArray) of
      {BCB_QUALIFCODEBARRE} 0 :   Result := '128';
    end;
  end;
end;

end.
