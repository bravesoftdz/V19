unit ErrorsManagement;

interface

type

  T_EMContext  = (temc_none, temc_DocValid);
  T_EMTypMsg   = (temtm_none, temtm_Information, temtm_Warning, temtm_Error);
  T_EMStartMsg = (temsm_none, temsm_Delete, temsm_Update, temsm_Inser, temsm_Recalc, temsm_Number, temsm_Load, temsm_Cpta, temsm_CptaS, temsm_Generate);

  TUtilErrorsManagement = class (TObject)
    class procedure ShowGenericMessage(pReinitAfter : boolean=True);
    class function GetGenericMessage(pReinitAfter : boolean=True) : string;
    class procedure SetGenericMessage(TemErrNumber : integer; pErrorMsg : string=''; pCriticality : T_EMTypMsg=temtm_Error; pReinitAfter : boolean=True);
    class function GetStartMsg(TypeStartMsg : T_EMStartMsg) : string;
    class function GetDeleteError : string;
    class function GetUpdateError : string;
    class function GetInsertError : string;
    class function GetRecalcError : string;
    class function GetNumberError : string;
    class function GetLoadError   : string;
    class function GetCptaError   : string;
    class function GetCptaSError  : string;
    class function GetGeneratEror : string;
  end;

  TErrorsManagement = class (Tobject)
  private
    ArrOfMsg     : array[1..100] of string;
    fContext     : T_EMContext;
    fErrorNumber : integer;
    fErrorMsg    : string;
    fCriticality : T_EMTypMsg;
    fReinitAfter : boolean;

    procedure SetArrOfMsg;
    function GetContextCaption : string;
    function MessageAlreadyCalculate : boolean;

  protected
    procedure SetContext(pContext : T_EMContext);

  public
    property Context     : T_EMContext read fContext     write SetContext;
    property ErrorNumber : integer     read fErrorNumber write fErrorNumber;
    property Criticality : T_EMTypMsg  read fCriticality write fCriticality;
    property ReinitAfter : boolean     read fReinitAfter write fReinitAfter;
    property ErrorMsg    : string      read fErrorMsg    write fErrorMsg;

    Constructor Create(pContext : T_EMContext);
    Destructor Destroy ;   override;
    procedure Init(pContext : T_EMContext);
    function GetMessage : string;
    procedure ShowError;
    procedure WriteErrorInLog;
    procedure Reinit;
    procedure Duplicate(DestinationTErr : TErrorsManagement);
 end;

var
  TErrManagement    : TErrorsManagement;
  TErrManagementSav : TErrorsManagement;

const
  TemErr_MessagePreRempli             = 0;
  TemErr_CreafAFFFromDBT              = 1;
  TemErr_CalcStockOrigin              = 2;
  TemErr_DeleteCOMPTADIFFEREE         = 3;
  TemErr_UpdateSoldeEcriture          = 4;
  TemErr_UpdateANouveauxCpta          = 5;
  TemErr_DeleteECRITURE               = 6;
  TemErr_InsertExtourneEcriture       = 7;
  TemErr_InsertEcritureTiersPayeur    = 8;
  TemErr_DeleteLIGNEFORMULEPrec       = 9;
  TemErr_DeleteAFORMULEVARQTEPrec     = 10;
  TemErr_DeletePIECEPrec              = 11;
  TemErr_DeleteLIGNEPrec              = 11;
  TemErr_DeletePIEDECHEPrec           = 12;
  TemErr_DeleteLIGNENOMENPrec         = 13;
  TemErr_DeleteLIGNELOTPrec           = 14;
  TemErr_DeleteLIGNESERIEPrec         = 15;
  TemErr_DeleteACOMPTESPrec           = 16;
  TemErr_DeletePIEDPORTPrec           = 17;
  TemErr_DeleteLIGNETARIFPrec         = 18;
  TemErr_DeleteLIENSOLEPrec           = 19;
  TemErr_DeleteLIGNEOUVPrec           = 20;
  TemErr_RecalculStock                = 21;
  TemErr_CalcNumSouche                = 22;
  TemErr_CalcNumPreNumero             = 23;
  TemErr_UpdateERefGescom             = 24;
  TemErr_InsertPieceAnnulation        = 26;
  TemErr_InsertLIGNCOMPL              = 27;
  TemErr_InsertLIGNEOUV               = 28;
  TemErr_InsertLIENSOLE               = 29;
  TemErr_InsertPIECEADRESSE           = 30;
  TemErr_InsertDEMANDEPRIX            = 31;
  TemErr_UpdateDEMANDEPRIXGL          = 32;
  TemErr_RecalcDemandeDePrix          = 33;
  TemErr_UpdatePIECEPrec              = 34;
  TemErr_UpdateSatusPIECEPrec         = 35;
  TemErr_UpdateQteLIGNEPrec           = 36;
  TemErr_UpdateACTIVITE               = 37;
  TemErr_InsertLIGNELOT               = 38;
  TemErr_UpdateTIERS                  = 39;
  TemErr_CptaJournalVide              = 40;
  TemErr_CptaNatureVide               = 41;
  TemErr_CptaNumeroVide               = 42;
  TemErr_CptaCollectifVide            = 43;
  TemErr_CptaParamCptRGVide           = 47;
  TemErr_CptaCptRGInexistant          = 48;
  TemErr_OuvrageNonTrouve             = 50;
  TemErr_CptaCpteRGNonTrouve          = 51;
  TemErr_CptaCpteGLNonTrouve          = 52;
  TemErr_CptaCptTaxeNonTrouve         = 53;
  TemErr_CptaCptRemiseNonTrouve       = 54;
  TemErr_CptaCptEscompteNonTrouve     = 55;
  TemErr_CptaCptRGNonTrouve           = 56;
  TemErr_CptaCptRDNonTrouve           = 57;
  TemErr_CptaCptSTPaieDirectNonTrouve = 58;
  TemErr_CptaCptRestAcptNonTrouve     = 59;
  TemErr_CptaErrorInsertTobEcr        = 60;
  TemErr_CptaErrorInsertTobEcrDiff    = 61;
  TemErr_CptaSJournalVide             = 62;
  TemErr_CptaSNumeroVide              = 63;
  TemErr_CptaSCptEscompteNonTrouve    = 64;
  TemErr_CptaSErrorInsertTobEcr       = 65;
  TemErr_CptaSErrorInsertTobEcrV      = 66;
  TemErr_CptaSUpdateSoldeEcriture     = 67;
  TemErr_UpdatePIEDCOLLECTIF          = 68;
  TemErr_UpdateLIGNESERIE             = 69;
  TemErr_UpdateLIGNEFORMULE           = 70;
  TemErr_UpdateAFORMULEVARQTE         = 71;
  TemErr_UpdateACOMPTES               = 72;
  TemErr_UpdatePIEDPORT               = 73;
  TemErr_UpdatePIECETRAIT             = 74;
  TemErr_UpdatePIECEINTERV            = 75;
  TemErr_UpdatePIECERG                = 76;
  TemErr_UpdatePIEDBASERG             = 77;
  TemErr_UpdateTIMBRESPIECE           = 78;
  TemErr_UpdateLIGNENOMEN             = 79;
  TemErr_RecalcPIECE                  = 80;
  TemErr_UpdateBDETETUDE              = 81;
  TemErr_UpdateBSTATUSANALMOIS        = 82;
  TemErr_InsertExtourneEcritureRgt    = 83;


implementation

uses
  HMsgBox
  , SysUtils
  , CommonTools
  ;

{ TUtilErrorsManagement }
class procedure TUtilErrorsManagement.ShowGenericMessage(pReinitAfter : boolean=True);
begin
  if (Assigned(TErrManagement)) then
  begin
    TErrManagement.ShowError;
    if pReinitAfter then
      TErrManagement.Reinit;
  end;
end;

class function TUtilErrorsManagement.GetGenericMessage(pReinitAfter : boolean=True) : string;
begin
  if Assigned(TErrManagement) and (TErrManagement.ErrorNumber > -1) then
  begin
    Result := TErrManagement.GetMessage;
    if pReinitAfter then
      TErrManagement.Reinit;
  end;
end;

class procedure TUtilErrorsManagement.SetGenericMessage(TemErrNumber : integer; pErrorMsg : string=''; pCriticality : T_EMTypMsg=temtm_Error; pReinitAfter : boolean=True);
begin
  if assigned(TErrManagement) then
  begin
    TErrManagement.Criticality := pCriticality;
    TErrManagement.ReinitAfter := pReinitAfter;
    TErrManagement.ErrorMsg    := pErrorMsg;
    TErrManagement.ErrorNumber := TemErrNumber;
  end;
end;

class function TUtilErrorsManagement.GetStartMsg(TypeStartMsg : T_EMStartMsg) : string;
begin

// IL FAUT APPLELER CA AU LIEU DES FONCTIONS CI-DESSOUS

  case TypeStartMsg of
    temsm_Delete   : Result := 'Erreur lors de la suppression';
    temsm_Update   : Result := 'Erreur lors de la mise à jour';
    temsm_Inser    : Result := 'Erreur lors de la création';
    temsm_Recalc   : Result := 'Erreur lors du recalcul';
    temsm_Number   : Result := 'Erreur lors du calcul du numéro de pièce';
    temsm_Load     : Result := 'Erreur lors du chargement';
    temsm_Cpta     : Result := 'Erreur lors de la comptabilisaton de la pièce';
    temsm_CptaS    : Result := 'Erreur lors de la comptabilisaton des stocks';
    temsm_Generate : Result := 'Erreur lors de la génération';
  else ;
    Result := '';
  end;
end;


class function TUtilErrorsManagement.GetDeleteError : string;
begin
  Result := 'Erreur lors de la suppression';
end;

class function TUtilErrorsManagement.GetUpdateError : string;
begin
  Result := 'Erreur lors de la mise à jour';
end;

class function TUtilErrorsManagement.GetInsertError : string;
begin
  Result := 'Erreur lors de la création';
end;

class function TUtilErrorsManagement.GetRecalcError : string;
begin
  Result := 'Erreur lors du recalcul';
end;

class function TUtilErrorsManagement.GetNumberError : string;
begin
  Result := 'Erreur lors du calcul du numéro de pièce';
end;

class function TUtilErrorsManagement.GetLoadError   : string;
begin
  Result := 'Error lors du chargement';
end;

class function TUtilErrorsManagement.GetCptaError   : string;
begin
  Result := 'Erreur lors de la comptabilisaton de la pièce';
end;

class function TUtilErrorsManagement.GetCptaSError  : string;
begin
  Result := 'Erreur lors de la comptabilisaton des stocks';
end;

class function TUtilErrorsManagement.GetGeneratEror : string;
begin
  Result := 'Erreur lors de la génération';
end;

{ TErrorsManagement }

procedure TErrorsManagement.SetContext(pContext : T_EMContext);
begin
  fContext := pContext;
end;

procedure TErrorsManagement.SetArrOfMsg;
var
  PrevDoc : string;
begin
  PrevDoc := 'de la pièce précédente.';
  ArrOfMsg[TemErr_CreafAFFFromDBT]              := Format('InsertError de l''affaire associée au devis.'                                          , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_CalcStockOrigin]              := Format('%s du stock d''origine.'                                                               , [TUtilErrorsManagement.GetRecalcError]);
  ArrOfMsg[TemErr_DeleteCOMPTADIFFEREE]         := Format('%s des données de la comptabilisation différée.'                                       , [TUtilErrorsManagement.GetDeleteError]);
  ArrOfMsg[TemErr_UpdateSoldeEcriture]          := Format('%s des soldes comptables.'                                                             , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateANouveauxCpta]          := Format('%s des à-nouveaux comptables.'                                                         , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_DeleteECRITURE]               := Format('%s des écritures comptables.'                                                          , [TUtilErrorsManagement.GetDeleteError]);
  ArrOfMsg[TemErr_InsertExtourneEcriture]       := Format('Erreur lors de l''extourne de la pièce comptable.'                                     , []);
  ArrOfMsg[TemErr_InsertEcritureTiersPayeur]    := Format('%s des écritures du tiers payeur.'                                                     , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_DeleteLIGNEFORMULEPrec]       := Format('%s des formules associées aux lignes %s'                                               , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteAFORMULEVARQTEPrec]     := Format('%s des variables des formules associée à l''affaire %s'                                , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeletePIECEPrec]              := Format('%s %s'                                                                                 , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteLIGNEPrec]              := Format('%s des lignes %s'                                                                      , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeletePIEDECHEPrec]           := Format('%s des échéances %s'                                                                   , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteLIGNENOMENPrec]         := Format('%s des nomenclature %s'                                                                , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteLIGNELOTPrec]           := Format('%s des lots %s'                                                                        , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteLIGNESERIEPrec]         := Format('%s des numéros de séries %s'                                                           , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteACOMPTESPrec]           := Format('%s des acomptes et règlements %s'                                                      , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeletePIEDPORTPrec]           := Format('%s des ports et frais %s'                                                              , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteLIGNETARIFPrec]         := Format('%s des tarifs %s'                                                                      , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteLIENSOLEPrec]           := Format('%s des blocs-notes %s'                                                                 , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_DeleteLIGNEOUVPrec]           := Format('%s des ouvrages %s'                                                                    , [TUtilErrorsManagement.GetDeleteError, PrevDoc]);
  ArrOfMsg[TemErr_RecalculStock]                := Format('%s du stock de la pièce.'                                                              , [TUtilErrorsManagement.GetRecalcError]);
  ArrOfMsg[TemErr_CalcNumSouche]                := Format('%s - Le code de la souche est vide.'                                                   , [TUtilErrorsManagement.GetNumberError]);
  ArrOfMsg[TemErr_CalcNumPreNumero]             := Format('%s - Le pré-numéro est vide.'                                                          , [TUtilErrorsManagement.GetNumberError]);
  ArrOfMsg[TemErr_UpdateERefGescom]             := Format('%s de la référence de la pièce sur l''écriture comptable.'                             , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_InsertPieceAnnulation]        := Format('%s de la pièce d''annulation.'                                                         , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_InsertLIGNCOMPL]              := Format('%s des informations complémentaires des lignes.'                                       , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_InsertLIGNEOUV]               := Format('%s des ouvrages.'                                                                      , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_InsertLIENSOLE]               := Format('%s des blocs-notes.'                                                                   , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_InsertPIECEADRESSE]           := Format('%s des adresses.'                                                                      , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_InsertDEMANDEPRIX]            := Format('%s des demandes de prix.'                                                              , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_UpdateDEMANDEPRIXGL]          := Format('%s des données de demandes de prix dans les lignes.'                                   , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_RecalcDemandeDePrix]          := Format('%s des demandes de prix.'                                                              , [TUtilErrorsManagement.GetRecalcError]);
  ArrOfMsg[TemErr_UpdatePIECEPrec]              := Format('%s de la pièce précédente.'                                                            , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateSatusPIECEPrec]         := Format('%s de la Date de modification et de l''Etat de la pièce précédente.'                   , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateQteLIGNEPrec]           := Format('%s des quantités d''une ligne de la pièce précédente.'                                 , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateACTIVITE]               := Format('%s de l''activité.'                                                                    , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_InsertLIGNELOT]               := Format('%s des lots.'                                                                          , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_UpdateTIERS]                  := Format('%s du tiers.'                                                                          , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_CptaJournalVide]              := Format('%s - Le journal associé à la pièce est vide.'                                          , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaNatureVide]               := Format('%s - La nature comptable associée à la pièce est vide.'                                , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaNumeroVide]               := Format('%s - Erreur lors de la récupération du numéro de la pièce comptable.'                  , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCollectifVide]            := Format('%s - Le compte collectif associé au tiers est inexistant.'                             , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMSg[TemErr_CptaParamCptRGVide]           := Format('%s - Le compte de retenue de garantie est vide.'                                       , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMSg[TemErr_CptaCptRGInexistant]          := Format('%s - Le compte de retenue de garantie paramétré est inexistant.'                       , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_OuvrageNonTrouve]             := Format('%s des ouvrages.'                                                                      , [TUtilErrorsManagement.GetLoadError]);
  ArrOfMsg[TemErr_CptaCpteRGNonTrouve]          := Format('%s - Le compte associé à la ligne de retenue de garantie est inexistant.'              , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCpteGLNonTrouve]          := Format('%s - Le compte associé à la ligne de pièce est inexistant.'                            , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCptTaxeNonTrouve]         := Format('%s - Le compte de taxe est inexistant.'                                                , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCptRemiseNonTrouve]       := Format('%s - Le compte de remise est inexistant.'                                              , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCptEscompteNonTrouve]     := Format('%s - Le compte d''escompte est inexistant.'                                            , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCptRGNonTrouve]           := Format('%s - Le compte de retenue de garantie est inexistant.'                                 , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCptRDNonTrouve]           := Format('%s - Le compte de retenue divers est inexistant.'                                      , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCptSTPaieDirectNonTrouve] := Format('%s - Le compte général associé aus sous-traitants en paiment direct est inexistant.'   , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaCptRestAcptNonTrouve]     := Format('%s - Le compte général de restitution d''acompte est inexistant.'                      , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaErrorInsertTobEcr]        := Format('%s - Erreur lors de la création des écritures comptables.'                             , [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaErrorInsertTobEcrDiff]    := Format('%s - Erreur lors de la création des écritures comptables en comptabilisation différée.', [TUtilErrorsManagement.GetCptaError]);
  ArrOfMsg[TemErr_CptaSJournalVide]             := Format('%s - Le journal est vide.'                                                             , [TUtilErrorsManagement.GetCptaSError]);
  ArrOfMsg[TemErr_CptaSNumeroVide]              := Format('%s - Erreur lors de la récupération du numéro de la pièce comptable.'                  , [TUtilErrorsManagement.GetCptaSError]);
  ArrOfMsg[TemErr_CptaSCptEscompteNonTrouve]    := Format('%s - Le compte de stock est inexistant..'                                              , [TUtilErrorsManagement.GetCptaSError]);
  ArrOfMsg[TemErr_CptaSErrorInsertTobEcr]       := Format('%s - %s  des écritures comptables.'                                                    , [TUtilErrorsManagement.GetCptaSError, TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_CptaSErrorInsertTobEcrV]      := Format('%s - %s des écritures comptables des variations de stock.'                             , [TUtilErrorsManagement.GetCptaSError, TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_CptaSUpdateSoldeEcriture]     := Format('%s - %s des soldes comptables.'                                                        , [TUtilErrorsManagement.GetCptaSError, TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdatePIEDCOLLECTIF]          := Format('%s de la répartition des collectifs sur la pièce (PIEDCOLLECTIF).'                     , [TUtilErrorsManagement.GetInsertError]);
  ArrOfMsg[TemErr_UpdateLIGNESERIE]             := Format('%s des numéros de série (LIGNESERIE).'                                                 , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateLIGNEFORMULE]           := Format('%s des formules (LIGNEFORMULE).'                                                       , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateAFORMULEVARQTE]         := Format('%s des variables des formules (AFORMULEVARQTE).'                                       , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateACOMPTES]               := Format('%s des acomptes et règlements (ACOMPTES).'                                             , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdatePIEDPORT]               := Format('%s des ports et frais (PIEDPORT).'                                                     , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdatePIECETRAIT]             := Format('%s des éléments de co-traitance/sous-traitance (PIECETRAIT).'                          , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdatePIECEINTERV]            := Format('%s des intervenants de la sous-traitance (PIECEINTERV).'                               , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdatePIECERG]                := Format('%s des retenues de garanties (PIECERG).'                                               , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdatePIEDBASERG]             := Format('%s des taxes associées aux retenues de garanties (PIEDBASERG).'                        , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateTIMBRESPIECE]           := Format('%s des droits de timbres (TIMBRESPIECE).'                                              , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_UpdateLIGNENOMEN]             := Format('%s du détail des nomenclatures (LIGNENOMEN).'                                          , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_RecalcPIECE]                  := Format('%s de la pièce (PIECE).'                                                               , [TUtilErrorsManagement.GetRecalcError]);
  ArrOfMsg[TemErr_UpdateBDETETUDE]              := Format('%s des en-tête de borderau (BDETETUDE).'                                               , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMSg[TemErr_UpdateBSTATUSANALMOIS]        := Format('%s des cumuls mensuels de consommations (BSTATUSANALMOIS).'                            , [TUtilErrorsManagement.GetUpdateError]);
  ArrOfMsg[TemErr_InsertExtourneEcritureRgt]    := Format('Erreur lors de l''extourne du règlement.'                                              , []);
end;

function TErrorsManagement.GetContextCaption : string;
begin
  case Context of
    temc_DocValid : Result := 'Validation de pièce.';
  else
    Result := '';
  end;
end;

function TErrorsManagement.MessageAlreadyCalculate : boolean;
begin
  Result := (fErrorNumber = TemErr_MessagePreRempli);
end;

constructor TErrorsManagement.Create;
begin
  SetContext(pContext);
  ReInit;
  SetArrOfMsg;
end;

destructor TErrorsManagement.Destroy;
begin
  inherited;
end;

procedure TErrorsManagement.Init(pContext : T_EMContext);
begin
  SetContext(pContext);
  ReInit;
  SetArrOfMsg;
end;

function TErrorsManagement.GetMessage : string;
begin
  Result := Tools.iif(MessageAlreadyCalculate, fErrorMsg, ArrOfMsg[fErrorNumber]);
end;

procedure TErrorsManagement.ShowError;
begin
  if fErrorNumber > -1  then
  begin
    case fCriticality of
      temtm_Information : PGIInfo (GetMessage, GetContextCaption);
      temtm_Warning     : PGIBox  (GetMessage, GetContextCaption);
      temtm_Error       : PGIError(GetMessage, GetContextCaption);
    end;
    Reinit;
  end;
end;

procedure TErrorsManagement.WriteErrorInLog;
begin
  if fErrorNumber > -1  then
  begin
  end;
end;

procedure TErrorsManagement.ReInit;
begin
  fErrorNumber := -1;
  fErrorMsg    := '';
  fCriticality := temtm_none;
  fReinitAfter := True;
end;

procedure TErrorsManagement.Duplicate(DestinationTErr: TErrorsManagement);
begin
  DestinationTErr.SetContext(fContext);
  DestinationTErr.SetArrOfMsg;
  DestinationTErr.fErrorNumber := fErrorNumber;
  DestinationTErr.fErrorMsg    := fErrorMsg;
  DestinationTErr.fCriticality := fCriticality;
  DestinationTErr.fReinitAfter := fReinitAfter;
end;

end.
