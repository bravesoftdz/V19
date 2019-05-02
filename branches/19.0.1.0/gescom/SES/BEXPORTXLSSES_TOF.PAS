{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/11/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BEXPORTXLSSES ()
Mots clefs ... : TOF;BEXPORTXLSSES
*****************************************************************}
Unit BEXPORTXLSSES_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul,
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
     Windows,
     Messages,
     UTOF ; 
const
  StdFileDATASES = 'C:\PGI00\STD\DATAS_SES.xlsx';

Type
  TOF_BEXPORTXLSSES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    CodeAffaire : string;
    fWinExcel   : OleVariant;
    WorkBook    : Variant;
    procedure XLSNAMEClick (Sender : Tobject);
    procedure AlimChantierClick (Sender  :TObject);
    function  AlimenteFicheChantier : boolean;
    function  OuvreExcel : boolean;
    procedure FermeExcel;
    procedure RemplitDatas (WB : variant;CodeAffaire : string);
    function  ChargeAffaire(CodeAffaire: String; TOBAFFAIRE : TOB) : Boolean;
    procedure ChargeTiers(CodeTiers : string;TOBTIER,TOBADRESSES : TOB);
    procedure ChargeLesDevis(TOBAffaire,TOBDEVIS,TOBSOUSTRAITS : TOB);
    procedure ChargeAffaireAdresses (TOBAFFAIRE,TOBADRESSES : TOB);
    procedure RAZTableauDEVIS(WB: Variant);
    procedure RAZTableauSOUSTRAIT(WB: Variant);

  end ;

Implementation

uses CalcOLEGenericBTP,Dialogs,utilxlsBTP,Variants,TiersUtil,FactAdresse,SHellApi;

procedure TOF_BEXPORTXLSSES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BEXPORTXLSSES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BEXPORTXLSSES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BEXPORTXLSSES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BEXPORTXLSSES.OnArgument (S : String ) ;
var critere,LibChantier,Champ,Valeur : string;
    X : integer;
begin
  Inherited ;
  Critere := (Trim (ReadTokenSt (S) ) ) ;
  CodeAffaire := '';
  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos ('=', Critere) ;
      if x <> 0 then
      begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        if champ = 'CHANTIER' then
        begin
          CodeAffaire := Valeur;
        end else if champ = 'LIBCHANTIER' then
        begin
          LibChantier := Valeur;
        end;
      end;
    end;
    Critere := (Trim (ReadTokenSt (S) ) ) ;
  end;
  THLabel(GetControl('LDESCRIPTION')).Caption := 'Chantier : '+BTPCodeAffaireAffiche(CodeAffaire,' ') + ' - '+LibCHantier;
  //
  THEdit(GetControl('XLSNAME')).OnElipsisClick := XLSNAMEClick;
  TToolbarButton97 (GetControl('BALIMFICCHANTIER')).onclick := AlimChantierClick;
end ;

procedure TOF_BEXPORTXLSSES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BEXPORTXLSSES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BEXPORTXLSSES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BEXPORTXLSSES.FermeExcel;
begin
  if not VarIsEmpty(fWinExcel) then fWinExcel.Quit;
  fWinExcel := unassigned;
end;

procedure TOF_BEXPORTXLSSES.XLSNAMEClick(Sender: Tobject);
var TT : TOpenDialog;
begin
	TT := TOpenDialog.Create(self.Ecran);
  TRY
    TT.DefaultExt := '.xlsx';
    TT.Filter := 'Modèle de document word (*.xlsx)|*.xlsx';
    if TT.Execute then
    begin
      THEdit(GetControl('XLSNAME')).Text := TT.FileName;
    end;
  FINALLY
  	TT.Free;
  end;
end;

procedure TOF_BEXPORTXLSSES.AlimChantierClick(Sender: TObject);
var fileName : string;
begin
  FileName := THEdit(GetControl('XLSNAME')).Text;
  if not FileExists(THEdit(GetControl('XLSNAME')).Text) then
  begin
    PGIInfo('La fiche Chantier n''est pas défini ou n''existe pas');
    exit;
  end;
  if not FileExists(StdFileDATASES) then
  begin
    PGIInfo('Le Fichier DATA SES.xlsx n''est pas présent dans C:\PGI00\APP');
    exit;
  end;
  if AlimenteFicheChantier then
  begin
    if PgiAsk('La fiche chantier est prête.Désirez-vous l''ouvrir ?') = mrYes then
    begin
      ShellExecute(0, PCHAR('open'), PChar(FileName), nil, nil, SW_SHOWNORMAL);
    end;
  end;
  SendMessage(Ecran.Handle,WM_CLOSE,0,0);
end;

function TOF_BEXPORTXLSSES.AlimenteFicheChantier : boolean;
var WB2,TS : Variant;
begin
  Result := false;

  if not OfficeExcelDispo then
  begin
    PGIInfo('Vous devez installer office sur ce poste, pour réaliser cette opération');
    exit;
  end;
  if Not OuvreExcel then
  begin
    PGIInfo('Ouverture d''Excel impossible');
    exit;
  end;
  TRY
    WorkBook := OpenWorkBook (THEdit(GetControl('XLSNAME')).Text ,fWinExcel);
    // Modified by f.vautrain 04/01/2019 16:39:27 - FS#3430 - SES ETANCHEITE - Erreur en intégration des fiches affaire si fait plusieurs fois
    if not TrouveSheet (WorkBook,'DONNEES_LSE') Then
    begin
      WB2 := OpenWorkBook (StdFileDATASES,fWinExcel);
      TS := SelectSheet (WB2,'DONNEES_LSE');
      TS.copy(WorkBook.sheets[2]);
      WB2.Close(0);
      TS := UnAssigned;
    end
    else
    begin
      TS := SelectSheet(WorkBook,'DONNEES_LSE');
      RAZTableauDEVIS(WorkBook);
      RAZTableauSOUSTRAIT(WorkBook);
      TS := UnAssigned;
    end;
    // Modified by f.vautrain 04/01/2019 16:39:27 - FS#3430 - SES ETANCHEITE - Erreur en intégration des fiches affaire si fait plusieurs fois
    TS := SelectSheet(WorkBook,'DONNEES_LSE');
    RemplitDatas (WorkBook,CodeAffaire);
    WorkBook.save;
    Result := True;
  FINALLY
    FermeExcel;
  END;
end;

function TOF_BEXPORTXLSSES.OuvreExcel : boolean;
var WinNew : Boolean;
begin
  Result := OpenExcel (true,fWinExcel,WinNew);
end;

// Modified by f.vautrain 04/01/2019 16:39:27 - FS#3430 - SES ETANCHEITE - Erreur en intégration des fiches affaire si fait plusieurs fois
Procedure TOF_BEXPORTXLSSES.RAZTableauDEVIS(WB : Variant);
Var Ind : Integer;
Begin

  For Ind := 1 To 24 do
  begin
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,1);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,2);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,3);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,4);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,5);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,6);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,7);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,8);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,9);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,10);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,11);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,12);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,13);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,14);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,15);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,16);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,17);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,18);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,19);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,20);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,21);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,22);
    ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS','',Ind,23);
  end;

end;

// Modified by f.vautrain 04/01/2019 16:39:27 - FS#3430 - SES ETANCHEITE - Erreur en intégration des fiches affaire si fait plusieurs fois
Procedure TOF_BEXPORTXLSSES.RAZTableauSOUSTRAIT(WB : Variant);
Var Ind : Integer;
Begin

  For Ind := 1 To 33 do
  begin
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,1);
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,2);
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,3);
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,4);
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,5);
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,6);
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,7);
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,8);
    ExcelRangeValue(WB, 'DONNEES_LSE','SOUSTRAIT','',Ind,9);
  end;

end;


procedure TOF_BEXPORTXLSSES.RemplitDatas(WB: variant; CodeAffaire: string);
var TOBAFFAIRE,TOBTIER,TOBADRESSES,TOBDEVIS,TOBSOUSTRAITS : TOB;

  procedure AlimenteFeuille;
  var Texte : string;
      II : Integer;
      TOBD : TOB;
  begin
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_AFFAIRE',TOBAFFAIRE.GetString('AFF_AFFAIRE'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBELLE',TOBAFFAIRE.GetString('AFF_LIBELLE'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_REFEXTERNE',TOBAFFAIRE.GetString('AFF_REFEXTERNE'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_NOMRESPONSABLE',TOBAFFAIRE.GetString('AFF_NOMRESPONSABLE'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_NOMAPPORTEUR',TOBAFFAIRE.GetString('AFF_NOMAPPORTEUR'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_MAILPPORTEUR',TOBAFFAIRE.GetString('AFF_MAILAPPORTEUR'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE1',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE1'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE2',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE2'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE3',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE3'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE4',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE4'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE5',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE5'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE6',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE6'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE7',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE7'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE8',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE8'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIRE9',TOBAFFAIRE.GetString('AFF_LIBREAFFAIRE9'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREAFFAIREA',TOBAFFAIRE.GetString('AFF_LIBREAFFAIREA'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREDATE1',TOBAFFAIRE.GetString('AFF_LIBREDATE1'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREDATE2',TOBAFFAIRE.GetString('AFF_LIBREDATE2'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_LIBREDATE3',TOBAFFAIRE.GetString('AFF_LIBREDATE3'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_CHARLIBRE1',TOBAFFAIRE.GetString('AFF_CHARLIBRE1'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_CHARLIBRE2',TOBAFFAIRE.GetString('AFF_CHARLIBRE2'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_CHARLIBRE3',TOBAFFAIRE.GetString('AFF_CHARLIBRE3'),1,1);
    ExcelRangeFormat(WB,'DONNEES_LSE','AFF_VALLIBRE1','# ##0,00 €',1,1);
    ExcelRangeValue(WB, 'DONNEES_LSE','AFF_VALLIBRE1',StringReplace(TOBAFFAIRE.GetString('AFF_VALLIBRE1'),',','.',[rfreplaceAll]),1,1);
    ExcelRangeFormat(WB,'DONNEES_LSE','AFF_VALLIBRE2','# ##0,00 €',1,1);
    ExcelRangeValue(WB, 'DONNEES_LSE','AFF_VALLIBRE2',StringReplace(TOBAFFAIRE.GetString('AFF_VALLIBRE2'),',','.',[rfreplaceAll]),1,1);
    ExcelRangeFormat(WB,'DONNEES_LSE','AFF_VALLIBRE3','# ##0,00 €',1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_VALLIBRE3',StringReplace(TOBAFFAIRE.GetString('AFF_VALLIBRE3'),',','.',[rfreplaceAll]),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_BOOLLIBRE1',TOBAFFAIRE.GetString('AFF_BOOLLIBRE1'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_BOOLLIBRE2',TOBAFFAIRE.GetString('AFF_BOOLLIBRE2'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','AFF_BOOLLIBRE3',TOBAFFAIRE.GetString('AFF_BOOLLIBRE3'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','T_TIERS',TOBTIER.GetString('T_TIERS'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','T_LIBELLE',TOBTIER.GetString('T_LIBELLE'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','T_ADRESSE',TOBTIER.GetString('T_ADRESSE1'),1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','T_ADRESSE',TOBTIER.GetString('T_ADRESSE2'),2,1);
    ExcelRangeValue(WB,'DONNEES_LSE','T_ADRESSE',TOBTIER.GetString('T_ADRESSE3'),3,1);
    ExcelRangeValue(WB,'DONNEES_LSE','T_ADRESSE',TOBTIER.GetString('T_CODEPOSTAL'),4,1);
    ExcelRangeValue(WB,'DONNEES_LSE','T_ADRESSE',TOBTIER.GetString('T_VILLE'),5,1);
    // Adresse de facturation
    if TOBADRESSES.detail.Count > 1 then
    begin
      if TOBADRESSES.detail[1].GetString('ADR_JURIDIQUE') <> '' then Texte := TOBADRESSES.detail[1].GetString('ADR_JURIDIQUE')+' '+TOBADRESSES.detail[1].GetString('ADR_LIBELLE')
                                                                else Texte := TOBADRESSES.detail[1].GetString('ADR_LIBELLE');
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',Texte,1,1);
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',TOBADRESSES.detail[1].GetString('ADR_ADRESSE1'),2,1);
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',TOBADRESSES.detail[1].GetString('ADR_ADRESSE2'),3,1);
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',TOBADRESSES.detail[1].GetString('ADR_ADRESSE3'),4,1);
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',TOBADRESSES.detail[1].GetString('ADR_CODEPOSTAL'),5,1);
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',TOBADRESSES.detail[1].GetString('ADR_VILLE'),6,1);
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',TOBADRESSES.detail[1].GetString('NOMCONTACT'),7,1);
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',TOBADRESSES.detail[1].GetString('TELCONTACT'),8,1);
      ExcelRangeValue(WB,'DONNEES_LSE','ADR_FACTURATION',TOBADRESSES.detail[1].GetString('MAILCONTACT'),9,1);
    end;
    // Adresse de Livraison
    if TOBADRESSES.detail[0].GetString('ADR_JURIDIQUE') <> '' then Texte := TOBADRESSES.detail[0].GetString('ADR_JURIDIQUE')+' '+TOBADRESSES.detail[0].GetString('ADR_LIBELLE')
                                                              else Texte := TOBADRESSES.detail[0].GetString('ADR_LIBELLE');
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',Texte,1,1);
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',TOBADRESSES.detail[0].GetString('ADR_ADRESSE1'),2,1);
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',TOBADRESSES.detail[0].GetString('ADR_ADRESSE2'),3,1);
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',TOBADRESSES.detail[0].GetString('ADR_ADRESSE3'),4,1);
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',TOBADRESSES.detail[0].GetString('ADR_CODEPOSTAL'),5,1);
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',TOBADRESSES.detail[0].GetString('ADR_VILLE'),6,1);
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',TOBADRESSES.detail[0].GetString('NOMCONTACT'),7,1);
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',TOBADRESSES.detail[0].GetString('TELCONTACT'),8,1);
    ExcelRangeValue(WB,'DONNEES_LSE','ADR_LIVRAISON',TOBADRESSES.detail[0].GetString('MAILCONTACT'),9,1);
    //Grille des devis
    for II := 0 to TOBDEVIS.Detail.count -1 do
    begin
      TOBD := TOBDEVIS.detail[II];
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('GP_NUMERO'),II+1,1);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('ETATAFFAIRE'),II+1,2);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('GP_REFINTERNE'),II+1,3);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('GP_REFEXTERNE'),II+1,4);
      ExcelRangeFormat(WB,'DONNEES_LSE','DEVIS','# ##0,00 €',II+1,5);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',StringReplace(TOBD.GetString('GP_TOTALHTDEV'),',','.',[rfreplaceAll]),II+1,5);
      ExcelRangeFormat(WB,'DONNEES_LSE','DEVIS','# ##0,00 €',II+1,6);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',StringReplace(TOBD.GetString('GP_TOTALTTCDEV'),',','.',[rfreplaceAll]),II+1,6);
      ExcelRangeFormat(WB,'DONNEES_LSE','DEVIS','# ##0,00 €',II+1,7);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',StringReplace(TOBD.GetString('MONTANTTVA'),',','.',[rfreplaceAll]),II+1,7);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('REPRESENTANT'),II+1,8);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS1'),II+1,9);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS2'),II+1,10);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS3'),II+1,11);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS4'),II+1,12);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS5'),II+1,13);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS6'),II+1,14);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS7'),II+1,15);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS8'),II+1,16);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERS9'),II+1,17);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBRETIERSA'),II+1,18);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBREPIECE1'),II+1,19);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBREPIECE2'),II+1,20);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('LIBREPIECE3'),II+1,21);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('GP_DATELIBREPIECE1'),II+1,22);
      ExcelRangeValue(WB, 'DONNEES_LSE','DEVIS',TOBD.GetString('GP_DATELIBREPIECE2'),II+1,23);
    end;
    //Grille des sous-traitants
    for II := 0 to TOBSOUSTRAITS.Detail.count -1 do
    begin
      TOBD := TOBSOUSTRAITS.detail[II];
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',TOBD.GetString('T_LIBELLE'),II+1,1);
      ExcelRangeFormat(WB,'DONNEES_LSE','SOUSTRAIT','@',II+1,8);
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',TOBD.GetString('T_TELEPHONE'),II+1,2);
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',TOBD.GetString('T_EMAIL'),II+1,3);
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',TOBD.GetString('T_ADRESSE1'),II+1,4);
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',TOBD.GetString('T_ADRESSE2'),II+1,5);
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',TOBD.GetString('T_CODEPOSTAL'),II+1,6);
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',TOBD.GetString('T_VILLE'),II+1,7);
      ExcelRangeFormat(WB,'DONNEES_LSE','SOUSTRAIT','@',II+1,8);
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',TOBD.GetString('T_SIRET'),II+1,8);
      ExcelRangeFormat(WB,'DONNEES_LSE','SOUSTRAIT','# ##0,00 €',II+1,9);
      ExcelRangeValue(WB,'DONNEES_LSE','SOUSTRAIT',StringReplace(FloatToStr(TOBD.GetDouble('MONTANT')),',','.',[rfreplaceAll]),II+1,9);
    end;
  end;

  procedure DefiniLesInfosContactAdr(TOBADRESSES,TOBTIERS : TOB);
  var II : Integer;
      TOBADR : TOB;
      QQ : TQuery;
  begin
    for II := 0 to TOBADRESSES.Detail.count -1 do
    begin
      TOBADR := TOBADRESSES.detail[II];
      TOBADR.AddChampSupValeur('NOMCONTACT','');
      TOBADR.AddChampSupValeur('TELCONTACT','');
      TOBADR.AddChampSupValeur('MAILCONTACT','');
      if TOBADR.GetInteger('ADR_NUMEROCONTACT')<> 0 then
      begin
        QQ := OpenSql ('SELECT C_NOM,C_TELEPHONE,C_RVA FROM CONTACT WHERE C_TYPECONTACT="T" AND C_AUXILIAIRE="'+TOBTIERS.GetString('T_AUXILIAIRE')+'" AND C_NUMEROCONTACT='+TOBADR.getString('ADR_NUMEROCONTACT'),True,1,'',True);
        if not QQ.eof then
        begin
          TOBADR.SetString('NOMCONTACT',QQ.fields[0].AsString);
          TOBADR.SetString('TELCONTACT',QQ.fields[1].AsString);
          TOBADR.SetString('MAILCONTACT',QQ.fields[2].AsString);
        end;
        Ferme(QQ);
      end;
    end;
  end;


begin
  TOBAFFAIRE := TOB.create ('AFFAIRE',nil,-1);
  TOBTIER := TOB.Create ('TIERS',nil,-1);
  TOBADRESSES := TOB.Create ('LES ADRESSES',nil,-1);
  TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
  TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
  TOBDEVIS := TOB.Create ('LES DEVIS',nil,-1);
  TOBSOUSTRAITS := TOB.Create ('LES SOUS TRAITS',nil,-1);
  TRY
    if not ChargeAffaire(CodeAffaire,TOBAFFAIRE) then Exit;
    ChargeTiers(TOBAFFAIRE.GetString('AFF_TIERS'),TOBTIER,TOBADRESSES);
    ChargeAffaireAdresses (TOBAFFAIRE,TOBADRESSES);
    DefiniLesInfosContactAdr(TOBADRESSES,TOBTIER);
    ChargeLesDevis(TOBAffaire,TOBDEVIS,TOBSOUSTRAITS);
    AlimenteFeuille;
  FINALLY
    TOBAFFAIRE.free;
    TOBTIER.free;
    TOBADRESSES.free;
    TOBDEVIS.free;
    TOBSOUSTRAITS.free;
  END;
end;


function TOF_BEXPORTXLSSES.ChargeAffaire(CodeAffaire: String;TOBAFFAIRE: TOB): Boolean;
var QQ : TQuery;
    SQL : String;
begin
  Result := false;
  SQL := 'SELECT AFF_TIERS,AFF_AFFAIRE,AFF_LIBELLE,AFF_REFEXTERNE,'+
         '(SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE=AFF_RESPONSABLE) AS AFF_NOMRESPONSABLE,'+
         '(SELECT GCL_LIBELLE FROM COMMERCIAL WHERE GCL_TYPECOMMERCIAL="APP" AND GCL_COMMERCIAL=AFF_APPORTEUR) AS AFF_NOMAPPORTEUR,'+
         '(SELECT GCL_EMAIL FROM COMMERCIAL WHERE GCL_TYPECOMMERCIAL="APP" AND GCL_COMMERCIAL=AFF_APPORTEUR) AS AFF_MAILAPPORTEUR,'+
         'IIF (AFF_LIBREAFF1<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF1" AND YX_CODE=AFF_LIBREAFF1),"") AS AFF_LIBREAFFAIRE1,'+
         'IIF (AFF_LIBREAFF2<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF2" AND YX_CODE=AFF_LIBREAFF2),"") AS AFF_LIBREAFFAIRE2,'+
         'IIF (AFF_LIBREAFF3<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF3" AND YX_CODE=AFF_LIBREAFF3),"") AS AFF_LIBREAFFAIRE3,'+
         'IIF (AFF_LIBREAFF4<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF4" AND YX_CODE=AFF_LIBREAFF4),"") AS AFF_LIBREAFFAIRE4,'+
         'IIF (AFF_LIBREAFF5<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF5" AND YX_CODE=AFF_LIBREAFF5),"") AS AFF_LIBREAFFAIRE5,'+
         'IIF (AFF_LIBREAFF6<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF6" AND YX_CODE=AFF_LIBREAFF6),"") AS AFF_LIBREAFFAIRE6,'+
         'IIF (AFF_LIBREAFF7<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF7" AND YX_CODE=AFF_LIBREAFF7),"") AS AFF_LIBREAFFAIRE7,'+
         'IIF (AFF_LIBREAFF8<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF8" AND YX_CODE=AFF_LIBREAFF8),"") AS AFF_LIBREAFFAIRE8,'+
         'IIF (AFF_LIBREAFF9<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LF9" AND YX_CODE=AFF_LIBREAFF9),"") AS AFF_LIBREAFFAIRE9,'+
         'IIF (AFF_LIBREAFFA<>"",(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LFA" AND YX_CODE=AFF_LIBREAFFA),"") AS AFF_LIBREAFFAIREA,'+
         'IIF (AFF_DATELIBRE1="01/01/1900","",AFF_DATELIBRE1) AS AFF_LIBREDATE1,'+
         'IIF (AFF_DATELIBRE2="01/01/1900","",AFF_DATELIBRE2) AS AFF_LIBREDATE2,'+
         'IIF (AFF_DATELIBRE3="01/01/1900","",AFF_DATELIBRE3) AS AFF_LIBREDATE3,'+
         'AFF_CHARLIBRE1,AFF_CHARLIBRE2,AFF_CHARLIBRE3,'+
         'AFF_VALLIBRE1,AFF_VALLIBRE2,AFF_VALLIBRE3,'+
         'AFF_BOOLLIBRE1,AFF_BOOLLIBRE2,AFF_BOOLLIBRE3 '+
         'FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeAffaire+'"';
  //
  QQ := OpenSQL(SQL,True,1,'',true);
  if not QQ.eof then
  begin
    Result := True;
    TOBAFFAIRE.SelectDB('',QQ);
  end;
  Ferme(QQ);
end;

procedure TOF_BEXPORTXLSSES.ChargeLesDevis(TOBAffaire, TOBDEVIS,TOBSOUSTRAITS: TOB);
var QQ : TQuery;
    SQL : String;
    TOBINT : TOB;
begin
  TOBINT := TOB.Create ('LES INFOS S/T',nil,-1);
  TRY
    SQL := 'SELECT GP_NATUREPIECEG,GP_NUMERO,'+
           '(SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="AET" AND CC_CODE=(SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE=GP_AFFAIREDEVIS)) AS ETATAFFAIRE,'+
           'GP_REFINTERNE,GP_REFEXTERNE,GP_TOTALHTDEV,GP_TOTALTTCDEV,(GP_TOTALTTCDEV-GP_TOTALHTDEV) AS MONTANTTVA,'+
           '(SELECT GCL_LIBELLE FROM COMMERCIAL WHERE GCL_COMMERCIAL=GP_REPRESENTANT) AS REPRESENTANT,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT1" AND YX_CODE=GP_LIBRETIERS1) AS LIBRETIERS1,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT2" AND YX_CODE=GP_LIBRETIERS2) AS LIBRETIERS2,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT3" AND YX_CODE=GP_LIBRETIERS3) AS LIBRETIERS3,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT4" AND YX_CODE=GP_LIBRETIERS4) AS LIBRETIERS4,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT5" AND YX_CODE=GP_LIBRETIERS5) AS LIBRETIERS5,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT6" AND YX_CODE=GP_LIBRETIERS6) AS LIBRETIERS6,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT7" AND YX_CODE=GP_LIBRETIERS7) AS LIBRETIERS7,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT8" AND YX_CODE=GP_LIBRETIERS8) AS LIBRETIERS8,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LT9" AND YX_CODE=GP_LIBRETIERS9) AS LIBRETIERS9,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="LTA" AND YX_CODE=GP_LIBRETIERSA) AS LIBRETIERSA,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="CA1" AND YX_CODE=GP_LIBREPIECE1) AS LIBREPIECE1,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="CA2" AND YX_CODE=GP_LIBREPIECE2) AS LIBREPIECE2,'+
           '(SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="CA3" AND YX_CODE=GP_LIBREPIECE3) AS LIBREPIECE3,'+
           'GP_DATELIBREPIECE1,GP_DATELIBREPIECE2,GP_DATELIBREPIECE3 '+
           'FROM PIECE WHERE GP_NATUREPIECEG="DBT" AND GP_AFFAIRE="'+TOBAffaire.getString('AFF_AFFAIRE')+'"';
    QQ := OpenSQL(SQL,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBDEVIS.LoadDetailDB('PIECE','','',QQ,false);
    end;
    Ferme(QQ);
    //
    SQL := 'SELECT BPE_FOURNISSEUR,T_LIBELLE,T_TELEPHONE,T_EMAIL,T_ADRESSE1,T_ADRESSE2,T_CODEPOSTAL,T_VILLE,T_SIRET, SUM(BPE_TOTALHTDEV) AS MONTANT '+
           'FROM PIECETRAIT '+
           'LEFT JOIN TIERS ON BPE_FOURNISSEUR=T_TIERS AND T_NATUREAUXI="FOU" '+
           'LEFT JOIN PIECE ON GP_NATUREPIECEG=BPE_NATUREPIECEG AND GP_SOUCHE=BPE_SOUCHE AND GP_NUMERO=BPE_NUMERO AND GP_INDICEG=BPE_INDICEG '+
           'WHERE '+
           'BPE_NATUREPIECEG="DBT" AND '+
           'BPE_TYPEINTERV="Y00" AND '+
           'GP_NATUREPIECEG="DBT" AND GP_AFFAIRE="'+TOBAffaire.getString('AFF_AFFAIRE')+'" '+
           'GROUP BY BPE_FOURNISSEUR,T_LIBELLE,T_TELEPHONE,T_EMAIL,T_ADRESSE1,T_ADRESSE2,T_CODEPOSTAL,T_VILLE,T_SIRET '+
           'ORDER BY BPE_FOURNISSEUR ';
    QQ := OpenSQL(SQL,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBSOUSTRAITS.LoadDetailDB('PIECETRAIT','','',QQ,false);
    end;
    Ferme(QQ);
  FINALLY
    TOBINT.free;
  END;
end;

procedure TOF_BEXPORTXLSSES.ChargeTiers(CodeTiers: string; TOBTIER,TOBADRESSES: TOB);
var QQ: TQuery;
    iNadresseLivr,iNadresseFact : Integer;
    AuxiLivr,CodeLivr,AuxiFact,CodeFact : string;
    TobAdresse : TOB;
begin
  iNadresseLivr := 0;
  iNadresseFact := 0;
  //
  QQ := OpenSQL('SELECT * FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS WHERE T_TIERS="'+CodeTIers+'" AND T_NATUREAUXI="CLI"',True,1,'',true);
  if not QQ.eof then
  begin
    TOBTIER.SelectDB('',QQ);

    if TOBTIER.FieldExists('YTC_NADRESSELIV') then
    begin
      iNadresseLivr := TOBTIER.GetInteger('YTC_NADRESSELIV');
      iNadresseFact := TOBTIER.GetInteger('YTC_NADRESSEFAC');
    end;

    if (TOBTIER.GetString('YTC_TIERSLIVRE') <> '') then
    begin
      AuxiLivr := TOBTIER.GetString('YTC_TIERSLIVRE');
      CodeLivr := TiersAuxiliaire(AuxiLivr, True);
    end
    else
    begin
      AuxiLivr := TOBTIER.GetString('T_AUXILIAIRE');
      CodeLivr := TOBTIER.GetString('T_TIERS');
    end;
    //
    if CodeLivr = '' then Exit;
    //
    AuxiFact := TOBTIER.GetSTring('T_FACTURE');
    if ((AuxiFact = '') or (AuxiFact = AuxiLivr)) then
    begin
      AuxiFact := AuxiLivr;
      CodeFact := CodeLivr;
    end else
    begin
      CodeFact := TiersAuxiliaire(AuxiFact, True);
    end;
    if (iNAdresseLivr = 0) then
    begin
      if Auxilivr = TOBTier.GetValue('T_AUXILIAIRE') then
        GetAdrFromTOB(TOBAdresses.Detail[0], TOBTier,true)
      else
        GetAdrFromCode(TOBAdresses.Detail[0], CodeLivr,True);
    end else
    begin
      TobAdresse := Tob.create('_ADRESSES_', nil, -1);
      try
        GetTobAdresseFromAdresses(TobAdresse, 'TIE', CodeLivr, iNAdresseLivr, 'Livraison');
        if (TobAdresses.detail.count > 0) and (TobAdresse.detail.count > 0) then GetTobPieceAdresseFromTobAdresses(TobAdresses.Detail[0], TobAdresse.Detail[0],true);
      finally
        TobAdresse.Free;
      end;
    end;

    // Facturation
    if (CodeFact = CodeLivr) and (iNAdresseLivr = 0) and (iNAdresseFact = 0) then
    begin
      if TOBADRESSEs.Detail.count > 1 then GetAdrFromTOB(TOBAdresses.Detail[1], TOBTier,true);
    end else
    begin
      if (iNAdresseFact <> 0) then
      begin
        TobAdresse := Tob.create('_ADRESSES_', nil, -1);
        try
          GetTobAdresseFromAdresses(TobAdresse, 'TIE', CodeFact, iNAdresseFact, 'Facturation');
          if (TobAdresses.detail.count > 0) and (TobAdresse.detail.count > 0) then GetTobPieceAdresseFromTobAdresses(TobAdresses.Detail[1], TobAdresse.Detail[0],true);
        finally
          TobAdresse.Free;
        end;
      end
      else
      begin
        if TOBADRESSEs.Detail.count > 0 then
        begin
          if CodeFact = TOBTier.GetValue('T_AUXILIAIRE') then
            GetAdrFromTOB(TOBAdresses.Detail[1], TOBTier,True)
          else
            GetAdrFromCode(TOBAdresses.Detail[1], CodeFact,true);
        end;
      end;
    end;
  end;
  Ferme(QQ);

end;

procedure TOF_BEXPORTXLSSES.ChargeAffaireAdresses(TOBAFFAIRE,TOBADRESSES: TOB);
var Q: TQuery;
    TOBAdrAff,TOBAdrLiv,TOBAdr: TOB;
begin
  Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_TYPEADRESSE="AFA" AND ADR_REFCODE="' +
                TobAffaire.GetValue('AFF_AFFAIRE') + '"', True,-1, '', True);
  TOBAdrAff := TOB.Create('ADRESSES', nil, -1);
  try
    if not Q.EOF then
    begin
      TOBAdrAff.SelectDB('', Q);
      TOBAdr := TOBAdresses.Detail[1];
      TOBAdr.Dupliquer(TobAdrAff, False, True);
      // reinitialisation de la clé
      TOBAdr.PutValue('ADR_NUMEROADRESSE', 0);
      TOBAdr.PutValue('ADR_REFCODE', '');
      TOBAdr.PutValue('ADR_TYPEADRESSE', '');
      TOBAdr.PutValue('ADR_NUMADRESSEREF', 0);
    end;
  finally
    Ferme(Q);
    TobAdrAff.Free;
  end;
  //
  Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_TYPEADRESSE="INT" AND ADR_REFCODE="' +
                TOBAFFAIRE.GetValue('AFF_AFFAIRE') + '"', True,-1, '', True);
  TOBAdrLiv := TOB.Create('LES ADRESSES', nil, -1);
  try
    if not Q.EOF then
    begin
      TOBAdrLiv.LoadDetailDb('ADRESSES','','', Q,false);
      TOBAdr := TOBAdresses.Detail[0];
      TOBAdr.Dupliquer(TOBAdrLiv.detail[0], False, True);
      // reinitialisation de la clé
      TOBAdr.PutValue('ADR_NUMEROADRESSE', 0);
      TOBAdr.PutValue('ADR_REFCODE', '');
      TOBAdr.PutValue('ADR_TYPEADRESSE', '');
      TOBAdr.PutValue('ADR_NUMADRESSEREF', 0);
    end;
  finally
    Ferme(Q);
    TOBAdrLiv.Free;
  end;
  //
end;

Initialization
  registerclasses ( [ TOF_BEXPORTXLSSES ] ) ;
end.
