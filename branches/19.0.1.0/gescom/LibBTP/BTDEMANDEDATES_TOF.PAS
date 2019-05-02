{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/09/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDEMANDEDATES ()
Mots clefs ... : TOF;BTDEMANDEDATES
*****************************************************************}
Unit BTDEMANDEDATES_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
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
     AGLInit,
     UTOF ;

Type
  TOF_BTDEMANDEDATES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  	private
    	TOBData : TOB;
      DateFac : THEdit;
      TDateFac: THLabel;
  end ;

function DemandeDateLivraison (TOBpiece : TOB) : Boolean;


Implementation
uses saisUtil;

function DemandeDateLivraison (TOBpiece : TOB) : Boolean;
var TobDates  : TOB;
begin
  TOBDates := TOB.Create ('LES DATES', nil,-1);
  TOBDates.AddChampSupValeur('RETOUROK','-');
  TOBDates.AddChampSupValeur('DATFAC', TOBpiece.GetDateTime('GP_DATELIVRAISON'));
  TOBDates.AddChampSupValeur('DATESITUATION','-');
  TOBDates.AddChampSupValeur('ANTERIORITE','-');
  TOBDates.AddChampSupValeur('NUMSIT',0);
  TOBDates.AddChampSupValeur('TYPEDATE','Date de livraison');
  TRY
    TheTOB := TOBDates;
    AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
    TheTOB := nil;
    if TOBDates.getValue('RETOUROK')='X' then
    begin
    	TOBpiece.SetDateTime('GP_DATELIVRAISON', TOBDates.GetDateTime('DATFAC'));
    end;
  FINALLY
  	result := (TOBDates.getValue('RETOUROK')='X');
  	freeAndNil(TOBDates);
  END;
end;

procedure TOF_BTDEMANDEDATES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnUpdate ;
begin
  Inherited ;
  if TOBData.getValue('ANTERIORITE')='X' then
  begin
    if THSpinEdit(GEtControl('NUMSIT')).Value =0 then
    begin
      PgiInfo ('Merci de renseigner le numéro de la dernière situation');
      Ecran.ModalResult := 0;
      Exit;
    end;
  end;
  if Not IsValidDate(DateFac.text) Then
  begin
    PgiInfo ('La date est incorrecte. Merci de la modifier');
    Ecran.ModalResult := 0;
    Exit;
  end;
  if (ControleDate(DateFac.text)<>0) Then
  begin
    PgiInfo ('La date n''est pas dans un exercice ouvert. Merci de la modifier');
    Ecran.ModalResult := 0;
    Exit;
  end;

  TOBData.SetDateTime('DATFAC', StrToDate(DateFac.text));
  TOBData.SetInteger('NUMSIT',  THSpinEdit(GEtControl('NUMSIT')).Value);

  if StrToDate(DateFac.text) = IDate2099 then
    TOBData.putValue('RETOUROK','-')
  else
    TOBData.putValue('RETOUROK','X');
  
end ;

procedure TOF_BTDEMANDEDATES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  DateFac := THEdit(GetControl('DATEFAC'));
  TDateFac:= THLabel(GetControl('TDATEFAC'));

  TOBData := LaTOB;
  //
  if TOBData.getValue('DATFAC') <> iDate1900 then
    DateFac.Text := DateToStr(TOBData.GetValue('DATFAC'))
  else
    DateFac.text := '  /  /    ';
  //
  if TOBData.getValue('DATESITUATION')='X' then
  begin
  	ecran.Caption     := TraduireMemoire ('Date de situation');
    TDateFac.Caption  := 'Date situation';
  end;
  if TOBData.getValue('TYPEDATE')<>'' then
  begin
  	ecran.Caption     := TraduireMemoire (TOBData.GetString('TYPEDATE'));
    TDateFac.Caption  := TOBData.getString('TYPEDATE');
  end;
  if TOBData.getValue('ANTERIORITE')='X' then
  begin
    SetControlVisible('LNUMSIT',true);
    SetControlVisible('NUMSIT',true);
  end;
end ;

procedure TOF_BTDEMANDEDATES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTDEMANDEDATES ] ) ; 
end.
