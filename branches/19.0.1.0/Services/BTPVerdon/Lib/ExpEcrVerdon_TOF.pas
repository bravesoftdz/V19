{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMWS ()
Mots clefs ... : TOF;BTPARAMWS
*****************************************************************}
Unit ExpEcrVerdon_TOF ;

Interface

Uses
  StdCtrls
  , Controls
  , Classes
  {$IFNDEF EAGLCLIENT}
  , db
  , uDbxDataSet
  , mul
  , FE_main
  , uTob
  {$ELSE EAGLCLIENT}
  , eMul
  , MaineAGL
  {$ENDIF EAGLCLIENT}
  , forms
  , sysutils
  , ComCtrls
  , HCtrls
  , HMsgBox
  , UTOF
  , HTB97
  ;

function BTLanceFicheExpEcrVerdon(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_ExpEcrVerdon = Class (TOF)
  private
    StartDate : THEdit;
    EndDate   : THEdit;
    bValider  : TToolbarButton97;

    procedure LoadDates;
    procedure bValider_OnClick(Sender : Tobject);


  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

uses
  UtilBTPVerdon
  , ParamSoc
  , HEnt1
  , TntStdCtrls;

function BTLanceFicheExpEcrVerdon(Nat, Cod : String ; Range,Lequel,Argument : string) : string;
begin
  if (Nat <> '') and (Cod <> '') then
  begin
    V_PGI.ZoomOle := True;
    Result := AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
    V_PGI.ZoomOle := False;
  end else
    Result := '';
end;

procedure TOF_ExpEcrVerdon.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ExpEcrVerdon.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ExpEcrVerdon.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_ExpEcrVerdon.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_ExpEcrVerdon.OnArgument (S : String ) ;
begin
  Inherited ;
  StartDate        := THEdit(GetControl('STARTDATE'));
  EndDate          := THEdit(GetControl('STARTDATE_'));
  bValider         := TToolbarButton97(GetControl('BValider'));
  bValider.OnClick := bValider_OnClick;
  LoadDates;
end ;

procedure TOF_ExpEcrVerdon.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_ExpEcrVerdon.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_ExpEcrVerdon.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_ExpEcrVerdon.LoadDates;
begin
  StartDate.Text := GetParamSocSecur('SO_EXPXMLDE', DateToStr(iDate1900));
  EndDate.Text   := GetParamSocSecur('SO_EXPXMLA' , DateToStr(iDate2099));
end;

procedure TOF_ExpEcrVerdon.bValider_OnClick(Sender: Tobject);
{$IFNDEF APPSRV}
var
  ExpEcr : TExportEcr;
{$ENDIF APPSRV}
begin
{$IFNDEF APPSRV}
  { Test la cohérence des dates }
  if StrToDateTime(EndDate.Text) < StrToDateTime(StartDate.Text) then
  begin
    PGIError('Erreur dans la période.');
    StartDate.SetFocus;
  end else
  begin
    ExpEcr := TExportEcr.create;
    try                                                                      
      ExpEcr.StartDate   := StrToDateTime(StartDate.Text);
      ExpEcr.EndDate     := StrToDateTime(EndDate.Text);
      ExpEcr.ForceExport := ThCheckBox(GetControl('FORCEEXPORT')).Checked;
      ExpEcr.AccountingExpTreatment;
    finally
      ExpEcr.Free;
      LoadDates;
    end;
  end;
  TForm(Ecran).ModalResult := 0;
{$ENDIF APPSRV}
end;

Initialization
  registerclasses ( [ TOF_ExpEcrVerdon ] ) ;
end.                                                     

