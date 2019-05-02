{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/05/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTBIMEXPORT ()
Mots clefs ... : TOF;BTBIMEXPORT
*****************************************************************}
Unit BTBIMEXPORT_TOF ;

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
     HTB97,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     Paramsoc, 
     UTOF,
     UBimetre ; 

Type
  TOF_BTBIMEXPORT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure LanceExportBIM (Sender : TObject);

  end ;

Implementation

procedure TOF_BTBIMEXPORT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMEXPORT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMEXPORT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMEXPORT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMEXPORT.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolbarButton97 (GetControl('BLANCEEXPORT')).OnClick := LanceExportBIM;
  SetControlText('EMPLEXPORT',GetParamSocSecur('SO_BTEMPLBIMEXPORT',''));
end ;

procedure TOF_BTBIMEXPORT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMEXPORT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMEXPORT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTBIMEXPORT.LanceExportBIM (Sender : TObject);
var TheFile : string;
    Year,Month,Day : Word;
    TheDate : TDateTime;
begin
  if GetControlText('EMPLEXPORT')='' then
  begin
    PGIInfo('Vous devez renseigner un emplacement de destination');
    Exit;
  end;
  if PgiAsk ('Attention : Vous allez lancer un traitement qui peut être long. Etes-vous sur(e) de voulloir lancer l''export BIM')<>mryes then exit;  
  //
  TheDate := Now;
  DecodeDate (TheDate,Year,Month,Day);
  TheFile := Format('%s-%.04d%.02d%.02d.bimetre',['BIBLSE',Year,Month,Day]);
  //
  ExportBimetre (TheDate,IncludeTrailingBackslash(GetControlText('EMPLEXPORT')),TheFile,
                THCheckbox(GetControl('CBEXPORTOUV')).Checked,
                THCheckbox(GetControl('CBEXPORTMAR')).Checked,
                THCheckbox(GetControl('CBEXPORTPREST')).Checked,
                THmultiValComboBox(getcontrol('FAMOUV')).value, THmultiValComboBox(getcontrol('FAMART')).value,THSpinEdit(GEtControl('NIVFAMMAX')).Value);
end;

Initialization
  registerclasses ( [ TOF_BTBIMEXPORT ] ) ; 
end.
