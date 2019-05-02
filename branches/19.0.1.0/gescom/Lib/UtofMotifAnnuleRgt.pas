Unit UtofMotifAnnuleRgt;

Interface

Uses
  UTOF,
  HCtrls,
  db, {$IFNDEF DBXPRESS} dbTables {$ELSE} uDbxDataSet {$ENDIF}
  ;

Type
  TOF_BTMOTIFANNULERGT = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    function GetNomUser (User : string) : string; 
  end ;

function BLanceFiche_MotifAnnuleRgt(Nat, Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

uses
  Classes
  , sysutils
  , HMsgBox
  , Vierge
  , utilPGI
  , HEnt1
  , CbpDates
  , CommonTools
  {$IFNDEF EAGLCLIENT}
  , FE_main
  {$ELSE EAGLCLIENT}
  , MaineAGL
  {$ENDIF EAGLCLIENT}
  ;

function BLanceFiche_MotifAnnuleRgt(Nat, Cod : String ; Range,Lequel,Argument : string) : string;
begin
  if (Nat <> '') and (Cod <> '') then
    Result := AGLLanceFiche(Nat,Cod,Range,Lequel,Argument)
  else
    Result := '';
end;

procedure TOF_BTMOTIFANNULERGT.OnUpdate ;
var
  Date  : string;
  Motif : string;

  procedure ShowMsg(FieldName : string);
  var
    Msg : string;
  begin
    case Tools.CaseFromString(FieldName, ['DATE', 'MOTIF']) of
      {DATE}  0 : Msg := 'La date';
      {MOTIF} 1 : Msg := 'Le motif';
    end;
    PGIError(Format('%s est obligatoire', [Msg]), Ecran.Caption);
    SetFocusControl(FieldName);
  end;

begin
  Inherited ;
  Date  := GetControlText('DATE');
  Motif := GetControlText('MOTIF');
  if ((Date = '') or (StrToDate(Date) = iDate1900)) or (Motif = '') then
  begin
    if (Date = '') or (StrToDate(Date) = iDate1900) then
      ShowMsg('DATE')
    else if Motif = '' then
      ShowMsg('MOTIF');
    TFVierge(Ecran).ModalResult := 0;
    Exit;
  end else
    TFVierge(Ecran).Retour := Format('%s;%s', [Date, Motif]);
end ;

procedure TOF_BTMOTIFANNULERGT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMOTIFANNULERGT.OnArgument (S : String ) ;
var
  sDate : string;
begin
  Inherited ;
  sDate := GetArgumentString(S, 'DATE');
  if (sDate = '') or (StrToDate(sDate) = iDate1900) then
    sDate := DateToStr(Now);
  SetControlText('DATE' , sDate);
  SetControlText('MOTIF', GetArgumentString(S, 'MOTIF'));
  SetControlText('USER', GetArgumentString(S, 'USER'));
  SetControlText('NOMUSER', GetNomUser(GetControlText('USER')));
end ;

procedure TOF_BTMOTIFANNULERGT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMOTIFANNULERGT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMOTIFANNULERGT.OnCancel () ;
begin
  Inherited ;
  TFVierge(Ecran).Retour := '';
end ;

function TOF_BTMOTIFANNULERGT.GetNomUser(User: string): string;
var QQ : TQuery;
begin
  Result := '?????';
  QQ := OpenSql ('SELECT US_LIBELLE FROM UTILISAT WHERE US_UTILISATEUR="'+User+'"',True,1,'',true);
  if not QQ.eof then
  begin
    Result := QQ.fields[0].AsString;
  end;
  Ferme(QQ);
end;

Initialization
  registerclasses ( [ TOF_BTMOTIFANNULERGT ] ) ;
end.