unit MDispUtil;

// majstruc umajstruct

interface

// ubob

uses Forms, dialogs,Hpanel,
   Classes,UApplication 
  ;

procedure TraitementEchangesBSV;
procedure InitApplication;

type
  Tgctie = class(TDatamodule)
    procedure FMDispADMCreate(Sender: TObject);
  end;

var
  gctie: Tgctie;
  okArg,OkLogin : boolean;
  st : string;
implementation

{$R *.DFM}

uses windows, MenuOLG, sysutils, Messages, Controls,HEnt1, Hctrls,
     uBTPVerrouilleDossier,HMsgBox,Ent1,PGIExec,MajTable,UTraiteTables,DB,
    {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF}
  ;

function prepareEnvUtilisateur : boolean;
var QQ: TQuery;
begin
  Result := false;
  QQ := OpenSql ('SELECT US_EMAIL, US_EMAILPASSWORD, US_EMAILSMTPSERVER, US_EMAILSMTPLOGIN, US_EMAILSMTPPWD, US_SMTPPORT '+
                 'FROM UTILISAT WHERE US_UTILISATEUR="'+V_PGI.User+'"',True,1,'',True);
  if not QQ.eof then
  begin
    V_PGI.EMailLogin := QQ.fields[0].asString;
    V_PGI.EMailPassword := QQ.fields[1].AsString;
    V_PGI.SMTPFrom  := QQ.fields[0].AsString;
    V_PGI.SMTPServer := QQ.fields[2].AsString;
    V_PGI.EMailSmtpLogin := QQ.fields[3].AsString;
    V_PGI.EmailSmtpPassword  := QQ.fields[4].AsString;
    V_PGI.SMTPPort := QQ.fields[5].AsString;
    result := ((V_PGI.EMailLogin<>'') or (V_PGI.EMailSmtpLogin<>'')) and (V_PGI.SMTPServer<>'') and ((V_PGI.EMailPassword<>'') or (V_PGI.EmailSmtpPassword<>''));
  end;
  ferme(QQ);
end;

procedure TraitementEchangesBSV;
var Nom, St, Value,Treatment    : string;
    i                 : integer;
    Connect           : Boolean;
begin
  V_PGI.Halley:=TRUE ;
  V_PGI.Versiondev:=FALSE ;
  VH^.ModeSilencieux := True;
  V_PGI.MultiUserLogin := TRUE;
  Connect := FALSE;
  V_PGI.MailMethod := mmSMTP;
  
  V_PGI.UserLogin  := 'LSE';
  V_PGI.DateEntree := Date();

  for i:=1 to ParamCount do
  begin
    St:=ParamStr(i) ;
    Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
    Value:=UpperCase(Trim(St)) ;
    //Paramètres de connexion
    if Nom='/USER'     then V_PGI.User :=Value;
    if Nom='/PASSWORD' then V_PGI.PassWord:=Value;
    if Nom='/DATE'     then BEGIN  V_PGI.DateEntree:=StrToDate(Value) ; END ;
    if Nom='/BASE'     then BEGIN  V_PGI.CurrentAlias:=Value ; END ;
    if Nom='/TRAIT'    then BEGIN  Treatment:= Value; END ;
  end;

  if (not Connect) then
  begin
    Connect := ConnecteHalley(V_PGI.CurrentAlias,FALSE,@ChargeMagHalley,NIL,NIL,NIL);
  end;

  if connect then
  begin
    if Treatment = 'RECUPBSV' then
    begin
      prepareEnvUtilisateur;
      ConstitueDocsFromDatasBSV(false,true);
    end else if Treatment = 'XXXXX' then
    begin
    end;
  end;

  Application.ProcessMessages;
  if Connect then
  begin
    Logout ;
    DeconnecteHalley ;
  end else
  begin
    // Fermeture de l'application
    Application.ProcessMessages;
    if FMenuG <> nil then
    begin
      FMenuG.ForceClose := True ;
      PostMessage(FmenuG.Handle,WM_CLOSE,0,0) ;
      FMenuG.Close ;
    end;
    VH^.STOPRSP:=TRUE ;
    Exit;
  end;
  VH^.STOPRSP:=TRUE ;
end;

procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
var
  CodeRetour : integer;
  stMessage, Versionbase : string;
  Ok1,Ok2 : Boolean;
begin
  case Num of
    10:
      begin
        if (V_PGI.MajStructAuto) and (V_PGI.RunWithParams) then
        begin
          FMenuG.ForceClose := True;
          PostMessage(FMenuG.Handle, WM_CLOSE, 0, 0);
          Application.ProcessMessages;
          Exit;
        end else
        begin
          if (not ISUtilisable) then
          begin
             stMessage := 'Base de données momentanément indisponible (mise à niveau en cours) ';
             stMessage := stMessage +#13#10 +'Veuillez réessayer ultérieurement';
             PGIInfo(stMessage,'Information de connection');
             FMenuG.ForceClose := True;
             PostMessage(FMenuG.Handle, WM_CLOSE, 0, 0);
             Application.ProcessMessages;
             Exit;
          end;
        end;
      end;

    11: ; //Après deconnection
    12: ;  // permet de faire qques chose avant connexion et seria
    13: ;
    15: ;
    16: ;
    100: ; // executer depuis le lanceur

    324110 : ConstitueDocsFromDatasBSV(true,false);
    
  else HShowMessage('2;?caption?;' + TraduireMemoire('Fonction non disponible : ') + ';W;O;O;O;', TitreHalley, IntToStr(Num));
  end;
end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT, Range: string);
begin
  case Num of
    1: ;
  end;
end;

procedure AfterChangeModule(NumModule: integer);
var
  VireCompta: Boolean;
begin
  V_PGI.VersionDemo := False;
  VireCompta := True;
  V_PGI.LaSerie := S5;
  if ctxPCL in V_PGI.PGIContexte then
  begin
    // maintenant qu'on peut le lancer en direct
  end else
  begin
    // PCS 13/10/2003 modules compta toujours présents.
    VireCompta := False;
  end;
  Case NumModule of
    324 : BEGIN
          end;
  end;
  //
end;


procedure InitApplication;
begin
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
  FMenuG.OnChangeModule := AfterChangeModule;
  FMenuG.SetModules([324], [39]);
  FMenuG.bSeria.Visible := False;
  V_PGI.DispatchTT := DispatchTT;
  VH^.OkModCompta := True;
  VH^.OkModBudget := True;
  VH^.OkModImmo := True;
  VH^.OkModEtebac := True;
end;

procedure Tgctie.FMDispADMCreate(Sender: TObject);
begin
  PGIAppAlone := True;
  CreatePGIApp;
end;

procedure InitialisationVPGI;
begin

  Apalatys := 'LSE';
  NomHalley := 'UTILSPOC';

  V_PGI.NumVersionBase := 998 ;
  TitreHalley := 'Utilitaires BTP-BSV    Base ' + IntToStr(V_PGI.NumVersionBase);
  HalSocIni := 'CEGIDPGI.ini';

  { Ce nom apparaît en bas à gauche de l'application }
  Copyright := '© Copyright L.S.E';

  V_PGI.PGIContexte:=[ctxGescom,ctxAffaire,ctxBTP,ctxGRC];
  V_PGI.StandardSurDp := False;  // indispensable pour l'aiguillage entre le DP et les bases dossiers (gamme Expert), voir SQL025
  V_PGI.OutLook:=TRUE ;
  V_PGI.OfficeMsg:=TRUE ;
  V_PGI.ToolsBarRight:=TRUE ;
  ChargeXuelib ;
  V_PGI.VersionDemo:=True ;
  V_PGI.MenuCourant:=0 ;
  V_PGI.VersionReseau:=True ;
  V_PGI.NumVersion:='19.0';
  V_PGI.NumBuild:='000.0';
  V_PGI.NumVersionBase:=998 ;
  V_PGI.CodeProduit:='032' ;
  V_PGI.DateVersion:=EncodeDate(2019,04,24);
  V_PGI.ImpMatrix := True ;
  V_PGI.OKOuvert:=FALSE ;
  V_PGI.Halley:=TRUE ;
  V_PGI.BlockMAJStruct:=TRUE ;
  V_PGI.PortailWeb:='http://www.lse.fr/' ;
  V_PGI.CegidApalatys:=FALSE ;
  V_PGI.QRMultiThread:=TRUE;
  V_PGI.MenuCourant:=0 ;
  V_PGI.RazForme:=TRUE;
  V_PGI.NumMenuPop:=27 ;
  V_PGI.CegidBureau:= True;
  V_PGI.LookUpLocate:= True;   // utile pour les ellipsis_click sur THEDIT
  V_PGI.QRPDFOptions:=4;        // pour tronquer les libellés en impression
  V_PGI.QRPdf:=True;           // Mode PDF Par défaut pour totalisations situations
end;

initialization
  InitialisationVPGI;

end.

