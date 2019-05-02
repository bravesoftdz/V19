{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 08/09/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFECTATION ()
Mots clefs ... : TOF;UTOFAFFECTATION
*****************************************************************}
Unit UTofAffectation;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     Lookup,
     uTob,
     HTB97,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Paramsoc,
     AGLInitGC,
     UTOF ;

Type
  TOF_AFFECTATION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private

    Responsable : THEdit;
    Ressource   : THEdit;

    Ok_GereSAV  : Boolean;

    procedure AfficheResponsable(TobRessource: tob);
    procedure LectureResponsable;

    procedure AppelResponsable(sender: Tobject);


  end ;

Implementation

procedure TOF_AFFECTATION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnUpdate ;
begin
  Inherited ;

  if Ressource.text <> '' then
  Begin
    LaTob.PutValue('RETOUR','X');
	  LaTob.PutValue('RESSOURCE',Ressource.text);
    exit;
  end;

end ;

procedure TOF_AFFECTATION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnArgument (S : String ) ;
begin
  Inherited ;

  Responsable := THEdit(GetControl('AFF_RESPONSABLE'));
  Responsable.OnElipsisClick := AppelResponsable;

  Ressource   := THEdit(GetControl('ARS_RESSOURCE'));

  //FV1 : 25/06/2018 - FS#3136 - SCETEC - ne proposer que les intervenants qui sont cochés en SAV en affectation d'un appel
  Ok_GereSAV := GetParamSocSecur('SO_AFFINTSAV', False);

  Ressource.text := '';

end ;

procedure TOF_AFFECTATION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.AppelResponsable(sender: Tobject);
Var StWhere     : String;
    StArgument  : string;
begin

  StWhere    := '';
  StArgument := '';

  if Trim(Ressource.Text) <> '' then
  Begin
    StArgument := 'ARS_RESSOURCE=' + Trim(Ressource.text);
    //FV1 : 25/06/2018 - FS#3136 - SCETEC - ne proposer que les intervenants qui sont cochés en SAV en affectation d'un appel
    If Ok_GereSAV then StArgument := StArgument + ';GERESAV=X';
  end
  else
  begin
    //FV1 : 25/06/2018 - FS#3136 - SCETEC - ne proposer que les intervenants qui sont cochés en SAV en affectation d'un appel
    If Ok_GereSAV then StArgument := 'GERESAV=X';
  end;

  //FV1 : 20/11/2015 - FS#1732 - ESPACS : Ne pas afficher les ressources fermées dans fiche Appel
  StWhere := ' AND (ARS_TYPERESSOURCE="SAL" OR ARS_TYPERESSOURCE="ST" OR ARS_TYPERESSOURCE="INT") AND ARS_FERME="-" ';

  DispatchRecherche(Ressource, 3, stwhere, StArgument, '');

  if Ressource.Text = '' then
  begin
    Responsable.text  := '';
    Ressource.text    := '';
  end
  else
    LectureResponsable; //Lecture des Ressources sélectionné et affichage des informations

end;

procedure TOF_AFFECTATION.LectureResponsable;
Var TobRessource : TOB;
    Req          : String;
Begin

  Req := '';

  //FV1 : 25/06/2018 - FS#3136 - SCETEC - ne proposer que les intervenants qui sont cochés en SAV en affectation d'un appel
  Req := 'SELECT * FROM RESSOURCE LEFT JOIN BRESSOURCE ON ARS_RESSOURCE=BRS_RESSOURCE ';
  Req := Req + 'WHERE ARS_RESSOURCE ="' + Ressource.Text + '"';

  if Ok_GereSAV then
  begin
    Req := Req + '  AND BRS_GERESAV ="X"';
  end;

  TobRessource := Tob.Create('LesRessources',Nil, -1);
  TobRessource.LoadDetailDBFromSQL('RESSOURCE',req,false);

  if TobRessource.Detail.Count <> 0 then AfficheResponsable(TobRessource);

  TobRessource.free;

end;

procedure TOF_AFFECTATION.AfficheResponsable(TobRessource: tob);
Var Nom : String;
begin

	Nom := TobRessource.Detail[0].GetValue('ARS_LIBELLE') + ' ' + TobRessource.Detail[0].GetValue('ARS_LIBELLE2');

	Responsable.text := Nom;

end;

Initialization
  registerclasses ( [ TOF_AFFECTATION ] ) ;
end.

