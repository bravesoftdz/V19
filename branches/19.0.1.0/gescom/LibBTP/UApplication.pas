unit UApplication;

interface
uses
  AglInit,
	hctrls,
{$IFDEF EAGLCLIENT}
     MenuOLX,Maineagl,
{$ELSE}
     MenuOLG,Fe_Main,
{$ENDIF}
	LicUtil,
  sysutils,
  HEnt1,
	windows,
  inifiles,
  CBPPath,
  Classes,
  forms
  ;

const INIFILE = 'GAMME-PRODUIT.LSE';
      INISEP = 'SEPARATION.LSE';
      PROFILE = 'PROTECTXXX.LSE';
var
  GCCodesSeria : array[1..12] of string;
  GCTitresSeria : array[1..12] of Hstring;
	GCCodeDomaine : string;
  NumRefObligatoire : string = '998.ZZZZI';
  VersionInterne : Boolean;

procedure GetInfoApplication;
procedure GetInfoMajApplication;


implementation

procedure InitPlace;
begin
  V_PGI.LaSerie:=S5 ;
  GCCodeDomaine := '00086010'; //
  TitreHalley:='LSE BTP eXtend 19' ;
  GCCodesSeria [1] :='00083190'; GCTitresSeria [1] := 'Etudes & Situations';
  GCCodesSeria [2] :='00311190'; GCTitresSeria [2] := 'Récupération des Appels d''offres';
  GCCodesSeria [3] :='00312190'; GCTitresSeria [3] := 'Achats/Stocks';
  GCCodesSeria [4] :='00313190'; GCTitresSeria [4] := 'Gestion des chantiers';
  GCCodesSeria [5] :='00314190'; GCTitresSeria [5] := 'Relations Clients';
  GCCodesSeria [6] :='00316190'; GCTitresSeria [6] := 'Contrats';
  GCCodesSeria [7] :='00416190'; GCTitresSeria [7] := 'Interventions';
  GCCodesSeria [8] :='14900190'; GCTitresSeria [8] := 'Cotraitance';
  GCCodesSeria [9] :='14902190'; GCTitresSeria [9] := 'Sous-traitance';    //
  GCCodesSeria [10] :='14905190'; GCTitresSeria [10] := 'Commerce de détail';    //
  GCCodesSeria [11] :='14906190'; GCTitresSeria [11] := 'Plan de charge- Planning';    //
  GCCodesSeria [12] :='14907190'; GCTitresSeria [12] := 'Parc matériel';    //
end;

procedure SetinitDefaut;
begin
  // définition par défaut
  InitPlace;
  HalSocIni:='cegidpgi.ini' ;
  VersionInterne := false;
end;

procedure GetInfoApplication;
var ExeRepert : string;
		Iexist : Integer;
    IniApplication : Tinifile;
    SeparationCEGIDLSE : Boolean;
    HalSocSpe : string;
begin
  SeparationCEGIDLSE := false;
  //
  SetInitDefaut;

  Exerepert := ExtractFilePath (Application.ExeName); // repertoire de l'application

  //
  Iexist := FileOpen(IncludeTrailingBackslash(Exerepert)+INISEP,fmOpenRead );
  if Iexist > 0 then SeparationCEGIDLSE := True;
  FileClose(Iexist);

  if SeparationCEGIDLSE then
  begin
    IniApplication := TiniFile.create (IncludeTrailingBackslash (Exerepert)+INISEP);
    HalSocSpe := IniApplication.ReadString('Application','CEGIDPGIINI','');
    if HalSocSpe <> '' then HalSocIni := HalSocSpe;
  	IniApplication.Free;
  end;

  Iexist := FileOpen(IncludeTrailingBackslash(Exerepert)+INIFILE,fmOpenRead );
  if Iexist < 0 then Exit;
  FileClose(Iexist);

  IniApplication := TiniFile.create (IncludeTrailingBackslash (Exerepert)+INIFILE);
  NomHalley := IniApplication.ReadString('Application','Nom','');
  VersionInterne := IniApplication.ReadString('Application','KIKONEST','')='CNOUS';
  if FileExists(IncludeTrailingBackslash (Exerepert)+PROFILE) then versionInterne := true;
  if NomHalley = 'CBTPS5' then
  begin
    // PLACE
    InitPlace;
    //
  end else if NomHalley = 'CBTPS3' then
  begin
    // SUITE
		V_PGI.LaSerie:=S3 ;
    GCCodeDomaine := '00056010';
    TitreHalley:='LSE BTP eXtend 19' ;
		GCCodesSeria [1] :='00055190'; GCTitresSeria [1] := 'Etudes & Situations';
    GCCodesSeria [2] :='00084190'; GCTitresSeria [2] := 'Récupération des Appels d''offres';
    GCCodesSeria [3] :='00094190'; GCTitresSeria [3] := 'Achats/Stocks';
    GCCodesSeria [4] :='00210190'; GCTitresSeria [4] := 'Gestion des chantiers';
    GCCodesSeria [5] :='00254190'; GCTitresSeria [5] := 'Relations Clients';
    GCCodesSeria [6] :='00255190'; GCTitresSeria [6] := 'Contrats';
    GCCodesSeria [7] :='00256190'; GCTitresSeria [7] := 'Interventions';
    GCCodesSeria [8] :='14901190'; GCTitresSeria [8] := 'Cotraitance';
    GCCodesSeria [9] :='14903190'; GCTitresSeria [9] := 'Sous-traitance';    //
  	GCCodesSeria [10] :='14904190'; GCTitresSeria [10] := 'Commerce de détail';    //
  	GCCodesSeria [11] :='14906190'; GCTitresSeria [11] := 'Plan de charge - Planning';    //
  	GCCodesSeria [12] :='14907190'; GCTitresSeria [12] := 'Parc matériel';    //
  end else if NomHalley = 'CBTPS1' then
  begin
    (*
    // START
		V_PGI.LaSerie:=S3 ;
    GCCodeDomaine := '00595010';
    TitreHalley:='L.S.E Business BTP LINE' ;
		GCCodesSeria [1] :='00055120'; GCTitresSeria [1] := 'Devis - Factures';
    GCCodesSeria [2] :='00084120'; GCTitresSeria [2] := 'Récupération des Appels d''offres';
    GCCodesSeria [3] :='00094120'; GCTitresSeria [3] := 'Achats/Stocks';
    GCCodesSeria [4] :='00210120'; GCTitresSeria [4] := 'Gestion des chantiers';
    GCCodesSeria [5] :='00254120'; GCTitresSeria [5] := 'Relations Clients';
    GCCodesSeria [6] :='00255120'; GCTitresSeria [6] := 'Contrats';
    GCCodesSeria [7] :='00256120'; GCTitresSeria [7] := 'Interventions';
    GCCodesSeria [8] :='14901120'; GCTitresSeria [8] := 'Cotraitance';
    GCCodesSeria [9] :='14903120'; GCTitresSeria [9] := 'Sous-traitance';    //
  	GCCodesSeria [10] :='14904120'; GCTitresSeria [10] := 'Commerce de détail';    //
  	GCCodesSeria [11] :='14906120'; GCTitresSeria [11] := 'Plan de charge - Planning';    //
  	GCCodesSeria [12] :='14907120'; GCTitresSeria [12] := 'Parc matériel';    //
    *)
  end;
  IniApplication.Free;
end;

procedure GetInfoMajApplication;
var ExeRepert : string;
		Iexist : Integer;
    IniApplication : Tinifile;
    SeparationCEGIDLSE : Boolean;
    HalSocSpe : string;
begin
  SeparationCEGIDLSE := false;
  //
  HalSocIni:='cegidpgi.ini' ;
  VersionInterne := false;

  Exerepert := ExtractFilePath (Application.ExeName); // repertoire de l'application

  //
  Iexist := FileOpen(IncludeTrailingBackslash(Exerepert)+INISEP,fmOpenRead );
  if Iexist > 0 then SeparationCEGIDLSE := True;
  FileClose(Iexist);

  if SeparationCEGIDLSE then
  begin
    IniApplication := TiniFile.create (IncludeTrailingBackslash (Exerepert)+INISEP);
    HalSocSpe := IniApplication.ReadString('Application','CEGIDPGIINI','');
    if HalSocSpe <> '' then HalSocIni := HalSocSpe;
  	IniApplication.Free;
  end;
end;


end.
