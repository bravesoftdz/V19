unit UBimetre;
interface

Uses StdCtrls,
     Controls,
     Classes,
     db,
     uDbxDataSet,
     uTob,
     HTB97,
     forms,
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     Paramsoc,
     XMLDoc,xmlintf,
     HRichOLE,
     Ulog,
     SAISUTIL;

type

  TDataTransport = class (TObject)
  public
    DEV : RDevise;
    //
    TOBOuvragesP : TOB;
    TOBPieceTrait : TOB;
    TOBBases       : TOB;
    TOBBasesL      : TOB;
    TOBEches       : TOB;
    TOBPorcs       : TOB;
    TOBTiers       : TOB;
    TOBArticles    : TOB;
    TOBConds       : TOB;
    TOBTarif       : TOB;
    TOBComms       : TOB;
    TOBCatalogu    : TOB;
    TOBAdresses    : TOB;
    TOBCXV         : TOB;
    TOBNomen       : TOB;
    TOBDim         : TOB;
    TOBLOT         : TOB;
    TOBSerie       : TOB;
    TOBSerRel      : TOB;
    TOBAcomptes    : TOB;
    TOBDispoContreM: TOB;
    TOBAffaire     : TOB;
    TOBCPTA        : TOB;
    TOBANAS        : TOB;
    TOBANAP        : TOB;
    TOBGSA         : TOB;
    TOBOuvrage     : TOB;
    TOBLIENOLE     : TOB;
    TOBPieceRG     : TOB;
    TobBasesRG     : TOB;
    TOBLienXls     : TOB;
    TOBPIECE_O     : TOB;
    TOBSSTRAIT     : TOB;
    TOBPIECE       : TOB;
    //
    constructor create (NaturePiece : string);
    destructor destroy; override;
    procedure SetModified;
    procedure MajTobs;
  end;


procedure ExportBimetre (DateTrait : TDateTime; DirOut,FileName : string; ExportOUV,ExportMAR,ExportPre : boolean; FamOuv,FamMAr : String; NivFamMax : integer);
procedure RecupInfosEnteteBIm(Repertoire,NomFic : string; TOBL : TOB);
function DecodeDateBim (TheDateBIM: string) : TDateTime;
function ConstitueDocument(NaturePiece : string; Repertoire : string; TOBL : TOB; RAPPORT: THRichEditOLE) : boolean;

var CodeDefaut : string;
implementation
uses FactTiers,
     FactTOB,
     Ent1,
     EntGC,
     FactPiece,
     FactUtil,
     FactAdresse,
     FactOuvrage,
     FactAdresseBTP,
     FactSpec,
     factcalc,
     FactureBtp,
     BTPUtil,
     FactComm,
     FactArticle,
     BTStructChampSup,
     UtilArticle,
     utilPGI,
     galPatience,
     Variants
     ;

function DecodeDateBim (TheDateBIM: string) : TDateTime;
var TheDatePart,TheTimePart,TheDateReal : string;
    YY,MM,DD,HH,MN,SS : string;
begin
  TheDatePart := '';
  TheTimePart := '';
  Result := IDate1900;
  if Length(TheDateBim) < 10 then Exit;
  TheDatePart := Copy(TheDateBIM,1,10);
  if Length(TheDateBIM) > 12 then
  begin
    TheTimePart := Copy(TheDateBIM,12,Length(TheDateBIM));
  end;
  YY := Copy(TheDatePart,1,4);
  MM := copy(TheDatePart,6,2);
  DD := Copy(TheDatePart,9,2);
  if TheTimePart <> '' then
  begin
    HH := Copy(TheTimePart,1,2);
    MN := Copy(TheTimePart,4,2);
    SS := Copy(TheTimePart,7,2);
  end;
  TheDateReal := Format('%s/%s/%s',[DD,MM,YY]);
  if TheTimePart <> '' then
  begin
    TheDateReal := TheDateReal + Format('%s:%s;%s',[HH,MN,SS]);
  end;
  Result := StrToDateTime(TheDateReal);
end;

function AjouteNoeudFamille ( From : string; NoeudPere : IXMLNode; IDLigPere : Integer; IDLig : integer ; Famille1: string; Famille2: string = ''; Famille3: string='') : IXMLNode;
var QQ : TQuery;
    SQL : string;
    TheFam : string;
    TheSSFam : string;
    Niv,NivF : Integer;
    DD : IXMLNode;
    Libelle : string;
begin
  if Famille3 = '' then
  begin
    if Famille2 = '' then
    begin
      TheFam := Famille1;
      TheSSFam := '';
      Niv := 2;
      NivF := 1;
    end else
    begin
      TheFam := Famille2;
      TheSSFam := Famille1;
      Niv := 3;
      NivF := 2;
    end;
  end else
  begin
    TheFam := Famille3;
    TheSSFam := Famille1+';'+Famille2;
    Niv := 4;
    NivF := 3;
  end;
  //
  SQl := 'SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="'+From+InttoStr(NivF)+'" AND CC_CODE="'+TheFam+'"';// AND CC_LIBRE="'+TheSSFam+'"';
  QQ := OpenSQL(sql,True,1,'',true);
  if not QQ.eof then
  begin
    Libelle:= QQ.Fields[0].AsString;
  end else
  begin
    Libelle := '?????????????';
  end;
  Ferme(QQ);
  Result := NoeudPere.AddChild('LIGNE');
  DD := Result.AddChild('IDLIG'); DD.Text := IntToStr(IDlig);
  DD := Result.AddChild('IDPARENT'); DD.Text := IntToStr(IDLigPere);
  DD := Result.AddChild('TYPELIG'); DD.Text := 'C';
  DD := Result.AddChild('NIVCHAP'); DD.Text := IntToStr(Niv);
  DD := Result.AddChild('CODEARTICLE');
  DD := Result.AddChild('CODE_DANS_BIBLIO');
  DD := Result.AddChild('LIBELLE'); DD.Text := Libelle;
  DD := Result.AddChild('LIBELLECOM'); DD.Text := Libelle ;
  DD := Result.AddChild('LIBELLETEC');
  DD := Result.AddChild('EDITEE'); DD.Text := 'true';
  DD := Result.AddChild('INFO');
end;

function AjouteNoeudOuvrage ( NoeudPere : IXMLNode;IDLigPere,IDLig,LastNiv : integer ; CodeOuvrage,Libelle,BlocNote,Unite : string;  TauxTva : double) : IXMLNode;
var DD : IXMLNode;
begin
  //
  Result := NoeudPere.AddChild('LIGNE');
  DD := Result.AddChild('IDLIG'); DD.Text := IntToStr(IDlig);
  DD := Result.AddChild('IDPARENT'); DD.Text := IntToStr(IDLigPere);
  DD := Result.AddChild('TYPELIG'); DD.Text := 'A';
  DD := Result.AddChild('NIVCHAP'); DD.Text := IntToStr(LastNiv);
  DD := Result.AddChild('CODEARTICLE'); DD.text := CodeOuvrage;
  DD := Result.AddChild('CODE_DANS_BIBLIO');
  DD := Result.AddChild('LIBELLE'); DD.Text := Libelle;
  DD := Result.AddChild('LIBELLECOM'); DD.Text := BlocNote;
  DD := Result.AddChild('LIBELLETEC');
  DD := Result.AddChild('PVHT'); DD.text := '0';
  DD := Result.AddChild('TAUXTVA'); DD.text := STRFPOINT(TauxTva);
  DD := Result.AddChild('UNITE'); DD.text := Unite;
  DD := Result.AddChild('DECUNITE'); DD.text := IntToStr(V_PGI.okdecP);
  DD := Result.AddChild('EDITEE'); DD.Text := 'true';
  DD := Result.AddChild('INFO');
end;

function DecomposeFamille (Familles : string) : string;
var TheFams,Fam : string;
begin
  Result := '';
  TheFams := Familles;
  repeat
    Fam := READTOKENPipe (Familles,';');
    if Fam = '' then break;
    if result <> '' Then Result := Result + ',';
    Result := Result + '"'+Fam+'"';
  until Fam= '';
end;



procedure ExportePrestations (LIGNES : IXMLNode; FamMar : string; IdLigPere : integer; NivFamMAx : Integer; Var Idlig : integer);

  function AddEnteteMateriaux (ROOT : IXMLNode; IDLigPere : Integer; var IdLig : integer) : IXMLNode ;
  var DD : IXMLNode;
  begin
    Result := ROOT.AddChild('LIGNE');
    DD := Result.AddChild('IDLIG'); DD.Text := IntToStr(IDlig);
    DD := Result.AddChild('IDPARENT'); DD.Text := IntToStr(IDLigPere);
    DD := Result.AddChild('TYPELIG'); DD.Text := 'C';
    DD := Result.AddChild('NIVCHAP'); DD.Text := '1';
    DD := Result.AddChild('CODEARTICLE');
    DD := Result.AddChild('CODE_DANS_BIBLIO');
    DD := Result.AddChild('LIBELLE'); DD.Text := 'Prestations';
    DD := Result.AddChild('LIBELLECOM'); DD.Text := 'Prestations';
    DD := Result.AddChild('LIBELLETEC');
    DD := Result.AddChild('EDITEE'); DD.Text := 'true';
    DD := Result.AddChild('INFO');
    inc(IdLig);
  end;
var QQ: TQuery;
    SQl : string;
    FamNiv1,FamNiv2,FamNiv3 : string;
    NFAM1,NFAM2,NFAm3,NOUV,EntOuvrage,TheP: IXMLNode;
    IdLigP,IdLig1,IdLig2,IdLig3,LastNiv,LastP : Integer;
begin
  //
  FamNiv1 := '---';
  FamNiv2 := '---';
  FamNiv3 := '---';
  LastNiv := 1;
  SQL := 'SELECT GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_CODEARTICLE,GA_LIBELLE,GA_BLOCNOTE,GA_PVHT,GA_QUALIFUNITEVTE,'+
         '('+
          'SELECT TV_TAUXVTE '+
          'FROM TXCPTTVA '+
          'WHERE '+
          'TV_TVAOUTPF="TX1" AND TV_REGIME="FRA" AND TV_CODETAUX=GA_FAMILLETAXE1'+
          ') AS TAUXTVA '+
          'FROM ARTICLE WHERE GA_TYPEARTICLE="PRE"';
  if FamMar <> '' then
  begin
    SQL := SQL + ' AND GA_FAMILLENIV1 IN ('+DecomposeFamille(FamMar)+')';
  end;
  SQL := SQL + ' ORDER BY GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_CODEARTICLE';
  QQ := OpenSQL(SQL,True,-1,'',true);
  if not QQ.eof then
  begin
    //
    IdLigP := Idlig;
    LastP := IdLig;
    EntOuvrage := AddEnteteMateriaux (LIGNES,IdLigPere,IdLig);
    //
    QQ.First;
    repeat
      if (QQ.Fields[0].AsString <> '') then
      begin
        if (FamNiv1 <> QQ.Fields[0].AsString) and (QQ.Fields[0].AsString <> '') then
        begin
          IdLig1 := IdLig;
          NFAM1 := AjouteNoeudFamille ('FN',EntOuvrage,IdLigP,IdLig,QQ.Fields[0].AsString);
          Inc(IDLig);
          FamNiv1 := QQ.Fields[0].AsString;
          FamNiv2 := '---';
          FamNiv3 := '---';
          LastNiv := 2;
          LastP := IdLig1;
        end;
      end else
      begin
        NFAM1 := nil;
      end;

      if (QQ.Fields[1].AsString <> '') and (NivFamMax > 1) then
      begin
        if (FamNiv2 <> QQ.Fields[1].AsString) and (QQ.Fields[1].AsString <> '') then
        begin
          IdLig2 := IdLig;
          NFAM2 := AjouteNoeudFamille ('FN',NFAM1,IdLig1,IdLig,QQ.Fields[0].AsString,QQ.Fields[1].AsString);
          Inc(IDLig);
          FamNiv2 := QQ.Fields[1].AsString;
          FamNiv3 := '---';
          LastNiv := 3;
          LastP := IdLig2;
        end;
      end else NFAM2 := nil;

      if (QQ.Fields[2].AsString <> '') and (NivFamMax > 2) then
      begin
        if (FamNiv3 <> QQ.Fields[2].AsString) and (QQ.Fields[2].AsString <> '') then
        begin
          IdLig3 := IdLig;
          NFAM3 := AjouteNoeudFamille ('FN',NFAM2,IdLig2,IdLig,QQ.Fields[0].AsString,QQ.Fields[1].AsString,QQ.Fields[2].AsString);
          Inc(IDLig);
          FamNiv3 := QQ.Fields[2].AsString;
          LastNiv := 4;
          LastP := IdLig3;
        end;
      end else NFAM3 := nil;
      
      if NFAM3 <> nil then TheP := NFAm3
      else if NFAM2 <> nil Then TheP := NFAM2
      else if NFAM1 <> nil then TheP := NFAM1
      else TheP := EntOuvrage;
      NOUV := AjouteNoeudOuvrage ( TheP ,LastP,IdLig,LastNiv + 1, QQ.fields[3].asString,QQ.fields[4].asString,QQ.fields[5].asString,QQ.fields[7].asString,QQ.fields[8].asfloat  );
      inc(IdLig);
      QQ.next;
    until QQ.Eof;
  end;
  ferme (QQ);
end;


procedure ExporteMateriaux (LIGNES : IXMLNode; FamMar : string; IdLigPere : integer; NivFamMax : Integer; Var Idlig : integer);

  function AddEnteteMateriaux (ROOT : IXMLNode; IDLigPere : Integer; var IdLig : integer) : IXMLNode ;
  var DD : IXMLNode;
  begin
    Result := ROOT.AddChild('LIGNE');
    DD := Result.AddChild('IDLIG'); DD.Text := IntToStr(IDlig);
    DD := Result.AddChild('IDPARENT'); DD.Text := IntToStr(IDLigPere);
    DD := Result.AddChild('TYPELIG'); DD.Text := 'C';
    DD := Result.AddChild('NIVCHAP'); DD.Text := '1';
    DD := Result.AddChild('CODEARTICLE');
    DD := Result.AddChild('CODE_DANS_BIBLIO');
    DD := Result.AddChild('LIBELLE'); DD.Text := 'Matériaux';
    DD := Result.AddChild('LIBELLECOM'); DD.Text := 'Matériaux';
    DD := Result.AddChild('LIBELLETEC');
    DD := Result.AddChild('EDITEE'); DD.Text := 'true';
    DD := Result.AddChild('INFO');
    inc(IdLig);
  end;
var QQ: TQuery;
    SQl : string;
    FamNiv1,FamNiv2,FamNiv3 : string;
    NFAM1,NFAM2,NFAm3,NOUV,EntOuvrage,TheP: IXMLNode;
    IdLigP,IdLig1,IdLig2,IdLig3,LastNiv,LastP : Integer;
begin
  //
  FamNiv1 := '---';
  FamNiv2 := '---';
  FamNiv3 := '---';
  LastNiv := 1;
  SQL := 'SELECT GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_CODEARTICLE,GA_LIBELLE,GA_BLOCNOTE,GA_PVHT,GA_QUALIFUNITEVTE,'+
         '('+
          'SELECT TV_TAUXVTE '+
          'FROM TXCPTTVA '+
          'WHERE '+
          'TV_TVAOUTPF="TX1" AND TV_REGIME="FRA" AND TV_CODETAUX=GA_FAMILLETAXE1'+
          ') AS TAUXTVA '+
          'FROM ARTICLE WHERE GA_TYPEARTICLE="MAR"';
  if FamMar <> '' then
  begin
    SQL := SQL + ' AND GA_FAMILLENIV1 IN ('+DecomposeFamille(FamMar)+')';
  end;
  SQL := SQL + ' ORDER BY GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_CODEARTICLE';
  QQ := OpenSQL(SQL,True,-1,'',true);
  if not QQ.eof then
  begin
    //
    IdLigP := Idlig;
    LastP := IdLig;
    EntOuvrage := AddEnteteMateriaux (LIGNES,IdLigPere,IdLig);
    //
    QQ.First;
    repeat
      if (QQ.Fields[0].AsString <> '') then
      begin
        if (FamNiv1 <> QQ.Fields[0].AsString) and (QQ.Fields[0].AsString <> '') then
        begin
          IdLig1 := IdLig;
          NFAM1 := AjouteNoeudFamille ('FN',EntOuvrage,IdLigP,IdLig,QQ.Fields[0].AsString);
          Inc(IDLig);
          FamNiv1 := QQ.Fields[0].AsString;
          FamNiv2 := '---';
          FamNiv3 := '---';
          LastNiv := 2;
          LastP := IdLig1;
        end;
      end else NFAM1 := nil;

      if (QQ.Fields[1].AsString <> '') and (NivFamMax > 1) then
      begin
        if (FamNiv2 <> QQ.Fields[1].AsString) and (QQ.Fields[1].AsString <> '') then
        begin
          IdLig2 := IdLig;
          NFAM2 := AjouteNoeudFamille ('FN',NFAM1,IdLig1,IdLig,QQ.Fields[0].AsString,QQ.Fields[1].AsString);
          Inc(IDLig);
          FamNiv2 := QQ.Fields[1].AsString;
          FamNiv3 := '---';
          LastNiv := 3;
          LastP := IdLig2;
        end;
      end else NFAM2 := nil;

      if (QQ.Fields[2].AsString <> '') and (NivFamMax > 2) then
      begin
        if (FamNiv3 <> QQ.Fields[2].AsString) and (QQ.Fields[2].AsString <> '') then
        begin
          IdLig3 := IdLig;
          NFAM3 := AjouteNoeudFamille ('FN',NFAM2,IdLig2,IdLig,QQ.Fields[0].AsString,QQ.Fields[1].AsString,QQ.Fields[2].AsString);
          Inc(IDLig);
          FamNiv3 := QQ.Fields[2].AsString;
          LastNiv := 4;
          LastP := IdLig3;
        end;
      end else NFAm3 := nil;

      if NFAM3 <> nil then TheP := NFAm3
      else if NFAM2 <> nil Then TheP := NFAM2
      else if NFAM1 <> nil then TheP := NFAM1
      else TheP := EntOuvrage;
      NOUV := AjouteNoeudOuvrage ( TheP,LastP,IdLig,LastNiv + 1, QQ.fields[3].asString,QQ.fields[4].asString,QQ.fields[5].asString,QQ.fields[7].asString,QQ.fields[8].asfloat  );
      inc(IdLig);
      QQ.next;
    until QQ.Eof;
  end;
  ferme (QQ);
end;


procedure ExporteOuvrages (LIGNES : IXMLNode; FamOuv : string; IdLigPere : integer; NivFamMax : Integer; Var Idlig : integer);

  function AddEnteteOuvrage (ROOT : IXMLNode; IDLigPere : Integer; var IdLig : integer) : IXMLNode ;
  var DD : IXMLNode;
  begin
    Result := ROOT.AddChild('LIGNE');
    DD := Result.AddChild('IDLIG'); DD.Text := IntToStr(IDlig);
    DD := Result.AddChild('IDPARENT'); DD.Text := IntToStr(IDLigPere);
    DD := Result.AddChild('TYPELIG'); DD.Text := 'C';
    DD := Result.AddChild('NIVCHAP'); DD.Text := '1';
    DD := Result.AddChild('CODEARTICLE');
    DD := Result.AddChild('CODE_DANS_BIBLIO');
    DD := Result.AddChild('LIBELLE'); DD.Text := 'Ouvrages';
    DD := Result.AddChild('LIBELLECOM'); DD.Text := 'Ouvrages';
    DD := Result.AddChild('LIBELLETEC');
    DD := Result.AddChild('EDITEE'); DD.Text := 'true';
    DD := Result.AddChild('INFO');
    inc(IdLig);
  end;

var QQ: TQuery;
    SQl: string;
    FamNiv1,FamNiv2,FamNiv3 : string;
    NFAM1,NFAM2,NFAm3,NOUV,EntOuvrage,TheP: IXMLNode;
    IdLigP,IdLig1,IdLig2,IdLig3,LastNiv,LastP : Integer;
begin
  //
  FamNiv1 := '---';
  FamNiv2 := '---';
  FamNiv3 := '---';
  LastNiv := 1;
  SQL := 'SELECT GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_CODEARTICLE,GA_LIBELLE,GA_BLOCNOTE,GA_PVHT,GA_QUALIFUNITEVTE,'+
         '('+
          'SELECT TV_TAUXVTE '+
          'FROM TXCPTTVA '+
          'WHERE '+
          'TV_TVAOUTPF="TX1" AND TV_REGIME="FRA" AND TV_CODETAUX=GA_FAMILLETAXE1'+
          ') AS TAUXTVA '+
          'FROM ARTICLE WHERE GA_TYPEARTICLE="OUV"';
  if FamOuv <> '' then
  begin
    SQL := SQL + ' AND GA_FAMILLENIV1 IN ('+DecomposeFamille(FamOuv)+')';
  end;
  SQL := SQL + ' ORDER BY GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_CODEARTICLE';
  QQ := OpenSQL(SQL,True,-1,'',true);
  if not QQ.eof then
  begin
    //
    IdLigP := Idlig;
    EntOuvrage := AddEnteteOuvrage (LIGNES,IdLigPere,IdLig);
    //
    QQ.First;
    repeat
      if (QQ.Fields[0].AsString <> '') then
      begin
        if (FamNiv1 <> QQ.Fields[0].AsString) and (QQ.Fields[0].AsString <> '') then
        begin
          IdLig1 := IdLig;
          NFAM1 := AjouteNoeudFamille ('BO',EntOuvrage,IdLigP,IdLig,QQ.Fields[0].AsString);
          Inc(IDLig);
          FamNiv1 := QQ.Fields[0].AsString;
          FamNiv2 := '---';
          FamNiv3 := '---';
          LastNiv := 2;
          LastP := IdLig1;
        end;
      end else NFAM1 := nil;

      if (QQ.Fields[1].AsString <> '') and (NivFamMax > 1) then
      begin
        if (FamNiv2 <> QQ.Fields[1].AsString) and (QQ.Fields[1].AsString <> '') then
        begin
          IdLig2 := IdLig;
          NFAM2 := AjouteNoeudFamille ('BO',NFAM1,IdLig1,IdLig,QQ.Fields[0].AsString,QQ.Fields[1].AsString);
          Inc(IDLig);
          FamNiv2 := QQ.Fields[1].AsString;
          FamNiv3 := '---';
          LastNiv := 3;
          LastP := IdLig2;
        end;
      end else NFAM2 := nil;

      if (QQ.Fields[2].AsString <> '') and (NivFamMax > 2) then
      begin
        if (FamNiv3 <> QQ.Fields[2].AsString) and (QQ.Fields[2].AsString <> '') then
        begin
          IdLig3 := IdLig;
          NFAM3 := AjouteNoeudFamille ('BO',NFAM2,IdLig2,IdLig,QQ.Fields[0].AsString,QQ.Fields[1].AsString,QQ.Fields[2].AsString);
          Inc(IDLig);
          FamNiv3 := QQ.Fields[2].AsString;
          LastNiv := 4;
          LastP := IdLig3;
        end;
      end else NFAM3 := nil;

      if NFAM3 <> nil then TheP := NFAm3
      else if NFAM2 <> nil Then TheP := NFAM2
      else if NFAM1 <> nil then TheP := NFAM1
      else TheP := EntOuvrage;
      NOUV := AjouteNoeudOuvrage ( TheP,LastP,IdLig,LastNiv + 1, QQ.fields[3].asString,QQ.fields[4].asString,QQ.fields[5].asString,QQ.fields[7].asString,QQ.fields[8].asfloat  );
      inc(IdLig);
      QQ.next;
    until QQ.Eof;
  end;
  ferme (QQ);
end;

function FormatDateBim (TheDate : TDateTime) : string;
var Year,Month,Day : word;
begin
  DecodeDate(TheDate,Year,Month,Day);
  Result := Format('%04d-%.02d-%.02d',[Year,Month,Day]);
end;

procedure ExportBimetre (DateTrait : TDateTime; DirOut,FileName : string; ExportOUV,ExportMAR,ExportPre : boolean;  FamOuv,FamMAr : String; NivFamMax : integer);
var XmlDoc : IXMLDocument ;
    Root,NodeDocs,NodeDoc,headerDoc,TD,LD,CB,DD,TE,CT,Body,LIGNES,Localisation : IXMLNode;
    IdLigP,IDLig : Integer;
begin
  IDLig := 1;
  XmlDoc := NewXMLDocument();
  XmlDoc.Encoding := 'ISO-8859-1';
  XmlDoc.Options := [doNodeAutoIndent];
  Root := XmlDoc.AddChild('BIMETRE');
  DD := Root.AddChild('SOURCE');
  DD.Attributes ['APPLICATION'] := 'LSE BUSINESS BTP';
  DD.Attributes ['VERSION'] := '10.00';
  NodeDocs := Root.AddChild('DOCUMENTS');
  NodeDoc := NodeDocs.AddChild('DOCUMENT');
  headerDoc := NodeDoc.AddChild('HEADER');
  TD := headerDoc.AddChild('TYPE_DOCUMENT'); TD.Text := '4';
  LD := headerDoc.AddChild('LIBELLE'); LD.text := 'Bibliothèque LSE BUSINESS BTP';
  CB := headerDoc.AddChild('CODEBBIBLIO');
  CB.Text := 'PER-LSE';
  DD := headerDoc.AddChild('DATEDOC'); DD.text := FormatDateBim(DateTrait);
  TE := headerDoc.AddChild('TIERS');
  CT := TE.AddChild('CODECLIENT'); CT.Text := '0';
  CT := TE.AddChild('CODECATEGORIE'); CT.text := '0';
  CT := TE.AddChild('STATUTCLIENT'); CT.text := '2';
  Localisation := headerDoc.AddChild('LOCALISATIONS');
  DD := Localisation.AddChild('LOCALISATION');
  CT := Localisation.AddChild('IDLOC'); CT.Text := '1';
  CT := Localisation.AddChild('IDLOCSUP'); CT.Text := '-1';
  CT := Localisation.AddChild('NIVEAU'); CT.Text := '0';
  CT := Localisation.AddChild('LIBELLE'); CT.Text := 'Bibliothèque L.S.E Business BTP';
  Body := NodeDoc.AddChild('BODY');
  LIGNES := Body.AddChild('LIGNE');
  DD := LIGNES.AddChild('IDLIG'); DD.Text := IntToStr(IDlig);
  DD := LIGNES.AddChild('IDPARENT'); DD.Text := '0';
  DD := LIGNES.AddChild('TYPELIG'); DD.Text := 'P';
  DD := LIGNES.AddChild('NUMEROTATION'); DD.Text := '1';
  DD := LIGNES.AddChild('NIVCHAP'); DD.Text := '-1';
  DD := LIGNES.AddChild('CODEARTICLE');
  DD := LIGNES.AddChild('CODE_DANS_BIBLIO');
  DD := LIGNES.AddChild('LIBELLE'); DD.Text := 'Bibliothèque L.S.E Business BTP';
  DD := LIGNES.AddChild('LIBELLECOM');
  DD := LIGNES.AddChild('LIBELLETEC');
  DD := LIGNES.AddChild('EDITEE'); DD.Text := 'true';       
  DD := LIGNES.AddChild('INFO');
  IdLigP := IDLig;
  inc(IDLig);
  if ExportOUV then ExporteOuvrages (LIGNES, FamOuv, IdLigP,NivFamMax, IdLig);
  if ExportMAR then ExporteMateriaux (LIGNES, FamMAr, IdLigP,NivFamMax, IdLig);
  if ExportPre  then ExportePrestations (LIGNES, FamMAr, IdLigP,NivFamMax, IdLig);
  //
  XmlDoc.SaveToFile(IncludeTrailingBackslash(DirOut)+FileName);
  XmlDoc:= nil;
end;

procedure RecupInfosEnteteBIm(Repertoire,NomFic : string; TOBL : TOB);
var XmlDoc : IXMLDocument ;
    DOCUMENTS,DOCUMENT,HEADER,INFO,INFOC : IXMLNode;
    II,JJ,KK,LL,MM : Integer;
    TypeDoc : string;
begin
  XmlDoc := NewXMLDocument();
  TRY
    TRY
      XmlDoc.LoadFromFile (IncludeTrailingBackslash(Repertoire)+NomFic);
    EXCEPT
      On E: Exception do
      begin
        PgiError('Erreur durant Chargement XML : ' + E.Message );
      end;
    end;
    if not XmlDoc.IsEmptyDoc then
    begin
      For II := 0 to Xmldoc.DocumentElement.ChildNodes.Count -1 do
      begin
        DOCUMENTS := XmlDoc.DocumentElement.ChildNodes[II]; // Liste des <Folder>
        if DOCUMENTS.NodeName = 'DOCUMENTS' then
        begin
          for JJ := 0 to DOCUMENTS.ChildNodes.Count -1 do
          begin
            DOCUMENT := DOCUMENTS.ChildNodes [JJ];
            if DOCUMENT.NodeName = 'DOCUMENT' then
            begin
              for KK := 0 to DOCUMENT.ChildNodes.Count -1 do
              begin
                HEADER := DOCUMENT.ChildNodes [KK];
                if HEADER.NodeName = 'HEADER' then
                begin
                  for LL := 0 to HEADER.ChildNodes.Count -1 do
                  begin
                    INFO := HEADER.ChildNodes[LL];
                    if INFO.NodeName = 'TYPE_DOCUMENT' then
                    begin
                      TypeDoc := INFO.NodeValue;
                      TOBL.SetString('TYPE',TypeDoc);

                      if TypeDoc = '1' then TOBL.SetString('LIBTYPE','Devis étude')
                      else if TypeDoc = '2' then TOBL.SetString('LIBTYPE','Devis Marché')
                      else if TypeDoc = '3' then TOBL.SetString('LIBTYPE','Métré Chantier')
                      else if TypeDoc = '4' then TOBL.SetString('LIBTYPE','Bibliothèque')
                      else TOBL.SetSTring('LIBTYPE','Inconnu');
                    end else if (INFO.NodeName ='DATECREATION') or (INFO.NodeName ='DATEDOC') then
                    begin
                      if not VarIsNull(INFO.NodeValue) then TOBL.SetDateTime(INFO.NodeName,DecodeDateBIM(INFO.NodeValue));
                    end else if INFO.NodeName = 'TIERS' then
                    begin
                      for MM := 0 to INFO.ChildNodes.count -1 do
                      begin
                        INFOC := INFO.ChildNodes[MM];
                        if TOBL.FieldExists(INFOC.NodeName) then
                        begin
                          if not VarIsNull(INFOC.NodeValue) then TOBL.SetString(INFOC.NodeName,INFOC.NodeValue);
                        end;
                      end;
                    end else if TOBL.FieldExists(INFO.NodeName) then
                    begin
                      if not VarIsNull(INFO.NodeValue) then TOBL.SetString(INFO.NodeName,INFO.NodeValue);
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
      end;
    end;
  FINALLY
    XmlDoc:= nil;
  end;
end;

function ConstitueDocument(NaturePiece : string; Repertoire : string; TOBL : TOB; RAPPORT: THRichEditOLE) : boolean;

  function FindParent (IdParent: Integer; TOBpiece : TOB) : TOB;
  var II : Integer;
  begin
    Result := nil;
    for II := TOBPiece.detail.count -1 downto 0 do
    begin
      if TOBpiece.detail[II].GetInteger('UNIQUEID')=IdParent then
      begin
        Result := TOBpiece.detail[II];
        break;
      end;
    end;
  end;

  function ConstitueEntete(TX : TDataTransport;TOBL : TOB) : boolean;
  var CodeTiers,DomaineSoc,Etablissement : string;
      QQ : TQuery;
      ValNumP : T_ValideNumPieceVide;
      io: TIoErr;
  begin
    CodeTiers := '';
    Etablissement :=VH^.EtablisDefaut;
    InitTOBPiece (TX.TOBPiece);
    if TOBL.GetString('AFFAIRE')<> '' then
    begin
      QQ := OpenSQL('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBL.GetString('AFFAIRE')+'"',True,1,'',true);
      if not QQ.eof then
      begin
        TX.TOBAFFAIRE.SelectDb('',QQ);
      end;
      ferme (QQ);
      CodeTiers :=TX.TOBAffaire.getString('AFF_TIERS');
      DomaineSoc := TX.TOBAffaire.getString('AFF_DOMAINE');
      AffaireVersPiece(TX.TobPiece, TX.TobAffaire);
      TX.TOBPiece.SetString('GP_AFFAIRE',TX.TobAffaire.getString('AFF_AFFAIRE'));
      TX.TOBPiece.SetString('GP_AFFAIRE1',TX.TobAffaire.getString('AFF_AFFAIRE1'));
      TX.TOBPiece.SetString('GP_AFFAIRE2',TX.TobAffaire.getString('AFF_AFFAIRE2'));
      TX.TOBPiece.SetString('GP_AFFAIRE3',TX.TobAffaire.getString('AFF_AFFAIRE3'));
      TX.TOBPiece.SetString('GP_AVENANT',TX.TobAffaire.getString('AFF_AVENANT'));
    end;

    if CodeTiers = ''  then CodeTiers := GetParamSocSecur('SO_BTBIMCLIENTDEF','');
    if DomaineSoc = '' then DomaineSoc:= GetParamSocSecur('SO_BTDOMAINEDEFAUT','');
    //
    Result := (RemplirTOBTiers (TX.TOBTiers,CodeTiers,NaturePiece,True)=trtOk);
    if not result then
    begin
      PGIInfo('Le client par défaut n''existe pas');
      Exit;
    end;
    TX.TOBPiece.SetString ('GP_NATUREPIECEG',NaturePiece);
    TX.TOBPiece.SetString ('GP_SOUCHE',GetSoucheG(NaturePiece,Etablissement,DomaineSoc));
    //
    ValNumP := T_ValideNumPieceVide.Create;
    ValNumP.Souche := TX.TOBPiece.GetString ('GP_SOUCHE');
    io:=Transactions(ValNumP.ValideNumPieceVide,5) ;
    if io=oeOk then TX.TOBPiece.SetInteger('GP_NUMERO',ValNumP.NewNum)  else TX.TOBPiece.SetInteger('GP_NUMERO',0);
    ValNumP.Free;
    //
    if TX.TOBPiece.GetInteger('GP_NUMERO') = 0 then begin Result := false; Exit; end;
    //
    TX.TOBpiece.setSTring('GP_VENTEACHAT',GetInfoParPiece(NaturePiece, 'GPP_VENTEACHAT'));
    TX.TOBPIECE.setString('GP_DEVISE',TX.TOBTiers.GetString('T_DEVISE'));
    TX.DEV.Code := TX.TOBPIECE.GetValue('GP_DEVISE');
    GetInfosDevise(TX.DEV);
    TX.DEV.Taux := GetTaux(TX.DEV.Code, TX.DEV.DateTaux, TOBL.GetDateTime('DATEDOC'));
    //
    TX.TOBPiece.SetString ('GP_DEVISE', TX.DEV.Code);
    TX.TOBPiece.SetDouble ('GP_TAUXDEV', TX.DEV.Taux);
    TX.TOBPiece.SetDatetime('GP_DATETAUXDEV', TX.DEV.DateTaux);
    //
    TiersVersAdresses(TX.TOBTiers, TX.TOBAdresses, TX.TOBPiece);
    AffaireVersAdresses(TX.TOBAffaire,TX.TOBAdresses,TX.TOBPiece);
    LivAffaireVersAdresses (TX.TOBAffaire,TX.TOBAdresses,TX.TOBPiece);
    //
    TX.TOBPiece.SetString('GP_DOMAINE',DomaineSoc);
    TX.TOBPiece.PutValue('GP_TIERS', CodeTiers);
    TX.TOBPiece.SetString('GP_REGIMETAXE', VH^.RegimeDefaut);
    if VH^.EtablisDefaut <> '' then TX.TOBPiece.SetString('GP_ETABLISSEMENT', VH^.EtablisDefaut);
    if VH_GC.GCDepotDefaut <> '' then TX.TOBPiece.SetString('GP_DEPOT', VH_GC.GCDepotDefaut);
    TX.TOBPiece.setString('GP_SOUCHEG',GetSoucheG(NaturePiece, TX.TOBPiece.GetValue('GP_ETABLISSEMENT'), DomaineSoc));
    //
    TiersVersPiece(TX.TOBTiers, TX.TOBPiece);
    // Représentant du tiers
    if TX.TobPiece.GetValue('GP_REPRESENTANT') = '' then
    begin
      if TX.TOBTiers.FieldExists('T_REPRESENTANT') then
        TX.TobPiece.PutValue('GP_REPRESENTANT', TX.TOBTiers.GetValue('T_REPRESENTANT'));
    end;
    TX.TOBPiece.SetDateTime('GP_DATEPIECE',TOBL.GetDateTime('DATEDOC'));
    TX.TOBPiece.SetString('GP_REFINTERNE',TOBL.GetString('CODEAFFAIRE'));
    MajTypeFact (TX.TOBPiece, TX.TOBAffaire.GetValue('AFF_GENERAUTO'));
  end;

  function AddLigne (TX :TDataTransport; Indice : integer) : TOB;
  begin
    result := TOB.Create('LIGNE', TX.TOBPiece, Indice);
    NewTOBLigneFille(result);
    AddLesSupLigne(result, False,false);
    InitLesSupLigne(result);
    result.AddChampSupValeur('UNIQUEID',-1);
    //
    InitLigneVide(TX.TOBPiece, result, TX.TOBTiers, TX.TOBAffaire,result.GetIndex+1, 1);
    result.PutValue('GL_ENCONTREMARQUE', '-');
    if TX.TOBPiece.getValue('GP_DOMAINE')<>'' then result.PutValue('GL_DOMAINE', TX.TOBPiece.GetValue('GP_DOMAINE'));
    PieceVersLigne(TX.TobPiece, result);
  end;

  procedure AjouteParagraphe (TX :TDataTransport ; Numerotation,Libelle : string; IdParent,UniqueID : integer);
  var NIv : Integer;
      Indice : Integer;
      TOBL : TOB;
      TOBpere: TOB;
  begin
    if TX.TOBPiece.findfirst(['UNIQUEID'],[UniqueID],true)<> nil then exit;
    TOBpere := findParent (IdParent , TX.Tobpiece);
    if TOBpere = nil then
    begin
      NIv := 1;
      Indice := -1; // a la fin du document
      TOBL := AddLIgne (TX,Indice);
      TOBL.SetString('GL_TYPELIGNE','DP'+InttoStr(NIv));
      TOBL.SetString('GL_LIBELLE',Libelle);
      TOBL.SetInteger('GL_NIVEAUIMBRIC',NIV);
      TOBL.SetString('GLC_NUMEROTATION',Numerotation);
      TOBL.SetString('GLC_NUMFORCED','X');
      //
      TOBL := AddLIgne (TX,Indice);
      TOBL.SetInteger('UNIQUEID',UniqueID);
      TOBL.SetString('GL_TYPELIGNE','TP'+InttoStr(NIv));
      TOBL.SetString('GL_LIBELLE','TOTAL ' +Libelle);
      TOBL.SetInteger('GL_NIVEAUIMBRIC',NIV);
    end else
    begin
      NIv := TOBpere.GetInteger('GL_NIVEAUIMBRIC')+1;
      indice := TOBpere.GetIndex;
      TOBL := AddLIgne (TX,Indice);
      TOBL.SetString('GL_TYPELIGNE','TP'+InttoStr(NIv));
      TOBL.SetString('GL_LIBELLE','TOTAL ' +Libelle);
      TOBL.SetInteger('GL_NIVEAUIMBRIC',NIV);
      TOBL.SetInteger('UNIQUEID',UniqueID);
      //
      TOBL := AddLIgne (TX,Indice);
      TOBL.SetString('GL_TYPELIGNE','DP'+InttoStr(NIv));
      TOBL.SetString('GL_LIBELLE',Libelle);
      TOBL.SetInteger('GL_NIVEAUIMBRIC',NIV);
      TOBL.SetString('GLC_NUMEROTATION',Numerotation);
      TOBL.SetString('GLC_NUMFORCED','X');
    end;
  end;

  function GetArticle (TX : TDataTransport; CodeArticle : string) : TOB;
  var ArticleUnique : string;
  begin
    ArticleUnique := CodeArticleUnique(CodeArticle,'','','','','');
    result := findTobArt(ArticleUnique,TX.TOBArticles);
  end;

  procedure TraitePrixLigneCourante (TX : TDataTransport; TOBL : TOB;Pu : double);
  var TypeArticle : string;
      EnHt : boolean;
  begin
  //
    TypeArticle := TOBL.getValue('GL_TYPEARTICLE');
    EnHt := (TX.TOBPiece.getValue('GP_FACTUREHT')='X');
    if (TypeArticle = 'OUV') or (TypeArticle = 'ARP') then
    begin
      TraitePrixOuvrage(TX.TOBPiece,TOBL,TX.TOBBases,TX.TOBBasesL,TX.TOBOuvrage, EnHt, Pu,0,TX.DEV,false,False,True,GetParamSocSecur('SO_BTAPPOFFPAEQPV',false));
      ReinitCoefMarg (TOBL,TX.TOBOuvrage);
      TOBL.PutValue('GL_REMISABLELIGNE','X');
    end;
  //
    if EnHt then
    begin
      TOBL.PutValue('GL_PUHTDEV',Pu);
      TOBL.putValue('GL_PUHT',DEVISETOPIVOTEx (Pu,TX.DEV.Taux,TX.DEV.Quotite,V_PGI.okdecP));
      TOBL.PutValue('GL_PUHTNETDEV',Pu);
      TOBL.putValue('GL_PUHTNET',DEVISETOPIVOTEx (Pu,TX.DEV.Taux,TX.DEV.Quotite,V_PGI.okdecP));
    end else
    begin
      TOBL.PutValue('GL_PUTTCDEV',Pu);
      TOBL.putValue('GL_PUTTC',DEVISETOPIVOTEx (Pu,TX.DEV.Taux,TX.DEV.Quotite,V_PGI.okdecP));
      TOBL.PutValue('GL_PUTTCNETDEV',Pu);
      TOBL.putValue('GL_PUTTCNET',DEVISETOPIVOTEx (Pu,TX.DEV.Taux,TX.DEV.Quotite,V_PGI.okdecP));
    end;
    if GetParamSocSecur('SO_BTAPPOFFPAEQPV',false) then
    begin
      TOBL.PutValue('GL_DPA',Pu);
      TOBL.SetDouble('GL_COEFFG',0);
      TOBL.SetDouble('GL_COEFMARG',1);
    end;
  //
    TOBL.putValue('GL_BLOQUETARIF','X');
  end;


  procedure AjouteLigneDet (IdParent : Integer; TX :TDataTransport ; CodeArticle,CodeBiblio,Libelle,LibelleTec,LibelleCom,Numerotation : string; PVHT,TauxTva,QuantiteTot : Double; Unite : string);
  var Niv,Indice : Integer;
      TOBL,TOBA : TOB;
      TOBpere : TOB;
  begin
    TOBPere := FindParent (IdParent,TX.TOBPiece); if TOBpere = nil then Exit;
    Niv := TOBpere.GetInteger('GL_NIVEAUIMBRIC');
    indice := TOBpere.GetIndex;
    TOBL := AddLIgne (TX,Indice);
    TOBA := GetArticle (TX,CodeBiblio);
    if TOBA = nil then
    begin
      TOBA := GetArticle (TX,CodeDefaut);
    end;
    if TOBA <> nil then
    begin
      TOBL.PutValue('GL_ARTICLE',TOBA.GetString('GA_ARTICLE'));
      TOBL.PutValue('GL_TYPELIGNE','ART');
      TOBL.PutValue('GL_TYPEARTICLE',TOBA.GetValue('GA_TYPEARTICLE'));
      TOBL.PutValue('GL_CODEARTICLE',TOBA.GetValue('GA_CODEARTICLE'));
      TOBL.SetString('GL_REFARTSAISIE',CodeArticle);
      TOBL.PutValue('GL_REFARTBARRE', TOBA.GetValue('GA_CODEBARRE'));
      TOBL.PutValue('GL_PRIXPOURQTE', TOBA.GetValue('GA_PRIXPOURQTE'));
      TOBL.PutValue('GL_TYPEREF', 'ART');
      TOBL.PutValue('GL_RECALCULER','X');
      ArticleVersLigne (TX.TOBPiece,TOBA,TX.TobConds,TOBL,TX.TOBTiers);
      TraiteLesOuvrages(nil,TX.TOBPiece, TX.TOBArticles, TX.TOBOuvrage,nil,nil, TOBL.GetIndex+1, False, TX.DEV, true);
    end;
    TOBL.putValue('GL_COEFFC',0.0);
    TOBL.putValue('GL_COEFFR',0.0);
    TOBL.putValue('GL_COEFFG',0.0);
    TOBL.putValue('GL_COEFMARG',1.0);
    //
    //if PVHT <> 0 then TraitePrixLigneCourante (TX,TOBL,PVHT);
    //
    TOBL.SetString('GL_LIBELLE',copy(libelle,1,70));
    if LibelleCom <> '' then TOBL.SetString('GL_BLOCNOTE',LibelleCom);
    TOBL.SetString('GL_TYPELIGNE','ART');
    TOBL.SetDouble('GL_QTEFACT',QuantiteTot);
    TOBL.SetDouble('GL_QTESTOCK',QuantiteTot);
    TOBL.SetDouble('GL_QTERESTE',QuantiteTot);
    TOBL.SetString('GL_QUALIFUNITEVTE',Unite);
    TOBL.SetString('GLC_NUMEROTATION',Numerotation);
    TOBL.SetString('GLC_NUMFORCED','X');
    TOBL.SetInteger('GL_NIVEAUIMBRIC',NIV);
  end;

  procedure TraiteLesLignes (LIGNES : IXMLNode; TX :TDataTransport );
  var II : Integer;
      Libelle,LibelleCom,LibelleTec,CodeArticle,CodeBiblio,Numerotation,TypeLig,Unite : string;
      NivChap,IdLig,IdParent : Integer;
      PVHT,TauxTva,QuantiteTot : Double;
      NodeDetail,NodeCur : IXMLNode;
  begin
    NodeDetail := nil;
    for II := 0 to LIGNES.ChildNodes.count -1 do
    begin
      NodeCur := LIGNES.ChildNodes [II];
      if NodeCur.NodeName = 'IDPARENT' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then IdParent := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'IDLIG' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then IdLig := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'TYPELIG' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then TypeLig := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'NUMEROTATION' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then Numerotation := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'NIVCHAP' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then NivChap := StrToInt(NodeCur.NodeValue);
      end else if NodeCur.NodeName = 'CODEARTICLE' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then CodeArticle := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'CODE_DANS_BIBLIO' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then CodeBiblio := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'LIBELLE' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then Libelle := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'LIBELLECOM' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then LibelleCom := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'LIBELLETEC' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then LibelleTec := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'PVHT' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then PvHT := valeur(NodeCur.NodeValue);
      end else if NodeCur.NodeName = 'TAUXTVA' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then TauxTva := valeur(NodeCur.NodeValue);
      end else if NodeCur.NodeName = 'QUANTITETOT' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then QuantiteTot := valeur(NodeCur.NodeValue);
      end else if NodeCur.NodeName = 'UNITE' then
      begin
        if not VarIsNull(NodeCur.NodeValue) then Unite := NodeCur.NodeValue;
      end else if NodeCur.NodeName = 'LIGNE' then
      begin
        NodeDetail := NodeCur;
        // -- fin de parcours --> on a tous les elements pour savoir ce que l'on fait
        if (TypeLig='P') or (NivChap <= 0)then
        begin
          TraiteLesLignes (NodeDetail,TX);
        end else if (TypeLig = 'C') then
        begin
          AjouteParagraphe (TX,Numerotation,Libelle,IdParent,IDLig);
          TraiteLesLignes (NodeDetail,TX);
        end;
      end;
    end;
    if TypeLig= 'A' then
    begin
      AjouteLigneDet (IdParent,TX,CodeArticle,CodeBiblio,Libelle,LibelleTec,LibelleCom,Numerotation,PVHT,TauxTva,QuantiteTot,Unite);
    end;

  end;

  procedure AjusteDocument (TX :TDataTransport) ;
  var indice : integer;
      TOBL : TOB;
      IndiceOuv : integer;
  begin
    {$IFDEF BTP}
    for Indice := 0 to TX.TOBPiece.detail.count -1 do
    begin
      TOBL := TX.TOBpiece.detail[indice];
      IndiceOuv := TOBL.GetValue('GL_INDICENOMEN');
      if (IndiceOuv > 0) and (IsOuvrage(TOBL)) then ReAffecteLigOuv(IndiceOuv, TobL, TX.TobOuvrage);
      ProtectionCoef  (TOBL);
    end;
    {$ENDIF}
  end;

var TX : TDataTransport;
    XmlDoc : IXMLDocument ;
    DOCUMENTS,D1,DOCUMENT,LIGNES : IXMLNode;
    II,JJ,KK,LL : Integer;
begin
  CodeDefaut := GetParamSocSecur('SO_BTARTICLEDIV','');
  if CodeDefaut = '' then
  begin
    result := false;
    PgiError('l''article par défaut des appels d''offres n''est pas renseigné');
    Exit;
  end;
  result := true;
  TX := TDataTransport.create (NaturePiece);
  XmlDoc := NewXMLDocument();
  try
    if not ConstitueEntete(TX,TOBL) then Exit;
    TRY
      XmlDoc.LoadFromFile (IncludeTrailingBackslash(Repertoire)+TOBL.GetString('NOMFIC'));
    EXCEPT
      On E: Exception do
      begin
        result := false;
        PgiError('Erreur durant Chargement XML : ' + E.Message );
      end;
    end;
    if not XmlDoc.IsEmptyDoc then
    begin
      For II := 0 to Xmldoc.DocumentElement.ChildNodes.Count -1 do
      begin
        DOCUMENTS := XmlDoc.DocumentElement.ChildNodes[II]; // Liste des <Folder>
        if DOCUMENTS.NodeName = 'DOCUMENTS' then
        begin
          for JJ := 0 to DOCUMENTS.ChildNodes.Count -1 do
          begin
            DOCUMENT := DOCUMENTS.ChildNodes [JJ];
            if DOCUMENT.NodeName = 'DOCUMENT' then
            begin
              for KK := 0 to DOCUMENT.ChildNodes.Count -1 do
              begin
                D1 := DOCUMENT.ChildNodes [KK];
                if D1.NodeName = 'BODY' then
                begin
                  for LL := 0 to D1.ChildNodes.Count -1 do
                  begin
                    LIGNES := D1.ChildNodes[LL];
                    if LIGNES.NodeName = 'LIGNE' then
                    begin
                      TraiteLesLignes (LIGNES,TX);
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
      end;
      if TX.TOBpiece.detail.count > 0 then
      begin
        NumeroteLignesGC (nil,TX.TOBPiece,true);
        PutValueDetail (TX.TOBPiece,'GP_RECALCULER','X');
        CalculeMontantsDoc (TX.TOBPiece,TX.TOBOuvrage,true,true);
        CalculFacture (TX.TOBAFFAire,TX.TOBPiece,TX.TOBPieceTrait,TX.TOBSSTRAIT,TX.TOBOuvrage,TX.TOBOUvragesP,TX.TOBBases,TX.TOBBasesL,TX.TOBTiers,TX.TOBArticles,TX.TOBPorcs,TX.TOBPieceRG,TX.TOBBasesRG,nil,TX.DEV);
        ValideLaPeriode (TX.TOBPIECE);
        AjusteDocument (TX);
        TX.SetModified;
        if TRANSACTIONS(TX.MajTobs,0) = oeOk then
        begin
          Rapport.lines.Add(Format(' Importation du fichier %s  ---> %s Numéro %d généré -- OK --',[TOBL.GetString('NOMFIC'), RechDom('GCNATUREPIECEG',TX.TOBpiece.getString('GP_NATUREPIECEG'),False),TX.TOBPIECE.GetInteger('GP_NUMERO')]));
        end else
        begin
          result := false;
        end;
      end;
    end;
  finally
    TX.Free;
    XmlDoc:= nil;
  end;
end;


{ TDataTransport }

constructor TDataTransport.create(NaturePiece : string);
begin
  TOBPiece:=TOB.Create('PIECE',Nil,-1)     ;
  AddLesSupEntete (TOBPiece);
  TOBPiece_O := TOB.create ('PIECEORIG',nil,-1);
  // ---
  TOBBases:=TOB.Create('BASES',Nil,-1)     ;
  TOBBasesL:=TOB.Create('LES BASES LIGNE',Nil,-1)     ;
  TOBEches:=TOB.Create('LES ECHEANCES',Nil,-1) ;
  TOBPorcs:=TOB.Create('PORCS',Nil,-1)     ;
  // Fiches
  TOBTiers:=TOB.Create('TIERS',Nil,-1) ;
  TOBTiers.AddChampSup('RIB',False) ;

  TOBArticles:=TOB.Create('ARTICLES',Nil,-1) ;
  TOBConds:=TOB.Create('CONDS',Nil,-1) ;
  TOBTarif:=TOB.Create('TARIF',Nil,-1) ;
  TOBComms:=TOB.Create('COMMERCIAUX',Nil,-1) ;
  TOBCatalogu:=TOB.Create('LECATALOGUE',Nil,-1) ;
  // Adresses
  TOBAdresses:=TOB.Create('LESADRESSES',Nil,-1) ;
  if GetParamSoc('SO_GCPIECEADRESSE') then
  BEGIN
     TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Livraison}
     TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Facturation}
  END else
  BEGIN
     TOB.Create('ADRESSES',TOBAdresses,-1) ; {Livraison}
     TOB.Create('ADRESSES',TOBAdresses,-1) ; {Facturation}
  END ;
  // Divers
  TOBCXV:=TOB.Create('',Nil,-1) ;
  TOBNomen:=TOB.Create('NOMENCLATURES',Nil,-1) ;
  TOBDim:=TOB.Create('',Nil,-1) ;
  TOBLOT:=TOB.Create('',Nil,-1) ;
  TOBSerie:=TOB.Create('',Nil,-1) ;
  TOBSerRel:=TOB.Create('',Nil,-1) ;
  TOBAcomptes:=TOB.Create('',Nil,-1) ;
  TOBDispoContreM:=TOB.Create('',Nil,-1);
  // Affaires
  TOBAffaire:=TOB.Create('AFFAIRE',Nil,-1) ;
  // Comptabilité
  TOBANAP:=TOB.Create('',Nil,-1) ;
  TOBANAS:=TOB.Create('',Nil,-1) ;
  //Saisie Code Barres
  TOBGSA:=TOB.Create('',Nil,-1);
  // Ouvrages
  TOBOuvrage:=TOB.Create ('OUVRAGES',nil,-1);
  // Ouvrages^plat
  TOBOuvragesP:=TOB.Create ('OUVRAGES PLAT',nil,-1);
  // textes debut et fin
  TOBLIENOLE:=TOB.Create ('LIENS',nil,-1);
  // retenues de garantie
  TOBPieceRG:=TOB.create('PIECERRET',Nil,-1);
  // Bases de tva sur RG
  TOBBasesRG:=TOB.create('BASESRG',Nil,-1);
  TOBLienXls := TOB.Create ('LIENSPGIEXCEL',nil,-1);
  TOBPieceTrait := TOB.Create ('LES PIECES TRAITS',nil,-1);
  TOBSSTrait := TOB.Create ('LES SOUS TRAITS',nil,-1);

  InitStructureETL;
  InitStructureOUV;
  MemoriseChampsSupLigneETL (NaturePiece,True);
  MemoriseChampsSupLigneOUV (NaturePiece);
end;

destructor TDataTransport.destroy;
begin
  InitStructureETL;
  InitStructureOUV;
  TOBPiece.free; TOBPiece_O.free;
  TOBBases.free;
  TOBBasesL.free;
  TOBEches.free;
  TOBPorcs.free;
  // Fiches
  TOBTiers.free;
  TOBArticles.free;
  TOBConds.free;
  TOBTarif.free;
  TOBComms.free;
  TOBCatalogu.free;
  // Adresses
  TOBAdresses.free;
  // Divers
  TOBCXV.free;
  TOBNomen.free;
  TOBDim.free;
  TOBLOT.free;
  TOBSerie.free;
  TOBSerRel.free;
  TOBAcomptes.free;
  TOBDispoContreM.free;
  // Affaires
  TOBAffaire.free;
  // Comptabilité
//  TOBCPTA.free;
  TOBANAP.free;
  TOBANAS.free;
  //Saisie Code Barres
  TOBGSA.free;
  // Ouvrages
  TOBOuvrage.free;
  // Ouvrages plat
  TOBOuvragesP.free;
  // textes debut et fin
  TOBLIENOLE.free;
  // retenues de garantie
  TOBPieceRG.free;
  // Bases de tva sur RG
  TOBBasesRG.free;
  // Lien PGI-Excel
  TOBLienXls.free;
  TOBPieceTrait.free;
  TOBSSTrait.free;
  inherited;
end;



procedure TDataTransport.MajTobs;
begin
  V_PGI.ioError := OeOk;
  IF TOBLienXls.Detail.count > 0 then
  begin
    TOBLienXls.InsertDB(nil,true);
  end;
  // --
  if V_PGI.IoError=oeOk then ValideLesAdresses (ToBPiece,TOBPiece_O,TOBadresses);
  if V_PGI.IoError = oeOk then ValideLesOuv(TOBOuvrage, TOBPiece);
  if V_PGI.IoError = oeOk then ValideLesBases(TOBPiece,TobBases,TOBBasesL);
  if V_PGI.IoError = oeOk then ValideLesArticlesFromOuv(TOBarticles,TOBOuvrage);
  if V_PGI.IOError = OeOk then ValideLesLignesCompl (TOBpiece,TOBPiece_O);
  if V_PGI.IoError=oeOk then if not TOBPiece.InsertDBByNivel (true) then V_PGI.IoError := OeUnknown;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBBases.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBBasesL.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBEches.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBAnaP.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBAnaS.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  //
  if V_PGI.IoError=oeOk then MajSousAffaire(TobPiece,nil,'00',taCreat,false,false);
end;

procedure TDataTransport.SetModified;
begin
  TOBPiece.SetAllModifie(True) ;
  TOBPiece_O.SetAllModifie(True) ;
  TOBBases.SetAllModifie(True) ;
  TOBBasesL.SetAllModifie(True) ;
  TOBEches.SetAllModifie(True) ;
  TOBPorcs.SetAllModifie(True) ;
  TOBTiers.SetAllModifie(True) ;
  TOBArticles.SetAllModifie(True) ;
  TOBConds.SetAllModifie(True) ;
  TOBTarif.SetAllModifie(True) ;
  TOBComms.SetAllModifie(True) ;
  TOBCatalogu.SetAllModifie(True) ;
  TOBAdresses.SetAllModifie(True) ;
  TOBCXV.SetAllModifie(True) ;
  TOBNomen.SetAllModifie(True) ;
  TOBDim.SetAllModifie(True) ;
  TOBLOT.SetAllModifie(True) ;
  TOBSerie.SetAllModifie(True) ;
  TOBSerRel.SetAllModifie(True) ;
  TOBAcomptes.SetAllModifie(True) ;
  TOBDispoContreM.SetAllModifie(True) ;
  TOBAffaire.SetAllModifie(True) ;
//  TOBCPTA.SetAllModifie(True) ;
  TOBANAP.SetAllModifie(True) ;
  TOBANAS.SetAllModifie(True) ;
  TOBGSA.SetAllModifie(True) ;
  TOBOuvrage.SetAllModifie(True) ;
  TOBOuvragesP.SetAllModifie(True) ;
  TOBLIENOLE.SetAllModifie(True) ;
  TOBPieceRG.SetAllModifie(True) ;
  TOBBasesRG.SetAllModifie(True) ;
  TOBLienXls.SetAllModifie(True) ;
  TOBPieceTrait.SetAllModifie(True) ;
  TOBSSTrait.SetAllModifie(True) ;
end;

end.
