unit UShareDB;

interface
uses
  classes
  ,forms
  ,DB
  ,ADODB
  ,Windows
  ,SysUtils
  ,Messages
  ,UdefImportDatas
  ,HDB
  ,DBCtrls
  ,CBPDatabases
  ,uTob
  {$IFNDEF DBXPRESS}, dbtables {$ELSE}, uDbxDataSet {$ENDIF}
  ;

type

  ShareDb = class
    class function IsShareMaster (Application: TApplication; ServerName,DBName : string) : boolean;
    class function IsY2DB (DbName : string) : Boolean;
    class function ISDBPresent (ServerName,DbName : string; var ErrorString : string) : boolean;
    class function GrpIsPresent (ServerName,DBName, Grp: string) : Boolean;
    class function AppliqueDBPrinc(SERVERNAME,DbName : string; ListeTables,ListeGrps : TListDBStatus; LDossiers,LDatas : TStringList ) : boolean ;
    class function AppliqueDBSecond(SERVERNAME, DBMaster, DbName: string;LDatas: TStringList): boolean;
  end;

  AdoCnx = class
    class function GetConnectionString (ServerName,DbName : string) : string;
    class function PrepareSQl (SQL : string) : string;
    class function USDateTimeToStr (TheDate : TDateTime) : string;
  end;

implementation

uses DateUtils,HCtrls;


function ConstituePartageDatas (CNX : TADOConnection; DBPrinc : string; LDatas : TStringList) : Boolean;
var SQL,SQLEx : String;
    OneShare : string;
    NOMTABLE,MODEFONC,NOMBASE,TYPTABLE,VUE : string;
    II : Integer;
begin
  Result := True;
  //
  SQL := 'INSERT INTO DESHARE (DS_NOMTABLE, DS_MODEFONC, DS_NOMBASE, DS_TYPTABLE, DS_VUE) VALUES ("%s","%s","%s","%s","%s")';
  TRY
    for II := 0 to LDatas.Count -1 do
    begin
      OneShare  := LDatas.Strings[II];
      NOMTABLE     := READTOKENPipe(OneShare,'|');
      MODEFONC   := READTOKENPipe(OneShare,'|');
      NOMBASE     := READTOKENPipe(OneShare,'|'); // provenance
      TYPTABLE     := READTOKENPipe(OneShare,'|');
      VUE := READTOKENPipe(OneShare,'|');
      SQLEx := Format(SQL,[NOMTABLE,MODEFONC,DBPrinc,TYPTABLE,VUE]);
      TRY
        SQLEx := AdoCnx.PrepareSQl(SQLEx);
        
        CNX.Execute(SQLEx);
      EXCEPT
        Result := false;
        break;
      END;
    end;
    
  FINALLY
  END;
end;


function ConstitueDossier (CNX : TADOConnection; LDossiers : TStringList) : Boolean;
var SQL,SQLEx : String;
    OneDossier : string;
    NOMBASE,NODOSSIER,SOCIETE,LIBELLE,UTILISATEUR,GUIDDOSSIER,TYPEDOSSIER : string;
    II : Integer;
begin
  Result := True;
  //
  SQL := 'INSERT INTO DOSSIER '+
         '(DOS_NODOSSIER,DOS_CODEPER,DOS_GUIDPER,DOS_SOCIETE,DOS_LIBELLE,'+
         'DOS_NODISQUE,DOS_NODISQUELOC,DOS_LASERIE,DOS_GROUPECONF,DOS_VERROU,'+
         'DOS_DATEDEPART,DOS_DATERETOUR,DOS_UTILISATEUR,DOS_CABINET,DOS_NOMSERVEUR, '+
         'DOS_NETEXPERT,DOS_NECPSEQ,DOS_NERECDATE,DOS_NERECNBFIC,DOS_NECPDATEARRET, '+
         'DOS_PASSWORD,DOS_PWDGLOBAL,DOS_NETEXPERTGED,DOS_NECDKEY,DOS_USRS1, '+
         'DOS_PWDS1,DOS_EWSCREE,DOS_SERIAS1,DOS_WINSTALL,DOS_ABSENT, '+
         'DOS_VERSIONBASE,DOS_TYPEMAJ,DOS_DETAILMAJ,DOS_NOMBASE, '+
         'DOS_GUIDDOSSIER,DOS_WINOUV,DOS_WINSTR,DOS_NONTRAITE, '+
         'DOS_STATUTEDIFIS,DOS_STATUTEDISOC,DOS_STATUTEDIJUR,DOS_SWSACTIVE, '+
         'DOS_SWSDATEACTIV,DOS_TYPEDOSSIER) values ('+
         '"%s",0,"","%s","%s",'+
         '0,0,"","","",'+
         '"19000101","19000101","%s","-","",'+
         '"-",0,"19000101",0,"19000101",'+
         '"","-","-","","",'+
         '"","-","-","","-",'+
         '0,"","","%s",'+
         '"%s","","","-",'+
         '"","","","-",'+
         '"19000101","%s")';
  TRY
    for II := 0 to LDossiers.Count -1 do
    begin
      OneDossier  := LDossiers.Strings[II];
      NOMBASE     := READTOKENPipe(OneDossier,'|');
      NODOSSIER   := READTOKENPipe(OneDossier,'|');
      SOCIETE     := READTOKENPipe(OneDossier,'|');
      LIBELLE     := READTOKENPipe(OneDossier,'|');
      UTILISATEUR := READTOKENPipe(OneDossier,'|');
      GUIDDOSSIER := READTOKENPipe(OneDossier,'|');
      TYPEDOSSIER := READTOKENPipe(OneDossier,'|');
      SQLEx := Format(SQL,[NODOSSIER,SOCIETE,LIBELLE,UTILISATEUR,NOMBASE,GUIDDOSSIER,TYPEDOSSIER]);
      TRY
        SQLEx := AdoCnx.PrepareSQl(SQLEx);
        
        CNX.Execute(SQLEx);
      EXCEPT
        Result := false;
        break;
      END;
    end;
    
  FINALLY
  END;
end;

function ConstitueYmultiDossier (CNX: TADOConnection; ListeTables,ListeGrps :TListDBStatus ) : boolean;
var SQL : String;
    MemoST : TStringList;
    DbList,GrpList : string;
    II : Integer;
begin
  Result := True;
  //
  MemoST := TStringList.Create;
  TRY
    DBList := '';
    for II := 0 to ListeTables.Count -1 do
    begin
      DbList := DBList + TStatus(ListeTables.Items[II]).DB+'|'+TStatus(ListeTables.Items[II]).DB+';';
    end;
    MemoST.Add(DbList);
    //
    GrpList := '';
    for II := 0 to ListeGrps.Count -1 do
    begin
      GrpList := GrpList + TStatus(ListeGrps.Items[II]).DB+'|'+TStatus(ListeGrps.Items[II]).DB+';';
    end;
    MemoST.Add(GrpList);
    //
    SQL := 'INSERT INTO YMULTIDOSSIER (YMD_CODE,YMD_LIBELLE,YMD_UTILISATEUR,YMD_CREATEUR,YMD_DATECREATION,YMD_DATEMODIF,YMD_DETAILS,YMD_BOOLEEN) values ("%s","%s","%s","%s","%s","%s","%s"+CHAR(13)+CHAR(10)+"%s","-")';
    SQL := Format(SQL,['##MULTISOC','Multi société','CEG','CEG',AdoCnx.USDateTimeToStr(Now),AdoCnx.USDateTimeToStr(Now),MemoST.Strings[0],MemoST.Strings[1]]);
    TRY
      SQL := AdoCnx.PrepareSQl(SQL);
      CNX.Execute(SQL);
    EXCEPT
      Result := false;
    END;
  FINALLY
    MemoST.free;
  END;
end;

class function AdoCnx.GetConnectionString (ServerName,DbName : string) : string;
begin
  Result := 'Provider=SQLOLEDB.1'
          + ';Password=ADMIN'
          + ';Persist Security Info=True'
          + ';User ID=ADMIN'
          + ';Initial Catalog=' + DBName
          + ';Data Source=' + ServerName
          + ';Use Procedure for Prepare=1'
          + ';Auto Translate=True'
          + ';Packet Size=4096'
          + ';Workstation ID=LOCALST'
          + ';Use Encryption for Data=False'
          + ';Tag with column collation when possible=False'
          ;
end;

class function AdoCnx.PrepareSQl (SQL : string) : string;
begin
  Result := StringReplace(SQL,'"','''',[rfReplaceAll]);
end;


class function ShareDb.IsY2DB (DbName : string) : boolean;
begin
  Result := (Copy(DbName,Length(DbName)-1,2)='Y2');
end;


class function ShareDb.IsShareMaster (Application: TApplication; ServerName,DBName : string) : boolean;
var QQ : TADOQuery;
begin
  Result := false;
  QQ := TADOQuery.Create(application);
  TRY
    QQ.ConnectionString := AdoCnx.GetConnectionString (ServerName,DBName);
    QQ.SQL.Add(AdoCnx.PrepareSQl('SELECT 1 FROM YMULTIDOSSIER WHERE YMD_CODE="##MULTISOC" AND EXISTS(SELECT 1 FROM DESHARE) '));
    QQ.Prepared := true;
    TRY
      QQ.Open;
      Result := not QQ.eof;
      QQ.close;
    EXCEPT
      (*
      ON E:Exception do
      begin
        MessageBox(application.handle,PAnsiChar(E.Message),PAnsiChar(Application.Name),MB_OK);
      end;
      *)
    END;
  FINALLY
    QQ.free;
  END;
end;

class function ShareDb.ISDBPresent(ServerName, DbName: string; var ErrorString: string): boolean;
var CNX : TADOConnection;
begin
  Result := True;
  CNX := TADOConnection.Create(Application);
  TRY
    CNX.ConnectionString := AdoCnx.GetConnectionString(SERVERNAME,DBName);
    CNX.LoginPrompt := false;
    TRY
      CNX.Connected := True;
    EXCEPT
      on E: Exception do
      begin
        ErrorString := E.Message;
        result:= false;
      end;
    END;
  FINALLY
    CNX.free;
  END;
end;

class function ShareDb.GrpIsPresent(ServerName,DBName, Grp: string): Boolean;
var QQ : TADOQuery;
begin
  Result := false;
  QQ := TADOQuery.Create(application);
  TRY
    QQ.ConnectionString := AdoCnx.GetConnectionString (ServerName,DBName);
    QQ.SQL.Add(AdoCnx.PrepareSQl('SELECT 1 FROM USERGRP WHERE UG_GROUPE="'+Grp+'"'));
    QQ.Prepared := true;
    TRY
      QQ.Open;
      Result := not QQ.eof;
      QQ.close;
    EXCEPT
    END;
  FINALLY
    QQ.free;
  END;
end;

class function ShareDb.AppliqueDBPrinc(SERVERNAME,DbName: string;ListeTables,ListeGrps: TListDBStatus; LDossiers, LDatas: TStringList): boolean;
var CNX : TADOConnection;
    okok : boolean;
begin
  CNX := TADOConnection.Create(application);
  CNX.ConnectionString := AdoCnx.GetConnectionString (ServerName,DBName);
  TRY
    CNX.LoginPrompt := false;
    CNX.Connected := True;
    okok := ConstitueYmultiDossier (CNX,ListeTables,ListeGrps);
    if okok then okok := ConstitueDossier (CNX,LDossiers);
    if okok then okok := ConstituePartageDatas (CNX,DBName,LDatas);
    Result := okok;
  FINALLY
    CNX.free;
  end;
end;

class function ShareDb.AppliqueDBSecond(SERVERNAME,DBMaster,DbName: string; LDatas: TStringList): boolean;
var CNX : TADOConnection;
begin
  CNX := TADOConnection.Create(application);
  CNX.ConnectionString := AdoCnx.GetConnectionString (ServerName,DBName);
  TRY
    CNX.LoginPrompt := false;
    CNX.Connected := True;
    result := ConstituePartageDatas (CNX,DBMaster,LDatas);
  FINALLY
    CNX.free;
  end;
end;

class function AdoCnx.USDateTimeToStr(TheDate: TDateTime): string;
var YY,MM,DD,HH,MN,SS,ML : word;
begin
  DecodeDateTime(TheDate,YY,MM,DD,HH,MN,SS,ML);
  Result := Format('%.04d%.02d%.02d %.02d:%.02d:%.02d:%.03d',[YY,MM,DD,HH,MN,SS,ML]);
end;

end.
