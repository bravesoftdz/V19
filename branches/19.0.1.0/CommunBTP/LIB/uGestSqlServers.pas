unit uGestSqlServers;
interface
uses
  Windows, Classes, SysUtils,
  ActiveX,
  ComObj,
  OleDB,
  DB,
  ADOInt,
  ADODB;
// broadcasted so that any single unit can expose method to handle enumerating
// their type of server

type TServerInfo101 = record
    platform_id: DWORD; 
    name: PWideChar; 
    version_major: DWORD; 
    version_minor: DWORD; 
    server_type: DWORD; 
    comment: PWideChar; 
  end; 
    PServerInfo101 = ^TServerInfo101; 
const 
  NERR_SUCCESS = 0; 
  MAX_PREFERRED_LENGTH = DWORD(-1); 
  SV_TYPE_WORKSTATION = $00000001; 
  SV_TYPE_SERVER = $00000002;
  SV_TYPE_SQLSERVER = $00000004;
  SV_TYPE_DOMAIN_CTRL = $00000008; 
  SV_TYPE_DOMAIN_BAKCTRL = $00000010; 
  SV_TYPE_TIME_SOURCE = $00000020; 
  SV_TYPE_AFP = $00000040; 
  SV_TYPE_NOVELL = $00000080; 
  SV_TYPE_DOMAIN_MEMBER = $00000100; 
  SV_TYPE_PRINTQ_SERVER = $00000200; 
  SV_TYPE_DIALIN_SERVER = $00000400; 
  SV_TYPE_XENIX_SERVER = $00000800; 
  SV_TYPE_SERVER_UNIX = SV_TYPE_XENIX_SERVER; 
  SV_TYPE_NT = $00001000; 
  SV_TYPE_WFW = $00002000; 
  SV_TYPE_SERVER_MFPN = $00004000; 
  SV_TYPE_SERVER_NT = $00008000; 
  SV_TYPE_POTENTIAL_BROWSER = $00010000; 
  SV_TYPE_BACKUP_BROWSER = $00020000; 
  SV_TYPE_MASTER_BROWSER = $00040000; 
  SV_TYPE_DOMAIN_MASTER = $00080000; 
  SV_TYPE_SERVER_OSF = $00100000; 
  SV_TYPE_SERVER_VMS = $00200000; 
  SV_TYPE_WINDOWS = $00400000; // Windows95 and above 
  SV_TYPE_DFS = $00800000; // Root of a DFS tree 
  SV_TYPE_CLUSTER_NT = $01000000; // NT Cluster 
  SV_TYPE_DCE = $10000000; // IBM DSS (Directory and Security Services) or equivalent
  SV_TYPE_ALTERNATE_XPORT = $20000000; // return list for alternate transport
  SV_TYPE_LOCAL_LIST_ONLY = $40000000; // Return local list only
  SV_TYPE_DOMAIN_ENUM = $80000000; 
  SV_TYPE_ALL = $FFFFFFFF; // handy for NetServerEnum2

function NetServerEnum(const ServerName: PWideString;
                       level: DWORD;
                       var Buffer: pointer;
                       PrefMaxLen: DWORD;
                       var EntriesRead: DWORD;
                       var TotalEntries: DWORD;
                       ServerType: DWORD;
                       const Domain: PWideChar;
                       var ResumeHandle: DWORD) : DWORD; stdcall; external 'netapi32.dll';

function NetApiBufferFree(Buffer: pointer): DWORD; stdcall; external 'netapi32.dll';
procedure GetSQLInstances (var TL : TStringList);
procedure GetServer( var TL : TStringList);

implementation

procedure GetServerNames(var TL : TStringList; const ServerType:DWORD);
var
   Buffer: pointer;
   EntriesRead,i,ErrCode,ResumeHandle,TotalEntries: DWORD;
   PDomainUnicode: PWideChar;
   ServerInfo: PServerInfo101;
begin
   ResumeHandle := 0;
   PDomainUnicode := nil;
   errCode := NetServerEnum(nil, 101, Buffer, MAX_PREFERRED_LENGTH,
                            EntriesRead,
                            TotalEntries, ServerType, PDomainUnicode, ResumeHandle);
   if (errCode <> NERR_SUCCESS) then Exit;
   try
      ServerInfo := Buffer;
      for i := 1 to EntriesRead do
      begin
         TL.Add(ServerInfo^.name);
         Inc(ServerInfo);
      end;
      if TL.Count = 0 then TL.Add('No servers available!');
   finally
      NetApiBufferFree(Buffer);
   end; // end of try finally
end;

procedure GetSQLInstances (var TL : TStringList);
var TT : TStringList;
    II : Integer;
begin
  TT := TStringList.Create;
  for II := 0 to TT.Count -1 do
  begin
  
  end;
end;



procedure ListAvailableSQLServers(var Names: TStringList);
var
  RSCon: ADORecordsetConstruction;
  Rowset: IRowset;
  SourcesRowset: ISourcesRowset;
  SourcesRecordset: _Recordset;
  SourcesName, SourcesType: TField;

  function PtCreateADOObject(const ClassID: TGUID): IUnknown;
  var
    Status: HResult;
    FPUControlWord: Word;
  begin
    asm
      FNSTCW FPUControlWord
    end;
    Status := CoCreateInstance(
                CLASS_Recordset,
                nil,
                CLSCTX_INPROC_SERVER or
                CLSCTX_LOCAL_SERVER,
                IUnknown,
                Result);
    asm
      FNCLEX
      FLDCW FPUControlWord
    end;
    OleCheck(Status);
  end;

begin
  SourcesRecordset :=
       PtCreateADOObject(CLASS_Recordset)
       as _Recordset;
  RSCon :=
       SourcesRecordset
       as ADORecordsetConstruction;
   SourcesRowset :=
       CreateComObject(ProgIDToClassID('SQLOLEDB Enumerator'))
       as ISourcesRowset;
   OleCheck(SourcesRowset.GetSourcesRowset(
            nil,
            IRowset, 0,
            nil,
            IUnknown(Rowset)));
   RSCon.Rowset := RowSet;
   with TADODataSet.Create(nil) do
   try
     Recordset := SourcesRecordset;
     SourcesName := FieldByName('SOURCES_NAME');
     SourcesType := FieldByName('SOURCES_TYPE');
     Names.BeginUpdate;
     Names.Clear;
     try
        while not EOF do
        begin
          if (SourcesType.AsInteger = DBSOURCETYPE_DATASOURCE) and
             (SourcesName.AsString <> '') then
            Names.Add(SourcesName.AsString);
          Next;
        end;
     finally
        Names.EndUpdate;
     end;
  finally
     Free;
  end;
end;

procedure GetServer( var TL : TStringList);
begin
  ListAvailableSQLServers(TL);
end;

end.
