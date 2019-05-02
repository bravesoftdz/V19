unit UNetAPI32;

interface

uses
  Windows, SysUtils;

  function GetCurrentUser(): String;
  function GetDomainServerName(): String;
  function GetDomainFullName(ServerName, UserName: String): String;

implementation

  function NetUserGetInfo(ServerName, UserName: PWideChar; Level: DWORD; var Buffer: Pointer): DWORD; stdcall; external 'netapi32.dll' name 'NetUserGetInfo';
  function NetApiBufferFree(Buffer: pointer): DWORD; stdcall; external 'netapi32.dll' name 'NetApiBufferFree';
  function NetWkstaUserGetInfo(ServerName: PWideChar; Level: DWORD; var Buffer: Pointer): Longint; stdcall; external 'netapi32.dll' name 'NetWkstaUserGetInfo';

type
  TUserInfo1 = packed record
    UserName: PWideChar;
    DomainName : PWideChar;
    OtherDomainNames: PWideChar;
    ServerName: PWideChar;
  end;
  PUserInfo1 = ^TUserInfo1;

  TUserInfo2 = packed record
    Name: PWideChar;
    Password: PWideChar;
    PasswordAge: DWORD;
    Priv: DWORD;
    HomeDir: PWideChar;
    Comment: PWideChar;
    Flags: DWORD;
    ScriptPath: PWideChar;
    AuthorFlags: DWORD;
    FullName: PWideChar;
    UserComment: PWideChar;
    Params: PWideChar;
    WorkStations: PWideChar;
    LastLogon: DWORD;
    LastLogoff: DWORD;
    AccountExpires: DWORD;
    MaxStorage: DWORD;
    UnitsPerWeek: DWORD;
    LogonHours: DWORD;
    BadPasswordCount: DWORD;
    LogonCount: DWORD;
    Server: PWideChar;
    CountryCode: DWORD;
    Codepage: DWORD;
  end;
  PUserInfo2 = ^TUserInfo2;

function GetCurrentUser(): String;
var
  username: String;
  size: DWORD;
begin
  size := 255;
  SetLength(username, size) ;
  if GetUserName(PChar(username), size) then
    Result := Copy(username, 1, size - 1)
  else
    Result := '';
end;

function GetDomainServerName(): String;
var
  PUI1: PUserInfo1;
begin
  Result := '';
  if NetWkstaUserGetInfo(nil, 1, Pointer(PUI1)) = 0 then
  begin
    try
      Result := WideCharToString(PUI1^.ServerName);
    finally
      NetApiBufferFree(PUI1);
    end;
  end;
end;

function GetDomainFullName(ServerName, UserName: String): String;
var
  PUI2: PUserInfo2;
begin
  Result := '';
  if NetUserGetInfo(PWideChar(WideString(ServerName)), PWideChar(WideString(UserName)), 2, Pointer(PUI2)) = 0 then
    try
      Result := WideString(PUI2^.FullName);
    finally
      NetApiBufferFree(PUI2);
    end;
end;

end.
