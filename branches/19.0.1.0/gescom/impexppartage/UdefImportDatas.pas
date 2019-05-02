unit UdefImportDatas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, StdCtrls, Mask, Hctrls, TntDialogs, TntStdCtrls,ADODB,
  ExtCtrls, TntExtCtrls, HPanel, ComCtrls, TntComCtrls;

type
  TStatus = class (TObject)
  private
    fDB : string;
    fstatus : boolean;
  public
    property DB : string read fDB write fDB;
    property Status : boolean read fstatus write fstatus;
  end;

  TListDBStatus = class(TList)
  private
    function GetItems(Indice : integer): TStatus;
    procedure SetItems(Indice : integer; const Value: TStatus);
    function Add(AObject: TStatus): Integer;
  public
    destructor destroy; override;
    property Items [Indice : integer] : TStatus read GetItems write SetItems;
    procedure clear; override;
  end;


implementation
uses ed_tools, DB, Ulog;

{ TListDBStatus }

function TListDBStatus.Add(AObject: TStatus): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TListDBStatus.clear;
var indice : integer;
begin
  for Indice := 0 to Count -1 do
  begin
    TStatus(Items [Indice]).free;
  end;
  inherited;
end;

destructor TListDBStatus.destroy;
begin
  clear;
  inherited;
end;

function TListDBStatus.GetItems(Indice: integer): TStatus;
begin
 result := TStatus (Inherited Items[Indice]);
end;

procedure TListDBStatus.SetItems(Indice: integer; const Value: TStatus);
begin
  Inherited Items[Indice]:= Value;
end;
end.
