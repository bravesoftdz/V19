unit tThreadTiers;

interface

uses
  Classes
  , UtilBTPVerdon
  , ConstServices
  , uTob
  , CbpMCD
  , AdoDB
  , CommonTools
  {$IFDEF MSWINDOWS}
  , Windows
  {$ENDIF}
  ;

type
  ThreadTiers = class(TThread)
  public
    lTn          : T_TablesName;
    TableValues  : T_TiersValues;
    LogValues    : T_WSLogValues;
    FolderValues : T_FolderValues;
    AdoQryBTP    : AdoQry;
    AdoQryTMP    : AdoQry;
    TobT         : TOB;
    TobAdd       : TOB;
    TobQry       : TOB;

    constructor Create(CreateSuspended : boolean);
    destructor Destroy; override;

  protected
    procedure Execute; override;
  end;

implementation

uses
  SysUtils
  , hCtrls
  , hEnt1
  , StrUtils
  , SvcMgr
  , ActiveX
  ;

{ Important : les m�thodes et propri�t�s des objets de la VCL peuvent uniquement �tre
  utilis�s dans une m�thode appel�e en utilisant Synchronize, comme : 

      Synchronize(UpdateCaption);

  o� UpdateCaption serait de la forme

    procedure ThreadTiers.UpdateCaption;
    begin
      Form1.Caption := 'Mis � jour dans un thread';
    end; }

{$IFDEF MSWINDOWS}
type
  TThreadNameInfo = record
    FType: LongWord;     // doit �tre 0x1000
    FName: PChar;        // pointeur sur le nom (dans l'espace d'adresse de l'utilisateur)
    FThreadID: LongWord; // ID de thread (-1=thread de l'appelant)
    FFlags: LongWord;    // r�serv� pour une future utilisation, doit �tre z�ro
  end;
{$ENDIF}

{ ThreadTiers }

constructor ThreadTiers.Create(CreateSuspended : boolean);
begin
  inherited Create(CreateSuspended);
end;

destructor ThreadTiers.Destroy;
begin
  inherited;
  TUtilBTPVerdon.AddLog(ServiceName_BTPVerdonExp, lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, False, TableValues.LastSynchro), LogValues, 0);
end;

procedure ThreadTiers.Execute;
var
  TobT           : TOB;
  TobAdd         : TOB;
  TobQry         : TOB;
  AdoQryBTP      : AdoQry;
  AdoQryTMP      : AdoQry;
  Treatment      : TExportTreatment;
begin
  Coinitialize(nil);
  try
    TobQry := TOB.Create('_QRY', nil, -1);
    try
      TobT := TOB.Create('_TABLE', nil, -1);
      try
        TobAdd := TOB.Create('_ADDFIEDS', nil, -1);
        try
          AdoQryBTP := AdoQry.Create;
          try
            AdoQryTMP := AdoQry.Create;
            try
              AdoQryBTP.Qry := TADOQuery.Create(nil);
              try
                AdoQryTMP.Qry := TADOQuery.Create(nil);
                try
                  Treatment := TExportTreatment.Create;
                  try
                    Treatment.AssignAdoQry(AdoQryBTP, AdoQryTMP, FolderValues, LogValues);
                    Treatment.TreatmentType := 'EXPORT';
                    Treatment.Tn            := tnTiers;
                    Treatment.FolderValues  := FolderValues;
                    Treatment.LogValues     := LogValues;
                    Treatment.LastSynchro   := TableValues.LastSynchro;
                    Treatment.AdoQryBTP     := AdoQryBTP;
                    Treatment.AdoQryTMP     := AdoQryTMP;
                    Treatment.ExportTreatment(TobT, TobAdd, TobQry);
                  finally
                    Treatment.Free;
                  end;
                finally
                  AdoQryTMP.Qry.Free;
                end;
              finally
                AdoQryBTP.Qry.Free;
              end;
            finally
              AdoQryTMP.free;
            end;
          finally
            AdoQryBTP.Free;
          end;
        finally
          FreeAndNil(TobAdd);
        end;
      finally
        FreeAndNil(TobT);
      end;
    finally
      FreeAndNil(TobQry);
    end;
  finally
    CoUninitialize();
  end;
end;

end.
