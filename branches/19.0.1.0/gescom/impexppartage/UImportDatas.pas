unit UImportDatas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, StdCtrls, Mask, Hctrls, TntDialogs, TntStdCtrls,ADODB,
  ExtCtrls, TntExtCtrls, HPanel, ComCtrls, TntComCtrls, UShareDb, UdefImportDatas;

type
  TFImportDatas = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bDefaire: TToolbarButton97;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BImprimer: TToolbarButton97;
    HPanel1: THPanel;
    HLabel1: THLabel;
    BConnect: TToolbarButton97;
    SERVERNAME: TEdit;
    PGCONTROL: THPageControl2;
    TSDB: TTabSheet;
    TSGRP: TTabSheet;
    LbSaveFile: THLabel;
    SaveFileName: THCritMaskEdit;
    BLanceExport: TToolbarButton97;
    BLANCE: TToolbarButton97;
    LBDB: THListView;
    LBGRPS: THListView;
    procedure BConnectClick(Sender: TObject);
    procedure SaveFileNameElipsisClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BLanceExportClick(Sender: TObject);
    procedure BLANCEClick(Sender: TObject);
  private
    { Déclarations privées }
    CNX : TADOConnection;
    LDOSSIER,LDATA : TStringList;
    TLS,TGRPS : TListDBStatus;
    procedure ConnecteServeur;
    procedure ChargeLesDonnees;
    procedure DecodeDb (Datas : WideString);
    procedure Decodegrps(Datas : WideString);
    procedure DecodeDossier(Datas : WideString);
    procedure DecodeDatas (Datas : widestring);
    procedure PositionneLesPartages;
  public
    { Déclarations publiques }
  end;


implementation
uses ed_tools, DB, Ulog;

{$R *.dfm}


procedure TFImportDatas.BConnectClick(Sender: TObject);
begin
  if SERVERNAME.Text = '' then exit;
  TRY
    ConnecteServeur;
    MessageBox(application.handle,'Connection au serveur effectuée',PAnsiChar(Application.MainForm.Caption),MB_OK);
  EXCEPT
  END;
end;

procedure TFImportDatas.SaveFileNameElipsisClick(Sender: TObject);
var TT : TOpenDialog;
begin

  //ouverture du sélecteur de fichier windows dans le répertoire des modèles...
	TT := TOpenDialog.Create(application);
  TRY
    TT.DefaultExt := '.lss';
    TT.Filter := 'Fichiers sauvegarde (*.lss)|*.lss';
    if TT.Execute then
    begin
      SaveFileName.Text :=  TT.FileName;
    end;
  FINALLY
  	TT.Free;
  end;
end;

procedure TFImportDatas.FormCreate(Sender: TObject);
begin
  TGRPS := TListDBStatus.Create;
  TLS := TListDBStatus.Create;
  CNX := TADOConnection.Create(application);
  LDOSSIER := TStringList.Create;
  LDATA := TStringList.Create;
end;

procedure TFImportDatas.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  TGRPS.Free;
  TLS.Free;
  CNX.free;
  LDOSSIER.free;
  LDATA.Free;
end;

procedure TFImportDatas.BLanceExportClick(Sender: TObject);
begin
  TRY
    TLS.clear;
    if not CNX.Connected then
    begin
      ConnecteServeur;
    end;
    LDOSSIER.Clear;
    LDATA.Clear;
    ChargeLesDonnees;
  EXCEPT
    raise;
  END;
end;

procedure TFImportDatas.ConnecteServeur;
begin
  if SERVERNAME.text = '' then exit;
  CNX.ConnectionString := AdoCnx.GetConnectionString(SERVERNAME.text,'master');
  TRY
    CNX.LoginPrompt := false;
    TRY
      CNX.Connected := True;
    EXCEPT
      on E: Exception do
      begin
        MessageBox(application.handle,PAnsiChar(E.Message),PAnsiChar(Application.MainForm.Caption),MB_OK);
        raise;
      end;
    END;
  finally
    CNX.connected := false;
  end;
end;

procedure TFImportDatas.ChargeLesDonnees;
var f : TextFile;
    Datas : Widestring;
    fOk,ffirst  : boolean;
    xx : string;
begin
  if not FileExists(SaveFileName.Text) then Exit;
  //
  ffirst := True;
  TRY
    InitMoveProgressForm(nil, 'Traitements des données', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
    TRY
      //
      MoveCurProgressForm('Récupération des données');
      //
      AssignFile(f, SaveFileName.Text);
      TRY
        Reset(f);
        while not Eof(f) do
        begin
          readln ( f, Datas);
          if ffirst then
          begin
            if Copy(Datas,1,2)<> 'DB' then
            begin
              Break;
            end;
            ffirst := false;
          end;
          xx := READTOKENPipe (Datas,':');
          if Copy(xx,1,2)='DB' then
          begin
            DecodeDb(Datas);
          end else if Copy(xx,1,4)='GRPS' then
          begin
            Decodegrps(Datas);
          end else if Copy(xx,1,7)='DOSSIER' then
          begin
            DecodeDossier(Datas);
          end else if Copy(xx,1,4)='DATA' then
          begin
            DecodeDatas(Datas);
          end;
        end;
      FINALLY
        CloseFile(f);
      end;
      //
      MoveCurProgressForm('Vérification des données');
      //
    FINALLY
      FiniMoveProgressForm();
      PGCONTROL.ActivePageIndex := 0;
      fOk := (LBDB.Items.Count>0) and (LBGRPS.Items.count >0) and (LDOSSIER.Count >0) and (LDATA.count > 0);
      if not fok then
      begin
        MessageBox(application.handle,PAnsiChar('Informations de partage non utilisable'),PAnsiChar(Application.MainForm.Caption),MB_OK);
      end;
    end;
  EXCEPT
    On E:Exception do
    begin
      MessageBox(application.handle,PAnsiChar(E.Message),PAnsiChar(Application.MainForm.Caption),MB_OK);
    end;
  END;
end;

procedure TFImportDatas.DecodeDb(Datas: WideString);
var DB : string;
    Error : string;
    ListItem : TListItem;
    TS : TStatus;
begin
  TLS.clear;
  LBDB.Clear;
  repeat
    DB := READTOKENPipe(Datas,'|');
    if DB <> '' then
    begin
      DB := DB+'_Y2';
      TS := TStatus.Create;
      TS.DB := DB;
      LBDB.Items.BeginUpdate;
      try
        with LBDB do
          begin
            ListItem := LBDB.Items.Add;
            Listitem.Caption := DB;
            if ShareDb.ISDBPresent (SERVERNAME.Text,DB,Error) then
            begin
              Listitem.SubItems.Add  ('Présente');
              TS.Status := True;
            end else
            begin
              Listitem.SubItems.add  ('Absente');
              TS.Status := false;
            end;
          end;
      finally
        TLS.Add(TS);
        LBDB.Items.EndUpdate;
      end;
    end;
  until db = '';
end;

procedure TFImportDatas.Decodegrps(Datas: WideString);
var GRP : string;
    TS : TStatus;
    ListItem : TListItem;
    DBName : string;
begin
  TGRPS.clear;
  LBGRPS.Items.BeginUpdate;
  DBName :='';
  if TLS.Count > 0 then DBName := TStatus(TLS.items[0]).DB;
  repeat
    GRP := READTOKENPipe(Datas,'|');
    if GRP <> '' then
    begin
      TS := TStatus.Create;
      ListItem := LBGRPS.Items.Add;
      ListItem.Caption := GRP;
      TS.DB := GRP;
      if (ShareDb.GrpIsPresent (SERVERNAME.Text,DBName,GRP)) and (DbName <> '') then
      begin
        TS.Status := True;
        Listitem.SubItems.Add  ('Présent');
      end else
      begin
        Listitem.SubItems.Add  ('Absent');
        TS.Status := false;
      end;
      TGRPS.Add(TS); 
    end;
  until GRP = '';
  LBGRPS.Items.EndUpdate;
end;

procedure TFImportDatas.DecodeDossier(Datas: WideString);
begin
  LDOSSIER.Add(Datas);
end;

procedure TFImportDatas.DecodeDatas(Datas: widestring);
begin
  LDATA.Add(Datas);
end;

procedure TFImportDatas.BLANCEClick(Sender: TObject);
var II : Integer;
    TS : TStatus;
    fOK : Boolean;
begin
//
  fOk := True;
  for II := 0 to TLS.Count-1 do
  begin
    TS := TStatus(TLS.Items[II]);
    if TS.Status = false then
    begin
      fOK := false;
      break;
    end;
  end;
  if fok then
  begin
    for II := 0 to TGRPS.Count-1 do
    begin
      TS := TStatus(TGRPS.Items[II]);
      if TS.Status = false then
      begin
        fOK := false;
        break;
      end;
    end;
  end;
  if not Fok then
  begin
    MessageBox(application.handle,PAnsiChar('Les éléments présents ne permettent pas de définir le partage'),PAnsiChar(Application.MainForm.Caption),MB_OK);
    exit;
  end;
  PositionneLesPartages;
end;

procedure TFImportDatas.PositionneLesPartages;
var II : Integer;
    okok : boolean;
begin
  okok := false;
  for II := 0 to LBDB.Items.Count -1 do
  begin
    if II = 0 then
    begin
      // positionnement sur base Maitre
      okok := ShareDb.AppliqueDBPrinc (SERVERNAME.Text,LBDB.Items[0].Caption,TLS,TGRPS,LDOSSIER,LDATA);
    end else
    begin
      if okok then okok := ShareDb.AppliqueDBSecond (SERVERNAME.Text,LBDB.Items[0].caption,LBDB.Items[II].caption,LDATA);
    end;
  end;
  if okok then
  begin
    MessageBox(application.handle,PAnsiChar('Importation des partages dans les bases effectué'),PAnsiChar(Application.MainForm.Caption),MB_OK);
  end else
  begin
    MessageBox(application.handle,PAnsiChar('Importation des partages non effectué di fait d''erreur en écriture'),PAnsiChar(Application.MainForm.Caption),MB_OK);
  end;
end;

end.
