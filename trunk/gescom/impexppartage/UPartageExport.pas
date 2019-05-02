unit UPartageExport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, StdCtrls, Mask, Hctrls, TntDialogs, TntStdCtrls,ADODB,
  ExtCtrls, TntExtCtrls, HPanel, ComCtrls, TntComCtrls, UShareDb;


type
  TFExportData = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    HPanel1: THPanel;
    LNomDB: THLabel;
    HPanel4: THPanel;
    SaveFileName: THCritMaskEdit;
    LbSaveFile: THLabel;
    BLanceExport: TToolbarButton97;
    HLabel1: THLabel;
    SERVERNAME: TEdit;
    BConnect: TToolbarButton97;
    PGCONTROL: THPageControl2;
    TSDB: TTabSheet;
    TSGRP: TTabSheet;
    LBDB: TListBox;
    TSDATAS: TTabSheet;
    LBGRPS: TListBox;
    LBDATAS: TListBox;
    DBNAME: THValComboBox;
    procedure BFermeClick(Sender: TObject);
    procedure SaveFileNameElipsisClick(Sender: TObject);
    procedure BLanceExportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BConnectClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    DATABASES,GRPS : string;
    LSHARES,LBDOSSIER : TstringList;
    procedure SelectDb(Sender : TObject);
  public
    { Déclarations publiques }
  end;

implementation
uses ed_tools, DB, Ulog;

{$R *.dfm}

procedure TFExportData.BFermeClick(Sender: TObject);
begin
  close;
end;

procedure TFExportData.SaveFileNameElipsisClick(Sender: TObject);
var SD : TSaveDialog;
begin
  SD := TSaveDialog.Create(Application);
  SD.DefaultExt := '*.lss';
  SD.Title := 'Exportation de fichiers';
  SD.Filter := 'Fichiers sauvegarde (*.lss)|*.lss';
  if SD.execute then SaveFileName.Text := SD.FileName;
  SD.Free;
end;

procedure TFExportData.BLanceExportClick(Sender: TObject);
var II : Integer;
begin
  if SaveFileName.Text = '' then
  begin
    MessageBox(Application.Handle,PAnsiCHar('Vous devez renseigner le nom du fichier de sauvegarde'),PAnsiChar(Application.MainForm.Caption),MB_OK);
    Exit;
  end;
  if FileExists(SaveFileName.text) then
  begin
    if MessageDlg('Le fichier existe déjà. Confirmez-vous son ecrasement ?',mtWarning,[MByes,MBNO],0 )<> mryes then
    begin
      Exit;
    end;
    DeleteFile(SaveFileName.text);
  end;
  //
  TRY
    InitMoveProgressForm(nil, 'Ecriture des données', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);

    TRY
      MoveCurProgressForm('Ecriture DB ');
      for II := 0 to LBDB.Count -1 do
      begin
        if II=0 then DATABASES := 'DB:' else DATABASES := DATABASES+'|';
        DATABASES := DATABASES+ LBDB.Items [II];
      end;
      EcritDatas(SaveFileName.text,DATABASES);
      //
      MoveCurProgressForm('Ecriture GRPS');
      for II := 0 to LBGRPS.Count -1 do
      begin
        if II=0 then GRPS := 'GRPS:' else GRPS := GRPS+'|';
        GRPS := GRPS+ LBGRPS.Items [II];
      end;
      EcritDatas(SaveFileName.text,GRPS);
      //
      MoveCurProgressForm('Ecriture DOSSIER');
      for II := 0 to LBDOSSIER.count -1 do
      begin
        EcritDatas(SaveFileName.text,LBDOSSIER.Strings[II]);
      end;
      //
      MoveCurProgressForm('Ecriture SHARES');
      for II := 0 to LSHARES.count -1 do
      begin
        EcritDatas(SaveFileName.text,LSHARES.Strings[II]);
      end;
      //
      MessageBox(Application.Handle,PAnsiCHar('Les éléments de partages sont sauvegardés'),PAnsiChar(Application.MainForm.Caption),MB_OK);
    EXCEPT
      ON E:Exception do
      begin
        MessageBox(Application.Handle,PAnsiCHar(E.Message),PAnsiChar(Application.MainForm.Caption),MB_OK);
      end;
    END;
  FINALLY
    FiniMoveProgressForm();
    close;
  END;
end;

procedure TFExportData.FormShow(Sender: TObject);
begin
  SERVERNAME.Text := '';
end;

procedure TFExportData.BConnectClick(Sender: TObject);

var QQ : TADOQuery;
    CNX : TADOConnection;
begin
  DBNAME.Clear;
  if SERVERNAME.Text = '' then exit;

  CNX := TADOConnection.Create(application);
  TRY
    CNX.ConnectionString := AdoCnx.GetConnectionString(SERVERNAME.text,'master');
    CNX.LoginPrompt := false;
    TRY
      CNX.Connected := True;
      QQ := TADOQuery.Create(application);
      TRY
        QQ.Connection := CNX;
        QQ.SQL.Add('SELECT name, database_id FROM sys.databases where database_id > 4');
        QQ.Prepared := true;
        QQ.Open;
        InitMoveProgressForm(nil, 'Analyse des Bases de données', 'Veuillez patienter SVP ...', QQ.RecordCount, FALSE, TRUE);
        QQ.First;
        TRY
          while not QQ.Eof do
          begin
            MoveCurProgressForm('Analyse de la base ' + QQ.fields[0].AsString);
            if Not ShareDb.IsY2DB(QQ.fields[0].AsString) then
            begin
              if ShareDb.IsShareMaster(Application,SERVERNAME.text,QQ.fields[0].AsString) then DBName.AddItem(QQ.fields[0].AsString,nil);
            end;
            QQ.next;
          end;
        FINALLY
          FiniMoveProgressForm();
          QQ.Close;
        end;
      finally
        QQ.Free;
      end;
    EXCEPT
    //
    END;
  FINALLY
    CNX.free;
    if DBNAME.Items.Count > 0 then DBNAME.OnChange  := SelectDB;
  END;
end;

procedure TFExportData.SelectDb(Sender: TObject);
var QQ : TADOQuery;
    CNX : TADOConnection;
    STL : TStringList;
    DBList,GrpList : string;
    TL,ST,LogicName,realName : string;
    NameDB : string;
    Stype,SData : string; 
begin
  //
  InitMoveProgressForm(nil, 'Analyse des des données de partages', 'Veuillez patienter SVP ...', 3, FALSE, TRUE);
  NameDB := DbName.Text ;
  CNX := TADOConnection.Create(application);
  TRY
    CNX.ConnectionString := AdoCnx.GetConnectionString(SERVERNAME.text,NameDB);
    CNX.LoginPrompt := false;
    TRY
      CNX.Connected := True;
      QQ := TADOQuery.Create(application);
      TRY
        MoveCurProgressForm('Récupération des bases et des groupes du partage ');
        QQ.Connection := CNX;
        QQ.SQL.Add(AdoCnx.PrepareSQl('SELECT YMD_DETAILS FROM YMULTIDOSSIER WHERE YMD_CODE="##MULTISOC"'));
        QQ.Prepared := true;
        TRY
          QQ.Open;
        EXCEPT
          On E : Exception do
          begin
            MessageBox(application.handle,PAnsiChar(E.Message),PAnsiChar(Application.MainForm.Caption),MB_OK);
          end;
        end;
        QQ.First;
        STL := TStringList.Create ;
        TRY
          STL.Text := QQ.fields[0].AsString;
          DBList       := STL.Strings[0] ;
          GRPList      := STL.Strings[1] ;
          //
          LBDB.Clear;
          repeat
            TL := READTOKENST(DBLIST);
            if TL <> '' then
            begin
              ST := TL;
              LogicName := READTOKENPipe(ST,'|');
              RealName := READTOKENPipe(ST,'|');
              LBDB.AddItem(RealName,nil);
            end;
          until TL='' ;
          //
          LBGRPS.Clear;
          repeat
            ST := READTOKENST(GRPList);
            if ST <> '' then
            begin
              LBGRPS.AddItem(ST,nil);
            end;
          until ST='' ;
          //
        FINALLY
          STL.Free;
        end;
        QQ.Close;
        //
        MoveCurProgressForm('Récupération des données du partage ');
        QQ.SQL.Clear;
        QQ.SQL.Add(AdoCnx.PrepareSQl('SELECT DS_NOMTABLE,DS_MODEFONC,DS_NOMBASE,DS_TYPTABLE,DS_VUE FROM DESHARE'));
        QQ.Prepared := true;
        TRY
          QQ.Open;
        EXCEPT
          On E : Exception do
          begin
            MessageBox(application.handle,PAnsiChar(E.Message),PAnsiChar(Application.MainForm.Caption),MB_OK);
          end;
        end;
        QQ.First;
        LBDATAS.Clear;
        while not QQ.Eof do
        begin
          LSHARES.Add(Format('DATA:%s|%s|%s|%s|%s',[QQ.fields[0].AsString,QQ.fields[1].AsString,QQ.fields[2].AsString,QQ.fields[3].AsString,QQ.fields[4].AsString]));
          //
          Stype := QQ.fields[3].AsString;
          if Stype = 'TTE' then SData := 'Tablette'
          else if Stype = 'TAB' then SData := 'Table'
          else if sType = 'PAR' then SData := 'Paramsoc';

          LBDATAS.AddItem (Format('%s : %s',[sData,QQ.fields[0].AsString]),nil);
          QQ.Next;
        end;
        //
        MoveCurProgressForm('Récupération des données sociétés ');
        QQ.SQL.Clear;
        QQ.SQL.Add(AdoCnx.PrepareSQl('SELECT DOS_NOMBASE,DOS_NODOSSIER,DOS_SOCIETE,DOS_LIBELLE,DOS_UTILISATEUR,DOS_GUIDDOSSIER,DOS_TYPEDOSSIER FROM DOSSIER'));
        QQ.Prepared := true;
        TRY
          QQ.Open;
        EXCEPT
          On E : Exception do
          begin
            MessageBox(application.handle,PAnsiChar(E.Message),PAnsiChar(Application.MainForm.Caption),MB_OK);
          end;
        end;
        QQ.First;
        LBDOSSIER.Clear;
        while not QQ.Eof do
        begin
          LBDOSSIER.Add(Format('DOSSIER:%s|%s|%s|%s|%s|%s|%s',[QQ.fields[0].AsString,QQ.fields[1].AsString,QQ.fields[2].AsString,QQ.fields[3].AsString,QQ.fields[4].AsString,QQ.fields[5].AsString,QQ.fields[6].AsString]));
          QQ.Next;
        end;
        //
      finally
        QQ.Free;
      end;
    EXCEPT
      MessageBox(application.handle,'Traitement annulé',PAnsiChar(Application.MainForm.Caption),MB_OK);
    END;
  FINALLY
    CNX.free;
    FiniMoveProgressForm();
    PGCONTROL.ActivePageIndex := 0;
  END;

end;

procedure TFExportData.FormCreate(Sender: TObject);
begin
  LSHARES := TStringList.Create;
  LBDOSSIER := TStringList.create;
end;

procedure TFExportData.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  LSHARES.free;
  LBDOSSIER.free;
end;

end.


