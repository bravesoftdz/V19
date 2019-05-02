unit UPrincipale;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, ExtCtrls, TntExtCtrls, HPanel, StdCtrls, TntStdCtrls,
  Hctrls;

type
  TFprincipale = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    HLabel3: THLabel;
    HLabel4: THLabel;
    ToolbarButton972: TToolbarButton97;
    procedure BFermeClick(Sender: TObject);
    procedure BExportClick(Sender: TObject);
    procedure BImportClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  Fprincipale: TFprincipale;

implementation
uses UpartageExport,UimportDatas;

{$R *.dfm}

procedure TFprincipale.BFermeClick(Sender: TObject);
begin
  close;
end;

procedure TFprincipale.BExportClick(Sender: TObject);
var XX: TFExportData;
begin
  //
  XX := TFExportData.create (application);
  XX.ShowModal;
  XX.free;
end;

procedure TFprincipale.BImportClick(Sender: TObject);
var XX: TFImportDatas;
begin
  XX := TFImportDatas.create (application);
  XX.ShowModal;
  XX.free;
//
end;

end.
