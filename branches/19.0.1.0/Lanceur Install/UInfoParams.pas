unit UInfoParams;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TInfoParams = class(TForm)
    Label1: TLabel;
    TEMPPATH: TLabel;
    Label2: TLabel;
    CURPATH: TLabel;
    Label4: TLabel;
    PARAMS: TLabel;
    DRIVETYPE: TLabel;
    Label3: TLabel;
    DRIVE: TLabel;
    Label6: TLabel;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.dfm}

end.
