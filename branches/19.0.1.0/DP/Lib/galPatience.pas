unit galPatience;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, ComCtrls, HStatus, ExtCtrls;

type K2KThread = Class(TThread)
private
   ProgressBar : TProgressBar ;
   procedure Execute ; override ;
public
   constructor Create (CreateSuspended:Boolean ; PBar : TProgressBar);
end;

type TAfficheOu = (aoMilieu,aoCentreBas, aoCentreHaut) ;

type
  TFPatience = class(TForm)
    lAide: THLabel;
    lCreation: THLabel;
    lTitre: THLabel;
    sbBarre: THStatusBar;
    ProgressBar1: TProgressBar;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    Th : K2KThread ;
    procedure SetPos (AfficheOu : TAfficheOu) ;
    { Déclarations privées }
  public
    { Déclarations publiques }
    procedure SetTitre(texte: String);
    procedure Affiche(texte: String);
    procedure StartK2000 ;
    procedure StopK2000 ;
  end;

function FenetrePatience(titre : String ; AfficheOu : TAfficheOu = aoMilieu; ShowStatus : Boolean = True ; ShowIt : Boolean = True ): TFPatience;


/////////// IMPLEMENTATION //////////
implementation

{$R *.DFM}

function FenetrePatience(titre : String ; AfficheOu : TAfficheOu = aoMilieu; ShowStatus : Boolean = True ; ShowIt : Boolean = True ): TFPatience;
var F: TFPatience;
begin
  //$$$ JP 06/12/05 - warning delphi
  Result := Nil;
  F := TFPatience.Create(Application);
  if F <> nil then
  begin
       F.lTitre.caption := titre;
       F.SetPos(AfficheOu);
       F.sbBarre.Visible := ShowStatus ;
       if ShowIt then
          F.Show;
       Result := F;
  end;
end;

procedure TFPatience.Affiche(texte: String);
begin
  lAide.caption := texte;
  Self.Invalidate;
  Application.ProcessMessages;
end;

procedure TFPatience.SetTitre(texte: String);
begin
  lTitre.caption := texte;
  lTitre.Invalidate ;
end;

procedure TFPatience.StartK2000;
begin
  if not ProgressBar1.Visible then
     ProgressBar1.Visible := True ;
  Th.Resume ;
end;

procedure TFPatience.StopK2000;
begin
  if ProgressBar1.Visible then
     ProgressBar1.Visible := False ;
  Th.Terminate ;
end;


{ K2KThread }

constructor K2KThread.Create(CreateSuspended: Boolean; PBar: TProgressBar);
begin
  inherited Create(CreateSuspended);
  Priority := tpNormal	;
  FreeOnTerminate := False ;
  ProgressBar := PBar ;
end;

procedure K2KThread.Execute;
begin
  if ProgressBar = Nil then exit ;
  ProgressBar.Position := ProgressBar.Min ;
  while (not Terminated) do
  begin
    if ProgressBar.Position < ProgressBar.Max then
       ProgressBar.Position := ProgressBar.Position + 1
    else
       ProgressBar.Position := ProgressBar.Min ;
    ProgressBar.Invalidate ;
    Sleep(50);
  end;
  ProgressBar.Position := ProgressBar.Max ;
end;

procedure TFPatience.FormCreate(Sender: TObject);
begin
  Th := K2KThread.Create(True,ProgressBar1) ;
end;

procedure TFPatience.FormDestroy(Sender: TObject);
begin
  if not Th.Suspended then
     Th.Terminate ;
  Th.Free ;
end;

procedure TFPatience.SetPos(AfficheOu: TAfficheOu);
begin
  case AfficheOu of
  aoMilieu : Position := poMainFormCenter ;
  aoCentreBas :
  begin
     Position := poDesigned	 ;
     Top := ( Screen.Height div 2 ) + ( Screen.Height div 4 ) - ( Height div 2 ) ;
     Left := ( Screen.Width div 2 ) - ( Width div 2 ) ;
  end;
  aoCentreHaut :
  begin
     Position := poDesigned	;
     Top := ( Screen.Height div 4 ) - ( Height div 2 ) ;
     Left := ( Screen.Width div 2 ) - ( Width div 2 ) ;
  end;
  end;
end;

end.
