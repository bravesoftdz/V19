unit tWaitingMessage;

interface

uses
  Classes
  , Forms
  , StdCtrls
  , ExtCtrls
  , ComCtrls
  , Dialogs
  ;

type
  thWaitingMessage = class(TThread)
  private

  protected
    procedure Execute; override;

  public
    Caption         : string;
    Msg             : string;
    WaitMsgForm     : TForm;
    WaitLabel       : TLabel;
    WaitProgressBar : TProgressBar;

    constructor Create(CreateSuspended : boolean; pCaption, pMessage : string);
    destructor Destroy; override;
  end;

implementation

uses
  CommonTools
  , SysUtils
  , ActiveX
  ;

{ WaintingMessage }

constructor thWaitingMessage.Create(CreateSuspended : boolean; pCaption, pMessage : string);
begin
  inherited Create(CreateSuspended);
  Caption         := pCaption;
  Msg             := pMessage;
  Priority        := tpNormal;
  FreeOnTerminate := False;
  { Création TForm }
  WaitMsgForm              := CreateMessageDialog(PChar(Msg), mtCustom, []);
  WaitMsgForm.BorderIcons  := [];
  WaitMsgForm.Caption      := Caption;
  WaitMsgForm.BorderStyle  := bsDialog;
  WaitMsgForm.Width        := Length(Msg)*10;
  { Création du message }
  WaitLabel                := TLabel.Create(WaitMsgForm);
  WaitLabel.Caption        := '';
  WaitLabel.Width          := (WaitMsgForm.Width div 10)*8;
  WaitLabel.Top            := (WaitMsgForm.ClientHeight div 2)-(WaitLabel.Height div 2);
  WaitLabel.Left           := (WaitMsgForm.ClientWidth div 2)-(WaitLabel.Width div 2);
  { Création du ProgressBar }
  WaitProgressBar          := TProgressBar.Create(WaitMsgForm);
  WaitProgressBar.Parent   := WaitMsgForm;
  WaitProgressBar.Width    :=(WaitMsgForm.Width div 10)*8;
  WaitProgressBar.Top      :=(WaitMsgForm.ClientHeight div 2)-(WaitProgressBar.Height div 2);
  WaitProgressBar.Left     :=(WaitMsgForm.ClientWidth div 2)-(WaitProgressBar.Width div 2);
  WaitProgressBar.Max      := 50;
  WaitMsgForm.Show;
  WaitMsgForm.Repaint;
end;

destructor thWaitingMessage.Destroy;
begin
  inherited;
  Terminate;
  FreeAndNil(WaitLabel);
  FreeAndNil(WaitProgressBar);
  WaitMsgForm.Hide();
  WaitMsgForm.Release();
  FreeAndNil(WaitMsgForm);
end;

procedure thWaitingMessage.Execute;
begin
  while WaitProgressBar.Position < WaitProgressBar.Max do
  begin
    WaitProgressBar.Position := Tools.iif(WaitProgressBar.Position = WaitProgressBar.Max-1, 0, WaitProgressBar.Position + 1);
    Sleep(100);
    if Terminated then break;
  end;
end;

end.
