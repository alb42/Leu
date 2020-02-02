unit AskTextUnit;
{$mode objfpc}{$H+}
interface
uses
  Classes, SysUtils,
  MUIClass.Base,
  MUIClass.Area,
  MUIClass.Gadget,
  MUIClass.Group,
  MUIClass.Window;

type

  TAskTextWin = class(TMUIWindow)
  private
    FOnFinishTextEntry: TNotifyEvent;
    TopText: TMUIText;
    InputText: TMUIString;
    FIsOK: Boolean;
    procedure OKClicked(Sender: TObject);
    procedure CancelClicked(Sender: TObject);
    procedure DeactiveEvent(Sender: TObject);
    function GetValue: string;
  public
    constructor Create; override;

    function Execute(ATitle, ATopText, ADefaultText: string): Boolean;

    property OnFinishTextEntry: TNotifyEvent read FOnFinishTextEntry write FOnFinishTextEntry;
    property IsOK: Boolean read FIsOK;
    property Value: string read GetValue;
  end;
var
  AskTextWin: TAskTextWin;

implementation

constructor TAskTextWin.Create;
var
  Grp: TMUIGroup;
begin
  inherited;
  Horizontal := False;
  CloseGadget := False;
  OnDeactivate := @DeactiveEvent;
  Width := 300;

  TopText := TMUIText.Create('Enter a text: ');
  TopText.Parent := self;
  //
  InputText := TMUIString.Create;
  InputText.Parent := self;

  Grp := TMUIGroup.Create;
  Grp.Horiz := True;
  Grp.Parent := Self;

  with TMUIButton.Create('Ok') do
  begin
    Parent := Grp;
    OnClick := @OKClicked;
  end;

  TMUIRectangle.Create.Parent := Grp;

  with TMUIButton.Create('Cancel') do
  begin
    Parent := Grp;
    OnClick := @CancelClicked;
  end;
end;

procedure TAskTextWin.DeactiveEvent(Sender: TObject);
begin
  Activate := True;
end;

procedure TAskTextWin.OKClicked(Sender: TObject);
begin
  FIsOK := True;
  Close;
  if Assigned(OnFinishTextEntry) then
    OnFinishTextEntry(Self);
end;

procedure TAskTextWin.CancelClicked(Sender: TObject);
begin
  FIsOK := False;
  Close;
end;

function TAskTextWin.Execute(ATitle, ATopText, ADefaultText: string): Boolean;
begin
  FIsOK := False;
  Open := False;
  Title := ATitle;
  TopText.Contents := ATopText;
  InputText.Contents := ADefaultText;
  Show;
  repeat
    MUIApp.InputBuffered;
    SysUtils.Sleep(25);
  until Open = False;
  Result := FIsOK;
end;

function TAskTextWin.GetValue: string;
begin
  Result := InputText.Contents;
end;


end.
