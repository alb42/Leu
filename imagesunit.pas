unit imagesunit;
{$mode objfpc}{$H+}
interface
uses
  Types, AGraphics, MUI,
  MUIClass.Base,
  MUIClass.DrawPanel, MUIClass.Area, MUIClass.Image, MUIClass.Window,
  MUIClass.Group;

const
  DistToBorder = 4;

type
  TDrawBorder = (dbLeft, dbRight, dbTop, dbBottom, dbHMid, dbVMid);
  TDrawBorders = set of TDrawBorder;

  TBorderSelectEvent = procedure(Sender: TObject; Border: TDrawBorders) of object;

  TBorderImage = class(TMUIBitmap)
  private
    FMyBorder: TDrawBorders;
    DB: TDrawBuffer;
  public
    destructor Destroy; override;
    procedure LoadImage(Num: Integer);
    property MyBorder: TDrawBorders read FMyBorder;
  end;

  TBorderButton = class(TBorderImage)
  private
    FOnBorderSelect: TBorderSelectEvent;
    BWin: TMUIWindow;
    Grp: TMUIGroup;
    Buttons: array[0..11] of TBorderImage;
    procedure ClickButton(Sender: TObject);
    procedure DeactiveEvent(Sender: TObject);
    procedure ClickBorder(Sender: TObject);
  public
    constructor Create; override;

    property OnBorderSelect: TBorderSelectEvent read FOnBorderSelect write FOnBorderSelect;
  end;

implementation

destructor TBorderImage.Destroy;
begin
  DB.Free;
  inherited;
end;

procedure TBorderImage.LoadImage(Num: Integer);
var
  Border: TDrawBorders;
begin
  InputMode := MUIV_InputMode_RelVerify;
  Frame := MUIV_FRAME_BUTTON;
  Background.SetStdPattern(MUII_BACKGROUND);
  Width := 24;
  Height := 24;
  FixWidth := Width;
  FixHeight := Height;
  DB := TDrawBuffer.Create(Width,Height, 2, nil);
  DB.Clear(2); // everything white
  Border := [];
  case Num of
    0: Border := [];
    1: Border := [dbLeft];
    2: Border := [dbRight];
    3: Border := [dbLeft, dbRight];
    4: Border := [dbTop];
    5: Border := [dbBottom];
    6: Border := [dbTop, dbBottom];
    7: Border := [dbLeft, dbRight, dbTop, dbBottom];
    8: Border := [dbTop, dbBottom, dbVMid];
    9: Border := [dbLeft, dbRight, dbTop, dbBottom, dbVMid];
    10: Border := [dbLeft, dbRight, dbTop, dbBottom, dbHMid];
    11: Border := [dbLeft, dbRight, dbTop, dbBottom, dbVMid, dbHMid];
  end;
  FMyBorder := Border;
  DB.APen := 1;
  if dbLeft in Border then
    DB.Line(DistToBorder, DistToBorder, DistToBorder, Height - DistToBorder);
  if dbRight in Border then
    DB.Line(Width - DistToBorder, DistToBorder, Width - DistToBorder, Height - DistToBorder);
  if dbTop in Border then
    DB.Line(DistToBorder, DistToBorder, Width - DistToBorder, DistToBorder);
  if dbBottom in Border then
    DB.Line(DistToBorder, Height - DistToBorder, Width - DistToBorder, Height - DistToBorder);
  if dbVMid in Border then
    DB.Line(DistToBorder, Height div 2, Width - DistToBorder, Height div 2);
  if dbHMid in Border then
    DB.Line(Width div 2, DistToBorder, Width div 2, Height - DistToBorder);
  Bitmap := DB.RP^.Bitmap;
end;

{ TBorderButton }
constructor TBorderButton.Create;
var
  i: Integer;
begin
  inherited;
  LoadImage(11);
  Transparent := 0;
  OnClick := @ClickButton;

  BWin := TMUIWindow.Create;
  BWin.BorderLess := True;
  //
  Grp := TMUIGroup.Create;
  Grp.Columns := 4;
  Grp.Parent := BWin;
  //
  for i := 0 to High(Buttons) do
  begin
    Buttons[i] := TBorderImage.Create;
    Buttons[i].LoadImage(i);
    Buttons[i].OnClick := @ClickBorder;
    Buttons[i].Parent := Grp;
  end;
end;

procedure TBorderButton.ClickButton(Sender: TObject);
var
  T: Types.TPoint;
  W: TMUIWindow;
begin
  // Calculate the window position
  T := Point(LeftEdge, TopEdge + Height);
  w := WindowObject;
  if Assigned(w) then
  begin
    T.X := T.X + W.LeftEdge;
    T.Y := T.Y + W.TopEdge;
  end;
  // recreate Window
  Grp.Parent := nil;
  BWin.Free;
  BWin := TMUIWindow.Create;
  BWin.BorderLess := True;
  BWin.CloseGadget := False;
  BWin.DragBar := False;
  BWin.DepthGadget := False;
  BWin.SizeGadget := False;
  BWin.LeftEdge := T.X;
  BWin.TopEdge := T.Y;

  BWin.OnDeactivate := @DeactiveEvent;

  Grp.Parent := BWin;

  BWin.Show;
end;

procedure TBorderButton.DeactiveEvent(Sender: TObject);
begin
  BWin.Close;
end;

procedure TBorderButton.ClickBorder(Sender: TObject);
begin
  if Assigned(Sender) and (Sender is TBorderImage) and Assigned(FOnBorderSelect) then
  begin
    FOnBorderSelect(Self, TBorderImage(Sender).MyBorder);
  end;
  BWin.Close;
end;

initialization

finalization

end.
