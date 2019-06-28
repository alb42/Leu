unit ColorUnit;
{$mode objfpc}{$H+}
interface
uses
  Types, AGraphics, MUI, intuition, Classes, utility, exec, Math,
  MUIClass.Base,
  MUIClass.DrawPanel, MUIClass.Area, MUIClass.Image, MUIClass.Window,
  MUIClass.Group;
const
  ImagesSize = 10;

type
  TColorImage = class(TMUIBitmap)
  private
    FImagesSize: Integer;
    DB: TDrawBuffer;
    APen: LongInt;
    ColorMode: Boolean;
  public
    constructor Create(Size: Integer); virtual reintroduce;
    destructor Destroy; override;
    procedure SetColorAsPen(Pen: Integer);
    procedure SetColor(AColor: LongWord);
    property Pen: Integer read APen write SetColorAsPen;
  end;

  TColorButton = class(TMUIGroup)
  private
    ArrowButton: TMUIImage;

    NameTag: TMUIText;
    BWin: TMUIWindow;
    HiddenWin: TMuiWindow;
    Grp: TMUIGroup;
    FTitle: string;

    FOnColorClick: TNotifyEvent;
    //
    ColChooseWin: TMUIWindow;
    CA: TMUIColorAdjust;
    //
    Buttons: array of TColorImage;
    Last: array[0..11] of TColorImage;
    procedure ClickButton(Sender: TObject);
    procedure DeactiveEvent(Sender: TObject);
    procedure SetTitle(ATitle: string);
    procedure ClickColor(Sender: TObject);
    procedure ClickLastColor(Sender: TObject);
    procedure ChooseCol(Sender: TObject);
    procedure DeactiveColEvent(Sender: TObject);
    procedure ColChooseOKClick(Sender: TObject);
    function GetColor: LongWord;
  public
    ColorImage: TColorImage;
    constructor Create; override;
    property Title: string read FTitle write SetTitle;
    property OnColorClick: TNotifyEvent read FOnColorClick write FOnColorClick;
    property Color: LongWord read GetColor;
  end;

implementation


function GetAmigaRed(AColor: LongWord): LongWord; Inline;
begin
  Result := ((AColor and $0000FF) shl 24);
end;

function GetAmigaGreen(AColor: LongWord): LongWord; Inline;
begin
  Result := ((AColor and $00FF00) shl 16);
end;

function GetAmigaBlue(AColor: LongWord): LongWord; Inline;
begin
  Result := ((AColor and $FF0000) shl 8);
end;

function GetFromAmigaColor(Red, Green, Blue: LongWord): LongWord; inline;
begin
  Result := (Red and $FF000000) shr 24 or
            (Green and $FF000000) shr 16 or
            (Blue and $FF000000) shr 8;
end;


constructor TColorImage.Create(Size: Integer);
begin
  inherited Create;
  FImagesSize := Size;
  //
  InputMode := MUIV_InputMode_RelVerify;
  //
  Width := FImagesSize;
  Height := FImagesSize;
  FixWidth := FImagesSize;
  FixHeight := FImagesSize;
  DB := TDrawBuffer.Create(FImagesSize, FImagesSize, 8, nil);
  DB.Clear(2);
  Bitmap := DB.RP^.Bitmap;
  ColorMode := False;
end;

destructor TColorImage.Destroy;
begin
  if ColorMode and (Pen >= 0) then
    ReleasePen(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, Pen);
  DB.Free;
  inherited;
end;

procedure TColorImage.SetColorAsPen(Pen: Integer);
begin
  if ColorMode and (Pen >= 0) then
    ReleasePen(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, Pen);
  ColorMode := False;
  APen := Pen;
  if Pen >= 0 then
    DB.Clear(Pen)
  else
  begin
    DB.Clear(0);
    DB.APen := 1;
    DB.Line(0,0, Width, Height);
    DB.Line(Width,0, 0, Height);
  end;
end;

procedure TColorImage.SetColor(AColor: LongWord);
begin
  if ColorMode and (Pen >= 0) then
    ReleasePen(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, Pen);
  if AColor and $FF000000 <> 0 then
  begin
    ColorMode := False;
    Pen := -1;
  end
  else
  begin
    ColorMode := True;
    Pen := ObtainBestPenA(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, GetAmigaRed(AColor), GetAmigaGreen(AColor), GetAmigaBlue(AColor), nil);
  end;
  if Pen >= 0 then
    DB.Clear(Pen)
  else
  begin
    DB.Clear(0);
    DB.APen := 1;
    DB.Line(0,0, Width, Height);
    DB.Line(Width,0, 0, Height);
  end;
end;



{ TColorButton }

constructor TColorButton.Create;
var
  i: Integer;
  Grp2, Grp3: TMUIGroup;
  Depth: Integer;
begin
  inherited;
  Frame := MUIV_FRAME_NONE;
  Horiz := True;
  ColorImage := TColorImage.Create(24);
  With ColorImage do
  begin
    Frame := MUIV_FRAME_BUTTON;
    Transparent := 0;
    Parent := Self;
    OnClick := @ClickLastColor;
  end;

  ArrowButton := TMUIImage.Create;
  with ArrowButton do
  begin
    FixHeight := 24;
    InputMode := MUIV_InputMode_RelVerify;
    Frame := MUIV_FRAME_BUTTON;
    OnClick := @ClickButton;
    Spec.Spec := MUII_ArrowDown;
    Parent := Self;
  end;

  BWin := TMUIWindow.Create;
  BWin.BorderLess := True;

  HiddenWin := TMUIWindow.Create;
  HiddenWin.BorderLess := True;
  //
  Grp := TMUIGroup.Create;
  Grp.Parent := BWin;

  FTitle := MUIX_B + 'Color';
  NameTag := TMUIText.Create(FTitle);
  NameTag.Parent := Grp;

  Grp2 := TMUIGroup.Create;
  Grp2.Parent := Grp;
  Grp2.Columns := 16;

  Depth := Min(8, IntuitionBase^.ActiveScreen^.Bitmap.Depth);

  Depth := 2 ** Depth;
  SetLength(Buttons, Depth);

  //
  for i := 0 to High(Buttons) do
  begin
    Buttons[i] := TColorImage.Create(ImagesSize);
    Buttons[i].Pen := i - 1;
    Buttons[i].OnClick := @ClickColor;
    Buttons[i].Parent := Grp2;
  end;

  Grp3 := TMUIGroup.Create;
  Grp3.Horiz := True;
  Grp3.Parent := Grp;

  TMUIText.Create('Last:').Parent := Grp3;
  //
  for i := 0 to High(Last) do
  begin
    Last[i] := TColorImage.Create(ImagesSize);
    Last[i].SetColorAsPen(1);
    Last[i].OnClick := @ClickLastColor;
    Last[i].Parent := Grp3;
  end;

  with TMUIButton.Create('Choose Color') do
  begin
    OnClick := @ChooseCol;
    Parent := Grp;
  end;

  // color choose window
  ColChooseWin := TMUIWindow.Create;
  ColChooseWin.OnDeactivate := @DeactiveColEvent;
  ColChooseWin.LeftEdge := MUIV_Window_LeftEdge_Moused;
  ColChooseWin.TopEdge := MUIV_Window_TopEdge_Moused;
  ColChooseWin.Height := MUIV_Window_Height_Visible(30);
  ColChooseWin.Width := MUIV_Window_Width_Visible(30);
  CA := TMUIColorAdjust.Create;
  Ca.Parent := ColChooseWin;

  Grp2 :=TMUIGroup.Create;
  Grp2.Horiz := True;
  Grp2.Parent := ColChooseWin;

  with TMUIButton.Create('Ok') do
  begin
    OnClick := @ColChooseOKClick;
    Parent := Grp2;
  end;

  with TMUIButton.Create('Cancel') do
  begin
    OnClick := @DeactiveColEvent;
    Parent := Grp2;
  end;

end;

procedure TColorButton.ClickButton(Sender: TObject);
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
  Grp.Parent := HiddenWin;
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

procedure TColorButton.DeactiveEvent(Sender: TObject);
begin
  BWin.Close;
end;

procedure TColorButton.DeactiveColEvent(Sender: TObject);
begin
  ColChooseWin.Close;
end;

procedure TColorButton.SetTitle(ATitle: string);
begin
  if FTitle = ATitle then
    Exit;
  FTitle := ATitle;
  NameTag.Contents := ATitle;
end;

procedure TColorButton.ClickColor(Sender: TObject);
var
  i: Integer;
begin
  if Last[0].Pen <> TColorImage(Sender).Pen then
  begin
    // Put Last
    for i := High(Last) downto 1 do
    begin
      Last[i].Pen := Last[i - 1].Pen;
    end;
    Last[0].Pen := TColorImage(Sender).Pen;
  end;
  //
  ClickLastColor(Sender);
end;

procedure TColorButton.ClickLastColor(Sender: TObject);
begin
  if ColorImage.Pen <> TColorImage(Sender).Pen then
  begin
    ColorImage.Pen := TColorImage(Sender).Pen;
    if HasObj then
      MUI_Redraw(MUIObj, MADF_DRAWOBJECT);
  end;
  BWin.Close;
  if Assigned(FOnColorClick) then
    FOnColorClick(Self);
end;

function TColorButton.GetColor: LongWord;
var
  col: array[0..2] of LongWord;
begin
  if ColorImage.Pen < 0 then
    Result := $80000000
  else
  begin
    {$ifdef MorphOS}
    GetRGB32(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, ColorImage.Pen, 1, Col[0]);
    {$else}
    GetRGB32(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, ColorImage.Pen, 1, @Col[0]);
    {$endif}
    Result := GetFromAmigaColor(Col[0], Col[1], Col[2]);
  end;
end;

procedure TColorButton.ChooseCol(Sender: TObject);
var
  Col: LongWord;
begin

  if ColChooseWin.Open then
    ColChooseWin.Open := False;
  Col := Self.GetColor;
  CA.Red := GetAmigaRed(Col);
  CA.Green := GetAmigaGreen(Col);
  CA.Blue := GetAmigaBlue(Col);
  //
  ColChooseWin.Open := True;
end;

procedure TColorButton.ColChooseOKClick(Sender: TObject);
var
  i: Integer;
begin
  Self.ColorImage.SetColor(GetFromAmigaColor(CA.Red, CA.Green, CA.Blue));
  if Last[0].Pen <> Self.ColorImage.Pen then
  begin
    // Put Last
    for i := High(Last) downto 1 do
    begin
      Last[i].Pen := Last[i - 1].Pen;
    end;
    Last[0].Pen := Self.ColorImage.Pen;
  end;
  if HasObj then
    MUI_Redraw(MUIObj, MADF_DRAWOBJECT);
  ColChooseWin.Close;
  if Assigned(FOnColorClick) then
    FOnColorClick(Self);
end;

end.
