program Leu;
{$mode objfpc}{$H+}
uses
  Types, Classes, SysUtils, intuition, agraphics, mui,
  Math,
  OfficeGridUnit,
  MUIClass.Base,
  MUIClass.Window,
  MUIClass.Grid,
  MUIClass.Gadget,
  MUIClass.Image,
  MUIClass.Group,
  MUIClass.Area,
  MUIClass.Menu,
  MUIClass.Dialog,
  MUIClass.DrawPanel,
  imagesunit, colorunit,
  fpstypes, fpspreadsheet, fpsutils, fpsnumformat,
  variants, SetSizeUnit;

type
  TMyWindow = class(TMUIWindow)
  private
    // Menu Entries
    procedure MenuNew(Sender: TObject); // New
    procedure MenuLoad(Sender: TObject); // Load
    procedure MenuSave(Sender: TObject); // Save
    procedure MenuSetSize(Sender: TObject); // SetSize
    procedure MenuQuit(Sender: TObject); // Quit
    procedure MenuCopy(Sender: TObject); // Copy
    procedure MenuCut(Sender: TObject); // Cut
    procedure MenuPaste(Sender: TObject); // Paste
    procedure MenuDelete(Sender: TObject); // Delete

    //
    procedure SetSizeEvent(Sender: TObject);
    procedure AutoColsRowsEvent(Sender: TObject);
    // Font Buttons
    procedure FontPropChanged(Sender: TObject);
    procedure TextAlignChanged(Sender: TObject);
    procedure BorderSelectChanged(Sender: TObject; Border: TDrawBorders);
    procedure BGColorClick(Sender: TObject);
    procedure FontColorClick(Sender: TObject);
    //
    procedure ListClick(Sender: TObject);
    procedure SetClick(Sender: TObject);
    procedure ChangeWorksheet(Sender: TObject);
    procedure DblClick(Sender: TObject);
  public
    BlockEvents: Boolean;
    WColRow: TMUIText;
    PrevSheet, NextSheet: TMUIImage;
    SheetName: TMUIText;
    WText: TMUIString;
    SG: TOfficeGrid;  // A StringGrid
    BoldBtn, ItalicBtn, ULineBtn: TMUIButton;
    LeftAlignBtn, CenterAlignBtn, RightAlignBtn: TMUIButton;
    function LoadFile(AFilename: string): Boolean;
    constructor Create; override;
  end;

var
  Win: TMyWindow;     // Window

procedure TMyWindow.MenuNew(Sender: TObject);
begin
  Unused(Sender);
  SG.NewWorkbook;
  Title := 'LEU <> Columns:' + IntToStr(SG.NumCols) + ' Rows:' + IntToStr(SG.NumRows);
  WColRow.Contents := '';
  WText.Contents := '';
  SheetName.Contents := IntToStr(SG.WorksheetIdx + 1) + '(' + SG.Worksheet.Name + ') /' + IntToStr(SG.WorksheetCount);
  PrevSheet.Disabled := SG.WorksheetIdx <= 0;
  NextSheet.Disabled := SG.WorksheetIdx >= SG.WorksheetCount;
end;

procedure TMyWindow.MenuLoad(Sender: TObject);
var
  FD: TFileDialog;
begin
  Unused(Sender);
  FD := TFileDialog.Create;
  FD.MultiSelect := False;
  FD.SaveMode := False;
  FD.TitleText := 'Select Spreadsheet file to load';
  FD.Pattern := '#?.(ods|xls|xlsx|csv|tcd|html|wikitable_pipes|wikitable_wikimedia)';
  if FD.Execute then
  begin
    try
      SG.LoadFile(FD.FileName);
    except
      on e: Exception do
      begin
        ShowMessage('Cannot load file: ' + ExtractFileName(FD.FileName) + #10 + E.Message);
        MenuNew(Self);
      end;
    end;
  end;
  FD.Free;
end;

procedure TMyWindow.MenuSave(Sender: TObject);
var
  FD: TFileDialog;
begin
  Unused(Sender);
  FD := TFileDialog.Create;
  FD.MultiSelect := False;
  FD.SaveMode := True;
  FD.TitleText := 'Select file name to save';
  FD.Pattern := '#?.(ods|xls|xlsx|csv|html|wikitable_pipes|wikitable_wikimedia)';
  if FD.Execute then
  begin
    try
      SG.SaveFile(FD.FileName);
      Title := 'LEU <'+ExtractFileName(FD.FileName)+'> Columns:' + IntToStr(SG.NumCols) + ' Rows:' + IntToStr(SG.NumRows);
      WColRow.Contents := '';
      WText.Contents := '';
      SheetName.Contents := IntToStr(SG.WorksheetIdx + 1) + '(' + SG.Worksheet.Name + ') /' + IntToStr(SG.WorksheetCount);
      PrevSheet.Disabled := SG.WorksheetIdx <= 0;
      NextSheet.Disabled := SG.WorksheetIdx >= SG.WorksheetCount;
    except
      on e: Exception do
      begin
        ShowMessage('Cannot save file: ' + ExtractFileName(FD.FileName) + #10 + E.Message);
      end;
    end;
  end;
  FD.Free;
end;

procedure TMyWindow.MenuSetSize(Sender: TObject);
begin
  Unused(Sender);
  SetSizeWin.OnAccept := @SetSizeEvent;
  SetSizeWin.OnAutoSize := @AutoColsRowsEvent;
  SetSizeWin.ShowWindow(SG.NumCols, SG.NumRows);
end;

procedure TMyWindow.MenuQuit(Sender: TObject);
begin
  Unused(Sender);
  Self.Close;
end;

procedure TMyWindow.MenuCopy(Sender: TObject); // Copy
begin
  Unused(Sender);
  SG.CopyToClip;
end;

procedure TMyWindow.MenuCut(Sender: TObject); // Cut
begin
  Unused(Sender);
  SG.CutToClip;
end;

procedure TMyWindow.MenuPaste(Sender: TObject); // Paste
begin
  Unused(Sender);
  SG.PasteFromClip;
end;

procedure TMyWindow.MenuDelete(Sender: TObject);
begin
  Unused(Sender);
  SG.DeleteSelectedCells;
end;

procedure TMyWindow.SetClick(Sender: TObject);
begin
  Unused(Sender);
  SG.Cells[SG.Col, SG.Row] := WText.Contents;
  SG.Row := SG.Row + 1;
  SG.SelectAll(False);
end;

function TMyWindow.LoadFile(AFileName: string): Boolean;
begin
  Result := True;
  SG.LoadFile(AFileName);
  Title := 'LEU <'+ExtractFileName(AFileName)+'> Columns:' + IntToStr(SG.NumCols) + ' Rows:' + IntToStr(SG.NumRows);
  WColRow.Contents := '';
  WText.Contents := '';
  SheetName.Contents := IntToStr(SG.WorksheetIdx + 1) + '(' + SG.Worksheet.Name + ') /' + IntToStr(SG.WorksheetCount);
  PrevSheet.Disabled := SG.WorksheetIdx <= 0;
  NextSheet.Disabled := SG.WorksheetIdx >= SG.WorksheetCount;
end;

procedure TMyWindow.ChangeWorksheet(Sender: TObject);
var
  Idx: Integer;
begin
  if TMUIArea(Sender).Disabled then
    Exit;
  Idx := SG.WorksheetIdx;
  if Sender = PrevSheet then
    Idx := SG.WorksheetIdx - 1;
  if Sender = NextSheet then
    Idx := SG.WorksheetIdx + 1;
  SG.LoadWorksheet(Idx);
  SheetName.Contents := IntToStr(SG.WorksheetIdx + 1) + '(' + SG.Worksheet.Name + ') /' + IntToStr(SG.WorksheetCount);
  PrevSheet.Disabled := SG.WorksheetIdx <= 0;
  NextSheet.Disabled := SG.WorksheetIdx >= SG.WorksheetCount;
  Title := 'LEU <'+ExtractFileName(SG.FileName)+'> Columns:' + IntToStr(SG.NumCols) + ' Rows:' + IntToStr(SG.NumRows);
end;

procedure TMyWindow.DblClick(Sender: TObject);
begin
  Unused(Sender);
  ActiveObject := WText;
end;

procedure TMyWindow.SetSizeEvent(Sender: TObject);
begin
  Unused(Sender);
  SG.SetSize(SetSizeWin.Columns, SetSizeWin.Rows);
end;

procedure TMyWindow.AutoColsRowsEvent(Sender: TObject);
var
  MaxR, MaxC: Integer;
  Cell: PCell;
begin
  Unused(Sender);
  MaxR := 2;
  MaxC := 2;
  for Cell in SG.Worksheet.Cells do
  begin
    MaxR := Max(Cell^.Row, MaxR);
    MaxC := Max(Cell^.Col, MaxC);
  end;
  SetSizeWin.Columns := MaxC + 2;
  SetSizeWin.Rows := MaxR + 2;
end;

//#############
// Lick click
procedure TMyWindow.ListClick(Sender: TObject);
var
  Style: TsFontStyles;
  ha: TsHorAlignment;
begin
  Unused(Sender);
  BlockEvents := True;
  try
    WColRow.Contents := GetColString(SG.Col - SG.FixedCols) + IntToStr(SG.Row);
    WText.Contents := SG.Cells[SG.Col, SG.Row];
    //
    // Update Font Parameter
    Style := SG.CellFontStyle[SG.Col, SG.Row];
    BoldBtn.Selected := fssBold in Style;
    ItalicBtn.Selected := fssItalic in Style;
    ULineBtn.Selected := fssUnderline in Style;
    // Update Text alignment
    ha := SG.HorAlignment[SG.Col, SG.Row];
    RightAlignBtn.Selected := ha = haRight;
    CenterAlignBtn.Selected := ha = haCenter;
    LeftAlignBtn.Selected := ha = haLeft;
  finally
    BlockEvents := False;
  end;
end;

//################
// Font Properties
procedure TMyWindow.FontPropChanged(Sender: TObject);
var
  Style: TsFontStyles;
  i: Integer;
  P: Types.TPoint;
begin
  Unused(Sender);
  if BlockEvents then
    Exit;
  // update the Font Parameter
  Style := [];
  if BoldBtn.Selected then
    Style := Style + [fssBold];
  if ItalicBtn.Selected then
    Style := Style + [fssItalic];
  if ULineBtn.Selected then
    Style := Style + [fssUnderline];
  SG.BeginUpdate;
  SG.CellFontStyle[SG.Col, SG.Row] := Style;
  SG.ReDrawCell(SG.Col, SG.Row);
  // redraw selection
  for i := 0 to SG.SelectionCount - 1 do
  begin
    P := SG.Selection[i];
    SG.CellFontStyle[P.X, P.Y] := Style;
    SG.ReDrawCell(P.X, P.Y);
  end;
  SG.EndUpdate;


end;

//################
// Text Alignment
procedure TMyWindow.TextAlignChanged(Sender: TObject);
var
  ha: TsHorAlignment;
  i: Integer;
  P: Types.TPoint;
begin
  if BlockEvents then
    Exit;
  BlockEvents := True;
  try
    if (Sender = LeftAlignBtn) and LeftAlignBtn.Selected then
    begin
      RightAlignBtn.Selected := False;
      CenterAlignBtn.Selected := False;
    end else
    if (Sender = CenterAlignBtn) and CenterAlignBtn.Selected then
    begin
      LeftAlignBtn.Selected := False;
      RightAlignBtn.Selected := False;
    end else
    if (Sender = RightAlignBtn) and RightAlignBtn.Selected then
    begin
      CenterAlignBtn.Selected := False;
      LeftAlignBtn.Selected := False;
    end;
    ha := haDefault;
    if LeftAlignBtn.Selected then
      ha := haLeft
    else
    if CenterAlignBtn.Selected then
      ha := haCenter
    else
    if RightAlignBtn.Selected then
      ha := haRight;
    // do not redraw
    SG.BeginUpdate;
    // assign focus
    SG.HorAlignment[SG.Col, SG.Row] := ha;
    SG.ReDrawCell(SG.Col, SG.Row);

    // assign selection
    for i := 0 to SG.SelectionCount - 1 do
    begin
      P := SG.Selection[i];
      SG.HorAlignment[P.X, P.Y] := ha;
      SG.ReDrawCell(P.X, P.Y);
    end;
    // redraw
    SG.EndUpdate;
  finally
    BlockEvents := False;
  end;
end;

procedure TMyWindow.BorderSelectChanged(Sender: TObject; Border: TDrawBorders);
var
  i: Integer;
  x, y, XMin, XMax, YMin, YMax: Integer;
  sBorder: TsCellBorders;
begin
  Unused(Sender);
  YMin := SG.Row;
  XMin := SG.Col;
  YMax := SG.Row;
  XMax := SG.Col;
  SG.BeginUpdate;
  try
    for i := 0 to SG.SelectionCount - 1 do
    begin
      YMin := Min(YMin, SG.Selection[i].Y);
      YMax := Max(YMax, SG.Selection[i].Y);
      XMin := Min(XMin, SG.Selection[i].X);
      XMax := Max(XMax, SG.Selection[i].X);
    end;
    for x := XMin to XMax do
    begin
      for y := YMin to YMax do
      begin
        sBorder := [];
        if dbLeft in Border then
        begin
          if x= xMin then
            Include(sBorder, cbWest);
        end;
        if dbRight in Border then
        begin
          if x= xMax then
            Include(sBorder, cbEast);
        end;
        if dbTop in Border then
        begin
          if y = YMin then
            Include(sBorder, cbNorth);
        end;
        if dbBottom in Border then
        begin
          if y = YMax then
            Include(sBorder, cbSouth);
        end;
        if dbVMid in Border then
        begin
          Include(sBorder, cbNorth);
          Include(sBorder, cbSouth);
        end;
        if dbHMid in Border then
        begin
          Include(sBorder, cbWest);
          Include(sBorder, cbEast);
        end;
        SG.SetCellBorder(x,y, sBorder);
      end;
    end;
  finally
    SG.EndUpdate;
  end;
end;

procedure TMyWindow.BGColorClick(Sender: TObject);
var
  i: Integer;
  Col: LongWord;
  BGButton: TColorButton;
begin
  if Sender is TColorButton then
  begin
    BGButton := TColorButton(Sender);
    Col := BGButton.Color;
    SG.BeginUpdate;
    try
      for i := 0 to SG.SelectionCount - 1 do
      begin
        SG.Worksheet.WriteBackgroundColor(SG.Selection[i].Y - SG.FixedRows, SG.Selection[i].X - SG.FixedCols, Col);
        SG.RedrawCell(SG.Selection[i].X, SG.Selection[i].Y);
      end;
    finally
      SG.EndUpdate;
    end;
  end;
end;

procedure TMyWindow.FontColorClick(Sender: TObject);
var
  i: Integer;
  Col: LongWord;
  FontButton: TColorButton;
begin
  if Sender is TColorButton then
  begin
    FontButton := TColorButton(Sender);
    Col := FontButton.Color;
    SG.BeginUpdate;
    try
      for i := 0 to SG.SelectionCount - 1 do
      begin
        SG.Worksheet.WriteFontColor(SG.Selection[i].Y - SG.FixedRows, SG.Selection[i].X - SG.FixedCols, Col);
        SG.RedrawCell(SG.Selection[i].X, SG.Selection[i].Y);
      end;
    finally
      SG.EndUpdate;
    end;
  end;
end;

constructor TMyWindow.Create;
var
  HeadGroup: TMUIGroup;
  FootGroup: TMUIGroup;
  ProjectMenu: TMUIMenu;
  EditMenu: TMUIMenu;
  Grp1: TMUIGroup;
begin
  inherited;
  BlockEvents := False;

  MenuStrip := TMUIMenuStrip.Create;

  ProjectMenu := TMUIMenu.Create;
  ProjectMenu.Title := 'Project';
  ProjectMenu.Parent := MenuStrip;

  with TMUIMenuItem.Create do
  begin
    Title := 'New...';
    ShortCut := 'N';
    OnTrigger := @MenuNew;
    Parent := ProjectMenu;
  end;

  with TMUIMenuItem.Create do
  begin
    Title := 'Load...';
    ShortCut := 'L';
    OnTrigger := @MenuLoad;
    Parent := ProjectMenu;
  end;

  with TMUIMenuItem.Create do
  begin
    Title := 'Save As...';
    ShortCut := 'S';
    OnTrigger := @MenuSave;
    Parent := ProjectMenu;
  end;

  with TMUIMenuItem.Create do
  begin
    Title := 'Set Size...';
    ShortCut := 'L';
    OnTrigger := @MenuSetSize;
    Parent := ProjectMenu;
  end;

  with TMUIMenuItem.Create do
  begin
    Title := 'Quit';
    ShortCut := 'Q';
    OnTrigger := @MenuQuit;
    Parent := ProjectMenu;
  end;

  EditMenu := TMUIMenu.Create;
  EditMenu.Title := 'Edit';
  EditMenu.Parent := MenuStrip;

  with TMUIMenuItem.Create do
  begin
    Title := 'Copy';
    ShortCut := 'C';
    OnTrigger := @MenuCopy;
    Parent := EditMenu;
  end;

  with TMUIMenuItem.Create do
  begin
    Title := 'Cut';
    ShortCut := 'X';
    OnTrigger := @MenuCut;
    Parent := EditMenu;
  end;

  with TMUIMenuItem.Create do
  begin
    Title := 'Paste';
    ShortCut := 'V';
    OnTrigger := @MenuPaste;
    Parent := EditMenu;
  end;

  with TMUIMenuItem.Create do
  begin
    Title := '-';
    Parent := EditMenu;
  end;

  with TMUIMenuItem.Create do
  begin
    Title := 'Erase';
    ShortCut := 'Delete';
    CommandString := True;
    OnTrigger := @MenuDelete;
    Parent := EditMenu;
  end;

  //

  HeadGroup := TMUIGroup.Create;
  HeadGroup.Horiz := True;
  HeadGroup.Frame := MUIV_Frame_None;
  HeadGroup.Parent := Self;

  WColRow := TMUIText.Create;
  WColRow.FixWidthTxt := 'Www1000';
  WColRow.Parent := HeadGroup;

  WText := TMUIString.Create;
  WText.CycleChain := 1;
  WText.Weight := 300;
  WText.OnAcknowledge := @SetClick;
  WText.Parent := HeadGroup;

  with TMUIButton.Create('Set') do
  begin
    FixWidthTxt := ' Set ';
    CycleChain := 0;
    OnClick := @SetClick;
    Parent := HeadGroup;
  end;

  // Font Buttons

  Grp1 := TMUIGroup.Create;
  with Grp1 do
  begin
    Horiz := True;
    FrameTitle := 'Font';
    Parent := HeadGroup;
  end;

  BoldBtn := TMUIButton.Create(MUIX_B + 'B');
  with BoldBtn do
  begin
    FixWidthTxt := ' B ';
    CycleChain := 0;
    InputMode := MUIV_InputMode_Toggle;
    OnSelected := @FontPropChanged;
    Parent := Grp1;
  end;

  ItalicBtn := TMUIButton.Create(MUIX_I + 'I');
  with ItalicBtn do
  begin
    FixWidthTxt := ' I ';
    CycleChain := 0;
    InputMode := MUIV_InputMode_Toggle;
    OnSelected := @FontPropChanged;
    Parent := Grp1;
  end;

  ULineBtn := TMUIButton.Create(MUIX_U + 'U');
  with ULineBtn do
  begin
    FixWidthTxt := ' U ';
    CycleChain := 0;
    InputMode := MUIV_InputMode_Toggle;
    OnSelected := @FontPropChanged;
    Parent := Grp1;
  end;

  with TColorButton.Create do
  begin
    Parent := Grp1;
    ColorImage.Height := 15;
    ColorImage.Width := 15;
    ColorImage.FixHeight := 15;
    ColorImage.FixWidth := 15;
    Title := MUIX_B + MUIX_C + 'Font color';
    OnColorClick := @FontColorClick;
  end;

  // Text Alignment stuff

  Grp1 := TMUIGroup.Create;
  with Grp1 do
  begin
    Horiz := True;
    FrameTitle := 'Text Alignment';
    Parent := HeadGroup;
  end;

  LeftAlignBtn := TMUIButton.Create('Left');
  with LeftAlignBtn do
  begin
    FixWidthTxt := 'Left';
    CycleChain := 0;
    InputMode := MUIV_InputMode_Toggle;
    OnSelected := @TextAlignChanged;
    Parent := Grp1;
  end;

  CenterAlignBtn := TMUIButton.Create('Center');
  with CenterAlignBtn do
  begin
    FixWidthTxt := 'Center';
    CycleChain := 0;
    InputMode := MUIV_InputMode_Toggle;
    OnSelected := @TextAlignChanged;
    Parent := Grp1;
  end;

  RightAlignBtn := TMUIButton.Create('Right');
  with RightAlignBtn do
  begin
    FixWidthTxt := 'Right';
    CycleChain := 0;
    InputMode := MUIV_InputMode_Toggle;
    OnSelected := @TextAlignChanged;
    Parent := Grp1;
  end;

  with TBorderButton.Create do
  begin
    Parent := HeadGroup;
    OnBorderSelect := @BorderSelectChanged;
  end;


  with TColorButton.Create do
  begin
    Parent := HeadGroup;
    Title := MUIX_B + MUIX_C + 'Background color';
    OnColorClick := @BGColorClick;
  end;

  SG := TOfficeGrid.Create;
  SG.DefCellWidth := 100;
  SG.NumCols := 12;
  SG.NumRows := 60;
  SG.OnDblClick := @DblClick;
  SG.OnCellFocus := @ListClick;
  SG.Frame := MUIV_Frame_None;
  SG.Parent := Self;

  FootGroup := TMUIGroup.Create;
  FootGroup.Frame := MUIV_Frame_None;
  FootGroup.Horiz := True;
  FootGroup.Parent := Self;

  PrevSheet := TMUIImage.Create;
  with PrevSheet do
  begin
    Frame := MUIV_Frame_Button;
    Font := MUIV_Font_Button;
    InputMode := MUIV_InputMode_RelVerify;
    FreeHoriz := False;
    Spec.SetStdPattern(MUII_ArrowLeft);
    OnClick := @ChangeWorksheet;
    Parent := FootGroup;
  end;

  SheetName := TMUIText.Create;
  SheetName.Parent := FootGroup;

  NextSheet := TMUIImage.Create;
  with nextSheet do
  begin
    Frame := MUIV_Frame_Button;
    Font := MUIV_Font_Button;
    InputMode := MUIV_InputMode_RelVerify;
    FreeHoriz := False;
    Spec.SetStdPattern(MUII_ArrowRight);
    OnClick := @ChangeWorksheet;
    Parent := FootGroup;
  end;

  TMUIRectangle.Create.Parent := FootGroup;

  Title := 'LEU <> Columns:' + IntToStr(SG.NumCols) + ' Rows:' + IntToStr(SG.NumRows);

  MenuNew(nil);
end;

procedure StartMe;
begin
  // Create a Window, with a title bar text
  Win := TMyWindow.Create;
  SetSizeWin := TSetSizeWin.Create;

  if ParamCount > 0 then
  begin
    try
      Win.LoadFile(ParamStr(1));
    except
      on E: Exception do
      begin
        ShowMessage('Cannot load file: ' + ExtractFileName(ParamStr(1)) + #10 + E.Message);
        Exit;
      end;
    end;
  end;

  MUIApp.Title := 'LEU';
  MUIApp.Version := '$VER: LEU 0.07 (05.07.2019)';
  MUIApp.Copyright := 'CC0';
  MUIApp.Author := 'Marcus "ALB42" Sackrow';
  MUIApp.Description := 'Simple Spreadsheet.';
  MUIApp.Base := 'LEU';

  MUIApp.Run;
end;

begin
  StartMe;
end.
