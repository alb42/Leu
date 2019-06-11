program Leu;
{$mode objfpc}{$H+}
uses
  Types, Classes, SysUtils, intuition, agraphics, exec, utility, mui,
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
  fpstypes, fpspreadsheet, fpsallformats, fpsutils, fpsnumformat, variants;

type
  TMyWindow = class(TMUIWindow)
  private
    // Menu Entries
    procedure MenuNew(Sender: TObject); // New
    procedure MenuLoad(Sender: TObject); // Load
    procedure MenuSave(Sender: TObject); // Save
    procedure MenuQuit(Sender: TObject); // Quit
    // Font Buttons
    procedure FontPropChanged(Sender: TObject);
    procedure TextAlignChanged(Sender: TObject);
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

procedure TMyWindow.MenuQuit(Sender: TObject);
begin
  Self.Close;
end;

procedure TMyWindow.SetClick(Sender: TObject);
begin
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
  ActiveObject := WText;
end;

//#############
// Lick click
procedure TMyWindow.ListClick(Sender: TObject);
var
  Style: TsFontStyles;
  ha: TsHorAlignment;
begin
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
begin
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
  SG.CellFontStyle[SG.Col, SG.Row] := Style;
  SG.ReDrawCell(SG.Col, SG.Row);
end;

//################
// Text Alignment
procedure TMyWindow.TextAlignChanged(Sender: TObject);
var
  ha: TsHorAlignment;
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
    SG.HorAlignment[SG.Col, SG.Row] := ha;
    SG.ReDrawCell(SG.Col, SG.Row);
  finally
    BlockEvents := False;
  end;
end;

constructor TMyWindow.Create;
var
  HeadGroup: TMUIGroup;
  FootGroup: TMUIGroup;
  ProjectMenu: TMUIMenu;
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
    Title := 'Quit';
    ShortCut := 'Q';
    OnTrigger := @MenuQuit;
    Parent := ProjectMenu;
  end;

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
  MUIApp.Version := '$VER: LEU 0.05 (09.06.2019)';
  MUIApp.Copyright := 'CC0';
  MUIApp.Author := 'Marcus "ALB42" Sackrow';
  MUIApp.Description := 'Simple Spreadsheet.';
  MUIApp.Base := 'LEU';

  MUIApp.Run;
end;

begin
  StartMe;
end.
