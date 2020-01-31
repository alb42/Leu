unit FormatWinUnit;
{$mode objfpc}{$H+}
interface
uses
  Classes, SysUtils,
  locale, intuition, agraphics, MUI,
  fgl,
  fpstypes, fpsnumformat,
  typinfo, Math,
  MUIClass.Area,
  MUIClass.StringGrid,
  MUIClass.Window,
  MUIClass.Gadget,
  MUIClass.Group,
  MUIClass.Image,
  MUIClass.DrawPanel,
  OfficeGridUnit,
  Types;

const
  NumFormats = 9;
  FormatTypes: array[0..NumFormats - 1] of TsNumberFormat = (nfFixed, nfPercentage, nfCurrency, nfLongDate, nfLongTime, nfExp, nfFraction, nfCustom, nfText);
  FormatStrings: array[0..NumFormats - 1] of string = ('Fixed', 'Percentage', 'Currency', 'Date', 'Time', 'Exponential', 'Fractional', 'Custom', 'Text');


  FixedFormats: string = ''#10'0'#10'0.00'#10'#,##0'#10'#,##0.00'#10'#,###.00'#10'#,##0_);(#,##0)'#10'#,##0.00_);(#,##0.00)';
  PercentFormats: string = '0%'#10'0.00%';
  MoneyFormats: string = '#,##0 [$-407];-#,##0 [$-407]'#10'#,##0.00 [$-407];-#,##0.00 [$-407]'#10 +
                         '#,##0 [$-407];[RED]-#,##0 [$-407]'#10'#,##0.00 [$-407];[RED]-#,##0.00 [$-407]'#10 +
                         '#,##0.-- [$-407];[RED]-#,##0.-- [$-407]'#10'#,##0.00 [$];-#,##0.00 [$]'#10 +
                         '#,##0.00 [$];[RED]-#,##0.00 [$]';
  DateFormats: string = 'DD/MM/YY'#10'DD.MM.YY'#10'MM/DD/YY'#10 +
                        'DD/MM/YYYY'#10'DD.MM.YYYY'#10'MM/DD/YYYY'#10 +
                        'YY-MM-DD'#10'YYYY-MM-DD'#10 +
                        'D MMM YY'#10'DD MMM YYYY'#10'D MMMM YYYY'#10 +
                        'NN D MMM YY'#10'NNNND MMMM YYYY'#10'YYYY-MM-DD HH:MM:SS';
  TimeFormats: string = 'HH:MM'#10'HH:MM:SS'#10'HH:MM AM/PM'#10'HH:MM:SS AM/PM';
  ExpFormats: string = '0.00E+000'#10'0.00E+00'#10'0.0000E+00'#10'##0.00E+00';
  FracFormats: string = '# ?/?'#10'# ??/??'#10'# ???/???'#10'# ?/2'#10'# ?/4'#10'# ?/10'#10'# ??/100'#10'# ???/1000';
  BoolFormats: string = ' ';
  TextFormats: string = '@';

type
  TFormatDesc = class
  public
    Name: string;
    Fmt: TsNumberFormat;
    FormatCodes: TStringList;
    constructor Create;
    destructor Destroy; override;
  end;

  //TFormatSpec = specialize TFPGObjectList<TFormatDesc>;

  //TAllFormats = class


type
  TFormatWin = class(TMUIWindow)
  private
    FormatList: TMUiStringGrid;
    SubFormatList: TMUiStringGrid;
    FormatEdit: TMUIString;
    SubFormats: TStringList;
    OKButton: TMUIButton;
    CancelButton: TMUIButton;
    Preview: TMUIBitmap;
    DB: TDrawBuffer;

    function GetExampleText(FmtStr: string; Dummy: Boolean): string;
    procedure ClickEvent(Sender: TObject);
    procedure ClickFormatEvent(Sender: TObject);
    procedure AckEdit(Sender: TObject);
    procedure OKClick(Sender: TObject);
    procedure CancelClick(Sender: TObject);
  public
    SG: TOfficeGrid;
    Cell: TPoint;
    OrgValue: Double;
    OrgText: string;
    SelectedCells: TsCellRangeArray;
    NewFmt: TsNumberFormat;
    constructor Create; override;
    destructor Destroy; override;

    function Execute: Boolean;
  end;

var
  FormatWin: TFormatWin;
implementation

uses
  fpsutils;

var
  DefMoneySymb: string = '$';

constructor TFormatWin.Create;
var
  i: Integer;
  Grp: TMUIGroup;
begin
  inherited;

  Horizontal := False;

  Grp := TMUIGroup.Create;
  Grp.Horiz := True;
  Grp.Parent := Self;

  FormatList := TMUiStringGrid.Create;
  FormatList.NumColumns := 1;
  FormatList.ShowLines := False;
  FormatList.ShowTitle := False;
  FormatList.Parent := Grp;
  FormatList.OnClick := @ClickEvent;
  FormatList.OnDoubleClick := @ClickEvent;

  FormatList.Quiet := True;
  FormatList.NumRows := Length(FormatStrings);
  for i := 0 to High(FormatStrings) do
  begin
    FormatList.Cells[0, i] := FormatStrings[i];
  end;
  FormatList.Quiet := False;

  SubFormatList := TMUiStringGrid.Create;
  SubFormatList.NumColumns := 1;
  SubFormatList.ShowLines := False;
  SubFormatList.ShowTitle := False;
  SubFormatList.Parent := Grp;
  SubFormatList.OnClick := @ClickFormatEvent;
  SubFormatList.OnDoubleClick := @ClickFormatEvent;

  SubFormats := TStringList.Create;

  Grp := TMUIGroup.Create;
  Grp.Horiz := True;
  Grp.Parent := Self;

  with TMUIRectangle.Create do
   Parent := Grp;

  Preview := TMUIBitmap.Create;
  With Preview do
  begin
    FixWidth := 200;
    FixHeight := 20;
    Width := FixWidth;
    Height := FixHeight;
    DB := TDrawBuffer.Create(FixWidth, FixHeight, 8, nil);
    DB.Clear(2);
    Bitmap := DB.RP^.Bitmap;
    Parent := Grp;
  end;

  with TMUIRectangle.Create do
   Parent := Grp;

  Grp := TMUIGroup.Create;
  Grp.Horiz := True;
  Grp.Parent := Self;

  FormatEdit := TMUIString.Create;
  FormatEdit.OnAcknowledge := @AckEdit;
  FormatEdit.Parent := Grp;

  Grp := TMUIGroup.Create;
  Grp.Horiz := True;
  Grp.Parent := Self;

  OKButton := TMUIButton.Create('OK');
  OKButton.OnClick := @OKClick;
  OKButton.Parent := Grp;

  with TMUIRectangle.Create do
   Parent := Grp;

  CancelButton := TMUIButton.Create('Cancel');
  CancelButton.OnClick := @CancelClick;
  CancelButton.Parent := Grp;

  {SL:= TStringList.Create;
  BuildCurrencyFormatList(SL, False, 12.34, DefMoneySymb);
  writeln('SL: ', SL.Text);
  SL.Free;}
  //writeln('BuildCurrencyFormatString: ', BuildCurrencyFormatString( ));
end;

destructor TFormatWin.Destroy;
begin
  SubFormats.Free;
  DB.Free;
  inherited;
end;

function TFormatWin.Execute: Boolean;
var
  i: Integer;
  fmt: TsCellFormat;
begin
  Cell := Point(SG.Col, SG.Row);
  SetLength(SelectedCells, 1);
  with SG do
  begin
    SelectedCells[0].Col1 := Col - FixedCols;
    SelectedCells[0].Col2 := Col - FixedCols;
    SelectedCells[0].Row1 := Row - FixedRows;
    SelectedCells[0].Row2 := Row - FixedRows;
    for i := 0 to SelectionCount - 1 do
    begin
      SelectedCells[0].Col1 := Min(SelectedCells[0].Col1, Selection[i].X - FixedCols);
      SelectedCells[0].Col2 := Max(SelectedCells[0].Col2, Selection[i].X - FixedCols);
      SelectedCells[0].Row1 := Min(SelectedCells[0].Row1, Selection[i].Y - FixedRows);
      SelectedCells[0].Row2 := Max(SelectedCells[0].Row2, Selection[i].Y - FixedRows);
    end;
    Worksheet.SetSelection(SelectedCells);
  end;
  if (SelectedCells[0].Col1 = SelectedCells[0].Col2) and (SelectedCells[0].Row1 = SelectedCells[0].Row2) then
    Title := 'Format for ' + GetColString(Cell.X - SG.FixedCols) + IntToStr(Cell.Y)
  else
    Title := 'Format for ' + GetColString(SelectedCells[0].Col1 - SG.FixedCols) + IntToStr(SelectedCells[0].Row1) + ':' + GetColString(SelectedCells[0].Col2 - SG.FixedCols) + IntToStr(SelectedCells[0].Row2);
  OrgValue := SG.Worksheet.ReadAsNumber(SG.Row - 1, SG.Col - 1);
  OrgText := SG.Worksheet.ReadAsText(SG.Row - 1, SG.Col - 1);
  try
    fmt := SG.Workbook.GetCellFormat(SG.Worksheet.GetCell(SG.Row - 1, SG.Col - 1)^.FormatIndex);
    case fmt.NumberFormat of
      nfFixed, nfFixedTh, nfGeneral: FormatList.Row := 0;
      nfPercentage: FormatList.Row := 1;
      nfCurrency, nfCurrencyRed: FormatList.Row := 2;
      nfExp: FormatList.Row := 5;
      nfShortDateTime, nfShortDate, nfLongDate, nfDayMonth, nfMonthYear: FormatList.Row := 3;
      nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM, nfTimeInterval: FormatList.Row := 4;
      nfFraction: FormatList.Row := 6;
      nfText: FormatList.Row := 8;
      else
        FormatList.Row := 7;
    end;
    ClickEvent(FormatList);
    FormatEdit.Contents := fmt.NumberFormatStr;
  except
  end;
  Result := True;
  Show;
end;


function TFormatWin.GetExampleText(FmtStr: string; Dummy: Boolean): string;
var
  parser: TsNumFormatParser;
  mClass: TsNumFormatParams;
  i: Integer;
begin
  Result := '';
  parser := TsNumFormatParser.Create(FmtStr, SG.Workbook.FormatSettings);
  try
    mClass := TsNumFormatParams.Create;
    SetLength(mClass.Sections, parser.ParsedSectionCount);
    for i:=0 to High(mClass.Sections) do
      mClass.Sections[i] := parser.ParsedSections[i];
    if Dummy then
    begin
      if Parser.IsDateTimeFormat then
        Result := ConvertFloatToStr(Now(), mClass, SG.Workbook.FormatSettings)
      else
        Result := ConvertFloatToStr(-123.456, mClass, SG.Workbook.FormatSettings);
      if Pos('RED', FmtStr) > 0 then
        Result := #27'b' +  Result;
    end
    else
    begin
      if isNan(OrgValue) then
        Result := OrgText
      else
        Result := ConvertFloatToStr(OrgValue, mClass, SG.Workbook.FormatSettings);
    end;
  finally
    parser.Free;
    mClass.Free;
  end;
end;

procedure TFormatWin.ClickEvent(Sender: TObject);
var
  i: Integer;
begin
  if (FormatList.Row < 0) or (FormatList.Row > High(FormatTypes)) then
    Exit;
  NewFmt := FormatTypes[FormatList.Row];
  SubFormats.Clear;
  case NewFmt of
    nfFixed: SubFormats.Text := FixedFormats;
    nfPercentage: SubFormats.Text := PercentFormats;
    nfCurrency: SubFormats.Text := MoneyFormats;
    nfLongDate: SubFormats.Text := DateFormats;
    nfLongTime: SubFormats.Text := TimeFormats;
    nfExp: SubFormats.Text := ExpFormats;
    nfFraction: SubFormats.Text := FracFormats;
    nfCustom: SubFormats.Text := BoolFormats;
    nfText: SubFormats.Text := TextFormats;
    else
      SubFormats.Text :=  ' ';
  end;
  SubFormatList.NumRows := SubFormats.Count;
  SubFormatList.Quiet := True;
  if (NewFmt = nfCustom) or (NewFmt = nfText) then
  begin
    SubFormatList.Cells[0, 0] := SubFormats.Text;
  end
  else
  begin
    for i := 0 to SubFormats.Count - 1 do
    begin
      SubFormats[i] := stringreplace(SubFormats[i], '$', '$' + DefMoneySymb, [rfReplaceAll]);
      SubFormatList.Cells[0, i] := GetExampleText(SubFormats[i], True);
    end;
  end;
  SubFormatList.Quiet := False;
  SubFormatList.Row := 0;
  ClickFormatEvent(nil);
  //writeln('Setformat: ', GetEnumName(TypeInfo(TsNumberFormat), Ord(NewFmt)));
  //SG.Worksheet.WriteNumberFormat(Cell.Y - 1, Cell.X - 1, NewFmt, '#,##0.00_);(#,##0.00)');
  //SG.RedrawCell(Cell.X, Cell.Y);

end;

procedure TFormatWin.ClickFormatEvent(Sender: TObject);
begin
  FormatEdit.Contents := SubFormats[SubFormatList.Row];
  AckEdit(nil);
end;

procedure TFormatWin.AckEdit(Sender: TObject);
var
  str: string;
  Pen: LongInt;
begin
  // make preview
  str := GetExampleText(FormatEdit.Contents, False);
  DB.Clear(2);
  Pen := -1;
  if Pos('RED', FormatEdit.Contents) > 0 then
  begin
    Pen := ObtainBestPenA(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, $ff000000, 0, 0, nil);
    SetAPen(DB.RP, Pen);
  end
  else
    SetAPen(DB.RP, 1);
  SetDrMd(DB.RP, JAM1);
  DB.DrawText(5, 12, str);
  if Pen >= 0 then
    ReleasePen(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, Pen);
  MUI_Redraw(Preview.MUIObj, MADF_DRAWOBJECT);
end;

procedure TFormatWin.OKClick(Sender: TObject);
var
  x,y: Integer;
begin

  if (SelectedCells[0].Col1 = SelectedCells[0].Col2) and (SelectedCells[0].Row1 = SelectedCells[0].Row2) then
  begin
    SG.Worksheet.WriteNumberFormat(Cell.Y - 1, Cell.X - 1, NewFmt, FormatEdit.Contents);
    SG.RedrawCell(Cell.X, Cell.Y);
  end
  else
  begin
    for x := SelectedCells[0].Col1 to SelectedCells[0].Col2 do
    begin
      for y := SelectedCells[0].Row1 to SelectedCells[0].Row2 do
      begin
        SG.Worksheet.WriteNumberFormat(Y, X, NewFmt, FormatEdit.Contents);
        SG.RedrawCell(X, Y);
      end;
    end;
  end;
  //
  Close;
end;

procedure TFormatWin.CancelClick(Sender: TObject);
begin
  Close;
end;

procedure GetLocaleSettings;
var
  Loc: PLocale;
begin
  Loc := OpenLocale(nil);
  if Assigned(Loc) then
  begin
    DefMoneySymb := Loc^.loc_MonCS;
    CloseLocale(Loc);
  end;
end;

{ TFormatDesc }
constructor TFormatDesc.Create;
begin
  FormatCodes := TStringList.Create;
end;

destructor TFormatDesc.Destroy;
begin

end;


initialization
  GetLocaleSettings;

end.
