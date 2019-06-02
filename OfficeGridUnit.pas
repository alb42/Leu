unit OfficeGridUnit;
{$mode objfpc}{$H+}

interface

uses
  Types, Classes, SysUtils, intuition, agraphics, exec, utility,
  Math,
  MUIClass.Base,
  MUIClass.Grid,
  fpstypes, fpspreadsheet, fpsallformats, fpsutils, fpsnumformat, variants;

type
  TOfficeGrid = class(TMUIStrGrid)
  private
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FShowHeaders: Boolean;
    FHeaderCount: Integer;
    FShowFormulas: Boolean;
    function GetCellText(ACol, ARow: Integer; ATrim: Boolean = true): String;
    procedure SetCellValue(ACol, ARow: Integer; AValue: Variant);
  protected
    function GetCell(ACol, ARow: Integer): string; override;
    procedure SetCell(ACol, ARow: Integer; AValue: string); override;
    procedure DoDrawCell(Sender: TObject; ACol, ARow: Integer; RP: PRastPort; ARect: TRect); override;
    function CalcColWidthFromSheet(AWidth: Single): Integer;
  public
    procedure LoadFile(FileName: string);

    property ShowHeaders: Boolean read FShowHeaders write FShowHeaders;
    property Worksheet: TsWorksheet read FWorksheet;
  end;

implementation

function TOfficeGrid.GetCellText(ACol, ARow: Integer; ATrim: Boolean = true): String;
var
  cell: PCell;
  r, c: Integer;
begin
  Result := '';

  if ShowHeaders then
  begin
    // Headers
    if (ARow = 0) and (ACol = 0) then
      exit;
    if (ARow = 0) then
    begin
      Result := GetColString(ACol - FHeaderCount);
      //if Assigned(FGetColHeaderText) then
      //  FGetColHeaderText(Self, ACol, Result);
      exit;
    end
    else
    if (ACol = 0) then
    begin
      Result := IntToStr(ARow);
      //if Assigned(FGetRowHeaderText) then
      //  FGetRowHeaderText(Self, ARow, Result);
      exit;
    end;
  end;

  if Worksheet <> nil then
  begin
    r := ARow - FHeaderCount;
    c := ACol - FHeaderCount;
    cell := Worksheet.FindCell(r, c);
    if cell <> nil then
    begin
      if (HasFormula(cell)  { and not (boAutoCalc in FWorkbook.Options)}) and FShowFormulas then
      begin
        Result := '=' + Worksheet.ReadFormula(cell);
        //sysdebugln('Formula as Formula');
      end
      else
      begin
        Result := Worksheet.ReadAsText(cell);
        //sysdebugln('Formula as text');
      end;
    end;
  end;
end;

procedure TOfficeGrid.SetCellValue(ACol, ARow: Integer; AValue: Variant);
var
  cell: PCell = nil;
  fmt: PsCellFormat = nil;
  idx: Integer;
  nfp: TsNumFormatParams;
  r, c: Cardinal;
  s: String;
  rtParams: TsRichTextParams;
begin
  if not Assigned(Worksheet) then
    exit;

  r := ARow - FixedRows;
  c := ACol - FixedCols;

  // If the cell already exists and contains a formula then the formula must be
  // removed. The formula would dominate over the data value.
  cell := Worksheet.FindCell(r, c);
  if HasFormula(cell) then
    Worksheet.UseformulaInCell(cell, nil);   //cell^.FormulaValue := '';

  if VarIsNull(AValue) then
    Worksheet.WriteBlank(r, c)
  else
  if VarIsStr(AValue) then
  begin
    s := VarToStr(AValue);
    if (s <> '') and (s[1] = '=') then
      Worksheet.WriteFormula(r, c, Copy(s, 2, Length(s)), true)
    else
    begin
      if cell = nil then cell := Worksheet.GetCell(r, c);
      if s = '' then
        Worksheet.WriteBlank(cell)
      else begin
        //HTMLToRichText(Workbook, Worksheet.ReadCellFont(cell), s, plain, rtParams);
        rtParams := [];
        Worksheet.WriteText(cell, s, rtParams);  // This will erase a non-formatted cell if s = ''
      end;
    end;
  end else
  if VarIsType(AValue, varDate) then
    Worksheet.WriteDateTime(r, c, VarToDateTime(AValue))
  else
  if VarIsNumeric(AValue) then
  begin
    // Check if the cell already exists and contains a format.
    // If it is a date/time format write a date/time cell...
    if cell <> nil then
    begin
      idx := Worksheet.GetEffectiveCellFormatIndex(cell);
      fmt := FWorkbook.GetPointerToCellFormat(idx);
      if fmt <> nil then
        nfp := FWorkbook.GetNumberFormat(fmt^.NumberFormatIndex);
      if (fmt <> nil) and IsDateTimeFormat(nfp) then
        FWorksheet.WriteDateTime(r, c, VarToDateTime(AValue)) else
        FWorksheet.WriteNumber(r, c, AValue);
    end
    else
      // ... otherwise write a number cell
      FWorksheet.WriteNumber(r, c, AValue);
  end else
  if VarIsBool(AValue) then
    FWorksheet.WriteBoolValue(r, c, AValue);
end;

function TOfficeGrid.CalcColWidthFromSheet(AWidth: Single): Integer;
var
  w_pts: Double;
begin
  w_pts := FWorkbook.ConvertUnits(AWidth, FWorkbook.Units, suPoints);
  Result := PtsToPx(w_pts, 72);
end;

procedure TOfficeGrid.LoadFile(FileName: string);
var
  cell: PCell;
  MaxC, MaxR: Integer;
  lCol: PCol;
  i, w: Integer;       // Col width at current zoom level
  w100: Integer;    // Col width at 100% zoom level
begin
  FShowHeaders := True;
  FHeaderCount := 1;
  if Assigned(FWorkbook) then
    FWorkBook.Free;
  FWorkbook := TsWorkbook.Create;

  FWorkbook.Options := FWorkbook.Options + [boReadFormulas];
  FWorkbook.ReadFromFile(Filename);
  FWorkbook.Options := FWorkbook.Options + [boAutoCalc];

  FWorksheet := FWorkbook.GetFirstWorksheet;

  MaxR := 2;
  MaxC := 2;
  for Cell in FWorksheet.Cells do
  begin
    MaxR := Max(Cell^.Row, MaxR);
    MaxC := Max(Cell^.Col, MaxC);
  end;
  BeginUpdate;
  NumRows := MaxR + 2;
  NumCols := MaxC + 2;

  for i := 0 to NumCols - 1 do
  begin
    lCol := Worksheet.FindCol(i - FHeaderCount);
    if (lCol <> nil) and lCol^.Hidden then
      w := 0
    else begin
      if (lCol <> nil) and (lCol^.ColWidthType = cwtCustom) then
        w100 := CalcColWidthFromSheet(lCol^.Width)
      else
        w100 := CalcColWidthFromSheet(FWorksheet.ReadDefaultColWidth(FWorkbook.Units));
      w := round(w100 * 1.0);
    end;
    CellWidth[i] := w;
  end;
  EndUpdate;
end;


function TOfficeGrid.GetCell(ACol, ARow: Integer): string;
begin
  Result := GetCellText(ACol, ARow);
end;

procedure TOfficeGrid.SetCell(ACol, ARow: Integer; AValue: string);
var
  f: Double;
begin
  if TryStrToFloat(AValue, f) then
    SetCellValue(ACol, ARow, f)
  else
    SetCellValue(ACol, ARow, AValue);
  inherited;
end;

procedure TOfficeGrid.DoDrawCell(Sender: TObject; ACol, ARow: Integer; RP: PRastPort; ARect: TRect);
var
  s: string;
  CS: TCellStatus;
  Color: TsColor;
  Cell: PCell;
  Pen: LongInt;
  SIdx, Idx: Integer;
  fmt: PsCellFormat;
  horAlign: TsHorAlignment;
  numfmt: TsNumFormatParams;
  TE: TTextExtent;
  fnt: TsFont;
  fst: LongWord;
  NStyle: LongWord;
begin
  Pen := -1;
  FShowFormulas := False;
  s := Cells[ACol, ARow];
  CS := CellStatus[ACol, ARow];

  fmt := nil;
  Color := 0;
  if (FWorkbook <> nil) and (FWorksheet <> nil) then
  begin
    cell := FWorksheet.FindCell(ARow - FixedRows, ACol - FixedCols);
    fnt :=FWorksheet.ReadCellFont(cell);
    Color := fnt.Color;
    idx := FWorksheet.GetEffectiveCellFormatIndex(Cell);
    fmt := FWorkbook.GetPointerToCellFormat(idx);
  end
  else
  begin
    inherited;
    Exit;
  end;

  if Assigned(Cell) and Assigned(fnt) and Assigned(fmt) then
  begin
    if (uffHorAlign in fmt^.UsedFormattingFields) then
      horAlign := fmt^.HorAlignment
    else
      horAlign := haDefault;
    if horAlign = haDefault then
    begin
      if (Cell^.ContentType in [cctNumber, cctDateTime]) then
        horAlign := haRight
      else
      begin
        if (Cell^.ContentType in [cctBool]) then
          horAlign := haCenter
        else
          horAlign := haLeft;
      end;
    end;

    // Font color as derived from number format
    if not IsNaN(Cell^.NumberValue) and (uffNumberFormat in fmt^.UsedFormattingFields) then
    begin
      numFmt := FWorkbook.GetNumberFormat(fmt^.NumberFormatIndex);
      if numFmt <> nil then
      begin
        sidx := 0;
        if (Length(numFmt.Sections) > 1) and (Cell^.NumberValue < 0) then
          sidx := 1
        else
        if (Length(numFmt.Sections) > 2) and (Cell^.NumberValue = 0) then
          sidx := 2;
        if (nfkHasColor in numFmt.Sections[sidx].Kind) then
          Color := numFmt.Sections[sidx].Color and $00FFFFFF;
      end;
    end;
  end;

  if CS = csFixed then
  begin
    horAlign := haCenter;
    Color := 0;
  end;

  SetDrMd(RP, JAM1);
  if (CS = csSelected) or (CS = csFocussed) then
  begin
    SetAPen(RP, 3);
    RectFill(RP, ARect.Left, ARect.Top, ARect.Right, ARect.Bottom);
    SetAPen(RP, 2);
  end
  else
  begin
    if Color <> 0 then
    begin
      {$ifdef ENDIAN_LITTLE}
      Pen := ObtainBestPenA(ViewPortAddress(IntuitionBase^.ActiveWindow)^.ColorMap, (Color and $ff) shl 24, (Color and $ff00) shl 16, (color and $ff0000) shl 8, nil);
      {$else}
      Pen := ObtainBestPenA(ViewPortAddress(IntuitionBase^.ActiveWindow)^.ColorMap, (color and $ff0000) shl 8, (Color and $ff00) shl 16, (Color and $ff) shl 24, nil);
      {$endif}
      SetAPen(RP, Pen);
    end
    else
      SetAPen(RP, 1);
  end;
  TextExtent(RP, PChar(s), Length(s), @TE);
  case horAlign of
    haLeft: GfxMove(RP, ARect.Left + 1, ARect.Top + ARect.Height div 2 + RP^.Font^.tf_Baseline div 2);
    haCenter: GfxMove(RP, (ARect.Left + ARect.Width div 2) - (TE.te_Width div 2), ARect.Top + ARect.Height div 2 + RP^.Font^.tf_Baseline div 2);
    haRight: GfxMove(RP, ARect.Right - TE.te_Width - 1, ARect.Top + ARect.Height div 2 + RP^.Font^.tf_Baseline div 2);
  end;
  if Assigned(fnt) then
  begin
    fst := AskSoftStyle(RP);
    NStyle := 0;
    if fssBold in fnt.Style then NStyle := NStyle or FSF_BOLD;
    if fssItalic in fnt.Style then NStyle := NStyle or FSF_ITALIC;
    if fssUnderline in fnt.Style then NStyle := NStyle or FSF_UNDERLINED;
    SetSoftStyle(RP, fst, NStyle);
  end;
  GfxText(RP, PChar(s), Length(s));
  SetSoftStyle(RP, fst, 0);


  if CS = csFocussed then
  begin
    SetAPen(RP, 1);
    GfxMove(RP, ARect.Left + 1, ARect.Top + 1);
    AGraphics.Draw(RP, ARect.Right, ARect.Top + 1);
    AGraphics.Draw(RP, ARect.Right, ARect.Bottom);
    AGraphics.Draw(RP, ARect.Left + 1, ARect.Bottom);
    AGraphics.Draw(RP, ARect.Left + 1, ARect.Top + 1);
  end;
  if Pen >= 0 then
    ReleasePen(ViewPortAddress(IntuitionBase^.ActiveWindow)^.ColorMap, Pen);

  FShowFormulas := True;
end;


end.
