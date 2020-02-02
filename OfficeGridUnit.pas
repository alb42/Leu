unit OfficeGridUnit;
{$mode objfpc}{$H+}

interface

uses
  Types, Classes, SysUtils, intuition, agraphics, exec, utility, iffparse,
  Math, imagesunit,
  MUIClass.Base,
  MUIClass.Grid,
  MUIClass.Dialog,
  fpstypes, fpspreadsheet, fpsallformats, fpsutils, fpsnumformat, variants;

type
  TOfficeGrid = class(TMUIStrGrid)
  private
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FShowHeaders: Boolean;
    FHeaderCount: Integer;
    FShowFormulas: Boolean;
    FFilename: string;
    FMinCols: Integer;
    FMinRows: Integer;
    RTGMode: Boolean;
    function GetCellText(ACol, ARow: Integer; ATrim: Boolean = true): String;
    procedure SetCellValue(ACol, ARow: Integer; AValue: Variant);
    function GetWorksheetCount: Integer;
    function GetWorksheetIdx: Integer;
    function GetCellFontStyle(ACol, ARow: Integer): TsFontStyles;
    procedure SetCellFontStyle(ACol, ARow: Integer;  AValue: TsFontStyles);
    function GetHorAlignment(ACol, ARow: Integer): TsHorAlignment;
    procedure SetHorAlignment(ACol, ARow: Integer; AValue: TsHorAlignment);

    procedure ChangedCellHandler(ASender: TObject; ARow, ACol:Cardinal);

  protected
    function GetCell(ACol, ARow: Integer): string; override;
    procedure SetCell(ACol, ARow: Integer; AValue: string); override;
    //
    function GetCellWidth(ACol: Integer): Integer; override;
    procedure SetCellWidth(ACol: Integer; AWidth: Integer); override;
    function GetCellHeight(ARow: Integer): Integer; override;
    procedure SetCellHeight(ARow: Integer; AHeight: Integer); override;
    //
    function GetCellBorderStyle(ACol, ARow: Integer; ABorder: TsCellBorder): TsCellBorderStyle;
    function HasBorder(ACell: PCell; ABorder: TsCellBorder): Boolean;
    function GetBorderStyle(ACol, ARow, ADeltaCol, ADeltaRow: Integer; ACell: PCell; out ABorderStyle: TsCellBorderStyle): Boolean;
    procedure DrawCellBorders(RP: PRastPort; ACol, ARow: Integer; ARect: TRect; ACell: PCell);
    procedure DoDrawCell(Sender: TObject; ACol, ARow: Integer; RP: PRastPort; ARect: TRect); override;
    function CalcColSizeFromSheet(ASize: Single): Integer;
    function CalcSizeToSheet(ASize: Integer): Single;

    procedure FixNeighborCellBorders(ACell: PCell);

    procedure DoDeleteCell(ACol, ARow: Integer); override;
  public
    constructor Create; override;
    destructor Destroy; override;

    procedure NewWorkbook(ANumCols: Integer = 20; ANumRows: Integer = 40);
    procedure LoadFile(AFileName: string);
    procedure SaveFile(AFileName: string);
    procedure LoadWorksheet(Idx: Integer);
    procedure AddWorksheet();
    procedure RemoveWorksheet();

    procedure CopyToClip; override;
    procedure CutToClip; override;
    procedure PasteFromClip; override;

    property ShowHeaders: Boolean read FShowHeaders write FShowHeaders;
    property Worksheet: TsWorksheet read FWorksheet;
    property WorkBook: TsWorkBook read FWorkBook;

    property WorksheetCount: Integer read GetWorksheetCount;
    property WorksheetIdx: Integer read GetWorksheetIdx;
    property Filename: string read FFilename;
    property MinCols: Integer read FMinCols write FMinCols;
    property MinRows: Integer read FMinRows write FMinRows;

    property CellFontStyle[ACol, ARow: Integer]: TsFontStyles read GetCellFontStyle write SetCellFontStyle;
    property HorAlignment[ACol, ARow: Integer]: TsHorAlignment read GetHorAlignment write SetHorAlignment;

    procedure SetCellBorder(ACol, ARow: Integer; AValue: TsCellBorders);
    procedure SetCellBorders(ALeft, ATop, ARight, ABottom: Integer; AValue: TsCellBorders);
  end;

function DrawBorder2CellBorder(c: TDrawBorders): TsCellBorders;

implementation

uses
  TCDReaderUnit, clipunit;

function CalcSelectionColor(c: TsColor; ADelta: Byte) : TsColor;
begin
  TRGBA(Result).A := 0;
  if TRGBA(c).R < 128
    then TRGBA(Result).R := TRGBA(c).R + ADelta
    else TRGBA(Result).R := TRGBA(c).R - ADelta;
  if TRGBA(c).G < 128
    then TRGBA(Result).G := TRGBA(c).G + ADelta
    else TRGBA(Result).G := TRGBA(c).G - ADelta;
  if TRGBA(c).B < 128
    then TRGBA(Result).B := TRGBA(c).B + ADelta
    else TRGBA(Result).B := TRGBA(c).B - ADelta;
end;

function DrawBorder2CellBorder(c: TDrawBorders): TsCellBorders;
begin
  Result := [];
  if dbTop in c then
    Include(Result, cbNorth);
  if dbBottom in c then
    Include(Result, cbSouth);
  if dbLeft in c then
    Include(Result, cbWest);
  if dbRight in c then
    Include(Result, cbEast);
end;

constructor TOfficeGrid.Create;
begin
  inherited;
  CycleChain := 1;
  FMinCols := 10;
  FMinRows := 20;
  FShowHeaders := True;
  FHeaderCount := 1;
  NewWorkbook;
  RTGMode := GetBitMapAttr(@(IntuitionBase^.ActiveScreen^.Bitmap), BMA_DEPTH) > 8;
end;


destructor TOfficeGrid.Destroy;
begin
  FWorkbook.Free;
  inherited;
end;

procedure TOfficeGrid.NewWorkbook(ANumCols: Integer = 20; ANumRows: Integer = 40);
begin
  InitChange;
  BeginUpdate;
  BlockRecalcSize := True;
  NumRows := 0;
  NumCols := 0;
  FWorkbook.Free;
  FFilename := '';
  FWorkbook := TsWorkbook.Create;
  FWorkbook.Options := FWorkbook.Options + [boAutoCalc];
  FWorksheet := FWorkbook.AddWorksheet('My Worksheet');
  FWorksheet.OnChangeCell := @ChangedCellHandler;
  AllToRedraw := True;
  BlockRecalcSize := False;
  LoadWorksheet(-1); // load first Worksheet
  NumRows := ANumRows;
  NumCols := ANumCols;
  RecalcSize;
  EndUpdate;
  ExitChange;
end;

procedure TOfficeGrid.ChangedCellHandler(ASender: TObject; ARow, ACol:Cardinal);
begin
  Unused(ASender);
  AddToRedraw(ACol + FixedCols, ARow + FixedRows);
end;

function TOfficeGrid.GetWorksheetCount: Integer;
begin
  Result := 0;
  if Assigned(FWorkbook) then
    Result := FWorkbook.GetWorksheetCount;
end;

function TOfficeGrid.GetWorksheetIdx: Integer;
begin
  Result := -1;
  if Assigned(FWorkbook) then
    Result := FWorkbook.GetWorksheetIndex(FWorksheet);
end;


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
  Result := UTF8ToAnsi(Result);
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
    begin
      try
        Worksheet.WriteFormula(r, c, Copy(s, 2, Length(s)), true)
      except
        On E:Exception do
        begin
          Worksheet.Formulas.DeleteFormula(R, c);
          ShowMessage('Error in Formula '#10 + e.Message);
          if cell = nil then cell := Worksheet.GetCell(r, c);
          try
            Worksheet.WriteText(cell, s, []);
          except
          end;
        end;
      end;
    end
    else
    begin
      if cell = nil then cell := Worksheet.GetCell(r, c);
      if s = '' then
        Worksheet.WriteBlank(cell)
      else begin
        //HTMLToRichText(Workbook, Worksheet.ReadCellFont(cell), s, plain, rtParams);
        rtParams := [];
        try
        Worksheet.WriteText(cell, s, rtParams);  // This will erase a non-formatted cell if s = ''
        except
        end;
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

function TOfficeGrid.GetCellFontStyle(ACol, ARow: Integer): TsFontStyles;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := [];
  if (FWorkbook <> nil) and (FWorksheet <> nil) then begin
    cell := FWorksheet.FindCell(ARow - FixedRows, ACol - FixedCols);
    fnt := FWorksheet.ReadCellFont(cell);
    Result := fnt.Style;
  end;
end;

procedure TOfficeGrid.SetCellFontStyle(ACol, ARow: Integer;
  AValue: TsFontStyles);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(ARow - FixedRows, ACol - FixedCols);
    Worksheet.WriteFontStyle(cell, AValue);
  end;
end;

function TOfficeGrid.GetHorAlignment(ACol, ARow: Integer): TsHorAlignment;
var
  cell: PCell;
begin
  Result := haDefault;
  if Assigned(Worksheet) then begin
    cell := Worksheet.FindCell(ARow - FixedRows, ACol - FixedCols);
    Result := Worksheet.ReadHorAlignment(cell);
  end;
end;

procedure TOfficeGrid.SetHorAlignment(ACol, ARow: Integer;
  AValue: TsHorAlignment);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(ARow - FixedRows, ACol - FixedCols);
    Worksheet.WriteHorAlignment(cell, AValue);
  end;
end;


function TOfficeGrid.CalcColSizeFromSheet(ASize: Single): Integer;
var
  w_pts: Double;
begin
  w_pts := FWorkbook.ConvertUnits(ASize, FWorkbook.Units, suPoints);
  Result := PtsToPx(w_pts, 72);
end;

function TOfficeGrid.CalcSizeToSheet(ASize: Integer): Single;
var
  h_pts: Single;
begin
  h_pts := PxToPts(ASize, 72);
  Result := FWorkbook.ConvertUnits(h_pts, suPoints, FWorkbook.Units);
end;

procedure TOfficeGrid.LoadWorksheet(Idx: Integer);
var
  cell: PCell;
  MaxC, MaxR: Integer;
begin
  if Idx < 0 then
    FWorksheet := FWorkbook.GetFirstWorksheet
  else
  if Idx >= WorksheetCount then
    FWorksheet := FWorkbook.GetLastWorksheet
  else
    FWorksheet := FWorkbook.GetWorksheetByIndex(Idx);
  if not Assigned(FWorksheet) then
    Exit;
  FWorksheet.OnChangeCell := @ChangedCellHandler;
  MaxR := NumRows;
  MaxC := NumCols;
  for Cell in FWorksheet.Cells do
  begin
    MaxR := Max(Cell^.Row, MaxR);
    MaxC := Max(Cell^.Col, MaxC);
  end;
  if (MaxR <> NumRows) or (MaxC <> NumCols) then
  begin
		BeginUpdate;
		BlockRecalcSize := True;
		try
			NumRows := MaxR;
			NumCols := MaxC;
		finally
			BlockRecalcSize := False;
			RecalcSize;
			EndUpdate;
		end;
  end;
  RedrawAllCells;
end;

procedure TOfficeGrid.AddWorksheet();
begin
  FWorksheet := FWorkBook.AddWorkSheet('New Worksheet', True);
  FWorksheet.OnChangeCell := @ChangedCellHandler;
  RedrawAllCells;
end;

procedure TOfficeGrid.RemoveWorksheet();
begin
  FWorkBook.RemoveWorkSheet(FWorksheet);
  LoadWorkSheet(-1);
end;


procedure TOfficeGrid.LoadFile(AFileName: string);
begin
  FFilename := '';
  FWorkBook.Free;
  FWorksheet := nil;

  FWorkbook := TsWorkbook.Create;
  FWorkbook.Options := FWorkbook.Options + [boReadFormulas, boAutoCalc];
  if LowerCase(ExtractFileExt(AFileName)) = '.tcd' then
  begin
    ReadTCD(FWorkbook, AFilename);
  end
  else
    FWorkbook.ReadFromFile(AFilename);
  LoadWorksheet(-1); // load first Worksheet
  FFilename := AFilename;
end;

procedure TOfficeGrid.SaveFile(AFileName: string);
begin
  if AFileName <> '' then
    FFileName := AFilename;
  if FFilename = '' then
    Exit;
  FWorkbook.WriteToFile(FFileName, True);
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
    SetCellValue(ACol, ARow, AnsiToUTF8(AValue));
  inherited;
end;

function TOfficeGrid.GetCellWidth(ACol: Integer): Integer;
var
  lCol: PCol;
  w: Integer;
  w100: Integer;
begin
  lCol := Worksheet.FindCol(ACol - FHeaderCount);
  if lCol = nil then
  begin
    Result := inherited;
    Exit;
  end;
  //
  if (lCol <> nil) and lCol^.Hidden then
    w := 0
  else
  begin
    if lCol^.ColWidthType = cwtCustom then
      w100 := CalcColSizeFromSheet(lCol^.Width)
    else
      w100 := CalcColSizeFromSheet(FWorksheet.ReadDefaultColWidth(FWorkbook.Units));
    w := round(w100 * 1.2);
  end;
  Result := w;
end;

procedure TOfficeGrid.SetCellWidth(ACol: Integer; AWidth: Integer);
var
  lCol: PCol;
begin
  lCol := Worksheet.FindCol(ACol - FHeaderCount);
  if lCol = nil then
  begin
    inherited;
    Exit;
  end;
  //
  if lCol^.Hidden then
    Exit;
  lCol^.ColWidthType := cwtCustom;
  lCol^.Width := CalcSizeToSheet(AWidth) / 1.2;
  RecalcSize;
end;

function TOfficeGrid.GetCellHeight(ARow: Integer): Integer;
var
  lRow: PRow;
  h: Integer;
  h100: Integer;
begin
  lRow := Worksheet.FindRow(ARow - FixedRows);
  if lRow = nil then
  begin
    Result := inherited;
    Exit;
  end;
  //
  if lRow^.Hidden then
    h := 0
  else
  begin
    if lRow^.RowHeightType = rhtCustom then
      h100 := CalcColSizeFromSheet(lRow^.Height)
    else
      h100 := CalcColSizeFromSheet(FWorksheet.ReadDefaultRowHeight(FWorkbook.Units));
    h := round(h100 * 1.2);
  end;
  Result := h;
end;

procedure TOfficeGrid.SetCellHeight(ARow: Integer; AHeight: Integer);
var
  lRow: PRow;
begin
  lRow := Worksheet.FindRow(ARow - FixedRows);
  if lRow = nil then
  begin
    inherited;
    Exit;
  end;
  if lRow^.Hidden then
    Exit;
  lRow^.RowHeightType := rhtCustom;
  lRow^.Height := CalcSizeToSheet(AHeight) / 1.2;
  RecalcSize;
end;

function TOfficeGrid.GetCellBorderStyle(ACol, ARow: Integer;
  ABorder: TsCellBorder): TsCellBorderStyle;
var
  cell: PCell;
begin
  Result := DEFAULT_BORDERSTYLES[ABorder];
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.FindCell(ARow - FHeaderCount, ACol);
    if Worksheet.IsMerged(cell) then
      cell := Worksheet.FindMergeBase(cell);
    Result := Worksheet.ReadCellBorderStyle(cell, ABorder);
  end;
end;

function TOfficeGrid.HasBorder(ACell: PCell; ABorder: TsCellBorder): Boolean;
var
  base: PCell;
  r1, c1, r2, c2: Cardinal;
begin
  if Worksheet = nil then
    result := false
  else
  if Worksheet.IsMerged(ACell) then
  begin
    Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    base := Worksheet.FindCell(r1, c1);
    Result := ABorder in Worksheet.ReadCellBorders(base);
    case ABorder of
      cbNorth : if ACell^.Row > r1 then Result := false;
      cbSouth : if ACell^.Row < r2 then Result := false;
      cbEast  : if ACell^.Col < c2 then Result := false;
      cbWest  : if ACell^.Col > c1 then Result := false;
    end;
  end else
    Result := ABorder in Worksheet.ReadCellBorders(ACell);
end;

{@@ ----------------------------------------------------------------------------
  Determines the style of the border between a cell and its neighbor given by
  ADeltaCol and ADeltaRow (one of them must be 0, the other one can only be +/-1).
  ACol and ARow are in grid units.
  Result is FALSE if there is no border line.
-------------------------------------------------------------------------------}
function TOfficeGrid.GetBorderStyle(ACol, ARow, ADeltaCol, ADeltaRow: Integer;
  ACell: PCell; out ABorderStyle: TsCellBorderStyle): Boolean;
var
  neighborcell: PCell;
  border, neighborborder: TsCellBorder;
  r, c: Cardinal;
begin
  Result := true;

  if (ADeltaCol = -1) and (ADeltaRow = 0) then
  begin
    border := cbWest;
    neighborborder := cbEast;
  end else
  if (ADeltaCol = +1) and (ADeltaRow = 0) then
  begin
    border := cbEast;
    neighborborder := cbWest;
  end else
  if (ADeltaCol = 0) and (ADeltaRow = -1) then
  begin
    border := cbNorth;
    neighborborder := cbSouth;
  end else
  if (ADeltaCol = 0) and (ADeltaRow = +1) then
  begin
    border := cbSouth;
    neighborBorder := cbNorth;
  end else
    raise Exception.Create('[TsCustomWorksheetGrid] Incorrect col/row for GetBorderStyle.');

  r := (ARow - FHeaderCount);
  c := (ACol - FHeaderCount);
  if (longint(r) + ADeltaRow < 0) or (longint(c) + ADeltaCol < 0) then
    neighborcell := nil
  else
    neighborcell := Worksheet.FindCell(longint(r) + ADeltaRow, longint(c) + ADeltaCol);

  // Only cell has border, but neighbor has not
  if HasBorder(ACell, border) and not HasBorder(neighborCell, neighborBorder) then
  begin
    if Worksheet.InSameMergedRange(ACell, neighborcell) then
      result := false
    else
      ABorderStyle := GetCellBorderStyle(ACol, ARow, border);
  end
  else
  // Only neighbor has border, cell has not
  if not HasBorder(ACell, border) and HasBorder(neighborCell, neighborBorder) then
  begin
    if Worksheet.InSameMergedRange(ACell, neighborcell) then
      result := false
    else
      ABorderStyle := GetCellBorderStyle(ACol+ADeltaCol, ARow+ADeltaRow, neighborborder);
  end
  else
  // Both cells have shared border -> use top or left border
  if HasBorder(ACell, border) and HasBorder(neighborCell, neighborBorder) then
  begin
    if Worksheet.InSameMergedRange(ACell, neighborcell) then
      result := false
    else
    if (border in [cbNorth, cbWest]) then
      ABorderStyle := GetCellBorderStyle(ACol+ADeltaCol, ARow+ADeltaRow, neighborborder)
    else
      ABorderStyle := GetCellBorderStyle(ACol, ARow, border);
  end else
    Result := false;
end;

procedure TOfficeGrid.DrawCellBorders(RP: PRastPort; ACol, ARow: Integer; ARect: TRect; ACell: PCell);
const
  drawHor = 0;
  drawVert = 1;
  drawDiagUp = 2;
  drawDiagDown = 3;

  procedure DoLine(L, T, R, B: Integer);
  begin
    GFXMove(RP, L, T);
    AGraphics.Draw(RP, R, B);
  end;

  procedure DrawBorderLine(ACoord: Integer; ARect: TRect; ADrawDirection: Byte;
    ABorderStyle: TsCellBorderStyle);
  var
    deltax, deltay: Integer;
    angle: Double;
    //savedCosmetic: Boolean;
  begin
    if ABorderStyle.Color = scWhite then
      SetAPen(RP, 2);
    if ABorderStyle.Color = scBlack then
      SetAPen(RP, 1);

    case ABorderStyle.LineStyle of
      lsThin, lsMedium, lsThick, lsDotted, lsDashed, lsDashDot, lsDashDotDot,
      lsMediumDash, lsMediumDashDot, lsMediumDashDotDot, lsSlantDashDot, lsHair:
        case ADrawDirection of
          drawHor     : DoLine(ARect.Left, ACoord, ARect.Right, ACoord);
          drawVert    : DoLine(ACoord, ARect.Top, ACoord, ARect.Bottom);
          drawDiagUp  : DoLine(ARect.Left, ARect.Bottom, ARect.Right, ARect.Top);
          drawDiagDown: DoLine(ARect.Left, ARect.Top, ARect.Right, ARect.Bottom);
        end;
      {
      lsHair:
        case ADrawDirection of
          drawHor     : DrawHairLineHor(Canvas, ARect.Left, ARect.Right, ACoord);
          drawVert    : DrawHairLineVert(Canvas, ACoord, ARect.Top, ARect.Bottom);
          drawDiagUp  : ;
          drawDiagDown: ;
        end;
      }
      lsDouble:
        case ADrawDirection of
          drawHor:
            begin
              DoLine(ARect.Left, ACoord-1, ARect.Right, ACoord-1);
              DoLine(ARect.Left, ACoord+1, ARect.Right, ACoord+1);
              //Canvas.Pen.Color := Color;
              DoLine(ARect.Left, ACoord, ARect.Right, ACoord);
            end;
          drawVert:
            begin
              DoLine(ACoord-1, ARect.Top, ACoord-1, ARect.Bottom);
              DoLine(ACoord+1, ARect.Top, ACoord+1, ARect.Bottom);
              //Canvas.Pen.Color := Color;
              DoLine(ACoord, ARect.Top, ACoord, ARect.Bottom);
            end;
          drawDiagUp:
            begin
              if ARect.Right = ARect.Left then
                angle := pi/2
              else
                angle := arctan((ARect.Bottom-ARect.Top) / (ARect.Right-ARect.Left));
              deltax := Max(1, round(1.0 / sin(angle)));
              deltay := Max(1, round(1.0 / cos(angle)));
              DoLine(ARect.Left, ARect.Bottom-deltay-1, ARect.Right-deltax, ARect.Top-1);
              DoLine(ARect.Left+deltax, ARect.Bottom-1, ARect.Right, ARect.Top+deltay-1);
            end;
          drawDiagDown:
            begin
              if ARect.Right = ARect.Left then
                angle := pi/2
              else
                angle := arctan((ARect.Bottom-ARect.Top) / (ARect.Right-ARect.Left));
              deltax := Max(1, round(1.0 / sin(angle)));
              deltay := Max(1, round(1.0 / cos(angle)));
              DoLine(ARect.Left, ARect.Top+deltay-1, ARect.Right-deltax, ARect.Bottom-1);
              DoLine(ARect.Left+deltax, ARect.Top-1, ARect.Right, ARect.Bottom-deltay-1);
            end;
        end;
    end;
    //Canvas.Pen.Cosmetic := savedCosmetic;
  end;

var
  bs: TsCellBorderStyle;
  fmt: PsCellFormat;
  idx: Integer;
  r1, c1, r2, c2: Cardinal;
begin
  if Assigned(Worksheet) then begin
    if Worksheet.IsMergeBase(ACell) then begin
      Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
      ARect := CellsToRect(c1 + FHeaderCount, r1 + FHeaderCount, c2 + FHeaderCount, r2 + FHeaderCount);
    end;
    // Left border
    if GetBorderStyle(ACol, ARow, -1, 0, ACell, bs) then
      DrawBorderLine(ARect.Left, ARect, drawVert, bs);
    // Right border
    if GetBorderStyle(ACol, ARow, +1, 0, ACell, bs) then
      DrawBorderLine(ARect.Right, ARect, drawVert, bs);
    // Top border
    if GetBorderstyle(ACol, ARow, 0, -1, ACell, bs) then
      DrawBorderLine(ARect.Top, ARect, drawHor, bs);
    // Bottom border
    if GetBorderStyle(ACol, ARow, 0, +1, ACell, bs) then
      DrawBorderLine(ARect.Bottom, ARect, drawHor, bs);

    if ACell <> nil then begin
      idx := Worksheet.GetEffectiveCellFormatIndex(ACell);
      fmt := FWorkbook.GetPointerToCellFormat(idx);
      // Diagonal up
      if cbDiagUp in fmt^.Border then begin
        bs := fmt^.Borderstyles[cbDiagUp];
        DrawBorderLine(0, ARect, drawDiagUp, bs);
      end;
      // Diagonal down
      if cbDiagDown in fmt^.Border then begin
        bs := fmt^.BorderStyles[cbDiagDown];
        DrawborderLine(0, ARect, drawDiagDown, bs);
      end;
    end;
  end;
end;


procedure TOfficeGrid.DoDrawCell(Sender: TObject; ACol, ARow: Integer; RP: PRastPort; ARect: TRect);
var
  s,s1: string;
  CS: TCellStatus;
  Color: TsColor;
  Cell: PCell;
  BGPen, Pen: LongInt;
  SIdx, Idx: Integer;
  fmt: PsCellFormat;
  horAlign: TsHorAlignment;
  numfmt: TsNumFormatParams;
  TE: TTextExtent;
  fnt: TsFont;
  fst: LongWord;
  NStyle: LongWord;
  BGColor: TsColor;
  f: Double;
  TempPen: LongInt;
  SL: TStringList;
  TH, I: Integer;
begin
  Pen := -1;
  BGPen := -1;
  TempPen := -1;
  FShowFormulas := False;
  CS := CellStatus[ACol, ARow];
  if EditMode and (CS = csFocussed) then
    s := EditText
  else
  begin
    s := Cells[ACol, ARow];
    if (Length(s) > 0) and (Ord(s[1]) = 39) then
    begin
      s1 := s;
      Delete(s1, 1,1);
      if TryStrToFloat(s1, f) then
        s := s1;
    end;
  end;

  fmt := nil;
  Color := scBlack; // Black as default text color
  BGColor := scWhite; // White as default background color
  if (FWorkbook <> nil) and (FWorksheet <> nil) then
  begin
    cell := FWorksheet.FindCell(ARow - FixedRows, ACol - FixedCols);
    fnt :=FWorksheet.ReadCellFont(cell);
    Color := fnt.Color;
    idx := FWorksheet.GetEffectiveCellFormatIndex(Cell);
    BGColor := Worksheet.ReadBackgroundColor(cell);
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
          Color := numFmt.Sections[sidx].Color and scRGBMask;
      end;
    end;
  end;

  if EditMode and (CS = csFocussed) then
    HorAlign := haRight;

  if CS = csFixed then
  begin
    horAlign := haCenter;
    Color := scBlack;
    BGColor := scWhite;
  end;


  if (BGColor <> scWhite) and (((BGColor and $FF000000) = 0) or ((BGColor and $FF000000) = $FF000000)) then
  begin
    BGColor := BGColor and scRGBMask;
    BGPen := ObtainBestPenA(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, (BGColor and $ff) shl 24, (BGColor and $ff00) shl 16, (BGColor and $ff0000) shl 8, nil);
    SetAPen(RP, BGPen);
    RectFill(RP, ARect.Left, ARect.Top, ARect.Right, ARect.Bottom);
  end;

  SetDrMd(RP, JAM1);
  if (CS = csSelected) then
  begin
    if RTGMode and ((BGColor and $FF000000) = 0) then
    begin
      BGColor := CalcSelectionColor(BGColor, 64);
      TempPen := ObtainBestPenA(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, (BGColor and $ff) shl 24, (BGColor and $ff00) shl 16, (BGColor and $ff0000) shl 8, nil);
      SetAPen(RP, TempPen);
    end
    else
      SetAPen(RP, 3);
    RectFill(RP, ARect.Left, ARect.Top, ARect.Right, ARect.Bottom);
    SetAPen(RP, 0);
    GfxMove(RP, ARect.Left, ARect.Bottom);
    AGraphics.Draw(RP, ARect.Left, ARect.Top);
    AGraphics.Draw(RP, ARect.Right, ARect.Top);
    SetAPen(RP, 2);
  end
  else
  begin
    if Color <> 0 then
    begin
      Pen := ObtainBestPenA(IntuitionBase^.ActiveScreen^.ViewPort.ColorMap, (Color and $ff) shl 24, (Color and $ff00) shl 16, (color and $ff0000) shl 8, nil);
      SetAPen(RP, Pen);
    end
    else
      SetAPen(RP, 1);
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
  // Print the actual text, multiline
  SL := TStringList.Create;
  SL.Text := s;
  TextExtent(RP, PChar(s), Length(s), @TE);
  TH := ARect.Top + ARect.Height div 2;
  TH := TH - ((SL.Count  - 1) * TE.te_Height) div 2;
  if TH < ARect.Top + 5 then
    TH := ARect.Top + 5;
  TH := TH + RP^.Font^.tf_Baseline div 2;

  for i := 0 to SL.Count - 1 do
  begin
    s := SL[i];
    TextExtent(RP, PChar(s), Length(s), @TE);
    case horAlign of
      haLeft: GfxMove(RP, ARect.Left + 1, TH + i * TE.te_Height);
      haCenter: GfxMove(RP, (ARect.Left + ARect.Width div 2) - (TE.te_Width div 2), TH + i * TE.te_Height);
      haRight: GfxMove(RP, ARect.Right - TE.te_Width - 1, TH + i * TE.te_Height);
    end;
    GfxText(RP, PChar(s), Length(s));
  end;
  SL.Free;
  SetSoftStyle(RP, 0, fst);

  DrawCellBorders(RP, ACol, ARow, ARect, Cell);

  if CS = csFocussed then
  begin
    SetAPen(RP, 1);
    // Draw cursor line
    if EditMode then
    begin
      GfxMove(RP, ARect.Right - 3, ARect.Top + ARect.Height div 2 + RP^.Font^.tf_Baseline div 2 - TE.te_Height);
      AGraphics.Draw(RP, ARect.Right - 3, ARect.Top + ARect.Height div 2 + RP^.Font^.tf_Baseline div 2 + TE.te_Height div 2);
    end;
    GfxMove(RP, ARect.Left + 1, ARect.Top + 1);
    //
    AGraphics.Draw(RP, ARect.Right, ARect.Top + 1);
    AGraphics.Draw(RP, ARect.Right, ARect.Bottom);
    AGraphics.Draw(RP, ARect.Left + 1, ARect.Bottom);
    AGraphics.Draw(RP, ARect.Left + 1, ARect.Top + 1);
    //
    GfxMove(RP, ARect.Left + 2, ARect.Top + 2);
    AGraphics.Draw(RP, ARect.Right - 1, ARect.Top + 2);
    AGraphics.Draw(RP, ARect.Right - 1, ARect.Bottom - 1);
    AGraphics.Draw(RP, ARect.Left + 2, ARect.Bottom - 1);
    AGraphics.Draw(RP, ARect.Left + 2, ARect.Top + 2);
  end;
  if Pen >= 0 then
    ReleasePen(ViewPortAddress(IntuitionBase^.ActiveWindow)^.ColorMap, Pen);
  if BGPen >= 0 then
    ReleasePen(ViewPortAddress(IntuitionBase^.ActiveWindow)^.ColorMap, BGPen);
  if TempPen >= 0 then
    ReleasePen(ViewPortAddress(IntuitionBase^.ActiveWindow)^.ColorMap, TempPen);
  FShowFormulas := True;
end;

procedure TOfficeGrid.SetCellBorders(ALeft, ATop, ARight, ABottom: Integer; AValue: TsCellBorders);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  for c := ALeft to ARight do
    for r := ATop to ABottom do
      SetCellBorder(c, r, AValue);
end;

procedure TOfficeGrid.SetCellBorder(ACol, ARow: Integer; AValue: TsCellBorders);
var
  cell: PCell;
  //sr1, sc1, sr2, sc2: Cardinal;
  //gr1, gc1, gr2, gc2: Integer;
  //styles, saved_styles: TsCellBorderStyles;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(ARow - FixedRows, ACol - FixedCols);
    if Worksheet.IsMergeBase(cell) then
    begin
      {styles := Worksheet.ReadCellBorderStyles(cell);
      saved_styles := styles;
      if not (cbEast in AValue) then
        styles[cbEast] := NO_CELL_BORDER;
      if not (cbWest in AValue) then styles[cbWest] := NO_CELL_BORDER;
      if not (cbNorth in AValue) then styles[cbNorth] := NO_CELL_BORDER;
      if not (cbSouth in AValue) then styles[cbSouth] := NO_CELL_BORDER;
      Worksheet.FindMergedRange(cell, sr1, sc1, sr2, sc2);
      gr1 := sr1 - FixedRows;
      gr2 := sr2 - FixedRows;
      gc1 := sc1 - FixedCols;
      gc2 := sc2 - FixedCols;
      // Set border flags and styles for all outer cells of the merged block
      // Note: This overwrites the styles of the base ...
      ShowCellBorders(gc1,gr1, gc2,gr2, styles[cbWest], styles[cbNorth],
        styles[cbEast], styles[cbSouth], NO_CELL_BORDER, NO_CELL_BORDER);
      // ... Restores base border style overwritten in prev instruction
      Worksheet.WriteBorderStyles(cell, saved_styles);
      Worksheet.WriteBorders(cell, AValue);}
    end else
    begin
      Worksheet.WriteBorders(cell, AValue);
      FixNeighborCellBorders(cell);
      AddToRedraw(ACol, ARow);
      if ACol > 0 then
        AddToRedraw(ACol - 1 , ARow);
      AddToRedraw(ACol + 1 , ARow);
      if ARow > 0 then
        AddToRedraw(ACol, ARow - 1);
      AddToRedraw(ACol, ARow + 1);
    end;
  end;
end;

procedure TOfficeGrid.FixNeighborCellBorders(ACell: PCell);

  procedure SetNeighborBorder(NewRow, NewCol: Cardinal;
    ANewBorder: TsCellBorder; const ANewBorderStyle: TsCellBorderStyle;
    AInclude: Boolean);
  var
    neighbor: PCell;
    border: TsCellBorders;
  begin
    neighbor := Worksheet.FindCell(NewRow, NewCol);
    if neighbor <> nil then
    begin
      border := Worksheet.ReadCelLBorders(neighbor);
      if AInclude then
      begin
        Include(border, ANewBorder);
        Worksheet.WriteBorderStyle(NewRow, NewCol, ANewBorder, ANewBorderStyle);
      end else
        Exclude(border, ANewBorder);
      Worksheet.WriteBorders(NewRow, NewCol, border);
      AddToRedraw(NewCol + FixedCols, NewRow + FixedRows);
    end;
  end;

var
  fmt: PsCellFormat;
  idx: Integer;
begin
  if (Worksheet = nil) or (ACell = nil) then
    exit;

  idx := Worksheet.GetEffectiveCellFormatIndex(ACell);
  fmt := Workbook.GetPointerToCellFormat(idx);
  begin
    if ACell^.Col > 0 then
      SetNeighborBorder(ACell^.Row, ACell^.Col-1, cbEast, fmt^.BorderStyles[cbWest], cbWest in fmt^.Border);
    SetNeighborBorder(ACell^.Row, ACell^.Col+1, cbWest, fmt^.BorderStyles[cbEast], cbEast in fmt^.Border);
    if Row > 0 then
      SetNeighborBorder(ACell^.Row-1, ACell^.Col, cbSouth, fmt^.BorderStyles[cbNorth], cbNorth in fmt^.Border);
    SetNeighborBorder(ACell^.Row+1, ACell^.Col, cbNorth, fmt^.BorderStyles[cbSouth], cbSouth in fmt^.Border);
  end;
end;

procedure TOfficeGrid.DoDeleteCell(ACol, ARow: Integer);
var
  RemoveFormat: Boolean;
  Cell: PCell;
begin
  RemoveFormat := Cells[ACol, ARow] = '';
  Cell := Worksheet.GetCell(ARow - FixedRows, ACol - FixedCols);
  if Assigned(Cell) and RemoveFormat then
    WorkSheet.DeleteCell(Cell);
  inherited;
end;

procedure TOfficeGrid.CopyToClip;
var
  MS: TMemoryStream;
  Data: TClipData;
  Idx: Integer;
  i: Integer;
  SelectedCells: TsCellRangeArray;
  SL: TStringList;
  x,y: Integer;
  Line: string;
begin
  SetLength(Data, 0);
  MS := TMemoryStream.Create;
  SL := TStringList.Create;
  try
    SetLength(SelectedCells, 1);
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

    Workbook.CopyToClipboardStream(MS, sfOOXML);
    MS.Position := 0;
    if MS.Size > 0 then
    begin
      Idx := Length(Data);
      SetLength(Data, Idx + 1);
      Data[Idx].Typ := MAKE_ID('OXML');
      Data[Idx].ID := MAKE_ID('BODY');
      Data[Idx].Buffer := AllocMem(MS.Size);
      Data[Idx].BufferSize := MS.Size;
      MS.Read(Data[Idx].Buffer^, MS.Size);
    end;
    MS.Clear;
    //
    // we put together our own list instead the CSV function, (it does not copy results of forumlas)
    FShowFormulas := False;
    for y := SelectedCells[0].Row1 to SelectedCells[0].Row2 do
    begin
      Line := '';
      for x := SelectedCells[0].Col1 to SelectedCells[0].Col2 do
      begin
         if Line = '' then
           Line := Cells[x + FixedCols, y + FixedRows]
         else
           Line := Line + #9 + Cells[x + FixedCols, y + FixedRows];
      end;
      SL.Add(Line);
    end;
    FShowFormulas := True;
    SL.SaveToStream(MS);

    //WorkBook.CopyToClipboardStream(MS, sfCSV); // Original CSV write
    MS.Position := 0;
    if MS.Size > 0 then
    begin
      Idx := Length(Data);
      SetLength(Data, Idx + 1);
      Data[Idx].Typ := ID_FTXT;
      Data[Idx].ID := ID_CHRS;
      Data[Idx].Buffer := AllocMem(MS.Size);

      Data[Idx].BufferSize := MS.Size;
      MS.Read(Data[Idx].Buffer^, MS.Size);
    end;
    MS.Clear;
    if Length(Data) > 0 then
      PutToClip(Data);
  finally
    MS.Free;
    SL.Free;
  end;
end;

procedure TOfficeGrid.CutToClip;
var
  Cell: PCell;
  i: Integer;
begin
  CopyToClip;
  BeginUpdate;
  for i := 0 to SelectionCount - 1 do
  begin
    Cell := Worksheet.GetCell(Selection[i].Y - FixedRows, Selection[i].X - FixedCols);
    if Assigned(Cell) then
      WorkSheet.DeleteCell(Cell);
    Cells[Selection[i].X, Selection[i].Y] := '';
  end;
  EndUpdate;
end;

procedure TOfficeGrid.PasteFromClip;
var
  MS: TMemoryStream;
  Data: TClipData;
  ExclIdx, TextIdx, i: Integer;
  SelectedCells: TsCellRangeArray;
  Text: AnsiString;
begin
  SetLength(Data, 0);
  MS := TMemoryStream.Create;
  try
    BeginUpdate;
    if SelectionCount = 0 then
    begin
      SetLength(SelectedCells, 0);
    end
    else
    begin
      SetLength(SelectedCells, 1);
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
    end;
    Worksheet.SetSelection(SelectedCells);
    Worksheet.SelectCell(Row - FixedRows, Col - FixedCols);
    //
    ExclIdx := -1;
    TextIdx := -1;
    Data := GetFromClip(0);
    for i := 0 to High(Data) do
    begin
      if (Data[i].Typ = MAKE_ID('OXML')) and (Data[i].BufferSize > 0) then
        ExclIdx := i;
      if (Data[i].Typ = ID_FTXT) and (Data[i].BufferSize > 0) then
        TextIdx := i;
    end;
    if ExclIdx >= 0 then
    begin
      MS.Clear;
      MS.Write(Data[ExclIdx].Buffer^, Data[ExclIdx].BufferSize);
      Workbook.PasteFromClipboardStream(MS, sfOOXML, coCopyCell);
    end
    else if TextIdx >= 0 then
    begin
      MS.Clear;
      Text := PChar(Data[TextIdx].Buffer);
      Text := stringreplace(Text, #9, ';', [rfReplaceAll]);

      MS.Write(Text[1], Length(Text));
      Workbook.PasteFromClipboardStream(MS, sfCSV, coCopyCell);
    end;
    EndUpdate;
  finally
    MS.Free;
  end;
end;

initialization
  RegisterClipType(MAKE_ID('OXML'), MAKE_ID('BODY'));
end.
