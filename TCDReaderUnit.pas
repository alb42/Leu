unit TCDReaderUnit;
{$mode objfpc}{$H+}

interface

uses
  SysUtils, Classes, Math, fpspreadsheet, fpstypes, fpsutils;

procedure ReadTCD(Workbook: TsWorkbook; Filename: string);

implementation

type
  TTCDHeader = packed record
    th_Type: Byte;
    th_Size: Word;
  end;

  TTCDVersion = packed record
    tv_Version: Byte;
    tv_Revision: Byte;
  end;
  PTCDVersion = ^TTCDVersion;

const
  TC_VERSION = 1;
  TC4_VERSION = 9;
  TC5_VERSION = 10;

type
  TCTypes = (
    FILE_END,
    FILE_START,     (* 12:TURBOCALC\8, floatmode, 0.b *)
    FILE_VERSION,   (* 2:(version, revision): see above! *)
    FILE_PASSWORD,
                   (* flags.b (0: ask_password), flags2.b 2+3*txt
               * pwopenread_txt
               * pwwriteread_txt
               * pwchange_txt *)
    FILE_OPT0,      (* 6:stdwidth.w, stdheight.w, flags1.b 0.b *)
    FILE_OPT1,      (* 8+2_txt: decimalpoint.b, 1000sep.b, datumsep.b, waehrung.b
               * sheet_waehrungprefix_txt, sheet_waehrungsuffix_txt
               * dateorder.b, localeflags.b, 0.b, 0.b *)
    FILE_WIDTH,     (* x.l, width.w only if differs from std! *)
    FILE_HEIGHT,    (* y.l, height.w *)
    FILE_CELL,      (* x.w,y.w,
               * type.b,
               * ieee.l (only for type==FLOAT!) see note at the end
               * (data1.l, data2.l) - for type==TEXT: text_txt, (skipped for type==EMPTY)
               * formel_txt (not for type==EMPTY)
               * format.Format_Size *)
    FILE_LCELL,     (* dito, but x.l, y.l *)
    FILE_NAME,      (* name_txt
               * typ.b
               * TYPE_TEXT: text_txt
               * TYPE_FORMEL_x: formel_txt
               * else: data1.l, data2.l *)
    FILE_OPT2,      (* print options: 42+3*txt!
               * printflags.b, printflags2.b
               * printLM.w, RM.w, UM.w, BM.w
               * printwidth.w, printheight.w
               * printfile_txt, printheader_txt, printfooter_txt
               * sheet_titlex0.l, sheet_titlex1.l, sheet_titley0.l, sheet_titley1.l
               * printflags3.b, 0.b, 0.b, 0.b
               * scalex.w, scaley.w, 0.w, 0.w *)
    FILE_WINDOW,    (* add this at the end of the file!!!
               * 14.w: x1.w, y1.w, x2.w, y2.w
               * flags.b, pad.b, pad.l *)
    FILE_FONTS,     (* 31 times: (font_txt, size.w, pad.w | or 0.w for unused) *)
    FILE_SCREEN,    (* flag.b (0=off, 1=wb-clone, 2=on)
               * width.w,height.w, depth.w, id.l, smartrefresh, pad.b *)
    FILE_COLOR,     (* num.w
               * r.b,g.b,b.b,pad.b *)
    FILE_DIAGRAM,   (* (72+2*DIAGRAM_MAXY+(4+2*DIAGRAM_MAXY)_txt+6_fnt)
               * x1.w, <1.w, x2.w, y2.w - window-dim
               * winflags.b (bit0: hidden), pad.b
               * name_txt
               * x1.l, y1.l, x2.l, y2.l - data-range
               * type.b, realtype.b, flags2.b, flags.b
               * min.l, min2.l
               * printflags.b, printpad.b, printwidth.w, printheight.w
               * max.l, max2.l (sorry, min & max should be together, but...)
               **
               * titleflags.b, pad.b
               * title_txt, title_font style.b pad.b *3
               **
               * legendeflags.b, style.b legende_fnt
               * DIAGRAM_MAXY*legende_txt
               **
               * patternflags.b, pad.b
               * colorpattern.w * DIAGRAM_MAXY
               **
               * achseflags.b, style.b achse_fnt
               * DIAGRAM_MAXY*achse_txt
               **
               * yachseflags.b, yachsepad.b, yachse_fnt
               * yachseticknum.w, yachsesubticknum.w *)
    FILE_STDFONTS,    (* font_txt, size.w, pad.w | or 0.w for unused *)
    FILE_PATTERNS,    (* 4*PATTERNMAXSIZE +_txt
               * std_pfad_txt
               * pattern0.4*PATTERNMAXSIZE *)
    FILE_COLUMNFLAGS,
    FILE_ROWFLAGS,    (* 12,row/column.l, Format.Format_Size *)
    FILE_SHEETSIZE,   (* 12:sheet_limitx/y.l, pad.l *)
    FILE_SYSTEMFONTS, (* font_txt, size.w, pad.w   or 0.w for unused for text
               * font_txt, size.w, pad.w   or 0.w for unused for menu *)
    FILE_FROZEN,    (* V3! - should be immedialty after FILE_WINDOW!!!
               * 24,freezex,freezey,freezepix,freezepiy,0,0  - all .l
               * this is only written, if there is freezing!!! *)
    FILE_SAVEOPT,   (* 20+txt,saveflags.b,0.b, 0.w autosavetime.w, 0.w, 0.l
               * sheet_autosavepath_txt, 0.l, 0.l *)
    FILE_CRYPT,     (* checkkey.w
               * flags0,1.b - reserved
               * long0,long1 for padding (filled with scratch!) *)
    FILE_DIAGRAM2,    (* dia_legendex1..y2 (16)
               * dia_achsex1..y2   (16)
               * symbolmode.b, pad.b, width.w height.w, pad.w (8)
               * kuchenexplode.b (MAXEXPLODE)
               * iffdepth/width/height.w  (6)
               * refreshflags0/1.b (2)
               * 2*pad.l (8)
               ** 3.10: (FEATURE_CHARTXY)
               * xachseflags, xachsestyle, xachseticknum, xachsesubticknum (6)
               * xmin.l, xmin2.l, xmax.l, xmax2.l (16) *)
    FILE_GLOBALFLAGS, (* GlobalFlags.l (0..3), GlobalFlags_autoopendir_txt, 0.w
               * Preview_flags.b 0.b, Preview_Width/Height/Depth.w, Preview_screenmode *)
    FILE_OBJECT,    (* (36+3*txt)
               * x1,y1,x2,y2.l (16)
               * objecttype.l (-1=extern) (4)
               * objectclass_txt, name_txt, macro_txt (3*txt)
               * backcolor, pattern, frame, pad.b (as long! 4)
               * flags0,1,2,3.b (as long! 4)
               * pad1.l, pad2.l (8) *)
    FILE_STDCHART,    (* (2+Stdchart_save_firstfont+3*fnt)
               * reserved.w (2)
               * StdChart (StdChart_save_firstfont)
               * fonts (3*fnt) *)
    FILE_OPT3,      (* (26)
               * sheet_printiffdept/with/height.w, undo_depth.w, printiff_scale (10)
               * 0.l, 0.l, 0.l, 0.l (16) *)
    FILE_LASTFILES,   (* only used for &#039;lastfiles&#039;
               * flags0/1/2/3.b (4)
               * filename_txt *)
    FILE_CURSOR,    (* (28) immediatly after FILE_WINDOW (or _FROZEN)
               * cellx/y.l
               * blockx1/y1/x2/y2.l (if blockx1=-1.l, then no block selected!)
               * flags0,1,2,3.b (reserved) *)
    FILE_STARTUPOPTIONS,  (* these options *must* follow immediatly
                 * after FILE_START & will be checked
                 * by the File_CheckType-routine!
                 * flags0-7 ; (8) reserved up to now! *)
    FILE_TURBOCALCOWNER,  (* (12+txt)
                 * reserved.l (0)       ;
                 * tc-serial-id_txt[6]   ; depends on TC-version!!!
                 * username_txt         ;
                 * reserved.w (0)   ; reserved for company-name! *)
    FILE_FILEINFO,    (* (20 + 4*txt)
               * reserved.l (0)
               * creation_date.l
               * creation_time.l
               * version.l
               * workingtime.l
               * author_txt
               * title_txt
               * subject_txt *)
   FILE_UNKNOWN
  ) ;

TCellType =
  (
    TCT_Empty,
    TCT_No,
    TCT_Float,
    TCT_Int,
    TCT_Date,
    TCT_Time,
    TCT_Bool,
    TCT_Text,
    TCT_unknown,
    TCT_Error
  );

TTCDCell = packed record
  tc_Color0: Byte;    //* 0 = std., values 1-63 */
  tc_Color1: Byte;    //* "                    */
  tc_Border: Byte;
  tc_Font: Byte;       //* 0 = std., bits 0-2: style, bits 3-7: font number, see FILE_FONTS */
  tc_TextFormat: Byte; //* 0 = std., -$FF (only for row & column_format): invalid format! */
  tc_Flags: Byte;
  //tc_Label: array[0..5] of Byte;   //* ????, there are only 6 bytes in the memory left so I guess this was a union before, wrongly converted */
end;
PTCDCell = ^TTCDCell;

const
  TC_COLOR_MASK = $3f;
  TC_PATTERN_MASK = $c0;
  TC_PATTERN_LSHIFT = $6;
  TC_BORDER_LEFT = $03;    // 00=off, 01= thin, 10 = medium, 11= bold
  TC_BORDER_RIGHT = $0c;
  TC_BORDER_UP = $30;
  TC_BORDER_DOWN = $c0;
  TC_FONT_LSHIFT = 3;
  TC_STYLE_ITALIC = 2;
  TC_STYLE_BOLD = 1;
  TC_STYLE_UNDERLINED = 0;
  TC_STYLE_MASK = $07;
  TCF_HIDDEN = 7;
  TCF_PROTECTED = 6;
  TCF_HIDEFORMEL = 5;
  TCF_VALIGNMENT = $18;   // 00=std, 01=upper, 10=mid, 11=lower
  TCF_ALIGNMENT = $07;

function TCDFunctionToStr(B: Byte): string;
begin
  case b of
    $01: Result := 'EXP';   // EXP
    $02: Result := 'LN';    // LN
    $03: Result := 'LOG10'; // LOG10
    $04: Result := 'LOG';   // LG
    $05: Result := 'Ln';    // LOG
    $06: Result := 'SQRT'; // Quadrat
    $07: Result := 'SUMSQ'; // Quadrat
    $08: Result := 'FACT';
    $09: Result := 'PI()';
    $0A: Result := 'SINH';
    $0B: Result := 'COSH';
    $0C: Result := 'TANH';
    $0D: Result := 'SIN';
    $0E: Result := 'COS';
    $0F: Result := 'TAN';
    //
    $10: Result := 'ASIN'; // ARCSIN
    $11: Result := 'ACOS'; // ARCCOS
    $12: Result := 'ATAN'; // ARCTAN
    $13: Result := 'DEGREES'; // WINKEL
    $14: Result := 'RADIANS'; // BOGEN
    $15: Result := 'IF'; // WENN
    $17: Result := 'MOD';
    $18: Result := 'INT'; // GANZZAHL
    $1a: Result := 'ABS';
    $1b: Result := 'SIGN';
    $1c: Result := 'ROUND'; // RUNDEN
    $1d: Result := 'NOT'; // NICHT
    $1e: Result := 'RAND()'; // ZUFALLSZAHL
    $1F: Result := 'TRUE'; // WAHR
    //
    $20: Result := 'FALSE'; // FALSCH
    $21: Result := 'TEXT("",'; // TEXT
    $22: Result := 'LEFT'; // LEFT
    $23: Result := 'RIGHT'; // RIGHT
    $24: Result := 'MID'; // Mitte
    $25: Result := 'AVERAGE'; // MITTELWERT
    $26: Result := 'CHAR'; // CHAR
    $27: Result := 'LEN'; // LÄNGE
    $28: Result := 'LOWER'; // KLEIN
    $29: Result := 'PROPER'; // GROSS2 missing PROPER
    $2A: Result := 'UPPER'; // GROSS
    $2B: Result := 'CODE'; // CODE
    $2D: Result := 'TRIM'; // TRIM
    $2C: Result := 'REPT'; // REPT
    $2E: Result := ''; // CLEAN missing CLEAN
    $2F: Result := 'VALUE'; // WERT
    //
    $49: Result := 'COUNTA'; // ANZAHL2
    $4A: Result := 'COUNT'; // ANZAHL
    $4B: Result := 'MIN'; // MIN
    $4C: Result := 'MAX'; // MAX
    $4E: Result := 'SUM'; // SUMME
    $4F: Result := 'PRODUCT'; // PRODUKT

    //
    $50: Result := 'MID'; // MITTE
    $51: Result := 'AND'; // UND
    $52: Result := 'OR'; // Oder
    $53: Result := ''; // XOR
    $55: Result := 'ISNUMBER'; // ISTZAHL
    $56: Result := 'ISTEXT'; // ISTTEXT
    $57: Result := ''; //   ISTDATUM  MISSING ISDATE
    $58: Result := ''; //   ISTZEIT   MISSING ISTIME
    $59: Result := 'ISBLANK'; // ISTLEER
    //
    $72: Result := ''; // INSTRING   missing INSTR
    $73: Result := ''; // SPIEGELN   missing
    $74: Result := 'SHIFTL'; // SHIFTL   missing ...
    $75: Result := 'SHIFTR'; // SHIFTR   missing ...
    $76: Result := ''; // KOMPRIMIEREN missing ...
    $79: Result := ''; // HEX   missing DEC2HEX
    $7B: Result := 'STDEV'; // STABW
    $7D: Result := 'VAR'; // VARIANZ


    else Result := 'U' + HexStr(b, 2) + 'U';
  end;

end;
type
  TCellColor = record
    Col, Row: LongInt;
    bgPen: LongInt;
    TextPen: LongInt;
  end;
  TCellColors = array of TCellColor;

var
  CellColors: TCellColors;
  Colors: array of LongWord;

procedure AddColors(ACol, ARow, ABGPen, ATextPen: LongInt);
var
  Idx: Integer;
begin
  Idx := Length(CellColors);
  SetLength(CellColors, Idx + 1);
  CellColors[Idx].Col := ACol;
  CellColors[Idx].Row := ARow;
  CellColors[Idx].bgPen := ABGPen;
  CellColors[Idx].TextPen := ATextPen;
end;


procedure ReadTCD(Workbook: TsWorkbook; Filename: string);
var
  f: TFileStream;
  Header: TTCDHeader;
  got, Len: Integer;
  Buffer: Pointer;
  HType: TCTypes;
  p, p2: PByte;
  ieeemode: Byte;
  col, row: Integer;
  c1,r1: Integer;
  CellType: TCellType;
  Cell: PTCDCell;
  Txt: string;
  i: Integer;
  fVal: Double;
  iVal: Integer;
  qw: QWord;
  ColAbs, RowAbs: boolean;
  WS: TsWorksheet;
  wcell: PCell;
  Formula: string;
  CellBorder: TsCellBorders;
  CellFontStyle: TsFontStyles;
  r,g,b: byte;
  c: Char;
begin
  SetLength(Colors, 0);
  SetLength(CellColors, 0);
  f := TFileStream.Create(Filename, fmOpenRead);
  //
  WS := Workbook.AddWorksheet(ExtractFilename(Filename));
  try
    // Read Header
    repeat
      //writeln('file pos: ', f.position, ' length: ', f.size);
      got := f.Read(Header, SizeOf(Header));
      if got < SizeOf(Header) then
        break;
      //writeln('read Header: ', got);

      try
        HType := TCTypes(Header.th_type);
      except
        HType := FILE_UNKNOWN;
      end;

      //writeln('Header type: ', HType);

      {$ifdef ENDIAN_LITTLE}
      Header.th_size := SwapEndian(Header.th_size);
      {$endif}
      //writeln('Read Buffer: ', Header.th_size);
      Buffer := AllocMem(Header.th_size);
      got := f.Read(Buffer^, Header.th_size);
      if got < Header.th_size then
        break;
      //writeln('  Data size: ', got);
      p := Buffer;
      case HType of
        FILE_START: begin
          Inc(p, 11);
          ieeemode := P^;
          //writeln('  ieeee mode ', ieeemode);
        end;
        FILE_END: begin
          break;
        end;
        FILE_VERSION: begin
          //writeln('  Version: ', PTCDVersion(Buffer)^.tv_Version, '.', PTCDVersion(Buffer)^.tv_Revision);
        end;

        // the actual cells
        FILE_CELL,
        FILE_LCELL:
        begin
          if HType = FILE_CELL then
          begin
            {$ifdef ENDIAN_LITTLE}
            col := SwapEndian(PWord(p)^);
            {$else}
            col := PWord(p)^;
            {$endif}
            Inc(p, 2);
            {$ifdef ENDIAN_LITTLE}
            Row := SwapEndian(PWord(p)^);
            {$else}
            Row := PWord(p)^;
            {$endif}
            Inc(p, 2);
          end
          else
          begin
            {$ifdef ENDIAN_LITTLE}
            col := SwapEndian(PLongInt(p)^);
            {$else}
            col := PLongInt(p)^;
            {$endif}
            Inc(p, 4);
            {$ifdef ENDIAN_LITTLE}
            Row := SwapEndian(PLongInt(p)^);
            {$else}
            Row := PLongInt(p)^;
            {$endif}
            Inc(p, 4);
          end;
          //writeln('  cell: ', col, ';', row);
          wcell := WS.GetCell(row, col);
          if not Assigned(wcell) then
            exit;
          try
            CellType := TCellType(P^);
          except
            CellType := TCT_No;
          end;
          //writeln('  cell type: ', CellType);
          Inc(P);
          case CellType of
            TCT_Float: begin
              Inc(P, 4);
              {$ifdef ENDIAN_LITTLE}
              QW := SwapEndian(PQWord(P)^);
              fVal := PDouble(@QW)^;
              {$else}
              fVal := PDouble(P)^;
              {$endif}
              Inc(P, 8);
              //writeln('  got float value: ', fVal);
              WS.WriteNumber(row, col, fVal);
            end;
            TCT_Int: begin
              {$ifdef ENDIAN_LITTLE}
              iVal := SwapEndian(PLongInt(P)^);
              {$else}
              iVal := PLongInt(P)^;
              {$endif}
              Inc(P, 8);
              //writeln('  got integer value: ', iVal);
              WS.WriteNumber(row, col, iVal);
            end;
            TCT_DATE: begin
              {$ifdef ENDIAN_LITTLE}
              iVal := SwapEndian(PLongInt(P)^);
              {$else}
              iVal := PLongInt(P)^;
              {$endif}
              Inc(P, 8);
              //writeln('  got Date value: ', iVal);
            end;
            TCT_Time: begin
              {$ifdef ENDIAN_LITTLE}
              iVal := SwapEndian(PLongInt(P)^);
              {$else}
              iVal := PLongInt(P)^;
              {$endif}
              Inc(P, 8);
              //writeln('  got Date value: ', iVal);
            end;
            TCT_Bool: begin
              {$ifdef ENDIAN_LITTLE}
              iVal := SwapEndian(PLongInt(P)^);
              {$else}
              iVal := PLongInt(P)^;
              {$endif}
              Inc(P, 8);
              //writeln('  got Bool value: ', iVal);
            end;
            TCT_Text: begin
              // next word is the text length
              {$ifdef ENDIAN_LITTLE}
              Len := SwapEndian(PWord(p)^);
              {$else}
              Len := PWord(p)^;
              {$endif}
              Inc(p, 2);
              //writeln('  Text Len: ', Len);
              SetLength(Txt, Len);
              Move(P^, Txt[1], Len);
              //writeln('  got text: "'+Txt+'"');
              Inc(P, Len);
              Txt := AnsitoUTF8(Txt);
              WS.WriteText(wcell, Txt, []);
            end;
          end;
          // read the formula
          if CellType <> TCT_Empty then
          begin
            {$ifdef ENDIAN_LITTLE}
            Len := SwapEndian(PWord(p)^);
            {$else}
            Len := PWord(p)^;
            {$endif}
            Inc(P, 2);
            //
            if Len > 0 then
            begin
              SetLength(Txt, Len);
              Move(P^, Txt[1], Len);
              //writeln('  it has a formula: ');
              i := 1;
              p2 := p;
              Formula := '';
              while i <= Len do
              begin
                case P2^ of
                  $8..$b: // Cell reference
                    begin
                      colAbs := (p2^ and 1) = 0;
                      rowAbs := (p2^ and 2) = 0;
                      Inc(p2); Inc(i);
                      {$ifdef ENDIAN_LITTLE}
                      C1 := SwapEndian(PSmallInt(p2)^);
                      {$else}
                      C1 := PSmallInt(p2)^;
                      {$endif}
                      Inc(P2, 2); Inc(i, 2);
                      {$ifdef ENDIAN_LITTLE}
                      R1 := SwapEndian(PSmallInt(p2)^);
                      {$else}
                      R1 := PSmallInt(p2)^;
                      {$endif}
                      Inc(P2, 2); Inc(i, 2);
                      if not RowAbs then
                        r1 := Row + R1;
                      if not ColAbs then
                        c1 := Col + C1;
                      //if (Col = 1) and (Row=17) then
                      //  writeln('    col: ', c1, ' Row: ', r1);
                      Formula := Formula + GetColString(c1) + IntToStr(r1 + 1);
                     //write('  ', HexStr(Ord(TXT[i]), 2), ' ', txt[i],'     ');
                    end;
                  $5,$4:
                    begin
                      Inc(P2); Inc(i);
                      //writeln('Funktion: ' + GetColString(Col) + IntToStr(Row) + ' -> $', HexStr(P2^, 2));
                      //writeln('    ' + TCDFunctionToStr(P2^));
                      Formula := Formula + TCDFunctionToStr(P2^);

                      Inc(P2); Inc(i);
                    end;
                  else
                  begin
                    //if (Col = 1) and (Row=17) then
                    //  writeln('    ', char(P2^) + ' ($' + HexStr(P2^, 2) +')');
                    c := Char(P2^);
                    if c = ';' then
                      c := ',';
                    Formula := Formula + c;
                    Inc(P2); Inc(i);
                  end;
                end;
              end;
              //WS.WriteText(row, col, 'FKT: ' +  Formula);
              try
                //if (Col = 1) and (Row=17) then
                //  writeln('Formula: ', Formula);
                if (Formula <> '') and (Formula[1] = '=') then
                  WS.WriteFormula(row, col, Copy(Formula, 2, Length(Formula)), false);
              except
                On E: Exception do
                begin
                  WS.Formulas.DeleteFormula(WCell^.Row, WCell^.Col);
                  WS.WriteText(Row, Col, Copy(Formula, 2, Length(Formula)));
                end;
              end;
            end;
            //writeln();
            Inc(P, Len);
          end;
          Cell := PTCDCell(P);
          Inc(P, SizeOf(TTCDCell));
          //writeln('  color BG Pen: ',  Cell^.tc_Color0 and TC_COLOR_MASK, ' Text Pen: ', Cell^.tc_color1 and TC_COLOR_MASK);
          if ((Cell^.tc_Color0 and TC_COLOR_MASK) <> 0) or ((Cell^.tc_color1 and TC_COLOR_MASK) <> 0) then
            AddColors(col, row, Cell^.tc_Color0 and TC_COLOR_MASK, Cell^.tc_color1 and TC_COLOR_MASK);
          //write('  Border:');
          CellBorder := [];
          if (cell^.tc_Border and TC_Border_Left) <> 0 then
            CellBorder := CellBorder + [cbWest];
            //write(' Left');
          if (cell^.tc_Border and TC_Border_Up) <> 0 then
            CellBorder := CellBorder + [cbNorth];
            //write(' Top');
          if (cell^.tc_Border and TC_Border_Right) <> 0 then
            CellBorder := CellBorder + [cbEast];
            //write(' Right');
          if (cell^.tc_Border and TC_Border_Down) <> 0 then
            CellBorder := CellBorder + [cbSouth];
            //write(' Bottom');
          WS.WriteBorders(row, col, CellBorder);
          //writeln('');
          CellFontStyle := [];
          //write('  Font Styles:');
          if (cell^.tc_Font and (1 shl TC_STYLE_BOLD)) <> 0 then
            CellFontStyle := CellFontStyle + [fssBold];
            //write(' Bold');
          if (cell^.tc_Font and (1 shl TC_STYLE_ITALIC)) <> 0 then
            CellFontStyle := CellFontStyle + [fssItalic];
            //write(' Italic');
          if (cell^.tc_Font and (1 shl TC_STYLE_UNDERLINED)) <> 0 then
            CellFontStyle := CellFontStyle + [fssUnderline];
            //write(' underlined');
          //writeln('');
          WS.WriteFontStyle(row, col, CellFontStyle);
          //writeln('  ALignment: ', cell^.tc_Flags);
          case cell^.tc_Flags and TCF_Alignment of
            1: WS.WriteHorAlignment(Row, Col, haLeft); //writeln('  left align');
            2: WS.WriteHorAlignment(Row, Col, haRight); //writeln('  right align');
            3: WS.WriteHorAlignment(Row, Col, haCenter); //writeln('  center align');
            0: ;//writeln('  default align');
          end;
          //writeln('  vertical Alignment: ', (cell^.tc_Flags and TCF_VALIGNMENT) shr 3);
          case (cell^.tc_Flags and TCF_VALIGNMENT) shr 3 of
            1: WS.WriteVertAlignment(Row, Col, vaTop); //writeln('  top valign');
            2: WS.WriteVertAlignment(Row, Col, vaCenter); //writeln('  center valign');
            3: WS.WriteVertAlignment(Row, Col, vaBottom); //writeln('  down valign');
            0: ;//writeln('  default valign');
          end;
          // Text format
          //writeln('  TextFormat: ', cell^.tc_TextFormat);
          //writeln('  used: ', NativeUInt(P) - NativeUInt(Buffer), ' Size: ', Header.th_size);
        end;
        FILE_COLOR:
        begin
          //writeln('File color');
          {$ifdef ENDIAN_LITTLE}
          Len := SwapEndian(PWord(p)^);
          {$else}
          Len := PWord(p)^;
          {$endif}
          Inc(p, 2);
          //writeln('Len: ', Len);
          SetLength(Colors, Len);
          for i := 0 to Len - 1 do
          begin
            {$ifdef ENDIAN_LITTLE}
            iVal := SwapEndian(PWord(p)^);
            {$else}
            iVal:= PWord(p)^;
            {$endif}
            //if i < 10 then
            //  writeln(i,'. Raw: ', HexStr(iVal, 4));
            Inc(p, 2);
            // make 8 bit per color
            {r := P^; Inc(P);
            g := P^; Inc(P);
            b := P^; Inc(P);
            Inc(P);}
            //if i < 10 then
              //writeln(i,'. Raw: ', HexStr(r, 2),' ', HexStr(g,2), '  ', HexStr(b,2));
            b := (iVal shr 8) and $f;
            g := (iVal shr 4) and $f;
            r := iVal and $f;
            r := r or (r shl 4);
            g := g or (g shl 4);
            b := b or (b shl 4);
            // create the color
            Colors[i] := (r shl 16) or (g shl 8) or b;
            if i < 10 then
            //writeln(i,'. color: ', HexStr(Colors[i], 8));
          end;
        end;
//
      end;
      FreeMem(Buffer);
    until false;
    for i := 0 to High(CellColors) do
    begin
      //writeln('set col, row: ', CellColors[i].col,', ', CellColors[i].row, ' to pen: ', CellColors[i].BgPen);
      if InRange(CellColors[i].BgPen, 1, Length(Colors)) then
        WS.WriteBackgroundColor(CellColors[i].row, CellColors[i].col,  Colors[CellColors[i].BgPen - 1]);
      if InRange(CellColors[i].TextPen, 1, Length(Colors)) then
        WS.WriteFontColor(CellColors[i].row, CellColors[i].col,  Colors[CellColors[i].TextPen - 1]);
    end;
  finally
    F.Free;
  end;
end;


end.
