unit FormatWinUnit;
{$mode objfpc}{$H+}
interface
uses
  Classes, SysUtils,
  locale,
  fgl,
  fpstypes, fpsnumformat,
  typinfo, Math,
  MUIClass.Area,
  MUIClass.StringGrid,
  MUIClass.Window,
  MUIClass.Gadget,
  MUIClass.Group,
  OfficeGridUnit;

const
  NumFormats = 8;
  FormatTypes: array[0..NumFormats - 1] of TsNumberFormat = (nfFixed, nfPercentage, nfCurrency, nfLongDate, nfLongTime, nfExp, nfCustom, nfText);
  FormatStrings: array[0..NumFormats - 1] of string = ('Fixed', 'Percentage', 'Currency', 'Date', 'Time', 'Exponential', 'Boolean', 'Text');


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
  BoolFormats: string = 'BOOLEAN';
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

    function GetExampleText(FmtStr: string): string;
    procedure ClickEvent(Sender: TObject);
    procedure ClickFormatEvent(Sender: TObject);
    procedure AckEdit(Sender: TObject);
  public
    SG: TOfficeGrid;
    Cell: TPoint;
    NewFmt: TsNumberFormat;
    constructor Create; override;
    destructor Destroy; override;
  end;

var
  FormatWin: TFormatWin;
implementation

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

  FormatEdit := TMUIString.Create;
  FormatEdit.OnAcknowledge := @AckEdit;
  FormatEdit.Parent := Grp;

  {SL:= TStringList.Create;
  BuildCurrencyFormatList(SL, False, 12.34, DefMoneySymb);
  writeln('SL: ', SL.Text);
  SL.Free;}
  //writeln('BuildCurrencyFormatString: ', BuildCurrencyFormatString( ));
end;

destructor TFormatWin.Destroy;
begin
  SubFormats.Free;
  inherited;
end;


function TFormatWin.GetExampleText(FmtStr: string): string;
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
    if Parser.IsDateTimeFormat then
      Result := ConvertFloatToStr(Now(), mClass, SG.Workbook.FormatSettings)
    else
      Result := ConvertFloatToStr(-123.456, mClass, SG.Workbook.FormatSettings);
    if Pos('RED', FmtStr) > 0 then
      Result := #27'b' +  Result;

  finally
    parser.Free;
    mClass.Free;
  end;
end;

procedure TFormatWin.ClickEvent(Sender: TObject);
var
  i: Integer;
begin
  NewFmt := FormatTypes[FormatList.Row];
  SubFormats.Clear;
  case NewFmt of
    nfFixed: SubFormats.Text := FixedFormats;
    nfPercentage: SubFormats.Text := PercentFormats;
    nfCurrency: SubFormats.Text := MoneyFormats;
    nfLongDate: SubFormats.Text := DateFormats;
    nfLongTime: SubFormats.Text := TimeFormats;
    nfExp: SubFormats.Text := ExpFormats;
    nfCustom: SubFormats.Text := BoolFormats;
    nfText: SubFormats.Text := TextFormats;
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
      SubFormatList.Cells[0, i] := GetExampleText(SubFormats[i]);
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
begin
  SG.Worksheet.WriteNumberFormat(Cell.Y - 1, Cell.X - 1, NewFmt, FormatEdit.Contents);
  SG.RedrawCell(Cell.X, Cell.Y);
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
