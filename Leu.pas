program Leu;
{$mode objfpc}{$H+}
uses
  Types, Classes, SysUtils, intuition, agraphics, exec, utility,
  Math,
  OfficeGridUnit,
  MUIClass.Base,
  MUIClass.Window,
  MUIClass.Grid,
  MUIClass.Gadget,
  MUIClass.Group,
  MUIClass.Area,
  MUIClass.DrawPanel,
  fpstypes, fpspreadsheet, fpsallformats, fpsutils, fpsnumformat, variants;

type
  TMyWindow = class(TMUIWindow)
  public
    WColRow: TMUIText;
    WText: TMUIString;
    SG: TOfficeGrid;  // A StringGrid
    procedure ListClick(Sender: TObject);
    procedure SetClick(Sender: TObject);
    function LoadFile(AFilename: string): Boolean;


    constructor Create; override;
  end;

var
  Win: TMyWindow;     // Window



procedure TMyWindow.SetClick(Sender: TObject);
begin
  SG.Cells[SG.Col, SG.Row] := WText.Contents;
end;

procedure TMyWindow.ListClick(Sender: TObject);
begin
  WColRow.Contents := IntToStr(SG.Col) + ';' + IntToStr(SG.Row);
  WText.Contents := SG.Cells[SG.Col, SG.Row]
end;

function TMyWindow.LoadFile(AFileName: string): Boolean;
begin
  Result := True;
  SG.LoadFile(AFileName);
  Title := 'LEU <'+ExtractFileName(AFileName)+'> Columns:' + IntToStr(SG.NumCols) + ' Rows:' + IntToStr(SG.NumRows);
end;

constructor TMyWindow.Create;
var
  HeadGroup: TMUIGroup;
begin
  inherited;
  HeadGroup := TMUIGroup.Create;
  HeadGroup.Horiz := True;
  HeadGroup.Parent := Self;

  WColRow := TMUIText.Create;
  WColRow.FixWidthTxt := '1000;1000';
  WColRow.Parent := HeadGroup;

  WText := TMUIString.Create;
  WText.Weight := 300;
  WText.OnAcknowledge := @SetClick;
  WText.Parent := HeadGroup;

  with TMUIButton.Create('Set') do
  begin
    FixWidthTxt := 'Set';
    OnClick := @SetClick;
    Parent := HeadGroup;
  end;

  SG := TOfficeGrid.Create;
  SG.DefCellWidth := 100;
  SG.NumCols := 12;
  SG.NumRows := 60;

  SG.OnCellFocus := @ListClick;
  SG.Parent := Self;
end;

procedure StartMe;

begin
  // Create a Window, with a title bar text
  Win := TMyWindow.Create;
  Win.Title := 'LEU <> Columns:' + IntToStr(Win.SG.NumCols) + ' Rows:' + IntToStr(Win.SG.NumRows);

  if ParamCount > 0 then
    Win.LoadFile(ParamStr(1))
  else
    Win.LoadFile(ExtractFilePath(ParamStr(0)) + 'test.ods');

  MUIApp.Run;
end;

begin
  StartMe;
end.
