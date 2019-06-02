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
    procedure ListClick(Sender: TObject);
    procedure SetClick(Sender: TObject);
  end;

var
  Win: TMyWindow;     // Window
  HeadGroup
  //,FootGroup
  : TMUIGroup;
  WColRow: TMUIText;
  WText: TMUIString;
  SG: TOfficeGrid;  // A StringGrid


procedure TMyWindow.SetClick(Sender: TObject);
begin
  SG.Cells[SG.Col, SG.Row] := WText.Contents;
end;

procedure TMyWindow.ListClick(Sender: TObject);
begin
  WColRow.Contents := IntToStr(SG.Col) + ';' + IntToStr(SG.Row);
  WText.Contents := SG.Cells[SG.Col, SG.Row]
end;

begin
  // Create a Window, with a title bar text
  Win := TMyWindow.Create;
  Win.Title := 'LEU <test.ods>';

  HeadGroup := TMUIGroup.Create;
  HeadGroup.Horiz := True;
  HeadGroup.Parent := Win;

  WColRow := TMUIText.Create;
  WColRow.FixWidthTxt := '1000;1000';
  WColRow.Parent := HeadGroup;

  WText := TMUIString.Create;
  WText.Weight := 300;
  WText.OnAcknowledge := @Win.SetClick;
  WText.Parent := HeadGroup;

  with TMUIButton.Create('Set') do
  begin
    FixWidthTxt := 'Set';
    OnClick := @Win.SetClick;
    Parent := HeadGroup;
  end;

  SG := TOfficeGrid.Create;
  SG.DefCellWidth := 100;
  SG.NumCols := 12;
  SG.NumRows := 60;

  SG.OnCellFocus := @Win.ListClick;
  SG.Parent := Win;

  SG.LoadFile(ExtractFilePath(ParamStr(0)) + 'test.ods');

  MUIApp.Run;
end.
