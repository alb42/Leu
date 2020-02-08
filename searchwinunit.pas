unit searchwinunit;
{$mode objfpc}{$H+}
interface
uses
  SysUtils,
  MUIClass.Base,
  MUIClass.Area,
  MUIClass.Gadget,
  MUIClass.Group,
  MUIClass.Image,
  MUIClass.Dialog,
  MUIClass.Window,
  Officegridunit,
  fpssearch, fpspreadsheet,
  fpstypes, fpsutils;

type
  TSearchWin = class(TMUIWindow)
  private
    SG: TOfficeGrid;
    SP: TsSearchParams;
    RP: TsReplaceParams;
    ARow, ACol: LongWord;
    WS: TsWorksheet;
    FSearchString: TMUIString;
    FReplaceString: TMUIString;
    FReplaceCheck: TMUICheckMark;
    FReplaceAll: TMUICheckMark;
    FConfirm: TMUICheckMark;
    FindButton: TMUIButton;
    FindAllButton: TMUIButton;
    procedure SelectedEvent(Sender: TObject);
    procedure FindEvent(Sender: TObject);
    procedure FindNextEvent(Sender: TObject);
    function GetIsReplace: Boolean;
    procedure ConfirmEvent(Sender: TObject; AWorksheet: TsWorksheet; FRow, FCol: Cardinal; const ASearchText, AReplaceText: String; var AResult: TsConfirmReplacementResult);
  public
    constructor Create; override;

    procedure Execute(ASG: TOfficeGrid);

    property IsReplace: Boolean read GetIsReplace;

  end;

var
  SearchWin: TSearchWin;

implementation

constructor TSearchWin.Create;
var
  Grp, Grp2: TMUIGroup;
begin
  inherited;

  Title := 'Search & Replace';

  Grp := TMUIGroup.Create;
  Grp.FrameTitle := 'Search';
  Grp.Parent := Self;

  FSearchString := TMUIString.Create;
  FSearchString.Parent := Grp;

  // Replace

  Grp := TMUIGroup.Create;
  Grp.Horiz := False;
  Grp.FrameTitle := 'Replace';
  Grp.Parent := Self;

  Grp2 := TMUIGroup.Create;
  Grp2.Horiz := True;
  Grp2.Frame := 0;
  Grp2.Parent := Grp;

  FReplaceCheck := TMUICheckMark.Create;
  FReplaceCheck.OnSelected := @SelectedEvent;
  FReplaceCheck.Parent := Grp2;

  FReplaceString := TMUIString.Create;
  FReplaceString.Disabled := True;
  FReplaceString.Parent := Grp2;

  Grp2 := TMUIGroup.Create;
  Grp2.Horiz := True;
  Grp2.Frame := 0;
  Grp2.Parent := Grp;

  FReplaceAll := TMUICheckMark.Create;
  FReplaceAll.OnSelected := @SelectedEvent;
  FReplaceAll.Parent := Grp2;

  TMUIText.Create('Replace All').Parent := Grp2;

  FConfirm := TMUICheckMark.Create;
  FConfirm.Parent := Grp2;

  TMUIText.Create('Confirmation').Parent := Grp2;

  // Buttons
  Grp := TMUIGroup.Create;
  Grp.Horiz := True;
  Grp.Frame := 0;
  Grp.Parent := Self;

  FindButton := TMUIButton.Create('Find');
  FindButton.FixWidthTxt := 'Replace     ';
  FindButton.OnClick := @FindEvent;
  FindButton.Parent := Grp;

  FindAllButton := TMUIButton.Create('Find Next');
  FindAllButton.FixWidthTxt := 'Replace Next  ';
  FindAllButton.OnClick := @FindNextEvent;
  FindAllButton.Parent := Grp;

  TMUIRectangle.Create.Parent := Grp;

end;

procedure TSearchWin.SelectedEvent(Sender: TObject);
begin
  FReplaceString.Disabled := not FReplaceCheck.Selected;
  if FReplaceCheck.Selected then
  begin
    FindAllButton.Contents := 'Replace Next';
    if FReplaceAll.Selected then
    begin
      FindButton.Contents := 'Replace All';
      FindAllButton.ShowMe := False;
    end
    else
    begin
      FindButton.Contents := 'Replace';
      FindAllButton.ShowMe := True;
    end;
  end
  else
  begin
    FindButton.Contents := 'Find';
    FindAllButton.Contents := 'Find Next';
    FindAllButton.ShowMe := True;
  end;
end;

function TSearchWin.GetIsReplace: Boolean;
begin
  Result := FReplaceCheck.Selected;
end;

procedure TSearchWin.Execute(ASG: TOfficeGrid);
begin
  if Open then
    Close;
  SG := ASG;
  SG.Search.OnConfirmReplacement := @ConfirmEvent;
  Show;
end;

procedure TSearchWin.FindEvent(Sender: TObject);
begin
  WS := nil;
  ARow := 0;
  ACol := 0;
  with SP do
  begin
    SearchText := FSearchString.Contents;
    Options := [soEntireDocument];
    Within := swWorkbook;
    ColsRows := ':';
  end;
  with RP do
  begin
    ReplaceText := FReplaceString.Contents;
    Options := [];
    if FReplaceAll.Selected then
      Options := Options + [roReplaceAll];
    if FConfirm.Selected then
      Options := Options + [roConfirm];
  end;
  if IsReplace then
  begin
    if SG.Search.ReplaceFirst(SP, RP, WS, ARow, ACol) then
    begin
      SG.SelectAll(False);
      SG.SetFocus(ACol + 1, ARow + 1);
    end
    else
      ShowMessage('Nothing found.');
  end
  else
  begin
    if SG.Search.FindFirst(SP, WS, ARow, ACol) then
    begin
      SG.SelectAll(False);
      SG.SetFocus(ACol + 1, ARow + 1);
    end
    else
      ShowMessage('Nothing found.');
  end;
end;

procedure TSearchWin.FindNextEvent(Sender: TObject);
begin
  if WS = nil then
    Exit;
  if IsReplace then
  begin
    if SG.Search.ReplaceNext(SP, RP, WS, ARow, ACol) then
    begin
      SG.SelectAll(False);
      SG.SetFocus(ACol + 1, ARow + 1);
    end
    else
    begin
      ShowMessage('No more matches found');
      WS := nil;
      ARow := 0;
      ACol := 0;
    end;
  end
  else
  begin
    if SG.Search.FindNext(SP, WS, ARow, ACol) then
    begin
      SG.SelectAll(False);
      SG.SetFocus(ACol + 1, ARow + 1);
    end
    else
    begin
      ShowMessage('No more matches found');
      WS := nil;
      ARow := 0;
      ACol := 0;
    end;
  end;
end;

procedure TSearchWin.ConfirmEvent(Sender: TObject; AWorksheet: TsWorksheet; FRow, FCol: Cardinal; const ASearchText, AReplaceText: String; var AResult: TsConfirmReplacementResult);
begin
  SG.SelectAll(False);
  SG.SetFocus(ACol + 1, ARow + 1);
  case MessageBox('Search & Replace', 'Replace "' + ASearchText + '" with "' + AReplaceText + '" in cell ' + GetColString(FCol) + IntToStr(FRow + 1), ['Yes', 'No', 'Cancel']) of
    1: AResult := crReplace;
    2: AResult := crIgnore;
    else
      AResult := crAbort;
  end;
end;

end.
