unit SortWinUnit;
{$mode objfpc}{$H+}
interface

uses
  SysUtils,
  MUIClass.Base,
  MUIClass.Window,
  MUIClass.Area,
  MUIClass.Image,
  MUIClass.Group,
  fpstypes;

const
  NUMLINES = 3;

type
  TSortWin = class(TMUIWindow)
  private
    FOKClicked: Boolean;
    TopGrp: TMUIGroup;
    SortDir: TMUIRadio;
    FRowFrom, FColFrom, FRowTo, FColTo: LongInt;
    Grps: array[0..NUMLINES - 1] of TMUIGroup;
    SubGrps: array[0..NUMLINES - 1] of TMUIGroup;
    Cbxs: array[0..NUMLINES - 1] of TMUICycle;
    SortOrder: array[0..NUMLINES - 1] of TMUIRadio;
    CaseSens: array[0..NUMLINES - 1] of TMUICheckmark;
    Indexe: array of Integer;
    procedure SortDirEvent(Sender: TObject);
    procedure MakeCBX(Index: Integer);
    procedure OKClicked(Sender: TObject);
    procedure CancelClicked(Sender: TObject);
    function GetSortParams: TsSortParams;
    procedure SortChanged(Sender: TObject);
  public
    constructor Create; override;
    function Execute(ARowFrom, AColFrom, ARowTo, AColTo: LongInt): Boolean;
    property SortParams: TsSortParams read GetSortParams;
  end;

var
  SortWin: TSortWin;

implementation

constructor TSortWin.Create;
var
  i: Integer;
  Grp: TMUIGroup;
begin
  inherited;
  Horizontal := False;

  TopGrp := TMUIGroup.Create;
  TopGrp.Horiz := True;
  TopGrp.Parent := Self;

  TMUIRectangle.Create.Parent := TopGrp;

  SortDir := TMUIRadio.Create;
  SortDir.Frame := 0;
  SortDir.Horiz := True;
  SortDir.Entries := ['by Column ', 'by Row '];
  SortDir.Active := 0;
  SortDir.OnActiveChange := @SortDirEvent;
  SortDir.Parent := TopGrp;


  TMUIRectangle.Create.Parent := TopGrp;


  for i := 0 to High(Grps) do
  begin
    Grps[i] := TMUIGroup.Create;
    Grps[i].FrameTitle := IntToStr(i + 1) + '. sort criterion';
    Grps[i].Horiz := True;
    Grps[i].Parent := Self;
    //
    SubGrps[i] := TMUIGroup.Create;
    SubGrps[i].Frame := 0;
    SubGrps[i].Horiz := False;
    SubGrps[i].Parent := Grps[i];
    //
    SortOrder[i] := TMUIRadio.Create;
    SortOrder[i].Frame := 0;
    SortOrder[i].Horiz := True;
    SortOrder[i].Entries := ['Ascending', 'Descending'];
    SortOrder[i].Active := 0;
    SortOrder[i].Parent := Grps[i];
    //
    CaseSens[i] := TMUICheckmark.Create;
    CaseSens[i].Selected := False;
    CaseSens[i].Parent := Grps[i];
  end;

  for i := 0 to High(Cbxs) do
  begin
    Cbxs[i] := nil;
  end;

  Grp := TMUIGroup.Create;
  Grp.Horiz := True;
  Grp.Parent := Self;

  with TMUIButton.Create('Ok') do
  begin
    Parent := Grp;
    OnClick := @OKClicked;
  end;

  TMUIRectangle.Create.Parent := Grp;

  with TMUIButton.Create('Cancel') do
  begin
    Parent := Grp;
    OnClick := @CancelClicked;
  end;
end;

function TSortWin.Execute(ARowFrom, AColFrom, ARowTo, AColTo: LongInt): Boolean;
begin
  FOKClicked := False;
  FRowFrom := ARowFRom;
  FColFrom := AColFrom;
  FRowTo := ARowTo;
  FColTo := AColTo;
  SortDirEvent(nil);
  Show;
  repeat
    MUIApp.InputBuffered;
    SysUtils.Sleep(25);
  until Open = False;
  Result := FOKClicked;
end;

procedure TSortWin.SortChanged(Sender: TObject);
var
  Num, i: Integer;
begin
  if Sender is TMUICycle then
  begin
    Num := TMUICycle(Sender).Tag;
    if (Num < 0) or (Num > High(Grps)) then
      Exit;
    if Cbxs[Num].Active = 0 then
    begin
      // deactivate under it
      for i := Num + 1 to High(Grps) do
        Grps[i].ShowMe := False;
    end
    else
    begin
      if Num + 1 <= High(Grps) then
      begin
        MakeCBX(Num + 1);
        Grps[Num + 1].ShowMe := True;
      end;
    end;

  end;
end;


procedure TSortWin.MakeCBX(Index: Integer);
var
  Entries: TStringArray;
  Base: string;
  i: Integer;
begin
  if (Index < 0) or (Index > High(Cbxs)) then
    Exit;
  if Assigned(Cbxs[Index]) then
    Cbxs[Index].Free;
  Cbxs[Index] := TMUICycle.Create;
  Cbxs[Index].Tag := Index;
  SetLength(Entries, Length(Indexe) + 1);
  Entries[0] := 'Nothing';
  if SortDir.Active = 0 then
    Base := 'Column '
  else
    Base := 'Row ';
  for i := 0 to High(Indexe) do
  begin
    Entries[i + 1] := Base + IntToStr(Indexe[i]);
  end;
  Cbxs[Index].Entries := Entries;
  Cbxs[Index].OnActiveChange := @SortChanged;
  Cbxs[Index].Parent := SubGrps[Index];
  Grps[Index].ShowMe := True;
end;

procedure TSortWin.SortDirEvent(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to High(Cbxs) do
  begin
    Grps[i].ShowMe := False;
    Cbxs[i].Free;
    Cbxs[i] := nil;
  end;
  if SortDir.Active = 0 then
  begin
    SetLength(Indexe, FColTo - FColFrom + 1);
    for i := FColFrom to FColTo do
      Indexe[i - FColFrom] := i;
  end
  else
  begin
    SetLength(Indexe, FRowTo - FRowFrom + 1);
    for i := FRowFrom to FRowTo do
      Indexe[i - FRowFrom] := i;
  end;
  MakeCBX(0);
end;

procedure TSortWin.OKClicked(Sender: TObject);
begin
  FOKClicked := True;
  Close;
end;

procedure TSortWin.CancelClicked(Sender: TObject);
begin
  FOKClicked := False;
  Close;
end;

function TSortWin.GetSortParams: TsSortParams;
var
  Idx, i: Integer;
begin
  with Result do
  begin
    SortByCols := SortDir.Active = 0;
    Priority := spNumAlpha;
    for i := 0 to High(Grps) do
    begin
      if Assigned(Cbxs[i]) and Grps[i].ShowMe and (Cbxs[i].Active > 0) then
      begin
        Idx := Length(Keys);
        SetLength(Keys, Idx + 1);
        Keys[Idx].ColRowIndex := Indexe[Cbxs[i].Active - 1] - 1;
        if SortOrder[i].Active <> 0 then
          Keys[Idx].Options := [ssoDescending]
        else
          Keys[Idx].Options := [];
        if not CaseSens[i].Selected then
          Keys[Idx].Options := Keys[Idx].Options + [ssoCaseInsensitive]
      end;
    end;
  end;
end;

end.
