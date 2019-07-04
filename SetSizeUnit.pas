unit SetSizeUnit;
{$mode objfpc}{$H+}
interface

uses
  Math, Classes,
  MUI,
  MUIClass.Base,
  MUIClass.Area,
  MUIClass.Group,
  MUIClass.Gadget,
  MUIClass.Window;

type
  TSetSizeWin = class(TMUIWindow)
  private
    BakCols: Integer;
    BakRows: Integer;
    EditCols: TMUIString;
    EditRows: TMUIString;
    FOnAccept: TNotifyEvent;
    FOnAutoSize: TNotifyEvent;

    function GetCols: Integer;
    procedure SetCols(AValue: Integer);
    function GetRows: Integer;
    procedure SetRows(AValue: Integer);
    procedure OKClick(Sender: TObject);
    procedure CancelClick(Sender: TObject);
    procedure AutoClick(Sender: TObject);
    procedure RevertClick(Sender: TObject);
  public
    constructor Create; override;

    procedure ShowWindow(ACols, ARows: Integer);

    property Columns: Integer read GetCols write SetCols;
    property Rows: Integer read GetRows write SetRows;
    property OnAccept: TNotifyEvent read FOnAccept write FOnAccept;
    property OnAutoSize: TNotifyEvent read FOnAutoSize write FOnAutoSize;
  end;

var
  SetSizeWin: TSetSizeWin = nil;

implementation

constructor TSetSizeWin.Create;
var
  Grp: TMUIGroup;
begin
  inherited;
  Horizontal := False;
  Title := 'Set Size';

  Grp := TMUIGroup.Create;
  with Grp do
  begin
    Horiz := True;
    Columns := 3;
    Parent := Self;
  end;

  with TMUIText.Create('Columns: ') do
  begin
    Parent := Grp;
  end;

  EditCols := TMUIString.Create;
  with EditCols do
  begin
    CycleChain := 1;
    Accept := '0123456789';
    Format := MUIV_String_Format_Right;
    IntegerValue := 0;
    Parent := Grp;
  end;

  with TMUIButton.Create('Auto') do
  begin
    CycleChain := 0;
    OnClick := @AutoClick;
    Parent := Grp;
  end;

  with TMUIText.Create('Rows: ') do
  begin
    Parent := Grp;
  end;

  EditRows := TMUIString.Create;
  with EditRows do
  begin
    CycleChain := 2;
    Accept := '0123456789';
    Format := MUIV_String_Format_Right;
    IntegerValue := 0;
    Parent := Grp;
  end;

  with TMUIButton.Create('Revert') do
  begin
    CycleChain := 0;
    OnClick := @RevertClick;
    Parent := Grp;
  end;

  Grp := TMUIGroup.Create;
  Grp.Horiz := True;
  Grp.Parent := Self;

  with TMUIButton.Create('OK') do
  begin
    CycleChain := 3;
    OnClick := @OKClick;
    Parent := Grp;
  end;

  with TMUIButton.Create('Cancel') do
  begin
    CycleChain := 4;
    OnClick := @CancelClick;
    Parent := Grp;
  end;

end;

function TSetSizeWin.GetCols: Integer;
begin
  Result := Max(2, EditCols.IntegerValue);
end;

procedure TSetSizeWin.SetCols(AValue: Integer);
begin
  EditCols.IntegerValue := Max(2, AValue);
end;

function TSetSizeWin.GetRows: Integer;
begin
  Result := Max(2, EditRows.IntegerValue);
end;

procedure TSetSizeWin.SetRows(AValue: Integer);
begin
  EditRows.IntegerValue := Max(2, AValue);
end;

procedure TSetSizeWin.ShowWindow(ACols, ARows: Integer);
begin
  if Open then
    Close;
  //
  BakCols := ACols;
  BakRows := ARows;
  Columns := ACols;
  Rows := ARows;
  Open := True;
end;

procedure TSetSizeWin.OKClick(Sender: TObject);
begin
  Close;
  if Assigned(FOnAccept) then
    FOnAccept(Self);
end;

procedure TSetSizeWin.CancelClick(Sender: TObject);
begin
  Close;
end;

procedure TSetSizeWin.AutoClick(Sender: TObject);
begin
  if Assigned(FOnAutoSize) then
    FOnAutoSize(Self);
end;

procedure TSetSizeWin.RevertClick(Sender: TObject);
begin
  Columns := BakCols;
  Rows := BakRows;
end;

initialization

end.
