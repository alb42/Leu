unit clipunit;
{$mode objfpc}{$H+}
interface
uses
  sysutils,iffparse;

type
  TClipEntry = record
    Typ: LongWord;
    ID: LongWord;
    Buffer: Pointer;
    BufferSize: Integer;
  end;
  TClipData = array of TClipEntry;

const
  ID_FTXT = 1179932756; // FTXT
  ID_CHRS = 1128813139; // CHRS

function PutToClip(Data: TClipData; ClipUnit: Integer = 0): Boolean;

function GetFromClip(ClipUnit: Integer = 0):TClipData;

procedure RegisterClipType(AType, AID: LongWord);

implementation

var
  KnownIDs: array of record
    Typ: LongWord;
    ID: LongWord;
  end;

procedure RegisterClipType(AType, AID: LongWord);
var
  Idx, i: Integer;
begin
  for i := 0 to High(KnownIDs) do
  begin
    if (KnownIDs[i].ID = AType) and (KnownIDs[i].ID = AID) then
      Exit;
  end;
  Idx := Length(KnownIDs);
  SetLength(KnownIDs, Idx + 1);
  KnownIDs[Idx].Typ := AType;
  KnownIDs[Idx].ID := AID;
end;


function PutToClip(Data: TClipData; ClipUnit: Integer = 0): Boolean;
var
  Iff: PIffHandle;
  i: Integer;
begin
  Result := False;
  Iff := AllocIff;
  if Assigned(Iff) then
  begin
    Iff^.iff_Stream := PtrUInt(OpenClipboard(ClipUnit));
    if Iff^.iff_Stream <> 0 then
    begin
      InitIffAsClip(iff);
      if OpenIff(Iff, IFFF_WRITE) = 0 then
      begin
        for i := 0 to High(Data) do
        begin
          if PushChunk(iff, Data[i].Typ, ID_FORM, IFFSIZE_UNKNOWN) = 0 then
          begin
            if PushChunk(iff, 0, Data[i].ID, IFFSIZE_UNKNOWN) = 0 then
            begin
              Result := WriteChunkBytes(iff, Data[i].Buffer, Data[i].BufferSize) = Data[i].BufferSize;
              if not Result then
                Break;
              PopChunk(iff);
            end;
          end;
        end;
        PopChunk(iff);
        CloseIff(iff);
      end;
      CloseClipboard(PClipBoardHandle(iff^.iff_Stream));
    end;
    FreeIFF(Iff);
  end;
end;

function GetFromClip(ClipUnit: Integer = 0):TClipData;
var
  Iff: PIffHandle;
  Error: LongInt;
  Cn: PContextNode;
  Buf: PChar;
  Len, i, Idx: Integer;
begin
  SetLength(Result, 0);
  Iff := AllocIff;
  if Assigned(Iff) then
  begin
    Iff^.iff_Stream := NativeUInt(OpenClipboard(ClipUnit));
    if Iff^.iff_Stream<>0 then
    begin
      InitIffAsClip(iff);
      if OpenIff(Iff, IFFF_READ) = 0 then
      begin
        for i := 0 to High(KnownIDs) do
        begin
          StopChunk(iff, KnownIDs[i].Typ, KnownIDs[i].ID);
        end;
        while True do
        begin
          Error := ParseIff(iff, IFFPARSE_SCAN);
          if (Error <> 0) and (Error <> IFFERR_EOC) then
            Break;
          Cn := CurrentChunk(Iff);
          if not Assigned(Cn) then
            Continue;
          Len := Cn^.cn_Size;
          if Len > 0 then
          begin
            Idx := Length(Result);
            SetLength(Result, Idx + 1);
            GetMem(Buf, Len + 1);
            FillChar(Buf^, Len + 1, #0);
            ReadChunkBytes(Iff, Buf, Len);
            //
            Result[Idx].Typ := cn^.cn_Type;
            Result[Idx].ID := cn^.cn_ID;
            Result[Idx].Buffer := Buf;
            Result[Idx].Buffersize := Len;
          end;
        end;
        CloseIff(Iff);
      end;
      CloseClipboard(PClipBoardHandle(iff^.iff_Stream));
    end;
    FreeIFF(Iff);
  end;
end;



initialization
  RegisterClipType(ID_FTXT, ID_CHRS);


end.
