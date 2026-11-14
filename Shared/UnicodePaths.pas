unit UnicodePaths;

interface

uses
  Classes,
  SysUtils,
  Windows;

function UnicodeConcatPath(const BasePath, FileName: UnicodeString): UnicodeString;
function UnicodeCreateDirectories(const Path: UnicodeString): Boolean;
function UnicodeCompareTextOrdinal(const Left, Right: UnicodeString): Boolean;
procedure UnicodeSaveStreamToFile(Stream: TStream; const FileName: UnicodeString);

implementation

uses
  LazFileUtils,
  LazUTF8;

const
  CSTR_EQUAL = 2;

function CompareStringOrdinalW(Left: PWideChar; LeftLength: LongInt; Right: PWideChar; RightLength: LongInt; IgnoreCase: LongInt): LongInt; stdcall; external 'kernel32.dll' name 'CompareStringOrdinal';

function UnicodeConcatPath(const BasePath, FileName: UnicodeString): UnicodeString;
begin
  if BasePath = '' then
    Exit(FileName);
  Result := SysUtils.IncludeTrailingPathDelimiter(BasePath) + FileName;
end;

function UnicodeCreateDirectories(const Path: UnicodeString): Boolean;
begin
  if Path = '' then
    Exit(False);
  Result := ForceDirectoriesUTF8(UTF16ToUTF8(Path));
end;

function UnicodeCompareTextOrdinal(const Left, Right: UnicodeString): Boolean;
begin
  if (Left = '') or (Right = '') then
    Exit(Left = Right);
  Result := CompareStringOrdinalW(PWideChar(Left), Length(Left), PWideChar(Right), Length(Right), 1) = CSTR_EQUAL;
end;

procedure UnicodeSaveStreamToFile(Stream: TStream; const FileName: UnicodeString);
var
  FileHandle: THandle;
  FileStream: THandleStream;
  Buffer: array[0..65535] of Byte;
  BytesRead: LongInt;
begin
  FileHandle := FileCreateUTF8(UTF16ToUTF8(FileName));
  if FileHandle = feInvalidHandle then
    RaiseLastOSError;
  try
    FileStream := THandleStream.Create(FileHandle);
    try
      repeat
        BytesRead := Stream.Read(Buffer, SizeOf(Buffer));
        if BytesRead = 0 then
          Break;
        FileStream.WriteBuffer(Buffer, BytesRead);
      until False;
    finally
      FileStream.Free;
    end;
  finally
    FileClose(FileHandle);
  end;
end;

end.
