unit Functions;

interface

uses
  Classes,
  Constants,
  Paths,
  Registry,
  RtlConsts,
  ShlObj,
  SysUtils,
  Windows;

type

  { TFunctions }

  TFunctions = class
  private
    class procedure HandleException(hWnd: HWND; Obj: TObject; Addr: Pointer; FrameCount: Longint; Frames: PPointer); static; overload;
  public
    class procedure HandleException(hWnd: HWND; E: Exception); static; overload;

    // Wrappers for windows functions
    class function MessageBox(hWnd: HWND; Text: UnicodeString; Caption: UnicodeString; uType: UINT): LongInt; static;
    class function GetSpecialFolder(const csidl: ShortInt): string; static;
    class function GetTempPath: string; static;

    // Functions only used by Launcher/Setup
    class procedure RunUninstall; static;
    class function GetFileVersion(const FileName: string): string; static;
  end;

implementation

class procedure TFunctions.HandleException(hWnd: HWND; Obj: TObject; Addr: Pointer; FrameCount: Longint; Frames: PPointer);
var
  i: LongInt;
  SL: TStringList;
begin
  SL := TStringList.Create;
  try
    if Obj is Exception then
      SL.Add('%s: %s'.Format([Exception(Obj).ClassName, Exception(Obj).Message]));
    SL.Add('  %s'.Format([StringReplace(Trim(BackTraceStrFunc(Addr)), '  ', ' ', [rfReplaceAll])]));
    for i := 0 to FrameCount - 1 do
      SL.Add('  %s'.Format([StringReplace(Trim(BackTraceStrFunc(Frames[i])), '  ', ' ', [rfReplaceAll])]));

    TFunctions.MessageBox(0, 'An unexpected exception occurred.'#13#10'%s'.Format([SL.Text]), SMsgDlgError, MB_ICONERROR);
  finally
    SL.Free;
  end;
end;

class procedure TFunctions.HandleException(hWnd: HWND; E: Exception);
begin
  TFunctions.HandleException(hWnd, E, ExceptAddr, ExceptFrameCount, ExceptFrames);
end;

class function TFunctions.MessageBox(hWnd: HWND; Text: UnicodeString; Caption: UnicodeString; uType: UINT): LongInt;
begin
  Result := MessageBoxW(hWnd, PWideChar(Text), PWideChar(Caption), uType);
end;

class function TFunctions.GetSpecialFolder(const csidl: ShortInt): string;
var
  Buf: UnicodeString;
begin
  SetLength(Buf, 1024);
  if Failed(SHGetFolderPathW(0, csidl, 0, SHGFP_TYPE_CURRENT, PWideChar(Buf))) then
    raise Exception.Create('SHGetFolderPathW() failed');
  Result := PWideChar(Buf);
end;

class function TFunctions.GetTempPath: string;
var
  Buf: UnicodeString;
begin
  SetLength(Buf, MAX_PATH + 1);
  SetLength(Buf, GetTempPathW(Length(Buf), PWideChar(Buf)));
  Result := Buf;
end;

class procedure TFunctions.RunUninstall;
var
  UninstallLibraryPath: string;
  Reg: TRegistry;
begin
  if FileExists(TPaths.WhatTheLockLibraryPath) and (not SysUtils.DeleteFile(TPaths.WhatTheLockLibraryPath)) then
  begin
    UninstallLibraryPath := ConcatPaths([ExtractFileDir(TPaths.WhatTheLockLibraryPath), Guid.NewGuid.ToString(True)]);
    SysUtils.RenameFile(TPaths.WhatTheLockLibraryPath, UninstallLibraryPath);
    MoveFileEx(PChar(UninstallLibraryPath), nil, MOVEFILE_DELAY_UNTIL_REBOOT);
  end;

  if not SysUtils.RemoveDir(TPaths.WhatTheLockDir) then
    MoveFileEx(PChar(TPaths.WhatTheLockDir), nil, MOVEFILE_DELAY_UNTIL_REBOOT);

  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    Reg.DeleteKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%s'.Format([APPNAME]));
  finally
    Reg.Free;
  end;
end;

class function TFunctions.GetFileVersion(const FileName: string): string;
var
  VerInfoSize: Integer;
  VerValueSize: DWord;
  Dummy: DWord;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
begin
  VerInfoSize := GetFileVersionInfoSizeW(PWideChar(UnicodeString(FileName)), Dummy);
  if VerInfoSize <> 0 then
  begin
    GetMem(VerInfo, VerInfoSize);
    try
      if GetFileVersionInfoW(PWideChar(UnicodeString(FileName)), 0, VerInfoSize, VerInfo) then
        if VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize) then
          Exit('%d.%d.%d.%d'.Format([VerValue.dwFileVersionMS shr 16, VerValue.dwFileVersionMS and $FFFF, VerValue.dwFileVersionLS shr 16, VerValue.dwFileVersionLS and $FFFF]));
    finally
      FreeMem(VerInfo, VerInfoSize);
    end;
  end;

  raise Exception.Create('Error reading file version.');
end;

end.
