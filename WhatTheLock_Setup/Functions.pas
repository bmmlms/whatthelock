unit Functions;

interface

uses
  ActiveX,
  Classes,
  Constants,
  Paths,
  Registry,
  ShlObj,
  StrUtils,
  SysUtils,
  Windows;

type
  // TODO: hier sehr viel aufräumen...
  // TODO: constants in dll selber auch nutzen.

  TIsWow64Process2 = function(hProcess: THandle; pProcessMachine: PUSHORT; pNativeMachine: PUSHORT): BOOL; stdcall;

  { TFunctions }

  TFunctions = class
  private
    class var
    FIsWow64Process2: TIsWow64Process2;
  public
    class procedure Init; static;

    // Wrappers for windows functions
    class function MessageBox(hWnd: HWND; Text: UnicodeString; Caption: UnicodeString; uType: UINT): LongInt; static;

    class function GetSpecialFolder(const csidl: ShortInt): string; static;    // TODO: kann weg
    class function GetTempPath: string; static;
    class function IsWindows64Bit: Boolean; static;
    class function FindCmdLineSwitch(const Name: string; var Value: string): Boolean; static; overload;
    class function FindCmdLineSwitch(const Name: string): Boolean; static; overload;

    // Functions only used by Launcher/Setup
    class procedure RunUninstall(const Quiet: Boolean); static;
    class function GetFileVersion(const FileName: string): string; static;
  end;

implementation

class procedure TFunctions.Init;
begin
  FIsWow64Process2 := GetProcAddress(GetModuleHandle('kernelbase.dll'), 'IsWow64Process2');

  if not Assigned(FIsWow64Process2) then
    raise Exception.Create('A required function could not be found, your windows version is most likely unsupported.');
end;

class function TFunctions.MessageBox(hWnd: HWND; Text: UnicodeString; Caption: UnicodeString; uType: UINT): LongInt;       // TODO: nutzen.
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

class function TFunctions.GetTempPath: string;  // TODO: relevant?
var
  Buf: UnicodeString;
begin
  SetLength(Buf, MAX_PATH + 1);
  SetLength(Buf, GetTempPathW(Length(Buf), PWideChar(Buf)));
  Result := Buf;
end;

class function TFunctions.IsWindows64Bit: Boolean;
var
  ProcessMachine, NativeMachine: USHORT;
begin
  FIsWow64Process2(GetCurrentProcess, @ProcessMachine, @NativeMachine);
  Result := NativeMachine = IMAGE_FILE_MACHINE_AMD64;
end;

class function TFunctions.FindCmdLineSwitch(const Name: string; var Value: string): Boolean;   // TODO: die beiden findcmdline funktionen hier relevant?
var
  i: Integer;
begin
  Result := False;
  Value := '';
  for i := 1 to ParamCount do
    if ParamStr(i).Equals('-' + Name) then
    begin
      Value := ParamStr(i + 1);
      Exit(True);
    end;
end;

class function TFunctions.FindCmdLineSwitch(const Name: string): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 1 to ParamCount do
    if ParamStr(i).Equals('-' + Name) then
      Exit(True);
end;

class procedure TFunctions.RunUninstall(const Quiet: Boolean);   // TODO: das hier in die dll direkt rein tun.

  procedure UninstallFile(const FileName: string);
  begin
    if not SysUtils.DeleteFile(FileName) then
      RemoveDir(FileName);
  end;

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

  MoveFileEx(PChar(TPaths.WhatTheLockDir), nil, MOVEFILE_DELAY_UNTIL_REBOOT);

  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    Reg.DeleteKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WhatTheLock');
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
