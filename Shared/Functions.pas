unit Functions;

interface

uses
  Classes,
  Constants,
  JwaWinReg,
  Paths,
  Registry,
  RtlConsts,
  ShellApi,
  ShlObj,
  SysUtils,
  UnicodePaths,
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
    class function GetSpecialFolder(const csidl: ShortInt): UnicodeString; static;
    class function GetTempPath: UnicodeString; static;
    class function RunElevated(const FileName, Parameters: UnicodeString; out ProcessExitCode: DWORD): Boolean; static;
    class function BeginPerUserRegistration: HKEY; static;
    class procedure EndPerUserRegistration(const ClassesKey: HKEY); static;

    // Functions only used by Launcher/Setup
    class function IsInstalled(const Scope: TInstallScope): Boolean; static;
    class procedure RunUninstall; static;
    class function GetFileVersion(const FileName: UnicodeString): UnicodeString; static;
    class procedure ScheduleDelete(const FileName, DirName: UnicodeString); static;
  end;

implementation

uses
  LazFileUtils,
  LazUTF8;

const
  SEE_MASK_NOCLOSEPROCESS = $00000040;
  ERROR_CANCELLED_BY_USER = 1223;

procedure OaEnablePerUserTLibRegistration; stdcall; external 'oleaut32.dll' name 'OaEnablePerUserTLibRegistration';

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

    TFunctions.MessageBox(0, UnicodeString('An unexpected exception occurred.'#13#10) + UnicodeString(SL.Text), UnicodeString(SMsgDlgError), MB_ICONERROR);
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

class function TFunctions.GetSpecialFolder(const csidl: ShortInt): UnicodeString;
var
  Buf: UnicodeString;
begin
  SetLength(Buf, MAX_PATH);
  if Failed(SHGetFolderPathW(0, csidl, 0, SHGFP_TYPE_CURRENT, PWideChar(Buf))) then
    raise Exception.Create('SHGetFolderPathW() failed');
  Result := UnicodeString(PWideChar(Buf));
end;

class function TFunctions.GetTempPath: UnicodeString;
var
  Buf: UnicodeString;
  BufferLength, CharCount: DWORD;
begin
  BufferLength := MAX_PATH + 1;
  repeat
    SetLength(Buf, BufferLength);
    CharCount := GetTempPathW(BufferLength, PWideChar(Buf));
    if CharCount = 0 then
      RaiseLastOSError;
    if CharCount < BufferLength then
      Break;
    BufferLength := CharCount + 1;
  until BufferLength > 32768;
  SetLength(Buf, CharCount);
  Result := Buf;
end;

class function TFunctions.RunElevated(const FileName, Parameters: UnicodeString; out ProcessExitCode: DWORD): Boolean;
var
  Info: TSHELLEXECUTEINFOW;
  Verb: UnicodeString;
  Error: DWORD;
begin
  ProcessExitCode := 0;

  Verb := 'runas';

  ZeroMemory(@Info, SizeOf(Info));
  Info.cbSize := SizeOf(Info);
  Info.fMask := SEE_MASK_NOCLOSEPROCESS;
  Info.lpVerb := PWideChar(Verb);
  Info.lpFile := PWideChar(FileName);
  Info.lpParameters := PWideChar(Parameters);
  Info.nShow := SW_SHOWNORMAL;

  if not ShellExecuteExW(@Info) then
  begin
    Error := GetLastError;
    if Error = ERROR_CANCELLED_BY_USER then
      Exit(False);
    raise Exception.Create('ShellExecuteExW() failed: %s'.Format([SysErrorMessage(Error)]));
  end;

  if Info.hProcess <> 0 then
  try
    WaitForSingleObject(Info.hProcess, INFINITE);
    if not GetExitCodeProcess(Info.hProcess, ProcessExitCode) then
      ProcessExitCode := 1;
  finally
    CloseHandle(Info.hProcess);
  end;

  Result := True;
end;

class function TFunctions.BeginPerUserRegistration: HKEY;
var
  Res: LongInt;
begin
  Res := RegOpenKeyExW(HKEY_CURRENT_USER, 'Software\Classes', 0, KEY_ALL_ACCESS, Result);
  if Res <> ERROR_SUCCESS then
    raise Exception.Create('RegOpenKeyExW() failed: %s'.Format([SysErrorMessage(Res)]));

  OaEnablePerUserTLibRegistration;

  Res := RegOverridePredefKey(HKEY_CLASSES_ROOT, Result);
  if Res <> ERROR_SUCCESS then
  begin
    RegCloseKey(Result);
    raise Exception.Create('RegOverridePredefKey() failed: %s'.Format([SysErrorMessage(Res)]));
  end;
end;

class procedure TFunctions.EndPerUserRegistration(const ClassesKey: HKEY);
begin
  RegOverridePredefKey(HKEY_CLASSES_ROOT, 0);
  RegCloseKey(ClassesKey);
end;

class function TFunctions.IsInstalled(const Scope: TInstallScope): Boolean;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := TPaths.GetRegistryRootKey(Scope);
    Result := Reg.KeyExists(UNINSTALL_KEY);
  finally
    Reg.Free;
  end;
end;

class procedure TFunctions.ScheduleDelete(const FileName, DirName: UnicodeString);
var
  Command, ValueName, Value: UnicodeString;
  Key: HKEY;
  Disposition: DWORD;
  RegistryResult: LongInt;
begin
  if TPaths.Scope = scAllUsers then
  begin
    if FileName <> '' then
      MoveFileExW(PWideChar(FileName), nil, MOVEFILE_DELAY_UNTIL_REBOOT);
    if DirName <> '' then
      MoveFileExW(PWideChar(DirName), nil, MOVEFILE_DELAY_UNTIL_REBOOT);
    Exit;
  end;

  // MOVEFILE_DELAY_UNTIL_REBOOT requires administrator rights, delete on next logon of the current user instead
  Command := '';
  if FileName <> '' then
    Command := 'del /f /q "' + FileName + '"';
  if DirName <> '' then
  begin
    if Command <> '' then
      Command := Command + ' & ';
    Command := Command + 'rmdir "' + DirName + '"';
  end;

  if Command = '' then
    Exit;

  if RegCreateKeyExW(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce', 0, nil, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, nil, Key, @Disposition) <> ERROR_SUCCESS then
    Exit;
  try
    ValueName := 'WhatTheLock_' + UnicodeString(Guid.NewGuid.ToString(True));
    Value := '"' + UnicodeConcatPath(GetSpecialFolder(CSIDL_SYSTEM), 'cmd.exe') + '" /c ' + Command;
    RegistryResult := RegSetValueExW(Key, PWideChar(ValueName), 0, REG_SZ, PByte(PWideChar(Value)), (Length(Value) + 1) * SizeOf(WideChar));
    if RegistryResult <> ERROR_SUCCESS then
    begin
      SetLastError(RegistryResult);
      RaiseLastOSError;
    end;
  finally
    RegCloseKey(Key);
  end;
end;

class procedure TFunctions.RunUninstall;
var
  UninstallLibraryPath: UnicodeString;
  Reg: TRegistry;
begin
  UninstallLibraryPath := '';

  if FileExistsUTF8(UTF16ToUTF8(TPaths.WhatTheLockLibraryPath)) and (not DeleteFileUTF8(UTF16ToUTF8(TPaths.WhatTheLockLibraryPath))) then
  begin
    UninstallLibraryPath := UnicodeConcatPath(SysUtils.ExtractFileDir(TPaths.WhatTheLockLibraryPath), UnicodeString(Guid.NewGuid.ToString(True)));
    if not RenameFileUTF8(UTF16ToUTF8(TPaths.WhatTheLockLibraryPath), UTF16ToUTF8(UninstallLibraryPath)) then
      UninstallLibraryPath := TPaths.WhatTheLockLibraryPath;
  end;

  if not RemoveDirUTF8(UTF16ToUTF8(TPaths.WhatTheLockDir)) then
    ScheduleDelete(UninstallLibraryPath, TPaths.WhatTheLockDir);

  Reg := TRegistry.Create;
  try
    Reg.RootKey := TPaths.GetRegistryRootKey(TPaths.Scope);
    Reg.DeleteKey(UNINSTALL_KEY);
  finally
    Reg.Free;
  end;
end;

class function TFunctions.GetFileVersion(const FileName: UnicodeString): UnicodeString;
var
  VerInfoSize: Integer;
  VerValueSize: DWord;
  Dummy: DWord;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
begin
  VerInfoSize := GetFileVersionInfoSizeW(PWideChar(FileName), Dummy);
  if VerInfoSize <> 0 then
  begin
    GetMem(VerInfo, VerInfoSize);
    try
      if GetFileVersionInfoW(PWideChar(FileName), 0, VerInfoSize, VerInfo) then
        if VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize) then
          Exit(UnicodeString('%d.%d.%d.%d'.Format([VerValue.dwFileVersionMS shr 16, VerValue.dwFileVersionMS and $FFFF, VerValue.dwFileVersionLS shr 16, VerValue.dwFileVersionLS and $FFFF])));
    finally
      FreeMem(VerInfo, VerInfoSize);
    end;
  end;

  raise Exception.Create('Error reading file version.');
end;

end.
