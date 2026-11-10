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
    class function RunElevated(const FileName, Parameters: string; out ProcessExitCode: DWORD): Boolean; static;
    class function BeginPerUserRegistration: HKEY; static;
    class procedure EndPerUserRegistration(const ClassesKey: HKEY); static;

    // Functions only used by Launcher/Setup
    class function IsInstalled(const Scope: TInstallScope): Boolean; static;
    class procedure RunUninstall; static;
    class function GetFileVersion(const FileName: string): string; static;
    class procedure ScheduleDelete(const FileName, DirName: string); static;
  end;

implementation

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

class function TFunctions.RunElevated(const FileName, Parameters: string; out ProcessExitCode: DWORD): Boolean;
var
  Info: TSHELLEXECUTEINFOW;
  Verb, FileNameW, ParametersW: UnicodeString;
  Error: DWORD;
begin
  ProcessExitCode := 0;

  Verb := 'runas';
  FileNameW := UnicodeString(FileName);
  ParametersW := UnicodeString(Parameters);

  ZeroMemory(@Info, SizeOf(Info));
  Info.cbSize := SizeOf(Info);
  Info.fMask := SEE_MASK_NOCLOSEPROCESS;
  Info.lpVerb := PWideChar(Verb);
  Info.lpFile := PWideChar(FileNameW);
  Info.lpParameters := PWideChar(ParametersW);
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

class procedure TFunctions.ScheduleDelete(const FileName, DirName: string);
var
  Command: string;
  Reg: TRegistry;
begin
  if TPaths.Scope = scAllUsers then
  begin
    if FileName <> '' then
      MoveFileEx(PChar(FileName), nil, MOVEFILE_DELAY_UNTIL_REBOOT);
    if DirName <> '' then
      MoveFileEx(PChar(DirName), nil, MOVEFILE_DELAY_UNTIL_REBOOT);
    Exit;
  end;

  // MOVEFILE_DELAY_UNTIL_REBOOT requires administrator rights, delete on next logon of the current user instead
  Command := '';
  if FileName <> '' then
    Command := 'del /f /q "%s"'.Format([FileName]);
  if DirName <> '' then
  begin
    if Command <> '' then
      Command := Command + ' & ';
    Command := Command + 'rmdir "%s"'.Format([DirName]);
  end;

  if Command = '' then
    Exit;

  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce', True) then
      Reg.WriteString('WhatTheLock_%s'.Format([Guid.NewGuid.ToString(True)]), '"%s" /c %s'.Format([ConcatPaths([GetSpecialFolder(CSIDL_SYSTEM), 'cmd.exe']), Command]));
  finally
    Reg.Free;
  end;
end;

class procedure TFunctions.RunUninstall;
var
  UninstallLibraryPath: string;
  Reg: TRegistry;
begin
  UninstallLibraryPath := '';

  if FileExists(TPaths.WhatTheLockLibraryPath) and (not SysUtils.DeleteFile(TPaths.WhatTheLockLibraryPath)) then
  begin
    UninstallLibraryPath := ConcatPaths([ExtractFileDir(TPaths.WhatTheLockLibraryPath), Guid.NewGuid.ToString(True)]);
    if not SysUtils.RenameFile(TPaths.WhatTheLockLibraryPath, UninstallLibraryPath) then
      UninstallLibraryPath := TPaths.WhatTheLockLibraryPath;
  end;

  if not SysUtils.RemoveDir(TPaths.WhatTheLockDir) then
    ScheduleDelete(UninstallLibraryPath, TPaths.WhatTheLockDir);

  Reg := TRegistry.Create;
  try
    Reg.RootKey := TPaths.GetRegistryRootKey(TPaths.Scope);
    Reg.DeleteKey(UNINSTALL_KEY);
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
