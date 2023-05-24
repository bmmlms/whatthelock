unit FileHandleUtil;

interface

uses
  Classes,
  Generics.Collections,
  JwaNative,
  JwaPsApi,
  JwaWinType,
  SysUtils,
  Windows;

type

  { TFileHandle }

  TFileHandle = class
  private
    FPID: Cardinal;
    FImageName: string;
    FNtFilename: string;
    FFilename: string;
    FHandle: THandle;
  public
    constructor Create(const PID: Cardinal; const ImageName, NtFilename, Filename: string; const Handle: THandle);

    property PID: Cardinal read FPID;
    property ImageName: string read FImageName;
    property NtFilename: string read FNtFilename;
    property Filename: string read FFilename;
    property Handle: THandle read FHandle;
  end;

  { TFileHandleUtil }

  TFileHandleUtil = class
  private
  class var
    FIsWow64Process2: function(hProcess: HANDLE; pProcessMachine: PUSHORT; pNativeMachine: PUSHORT): BOOL; stdcall;
    FGetFinalPathNameByHandleW: function(hFILE: HANDLE; lpszFilePath: LPWSTR; cchFilePath, dwFlags: DWORD): DWORD; stdcall;
    FEnumProcessModulesEx: function(hProcess: HANDLE; lphModule: PHMODULE; cb: DWORD; var lpcbNeeded: DWORD; dwFilterFlag: DWORD): BOOL; stdcall;
  private
    FOpenHandles: TList<TFileHandle>;

    function IsProcessWOW64(const Process: THandle): Boolean;
    function GetKernel32WOW64(const Process: THandle): HMODULE;
    function GetProcAddressWOW64(const Process, Module: THandle; const FunctionName: string): Pointer;

    procedure RemoteCloseHandle(const Process, Handle: THandle);
    procedure RemoteCloseHandleWOW64(const Process, Handle: THandle);
    procedure RunThread(const Process, Handle: THandle; const Func: Pointer);
  public
    class procedure Initialize; static;

    constructor Create;
    destructor Destroy; override;

    procedure CheckHandles(FilesOrDirectories: array of string; const ExcludeDeletable: Boolean);
    procedure CloseHandles(const OpenHandles: array of TFileHandle; const Kill: Boolean);

    property OpenHandles: TList<TFileHandle> read FOpenHandles;
  end;

implementation

type
  SYSTEM_HANDLE_INFORMATION_ENTRY = JwaNative.SYSTEM_HANDLE_INFORMATION;
  PSYSTEM_HANDLE_INFORMATION_ENTRY = ^SYSTEM_HANDLE_INFORMATION_ENTRY;

  SYSTEM_HANDLE_INFORMATION = record
    Count: ULONG;
    Handles: array[0..0] of SYSTEM_HANDLE_INFORMATION_ENTRY;
  end;
  PSYSTEM_HANDLE_INFORMATION = ^SYSTEM_HANDLE_INFORMATION;

{ TFileHandle }

constructor TFileHandle.Create(const PID: Cardinal; const ImageName, NtFilename, Filename: string; const Handle: THandle);
begin
  FPID := PID;
  FImageName := ImageName;
  FNtFilename := NtFilename;
  FFilename := Filename;
  FHandle := Handle;
end;

function TFileHandleUtil.IsProcessWOW64(const Process: THandle): Boolean;
var
  ProcessMachine, NativeMachine: USHORT;
begin
  FIsWow64Process2(Process, @ProcessMachine, @NativeMachine);
  Result := ProcessMachine <> IMAGE_FILE_MACHINE_UNKNOWN;
end;

function TFileHandleUtil.GetKernel32WOW64(const Process: THandle): HMODULE;
const
  LIST_MODULES_32BIT = $01;
var
  i: Integer;
  Needed: DWORD;
  hMods: array[0..1024] of HMODULE;
  FilePathPtr: PWideChar;
  FilePath: string;
begin
  FilePathPtr := GetMem($1000);

  try
    if FEnumProcessModulesEx(Process, @hMods[0], SizeOf(hMods), Needed, LIST_MODULES_32BIT) then
      for i := 0 to Trunc(Needed / SizeOf(HMODULE)) - 1 do
      begin
        if GetModuleFileNameExW(Process, hMods[i], FilePathPtr, MemSize(FilePathPtr)) = 0 then
          Continue;

        FilePath := FilePathPtr;
        if FilePath.EndsWith(kernel32, True) then
          Exit(hMods[i]);
      end;
  finally
    Freemem(FilePathPtr);
  end;

  raise Exception.Create('Kernel32 not found');
end;

function TFileHandleUtil.GetProcAddressWOW64(const Process, Module: THandle; const FunctionName: string): Pointer;
var
  ModuleBase: Pointer absolute Module;
  DosHeader: IMAGE_DOS_HEADER;
  NtHeaders: IMAGE_NT_HEADERS32;
  ExportDirectory: IMAGE_EXPORT_DIRECTORY;
  ExportDirectoryRva, ExportFunctionNameRva, ExportFunctionRva: Cardinal;
  FunctionTable, NameTable, NameOrdinalTable: Pointer;
  Ordinal: WORD;
  ExportFunctionName: string;
  i: Integer;
begin
  if not ReadProcessMemory(Process, ModuleBase, @DosHeader, sizeof(DosHeader), nil) then
    raise Exception.Create('ReadProcessMemory() failed: %s'.Format([SysErrorMessage(GetLastError)]));

  if DosHeader.e_magic <> IMAGE_DOS_SIGNATURE then
    raise Exception.Create('Invalid module');

  if not ReadProcessMemory(Process, ModuleBase + DosHeader.e_lfanew, @NtHeaders, SizeOf(NtHeaders), nil) then
    raise Exception.Create('ReadProcessMemory() failed: %s'.Format([SysErrorMessage(GetLastError)]));

  if NtHeaders.Signature <> IMAGE_NT_SIGNATURE then
    raise Exception.Create('Invalid module');

  ExportDirectoryRva := NtHeaders.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT].VirtualAddress;
  if ExportDirectoryRva = 0 then
    raise Exception.Create('Invalid module');

  if not ReadProcessMemory(Process, ModuleBase + ExportDirectoryRva, @ExportDirectory, SizeOf(ExportDirectory), nil) then
    raise Exception.Create('ReadProcessMemory() failed: %s'.Format([SysErrorMessage(GetLastError)]));

  FunctionTable := ModuleBase + ExportDirectory.AddressOfFunctions;
  NameTable := ModuleBase + ExportDirectory.AddressOfNames;
  NameOrdinalTable := ModuleBase + ExportDirectory.AddressOfNameOrdinals;

  SetLength(ExportFunctionName, FunctionName.Length);

  for  i := 0 to ExportDirectory.NumberOfNames - 1 do
  begin
    if not ReadProcessMemory(Process, NameTable + (i * SizeOf(ExportFunctionNameRva)), @ExportFunctionNameRva, SizeOf(ExportFunctionNameRva), nil) then
      raise Exception.Create('ReadProcessMemory() failed: %s'.Format([SysErrorMessage(GetLastError)]));

    if not ReadProcessMemory(Process, ModuleBase + ExportFunctionNameRva, @ExportFunctionName[1], ExportFunctionName.Length, nil) then
      raise Exception.Create('ReadProcessMemory() failed: %s'.Format([SysErrorMessage(GetLastError)]));

    if ExportFunctionName = FunctionName then
    begin
      if not ReadProcessMemory(Process, NameOrdinalTable + (i * SizeOf(WORD)), @Ordinal, SizeOf(WORD), nil) then
        raise Exception.Create('ReadProcessMemory() failed: %s'.Format([SysErrorMessage(GetLastError)]));

      if not ReadProcessMemory(Process, FunctionTable + (Ordinal * SizeOf(ExportFunctionRva)), @ExportFunctionRva, SizeOf(ExportFunctionRva), nil) then
        raise Exception.Create('ReadProcessMemory() failed: %s'.Format([SysErrorMessage(GetLastError)]));

      Exit(ModuleBase + ExportFunctionRva);
    end;
  end;

  raise Exception.Create('Address of function %s not found'.Format([FunctionName]));
end;

procedure TFileHandleUtil.RemoteCloseHandle(const Process, Handle: THandle);
const
  CloseHandleIdx = 11;
  FuncBytes: array[0..26] of byte = (
    $55,                                              // push rbp
    $48, $89, $E5,                                    // mov rbp,rsp
    $48, $8D, $64, $24, $D0,                          // lea rsp,qword ptr ss:[rsp-30]
    $48, $BB, $00, $00, $00, $00, $00, $00, $00, $00, // mov rbx,v
    $FF, $D3,                                         // call rbx
    $48, $8D, $65, $00,                               // lea rsp,qword ptr ss:[rbp]
    $5D,                                              // pop rbp
    $C3                                               // ret
    );
var
  Func, CloseHandleAddr: Pointer;
begin
  Func := GetMem(Length(FuncBytes));

  try
    Move(FuncBytes[0], Func^, MemSize(Func));

    CloseHandleAddr := GetProcAddress(GetModuleHandle(kernel32), 'CloseHandle');
    Move(CloseHandleAddr, Pointer(Func + CloseHandleIdx)^, SizeOf(CloseHandleAddr));

    RunThread(Process, Cardinal(Handle), Func);
  finally
    Freemem(Func);
  end;
end;

procedure TFileHandleUtil.RemoteCloseHandleWOW64(const Process, Handle: THandle);
const
  CloseHandleIdx = 7;
  FuncBytes: array[0..18] of byte = (
    $55,                      // push ebp
    $89, $E5,                 // mov ebp,esp
    $FF, $75, $08,            // push dword ptr ss:[ebp+8]
    $B8, $00, $00, $00, $00,  // mov eax,v
    $FF, $D0,                 // call eax
    $89, $EC,                 // mov esp,ebp
    $5D,                      // pop ebp
    $C2, $04, $00             // ret 4
    );
var
  Func: Pointer;
  CloseHandleAddr: Cardinal;
begin
  Func := GetMem(Length(FuncBytes));

  try
    Move(FuncBytes[0], Func^, MemSize(Func));

    CloseHandleAddr := Cardinal(GetProcAddressWOW64(Process, GetKernel32WOW64(Process), 'CloseHandle'));
    Move(CloseHandleAddr, Pointer(Func + CloseHandleIdx)^, SizeOf(CloseHandleAddr));

    RunThread(Process, Cardinal(Handle), Func);
  finally
    Freemem(Func);
  end;
end;

procedure TFileHandleUtil.RunThread(const Process, Handle: THandle; const Func: Pointer);
var
  RemoteFunc: Pointer;
  Thread: THandle;
begin
  RemoteFunc := VirtualAllocEx(Process, nil, MemSize(Func), MEM_COMMIT or MEM_RESERVE, PAGE_EXECUTE_READWRITE);
  if not Assigned(RemoteFunc) then
    raise Exception.Create('VirtualAllocEx() failed: %s'.Format([SysErrorMessage(GetLastError)]));

  try
    if not WriteProcessMemory(Process, RemoteFunc, Func, MemSize(Func), nil) then
      raise Exception.Create('WriteProcessMemory() failed: %s'.Format([SysErrorMessage(GetLastError)]));

    Thread := CreateRemoteThread(Process, nil, 0, RemoteFunc, Pointer(Handle), 0, nil);
    if Thread = 0 then
      raise Exception.Create('CreateRemoteThread() failed: %s'.Format([SysErrorMessage(GetLastError)]));

    WaitForSingleObject(Thread, INFINITE);

    GetExitCodeThread(Thread, @ExitCode);
    CloseHandle(Thread);

    if ExitCode = 0 then
      raise Exception.Create('CloseHandle() failed');
  finally
    VirtualFreeEx(Process, RemoteFunc, 0, MEM_RELEASE);
  end;
end;

class procedure TFileHandleUtil.Initialize;
begin
  FIsWow64Process2 := GetProcAddress(GetModuleHandle('kernelbase.dll'), 'IsWow64Process2');
  FGetFinalPathNameByHandleW := GetProcAddress(GetModuleHandle(kernel32), 'GetFinalPathNameByHandleW');
  FEnumProcessModulesEx := GetProcAddress(GetModuleHandle('psapi.dll'), 'EnumProcessModulesEx');

  if (not Assigned(FIsWow64Process2)) or (not Assigned(FGetFinalPathNameByHandleW)) or (not Assigned(FEnumProcessModulesEx)) then
    raise Exception.Create('A required function could not be found, your windows version is most likely unsupported.');
end;

constructor TFileHandleUtil.Create;
begin
  FOpenHandles := TList<TFileHandle>.Create;
end;

destructor TFileHandleUtil.Destroy;
var
  OpenHandle: TFileHandle;
begin
  for OpenHandle in FOpenHandles do
    OpenHandle.Free;
  FOpenHandles.Clear;

  inherited Destroy;
end;

procedure TFileHandleUtil.CheckHandles(FilesOrDirectories: array of string; const ExcludeDeletable: Boolean);

  function DosPathNameToNtPathName(DosPathName: UnicodeString; Res: PUnicodeString): Boolean;
  var
    NtName: WideString;
    NtNameArg: PUNICODE_STRING;
  begin
    Result := False;
    Result := RtlDosPathNameToNtPathName_U(PWideChar(DosPathName), Res^, nil, nil);
  end;

  function GetImageName(ProcessHandle: THandle): string;
  var
    Size: DWORD;
    Buf: UnicodeString;
  begin
    Size := 1024;
    SetLength(Buf, Size * 2);

    if not QueryFullProcessImageNameW(ProcessHandle, 0, PWideChar(Buf), @Size) then
      raise Exception.Create('QueryFullProcessImageNameW() failed: %s'.Format([SysErrorMessage(GetLastError)]));

    Result := PWideChar(Buf);
  end;

  function IsDeletable(Filename: TUnicodeString): Boolean;
  const
    Delete = $00010000;
    READ_CONTROL = $00020000;
    OBJ_CASE_INSENSITIVE = $00000040;
  var
    FileHandle: HFILE;
    ObjectAttributes: TObjectAttributes;
    IoStatusBlock: IO_STATUS_BLOCK;
    Res: NTSTATUS;
  begin
    InitializeObjectAttributes(@ObjectAttributes, @Filename, OBJ_CASE_INSENSITIVE, 0, nil);

    Res := NtCreateFile(@FileHandle, Delete or FILE_READ_ATTRIBUTES or READ_CONTROL or SYNCHRONIZE, @ObjectAttributes, @IoStatusBlock, nil, 0, FILE_SHARE_DELETE or FILE_SHARE_READ or
      FILE_SHARE_WRITE, FILE_OPEN, FILE_OPEN_FOR_BACKUP_INTENT or FILE_SYNCHRONOUS_IO_NONALERT, nil, 0);

    if NT_ERROR(Res) then
      Exit(False);

    NtClose(FileHandle);
    Exit(True);
  end;

type
  UNIVERSAL_NAME_INFOW = record
    lpUniversalName: LPWSTR;
  end;
  PUNIVERSALNAMEINFOW = ^UNIVERSAL_NAME_INFOW;
const
  STATUS_INFO_LENGTH_MISMATCH: NTSTATUS = $c0000004;
var
  HandleInfo: PSYSTEM_HANDLE_INFORMATION;
  QueryRes: NTSTATUS;
  HandleInfoEntry: PSYSTEM_HANDLE_INFORMATION_ENTRY;
  i: Integer;
  Process, DupHandle: THandle;
  ObjectInfo: POBJECT_TYPE_INFORMATION;
  FilePathPtr: PWideChar;
  FilePath, NtFilePath, CheckFilePath: string;
  FilePathUnicode: UnicodeString;
  NtFilePathUnicode: TUnicodeString;
  Res: Cardinal;
  OpenHandle: TFileHandle;
  UniversalNameBuf: PUNIVERSALNAMEINFOW absolute FilePathPtr;
  BufferSize: DWORD;
begin
  for OpenHandle in FOpenHandles do
    OpenHandle.Free;
  FOpenHandles.Clear;

  HandleInfo := GetMem($10000);
  ObjectInfo := GetMem($1000);
  FilePathPtr := GetMem($1000);

  for i := 0 to Length(FilesOrDirectories) - 1 do
  begin
    FilePathUnicode := FilesOrDirectories[i];

    BufferSize := MemSize(UniversalNameBuf);
    if WNetGetUniversalNameW(PWideChar(FilePathUnicode), UNIVERSAL_NAME_INFO_LEVEL, UniversalNameBuf, BufferSize) = NO_ERROR then
      FilePathUnicode := UniversalNameBuf.lpUniversalName;

    if DosPathNameToNtPathName(FilePathUnicode, @NtFilePathUnicode) then
    begin
      FilesOrDirectories[i] := NtFilePathUnicode.Buffer;
      RtlFreeUnicodeString(@NtFilePathUnicode);
    end;
  end;

  try
    QueryRes := NtQuerySystemInformation(SystemHandleInformation, HandleInfo, MemSize(HandleInfo), nil);
    while QueryRes = STATUS_INFO_LENGTH_MISMATCH do
    begin
      HandleInfo := ReAllocMem(HandleInfo, MemSize(HandleInfo) * 2);
      QueryRes := NtQuerySystemInformation(SystemHandleInformation, HandleInfo, MemSize(HandleInfo), nil);
    end;

    if not NT_SUCCESS(QueryRes) then
      raise Exception.Create('NtQuerySystemInformation() failed: %s'.Format([SysErrorMessage(RtlNtStatusToDosError(QueryRes))]));

    for i := 0 to HandleInfo^.Count - 1 do
    begin
      HandleInfoEntry := @HandleInfo^.Handles[i];

      if HandleInfoEntry.ProcessId = GetCurrentProcessId then
        Continue;

      Process := OpenProcess(PROCESS_DUP_HANDLE or PROCESS_QUERY_LIMITED_INFORMATION, False, HandleInfoEntry^.ProcessId);
      if Process = 0 then
        Continue;

      try
        if not NT_SUCCESS(NtDuplicateObject(Process, HandleInfoEntry^.Handle, GetCurrentProcess, @DupHandle, 0, 0, 0)) then
          Continue;

        try
          if GetFileType(DupHandle) <> FILE_TYPE_DISK then
            Continue;

          if not NT_SUCCESS(NtQueryObject(DupHandle, ObjectTypeInformation, ObjectInfo, MemSize(ObjectInfo), nil)) then
            Continue;

          if ObjectInfo^.Name.Buffer <> 'File' then
            Continue;

          Res := FGetFinalPathNameByHandleW(DupHandle, FilePathPtr, MemSize(FilePathPtr) - 1, 0);
          if (Res <= 4) or (Res > MemSize(FilePathPtr) - 1) then
            Continue;

          if not DosPathNameToNtPathName(FilePathPtr, @NtFilePathUnicode) then
            Continue;

          try
            if ExcludeDeletable and IsDeletable(NtFilePathUnicode) then
              Continue;

            FilePath := FilePathPtr;

            FilePath := FilePath.Replace('\\?\UNC\', '\\');
            FilePath := FilePath.Replace('\\?\', '');

            NtFilePath := NtFilePathUnicode.Buffer;
          finally
            RtlFreeUnicodeString(@NtFilePathUnicode);
          end;

          for CheckFilePath in FilesOrDirectories do
            if (CheckFilePath.ToLower = NtFilePath.ToLower) or (NtFilePath.StartsWith(IncludeTrailingPathDelimiter(CheckFilePath), True)) then
              FOpenHandles.Add(TFileHandle.Create(HandleInfoEntry^.ProcessId, GetImageName(Process), NtFilePath, FilePath, HandleInfoEntry^.Handle));
        finally
          CloseHandle(DupHandle);
        end;
      finally
        CloseHandle(Process);
      end;
    end;
  finally
    Freemem(HandleInfo);
    Freemem(ObjectInfo);
    Freemem(FilePathPtr);
  end;
end;

procedure TFileHandleUtil.CloseHandles(const OpenHandles: array of TFileHandle; const Kill: Boolean);
var
  ExMsg: string;
  Process: THandle;
  OpenHandle: TFileHandle;
begin
  ExMsg := '';

  for OpenHandle in OpenHandles do
    try
      Process := OpenProcess(PROCESS_ALL_ACCESS, False, OpenHandle.PID);
      if Process = 0 then
        raise Exception.Create('OpenProcess() failed: %s'.Format([SysErrorMessage(GetLastError)]));

      try
        if Kill then
        begin
          if not TerminateProcess(Process, 1) then
            if GetLastError <> 5 then
              raise Exception.Create('TerminateProcess() failed: %s'.Format([SysErrorMessage(GetLastError)]));
        end else if IsProcessWOW64(Process) then
          RemoteCloseHandleWOW64(Process, OpenHandle.Handle)
        else
          RemoteCloseHandle(Process, OpenHandle.Handle);
      finally
        CloseHandle(Process);
      end;
    except
      on E: Exception do
        ExMsg += '- %s%s'.Format([E.Message, LineEnding]);
    end;

  if ExMsg <> '' then
    raise Exception.Create(ExMsg.Trim);
end;

end.
