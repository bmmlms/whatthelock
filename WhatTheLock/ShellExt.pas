unit ShellExt;

interface

uses
  ActiveX,
  ComObj,
  FileHandleUtil,
  Functions,
  RtlConsts,
  shlobj,
  windows;

type

  { TWhatTheLockMenu }

  TWhatTheLockMenu = class(TComObject, IUnknown, IContextMenu, IShellExtInit)
  private
  class var
    FMenuBitmap: HBITMAP;
  private
    FFileName: UnicodeString;
  public
    class procedure CreateMenuBitmap; static;

    procedure Initialize; override;
  protected
    // IContextMenu
    function QueryContextMenu(hmenu: HMENU; indexMenu, idCmdFirst, idCmdLast, uFlags: UINT): HResult; stdcall;
    function InvokeCommand(var lpici: shlobj.TCMINVOKECOMMANDINFO): HResult; stdcall;
    function GetCommandString(idCmd: UINT_PTR; uType: UINT; pReserved: PUINT; pszName: LPSTR; cchMax: UINT): HResult; stdcall;

    // IShellExtInit
    function IShellExtInit.Initialize = InitShellExt;
    function InitShellExt(pidlFolder: PItemIDList; pdtobj: IDataObject; hKeyProgID: HKEY): HResult; stdcall;
  end;

  TWhatTheLockMenuFactory = class(TComObjectFactory)
  public
    procedure UpdateRegistry(Register: Boolean); override;
  end;

const
  CLSID_WhatTheLockMenu: TGUID = '{9684203C-406A-4228-8413-1D10E57D4ABD}';

procedure RunCheck(const CommandLine: UnicodeString);

resourcestring
  SHelpText = 'Try to unlock a currently used file/folder';
  SCheckFileFolder = 'Check with What The Lock...';
  SFolderNoHandles = 'The selected folder and its contents are not in use by any process.';
  SFileNoHandles = 'The selected file is not in use by any process.';

implementation

uses
  LazFileUtils,
  LazUTF8,
  App,
  ComServ,
  JwawinUser,
  LibraryInit,
  Paths,
  Registry,
  SysUtils,
  UnicodePaths;

function ParseCheckCommandLine(const CommandLine: UnicodeString; out ExcludePID: Cardinal; out ParentWindow: HWND; out FileName: UnicodeString): Boolean;

  function NextToken(var S: UnicodeString): UnicodeString;
  var
    P: Integer;
  begin
    S := TrimLeft(S);
    P := Pos(UnicodeString(' '), S);
    if P = 0 then
    begin
      Result := S;
      S := '';
    end else
    begin
      Result := Copy(S, 1, P - 1);
      Delete(S, 1, P);
    end;
  end;

var
  Rest: UnicodeString;
  Value: QWord;
  Code: Integer;
begin
  Result := False;
  ExcludePID := 0;
  ParentWindow := 0;
  FileName := '';

  Rest := CommandLine;

  Val(NextToken(Rest), Value, Code);
  if Code <> 0 then
    Exit;
  ExcludePID := Cardinal(Value);

  Val(NextToken(Rest), Value, Code);
  if Code <> 0 then
    Exit;
  ParentWindow := HWND(Value);

  // The file name is the remainder, it may contain spaces
  FileName := Trim(Rest);
  if (Length(FileName) >= 2) and (FileName[1] = '"') and (FileName[Length(FileName)] = '"') then
    FileName := Copy(FileName, 2, Length(FileName) - 2);

  Result := FileName <> '';
end;

procedure StartCheckProcess(const ParentWindow: HWND; const FileName: UnicodeString);
var
  ApplicationName, CommandLine: UnicodeString;
  StartupInfo: STARTUPINFOW;
  ProcessInfo: PROCESS_INFORMATION;
begin
  ApplicationName := UnicodeConcatPath(TFunctions.GetSpecialFolder(CSIDL_SYSTEM), 'rundll32.exe');
  CommandLine := UnicodeString('"') + ApplicationName + UnicodeString('" "') + TPaths.ExePath + UnicodeString('",Check ') +
    UnicodeString(IntToStr(GetCurrentProcessId)) + UnicodeString(' ') + UnicodeString(IntToStr(PtrUInt(ParentWindow))) + UnicodeString(' "') + FileName + UnicodeString('"');
  UniqueString(CommandLine);

  ZeroMemory(@StartupInfo, SizeOf(StartupInfo));
  StartupInfo.cb := SizeOf(StartupInfo);
  StartupInfo.dwFlags := STARTF_FORCEONFEEDBACK;

  ZeroMemory(@ProcessInfo, SizeOf(ProcessInfo));

  if not CreateProcessW(PWideChar(ApplicationName), PWideChar(CommandLine), nil, nil, False, CREATE_SUSPENDED, nil, nil, @StartupInfo, @ProcessInfo) then
    raise Exception.Create('CreateProcessW() failed: %s'.Format([SysErrorMessage(GetLastError)]));

  // Explorer is the foreground process, allow the check process to bring its dialog to the front
  AllowSetForegroundWindow(ProcessInfo.dwProcessId);

  ResumeThread(ProcessInfo.hThread);
  CloseHandle(ProcessInfo.hThread);
  CloseHandle(ProcessInfo.hProcess);
end;

procedure RunCheck(const CommandLine: UnicodeString);
var
  ExcludePID: Cardinal;
  ParentWindow: HWND;
  FileName: UnicodeString;
  IsDirectory: Boolean;
  FileHandleUtil: TFileHandleUtil;
begin
  if not ParseCheckCommandLine(CommandLine, ExcludePID, ParentWindow, FileName) then
    raise Exception.Create('Invalid command line: ' + string(CommandLine));

  if (ParentWindow <> 0) and (not IsWindow(ParentWindow)) then
    ParentWindow := 0;

  IsDirectory := DirectoryExistsUTF8(UTF16ToUTF8(FileName));

  FileHandleUtil := TFileHandleUtil.Create;
  try
    // Explorer's own handles were never reported when the check ran inside explorer, keep it that way
    FileHandleUtil.CheckHandles([FileName], False, ExcludePID);
  except
    FileHandleUtil.Free;
    raise;
  end;

  if FileHandleUtil.OpenHandles.Count > 0 then
    TApp.ShowTaskDialog(ParentWindow, FileHandleUtil, nil)
  else
  begin
    FileHandleUtil.Free;
    // No owner window: An owner in another process (Explorer) would attach both input queues
    TFunctions.MessageBox(0, IfThen<UnicodeString>(IsDirectory, UnicodeString(SFolderNoHandles), UnicodeString(SFileNoHandles)), UnicodeString(SMsgDlgInformation),
      MB_ICONINFORMATION or MB_SETFOREGROUND);
  end;
end;

  { TWhatTheLockMenu }

procedure TWhatTheLockMenu.Initialize;
begin
  inherited Initialize;

  EnsureLibraryInitialized(True);
end;

class procedure TWhatTheLockMenu.CreateMenuBitmap;
type
  TMenuBitmapSize = record
    Width: Longint;
    Height: Longint;
  end;

  BITMAPV5HEADER = record
    bV5Size: DWORD;
    bV5Width: Longint;
    bV5Height: Longint;
    bV5Planes: Word;
    bV5BitCount: Word;
    bV5Compression: DWORD;
    bV5SizeImage: DWORD;
    bV5XPelsPerMeter: Longint;
    bV5YPelsPerMeter: Longint;
    bV5ClrUsed: DWORD;
    bV5ClrImportant: DWORD;
    bV5RedMask: DWORD;
    bV5GreenMask: DWORD;
    bV5BlueMask: DWORD;
    bV5AlphaMask: DWORD;
    bV5CSType: DWORD;
    bV5Endpoints: TCIEXYZTriple;
    bV5GammaRed: DWORD;
    bV5GammaGreen: DWORD;
    bV5GammaBlue: DWORD;
    bV5Intent: DWORD;
    bV5ProfileData: DWORD;
    bV5ProfileSize: DWORD;
    bV5Reserved: DWORD;
  end;

  TCreateDIBSection = function(_para1: HDC; const _para2: BITMAPV5HEADER; _para3: UINT; var _para4: Pointer; _para5: HANDLE; _para6: DWORD): HBITMAP; stdcall;

  function GetBitmapHeader(const Size: TMenuBitmapSize): BITMAPV5HEADER;
  begin
    ZeroMemory(@Result, SizeOf(Result));
    Result.bV5Size := SizeOf(Result);
    Result.bV5Width := Size.Width;
    Result.bV5Height := -Size.Height;
    Result.bV5Planes := 1;
    Result.bV5BitCount := 32;
    Result.bV5Compression := BI_RGB;
  end;

var
  DC: HDC;
  Bitmap, BitmapOld: HBITMAP;
  Icon: HICON;
  BitmapStart: PRGBQUAD;
  BitmapHeader: BITMAPV5HEADER;
  BitmapSize: TMenuBitmapSize;
begin
  BitmapSize.Width := 16;
  BitmapSize.Height := 16;

  DC := CreateCompatibleDC(0);

  BitmapHeader := GetBitmapHeader(BitmapSize);
  Bitmap := TCreateDIBSection(@CreateDIBSection)(DC, BitmapHeader, DIB_RGB_COLORS, BitmapStart, 0, 0);

  BitmapOld := SelectObject(DC, Bitmap);

  Icon := LoadImage(HINSTANCE, 'WHATTHELOCK', IMAGE_ICON, 16, 16, 0);
  DrawIconEx(DC, 0, 0, Icon, BitmapSize.Width, BitmapSize.Height, 0, 0, DI_NORMAL);

  FMenuBitmap := SelectObject(DC, BitmapOld);

  DestroyIcon(Icon);
  DeleteDC(DC);
end;

function TWhatTheLockMenu.InitShellExt(pidlFolder: PItemIDList; pdtobj: IDataObject; hKeyProgID: HKEY): HResult; stdcall;
var
  Medium: TStgMedium;
  Format: TFormatEtc;
  FilePathPtr: PWideChar;
begin
  Result := E_FAIL;

  if not Assigned(pdtobj) then
    Exit;

  Format.cfFormat := CF_HDROP;
  Format.ptd := nil;
  Format.dwAspect := DVASPECT_CONTENT;
  Format.lindex := -1;
  Format.tymed := TYMED_HGLOBAL;

  Result := pdtobj.GetData(Format, Medium);
  if Failed(Result) then
    Exit;

  FilePathPtr := GetMem(4096);
  try
    if DragQueryFileW(Medium.hGlobal, $FFFFFFFF, nil, 0) <> 1 then
      Exit;

    DragQueryFileW(Medium.hGlobal, 0, FilePathPtr, MemSize(FilePathPtr) div 2);
    FFileName := UnicodeString(FilePathPtr);
    Result := NOERROR;
  finally
    Freemem(FilePathPtr);
    ReleaseStgMedium(Medium);
  end;
end;

function TWhatTheLockMenu.QueryContextMenu(hmenu: HMENU; indexMenu, idCmdFirst, idCmdLast, uFlags: UINT): HResult; stdcall;
var
  MenuItemInfo: TMENUITEMINFOW;
  MenuInfo: TMENUINFO;
  Text: UnicodeString;
begin
  if UnicodeCompareTextOrdinal(SysUtils.IncludeTrailingPathDelimiter(SysUtils.ExtractFileDrive(FFileName)), SysUtils.IncludeTrailingPathDelimiter(FFileName)) then
    Exit(MakeResult(SEVERITY_SUCCESS, 0, 0));

  Text := UnicodeString(SCheckFileFolder);

  MenuItemInfo.cbSize := SizeOf(MenuItemInfo);
  MenuItemInfo.fMask := MIIM_STRING or MIIM_ID or MIIM_BITMAP;
  MenuItemInfo.wID := idCmdFirst;
  MenuItemInfo.hbmpItem := FMenuBitmap;
  MenuItemInfo.dwTypeData := PWideChar(Text);
  MenuItemInfo.cch := Length(Text);

  MenuInfo.cbSize := SizeOf(MenuInfo);
  MenuInfo.fMask := MIM_STYLE;
  MenuInfo.dwStyle := MNS_CHECKORBMP;

  SetMenuInfo(hmenu, MenuInfo);

  InsertMenuItemW(hmenu, indexMenu, True, MenuItemInfo);

  Exit(MakeResult(SEVERITY_SUCCESS, 0, 1));
end;

function TWhatTheLockMenu.InvokeCommand(var lpici: shlobj.TCMINVOKECOMMANDINFO): HResult; stdcall;
var
  Verb: PtrUInt;
begin
  Verb := PtrUInt(lpici.lpVerb);

  // IS_INTRESOURCE(): Verbs passed as string are not supported
  if (Verb shr 16) <> 0 then
    Exit(E_FAIL);

  if Word(Verb) <> 0 then
    Exit(E_INVALIDARG);

  try
    // Run the check in a separate process since querying handles of other processes can block
    StartCheckProcess(lpici.hwnd, FFileName);
    Result := NOERROR;
  except
    on E: Exception do
    begin
      TFunctions.HandleException(lpici.hwnd, E);
      Result := E_FAIL;
    end;
  end;
end;

function TWhatTheLockMenu.GetCommandString(idCmd: UINT_PTR; uType: UINT; pReserved: PUINT; pszName: LPSTR; cchMax: UINT): HResult; stdcall;
var
  StrUc: UnicodeString;
begin
  Result := NOERROR;

  if idCmd <> 0 then
    Exit(E_INVALIDARG);

  case uType of
    GCS_HELPTEXTA:
      StrLCopy(pszName, PChar(SHelpText), cchMax);
    GCS_HELPTEXTW:
    begin
      StrUc := UnicodeString(SHelpText);
      StrLCopy(PWideChar(pszName), PWideChar(StrUc), cchMax);
    end;
    else
      Result := E_INVALIDARG;
  end;
end;

{ TWhatTheLockMenuFactory }

procedure TWhatTheLockMenuFactory.UpdateRegistry(Register: Boolean);
var
  Reg: TRegistry;
begin
  inherited UpdateRegistry(Register);

  Reg := TRegistry.Create;
  Reg.RootKey := HKEY_CLASSES_ROOT;
  try
    if Register then
    begin
      if Reg.OpenKey('\*\ShellEx\ContextMenuHandlers\WhatTheLock', True) then
      begin
        Reg.WriteString('', GUIDToString(CLSID_WhatTheLockMenu));
        Reg.CloseKey;
      end;

      if Reg.OpenKey('\Folder\ShellEx\ContextMenuHandlers\WhatTheLock', True) then
      begin
        Reg.WriteString('', GUIDToString(CLSID_WhatTheLockMenu));
        Reg.CloseKey;
      end;
    end else
    begin
      Reg.DeleteKey('\*\ShellEx\ContextMenuHandlers\WhatTheLock');
      Reg.DeleteKey('\Folder\ShellEx\ContextMenuHandlers\WhatTheLock');
    end;
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
end;

initialization
  TWhatTheLockMenuFactory.Create(ComServer, TWhatTheLockMenu, CLSID_WhatTheLockMenu, 'WhatTheLock', 'What The Lock Shell Extension', ciMultiInstance, tmApartment);

end.
