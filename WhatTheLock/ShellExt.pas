unit ShellExt;

interface

uses
  ActiveX,
  ComObj,
  FileHandleUtil,
  Functions,
  Graphics,
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
    FFileName: string;
    FIsDirectory: Boolean;

    class procedure CreateMenuBitmap; static;
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

implementation

uses
  App,
  ComServ,
  Registry,
  SysUtils;

resourcestring
  SHelpText = 'Try to unlock a currently used file/folder';
  SCheckFileFolder = 'Check with What The Lock...';
  SFolderNoHandles = 'The selected folder and its contents are not in use by any process.';
  SFileNoHandles = 'The selected file is not in use by any process.';

{ TWhatTheLockMenu }

class procedure TWhatTheLockMenu.CreateMenuBitmap;
type
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

  function GetBitmapHeader(const Size: TSize): BITMAPV5HEADER;
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
  BitmapSize: TSize;
begin
  BitmapSize := TSize.Create(16, 16);

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
    FFileName := FilePathPtr;
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
  if IncludeTrailingPathDelimiter(ExtractFileDrive(FFileName)) = IncludeTrailingPathDelimiter(FFileName) then
    Exit(MakeResult(SEVERITY_SUCCESS, 0, 0));

  FIsDirectory := DirectoryExists(FFileName);
  Text := SCheckFileFolder;

  MenuItemInfo.cbSize := SizeOf(MenuItemInfo);
  MenuItemInfo.fMask := MIIM_STRING or MIIM_ID or MIIM_BITMAP;
  MenuItemInfo.wID := idCmdFirst;
  MenuItemInfo.hbmpItem := FMenuBitmap;
  MenuItemInfo.dwTypeData := PWideChar(Text);
  MenuItemInfo.cch := Length(Text);

  MenuInfo.cbSize := SizeOf(MenuInfo);
  MenuInfo.fMask := MIM_STYLE;
  MenuInfo.dwStyle := MNS_CHECKORBMP;

  SetMenuInfo(hmenu, @MenuInfo);

  InsertMenuItemW(hmenu, indexMenu, True, @MenuItemInfo);

  Exit(MakeResult(SEVERITY_SUCCESS, 0, 1));
end;

function TWhatTheLockMenu.InvokeCommand(var lpici: shlobj.TCMINVOKECOMMANDINFO): HResult; stdcall;
var
  FileHandleUtil: TFileHandleUtil;
begin
  Result := NOERROR;

  if HiWord(LongInt(lpici.lpVerb)) <> 0 then
    Exit(E_FAIL);

  if LoWord(LongInt(lpici.lpVerb)) > 0 then
    Exit(E_INVALIDARG);

  if LoWord(LongInt(lpici.lpVerb)) <> 0 then
    Exit;

  FileHandleUtil := TFileHandleUtil.Create;

  try
    FileHandleUtil.CheckHandles([FFileName], False);
  except
    on E: Exception do
      TFunctions.HandleException(lpici.hwnd, E);
  end;

  if FileHandleUtil.OpenHandles.Count > 0 then
    TApp.ShowTaskDialog(lpici.hwnd, FileHandleUtil, nil)
  else
  begin
    FileHandleUtil.Free;
    TFunctions.MessageBox(lpici.hwnd, IfThen<string>(FIsDirectory, SFolderNoHandles, SFileNoHandles), SMsgDlgInformation, MB_ICONINFORMATION);
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
      StrUc := SHelpText;
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
  TWhatTheLockMenu.CreateMenuBitmap;
  TWhatTheLockMenuFactory.Create(ComServer, TWhatTheLockMenu, CLSID_WhatTheLockMenu, 'WhatTheLock', 'What The Lock Shell Extension', ciMultiInstance, tmApartment);

end.
