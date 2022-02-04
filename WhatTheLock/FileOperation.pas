unit FileOperation;

interface

uses
  ActiveX,
  Classes,
  ComObj,
  DDetours,
  Dialogs,
  FileHandleUtil,
  FileOperationProgressSink,
  Forms,
  Functions,
  RtlConsts,
  shlobj,
  SysUtils,
  windows;

const
  CLSID_FileOperation: TGUID = '{3ad05575-8857-4850-9277-11b85bdb8e09}';
  SID_IFileOperation = '{947AAB5F-0A5C-4C13-B4D6-4BF7836FC9F8}';

type

  { IFileOperation }

  IFileOperation = interface(IUnknown)
    [SID_IFileOperation]
    function Advise(const pfops: IFileOperationProgressSink; var pdwCookie: DWORD): HRESULT; stdcall;
    function Unadvise(dwCookie: DWORD): HRESULT; stdcall;
    function SetOperationFlags(dwOperationFlags: DWORD): HRESULT; stdcall;
    function SetProgressMessage(pszMessage: LPCWSTR): HRESULT; stdcall;
    function SetProgressDialog(const popd: Pointer): HRESULT; stdcall;
    function SetProperties(const pproparray: Pointer): HRESULT; stdcall;
    function SetOwnerWindow(hwndParent: HWND): HRESULT; stdcall;
    function ApplyPropertiesToItem(const psiItem: IShellItem): HRESULT; stdcall;
    function ApplyPropertiesToItems(const punkItems: IUnknown): HRESULT; stdcall;
    function RenameItem(const psiItem: IShellItem; pszNewName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function RenameItems(const pUnkItems: IUnknown; pszNewName: LPCWSTR): HRESULT; stdcall;
    function MoveItem(const psiItem: IShellItem; const psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function MoveItems(const punkItems: IUnknown; const psiDestinationFolder: IShellItem): HRESULT; stdcall;
    function CopyItem(const psiItem: IShellItem; const psiDestinationFolder: IShellItem; pszCopyName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function CopyItems(const punkItems: IUnknown; const psiDestinationFolder: IShellItem): HRESULT; stdcall;
    function DeleteItem(const psiItem: IShellItem; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function DeleteItems(const punkItems: IUnknown): HRESULT; stdcall;
    function NewItem(const psiDestinationFolder: IShellItem; dwFileAttributes: DWORD; pszName: LPCWSTR; pszTemplateName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function PerformOperations: HRESULT; stdcall;
    function GetAnyOperationsAborted(var pfAnyOperationsAborted: BOOL): HRESULT; stdcall;
  end;

  { TFileOperation }

  TFileOperation = class(TInterfacedObject, IFileOperation)
    type
      TProcessDataObjectArg = record
        Instance: TFileOperation;
        DataObject: IDataObject;
      end;
      PProcessDataObjectArg = ^TProcessDataObjectArg;

  private
    FTaskDialogWindowHandle, FParentWindowHandle: THandle;
    FWrappedFileOperation: IFileOperation;
    FDeleteFiles: TStringList;
    FFileHandleUtil: TFileHandleUtil;
    FProgressSink: TFileOperationProgressSink;
    FThreadHandle: THandle;

    class procedure ProcessDataObject(const Args: PProcessDataObjectArg); static;
    class procedure CheckHandlesThreadProcWrapper(const FileOperation: TFileOperation); stdcall; static;

    procedure CheckHandlesThreadProc;
    procedure FileOperationProgressSinkPreDeleteItem(Sender: TObject);
  public
    constructor Create(const WrappedFileOperation: IFileOperation);
    destructor Destroy; override;

    function Advise(const pfops: IFileOperationProgressSink; var pdwCookie: DWORD): HRESULT; stdcall;
    function Unadvise(dwCookie: DWORD): HRESULT; stdcall;
    function SetOperationFlags(dwOperationFlags: DWORD): HRESULT; stdcall;
    function SetProgressMessage(pszMessage: LPCWSTR): HRESULT; stdcall;
    function SetProgressDialog(const popd: Pointer): HRESULT; stdcall;
    function SetProperties(const pproparray: Pointer): HRESULT; stdcall;
    function SetOwnerWindow(hwndParent: HWND): HRESULT; stdcall;
    function ApplyPropertiesToItem(const psiItem: IShellItem): HRESULT; stdcall;
    function ApplyPropertiesToItems(const punkItems: IUnknown): HRESULT; stdcall;
    function RenameItem(const psiItem: IShellItem; pszNewName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function RenameItems(const pUnkItems: IUnknown; pszNewName: LPCWSTR): HRESULT; stdcall;
    function MoveItem(const psiItem: IShellItem; const psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function MoveItems(const punkItems: IUnknown; const psiDestinationFolder: IShellItem): HRESULT; stdcall;
    function CopyItem(const psiItem: IShellItem; const psiDestinationFolder: IShellItem; pszCopyName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function CopyItems(const punkItems: IUnknown; const psiDestinationFolder: IShellItem): HRESULT; stdcall;
    function DeleteItem(const psiItem: IShellItem; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function DeleteItems(const punkItems: IUnknown): HRESULT; stdcall;
    function NewItem(const psiDestinationFolder: IShellItem; dwFileAttributes: DWORD; pszName: LPCWSTR; pszTemplateName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
    function PerformOperations: HRESULT; stdcall;
    function GetAnyOperationsAborted(var pfAnyOperationsAborted: BOOL): HRESULT; stdcall;

    property FileHandleUtil: TFileHandleUtil read FFileHandleUtil;
    property TaskDialogWindowHandle: THandle read FTaskDialogWindowHandle write FTaskDialogWindowHandle;
    property ParentWindowHandle: THandle read FParentWindowHandle write FParentWindowHandle;
  end;

implementation

uses
  App;

class procedure TFileOperation.ProcessDataObject(const Args: PProcessDataObjectArg);

  function ContainFormat(AFormat: TClipFormat; ATymed: LongInt; AAspect: LongInt = DVASPECT_CONTENT; AIndex: LongInt = -1): Boolean;
  var
    Format: TFormatEtc;
  begin
    ZeroMemory(@Format, SizeOf(Format));
    Format.cfFormat := AFormat;
    Format.dwAspect := AAspect;
    Format.lindex := AIndex;
    Format.tymed := ATymed;
    Result := Args.DataObject.QueryGetData(Format) = S_OK;
  end;

var
  Format: TFormatEtc;
  Medium: TStgMedium;
  DropFiles: PDROPFILES;
  i, FileCount: Integer;
  FilePathPtr: PWideChar;
  FilePath: string;
begin
  if not ContainFormat(CF_HDROP, TYMED_HGLOBAL) then
    Exit;

  Format.cfFormat := CF_HDROP;
  Format.Ptd := nil;
  Format.dwAspect := DVASPECT_CONTENT;
  Format.lindex := -1;
  Format.tymed := TYMED_HGLOBAL;
  ZeroMemory(@Medium, SizeOf(Medium));

  OleCheck(Args.DataObject.GetData(Format, Medium));
  if (Medium.tymed and TYMED_HGLOBAL = 0) or (Medium.hGlobal = 0) then
    raise Exception.Create('Invalid medium');

  try
    DropFiles := GlobalLock(Medium.hGlobal);
    if not Assigned(DropFiles) then
      RaiseLastOSError;

    FilePathPtr := GetMem(4096);
    try
      FileCount := DragQueryFileW(HDROP(DropFiles), $FFFFFFFF, nil, 0);

      for i := 0 to FileCount - 1 do
      begin
        DragQueryFileW(HDROP(DropFiles), i, FilePathPtr, MemSize(FilePathPtr) div 2);
        FilePath := FilePathPtr;

        Args.Instance.FDeleteFiles.Add(FilePath);
      end;
    finally
      GlobalUnlock(Medium.hGlobal);
      Freemem(FilePathPtr);
    end;
  finally
    ReleaseStgMedium(Medium);
  end;
end;

class procedure TFileOperation.CheckHandlesThreadProcWrapper(const FileOperation: TFileOperation); stdcall;
begin
  FileOperation.CheckHandlesThreadProc;
end;

procedure TFileOperation.CheckHandlesThreadProc;
begin
  try
    FileHandleUtil.CheckHandles(FDeleteFiles.ToStringArray, True);
  except
    on E: Exception do
      TFunctions.MessageBox(FParentWindowHandle, E.Message, SMsgDlgError, MB_ICONERROR);
  end;

  if (FileHandleUtil.OpenHandles.Count > 0) then
    TApp.ShowTaskDialog(FParentWindowHandle, FileHandleUtil, Self);
end;

procedure TFileOperation.FileOperationProgressSinkPreDeleteItem(Sender: TObject);
var
  ThreadId: DWORD;
begin
  TFileOperationProgressSink(Sender).OnPreDeleteItem := nil;

  if FDeleteFiles.Count > 0 then
    FThreadHandle := CreateThread(nil, 0, @TFileOperation.CheckHandlesThreadProcWrapper, Self, 0, ThreadId);
end;

constructor TFileOperation.Create(const WrappedFileOperation: IFileOperation);
var
  dwCookie: DWORD;
begin
  inherited Create;

  _AddRef;

  FWrappedFileOperation := WrappedFileOperation;

  FDeleteFiles := TStringList.Create;
  FFileHandleUtil := TFileHandleUtil.Create;

  FProgressSink := TFileOperationProgressSink.Create;
  FProgressSink.OnPreDeleteItem := FileOperationProgressSinkPreDeleteItem;

  Advise(FProgressSink, dwCookie);
end;

destructor TFileOperation.Destroy;
begin
  if FThreadHandle > 0 then
  begin
    WaitForSingleObject(FThreadHandle, INFINITE);
    CloseHandle(FThreadHandle);
  end;

  if FTaskDialogWindowHandle > 0 then
    PostMessage(FTaskDialogWindowHandle, WM_CLOSE, 0, 0);

  FFileHandleUtil.Free;
  FDeleteFiles.Free;

  inherited Destroy;
end;

function TFileOperation.Advise(const pfops: IFileOperationProgressSink; var pdwCookie: DWORD): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.Advise(pfops, pdwCookie);
end;

function TFileOperation.Unadvise(dwCookie: DWORD): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.Unadvise(dwCookie);
end;

function TFileOperation.SetOperationFlags(dwOperationFlags: DWORD): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.SetOperationFlags(dwOperationFlags);
end;

function TFileOperation.SetProgressMessage(pszMessage: LPCWSTR): HRESULT;
  stdcall;
begin
  Result := FWrappedFileOperation.SetProgressMessage(pszMessage);
end;

function TFileOperation.SetProgressDialog(const popd: Pointer): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.SetProgressDialog(popd);
end;

function TFileOperation.SetProperties(const pproparray: Pointer): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.SetProperties(pproparray);
end;

function TFileOperation.SetOwnerWindow(hwndParent: HWND): HRESULT; stdcall;
begin
  ParentWindowHandle := hwndParent;

  Result := FWrappedFileOperation.SetOwnerWindow(hwndParent);
end;

function TFileOperation.ApplyPropertiesToItem(const psiItem: IShellItem): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.ApplyPropertiesToItem(psiItem);
end;

function TFileOperation.ApplyPropertiesToItems(const punkItems: IUnknown): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.ApplyPropertiesToItems(punkItems);
end;

function TFileOperation.RenameItem(const psiItem: IShellItem; pszNewName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT;
  stdcall;
begin
  Result := FWrappedFileOperation.RenameItem(psiItem, pszNewName, pfopsItem);
end;

function TFileOperation.RenameItems(const pUnkItems: IUnknown; pszNewName: LPCWSTR): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.RenameItems(pUnkItems, pszNewName);
end;

function TFileOperation.MoveItem(const psiItem: IShellItem; const psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.MoveItem(psiItem, psiDestinationFolder, pszNewName, pfopsItem);
end;

function TFileOperation.MoveItems(const punkItems: IUnknown; const psiDestinationFolder: IShellItem): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.MoveItems(punkItems, psiDestinationFolder);
end;

function TFileOperation.CopyItem(const psiItem: IShellItem; const psiDestinationFolder: IShellItem; pszCopyName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.CopyItem(psiItem, psiDestinationFolder, pszCopyName, pfopsItem);
end;

function TFileOperation.CopyItems(const punkItems: IUnknown; const psiDestinationFolder: IShellItem): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.CopyItems(punkItems, psiDestinationFolder);
end;

function TFileOperation.DeleteItem(const psiItem: IShellItem; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.DeleteItem(psiItem, pfopsItem);
end;

function TFileOperation.DeleteItems(const punkItems: IUnknown): HRESULT; stdcall;
var
  Arg: TProcessDataObjectArg;
begin
  // TODO: IFileOperation can be called in silent mode. If this is the case maybe What The Lock should do nothing

  Result := FWrappedFileOperation.DeleteItems(punkItems);

  if punkItems is IDataObject then
  begin
    Arg.DataObject := punkItems as IDataObject;
    Arg.Instance := Self;

    try
      ProcessDataObject(@Arg);
    except
      on E: Exception do
        TFunctions.MessageBox(0, E.Message, SMsgDlgError, MB_ICONERROR);
    end;
  end;
end;

function TFileOperation.NewItem(const psiDestinationFolder: IShellItem; dwFileAttributes: DWORD; pszName: LPCWSTR; pszTemplateName: LPCWSTR; const pfopsItem: IFileOperationProgressSink): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.NewItem(psiDestinationFolder, dwFileAttributes, pszName, pszTemplateName, pfopsItem);
end;

function TFileOperation.PerformOperations: HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.PerformOperations;
end;

function TFileOperation.GetAnyOperationsAborted(var pfAnyOperationsAborted: BOOL): HRESULT; stdcall;
begin
  Result := FWrappedFileOperation.GetAnyOperationsAborted(pfAnyOperationsAborted);
end;

end.
