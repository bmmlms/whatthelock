unit FileOperationProgressSink;

interface

uses
  Classes,
  shlobj,
  windows;

type

  { TFileOperationProgressSink }

  // TODO: Forward to original IFileOperationProgressSink

  TFileOperationProgressSink = class(TInterfacedObject, IFileOperationProgressSink)
  private
    FOnPreDeleteItem: TNotifyEvent;
  protected
    function StartOperations: HRESULT; stdcall;
    function FinishOperations(hrResult: HRESULT): HRESULT; stdcall;
    function PreRenameItem(dwFlags: DWORD; psiItem: IShellItem; pszNewName: LPCWSTR): HRESULT; stdcall;
    function PostRenameItem(dwFlags: DWORD; psiItem: IShellItem; pszNewName: LPCWSTR; hrRename: HRESULT; psiNewlyCreated: IShellItem): HRESULT; stdcall;
    function PreMoveItem(dwFlags: DWORD; psiItem: IShellItem; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR): HRESULT; stdcall;
    function PostMoveItem(dwFlags: DWORD; psiItem: IShellItem; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; hrMove: HRESULT; psiNewlyCreated: IShellItem): HRESULT; stdcall;
    function PreCopyItem(dwFlags: DWORD; psiItem: IShellItem; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR): HRESULT; stdcall;
    function PostCopyItem(dwFlags: DWORD; psiItem: IShellItem; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; hrCopy: HRESULT; psiNewlyCreated: IShellItem): HRESULT; stdcall;
    function PreDeleteItem(dwFlags: DWORD; psiItem: IShellItem): HRESULT; stdcall;
    function PostDeleteItem(dwFlags: DWORD; psiItem: IShellItem; hrDelete: HRESULT; psiNewlyCreated: IShellItem): HRESULT; stdcall;
    function PreNewItem(dwFlags: DWORD; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR): HRESULT; stdcall;
    function PostNewItem(dwFlags: DWORD; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; pszTemplateName: LPCWSTR; dwFileAttributes: DWORD; hrNew: HRESULT; psiNewItem: IShellItem): HRESULT; stdcall;
    function UpdateProgress(iWorkTotal: UINT; iWorkSoFar: UINT): HRESULT; stdcall;
    function ResetTimer: HRESULT; stdcall;
    function PauseTimer: HRESULT; stdcall;
    function ResumeTimer: HRESULT; stdcall;
  public
    property OnPreDeleteItem: TNotifyEvent read FOnPreDeleteItem write FOnPreDeleteItem;
  end;

implementation

{ TFileOperationProgressSink }

function TFileOperationProgressSink.StartOperations: HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.FinishOperations(hrResult: HRESULT): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PreRenameItem(dwFlags: DWORD; psiItem: IShellItem; pszNewName: LPCWSTR): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PostRenameItem(dwFlags: DWORD; psiItem: IShellItem; pszNewName: LPCWSTR; hrRename: HRESULT; psiNewlyCreated: IShellItem): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PreMoveItem(dwFlags: DWORD; psiItem: IShellItem; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PostMoveItem(dwFlags: DWORD; psiItem: IShellItem; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; hrMove: HRESULT; psiNewlyCreated: IShellItem): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PreCopyItem(dwFlags: DWORD; psiItem: IShellItem; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PostCopyItem(dwFlags: DWORD; psiItem: IShellItem; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; hrCopy: HRESULT; psiNewlyCreated: IShellItem): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PreDeleteItem(dwFlags: DWORD; psiItem: IShellItem): HRESULT; stdcall;
begin
  Result := S_OK;

  if Assigned(FOnPreDeleteItem) then
    FOnPreDeleteItem(Self);
end;

function TFileOperationProgressSink.PostDeleteItem(dwFlags: DWORD; psiItem: IShellItem; hrDelete: HRESULT; psiNewlyCreated: IShellItem): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PreNewItem(dwFlags: DWORD; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PostNewItem(dwFlags: DWORD; psiDestinationFolder: IShellItem; pszNewName: LPCWSTR; pszTemplateName: LPCWSTR; dwFileAttributes: DWORD; hrNew: HRESULT; psiNewItem: IShellItem): HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.UpdateProgress(iWorkTotal: UINT; iWorkSoFar: UINT): HRESULT;
  stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.ResetTimer: HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.PauseTimer: HRESULT; stdcall;
begin
  Result := S_OK;
end;

function TFileOperationProgressSink.ResumeTimer: HRESULT; stdcall;
begin
  Result := S_OK;
end;

end.
