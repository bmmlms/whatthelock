unit TaskDialog;

interface

uses
  JwaWinBase,
  Windows;

type
  TTaskDialogButton = packed record
    nButtonID: Integer;
    pszButtonText: PWideChar;
  end;
  PTaskDialogButton = ^TTaskDialogButton;

  TTaskDialogCallback = function(Window: HWND; Notification: UINT; WParam: WPARAM; LParam: LPARAM; RefData: LONG_PTR): HRESULT; stdcall;

  TTaskDialogConfig = packed record
    cbSize: UINT;
    hwndParent: HWND;
    hInstance: HMODULE;
    dwFlags: DWORD;
    dwCommonButtons: DWORD;
    pszWindowTitle: PWideChar;
    pszMainIcon: Pointer;
    pszMainInstruction: PWideChar;
    pszContent: PWideChar;
    cButtons: UINT;
    pButtons: PTaskDialogButton;
    nDefaultButton: Integer;
    cRadioButtons: UINT;
    pRadioButtons: PTaskDialogButton;
    nDefaultRadioButton: Integer;
    pszVerificationText: PWideChar;
    pszExpandedInformation: PWideChar;
    pszExpandedControlText: PWideChar;
    pszCollapsedControlText: PWideChar;
    pszFooterIcon: Pointer;
    pszFooter: PWideChar;
    pfCallback: TTaskDialogCallback;
    lpCallbackData: LONG_PTR;
    cxWidth: UINT;
  end;
  PTaskDialogConfig = ^TTaskDialogConfig;
  {$PACKRECORDS DEFAULT}

  TTaskDialogIndirectProc = function(Config: PTaskDialogConfig; Button: PInteger; RadioButton: PInteger; VerificationFlagChecked: PBOOL): HRESULT; stdcall;

const
  TDF_ENABLE_HYPERLINKS = $00000001;
  TDF_USE_HICON_MAIN = $00000002;
  TDF_USE_HICON_FOOTER = $00000004;
  TDF_ALLOW_DIALOG_CANCELLATION = $00000008;
  TDF_USE_COMMAND_LINKS = $00000010;
  TDF_USE_COMMAND_LINKS_NO_ICON = $00000020;
  TDF_EXPAND_FOOTER_AREA = $00000040;
  TDF_EXPANDED_BY_DEFAULT = $00000080;
  TDF_VERIFICATION_FLAG_CHECKED = $00000100;
  TDF_SHOW_PROGRESS_BAR = $00000200;
  TDF_SHOW_MARQUEE_PROGRESS_BAR = $00000400;
  TDF_CALLBACK_TIMER = $00000800;
  TDF_POSITION_RELATIVE_TO_WINDOW = $00001000;
  TDF_RTL_LAYOUT = $00002000;
  TDF_NO_DEFAULT_RADIO_BUTTON = $00004000;
  TDF_CAN_BE_MINIMIZED = $00008000;

  TDCBF_OK_BUTTON = $0001;
  TDCBF_YES_BUTTON = $0002;
  TDCBF_NO_BUTTON = $0004;
  TDCBF_CANCEL_BUTTON = $0008;
  TDCBF_RETRY_BUTTON = $0010;
  TDCBF_CLOSE_BUTTON = $0020;

  TD_WARNING_ICON = PWideChar(PtrUInt($FFFF));
  TD_ERROR_ICON = PWideChar(PtrUInt($FFFE));
  TD_INFORMATION_ICON = PWideChar(PtrUInt($FFFD));
  TD_SHIELD_ICON = PWideChar(PtrUInt($FFFC));

  TDN_CREATED = 0;
  TDN_NAVIGATED = 1;
  TDN_BUTTON_CLICKED = 2;
  TDN_HYPERLINK_CLICKED = 3;
  TDN_TIMER = 4;
  TDN_DESTROYED = 5;
  TDN_RADIO_BUTTON_CLICKED = 6;
  TDN_DIALOG_CONSTRUCTED = 7;
  TDN_VERIFICATION_CLICKED = 8;
  TDN_HELP = 9;
  TDN_EXPANDO_BUTTON_CLICKED = 10;

function IsTaskDialogAvailable: Boolean;
function ShowTaskDialog(const Config: TTaskDialogConfig; out Button: Integer): HRESULT;

implementation

const
  TASKDIALOG_MANIFEST_RESOURCE_ID = 2;
  TASKDIALOG_ACTCTX_RESOURCE_NAME_VALID = $00000008;
  TASKDIALOG_ACTCTX_HMODULE_VALID = $00000080;

function HResultFromWin32(ErrorCode: DWORD): HRESULT;
begin
  if ErrorCode = 0 then
    Result := HRESULT($80004005)
  else
    Result := HRESULT($80070000 or (ErrorCode and $0000FFFF));
end;

procedure CloseTaskDialogContext(var ActCtx: HANDLE; var Cookie: ULONG_PTR; var IsActive: Boolean; var Module: HMODULE);
begin
  if Module <> 0 then
  begin
    FreeLibrary(Module);
    Module := 0;
  end;

  if IsActive then
  begin
    DeactivateActCtx(0, Cookie);
    IsActive := False;
    Cookie := 0;
  end;

  if ActCtx <> 0 then
  begin
    ReleaseActCtx(ActCtx);
    ActCtx := 0;
  end;
end;

function OpenTaskDialogContext(out ActCtx: HANDLE; out Cookie: ULONG_PTR; out IsActive: Boolean; out Module: HMODULE; out DialogProc: TTaskDialogIndirectProc): HRESULT;
var
  Activation: ACTCTXW;
  EntryPoint: Pointer;
begin
  ActCtx := 0;
  Cookie := 0;
  IsActive := False;
  Module := 0;
  DialogProc := nil;

  ZeroMemory(@Activation, SizeOf(Activation));
  Activation.cbSize := SizeOf(Activation);
  Activation.dwFlags := TASKDIALOG_ACTCTX_RESOURCE_NAME_VALID or TASKDIALOG_ACTCTX_HMODULE_VALID;
  Activation.lpResourceName := PWideChar(PtrUInt(TASKDIALOG_MANIFEST_RESOURCE_ID));
  Activation.hModule := HInstance;

  ActCtx := CreateActCtxW(Activation);
  if ActCtx = INVALID_HANDLE_VALUE then
  begin
    ActCtx := 0;
    Exit(HResultFromWin32(GetLastError));
  end;

  if not ActivateActCtx(ActCtx, Cookie) then
  begin
    Result := HResultFromWin32(GetLastError);
    CloseTaskDialogContext(ActCtx, Cookie, IsActive, Module);
    Exit;
  end;
  IsActive := True;

  Module := LoadLibraryW('comctl32.dll');
  if Module = 0 then
  begin
    Result := HResultFromWin32(GetLastError);
    CloseTaskDialogContext(ActCtx, Cookie, IsActive, Module);
    Exit;
  end;


  EntryPoint := GetProcAddress(Module, 'TaskDialogIndirect');
  if EntryPoint = nil then
  begin
    Result := HResultFromWin32(ERROR_PROC_NOT_FOUND);
    CloseTaskDialogContext(ActCtx, Cookie, IsActive, Module);
    Exit;
  end;

  DialogProc := TTaskDialogIndirectProc(EntryPoint);
  Result := HRESULT(0);
end;

function IsTaskDialogAvailable: Boolean;
var
  ActCtx: HANDLE;
  Cookie: ULONG_PTR;
  IsActive: Boolean;
  Module: HMODULE;
  DialogProc: TTaskDialogIndirectProc;
begin
  Result := Succeeded(OpenTaskDialogContext(ActCtx, Cookie, IsActive, Module, DialogProc));
  if Result then
    CloseTaskDialogContext(ActCtx, Cookie, IsActive, Module);
end;

function ShowTaskDialog(const Config: TTaskDialogConfig; out Button: Integer): HRESULT;
var
  ActCtx: HANDLE;
  Cookie: ULONG_PTR;
  IsActive: Boolean;
  Module: HMODULE;
  DialogProc: TTaskDialogIndirectProc;
begin
  Button := 0;
  Result := OpenTaskDialogContext(ActCtx, Cookie, IsActive, Module, DialogProc);
  if Failed(Result) then
    Exit;

  try
    Result := DialogProc(@Config, @Button, nil, nil);
  finally
    CloseTaskDialogContext(ActCtx, Cookie, IsActive, Module);
  end;
end;

end.
