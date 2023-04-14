unit App;

interface

uses
  Classes,
  ComObj,
  Controls,
  DDetours,
  Dialogs,
  ExtCtrls,
  FileCtrl,
  FileHandleUtil,
  FileOperation,
  Forms,
  Functions,
  Generics.Collections,
  Graphics,
  LCLTaskDialog,
  RtlConsts,
  SysUtils,
  windows;

type

  { TApp }

  TApp = class
    type
      TTaskDialogConfig = packed record
        cbSize: Integer;
        hwndParent: HWND;
        hInstance: THandle;
        dwFlags: Cardinal;
        dwCommonButtons: Cardinal;
        pszWindowTitle: PWideChar;
        hMainIcon: HICON;
        pszMainInstruction: PWideChar;
        pszContent: PWideChar;
        cButtons: Integer;
        pButtons: Pointer;
        nDefaultButton: Integer;
        cRadioButtons: Integer;
        pRadioButtons: Pointer;
        nDefaultRadioButton: Integer;
        pszVerificationText: PWideChar;
        pszExpandedInformation: PWideChar;
        pszExpandedControlText: PWideChar;
        pszCollapsedControlText: PWideChar;
        hFooterIcon: HICON;
        pszFooter: PWideChar;
        pfCallback: Pointer;
        lpCallbackData: Pointer;
        cxWidth: Integer;
      end;
      PTaskDialogConfig = ^TTaskDialogConfig;

      TShowClassDialogArgs = record
        ParentWindowHandle: HWND;
        FileHandleUtil: TFileHandleUtil;
        FileOperation: TFileOperation;
      end;
      PShowClassDialogArgs = ^TShowClassDialogArgs;
  private
  class threadvar
    FNextTaskDialogArgs: PShowClassDialogArgs;
  class var
    OCoCreateInstance: function(_para1: PGUID; _para2: Pointer; _para3: DWORD; _para4, _para5: Pointer): HRESULT; stdcall;
    OTaskDialogIndirect: function(AConfig: PTASKDIALOGCONFIG; Res: PInteger; ResRadio: PInteger; VerifyFlag: PBOOL): HRESULT; stdcall;

    class function TaskDialogCallbackProc(hwnd: HWND; uNotification: UINT; wParam: WPARAM; lParam: LPARAM; dwRefData: Pointer): HRESULT; stdcall; static;

    class procedure ShowTaskDialogProc(Args: PShowClassDialogArgs); stdcall; static;
    class function HCoCreateInstance(_para1: PGUID; _para2: Pointer; _para3: DWORD; _para4, _para5: Pointer): HRESULT; stdcall; static;
    class function HTaskDialogIndirect(AConfig: PTASKDIALOGCONFIG; Res: PInteger; ResRadio: PInteger; VerifyFlag: PBOOL): HRESULT; static;
  public
    class procedure Initialize; static;
    class procedure ShowTaskDialog(const ParentWindowHandle: HWND; const FileHandleUtil: TFileHandleUtil; const FileOperation: TFileOperation); overload; static;
  end;

implementation

resourcestring
  SFileProcess = 'File "%s" is in use by process %s.';
  SFilesProcess = '%d files are in use by process %s.';
  SFileProcesses = 'File "%s" is in use by %d processes.';
  SFilesProcesses = '%d files are in use by %d processes.';
  SRelease = 'Release file(s)';
  STerminate = 'Terminate process(es)';
  SErrorMessageTerminate = 'At least one process could not be terminated:';
  SErrorMessageRelease = 'At least one handle could not be released:';

{ TApp }

class procedure TApp.ShowTaskDialogProc(Args: PShowClassDialogArgs); stdcall;

  function MinimizeName(const Filename: string; const MaxWidth: Integer): string;
  var
    Canvas: TCanvas;
  begin
    Canvas := TCanvas.Create;
    try
      Canvas.Handle := GetDC(GetDesktopWindow);
      try
        Result := FileCtrl.MinimizeName(Filename, Canvas, MaxWidth);
      finally
        ReleaseDC(GetDesktopWindow, Canvas.Handle);
      end;
    finally
      Canvas.Free;
    end;
  end;

const
  ProcessString = '"%s" (%d)';
var
  i: Integer;
  OpenedManually, Exists: Boolean;
  ProcessedFileHandle: TFileHandle;
  Text, ExpandedText: string;
  TaskDialog: Dialogs.TTaskDialog;
  Button: TTaskDialogBaseButtonItem;
  OpenHandle: TFileHandle;
  Processes, Filenames: TList<string>;
  Processed: TList<TFileHandle>;
begin
  OpenedManually := not Assigned(Args.FileOperation);

  try
    Processes := TList<string>.Create;
    Filenames := TList<string>.Create;
    Processed := TList<TFileHandle>.Create;
    try
      for OpenHandle in Args.FileHandleUtil.OpenHandles do
      begin
        Exists := False;
        for ProcessedFileHandle in Processed do
          if (ProcessedFileHandle.NtFilename = OpenHandle.NtFilename) and (ProcessedFileHandle.PID = OpenHandle.PID) then
          begin
            Exists := True;
            Break;
          end;

        if Exists then
          Continue;

        Processed.Add(TFileHandle.Create(OpenHandle.PID, ExtractFileName(OpenHandle.ImageName), OpenHandle.NtFilename, OpenHandle.Filename, 0));

        if not Processes.Contains(ProcessString.Format([ExtractFileName(OpenHandle.ImageName), OpenHandle.PID])) then
          Processes.Add(ProcessString.Format([ExtractFileName(OpenHandle.ImageName), OpenHandle.PID]));
        if not Filenames.Contains(OpenHandle.Filename) then
          Filenames.Add(OpenHandle.Filename);
      end;
      ExpandedText := '';

      if (Processes.Count = 1) and (Filenames.Count = 1) then
        Text := SFileProcess.Format([Filenames[0], Processes[0]])
      else if Processes.Count = 1 then
      begin
        Text := SFilesProcess.Format([Filenames.Count, Processes[0]]);
        for i := 0 to Filenames.Count - 1 do
          ExpandedText += '- %s'#13#10''.Format([Filenames[i]]);
      end else if Filenames.Count = 1 then
      begin
        Text := SFileProcesses.Format([Filenames[0], Processes.Count]);
        for i := 0 to Processes.Count - 1 do
          ExpandedText += '- %s'#13#10''.Format([Processes[i]]);
      end else
      begin
        Text := SFilesProcesses.Format([Filenames.Count, Processes.Count]);
        for ProcessedFileHandle in Processed do
          ExpandedText += ('- "%s" ➔ ' + ProcessString + #13#10).Format([MinimizeName(ProcessedFileHandle.Filename, 200), ProcessedFileHandle.ImageName, ProcessedFileHandle.PID]);
      end;
    finally
      for ProcessedFileHandle in Processed do
        ProcessedFileHandle.Free;

      Processes.Free;
      Filenames.Free;
      Processed.Free;
    end;

    TaskDialog := Dialogs.TTaskDialog.Create(Application);
    TaskDialog.Caption := IfThen<string>(OpenedManually, SMsgDlgInformation, SMsgDlgWarning);
    TaskDialog.MainIcon := IfThen < Dialogs.TTaskDialogIcon > (OpenedManually, tdiInformation, tdiWarning);
    TaskDialog.CommonButtons := [tcbCancel];
    TaskDialog.DefaultButton := tcbCancel;
    TaskDialog.Flags := TaskDialog.Flags + [tfUseCommandLinks];

    TaskDialog.Text := Text;
    TaskDialog.ExpandedText := ExpandedText;

    Button := TaskDialog.Buttons.Add;
    Button.Caption := SRelease;
    Button := TaskDialog.Buttons.Add;
    Button.Caption := STerminate;

    FNextTaskDialogArgs := Args;

    if not TaskDialog.Execute then
      Exit;

    try
      case TaskDialog.ModalResult of
        100: Args.FileHandleUtil.CloseHandles(Args.FileHandleUtil.OpenHandles.ToArray, False);
        101: Args.FileHandleUtil.CloseHandles(Args.FileHandleUtil.OpenHandles.ToArray, True);
      end;
    except
      on E: Exception do
        if TaskDialog.ModalResult = 100 then
          TFunctions.MessageBox(Args.ParentWindowHandle, SErrorMessageRelease + LineEnding + E.Message, SMsgDlgError, MB_ICONERROR)
        else
          TFunctions.MessageBox(Args.ParentWindowHandle, SErrorMessageTerminate + LineEnding + E.Message, SMsgDlgError, MB_ICONERROR);
    end;
  finally
    if not Assigned(Args.FileOperation) then
      Args.FileHandleUtil.Free;

    Dispose(Args);
  end;
end;

class procedure TApp.ShowTaskDialog(const ParentWindowHandle: HWND; const FileHandleUtil: TFileHandleUtil; const FileOperation: TFileOperation);
var
  ThreadId: DWORD;
  Args: PShowClassDialogArgs;
begin
  New(Args);
  Args.ParentWindowHandle := ParentWindowHandle;
  Args.FileHandleUtil := FileHandleUtil;
  Args.FileOperation := FileOperation;

  CloseHandle(CreateThread(nil, 0, @TApp.ShowTaskDialogProc, Args, 0, ThreadId));
end;

class function TApp.HCoCreateInstance(_para1: PGUID; _para2: Pointer; _para3: DWORD; _para4, _para5: Pointer): HRESULT; stdcall;
var
  FileOperation: TFileOperation;
begin
  Result := OCoCreateInstance(_para1, _para2, _para3, _para4, _para5);

  if Succeeded(Result) and IsEqualGUID(_para1^, CLSID_FileOperation) then
  begin
    FileOperation := TFileOperation.Create(IFileOperation(_para5));
    _para5 := FileOperation as IFileOperation;
  end;
end;

class function TApp.TaskDialogCallbackProc(hwnd: HWND; uNotification: UINT; wParam: WPARAM; lParam: LPARAM; dwRefData: Pointer): HRESULT; stdcall;
const
  TDN_DESTROYED = 5;
  TDN_DIALOG_CONSTRUCTED = 7;
var
  Args: PShowClassDialogArgs absolute dwRefData;
  TaskDialogRect, ParentRect: TRect;
begin
  Result := S_OK;

  case uNotification of
    TDN_DIALOG_CONSTRUCTED:
    begin
      if Assigned(Args.FileOperation) then
        Args.FileOperation.TaskDialogWindowHandle := hwnd;

      if Args.ParentWindowHandle <> 0 then
      begin
        GetWindowRect(hwnd, TaskDialogRect);
        GetWindowRect(Args.ParentWindowHandle, ParentRect);

        SetWindowLongPtrW(hwnd, GWLP_HWNDPARENT, Args.ParentWindowHandle);

        SetWindowPos(hwnd, 0, ParentRect.Left + (ParentRect.Width div 2) - (TaskDialogRect.Width div 2), ParentRect.Top + (ParentRect.Height div 2) - (TaskDialogRect.Height div 2), 0, 0, SWP_NOSIZE or SWP_NOZORDER);
      end;
    end;
    TDN_DESTROYED:
      if Assigned(Args.FileOperation) then
        Args.FileOperation.TaskDialogWindowHandle := 0;
  end;
end;

class function TApp.HTaskDialogIndirect(AConfig: PTASKDIALOGCONFIG; Res: PInteger; ResRadio: PInteger; VerifyFlag: PBOOL): HRESULT;
begin
  if Assigned(FNextTaskDialogArgs) then
  begin
    AConfig.pfCallback := @TaskDialogCallbackProc;
    AConfig.lpCallbackData := FNextTaskDialogArgs;
    AConfig.cxWidth := 250;

    FNextTaskDialogArgs := nil;
  end;

  Result := OTaskDialogIndirect(AConfig, Res, ResRadio, VerifyFlag);
end;

class procedure TApp.Initialize;
var
  Func: Pointer;
begin
  Func := GetProcAddress(GetModuleHandle('ole32.dll'), 'CoCreateInstance');
  @OCoCreateInstance := InterceptCreate(Func, @HCoCreateInstance);

  @OTaskDialogIndirect := @TaskDialogIndirect;
  TaskDialogIndirect := @HTaskDialogIndirect;
end;

end.
