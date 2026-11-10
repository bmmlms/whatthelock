unit App;

interface

uses
  Classes,
  ComObj,
  DDetours,
  FileHandleUtil,
  FileOperation,
  Functions,
  Generics.Collections,
  RtlConsts,
  SysUtils,
  TaskDialog,
  windows;

type

  { TShowClassDialogArgs }

  TShowClassDialogArgs = record
    ParentWindowHandle: HWND;
    FileHandleUtil: TFileHandleUtil;
    FileOperation: TFileOperation;
  end;
  PShowClassDialogArgs = ^TShowClassDialogArgs;

  { TApp }

  TApp = class
  private
  class var
    OCoCreateInstance: function(_para1: TCLSID; _para2: Pointer; _para3: DWORD; _para4: TIID; out _para5: Pointer): HRESULT; stdcall;

    class function TaskDialogCallbackProc(hwnd: HWND; uNotification: UINT; wParam: WPARAM; lParam: LPARAM; dwRefData: LONG_PTR): HRESULT; stdcall; static;

    class procedure ShowTaskDialogProc(Args: PShowClassDialogArgs); stdcall; static;
    class function HCoCreateInstance(_para1: TCLSID; _para2: Pointer; _para3: DWORD; _para4: TIID; out _para5: Pointer): HRESULT; stdcall; static;
  public
    class procedure Initialize; static;
    class procedure ShowTaskDialog(const ParentWindowHandle: HWND; const FileHandleUtil: TFileHandleUtil; const FileOperation: TFileOperation); overload; static;
  end;

resourcestring
  SFileProcess = 'File "%s" is in use by process %s.';
  SFilesProcess = '%d files are in use by process %s.';
  SFileProcesses = 'File "%s" is in use by %d processes.';
  SFilesProcesses = '%d files are in use by %d processes.';
  SRelease = 'Release file(s)';
  STerminate = 'Terminate process(es)';
  SErrorMessageTerminate = 'At least one process could not be terminated:';
  SErrorMessageRelease = 'At least one handle could not be released:';

implementation

{ TApp }

class procedure TApp.ShowTaskDialogProc(Args: PShowClassDialogArgs); stdcall;

  function MinimizeName(const FileName: UnicodeString; const MaxWidth: Integer): UnicodeString;
  var
    DC: HDC;
    Drive, Directory, Name, ShortName: UnicodeString;
    Separator, Width: Integer;

    function TextWidth(const Text: UnicodeString): Integer;
    var
      Size: TSize;
    begin
      if GetTextExtentPoint32W(DC, PWideChar(Text), Length(Text), Size) then
        Result := Size.cx
      else
        Result := -1;
    end;

  begin
    Result := FileName;
    if Pos(PathDelim, FileName) = 0 then
      Exit;

    DC := GetDC(0);
    if DC = 0 then
      Exit;

    try
      Width := TextWidth(FileName);
      if (Width < 0) or (Width <= MaxWidth) then
        Exit;

      Drive := ExtractFileDrive(FileName);
      Directory := ExtractFilePath(FileName);
      Name := ExtractFileName(FileName);
      Delete(Directory, 1, Length(Drive));
      while (Directory <> '') and (Directory[1] in ['\', '/']) do
      begin
        Drive := Drive + Directory[1];
        Delete(Directory, 1, 1);
      end;

      while Directory <> '' do
      begin
        Separator := Pos(PathDelim, Directory);
        if Separator = 0 then
          Break;
        Delete(Directory, 1, Separator);
        ShortName := Drive + '...' + PathDelim + Directory + Name;
        Width := TextWidth(ShortName);
        if Width < 0 then
          Exit;
        if Width <= MaxWidth then
        begin
          Result := ShortName;
          Exit;
        end;
      end;
      Result := Name;
    finally
      ReleaseDC(0, DC);
    end;
  end;

const
  ID_RELEASE = 100;
  ID_TERMINATE = 101;
  ProcessString = '"%s" (%d)';
var
  i: Integer;
  OpenedManually, Exists: Boolean;
  ProcessedFileHandle: TFileHandle;
  Text: UnicodeString;
  ExpandedText: UnicodeString = '';
  Title, ReleaseButtonText, TerminateButtonText: UnicodeString;
  Config: TTaskDialogConfig;
  Buttons: array[0..1] of TTaskDialogButton;
  Button: Integer = 0;
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

    Title := IfThen<string>(OpenedManually, SMsgDlgInformation, SMsgDlgWarning);
    ReleaseButtonText := SRelease;
    TerminateButtonText := STerminate;

    Buttons[0].nButtonID := ID_RELEASE;
    Buttons[0].pszButtonText := PWideChar(ReleaseButtonText);
    Buttons[1].nButtonID := ID_TERMINATE;
    Buttons[1].pszButtonText := PWideChar(TerminateButtonText);

    ZeroMemory(@Config, SizeOf(Config));
    Config.cbSize := SizeOf(Config);
    Config.hwndParent := Args.ParentWindowHandle;
    Config.dwFlags := TDF_USE_COMMAND_LINKS or TDF_ALLOW_DIALOG_CANCELLATION;
    Config.dwCommonButtons := TDCBF_CANCEL_BUTTON;
    Config.pszWindowTitle := PWideChar(Title);
    Config.pszMainIcon := IfThen<Pointer>(OpenedManually, TD_INFORMATION_ICON, TD_WARNING_ICON);
    Config.pszContent := PWideChar(Text);
    Config.cButtons := Length(Buttons);
    Config.pButtons := @Buttons[0];
    Config.nDefaultButton := idCancel;
    Config.pszExpandedInformation := PWideChar(ExpandedText);
    Config.pfCallback := @TaskDialogCallbackProc;
    Config.lpCallbackData := LONG_PTR(PtrUInt(Args));
    Config.cxWidth := 250;

    if Failed(TaskDialog.ShowTaskDialog(Config, Button)) then
      Exit;

    try
      case Button of
        ID_RELEASE: Args.FileHandleUtil.CloseHandles(Args.FileHandleUtil.OpenHandles.ToArray, False);
        ID_TERMINATE: Args.FileHandleUtil.CloseHandles(Args.FileHandleUtil.OpenHandles.ToArray, True);
      end;
    except
      on E: Exception do
        if Button = ID_RELEASE then
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
  Args: PShowClassDialogArgs;
begin
  New(Args);
  Args.ParentWindowHandle := ParentWindowHandle;
  Args.FileHandleUtil := FileHandleUtil;
  Args.FileOperation := FileOperation;

  TApp.ShowTaskDialogProc(Args);
end;

class function TApp.HCoCreateInstance(_para1: TCLSID; _para2: Pointer; _para3: DWORD; _para4: TIID; out _para5: Pointer): HRESULT; stdcall;
var
  FileOperation: TFileOperation;
begin
  Result := OCoCreateInstance(_para1, _para2, _para3, _para4, _para5);

  if Succeeded(Result) and IsEqualGUID(_para1, CLSID_FileOperation) then
  begin
    FileOperation := TFileOperation.Create(IFileOperation(_para5));
    _para5 := FileOperation as IFileOperation;
  end;
end;

class function TApp.TaskDialogCallbackProc(hwnd: HWND; uNotification: UINT; wParam: WPARAM; lParam: LPARAM; dwRefData: LONG_PTR): HRESULT; stdcall;
var
  Args: PShowClassDialogArgs;
  TaskDialogRect, ParentRect: TRect;
begin
  Args := PShowClassDialogArgs(PtrUInt(dwRefData));
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

class procedure TApp.Initialize;
var
  Func: Pointer;
begin
  Func := GetProcAddress(GetModuleHandle('ole32.dll'), 'CoCreateInstance');
  @OCoCreateInstance := InterceptCreate(Func, @HCoCreateInstance);
end;

end.
