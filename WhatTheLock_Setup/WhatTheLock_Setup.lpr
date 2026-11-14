program WhatTheLock_Setup;

uses
  ActiveX,
  Classes,
  Constants,
  Functions,
  LazFileUtils,
  LazUTF8,
  Paths,
  shlobj,
  SysUtils,
  TaskDialog,
  UnicodePaths,
  Windows;

  {$R *.res}
  {$R ..\Shared\CommonControls.rc}

const
  PARAM_ALLUSERS = '/AllUsers';

const
  TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE = WM_USER + 115;

  ID_ALLUSERS = 100;
  ID_CURRENTUSER = 101;

function TaskDialogCallback(hwnd: HWND; uNotification: UINT; wParam: WPARAM; lParam: LPARAM; dwRefData: LONG_PTR): HRESULT; stdcall;
begin
  Result := S_OK;

  if (uNotification = TDN_CREATED) and (not IsUserAnAdmin) then
    SendMessageW(hwnd, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, ID_ALLUSERS, 1);
end;

function AskScope(const DefaultScope: TInstallScope; out Scope: TInstallScope): Boolean;
var
  Config: TTaskDialogConfig;
  Buttons: array[0..1] of TTaskDialogButton;
  Title, Instruction, Content, AllUsersText, CurrentUserText: UnicodeString;
  Button: Integer;
  Status: HRESULT;
begin
  Result := False;
  Scope := DefaultScope;

  Title := UnicodeString(APPNAME) + ' Setup';
  Instruction := 'This will install/update ' + UnicodeString(APPNAME) + '.';
  Content := 'Who should be able to use it?';
  AllUsersText := 'For all users'#10'Installs to "' + TPaths.GetInstallDir(scAllUsers) + '". Requires administrator rights.';
  CurrentUserText := 'Only for me'#10'Installs to "' + TPaths.GetInstallDir(scCurrentUser) + '". No administrator rights required.';

  Buttons[0].nButtonID := ID_ALLUSERS;
  Buttons[0].pszButtonText := PWideChar(AllUsersText);
  Buttons[1].nButtonID := ID_CURRENTUSER;
  Buttons[1].pszButtonText := PWideChar(CurrentUserText);

  ZeroMemory(@Config, SizeOf(Config));
  Config.cbSize := SizeOf(Config);
  Config.hInstance := HInstance;
  Config.dwFlags := TDF_USE_COMMAND_LINKS or TDF_ALLOW_DIALOG_CANCELLATION;
  Config.dwCommonButtons := TDCBF_CANCEL_BUTTON;
  Config.pszWindowTitle := PWideChar(Title);
  Config.pszMainInstruction := PWideChar(Instruction);
  Config.pszContent := PWideChar(Content);
  Config.cButtons := Length(Buttons);
  Config.pButtons := @Buttons[0];
  Config.nDefaultButton := IfThen<LongInt>(DefaultScope = scCurrentUser, ID_CURRENTUSER, ID_ALLUSERS);
  Config.pfCallback := @TaskDialogCallback;

  Button := 0;
  Status := ShowTaskDialog(Config, Button);
  if Failed(Status) then
    raise Exception.Create('Unable to display the installation scope dialog (HRESULT 0x%s)'.Format([IntToHex(LongWord(Status), 8)]));

  case Button of
    ID_ALLUSERS: Scope := scAllUsers;
    ID_CURRENTUSER: Scope := scCurrentUser;
    else
      Exit(False);
  end;

  Result := True;
end;

function ExtractResource(const ResourceName: AnsiString; const FilePath: UnicodeString): Boolean;
var
  ResStream: TResourceStream;
begin
  Result := False;
  try
    ResStream := TResourceStream.Create(HInstance, ResourceName, RT_RCDATA);
    try
      UnicodeSaveStreamToFile(ResStream, FilePath);
      Result := True;
    finally
      ResStream.Free;
    end;
  except
  end;
end;

procedure CreateUninstallEntry(const Executable: UnicodeString);
var
  Disposition: DWORD;
  RegistryKey: HKEY;
  KeyName, Value: UnicodeString;
  RegistryResult: LongInt;

  procedure WriteStringValue(const Name, Data: UnicodeString);
  begin
    RegistryResult := RegSetValueExW(RegistryKey, PWideChar(Name), 0, REG_SZ, PByte(PWideChar(Data)), (Length(Data) + 1) * SizeOf(WideChar));
    if RegistryResult <> ERROR_SUCCESS then
    begin
      SetLastError(RegistryResult);
      RaiseLastOSError;
    end;
  end;

  procedure WriteIntegerValue(const Name: UnicodeString; const Data: DWORD);
  begin
    RegistryResult := RegSetValueExW(RegistryKey, PWideChar(Name), 0, REG_DWORD, PByte(@Data), SizeOf(Data));
    if RegistryResult <> ERROR_SUCCESS then
    begin
      SetLastError(RegistryResult);
      RaiseLastOSError;
    end;
  end;
begin
  KeyName := UnicodeString(UNINSTALL_KEY);
  RegistryResult := RegCreateKeyExW(TPaths.GetRegistryRootKey(TPaths.Scope), PWideChar(KeyName), 0, nil, REG_OPTION_NON_VOLATILE, KEY_WRITE, nil, RegistryKey, @Disposition);
  if RegistryResult <> ERROR_SUCCESS then
  begin
    SetLastError(RegistryResult);
    RaiseLastOSError;
  end;
  try
    WriteStringValue('DisplayIcon', Executable);
    WriteStringValue('DisplayName', UnicodeString(APPNAME));
    WriteStringValue('DisplayVersion', TFunctions.GetFileVersion(TPaths.ExePath));
    WriteStringValue('InstallDate', UnicodeString(FormatDateTime('yyyyMMdd', Now)));
    WriteStringValue('InstallLocation', SysUtils.ExtractFileDir(Executable));
    WriteIntegerValue('NoModify', 1);
    WriteIntegerValue('NoRepair', 1);
    Value := 'rundll32.exe "' + Executable + '" Uninstall';
    WriteStringValue('UninstallString', Value);
    WriteStringValue('Publisher', 'bmmlms');
  finally
    RegCloseKey(RegistryKey);
  end;
end;

procedure Install;
var
  UninstallLibraryPath: UnicodeString;
  Lib: HMODULE;
  DllRegisterServer: function: HRESULT; stdcall;
  ClassesKey: HKEY;
  Res: HRESULT;
begin
  if not DirectoryExistsUTF8(UTF16ToUTF8(TPaths.WhatTheLockDir)) then
    if not UnicodeCreateDirectories(TPaths.WhatTheLockDir) then
      raise Exception.Create('Error creating installation directory');

  if FileExistsUTF8(UTF16ToUTF8(TPaths.WhatTheLockLibraryPath)) and (not DeleteFileUTF8(UTF16ToUTF8(TPaths.WhatTheLockLibraryPath))) then
  begin
    UninstallLibraryPath := UnicodeConcatPath(SysUtils.ExtractFileDir(TPaths.WhatTheLockLibraryPath), UnicodeString(Guid.NewGuid.ToString(True)));
    RenameFileUTF8(UTF16ToUTF8(TPaths.WhatTheLockLibraryPath), UTF16ToUTF8(UninstallLibraryPath));
    TFunctions.ScheduleDelete(UninstallLibraryPath, '');
  end;

  if not ExtractResource('LIB_64', TPaths.WhatTheLockLibraryPath) then
    raise Exception.Create('Error installing library');

  Lib := LoadLibraryW(PWideChar(TPaths.WhatTheLockLibraryPath));
  if Lib = 0 then
    raise Exception.Create('Error loading library');

  @DllRegisterServer := GetProcAddress(Lib, 'DllRegisterServer');
  if not Assigned(@DllRegisterServer) then
    raise Exception.Create('Error loading library');

  if TPaths.Scope = scCurrentUser then
  begin
    // Redirect all writes to HKEY_CLASSES_ROOT (COM server, type library, shell extension) to HKEY_CURRENT_USER\Software\Classes
    ClassesKey := TFunctions.BeginPerUserRegistration;
    try
      Res := DllRegisterServer;
    finally
      TFunctions.EndPerUserRegistration(ClassesKey);
    end;
  end else
    Res := DllRegisterServer;

  if Failed(Res) then
    raise Exception.Create('Error registering library');

  CreateUninstallEntry(TPaths.WhatTheLockLibraryPath);

  TFunctions.MessageBox(0, UnicodeString('Installation completed successfully.'), UnicodeString('Information'), MB_ICONINFORMATION);
end;

var
  Scope, DefaultScope, OtherScope: TInstallScope;
  ProcessExitCode: DWORD;
begin
  try
    if not IsTaskDialogAvailable then
      raise Exception.Create('The required Windows Task Dialog API is not available. Setup cannot continue.');

    TPaths.Initialize;

    if (ParamCount >= 1) and SameText(ParamStr(1), PARAM_ALLUSERS) then
    begin
      // Elevated instance, the user already made the choice
      Scope := scAllUsers;

      if not IsUserAnAdmin then
        raise Exception.Create('Administrator rights are required to install for all users');
    end else
    begin
      if TFunctions.IsInstalled(scCurrentUser) and not TFunctions.IsInstalled(scAllUsers) then
        DefaultScope := scCurrentUser
      else
        DefaultScope := scAllUsers;

      if not AskScope(DefaultScope, Scope) then
        Exit;

      if Scope = scAllUsers then
        OtherScope := scCurrentUser
      else
        OtherScope := scAllUsers;

      if TFunctions.IsInstalled(OtherScope) then
        if OtherScope = scAllUsers then
          raise Exception.Create('%s is already installed for all users. Please uninstall it first'.Format([APPNAME]))
        else
          raise Exception.Create('%s is already installed for the current user. Please uninstall it first'.Format([APPNAME]));

      if (Scope = scAllUsers) and (not IsUserAnAdmin) then
      begin
        if not TFunctions.RunElevated(TPaths.ExePath, UnicodeString(PARAM_ALLUSERS), ProcessExitCode) then
          Exit;

        ExitProcess(ProcessExitCode);
      end;
    end;

    TPaths.SetScope(Scope);

    try
      Install;
    except
      try
        TFunctions.RunUninstall;
      except
      end;
      raise;
    end;

    ExitProcess(0);
  except
    on E: Exception do
    begin
      if not E.Message.EndsWith('.') then
        E.Message := E.Message + '.';

      TFunctions.MessageBox(0, 'Error installing ' + UnicodeString(APPNAME) + ': ' + UnicodeString(E.Message), UnicodeString('Error'), MB_ICONERROR);

      ExitProcess(1);
    end;
  end;
end.
