program WhatTheLock_Setup;

uses
  ActiveX,
  Classes,
  Constants,
  Functions,
  Paths,
  Registry,
  shlobj,
  SysUtils,
  TaskDialog,
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

  Title := UnicodeString('%s Setup'.Format([APPNAME]));
  Instruction := UnicodeString('This will install/update %s.'.Format([APPNAME]));
  Content := 'Who should be able to use it?';
  AllUsersText := UnicodeString('For all users'#10'Installs to "%s". Requires administrator rights.'.Format([TPaths.GetInstallDir(scAllUsers)]));
  CurrentUserText := UnicodeString('Only for me'#10'Installs to "%s". No administrator rights required.'.Format([TPaths.GetInstallDir(scCurrentUser)]));

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

function ExtractResource(const ResourceName, FilePath: string): Boolean;
var
  ResStream: TResourceStream;
begin
  Result := False;
  try
    ResStream := TResourceStream.Create(HInstance, ResourceName, RT_RCDATA);
    try
      ResStream.SaveToFile(FilePath);
      Result := True;
    finally
      ResStream.Free;
    end;
  except
  end;
end;

procedure CreateUninstallEntry(const Executable: string);
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;

  try
    Reg.RootKey := TPaths.GetRegistryRootKey(TPaths.Scope);
    if not Reg.OpenKey(UNINSTALL_KEY, True) then
      Exit;

    Reg.WriteString('DisplayIcon', Executable);
    Reg.WriteString('DisplayName', APPNAME);
    Reg.WriteString('DisplayVersion', TFunctions.GetFileVersion(TPaths.ExePath));
    Reg.WriteString('InstallDate', FormatDateTime('yyyyMMdd', Now));
    Reg.WriteString('InstallLocation', ExtractFileDir(Executable));
    Reg.WriteInteger('NoModify', 1);
    Reg.WriteInteger('NoRepair', 1);
    Reg.WriteString('UninstallString', 'rundll32.exe "%s" Uninstall'.Format([Executable]));
    Reg.WriteString('Publisher', 'bmmlms');
  finally
    Reg.Free;
  end;
end;

procedure Install;
var
  UninstallLibraryPath: string;
  Lib: HMODULE;
  DllRegisterServer: function: HRESULT; stdcall;
  ClassesKey: HKEY;
  Res: HRESULT;
begin
  if not DirectoryExists(TPaths.WhatTheLockDir) then
    if not ForceDirectories(TPaths.WhatTheLockDir) then
      raise Exception.Create('Error creating installation directory');

  if FileExists(TPaths.WhatTheLockLibraryPath) and (not SysUtils.DeleteFile(TPaths.WhatTheLockLibraryPath)) then
  begin
    UninstallLibraryPath := ConcatPaths([ExtractFileDir(TPaths.WhatTheLockLibraryPath), Guid.NewGuid.ToString(True)]);
    SysUtils.RenameFile(TPaths.WhatTheLockLibraryPath, UninstallLibraryPath);
    TFunctions.ScheduleDelete(UninstallLibraryPath, '');
  end;

  if not ExtractResource('LIB_64', TPaths.WhatTheLockLibraryPath) then
    raise Exception.Create('Error installing library');

  Lib := LoadLibrary(PChar(TPaths.WhatTheLockLibraryPath));
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

  TFunctions.MessageBox(0, 'Installation completed successfully.', 'Information', MB_ICONINFORMATION);
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
        if not TFunctions.RunElevated(TPaths.ExePath, PARAM_ALLUSERS, ProcessExitCode) then
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

      TFunctions.MessageBox(0, 'Error installing %s: %s'.Format([APPNAME, E.Message]), 'Error', MB_ICONERROR);

      ExitProcess(1);
    end;
  end;
end.
