library WhatTheLock;

uses
  App,
  Classes,
  ComServ,
  Constants,
  FileHandleUtil,
  Functions,
  LibraryInit,
  Paths,
  RtlConsts,
  ShellExt,
  ShlObj,
  SysUtils,
  UnicodePaths,
  windows;

  {$R *.tlb}
  {$R *.res}
  {$R ..\Shared\CommonControls.rc}

procedure Uninstall; stdcall;
var
  ClassesKey: HKEY;
  ProcessExitCode: DWORD;
begin
  // Uninstalling has to work even if parts of the initialization failed
  EnsureLibraryInitialized(False);

  try
    if (TPaths.Scope = scAllUsers) and (not IsUserAnAdmin) then
    begin
      TFunctions.RunElevated(UnicodeConcatPath(TFunctions.GetSpecialFolder(CSIDL_SYSTEM), 'rundll32.exe'), UnicodeString('"') + TPaths.WhatTheLockLibraryPath + '" Uninstall', ProcessExitCode);
      Exit;
    end;

    if TPaths.Scope = scCurrentUser then
    begin
      ClassesKey := TFunctions.BeginPerUserRegistration;
      try
        DllUnregisterServer;
      finally
        TFunctions.EndPerUserRegistration(ClassesKey);
      end;
    end else
      DllUnregisterServer;

    TFunctions.RunUninstall;

    TFunctions.MessageBox(0, UnicodeString(APPNAME) + ' was uninstalled successfully.', UnicodeString(SMsgDlgInformation), MB_ICONINFORMATION);
  except
    on E: Exception do
    begin
      if not E.Message.EndsWith('.') then
        E.Message := E.Message + '.';

      TFunctions.MessageBox(0, 'Error uninstalling ' + UnicodeString(APPNAME) + ': ' + UnicodeString(E.Message), UnicodeString(SMsgDlgError), MB_ICONERROR);
    end;
  end;
end;

// Started by the context menu via "rundll32.exe <library>,Check <ExcludePID> <ParentWindow> "<FileName>"".
// rundll32 prefers the export with suffix W and passes the command line as UTF-16 then.
procedure CheckW(Wnd: HWND; Instance: HINST; CmdLine: PWideChar; CmdShow: Integer); stdcall;
begin
  if not EnsureLibraryInitialized(False) then
    Exit;

  try
    RunCheck(UnicodeString(CmdLine));
  except
    on E: Exception do
      TFunctions.HandleException(0, E);
  end;
end;

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  Uninstall,
  CheckW;

begin
  // This block runs in DllMain while the loader lock is held, the real initialization happens in
  // EnsureLibraryInitialized() on first use (TWhatTheLockMenu.Initialize, Uninstall, CheckW).
  IsMultiThread := True;
end.
