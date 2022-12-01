program WhatTheLock_Setup;

uses
  ActiveX,
  Classes,
  Constants,
  Functions,
  Paths,
  Registry,
  SysUtils,
  Windows;

{$R *.res}

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
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if not Reg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WhatTheLock', True) then
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
begin
  if not DirectoryExists(TPaths.WhatTheLockDir) then
    if not CreateDir(TPaths.WhatTheLockDir) then
      raise Exception.Create('Error creating installation directory');

  if FileExists(TPaths.WhatTheLockLibraryPath) and (not SysUtils.DeleteFile(TPaths.WhatTheLockLibraryPath)) then
  begin
    UninstallLibraryPath := ConcatPaths([ExtractFileDir(TPaths.WhatTheLockLibraryPath), Guid.NewGuid.ToString(True)]);
    SysUtils.RenameFile(TPaths.WhatTheLockLibraryPath, UninstallLibraryPath);
    MoveFileEx(PChar(UninstallLibraryPath), nil, MOVEFILE_DELAY_UNTIL_REBOOT);
  end;

  if not ExtractResource('LIB_64', TPaths.WhatTheLockLibraryPath) then
    raise Exception.Create('Error installing library');

  Lib := LoadLibrary(PChar(TPaths.WhatTheLockLibraryPath));
  if Lib = 0 then
    raise Exception.Create('Error loading library');

  @DllRegisterServer := GetProcAddress(Lib, 'DllRegisterServer');
  if not Assigned(@DllRegisterServer) then
    raise Exception.Create('Error loading library');

  if Failed(DllRegisterServer) then
    raise Exception.Create('Error registering library');

  CreateUninstallEntry(TPaths.WhatTheLockLibraryPath);

  TFunctions.MessageBox(0, 'Installation completed successfully.', 'Information', MB_ICONINFORMATION);
end;

begin
  try
    TPaths.Init;
    TFunctions.Init;

    if TFunctions.MessageBox(0, 'This will install/update %s.'#13#10'Do you want to continue?'.Format([APPNAME]), 'Question', MB_ICONQUESTION or MB_YESNO) = IDNO then
      Exit;

    try
      Install;
    except
      try
        TFunctions.RunUninstall(True);
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
