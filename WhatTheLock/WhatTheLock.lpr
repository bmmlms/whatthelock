library WhatTheLock;

uses
  App,
  Classes,
  ComServ,
  Constants,
  FileHandleUtil,
  Functions,
  gettext,
  Paths,
  RtlConsts,
  ShellExt,
  ShlObj,
  SysUtils,
  Translations,
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
  try
    if (TPaths.Scope = scAllUsers) and (not IsUserAnAdmin) then
    begin
      // Installations for all users can only be removed with administrator rights
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

procedure Translate;
var
  Lang, FallbackLang: string;
  Res: TResourceStream;
  PoStringStream: TStringStream;
  PoFile: TPOFile;
begin
  GetLanguageIDs(Lang, FallbackLang);

  try
    Res := TResourceStream.Create(HInstance, 'WhatTheLock.' + LowerCase(Lang), RT_RCDATA);
  except
    on E: EResNotFound do
      Exit
    else
      raise Exception.Create('Error loading translations');
  end;

  PoStringStream := TStringStream.Create;
  Res.SaveToStream(PoStringStream);
  Res.Free;

  PoFile := TPOFile.Create(False);
  PoFile.ReadPOText(PoStringStream.DataString);
  PoStringStream.Free;

  TranslateResourceStrings(PoFile);
  PoFile.Free;
end;

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  Uninstall;

begin
  IsMultiThread := True;

  try
    Translate;

    TPaths.Initialize;
    TPaths.SetScopeFromLibraryPath(TPaths.ExePath);
    TFileHandleUtil.Initialize;
    TApp.Initialize;
  except
    on E: Exception do
    begin
      if not E.Message.EndsWith('.') then
        E.Message := E.Message + '.';

      TFunctions.MessageBox(0, UnicodeString(E.Message), UnicodeString(APPNAME) + ' error', MB_ICONERROR);
    end;
  end;
end.
