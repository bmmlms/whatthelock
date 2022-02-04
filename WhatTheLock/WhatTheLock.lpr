library WhatTheLock;

uses
  App,
  Classes,
  ComServ,
  Constants,
  FileHandleUtil,
  Functions,
  gettext,
  Interfaces,
  Paths,
  RtlConsts,
  ShellExt,
  SysUtils,
  Translations,
  windows;

{$R *.tlb}
{$R *.res}

procedure Uninstall; stdcall;
begin
  DllUnregisterServer;

  TFunctions.RunUninstall;

  TFunctions.MessageBox(0, '%s was uninstalled successfully.'.Format([APPNAME]), SMsgDlgInformation, MB_ICONINFORMATION);
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

var
  ThreadId: DWORD;
begin
  IsMultiThread := True;

  try
    Translate;

    TPaths.Initialize;
    TFileHandleUtil.Initialize;
    TApp.Initialize;
  except
    on E: Exception do
    begin
      if not E.Message.EndsWith('.') then
        E.Message := E.Message + '.';

      TFunctions.MessageBox(0, E.Message, '%s error'.Format([APPNAME]), MB_ICONERROR);
    end;
  end;
end.
