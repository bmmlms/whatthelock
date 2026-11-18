unit LibraryInit;

interface

// Initializes the library on first use. This must not happen in the main block of the library or in unit
// initialization sections: those run in DllMain while the loader lock is held, where loading other libraries, using
// the shell/COM, GDI or installing hooks (which suspends threads) can deadlock the host process.
// ShellExtension: also initialize what is only needed when running as shell extension inside explorer
// (menu bitmap, IFileOperation hook).
// Returns False if the basic initialization failed (an error message has been shown already).
function EnsureLibraryInitialized(const ShellExtension: Boolean): Boolean;

implementation

uses
  App,
  Classes,
  Constants,
  FileHandleUtil,
  Functions,
  gettext,
  Paths,
  ShellExt,
  SysUtils,
  Translations,
  Windows;

var
  InitLock: TRTLCriticalSection;
  BaseInitialized: Boolean = False;
  BaseInitSucceeded: Boolean = False;
  ShellExtensionInitialized: Boolean = False;

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

procedure ShowInitError(const E: Exception);
var
  Msg: string;
begin
  Msg := E.Message;
  if not Msg.EndsWith('.') then
    Msg := Msg + '.';

  TFunctions.MessageBox(0, UnicodeString(Msg), UnicodeString(APPNAME) + ' error', MB_ICONERROR);
end;

function EnsureLibraryInitialized(const ShellExtension: Boolean): Boolean;
begin
  System.EnterCriticalSection(InitLock);
  try
    if not BaseInitialized then
    begin
      // Only try once, otherwise every new context menu would show the error again
      BaseInitialized := True;
      try
        Translate;

        TPaths.Initialize;
        TPaths.SetScopeFromLibraryPath(TPaths.ExePath);

        BaseInitSucceeded := True;
      except
        on E: Exception do
          ShowInitError(E);
      end;
    end;

    if ShellExtension and (not ShellExtensionInitialized) then
    begin
      ShellExtensionInitialized := True;
      try
        TWhatTheLockMenu.CreateMenuBitmap;

        if BaseInitSucceeded then
          TApp.Initialize;
      except
        on E: Exception do
          ShowInitError(E);
      end;
    end;

    Result := BaseInitSucceeded;
  finally
    System.LeaveCriticalSection(InitLock);
  end;
end;

initialization
  // Initializing a critical section is fine in DllMain
  System.InitCriticalSection(InitLock);

finalization
  System.DoneCriticalSection(InitLock);

end.
