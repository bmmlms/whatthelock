unit Paths;

interface

uses
  Constants,
  ShlObj,
  Windows,
  SysUtils;

type
  TPaths = class
  private
  class var
    FExePath: string;
    FTempDir: string; // TODO: relevant?
    FWhatTheLockDir: string;
    FWhatTheLockLibraryPath: string;
  public
    class procedure Init; static;
    class property ExePath: string read FExePath;
    class property TempDir: string read FTempDir;
    class property WhatTheLockDir: string read FWhatTheLockDir;
    class property WhatTheLockLibraryPath: string read FWhatTheLockLibraryPath;
  end;

implementation

uses
  Functions;

class procedure TPaths.Init;
begin
  FExePath := ParamStr(0);
  FTempDir := TFunctions.GetTempPath;
  FWhatTheLockDir := ConcatPaths([TFunctions.GetSpecialFolder(CSIDL_PROGRAM_FILES), APPNAME]);
  FWhatTheLockLibraryPath := ConcatPaths([FWhatTheLockDir, LIBRARYNAME_64]);
end;

end.
