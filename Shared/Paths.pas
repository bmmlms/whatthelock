unit Paths;

interface

uses
  ShlObj,
  SysUtils;

type
  TPaths = class
  private
  class var
    FExePath: string;
    FWhatTheLockDir: string;
    FWhatTheLockLibraryPath: string;
  public
    class procedure Initialize; static;
    class property ExePath: string read FExePath;
    class property WhatTheLockDir: string read FWhatTheLockDir;
    class property WhatTheLockLibraryPath: string read FWhatTheLockLibraryPath;
  end;

implementation

uses
  Constants,
  Functions;

class procedure TPaths.Initialize;
begin
  FExePath := ParamStr(0);
  FWhatTheLockDir := ConcatPaths([TFunctions.GetSpecialFolder(CSIDL_PROGRAM_FILES), APPNAME]);
  FWhatTheLockLibraryPath := ConcatPaths([FWhatTheLockDir, LIBRARYNAME_64]);
end;

end.
