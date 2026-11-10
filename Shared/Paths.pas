unit Paths;

interface

uses
  ShlObj,
  SysUtils,
  Windows;

type
  TInstallScope = (scAllUsers, scCurrentUser);

  TPaths = class
  private
  class var
    FExePath: string;
    FAllUsersDir: string;
    FCurrentUserDir: string;
    FScope: TInstallScope;
    FWhatTheLockDir: string;
    FWhatTheLockLibraryPath: string;
  public
    class procedure Initialize; static;
    class procedure SetScope(const Scope: TInstallScope); static;
    class procedure SetScopeFromLibraryPath(const LibraryPath: string); static;

    class function GetInstallDir(const Scope: TInstallScope): string; static;
    class function GetRegistryRootKey(const Scope: TInstallScope): HKEY; static;
    class property ExePath: string read FExePath;
    class property Scope: TInstallScope read FScope;
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
  FAllUsersDir := ConcatPaths([TFunctions.GetSpecialFolder(CSIDL_PROGRAM_FILES), APPNAME]);
  FCurrentUserDir := ConcatPaths([TFunctions.GetSpecialFolder(CSIDL_LOCAL_APPDATA), 'Programs', APPNAME]);

  SetScope(scAllUsers);
end;

class procedure TPaths.SetScope(const Scope: TInstallScope);
begin
  FScope := Scope;
  FWhatTheLockDir := GetInstallDir(Scope);
  FWhatTheLockLibraryPath := ConcatPaths([FWhatTheLockDir, LIBRARYNAME_64]);
end;

class procedure TPaths.SetScopeFromLibraryPath(const LibraryPath: string);
begin
  if SameText(ExcludeTrailingPathDelimiter(ExtractFileDir(LibraryPath)), ExcludeTrailingPathDelimiter(FCurrentUserDir)) then
    SetScope(scCurrentUser)
  else
    SetScope(scAllUsers);
end;

class function TPaths.GetInstallDir(const Scope: TInstallScope): string;
begin
  if Scope = scCurrentUser then
    Result := FCurrentUserDir
  else
    Result := FAllUsersDir;
end;

class function TPaths.GetRegistryRootKey(const Scope: TInstallScope): HKEY;
begin
  if Scope = scCurrentUser then
    Result := HKEY_CURRENT_USER
  else
    Result := HKEY_LOCAL_MACHINE;
end;

end.
