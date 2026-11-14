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
    FExePath: UnicodeString;
    FAllUsersDir: UnicodeString;
    FCurrentUserDir: UnicodeString;
    FScope: TInstallScope;
    FWhatTheLockDir: UnicodeString;
    FWhatTheLockLibraryPath: UnicodeString;
  public
    class procedure Initialize; static;
    class procedure SetScope(const Scope: TInstallScope); static;
    class procedure SetScopeFromLibraryPath(const LibraryPath: UnicodeString); static;

    class function GetInstallDir(const Scope: TInstallScope): UnicodeString; static;
    class function GetRegistryRootKey(const Scope: TInstallScope): HKEY; static;
    class property ExePath: UnicodeString read FExePath;
    class property Scope: TInstallScope read FScope;
    class property WhatTheLockDir: UnicodeString read FWhatTheLockDir;
    class property WhatTheLockLibraryPath: UnicodeString read FWhatTheLockLibraryPath;
  end;

implementation

uses
  Constants,
  Functions,
  UnicodePaths;

class procedure TPaths.Initialize;
var
  BufferLength, CharCount: DWORD;
begin
  BufferLength := MAX_PATH;
  repeat
    SetLength(FExePath, BufferLength);
    CharCount := GetModuleFileNameW(HInstance, PWideChar(FExePath), BufferLength);
    if CharCount = 0 then
      RaiseLastOSError;
    if CharCount < BufferLength then
      Break;
    BufferLength := BufferLength * 2;
  until BufferLength > 32768;
  SetLength(FExePath, CharCount);

  FAllUsersDir := UnicodeConcatPath(TFunctions.GetSpecialFolder(CSIDL_PROGRAM_FILES), UnicodeString(APPNAME));
  FCurrentUserDir := UnicodeConcatPath(UnicodeConcatPath(TFunctions.GetSpecialFolder(CSIDL_LOCAL_APPDATA), 'Programs'), UnicodeString(APPNAME));

  SetScope(scAllUsers);
end;

class procedure TPaths.SetScope(const Scope: TInstallScope);
begin
  FScope := Scope;
  FWhatTheLockDir := GetInstallDir(Scope);
  FWhatTheLockLibraryPath := UnicodeConcatPath(FWhatTheLockDir, UnicodeString(LIBRARYNAME_64));
end;

class procedure TPaths.SetScopeFromLibraryPath(const LibraryPath: UnicodeString);
begin
  if UnicodeCompareTextOrdinal(SysUtils.ExtractFileDir(LibraryPath), FCurrentUserDir) then
    SetScope(scCurrentUser)
  else
    SetScope(scAllUsers);
end;

class function TPaths.GetInstallDir(const Scope: TInstallScope): UnicodeString;
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
