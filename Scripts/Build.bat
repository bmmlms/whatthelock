@ECHO OFF

REM Pathes of required tools
SET "FPCBIN=D:\Lazarus\fpc\3.2.2\bin\x86_64-win64"
SET "ZIP=D:\7-Zip\7z.exe"
SET "PLINK=C:\Program Files\PuTTY\plink.exe"
SET "MSYS2=D:\msys64\msys2_shell.cmd"

IF NOT EXIST %FPCBIN% GOTO FAILENVIRONMENT

REM Extend PATH
SET "PATH=%PATH%;%FPCBIN%;%FPCBIN%\..\..\..\.."

if exist ..\Build\ (
  rmdir /s /q ..\Build
)

REM Build libraries
call "%MSYS2%" -defterm -no-start -where "%cd%\..\SubModules\minlzma" -mingw64 -c "sed -i 's/\-Wconversion //g' minlzlib/CMakeLists.txt minlzdec/CMakeLists.txt && git checkout minlzlib/xzstream.h && echo '#define _In_' >> minlzlib/xzstream.h && mkdir -p build && cd build && cmake .. && make -j && exit"
if %ERRORLEVEL% GEQ 1 exit /B %ERRORLEVEL%

REM Build dlls
cd ..\WhatTheLock
lazbuild --build-all --cpu=x86_64 --os=Win64 --build-mode=Release WhatTheLock.lpi
IF ERRORLEVEL 1 GOTO FAIL

type ..\Build\Release\x86_64\WhatTheLock.dll | "%PLINK%" -batch gaia osslsigncode-sign.sh > ..\Build\Release\x86_64\WhatTheLock-signed.dll
IF ERRORLEVEL 1 GOTO FAIL
move /y ..\Build\Release\x86_64\WhatTheLock-signed.dll ..\Build\Release\x86_64\WhatTheLock.dll
IF ERRORLEVEL 1 GOTO FAIL

REM Compress setup resources
cd ..
mkdir Build\SetupResources
"%ZIP%" a dummy -txz -mx=9 -so Build\Release\x86_64\WhatTheLock.dll > Build\SetupResources\WhatTheLock-x86_64.dll.xz
IF ERRORLEVEL 1 GOTO FAIL

REM Build setup
cd WhatTheLock_Setup
lazbuild --build-all --cpu=x86_64 --os=Win64 --build-mode=Release WhatTheLock_Setup.lpi
IF ERRORLEVEL 1 GOTO FAIL

type ..\Build\Release\x86_64\WhatTheLock_Setup.exe | "%PLINK%" -batch gaia osslsigncode-sign.sh > ..\Build\Release\x86_64\WhatTheLock_Setup-signed.exe
IF ERRORLEVEL 1 GOTO FAIL
move /y ..\Build\Release\x86_64\WhatTheLock_Setup-signed.exe ..\Build\Release\x86_64\WhatTheLock_Setup.exe
IF ERRORLEVEL 1 GOTO FAIL

ECHO.
ECHO Build finished
ECHO.
GOTO END

:FAILENVIRONMENT
  ECHO.
  ECHO FPCBIN does not exist, please adjust variable
  ECHO.
  PAUSE
  GOTO END

:FAIL
  ECHO.
  ECHO Build failed
  ECHO.
  PAUSE

:END
