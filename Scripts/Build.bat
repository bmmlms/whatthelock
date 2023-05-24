@ECHO OFF

REM Pathes of required tools
SET "FPCBIN=D:\Lazarus\fpc\3.2.2\bin\x86_64-win64"
SET "ZIP=D:\7-Zip\7z.exe"
SET "PLINK=C:\Program Files\PuTTY\plink.exe"
SET "MODE=Release"

IF NOT EXIST %FPCBIN% GOTO FAILENVIRONMENT

REM Extend PATH
SET "PATH=%PATH%;%FPCBIN%;%FPCBIN%\..\..\..\.."

if exist ..\Build\ (
  rmdir /s /q ..\Build
)

REM Build dlls
cd ..\WhatTheLock
lazbuild --build-all --cpu=x86_64 --os=Win64 --build-mode=%MODE% WhatTheLock.lpi
IF ERRORLEVEL 1 GOTO FAIL

type ..\Build\%MODE%\x86_64\WhatTheLock.dll | "%PLINK%" -batch gaia osslsigncode-sign.sh > ..\Build\%MODE%\x86_64\WhatTheLock-signed.dll
IF ERRORLEVEL 1 GOTO FAIL
move /y ..\Build\%MODE%\x86_64\WhatTheLock-signed.dll ..\Build\%MODE%\x86_64\WhatTheLock.dll
IF ERRORLEVEL 1 GOTO FAIL

REM Build setup
cd ..\WhatTheLock_Setup
lazbuild --build-all --cpu=x86_64 --os=Win64 --build-mode=%MODE% WhatTheLock_Setup.lpi
IF ERRORLEVEL 1 GOTO FAIL

type ..\Build\%MODE%\x86_64\WhatTheLock_Setup.exe | "%PLINK%" -batch gaia osslsigncode-sign.sh > ..\Build\%MODE%\x86_64\WhatTheLock_Setup-signed.exe
IF ERRORLEVEL 1 GOTO FAIL
move /y ..\Build\%MODE%\x86_64\WhatTheLock_Setup-signed.exe ..\Build\%MODE%\x86_64\WhatTheLock_Setup.exe
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
