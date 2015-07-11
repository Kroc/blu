@echo off
CD /d %~dp0

::---

SET "WINBIT=32"
IF /I [%PROCESSOR_ARCHITECTURE%]==[x86]	  SET "WINBIT=32"
IF /I [%PROCESSOR_ARCHITECTURE%]==[EM64T] SET "WINBIT=64"
IF /I [%PROCESSOR_ARCHITECTURE%]==[AMD64] SET "WINBIT=64"

REM Set the program path based on the bit version.
IF [%WINBIT%]==[64] (
	SET "PROGRAMS=%PROGRAMFILES(x86)%"
) ELSE (
	SET "PROGRAMS=%PROGRAMFILES%"
)

::---

CD "%PROGRAMS%\Microsoft Visual Studio\VC98\Bin"
CALL VCVARS32.BAT

CD /d %~dp0

"%PROGRAMS%\Microsoft Visual Studio\VC98\Bin\MIDL.exe" WIN32.idl

PAUSE