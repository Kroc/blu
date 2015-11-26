@ECHO off
REM Change the directory to that of this script,
REM regardless of where it was called from
CD /d %~dp0

::Get the "Program Files" folder
::------------------------------

REM Assume a 32-bit system and look for 64-bit CPUs
SET "WINBIT=32"
IF /I [%PROCESSOR_ARCHITECTURE%]==[EM64T] SET "WINBIT=64"   REM Intel's x86_64
IF /I [%PROCESSOR_ARCHITECTURE%]==[AMD64] SET "WINBIT=64"   REM AMD's x86_64

REM Set the "Program Files" path based on the CPU bit-depth
IF [%WINBIT%]==[64] (
    REM On a 64-bit system, VS6 will be installed in the legacy folder
	SET "PROGRAMS=%PROGRAMFILES(x86)%"
) ELSE (
	SET "PROGRAMS=%PROGRAMFILES%"
)


::---

CD "%PROGRAMS%\Microsoft Visual Studio\VC98\Bin"
REM Import the Visual C++ environment
CALL VCVARS32.BAT
ECHO.

::---

CD /d %~dp0

"%PROGRAMS%\Microsoft Visual Studio\VC98\Bin\MIDL.exe" bluW32.idl

PAUSE