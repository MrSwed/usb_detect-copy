@echo off

SETLOCAL ENABLEDELAYEDEXPANSION

if "%~2"=="" (
 echo For start detect USB and copy data run file _usb_copy
 echo wrong parameters! It script used in "%detector%" as %~nx0 ^<disk_name^> ^<DeviceId^>
 GOTO :EOF
)
set errorcopy=0
set sources=sources.txt


for /f "eol=; usebackq tokens=1-2 delims=|" %%a in ("%sources%") do (
 if not "%%~a"=="" (
  echo copy "%%~a" "%~1\%%~b"
  xcopy /y /e /c /r /i /f /g /h /k "%%~a" "%~1\%%~b" 2>&1 || set errorcopy=1
  echo.
 )

)

set ejectTry=0

:eject
set /a ejectTry+=1
timeout 3 >nul 2>&1
if "!errorcopy!"=="0" (
 echo %ejectTry%. Attempt to eject.... 
 deveject -EjectId:"%~2" >nul 2>&1 && echo %~1 ejected  || (
  if %ejectTry% LEQ 10 goto eject
  echo Error eject %~1 
 )
) else (
 echo Some error was while copy 
)
