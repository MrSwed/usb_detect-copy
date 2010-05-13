@echo off

deveject >nul || (
 echo need deveject. See http://withopf.com/tools/deveject or http://www.google.com/search?q=deveject
 echo 
 pause
 exit /b
)
xcopy /? >nul || (
 echo need xcopy. http://www.microsoft.com/resources/documentation/windows/xp/all/proddocs/en-us/xcopy.mspx
 echo 
 pause
 exit /b
)

set detector=detect.vbs
set start_copy=start_copy.bat
set sources=sources.txt

pushd "%~dp0"
if not exist "%detector%"  echo "need %detector% for work" &  goto quit
if not exist "%start_copy%"  echo "need %start_copy% for work ^(copy and disconnect^)" &  goto quit

for %%a in ("%sources%") do (
 if "%%~za"=="" (
  set sources_size=0
 ) else ( 
  set sources_size=%%~za
 )
)

if NOT %sources_size% GTR 2 (
 echo please fill sources file ^(set sources for copy to USB Flash - one line - one source^[^|destination^]^)
 echo Examples:
 echo    d:\temp\data\*.*
 echo  all content of d:\temp\data will be copied to root of flash
 echo    d:\temp\data1^|\data\data1
 echo    d:\temp\data2^|\data2
 echo  for copied content of d:\temp\data1 and d:\temp\data2 will be created dirs \data\data1 and \data2
 echo.>"%sources%"
 start "" notepad "%sources%" >nul
 goto quit
)
chcp 1251>nul
cscript //nologo "%detector%"
chcp 866>nul
goto quit


:quit
popd
