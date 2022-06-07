@echo off

if %1==grey (
	set color=2
) else (
	set color=1
)

echo Scan color id: %color%

set configPath=%~dp0config.bat
call "%configPath%"

set counterFilePath=%savePath%counter.txt
echo Counter file: %counterFilePath%

for /f "tokens=1-4 delims=/. " %%a in ('date /t') do (set mydate=%%c-%%b-%%a)
for /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set mytime=%%a%%b)

for /f "usebackq delims=" %%x in ("%counterFilePath%") do set counter=%%x
echo %counter%
set /A counter=counter+1
echo %counter%

set BmpFilePath=%savePath%%counter%_%mydate%_%mytime%.bmp

echo Scanned file path: %BmpFilePath%

"C:/Windows/System32/cscript.exe" //X "%vbsFilePath%" "%BmpFilePath%" %color%

@echo %counter% > "%counterFilePath%"

exit /b 0
