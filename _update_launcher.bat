@echo off
setlocal
set "PID=4052"
set "NEW_EXE=C:\Users\QC60\AppData\Local\Temp\VIPQC_AI_Update_2.2.1.exe"
set "TARGET_EXE=E:\CODE\compare MATERIAL\VUNG Folder\VUNG Folder\BOM_Extractor.exe"

:: Wait up to 10 seconds for process to exit
set count=0
:wait_loop
timeout /t 1 /nobreak > nul
set /a count+=1
tasklist /fi "pid eq %PID%" 2>nul | findstr "%PID%" > nul
if not errorlevel 1 (
    if %count% leq 10 goto wait_loop
)

:: Copy new executable over current executable
copy /y "%NEW_EXE%" "%TARGET_EXE%" > nul
if exist "%NEW_EXE%" del /f /q "%NEW_EXE%" > nul

:: Also sync BOM_Extractor.exe if it exists in same folder
set "ALT_EXE=E:\CODE\compare MATERIAL\VUNG Folder\VUNG Folder\BOM_Extractor.exe"
if exist "%ALT_EXE%" (
    copy /y "%TARGET_EXE%" "%ALT_EXE%" > nul
)

:: Launch updated application
start "" "%TARGET_EXE%"

:: Self delete this batch script
(goto) 2>nul & del /f /q "%~f0"
