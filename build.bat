@echo off
set "PATH=C:\Program Files (x86)\Microsoft Visual Studio\Installer;%PATH%"
call "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\Tools\VsDevCmd.bat" -arch=x64 >nul 2>&1
echo Building CSV2ICS...
rc /nologo csv2ics.rc
cl /std:c17 /O2 /W4 /DUNICODE /D_UNICODE csv2ics.c csv2ics.res /Fe:csv2ics.exe /link user32.lib gdi32.lib comdlg32.lib comctl32.lib shell32.lib ole32.lib advapi32.lib
if errorlevel 1 (
    echo BUILD FAILED
    exit /b 1
) else (
    echo BUILD SUCCEEDED: csv2ics.exe
    del csv2ics.obj csv2ics.res 2>nul
)
