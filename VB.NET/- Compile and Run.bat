@echo off
:: Set Variables:
set EXENAME=SyMenuOrphanIcons.exe
set DESTDIR=D:\SyMenu\ProgramFiles\SPSSuite\SyMenuSuite\SyMenu_Orphan_Icons_Cleaner_sps

:: Set the working directory:
cd /d "%~dp0"

:: Close any running instances:
echo Closing any running instances...
taskkill /IM %EXENAME% /F 2>nul

:: Compile the project:
echo Compiling...
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe SyMenuOrphanIcons.vbproj /p:Configuration=Release
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Compilation successful! Copying exe...
    copy /Y "bin\Release\%EXENAME%" "%DESTDIR%\%EXENAME%"
    echo Starting app...
    timeout /t 2
    start "" "%DESTDIR%\%EXENAME%"
    exit
) else (
    echo.
    echo Compilation failed!
    pause
)