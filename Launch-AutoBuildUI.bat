@echo off
REM AutoBuild v3.0 - UI Launcher
REM -STA is required for WPF Window (Single Thread Apartment)
REM /D sets working directory to script location

cd /D "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0ui\AutoBuild.UI.ps1" %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo AutoBuild UI exited with error %ERRORLEVEL%
    pause
)
