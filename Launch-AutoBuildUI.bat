@echo off
REM AutoBuild Automation Interface Launcher
REM Double-click to open the GUI from the AutoBuild root directory.

SET "SCRIPT_DIR=%~dp0"
SET "UI_SCRIPT=%SCRIPT_DIR%AutoBuild.UI.ps1"

IF NOT EXIST "%UI_SCRIPT%" (
    echo ERROR: AutoBuild.UI.ps1 not found in %SCRIPT_DIR%
    pause
    exit /b 1
)

REM Launch WPF GUI (requires STA thread mode for WPF)
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass ^
    -WindowStyle Hidden ^
    -Command "& { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; . '%UI_SCRIPT%' -EnginePath '%SCRIPT_DIR%' }"

IF %ERRORLEVEL% NEQ 0 (
    echo AutoBuild UI exited with code %ERRORLEVEL%
    pause
)
