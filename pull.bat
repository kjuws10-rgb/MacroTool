@echo off
setlocal

set "PULL_SCRIPT=%~dp0scripts\pull.ps1"

if not exist "%PULL_SCRIPT%" (
    echo pull failed: scripts\pull.ps1 file was not found.
    exit /b 1
)

where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo pull failed: Windows PowerShell was not found.
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "%PULL_SCRIPT%" %*
set "PULL_EXIT_CODE=%ERRORLEVEL%"

if not "%PULL_EXIT_CODE%"=="0" (
    echo.
    echo pull.bat failed with exit code %PULL_EXIT_CODE%.
)

exit /b %PULL_EXIT_CODE%
