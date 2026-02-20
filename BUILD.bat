@echo off
chcp 437 > nul
:: ============================================================
::  BUILD.bat — Compile CombinedInstaller.ps1 into EXE
::  Run as Administrator
:: ============================================================
title Combined Installer - Builder

:: Переходим в папку где лежит BUILD.bat — это ключевое исправление
cd /d "%~dp0"

echo.
echo  +------------------------------------------+
echo  ^|      Combined Installer  -  Builder      ^|
echo  +------------------------------------------+
echo  Working directory: %CD%
echo.

:: 1. Check / install ps2exe
echo [1/3] Checking ps2exe...
powershell -Command "if (-not (Get-Module -ListAvailable -Name ps2exe)) { Write-Host 'Installing ps2exe...' -ForegroundColor Yellow; Install-Module ps2exe -Scope CurrentUser -Force -Confirm:$false } else { Write-Host 'ps2exe already installed.' -ForegroundColor Green }"

:: 2. Compile — используем абсолютные пути через %~dp0
echo.
echo [2/3] Compiling to EXE...
powershell -Command "Invoke-ps2exe -InputFile '%~dp0src/CombinedInstaller.ps1' -OutputFile '%~dp0CombinedInstaller.exe' -RequireAdmin -NoConsole:$false -Title 'Combined Installer' -Description 'Corporate software installer' -Company 'Your Company' -Version '1.0.0' -Verbose"

:: 3. Check result
echo.
echo [3/3] Result:
if exist "%~dp0CombinedInstaller.exe" (
    echo  [OK] CombinedInstaller.exe built successfully!
    for %%A in ("%~dp0CombinedInstaller.exe") do echo  Size: %%~zA bytes
    echo  Location: %~dp0CombinedInstaller.exe
) else (
    echo  [FAILED] Build error. Check the output above.
)

echo.
pause