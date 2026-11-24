@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

REM Double-click: read the SBIS folder from settings.json.
REM Drop another folder onto this BAT for a one-time run.

if "%~1"=="" goto run_preset
"%~dp0asud.exe" --headless --mode=sbis --folder "%~1"
goto done

:run_preset
"%~dp0asud.exe" --headless --preset=sbis

:done
echo.
pause
endlocal
