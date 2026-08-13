@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

rem Double-click: use the SBIS preset from settings.json.
rem Drag a folder onto this BAT: use that folder for this run only.

if "%~1"=="" goto run_preset
"%~dp0asud.exe" --headless --mode=sbis --folder "%~1"
goto done

:run_preset
"%~dp0asud.exe" --headless --preset=sbis

:done
echo.
pause
endlocal
