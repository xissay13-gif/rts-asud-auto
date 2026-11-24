@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

if "%~1"=="" goto configured
"%~dp0asud_reports.exe" --headless --folder "%~1"
goto done

:configured
"%~dp0asud_reports.exe" --headless

:done
echo.
pause
endlocal
