@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

REM Read the folder from settings.json or drop another folder onto this BAT.
if "%~1"=="" goto configured
"%~dp0asud.exe" --mode=attach-replies --folder "%~1"
goto done

:configured
"%~dp0asud.exe" --mode=attach-replies

:done
echo.
pause
endlocal
