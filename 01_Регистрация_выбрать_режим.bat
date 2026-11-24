@echo off
chcp 65001 > nul
REM Select a registration scenario from settings.json. Edge runs headless.

cd /d "%~dp0"
"%~dp0asud.exe" --headless

pause
