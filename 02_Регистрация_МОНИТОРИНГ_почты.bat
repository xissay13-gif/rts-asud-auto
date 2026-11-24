@echo off
chcp 65001 > nul
REM Continuously register new messages. Select a preset after startup.
REM Ctrl+C stops the monitor after the current document.

cd /d "%~dp0"
"%~dp0asud.exe" --headless --watch

pause
