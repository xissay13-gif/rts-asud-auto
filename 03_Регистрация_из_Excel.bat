@echo off
chcp 65001 > nul
REM Bulk registration from an XLSX registry without MSG attachments.

cd /d "%~dp0"
"%~dp0asud.exe" --headless --mode=auto-create

pause
