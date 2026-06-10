@echo off
chcp 65001 > nul
REM Непрерывный мониторинг (daemon) в headless-режиме.
REM Двойной клик → меню пресетов (ОЭК/ТЭС) → daemon-loop.
REM Ctrl+C для остановки.

cd /d "%~dp0"
"%~dp0asud.exe" --headless --watch

pause
