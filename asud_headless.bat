@echo off
chcp 65001 > nul
REM Запуск asud.exe в headless-режиме (Edge без GUI).
REM Двойной клик = запуск; pause в конце оставит окно для просмотра лога.

cd /d "%~dp0"
"%~dp0asud.exe" --headless

pause
