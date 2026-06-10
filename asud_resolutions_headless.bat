@echo off
chcp 65001 > nul
REM Запуск asud_resolutions.exe в headless-режиме (Edge без GUI).
REM Пресет и xlsx выбираются интерактивно в консоли, дальше работает в фоне.

cd /d "%~dp0"
"%~dp0asud_resolutions.exe" --headless

pause
