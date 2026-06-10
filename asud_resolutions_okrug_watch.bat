@echo off
chcp 65001 > nul
REM Daemon-режим: пресет 1 (Округа → начальницам) с непрерывным мониторингом
REM xlsx-реестра. Опрашивает файл каждые 30с, новые строки сразу подхватывает.
REM Ctrl+C — корректная остановка после текущего документа.

cd /d "%~dp0"
"%~dp0asud_resolutions.exe" --headless --preset 1 --yes --watch

pause
