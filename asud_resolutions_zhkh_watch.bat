@echo off
chcp 65001 > nul
REM Daemon-режим: пресет 2 (ГИСЖКХ → Халецкая под Басмановым) с непрерывным
REM мониторингом xlsx. Дополняет zhkh_daemon из основной сборки — может
REM работать параллельно (xlsx_lock сериализует доступ к реестру).
REM Ctrl+C — корректная остановка.

cd /d "%~dp0"
"%~dp0asud_resolutions.exe" --headless --preset 2 --yes --watch

pause
