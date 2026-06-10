@echo off
chcp 65001 > nul
REM Daemon-режим: пресет 2 «Басманов → Халецкая (все реестры)» с непрерывным
REM мониторингом xlsx. Опрашивает ОЭК + ТЭС + ГИСЖКХ из preset.watch.
REM Дополняет zhkh_daemon из основной сборки — может работать параллельно
REM (xlsx_lock сериализует доступ к реестру).
REM Ctrl+C — корректная остановка.

cd /d "%~dp0"
"%~dp0asud_resolutions.exe" --headless --preset 2 --yes --watch

pause
