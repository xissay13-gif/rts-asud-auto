@echo off
chcp 65001 > nul
REM SYNC-режим: пересмотр всех строк реестров и сверка с АСУД.
REM НЕ создаёт резолюций, только обновляет статус в xlsx если в АСУД
REM уже есть резолюция Халецкой (которую daemon мог пропустить из-за
REM PermissionError при сейве или ручной правки в АСУД).
REM
REM Запускать ПЕРЕД обычным daemon'ом утром или раз в неделю.
REM Пресет 2 — «Басманов → Халецкая (все реестры)».

cd /d "%~dp0"
"%~dp0asud_resolutions.exe" --headless --preset 2 --yes --sync-only

pause
