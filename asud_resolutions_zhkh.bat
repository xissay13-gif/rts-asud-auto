@echo off
chcp 65001 > nul
REM Пресет №2: «Басманов → Халецкая (все реестры)» — headless, без вопросов.
REM Опрашивает ВСЕ реестры из preset.watch (ОЭК + ТЭС + ГИСЖКХ).

cd /d "%~dp0"
"%~dp0asud_resolutions.exe" --headless --preset 2 --yes

pause
