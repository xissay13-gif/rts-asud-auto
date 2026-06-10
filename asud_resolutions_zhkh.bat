@echo off
chcp 65001 > nul
REM Пресет №2: «ГИСЖКХ → Халецкая (под Басмановым)» — headless, без вопросов.
REM Берёт первый xlsx из папки (приоритет имени с «_резолюции»).

cd /d "%~dp0"
"%~dp0asud_resolutions.exe" --headless --preset 2 --yes

pause
