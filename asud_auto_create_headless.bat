@echo off
chcp 65001 > nul
REM Auto-create — массовая регистрация из xlsx-реестра в headless-режиме.
REM Двойной клик → выбор xlsx (если их несколько в папке) → создание +
REM регистрация + На резолюцию для каждой строки. Без .msg-вложений
REM (используется пустышка как content).

cd /d "%~dp0"
"%~dp0registration.exe" --headless --mode=auto-create

pause
