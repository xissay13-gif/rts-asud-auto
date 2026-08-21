@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

rem Экспериментальный API backend ТОЛЬКО для папки ГИС ЖКХ.
rem Секреты этот BAT не принимает и не хранит: задайте ASUD_API_* env заранее.
set "ASUD_EMAIL_REGISTRATION_BACKEND=asud_api"
set "ASUD_DELETE_AFTER_DONE=0"
set "ASUD_API_MAX_DOCUMENTS=1"
set "ASUD_API_CHOICE="
set "ASUD_API_CONFIRM="

if /I "%~1"=="dry-run" goto dry_run
if /I "%~1"=="probe" goto probe
if /I "%~1"=="live-one" goto live_one
if not "%~1"=="" goto bad_mode

echo.
echo ГИС ЖКХ — тест регистрации через API
echo.
echo   1. dry-run  — разобрать одно письмо, без HTTP и без изменений в АСУД
echo   2. probe    — проверить API чтением, без создания/регистрации
echo   3. live-one — создать и зарегистрировать ровно ОДИН документ
echo.
set /p "ASUD_API_CHOICE=Выберите режим [1]: "
if "%ASUD_API_CHOICE%"=="" goto dry_run
if "%ASUD_API_CHOICE%"=="1" goto dry_run
if "%ASUD_API_CHOICE%"=="2" goto probe
if "%ASUD_API_CHOICE%"=="3" goto live_one
goto bad_mode

:dry_run
set "ASUD_API_ENABLED=1"
set "ASUD_API_MODE=dry-run"
set "ASUD_API_ALLOW_MUTATIONS=0"
echo.
echo DRY-RUN: HTTP-вызовов и изменений в АСУД не будет.
goto run

:probe
set "ASUD_API_ENABLED=1"
set "ASUD_API_MODE=probe"
set "ASUD_API_ALLOW_MUTATIONS=0"
echo.
echo PROBE: разрешены только проверочные запросы чтения к API.
goto run

:live_one
echo.
echo ВНИМАНИЕ: LIVE-ONE СОЗДАСТ И ЗАРЕГИСТРИРУЕТ ОДИН реальный документ.
echo Проверьте endpoint'ы, ID, папку ГИС ЖКХ и письмо перед продолжением.
set "ASUD_API_CONFIRM="
set /p "ASUD_API_CONFIRM=Для запуска введите LIVE-ONE: "
if /I not "%ASUD_API_CONFIRM%"=="LIVE-ONE" goto cancelled
set "ASUD_API_ENABLED=1"
set "ASUD_API_MODE=live-one"
set "ASUD_API_ALLOW_MUTATIONS=1"
goto run

:bad_mode
echo.
echo Неизвестный режим. Допустимо: dry-run, probe или live-one.
goto done

:cancelled
echo.
echo LIVE-ONE отменён. АСУД не изменялся.
goto done

:run
if not exist "%~dp0asud.exe" (
  echo.
  echo ОШИБКА: asud.exe не найден рядом с BAT.
  goto done
)
"%~dp0asud.exe" --preset=gis-api-test

:done
echo.
pause
endlocal
