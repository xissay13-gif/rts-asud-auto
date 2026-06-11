"""
flows/zhkh_daemon.py — Непрерывный мониторинг реестров и выдача резолюций
Халецкой под учёткой Басманова.

P2 в конвейере регистратор → Басманов → Халецкая → округа:
  • P1 (регистратор, разные учётки) пишет новые строки в xlsx-реестры
    с заполненной колонкой «Номер»
  • P2 (этот daemon) опрашивает реестры, для каждой новой строки выдаёт
    резолюцию Халецкой и помечает «Отписано Халецкой» с датой
  • P3 (clean-resolutions) опрашивает дальше и обрабатывает «Отписано в округ»

Конфигурация — секция `zhkh_daemon` в settings.json (см. settings.json.example).

Запуск: asud.exe --mode=zhkh-daemon [--headless]
        Ctrl+C для корректной остановки.
"""

import glob
import os
import signal
import sys
import time
import logging
from datetime import timedelta

import openpyxl
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService

from shared import config as cfg
from shared.ui import wait_asud_loaded, set_driver_timeout
from shared.xlsx_lock import xlsx_lock, is_xlsx_busy
from shared.xlsx_status import get_done_asud_ids, mark_status, COL_HALETSKAYA
from shared.asud_resolution import (
    switch_account, click_sidebar_section, find_doc_row, open_doc_card,
    click_create_resolution, has_existing_resolution_to, row_has_resolution_to,
    select_content_template, toggle_switch,
    set_stage_date_explicit, compute_control_date,
    fill_executor, click_add_btn, submit_resolution,
    click_complete_button, close_card_after_complete, clear_filter,
    _account_active,
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger("asud")
start_time = time.monotonic()


# ================= STOP HANDLING (Ctrl+C) =================
_stop_flag = False


def _on_sigint(signum, frame):
    global _stop_flag
    if _stop_flag:
        log.error("Повторный Ctrl+C — выхожу немедленно")
        sys.exit(1)
    _stop_flag = True
    log.info("Получен Ctrl+C — остановлюсь после текущего документа")


def _interruptible_sleep(seconds):
    """Sleep с проверкой _stop_flag раз в 0.5с."""
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        if _stop_flag:
            return
        time.sleep(0.5)


# ================= READING THE XLSX =================

def _read_todo(xlsx_path, status_column=COL_HALETSKAYA):
    """Читает реестр и возвращает список todo-строк.

    Включает только строки с:
      • непустым asud_id (Номер)
      • пустой колонкой status_column

    Каждая строка — dict с asud_id, planned_date, xlsx_path.
    Lock-acquire не нужен — read-only через openpyxl, не пишем.
    """
    out = []
    if not xlsx_path or not os.path.isfile(xlsx_path):
        return out
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            wb.close()
            return out
        headers = [str(c or '').strip() for c in rows[0]]
        headers_lower = [h.lower() for h in headers]
        # Колонка asud_id — Номер / ОПТС / ОРТС / асуд
        asud_keys = ('номер', 'опт', 'орт', 'асуд', 'asud', 'регистрацион')
        asud_col = None
        for i, h in enumerate(headers_lower):
            if any(k in h for k in asud_keys):
                asud_col = i
                break
        if asud_col is None:
            wb.close()
            return out
        try:
            status_col = headers.index(status_column)
        except ValueError:
            status_col = None  # колонки нет — все строки в todo
        # Опционально: «Планируемая дата исполнения» — если есть
        planned_col = None
        for i, h in enumerate(headers_lower):
            if 'планиру' in h:
                planned_col = i
                break

        for r in rows[1:]:
            if asud_col >= len(r):
                continue
            aid = str(r[asud_col] or '').strip()
            if not aid:
                continue
            if status_col is not None and status_col < len(r):
                status = str(r[status_col] or '').strip() if r[status_col] is not None else ''
                if status:
                    continue
            planned = None
            if planned_col is not None and planned_col < len(r):
                pv = r[planned_col]
                if pv:
                    planned = str(pv).strip()
            out.append({
                'asud_id': aid,
                'planned_date': planned,
                'xlsx_path': xlsx_path,
            })
        wb.close()
    except PermissionError:
        log.debug(f"_read_todo: {xlsx_path} занят (Excel?), скип")
    except Exception as e:
        log.warning(f"_read_todo({xlsx_path}): {e}")
    return out


def _list_watched_xlsx(watch_list, verbose=False):
    """Расхлопывает watch-list (list of {dir, xlsx_pattern}) в список xlsx-путей.

    verbose=True (на старте daemon'а) — пишет в лог почему пропускается
    каждая запись если она не отдала ни одного файла. По умолчанию (в loop)
    скипает молча чтобы не флудить.
    """
    out = []
    for entry in watch_list or []:
        d = (entry.get("dir") or "").strip()
        pat = (entry.get("xlsx_pattern") or "*.xlsx").strip()
        if not d:
            if verbose:
                log.warning(f"watch-entry без поля 'dir': {entry!r}")
            continue
        if not os.path.isdir(d):
            if verbose:
                # Помогаем юзеру понять что именно не так.
                # repr() покажет странные символы (mojibake/BOM/непечатные).
                hint = "сетевая шара не примонтирована?" if d[:2] in (r"\\", "//") \
                       else "проверь правильность пути и экранирование \\\\ в JSON"
                log.error(f"watch.dir не существует: {d!r} ({hint})")
            continue
        found = 0
        for p in glob.glob(os.path.join(d, pat)):
            # Скип временных файлов Excel ($файл.xlsx, ~$файл.xlsx)
            name = os.path.basename(p)
            if name.startswith('~$') or name.startswith('$'):
                continue
            out.append(p)
            found += 1
        if verbose and found == 0:
            log.warning(f"watch.dir {d!r}: папка есть, но '{pat}' не нашёл "
                        f"ни одного файла")
    return out


# ================= PROCESSING ONE DOC =================

def _process_one(driver, doc, executor_fio, content_template,
                  require_report, control_resolution, fallback_days):
    """Обработка одного документа: находит в списке, выдаёт резолюцию
    Халецкой и «Завершить». Возвращает True если резолюция выдана,
    False если skip/fail."""
    asud_id = doc['asud_id']
    planned = doc.get('planned_date')
    log.info(f"--- {asud_id} (планируемая: {planned}) ---")

    row = find_doc_row(driver, asud_id)
    if not row:
        log.info(f"  {asud_id}: не найден в списке (АСУД ещё не пропагировал?) — пропуск, попробуем позже")
        return False

    # БЫСТРАЯ проверка в строке списка (колонка «Направлено») — БЕЗ открытия
    # карточки. Если фамилия уже видна там — резолюция выдана, пропуск.
    if row_has_resolution_to(row, executor_fio):
        surname = executor_fio.split()[0] if executor_fio else '?'
        log.info(f"  {asud_id}: в «Направлено» уже видно «{surname}» — "
                 f"пропуск + синхрон xlsx (без открытия карточки)")
        xlsx_p = doc.get('xlsx_path')
        if xlsx_p:
            if mark_status(xlsx_p, asud_id, COL_HALETSKAYA):
                log.debug(f"  → xlsx: {asud_id} помечен «{COL_HALETSKAYA}» (синхрон по списку)")
        return True

    if not open_doc_card(driver, row):
        log.warning(f"  {asud_id}: не открылась карточка")
        return False

    # Защита от дублей: если в карточке уже есть резолюция на этого
    # исполнителя — синхронизируем xlsx и пропускаем без создания второй.
    # Fallback после row-проверки — на случай если «Направлено» по каким-то
    # причинам пустая, но в панели «Этапы исполнения» резолюция всё-таки есть.
    if has_existing_resolution_to(driver, executor_fio):
        surname = executor_fio.split()[0] if executor_fio else '?'
        log.info(f"  {asud_id}: в карточке уже есть резолюция на {surname} — "
                 f"синхронизирую xlsx, пропускаю создание дубля")
        xlsx_p = doc.get('xlsx_path')
        if xlsx_p:
            if mark_status(xlsx_p, asud_id, COL_HALETSKAYA):
                log.debug(f"  → xlsx: {asud_id} помечен «{COL_HALETSKAYA}» (синхрон)")
        close_card_after_complete(driver)
        return True  # засчитываем как обработано

    if not click_create_resolution(driver):
        log.warning(f"  {asud_id}: 'Создать резолюцию' не сработала")
        return False

    if not select_content_template(driver, content_template):
        log.warning(f"  {asud_id}: шаблон '{content_template}' не выбран")

    if require_report:
        toggle_switch(driver, "Требуется отчёт", "true")

    if control_resolution:
        toggle_switch(driver, "Контрольная резолюция", "true")
        deadline = compute_control_date(planned, fallback_days=fallback_days)
        set_stage_date_explicit(driver, deadline)

    if not fill_executor(driver, executor_fio):
        log.error(f"  {asud_id}: исполнитель не введён")
        return False

    if not click_add_btn(driver):
        log.error(f"  {asud_id}: 'Добавить' не сработала")
        return False

    if not submit_resolution(driver):
        log.error(f"  {asud_id}: 'Сохранить и отправить' не сработало")
        return False

    if not click_complete_button(driver):
        log.warning(f"  {asud_id}: 'Завершить' не нажалась (некритично)")

    close_card_after_complete(driver)
    log.info(f"  {asud_id}: ✓ резолюция выдана")
    return True


# ================= SESSION HEARTBEAT =================

HEARTBEAT_INTERVAL_SEC = 300  # каждые 5 минут проверяем не разлогинило ли


def _session_alive(driver, expected_user):
    """Проверяет что в шапке всё ещё видна нужная учётка. Использует
    _account_active с коротким таймаутом — если шапки нет, скорее всего
    нас разлогинило (АСУД редиректит на login-страницу)."""
    try:
        return _account_active(driver, expected_user, timeout=3)
    except Exception:
        return False


def _relogin(driver, url, switch_to, sidebar):
    """Полный re-login: открыть АСУД заново → переключить учётку → сайдбар.
    Используется когда _session_alive отдал False (idle/timeout разлогинил).
    Возвращает True если успешно восстановили сессию.
    """
    log.warning(f"Heartbeat: сессия с '{switch_to}' потерялась — переподключаюсь")
    try:
        driver.get(url)
        wait_asud_loaded(driver)
    except Exception as e:
        log.error(f"Re-login: driver.get упал: {e}")
        return False
    if not switch_account(driver, switch_to):
        log.error(f"Re-login: switch_account на '{switch_to}' не сработал")
        return False
    if not click_sidebar_section(driver, sidebar):
        log.error(f"Re-login: sidebar '{sidebar}' не открылся")
        return False
    log.info("Heartbeat: сессия восстановлена")
    return True


# ================= DAEMON LOOP =================

def main():
    settings = cfg.load()
    cfg.setup_file_logger("zhkh_daemon")
    cfg.keep_system_awake(True)

    dcfg = settings.get("zhkh_daemon") or {}
    if not dcfg:
        log.error("В settings.json нет секции 'zhkh_daemon'. "
                  "Прим. см. settings.json.example.")
        input("Enter...")
        sys.exit(1)

    watch_list      = dcfg.get("watch") or []
    poll_interval   = int(dcfg.get("poll_interval_sec", 300))
    switch_to       = dcfg.get("switch_to_user", "Басманов")
    sidebar         = dcfg.get("sidebar_section", "На резолюцию")
    executor        = dcfg.get("executor_fio", "Халецкая Юлия Владимировна")
    content_tpl     = dcfg.get("content_template", "Для работы")
    req_report      = bool(dcfg.get("require_report", True))
    ctrl_res        = bool(dcfg.get("control_resolution", True))
    fb_days         = int(dcfg.get("fallback_days", 3))
    round_robin     = bool(dcfg.get("round_robin", False))
    rr_idx          = 0

    if not watch_list:
        log.error("zhkh_daemon.watch пуст. Укажи папки и xlsx-pattern в settings.json.")
        input("Enter...")
        sys.exit(1)

    base_dir = cfg.get_base_dir()
    driver_path = os.path.join(base_dir, "msedgedriver.exe")
    if not os.path.exists(driver_path):
        log.error(f"msedgedriver.exe не найден в {base_dir}")
        input("Enter...")
        sys.exit(1)

    log.info("=" * 60)
    log.info(f"ZHKH-DAEMON: запуск (учётка {switch_to})")
    for entry in watch_list:
        log.info(f"  watch: {entry.get('dir')!r} pattern={entry.get('xlsx_pattern')!r}")
    log.info(f"  poll: каждые {poll_interval}с"
             + (" (round-robin по реестрам)" if round_robin else ""))
    log.info(f"  executor: {executor}, content: {content_tpl}, fallback: today+{fb_days} кал.дн.")
    log.info("=" * 60)

    # Проверка watch-путей на старте: если папка не существует / xlsx не нашёлся
    # — вылетит ERROR в логе с подсказкой. Без этого юзер видит только эхо
    # конфига и потом тихое «нет xlsx ни в одной watch-папке».
    initial_paths = _list_watched_xlsx(watch_list, verbose=True)
    if initial_paths:
        log.info(f"Watch-старт: найдено {len(initial_paths)} реестров")
        for p in initial_paths:
            log.info(f"  ✓ {p}")
    else:
        log.warning("Watch-старт: ни одного реестра не найдено. "
                    "Daemon всё равно стартует — будет переопрашивать каждые "
                    f"{poll_interval}с (на случай если папки появятся позже).")

    signal.signal(signal.SIGINT, _on_sigint)
    try:
        signal.signal(signal.SIGTERM, _on_sigint)
    except (AttributeError, ValueError):
        pass

    options = cfg.build_edge_options()
    service = EdgeService(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)
    set_driver_timeout(driver, settings.get("asud_load_timeout_sec",
                                              cfg.DEFAULTS["asud_load_timeout_sec"]))

    iters = 0
    totals = {"OK": 0, "SKIP": 0, "FAIL": 0}
    try:
        url = settings.get("asud_url", cfg.DEFAULTS["asud_url"])
        log.info(f"Открываю {url}")
        driver.get(url)
        wait_asud_loaded(driver)

        if not switch_account(driver, switch_to):
            log.error(f"Не удалось переключиться на '{switch_to}' — выход")
            sys.exit(1)

        if not click_sidebar_section(driver, sidebar):
            log.error(f"Не удалось перейти в '{sidebar}' — выход")
            sys.exit(1)

        # Кэш mtime: xlsx_path → last seen mtime. Если файл не менялся
        # с прошлой итерации, повторно не читаем (оптимизация).
        mtime_cache = {}
        last_heartbeat = time.monotonic()

        while not _stop_flag:
            iters += 1

            # Heartbeat-проверка сессии каждые HEARTBEAT_INTERVAL_SEC секунд.
            # Если АСУД разлогинило за idle/timeout — пробуем переподключиться.
            if time.monotonic() - last_heartbeat > HEARTBEAT_INTERVAL_SEC:
                if not _session_alive(driver, switch_to):
                    if not _relogin(driver, url, switch_to, sidebar):
                        log.error("Heartbeat: re-login не удался, sleep + retry")
                        _interruptible_sleep(60)
                        continue
                last_heartbeat = time.monotonic()

            xlsx_paths = _list_watched_xlsx(watch_list)
            if not xlsx_paths:
                log.info(f"[итер. {iters}] нет xlsx ни в одной watch-папке — sleep {poll_interval}с")
                _interruptible_sleep(poll_interval)
                continue

            # Round-robin (опционально): каждый тик берём ОДИН реестр в
            # порядке списка. Удобно когда хочется чёткое расписание
            # «тик1=ОЭК, тик2=ТЭС, тик3=ГИСЖКХ, тик4=ОЭК...».
            # Дефолт — ALL (все реестры на каждом тике + mtime-кэш).
            if round_robin and xlsx_paths:
                cur_xpath = xlsx_paths[rr_idx % len(xlsx_paths)]
                rr_idx += 1
                log.info(f"[итер. {iters}] round-robin: проверяю "
                         f"{os.path.basename(cur_xpath)}")
                paths_this_tick = [cur_xpath]
            else:
                paths_this_tick = xlsx_paths

            todo_all = []
            changed_files = 0
            busy_files = 0
            for xpath in paths_this_tick:
                # Приоритет регистрации: если рядом с реестром свежий .lock
                # (P1-регистратор пишет новые строки) — пропускаем на этом
                # тике, вернёмся через poll_interval.
                if is_xlsx_busy(xpath):
                    log.info(f"[итер. {iters}] {os.path.basename(xpath)} занят "
                             f"(регистратор пишет?) — пропускаю до след. тика")
                    busy_files += 1
                    continue
                try:
                    cur_mtime = os.path.getmtime(xpath)
                except OSError:
                    continue
                # Кэш: пропускаем если не менялся И были обработаны раньше
                last_mtime = mtime_cache.get(xpath, 0)
                if cur_mtime == last_mtime and iters > 1:
                    continue
                changed_files += 1
                todo_all.extend(_read_todo(xpath))
                mtime_cache[xpath] = cur_mtime

            if not todo_all:
                if iters == 1:
                    log.info(f"[итер. 1] просканировано {len(xlsx_paths)} xlsx, todo пусто — жду новые записи")
                else:
                    log.debug(f"[итер. {iters}] изменилось файлов: {changed_files}, todo: 0")
                _interruptible_sleep(poll_interval)
                continue

            log.info(f"[итер. {iters}] todo: {len(todo_all)} документов")
            for doc in todo_all:
                if _stop_flag:
                    break
                try:
                    if _process_one(driver, doc, executor, content_tpl,
                                     req_report, ctrl_res, fb_days):
                        totals["OK"] += 1
                        # Сразу помечаем в xlsx чтобы следующая итерация не подхватила
                        mark_status(doc['xlsx_path'], doc['asud_id'], COL_HALETSKAYA)
                    else:
                        totals["SKIP"] += 1
                except Exception as e:
                    log.error(f"  {doc['asud_id']}: exception {e}")
                    totals["FAIL"] += 1
                try:
                    clear_filter(driver)
                except Exception:
                    pass
                # mtime файла поменялся (мы записали статус) → форсируем
                # пере-чтение этого xlsx на следующей итерации.
                mtime_cache[doc['xlsx_path']] = 0

            log.info(f"[итер. {iters}] OK={totals['OK']} SKIP={totals['SKIP']} FAIL={totals['FAIL']}")
            _interruptible_sleep(poll_interval)

    except KeyboardInterrupt:
        log.info("KeyboardInterrupt — выхожу")
    except Exception as e:
        log.error(f"FATAL: {e}")
    finally:
        log.info("=" * 60)
        log.info(f"ZHKH-DAEMON: остановка после {iters} итераций")
        log.info(f"  итого: OK={totals['OK']} SKIP={totals['SKIP']} FAIL={totals['FAIL']}")
        log.info("=" * 60)
        try:
            driver.quit()
        except Exception:
            pass
        cfg.keep_system_awake(False)


if __name__ == "__main__":
    main()
