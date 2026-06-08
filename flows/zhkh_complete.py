"""
flows/zhkh_complete.py — Второй проход для ГИСЖКХ-пресета.

После того как email-flow зарегистрировал все письма и отправил их «На
резолюцию» под текущей учёткой, эта функция:
  1. Переключается на учётку Басманова
  2. Открывает sidebar «На резолюцию»
  3. По каждой строке списка (registered docs):
     - Фильтрует список по номеру (asud_id)
     - Открывает карточку документа
     - «Создать резолюцию» → выбирает шаблон «Для работы»
     - Включает тоггл «Требуется отчёт» + «Контрольная резолюция»
     - Ставит дату контрольного этапа (планируемая из ZHKH, или today+3)
     - Исполнитель — Халецкая Ю. В. (или конфиг)
     - «Добавить» → «Сохранить и отправить» → confirm «Да»
     - Кнопка «Завершить» → карточка закрывается сама
  4. (Опционально) переключается обратно на исходную учётку
"""

import os
import sys
import logging

import openpyxl

from shared.asud_resolution import (
    switch_account, click_sidebar_section, find_doc_row, open_doc_card,
    click_create_resolution, select_content_template, toggle_switch,
    set_stage_date_explicit, compute_control_date,
    fill_executor, click_add_btn, submit_resolution,
    click_complete_button, close_card_after_complete, clear_filter,
)

log = logging.getLogger("asud")


def _read_asud_ids(xlsx_path):
    """Читает все непустые АСУД-номера из колонки 'Номер' реестра."""
    ids = []
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        for r in rows[1:]:
            asud_id = str(r[0] or '').strip()
            if asud_id:
                ids.append({'asud_id': asud_id, 'planned_date': None})
        wb.close()
    except Exception as e:
        log.error(f"Не удалось прочитать {xlsx_path}: {e}")
    return ids


def _process_one(driver, doc, idx, total, executor_fio, content_template,
                  require_report, control_resolution, fallback_days):
    """Один документ: открыть → резолюция → исполнитель → завершить."""
    asud_id = doc['asud_id']
    planned = doc.get('planned_date')
    log.info(f"--- [{idx}/{total}] {asud_id} (планируемая: {planned}) ---")

    # 1. Найти в списке
    row = find_doc_row(driver, asud_id)
    if not row:
        log.warning(f"  {asud_id}: не найден в списке")
        return False

    # 2. Открыть карточку
    if not open_doc_card(driver, row):
        log.error(f"  {asud_id}: не открылась карточка")
        return False

    # 3. Создать резолюцию
    if not click_create_resolution(driver):
        log.error(f"  {asud_id}: 'Создать резолюцию' не сработала")
        return False

    # 4. Содержание
    if not select_content_template(driver, content_template):
        log.warning(f"  {asud_id}: шаблон '{content_template}' не выбран")

    # 5. Тоггл 'Требуется отчёт'
    if require_report:
        toggle_switch(driver, "Требуется отчёт", "true")

    # 6. Тоггл 'Контрольная резолюция'
    if control_resolution:
        toggle_switch(driver, "Контрольная резолюция", "true")

        # 7. Дата контрольного этапа — из планируемой или today+N
        deadline = compute_control_date(planned, fallback_days=fallback_days)
        set_stage_date_explicit(driver, deadline)

    # 8. Исполнитель
    if not fill_executor(driver, executor_fio):
        log.error(f"  {asud_id}: исполнитель не введён")
        return False

    # 9. Кнопка 'Добавить'
    if not click_add_btn(driver):
        log.error(f"  {asud_id}: 'Добавить' не сработала")
        return False

    # 10. Сохранить и отправить + confirm
    if not submit_resolution(driver):
        log.error(f"  {asud_id}: 'Сохранить и отправить' не сработало")
        return False

    # 11. «Завершить» (зелёная кнопка)
    if not click_complete_button(driver):
        log.warning(f"  {asud_id}: 'Завершить' не нажалась")
        # Не критично — карточка может закрыться сама

    # 12. Закрытие карточки (safety net)
    close_card_after_complete(driver)

    log.info(f"  {asud_id}: ✓ завершён")
    return True


def complete_resolutions(driver, docs=None, xlsx_path=None,
                          switch_to="Басманов", switch_back_to="",
                          sidebar_section="На резолюцию",
                          executor_fio="Халецкая Юлия Владимировна",
                          content_template="Для работы",
                          require_report=True,
                          control_resolution=True,
                          fallback_days=3):
    """Второй проход — выдача резолюции + 'Завершить' по списку.

    driver — уже залогиненная сессия
    docs — список dict'ов: [{'asud_id': 'ОРТС/8/...', 'planned_date': 'DD.MM.YYYY'}, ...]
            Если None — читается из xlsx_path (только asud_id, planned_date=None).
    xlsx_path — fallback путь к реестру
    switch_to — фрагмент ФИО учётки на которую переключаемся (Басманов)
    switch_back_to — фрагмент ФИО исходной учётки (пусто = не возвращать)
    sidebar_section — пункт sidebar после переключения ('На резолюцию')
    executor_fio — ФИО исполнителя для резолюции (Халецкая)
    content_template — шаблон 'Содержание' ('Для работы')
    require_report — включать тоггл 'Требуется отчёт'
    control_resolution — включать тоггл 'Контрольная резолюция'
    fallback_days — если планируемая дата в прошлом, ставим today + N кал. дней
    """
    if docs is None:
        if not xlsx_path or not os.path.isfile(xlsx_path):
            log.error(f"Нет ни docs, ни xlsx_path: {xlsx_path!r}")
            return
        docs = _read_asud_ids(xlsx_path)

    if not docs:
        log.warning("Список документов пуст — нечего обрабатывать")
        return

    log.info("=" * 60)
    log.info(f"ZHKH-COMPLETE: второй проход ({len(docs)} документов)")
    log.info(f"  исполнитель:  {executor_fio}")
    log.info(f"  содержание:   {content_template}")
    log.info(f"  тогглы:       отчёт={require_report}, контрольная={control_resolution}")
    log.info(f"  fallback дни: {fallback_days} (если планируемая просрочена)")
    log.info("=" * 60)

    # Переключение на Басманова
    if not switch_account(driver, switch_to):
        log.error(f"Не удалось переключиться на {switch_to} — отмена")
        return

    # Sidebar
    if not click_sidebar_section(driver, sidebar_section):
        log.error(f"Не удалось перейти в '{sidebar_section}' — отмена")
        if switch_back_to:
            switch_account(driver, switch_back_to)
        return

    done, failed = 0, 0
    for i, doc in enumerate(docs, 1):
        try:
            if _process_one(driver, doc, i, len(docs), executor_fio,
                             content_template, require_report,
                             control_resolution, fallback_days):
                done += 1
            else:
                failed += 1
        except Exception as e:
            log.error(f"Exception на {doc.get('asud_id')}: {e}")
            failed += 1
        # Очищаем фильтр для следующей итерации
        try:
            clear_filter(driver)
        except Exception:
            pass

    log.info("=" * 60)
    log.info(f"ZHKH-COMPLETE: выдано резолюций {done}/{len(docs)}, ошибок {failed}")
    log.info("=" * 60)

    if switch_back_to:
        switch_account(driver, switch_back_to)

    return done, failed


# ================= STANDALONE RUN =================

def main():
    """Восстановительный запуск второго прохода: открыть Edge, дождаться
    логина, выдать резолюции по реестру ГИСЖКХ.

    Используется когда первый проход (email-flow) уже завершился (все письма
    зарегистрированы), но второй проход прервался — например, компьютер ушёл
    в спящий режим. `_process_one` молча пропускает документы, которых уже
    нет в сайдбаре «На резолюцию», поэтому повторный запуск безопасен:
    добьёт только недоделанные.

    Путь к xlsx берётся из ASUD_XLSX env или argv[1] или интерактивно из
    Registered/. Драйвер пути и логин — как в остальных flow'ах.
    """
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service as EdgeService

    from shared import config as cfg
    from shared.ui import wait_asud_loaded, set_driver_timeout

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%H:%M:%S',
        handlers=[logging.StreamHandler()],
    )

    settings = cfg.load()
    cfg.setup_file_logger("zhkh_complete")
    cfg.keep_system_awake(True)

    log.info("=" * 50)
    log.info("АСУД ИК — ZHKH-COMPLETE (восстановление второго прохода)")
    log.info("=" * 50)

    base_dir = cfg.get_base_dir()

    # Путь к xlsx
    xlsx_path = os.environ.get('ASUD_XLSX') or (sys.argv[1] if len(sys.argv) > 1 else None)
    if not xlsx_path:
        registered_dir = os.path.join(base_dir, "Registered")
        candidates = []
        if os.path.isdir(registered_dir):
            candidates = sorted(
                [f for f in os.listdir(registered_dir)
                 if f.lower().endswith('.xlsx')
                 and 'гисжкх' in f.lower()
                 and not f.startswith('~')],
                reverse=True,
            )
        if candidates:
            print("\nРеестры ГИСЖКХ в Registered/ (свежие сверху):")
            for i, name in enumerate(candidates[:10], 1):
                print(f"  {i}. {name}")
            print(f"[Enter] = 1")
            choice = input(f"Номер (1-{min(10, len(candidates))}) или путь: ").strip().strip('"').strip("'")
            if not choice:
                xlsx_path = os.path.join(registered_dir, candidates[0])
            elif choice.isdigit():
                try:
                    xlsx_path = os.path.join(registered_dir, candidates[int(choice) - 1])
                except IndexError:
                    log.error(f"Неверный номер: {choice}")
                    sys.exit(1)
            else:
                xlsx_path = choice
        else:
            xlsx_path = input("\nПуть к ГИСЖКХ-реестру: ").strip().strip('"').strip("'")

    if not xlsx_path or not os.path.isfile(xlsx_path):
        log.error(f"Реестр не найден: {xlsx_path!r}")
        input("Enter...")
        sys.exit(1)

    log.info(f"Реестр: {xlsx_path}")

    # Браузер
    driver_path = os.path.join(base_dir, "msedgedriver.exe")
    if not os.path.exists(driver_path):
        log.error(f"msedgedriver.exe не найден в {base_dir}")
        input("Enter...")
        sys.exit(1)

    options = cfg.build_edge_options()
    service = EdgeService(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)
    set_driver_timeout(driver, settings.get("asud_load_timeout_sec",
                                              cfg.DEFAULTS["asud_load_timeout_sec"]))
    try:
        url = settings.get("asud_url", cfg.DEFAULTS["asud_url"])
        log.info(f"Открываю {url}")
        driver.get(url)
        wait_asud_loaded(driver)

        switch_to    = settings.get("zhkh_complete_user", "Басманов")
        switch_back  = settings.get("zhkh_initial_user", "")
        sidebar      = settings.get("zhkh_sidebar_section",
                                     cfg.DEFAULTS["zhkh_sidebar_section"])
        executor     = settings.get("zhkh_executor_fio",
                                     cfg.DEFAULTS["zhkh_executor_fio"])
        content_tpl  = settings.get("zhkh_content_template",
                                     cfg.DEFAULTS["zhkh_content_template"])
        req_report   = settings.get("zhkh_require_report",
                                     cfg.DEFAULTS["zhkh_require_report"])
        ctrl_res     = settings.get("zhkh_control_resolution",
                                     cfg.DEFAULTS["zhkh_control_resolution"])
        fb_days      = settings.get("zhkh_fallback_days",
                                     cfg.DEFAULTS["zhkh_fallback_days"])

        complete_resolutions(driver, xlsx_path=xlsx_path,
                             switch_to=switch_to,
                             switch_back_to=switch_back,
                             sidebar_section=sidebar,
                             executor_fio=executor,
                             content_template=content_tpl,
                             require_report=req_report,
                             control_resolution=ctrl_res,
                             fallback_days=fb_days)
        input("\nEnter для закрытия...")
    except Exception as e:
        log.error(f"Ошибка: {e}")
        input("Enter...")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        cfg.keep_system_awake(False)


if __name__ == "__main__":
    main()
