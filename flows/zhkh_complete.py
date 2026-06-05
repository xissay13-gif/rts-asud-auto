"""
flows/zhkh_complete.py — Второй проход для ГИСЖКХ-пресета.

После того как email-flow зарегистрировал все письма и сформировал реестр,
эта функция:
  1. Переключается на учётку Басманова
  2. Открывает sidebar «На резолюцию»
  3. По каждой строке реестра (с известным АСУД-ID):
     - Фильтрует список по номеру
     - Открывает карточку
     - Жмёт «Завершить» (карточка сама закрывается, без подтверждения)
  4. Переключается обратно на исходную учётку
"""

import os
import logging

import openpyxl

from shared.asud_resolution import (
    switch_account, click_sidebar_section, find_doc_row, open_doc_card,
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
        # rows[0] — заголовки: Номер | Link | Округ | Subject | Body
        for r in rows[1:]:
            asud_id = str(r[0] or '').strip()
            if asud_id:
                ids.append(asud_id)
        wb.close()
    except Exception as e:
        log.error(f"Не удалось прочитать {xlsx_path}: {e}")
    return ids


def complete_resolutions(driver, asud_ids=None, xlsx_path=None,
                          switch_to="Басманов", switch_back_to="",
                          sidebar_section="На исполнение"):
    """Второй проход — массово жмёт 'Завершить' по списку.

    driver — уже залогиненная сессия (после email-flow)
    asud_ids — список номеров ('ОРТС/8/...') от первого прохода.
               Если не передан — читается из xlsx_path.
    xlsx_path — путь к Registered/YYYY-MM-DD_ГИСЖКХ_резолюции.xlsx (fallback)
    switch_to — фрагмент ФИО учётки на которую переключаемся (Басманов А. В.)
    switch_back_to — фрагмент ФИО исходной учётки для возврата (пустая = не возвращать)
    sidebar_section — название пункта sidebar после переключения. У Басманова
                      документы лежат в 'На исполнение' (не 'На резолюцию' —
                      это у того кто только что зарегистрировал и отправил).
    """
    if asud_ids is None:
        if not xlsx_path or not os.path.isfile(xlsx_path):
            log.error(f"Нет ни asud_ids, ни xlsx_path: {xlsx_path!r}")
            return
        asud_ids = _read_asud_ids(xlsx_path)

    if not asud_ids:
        log.warning("Список АСУД-ID пуст — нечего завершать")
        return

    log.info("=" * 60)
    log.info(f"ZHKH-COMPLETE: второй проход ({len(asud_ids)} документов)")
    log.info("=" * 60)

    # Переключение на Басманова
    if not switch_account(driver, switch_to):
        log.error(f"Не удалось переключиться на {switch_to} — отмена второго прохода")
        return

    # Sidebar → нужный раздел (у Басманова это 'На исполнение')
    if not click_sidebar_section(driver, sidebar_section):
        log.error(f"Не удалось перейти в '{sidebar_section}' — отмена")
        if switch_back_to:
            switch_account(driver, switch_back_to)
        return

    # Цикл по документам
    done, failed = 0, 0
    for i, asud_id in enumerate(asud_ids, 1):
        log.info(f"--- [{i}/{len(asud_ids)}] {asud_id} ---")
        try:
            row = find_doc_row(driver, asud_id)
            if not row:
                log.warning(f"  {asud_id}: не найден в списке — пропуск")
                failed += 1
                clear_filter(driver)
                continue
            if not open_doc_card(driver, row):
                log.error(f"  {asud_id}: не открылась карточка")
                failed += 1
                clear_filter(driver)
                continue
            if not click_complete_button(driver):
                log.error(f"  {asud_id}: 'Завершить' не нажалась")
                failed += 1
                continue
            close_card_after_complete(driver)
            clear_filter(driver)
            done += 1
            log.info(f"  {asud_id}: завершён")
        except Exception as e:
            log.error(f"  {asud_id}: исключение {e}")
            failed += 1
            try:
                clear_filter(driver)
            except Exception:
                pass

    log.info("=" * 60)
    log.info(f"ZHKH-COMPLETE: завершено {done}/{len(asud_ids)}, ошибок {failed}")
    log.info("=" * 60)

    # Возврат на исходную учётку
    if switch_back_to:
        switch_account(driver, switch_back_to)

    return done, failed
