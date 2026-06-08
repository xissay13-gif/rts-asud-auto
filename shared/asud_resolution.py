"""
shared/asud_resolution.py — Хелперы для работы с разделом «На резолюцию» АСУД.

Перенесены из clean-resolutions ветки (resolutions.py). Используются для
второго прохода после email-flow в ГИСЖКХ-пресете:
  switch_account → click_sidebar_section → для каждого doc:
    filter_by_number → find_doc_row → open_doc_card → выдача резолюции +
    «Завершить» → close_card_after_complete
"""

import time
import logging
from datetime import date, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from shared.ui import click, find_input_near_label, js_set_value, close_open_modals

log = logging.getLogger("asud.resolution")


# ============================================================
# Константы
# ============================================================

# GXT-сетка АСУД: <tr> с обоими классами obj-list-rec и obj-list-task —
# это строки данных. Заголовки/служебные tr этих классов не имеют.
DATA_ROW_XPATH = ("//tr[contains(concat(' ',normalize-space(@class),' '),' obj-list-rec ')"
                  " and contains(concat(' ',normalize-space(@class),' '),' obj-list-task ')]")

# Фильтр-input под колонкой "Номер" — стабильный id из DOM
NUMBER_FILTER_INPUT_ID = "FCPC_Номер-input"
NUMBER_FILTER_CONTAINER_ID = "FCPC_Номер"


# ============================================================
# Переключение учётки
# ============================================================

def _account_active(driver, target_substring, timeout=0):
    """Проверяет, что в шапке страницы (y < 80, где лежит ФИО текущего юзера)
    виден элемент с target_substring.

    timeout > 0 — поллит каждые 0.5с пока шапка не отрендерится. Под Edge 148
    АСУД иногда показывает ФИО пользователя в шапке через 5-8с после
    _wait_profile_loaded. Без ожидания _account_active отдаст False и прога
    подумает что надо переключаться (а в шапке уже та учётка что нужна).
    """
    end = time.monotonic() + max(timeout, 0)
    first_pass = True
    while first_pass or time.monotonic() < end:
        first_pass = False
        try:
            elems = driver.find_elements(By.XPATH,
                f"//*[contains(normalize-space(text()), '{target_substring}')]")
            for e in elems:
                try:
                    if not e.is_displayed():
                        continue
                    rect = e.rect
                    if rect.get('y', 1000) < 80:
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        if time.monotonic() >= end:
            break
        time.sleep(0.5)
    return False


def switch_account(driver, target_substring):
    """Переключается на учётку, ФИО которой содержит target_substring.
    1. Если уже на этой учётке (видно в шапке) — пропускаем
    2. Клик по dropdown-стрелке в шапке профиля
    3. Клик по пункту с нужным ФИО в выпадашке (только в области dropdown)
    4. Пост-верификация: в шапке должен появиться target_substring
    """
    log.info(f"Переключение на учётку: {target_substring}")

    # Early return: уже под нужной учёткой. timeout=10 — ждём пока шапка
    # отрендерится (под Edge 148 ФИО появляется через 5-8с после wait_asud_loaded).
    if _account_active(driver, target_substring, timeout=10):
        log.info(f"Уже под учёткой '{target_substring}' — пропускаю переключение")
        return True

    # Шаг 1: клик ▼ рядом с именем профиля
    try:
        triggers = driver.find_elements(By.CSS_SELECTOR,
            "img[class*='trigger'], img[class*='Trigger'], div[class*='trigger']")
        candidate = None
        for t in triggers:
            try:
                if not t.is_displayed():
                    continue
                rect = t.rect
                if rect.get('y', 0) < 120 and rect.get('x', 0) < 400:
                    candidate = t
                    break
            except Exception:
                continue
        if not candidate:
            log.warning("Стрелка профиля не найдена по trigger-классам, пробую по позиции")
            all_imgs = driver.find_elements(By.TAG_NAME, "img")
            for im in all_imgs:
                try:
                    if not im.is_displayed():
                        continue
                    rect = im.rect
                    if rect.get('y', 0) < 120 and rect.get('x', 0) < 400 \
                            and 10 <= rect.get('width', 0) <= 30:
                        candidate = im
                        break
                except Exception:
                    continue
        if not candidate:
            log.error("Не нашёл dropdown-стрелку профиля")
            return False
        click(driver, candidate, "▼ профиль")
    except Exception as e:
        log.error(f"Ошибка клика по dropdown профиля: {e}")
        return False

    # Шаг 2: клик по пункту со строкой target_substring
    # Используем прямой XPath-search вместо predicate-обхода всех div/span/label
    # (раньше ждали ~25 сек, теперь должны находить за пару секунд).
    try:
        WebDriverWait(driver, 5).until(
            lambda d: any(
                e.is_displayed() for e in d.find_elements(
                    By.XPATH,
                    f"//*[contains(normalize-space(text()), '{target_substring}')]"
                )
            )
        )
    except Exception:
        log.warning(f"Пункт '{target_substring}' в выпадашке не появился за 5с")

    items = driver.find_elements(By.XPATH,
        f"//*[contains(normalize-space(text()), '{target_substring}')]")
    target = None
    for it in items:
        try:
            if not it.is_displayed():
                continue
            rect = it.rect
            # Dropdown профиля раскрывается под триггером (y=24+16=40), пункты
            # обычно на y=50-150. Карточки документов / sidebar — на y > 300.
            # Используем 40 < y < 300 + x < 500. Раньше нижняя граница была 100,
            # из-за чего легитимный пункт выпадашки на y=73 отфильтровывался.
            if 40 < rect.get('y', 0) < 300 and rect.get('x', 0) < 500:
                target = it
                break
        except Exception:
            continue
    if not target:
        # Диагностика: выведем координаты ВСЕХ видимых совпадений
        all_rects = []
        for it in items:
            try:
                if it.is_displayed():
                    r = it.rect
                    all_rects.append(f"(x={r.get('x')}, y={r.get('y')})")
            except Exception:
                pass
        log.error(f"Пункт с '{target_substring}' в dropdown'е профиля не найден "
                  f"(всего совпадений: {len(items)}, видимых: {len(all_rects)}). "
                  f"Координаты видимых: {', '.join(all_rects) or '—'}. "
                  f"Возможно, выпадашка не открылась или искомый пункт вне диапазона.")
        return False

    click(driver, target, f"учётка {target_substring}")
    log.info("Переключение запущено, жду перезагрузку АСУД")
    wait_profile_loaded(driver)

    # Пост-верификация: в шапке должна появиться целевая учётка
    end = time.monotonic() + 30
    while time.monotonic() < end:
        if _account_active(driver, target_substring):
            log.info(f"Учётка переключена на '{target_substring}'")
            return True
        time.sleep(1)
    log.error(f"Учётка НЕ переключилась на '{target_substring}' — "
              f"в шапке нет этого ФИО за 30с. Возможно кликнулся не тот элемент.")
    return False


def wait_profile_loaded(driver, max_wait=60):
    """Ждёт готовности АСУД после смены учётки: readyState + появление
    sidebar-items (не главная кнопка — она у разных юзеров может различаться).
    """
    log.info("Жду готовности АСУД после смены учётки...")
    try:
        WebDriverWait(driver, max_wait).until(
            lambda d: d.execute_script("return document.readyState === 'complete'"))
    except Exception:
        log.warning("readyState не complete")
    # Ждём sidebar (минимум 3 видимых <td> с текстом — это пункты меню)
    try:
        WebDriverWait(driver, max_wait).until(
            lambda d: sum(1 for el in d.find_elements(
                By.XPATH, "//td[normalize-space(text())]"
            ) if el.is_displayed()) >= 3
        )
        log.info("АСУД готов (sidebar отрисован)")
    except Exception:
        log.warning(f"Sidebar не отрисовался за {max_wait}с — продолжаю")


# ============================================================
# Sidebar
# ============================================================

def click_sidebar_section(driver, section_text):
    """Клик по пункту в левом сайдбаре. Пробует:
      1) Точное совпадение по тексту
      2) Несколько вариантов (с разными окончаниями)
      3) Contains-совпадение (если предыдущие не нашли)
    Подождёт до 15с пока пункт появится (sidebar может догружаться).
    """
    log.info(f"Сайдбар → '{section_text}'")

    # Варианты на случай небольших расхождений в названии
    base = section_text.rstrip('еиюуяай')  # отрезаем окончание
    variants = [section_text]
    # Добавим близкие варианты с разными окончаниями
    for suffix in ('', 'е', 'и', 'ю', 'у', 'я', 'й', 'а'):
        v = base + suffix
        if v and v not in variants:
            variants.append(v)

    target = None
    end = time.monotonic() + 15
    while time.monotonic() < end and not target:
        # Точные совпадения для каждой вариации
        for v in variants:
            try:
                items = driver.find_elements(By.XPATH,
                    f"//*[normalize-space(text())='{v}']")
                for it in items:
                    if it.is_displayed():
                        target = it
                        if v != section_text:
                            log.info(f"  найден по варианту: '{v}'")
                        break
                if target:
                    break
            except Exception:
                continue
        if target:
            break
        # Fallback — contains-search
        try:
            items = driver.find_elements(By.XPATH,
                f"//td[contains(normalize-space(text()), '{section_text}')]")
            for it in items:
                if it.is_displayed():
                    target = it
                    log.info(f"  найден по contains: '{it.text!r}'")
                    break
        except Exception:
            pass
        if target:
            break
        time.sleep(0.5)

    if not target:
        # Лог: какие пункты вообще есть в sidebar для диагностики
        try:
            visible_items = [el.text.strip() for el in driver.find_elements(
                By.XPATH, "//td[normalize-space(text())]")
                if el.is_displayed() and (el.text or '').strip()]
            log.error(f"Пункт сайдбара '{section_text}' не найден за 15с. "
                      f"Видимые пункты: {visible_items[:30]}")
        except Exception:
            log.error(f"Пункт сайдбара '{section_text}' не найден")
        return False
    click(driver, target, f"сайдбар: {section_text}")
    try:
        WebDriverWait(driver, 10).until(
            lambda d: len(d.find_elements(By.XPATH, DATA_ROW_XPATH)) > 0)
    except Exception:
        log.debug("Грид пустой за 10s — нет задач или ещё не отрисовался")
    return True


# ============================================================
# Поиск по номеру
# ============================================================

def _set_filter_value(driver, container_id, input_id, value):
    """JS-ввод в фильтр колонки — без эмуляции клавиатуры (headless-friendly)."""
    inp = None
    try:
        inp = driver.find_element(By.ID, input_id)
    except Exception:
        try:
            container = driver.find_element(By.ID, container_id)
            inp = container.find_element(By.CSS_SELECTOR, "input[type='text']")
        except Exception as e:
            log.warning(f"Фильтр {container_id}: input не найден ({e})")
            return False
    try:
        driver.execute_script("""
            var el = arguments[0], v = arguments[1];
            el.focus();
            el.value = v;
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('keyup', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
        """, inp, value)
        return True
    except Exception as e:
        log.warning(f"JS-ввод в фильтр упал: {e}")
        return False


def filter_by_number(driver, asud_id):
    """Вбивает ОПТС/ОРТС-номер в фильтр колонки 'Номер'."""
    log.info(f"Фильтр 'Номер' = {asud_id}")
    return _set_filter_value(driver, NUMBER_FILTER_CONTAINER_ID,
                              NUMBER_FILTER_INPUT_ID, asud_id)


def clear_filter(driver):
    """Очищает фильтр колонки 'Номер'."""
    _set_filter_value(driver, NUMBER_FILTER_CONTAINER_ID,
                       NUMBER_FILTER_INPUT_ID, "")
    time.sleep(0.5)


def find_doc_row(driver, asud_id, timeout=8):
    """Находит строку <tr> после применения фильтра по номеру.
    Возвращает elem или None."""
    if not filter_by_number(driver, asud_id):
        return None
    time.sleep(1.5)  # дебаунс GXT-фильтра
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        try:
            rows = driver.find_elements(By.XPATH, DATA_ROW_XPATH)
            visible = [r for r in rows if r.is_displayed() and (r.text or '').strip()]
            if visible:
                first_text = (visible[0].text or '').replace('\n', ' ')[:80]
                log.info(f"[find] МАТЧ: {asud_id} → {len(visible)} строк")
                log.debug(f"  первая строка: {first_text!r}")
                return visible[0]
        except Exception:
            pass
        time.sleep(0.5)
    log.warning(f"[find] {asud_id} → 0 строк за {timeout}s")
    return None


# ============================================================
# Открытие карточки
# ============================================================

def _card_opened(driver, timeout):
    """Признак открывшейся карточки — кнопка 'Создать резолюцию'."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, "header-action-btn-add_resolution")))
        return True
    except Exception:
        # Альтернативный маркер — header-close-btn (есть в любой карточке)
        try:
            els = driver.find_elements(By.ID, "header-close-btn")
            return any(e.is_displayed() for e in els)
        except Exception:
            return False


def _refresh_first_row(driver):
    """Возвращает свежий <tr> первой видимой строки таблицы (после фильтра)."""
    try:
        rows = driver.find_elements(By.XPATH, DATA_ROW_XPATH)
        for r in rows:
            try:
                if r.is_displayed() and (r.text or '').strip():
                    return r
            except Exception:
                continue
    except Exception:
        pass
    return None


def _meaningful_cell(row):
    """Берёт ячейку с текстом (subject/тип) — не чекбокс, не иконку."""
    try:
        tds = row.find_elements(By.XPATH, ".//td")
    except Exception:
        return None
    best = None
    for td in tds:
        try:
            if not td.is_displayed():
                continue
            txt = (td.text or '').strip()
            w = td.size.get('width', 0)
            if len(txt) > 10:
                return td
            if best is None and w > 50:
                best = td
        except Exception:
            continue
    return best


def open_doc_card(driver, row):
    """Открывает карточку документа из строки. 4 стратегии.

    GXT-сетка обычно реагирует так: одиночный клик выделяет строку,
    двойной открывает карточку. Стратегии — разные способы добиться открытия.
    """
    log.info("[open] открываю карточку")
    fresh = _refresh_first_row(driver) or row

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", fresh)
        time.sleep(0.3)
    except Exception:
        pass

    target_cell = _meaningful_cell(fresh) or fresh

    # Strat 1: click + Enter (самая надёжная для GXT)
    log.info("[open] strat1: click + Enter")
    try:
        ActionChains(driver).move_to_element(target_cell).pause(0.15).click().perform()
        time.sleep(0.4)
        ActionChains(driver).send_keys(Keys.ENTER).perform()
    except Exception as e:
        log.debug(f"  strat1 err: {e}")
    if _card_opened(driver, timeout=5):
        log.info("[open] УСПЕХ: strat1")
        return True

    # Strat 2: ActionChains double_click
    log.info("[open] strat2: dblclick")
    try:
        fresh2 = _refresh_first_row(driver) or fresh
        cell2 = _meaningful_cell(fresh2) or fresh2
        ActionChains(driver).move_to_element(cell2).pause(0.2).double_click().perform()
    except Exception as e:
        log.debug(f"  strat2 err: {e}")
    if _card_opened(driver, timeout=4):
        log.info("[open] УСПЕХ: strat2")
        return True

    # Strat 3: JS mouse event chain
    log.info("[open] strat3: JS event chain")
    try:
        fresh3 = _refresh_first_row(driver) or fresh
        cell3 = _meaningful_cell(fresh3) or fresh3
        driver.execute_script("""
            const el = arguments[0];
            const r = el.getBoundingClientRect();
            const x = r.left + r.width/2, y = r.top + r.height/2;
            const opts = {bubbles:true, cancelable:true, view:window,
                          button:0, buttons:1, clientX:x, clientY:y};
            for (const t of ['mousedown','mouseup','click']) {
                el.dispatchEvent(new MouseEvent(t, {...opts, detail:1}));
            }
            for (const t of ['mousedown','mouseup','click']) {
                el.dispatchEvent(new MouseEvent(t, {...opts, detail:2}));
            }
            el.dispatchEvent(new MouseEvent('dblclick', {...opts, detail:2}));
        """, cell3)
    except Exception as e:
        log.debug(f"  strat3 err: {e}")
    if _card_opened(driver, timeout=4):
        log.info("[open] УСПЕХ: strat3")
        return True

    # Strat 4: клик по <a>/anchor
    log.info("[open] strat4: anchor-click")
    try:
        fresh4 = _refresh_first_row(driver) or fresh
        links = fresh4.find_elements(By.XPATH,
            ".//a | .//*[contains(@class,'gwt-Anchor')] | .//*[contains(@class,'cellClickable')]")
        for lnk in links:
            if lnk.is_displayed():
                try:
                    driver.execute_script("arguments[0].click();", lnk)
                except Exception:
                    try:
                        lnk.click()
                    except Exception:
                        continue
                if _card_opened(driver, timeout=4):
                    log.info("[open] УСПЕХ: strat4")
                    return True
    except Exception as e:
        log.debug(f"  strat4 err: {e}")

    log.warning("[open] ПРОВАЛ: все 4 стратегии не сработали")
    return False


# ============================================================
# Выдача резолюции — порт из clean-resolutions
# ============================================================

def click_create_resolution(driver, timeout=10):
    """Клик 'Создать резолюцию' в открытой карточке. id=header-action-btn-add_resolution."""
    try:
        btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.ID, "header-action-btn-add_resolution")))
        click(driver, btn, "Создать резолюцию")
        # Ждём появления поля 'Содержание' (placeholder='Общие формулировки')
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                By.CSS_SELECTOR, "input[placeholder='Общие формулировки']")))
        except Exception:
            log.debug("Поле 'Содержание' не появилось за 10s")
        return True
    except Exception as e:
        log.error(f"Кнопка 'Создать резолюцию' не найдена: {e}")
        return False


def select_content_template(driver, template_text):
    """Выбирает в поле 'Содержание' пункт из выпадашки."""
    try:
        inp = None
        candidates = driver.find_elements(By.CSS_SELECTOR,
            "input[placeholder='Общие формулировки']")
        for c in candidates:
            if c.is_displayed():
                inp = c
                break
        if not inp:
            inp = find_input_near_label(driver, "Содержание")
        if not inp:
            log.error("Поле 'Содержание' не найдено")
            return False

        click(driver, inp, "Содержание")
        try:
            WebDriverWait(driver, 5).until(
                lambda d: any(it.is_displayed() for it in d.find_elements(
                    By.XPATH, f"//*[normalize-space(text())='{template_text}']")))
        except Exception:
            log.debug("Дропдаун 'Содержание' не появился за 5s")
        items = driver.find_elements(By.XPATH,
            f"//*[normalize-space(text())='{template_text}']")
        target = None
        for it in items:
            try:
                if it.is_displayed():
                    target = it
                    break
            except Exception:
                continue
        if not target:
            log.error(f"Пункт '{template_text}' в выпадашке не найден")
            return False
        click(driver, target, f"Содержание: {template_text}")
        time.sleep(0.5)
        return True
    except Exception as e:
        log.error(f"Ошибка выбора содержания: {e}")
        return False


def _wait_data_value(driver, container, target_value="true", timeout=5):
    """Ждёт пока у контейнера-тоггла data-value станет target_value."""
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        try:
            if container.get_attribute('data-value') == target_value:
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def toggle_switch(driver, label_text, target_value="true"):
    """Переключает тоггл рядом с label_text в нужное состояние."""
    try:
        label = driver.find_element(By.XPATH,
            f"//*[normalize-space(text())='{label_text}']")
        container = label.find_element(By.XPATH,
            "./following::*[contains(@class,'switcherContainer')][1]")
    except Exception as e:
        log.warning(f"Тоггл '{label_text}' не найден: {e}")
        return False

    cur = container.get_attribute('data-value')
    if cur == target_value:
        log.info(f"Тоггл '{label_text}' уже = {target_value}")
        return True

    click(driver, container, f"тоггл {label_text} → {target_value}")
    if _wait_data_value(driver, container, target_value, timeout=3):
        log.info(f"Тоггл '{label_text}' = {target_value}")
        return True
    log.warning(f"Тоггл '{label_text}' не переключился")
    return False


def set_stage_date_explicit(driver, deadline_str):
    """Заполняет дату в поле 'Контрольный этап' = deadline_str (формат DD.MM.YYYY)."""
    try:
        inp = driver.find_element(By.CSS_SELECTOR,
            "input[id*='stage_control_date']")
        js_set_value(driver, inp, deadline_str)
        log.info(f"Контрольный этап: {deadline_str}")
        return True
    except Exception as e:
        log.warning(f"Поле даты этапа не найдено: {e}")
        return False


def compute_control_date(planned_date_str, fallback_days=3):
    """Парсит DD.MM.YYYY. Если дата сегодня или в прошлом (<= today) —
    возвращает today + fallback_days КАЛЕНДАРНЫХ дней. Иначе — саму дату.
    Возвращает строку DD.MM.YYYY готовую для js_set_value.
    """
    today = date.today()
    if planned_date_str:
        try:
            parts = planned_date_str.strip().split('.')
            if len(parts) == 3:
                d = date(int(parts[2]), int(parts[1]), int(parts[0]))
                if d > today:
                    return d.strftime("%d.%m.%Y")
        except Exception:
            pass
    # Просрочено / сегодня / не распарсилось → today + N дней
    return (today + timedelta(days=fallback_days)).strftime("%d.%m.%Y")


def fill_executor(driver, fio):
    """Вбивает ФИО в поле 'Исполнитель' (combobox), выбирает из выпадашки."""
    from shared.correspondent import match_correspondent
    try:
        inp = find_input_near_label(driver, "Исполнитель")
        if not inp:
            inp = driver.find_element(By.ID, "select_combobox-input")
        if not inp:
            log.error("Поле 'Исполнитель' не найдено")
            return False

        surname = fio.split()[0]
        inp.click()
        try:
            inp.clear()
        except Exception:
            pass
        for ch in surname:
            inp.send_keys(ch)
        log.info(f"Введена фамилия исполнителя: {surname}")

        def _candidates():
            results = driver.find_elements(By.XPATH,
                f"//*[contains(text(),'{surname}')]")
            return [r for r in results
                    if r.is_displayed() and r != inp
                    and r.tag_name.lower() != 'input']

        candidates = []
        try:
            WebDriverWait(driver, 5).until(lambda d: len(_candidates()) > 0)
            candidates = _candidates()
        except Exception:
            try:
                inp.send_keys(Keys.ENTER)
                WebDriverWait(driver, 3).until(lambda d: len(_candidates()) > 0)
                candidates = _candidates()
            except Exception:
                pass

        log.info(f"Кандидатов: {len(candidates)}")
        target = None
        for idx, r in enumerate(candidates, 1):
            try:
                txt = (r.text or '').strip()
                if len(txt) > 200:
                    continue
                ok = match_correspondent(txt, fio)
                preview = txt.replace('\n', ' ')[:80]
                log.info(f"  [{idx}] {'OK' if ok else '--'} | {preview!r}")
                if ok and target is None:
                    target = r
            except Exception:
                continue

        if not target:
            log.error(f"Исполнитель '{fio}' не найден в выпадашке")
            return False

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
        ActionChains(driver).move_to_element(target).pause(0.2).click().perform()
        log.info(f"Исполнитель выбран: {fio}")
        return True
    except Exception as e:
        log.error(f"Ошибка ввода исполнителя: {e}")
        return False


def wait_button_enabled(driver, btn_id, timeout=15):
    """Ждёт пока кнопка с data-disabled='1' не станет активной."""
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        try:
            btn = driver.find_element(By.ID, btn_id)
            if btn.get_attribute('data-disabled') != '1' and btn.is_displayed():
                return btn
        except Exception:
            pass
        time.sleep(0.3)
    return None


def click_confirm_yes(driver, timeout=10):
    """Ждёт и кликает 'Да' в confirm-диалоге АСУД (с fallback'ами)."""
    yes_btn = None
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        try:
            btn = driver.find_element(By.ID, "confirm_dialog_btn_yes")
            if btn.is_displayed():
                yes_btn = btn
                break
        except Exception:
            pass
        try:
            btn = driver.find_element(By.CSS_SELECTOR,
                "[id*='confirm_dialog_btn_yes'], [id*='confirm'][id*='yes']")
            if btn.is_displayed():
                yes_btn = btn
                break
        except Exception:
            pass
        try:
            for b in driver.find_elements(By.XPATH,
                    "//*[normalize-space(text())='Да']"):
                if b.is_displayed():
                    yes_btn = b
                    break
        except Exception:
            pass
        if yes_btn:
            break
        time.sleep(0.5)
    if not yes_btn:
        log.warning("Confirm-диалог 'Да' не появился")
        return False
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", yes_btn)
        time.sleep(0.3)
    except Exception:
        pass
    clicked = False
    try:
        ActionChains(driver).move_to_element(yes_btn).pause(0.3).click().perform()
        log.info(f"Клик 'Да' (ActionChains)")
        clicked = True
    except Exception:
        pass
    if not clicked:
        try:
            driver.execute_script("arguments[0].click();", yes_btn)
            log.info("Клик 'Да' (JS)")
            clicked = True
        except Exception:
            pass
    if not clicked:
        try:
            yes_btn.click()
            log.info("Клик 'Да' (native)")
            clicked = True
        except Exception:
            pass
    if clicked:
        try:
            ActionChains(driver).send_keys(Keys.ENTER).perform()
        except Exception:
            pass
        try:
            WebDriverWait(driver, 5).until_not(EC.visibility_of(yes_btn))
        except Exception:
            pass
    return clicked


def click_add_btn(driver, timeout=10):
    """Клик 'Добавить' (id=add_btn) после fill_executor."""
    btn = wait_button_enabled(driver, "add_btn", timeout=timeout)
    if not btn:
        log.error("Кнопка 'Добавить' не активировалась")
        return False
    click(driver, btn, "Добавить")
    return True


def submit_resolution(driver):
    """Финальный шаг резолюции: 'Сохранить и отправить' → confirm 'Да'."""
    log.debug("[submit] жду активации 'Сохранить и отправить' (id=save_and_send_btn)")
    btn = wait_button_enabled(driver, "save_and_send_btn", timeout=15)
    if not btn:
        log.error("Кнопка 'Сохранить и отправить' не активировалась")
        return False
    log.info("[submit] клик 'Сохранить и отправить'")
    click(driver, btn, "Сохранить и отправить")
    log.info("[submit] жду confirm-диалог 'Да'")
    confirmed = click_confirm_yes(driver, timeout=10)
    log.info(f"[submit] confirm 'Да': {'OK' if confirmed else 'не появился'}")
    # Ждём пока модалка 'Корневая резолюция' закроется
    try:
        WebDriverWait(driver, 5).until_not(
            lambda d: any(t.is_displayed() for t in d.find_elements(
                By.XPATH, "//*[contains(text(),'Корневая резолюция')]")))
    except Exception:
        log.warning("[submit] модалка 'Корневая резолюция' ещё открыта")
        close_open_modals(driver)
    return True


# ============================================================
# Кнопка «Завершить»
# ============================================================

def click_complete_button(driver, timeout=10):
    """Клик по кнопке «Завершить» в открытой карточке.

    Кнопка — это <div> с текстом 'Завершить' (без id), зелёный фон #789440.
    Может быть несколько таких элементов в DOM, берём первый видимый.
    """
    log.info("Ищу кнопку 'Завершить'")
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: any(b.is_displayed() for b in d.find_elements(
                By.XPATH, "//div[normalize-space(text())='Завершить']")))
    except Exception:
        log.error("Кнопка 'Завершить' не появилась")
        return False

    candidates = driver.find_elements(By.XPATH,
        "//div[normalize-space(text())='Завершить']")
    btn = None
    for c in candidates:
        try:
            if c.is_displayed():
                btn = c
                break
        except Exception:
            continue
    if not btn:
        log.error("Кнопка 'Завершить' не найдена среди видимых")
        return False

    click(driver, btn, "Завершить")
    log.info("Кнопка 'Завершить' нажата")
    return True


def close_card_after_complete(driver):
    """После клика по 'Завершить' карточка должна закрыться сама.
    Эта функция — safety net: если осталась открыта, кликаем header-close-btn.
    """
    # Сначала ждём что главная вернулась сама (по словам юзера — закрывается)
    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "mainscreen-create-button")))
        log.info("[close] карточка закрылась сама — главная активна")
        return
    except Exception:
        pass

    # Резерв — клик по крестику
    log.info("[close] карточка ещё открыта, клик по header-close-btn")
    try:
        close_btn = driver.find_element(By.ID, "header-close-btn")
        if close_btn.is_displayed():
            try:
                ActionChains(driver).move_to_element(close_btn).pause(0.2).click().perform()
            except Exception:
                driver.execute_script("arguments[0].click();", close_btn)
    except Exception:
        log.debug("[close] header-close-btn не найден")

    try:
        WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.ID, "mainscreen-create-button")))
    except Exception:
        log.warning("[close] главная не загрузилась за 8s")
