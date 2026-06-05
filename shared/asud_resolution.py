"""
shared/asud_resolution.py — Хелперы для работы с разделом «На резолюцию» АСУД.

Перенесены из clean-resolutions ветки (resolutions.py). Используются для
второго прохода после email-flow в ГИСЖКХ-пресете:
  switch_account → click_sidebar_section → для каждого doc:
    filter_by_number → find_doc_row → open_doc_card → click_complete_button
                                                    → close_card_after_complete
"""

import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from shared.ui import click

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

def switch_account(driver, target_substring):
    """Переключается на учётку, ФИО которой содержит target_substring.
    1. Клик по dropdown-стрелке в шапке профиля
    2. Клик по пункту с нужным ФИО в выпадашке
    """
    log.info(f"Переключение на учётку: {target_substring}")

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
            if rect.get('y', 0) > 100:
                target = it
                break
        except Exception:
            continue
    if not target and items:
        target = next((i for i in items if i.is_displayed()), None)
    if not target:
        log.error(f"Пункт с '{target_substring}' не найден")
        return False

    click(driver, target, f"учётка {target_substring}")
    log.info("Переключение запущено, жду перезагрузку АСУД")
    wait_profile_loaded(driver)
    return True


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
