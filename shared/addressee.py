"""Выбор адресата с проверкой фактически добавленной строки АСУД.

В GXT поле поиска и таблица выбранных адресатов — соседние компоненты. После
успешного выбора input очищается, а запись появляется в ``#addressee_grid_id``.
Поэтому отправленный click и значение input сами по себе не являются успехом.
"""

import logging
import time

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from shared.correspondent import match_strict
from shared.ui import (
    cdp_click,
    find_dropdown_options,
    find_input_near_label,
    js_type_combobox,
)


log = logging.getLogger("asud.addressee")


_READ_COMMITTED_ADDRESSEES_JS = r"""
const rows = [];
const seen = new Set();

function visible(el) {
    if (!el || !el.isConnected) return false;
    const style = window.getComputedStyle(el);
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0 &&
        style.display !== 'none' && style.visibility !== 'hidden' &&
        parseFloat(style.opacity || '1') > 0.01;
}

for (const grid of document.querySelectorAll("#addressee_grid_id")) {
    if (!visible(grid)) continue;
    for (const content of grid.querySelectorAll("[data-attr='agents-content']")) {
        const row = content.closest('tr');
        if (!row || !visible(row) ||
                !row.querySelector("img[data-attr='remove-trigger']")) continue;
        const text = String(content.innerText || content.textContent || '')
            .replace(/\s+/g, ' ').trim();
        if (text && !seen.has(text)) {
            seen.add(text);
            rows.push(text);
        }
    }
}
return rows;
"""


def _committed_texts(driver):
    """Возвращает только строки уже выбранных адресатов из нужного grid."""
    try:
        state = driver.execute_script(_READ_COMMITTED_ADDRESSEES_JS)
        if isinstance(state, dict):
            state = state.get("rows") or ()
        if isinstance(state, (list, tuple)):
            return tuple(str(value or "").strip() for value in state if str(value or "").strip())
    except Exception:
        pass

    # Совместимость с ограниченными драйверами и тестовыми doubles. Поиск всё
    # равно жёстко ограничен таблицей адресатов; popup и текст карточки сюда не
    # попадают.
    try:
        grid = driver.find_element(By.ID, "addressee_grid_id")
        values = []
        for element in grid.find_elements(By.CSS_SELECTOR, "[data-attr='agents-content']"):
            try:
                if element.is_displayed() and (element.text or "").strip():
                    values.append(element.text.strip())
            except Exception:
                continue
        if values:
            return tuple(values)

        # Legacy/test fallback: реальные data rows, но никогда весь grid/body.
        for element in grid.find_elements(
                By.CSS_SELECTOR, "table[class*='GridStyle-dataTable'] tbody tr[class*='row']"):
            try:
                if element.is_displayed() and (element.text or "").strip():
                    values.append(element.text.strip())
            except Exception:
                continue
        return tuple(values)
    except Exception:
        return ()


def addressee_committed(driver, full_name):
    """True только если точное ФИО/инициалы уже есть в committed grid-row."""
    expected = str(full_name or "").strip()
    if not expected:
        return False
    return any(match_strict(value, expected) for value in _committed_texts(driver))


def wait_addressee_committed(driver, full_name, timeout=3.0, poll_interval=0.2):
    """Ждёт GXT-перерисовку grid после клика, каждый раз читая свежий DOM."""
    deadline = time.monotonic() + max(0.0, float(timeout))
    while True:
        if addressee_committed(driver, full_name):
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(max(0.01, float(poll_interval)))


def _usable_input(element):
    try:
        if str(getattr(element, "tag_name", "")).casefold() != "input":
            return False
        if not element.is_displayed():
            return False
        if hasattr(element, "is_enabled") and not element.is_enabled():
            return False
        if str(element.get_attribute("aria-hidden") or "").casefold() == "true":
            return False
        if element.get_attribute("readonly") is not None:
            return False
        return str(element.get_attribute("type") or "text").casefold() != "hidden"
    except Exception:
        return False


def _find_addressee_input(driver):
    # ID select_combobox-input повторяется у разных picker-ов, поэтому всегда
    # ограничиваем поиск семантическим marker конкретно адресата.
    selectors = (
        "[data-marker='AppointmentPicker_newAddressee'] input#select_combobox-input",
        "[data-marker='AppointmentPicker_newAddressee'] input[type='text']",
    )
    for selector in selectors:
        try:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                if _usable_input(element):
                    return element
        except Exception:
            continue
    candidate = find_input_near_label(driver, "Адресаты")
    return candidate if candidate is not None and _usable_input(candidate) else None


def _exact_option(options, person_name):
    matches = []
    for option in options:
        try:
            if match_strict(option.text or "", person_name):
                matches.append(option)
        except Exception:
            continue
    # Не выбираем первый случайный вариант и не разрешаем неоднозначность.
    return matches[0] if len(matches) == 1 else None


def _option_click_target(driver, target):
    """Поднимается от текста до кликабельной GXT list-item, если это нужно."""
    try:
        parent = driver.execute_script(r"""
            let el = arguments[0];
            for (let level = 0; level < 8 && el; level++, el = el.parentElement) {
                const role = String(el.getAttribute && el.getAttribute('role') || '').toLowerCase();
                const cls = String(el.className || '');
                if (role === 'option' ||
                        /ListViewStyle-item|boundlist-item|combo-list-item|menu-item|select-option/i.test(cls)) {
                    return el;
                }
            }
            return null;
        """, target)
        return parent or target
    except Exception:
        return target


def add_addressee(driver, person_name, *, logger=None):
    """Выбирает адресата и возвращает True только после появления grid-row."""
    logger = logger or log
    expected = str(person_name or "").strip()
    if not expected:
        logger.error("Адресат не задан")
        return False

    if addressee_committed(driver, expected):
        logger.info("Адресат уже подтверждён в таблице")
        return True

    inp = _find_addressee_input(driver)
    if inp is None:
        logger.error("Поле адресата не найдено")
        return False

    surname = expected.split()[0]
    try:
        inp.click()
        js_type_combobox(driver, inp, surname)
    except Exception as exc:
        logger.error("Не удалось ввести поиск адресата: %s", exc)
        return False

    options = ()
    try:
        options = WebDriverWait(driver, 5).until(
            lambda d: find_dropdown_options(d, surname, inp) or False)
    except Exception:
        try:
            inp.send_keys(Keys.ENTER)
            options = WebDriverWait(driver, 3).until(
                lambda d: find_dropdown_options(d, surname, inp) or False)
        except Exception:
            options = ()

    target = _exact_option(options, expected)
    if target is None:
        logger.error("Точный адресат не найден или результат неоднозначен")
        return False

    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center',inline:'center'});", target)
    except Exception:
        pass

    click_target = _option_click_target(driver, target)
    delivered = cdp_click(driver, click_target)
    if not delivered:
        try:
            ActionChains(driver).move_to_element(click_target).pause(0.2).click().perform()
            delivered = True
        except Exception as exc:
            logger.error("Клик по адресату не отправлен: %s", exc)
            return False

    if delivered and wait_addressee_committed(driver, expected):
        logger.info("Адресат выбран и подтверждён строкой таблицы")
        return True

    logger.error("Клик по адресату отправлен, но строка в таблице не появилась")
    return False
