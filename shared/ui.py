"""
ui.py — Selenium UI-хелперы для АСУД (GWT/GXT).

Единый click(), find_input_near_label(), ожидания, работа с модалками.
Все паузы сохранены — GWT без них ломается.
"""

import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

log = logging.getLogger("asud.ui")


def click(driver, element, description=""):
    """Единый клик с fallback'ами для GWT/GXT.
    Порядок: ActionChains → native → JS → mouse events.

    Без post-sleep — caller сам ждёт следующий элемент через WebDriverWait.
    """
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center',inline:'center'});", element)
    except Exception:
        pass

    # ActionChains — лучше всего для GXT dropdown/autocomplete
    try:
        ActionChains(driver).move_to_element(element).pause(0.15).click().perform()
        log.info(f"Клик (mouse): {description}")
        return True
    except Exception:
        pass

    # Обычный Selenium click
    try:
        element.click()
        log.info(f"Клик (native): {description}")
        return True
    except Exception:
        pass

    # JS .click()
    try:
        driver.execute_script("arguments[0].click();", element)
        log.info(f"Клик (JS): {description}")
        return True
    except Exception:
        pass

    # Полный набор mouse-событий
    try:
        driver.execute_script("""
            var el = arguments[0];
            ['mouseover','mousedown','mouseup','click'].forEach(function(type) {
                el.dispatchEvent(new MouseEvent(type, {bubbles:true, cancelable:true, view:window}));
            });
        """, element)
        log.info(f"Клик (events): {description}")
        return True
    except Exception as e:
        log.error(f"Клик не удался: {description}: {e}")
        return False


def wait_and_click(driver, by, selector, description="", timeout=20):
    """Ждёт элемент и кликает. Без post-sleep."""
    log.info(f"Ожидаю: {description or selector}")
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, selector))
    )
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    log.info(f"Клик: {description or selector}")
    return el


_FIND_INPUT_JS = """
const labelText = arguments[0];
const xp = `//*[normalize-space(text())='${labelText}']`;
const snap = document.evaluate(xp, document, null,
    XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
for (let i = 0; i < snap.snapshotLength; i++) {
    const label = snap.snapshotItem(i);
    if (!label.offsetParent) continue;
    let parent = label;
    for (let level = 1; level <= 5; level++) {
        parent = parent.parentElement;
        if (!parent) break;
        const inputs = parent.querySelectorAll(
            "input[id*='select_combobox-input'], input[type='text']");
        for (const inp of inputs) {
            if (inp.offsetParent && !inp.readOnly) return inp;
        }
    }
}
return null;
"""


def find_input_near_label(driver, label_text):
    """Находит input combobox рядом с лейблом — один JS-вызов вместо ~25 Selenium round-trips."""
    return driver.execute_script(_FIND_INPUT_JS, label_text)


def set_driver_timeout(driver, seconds):
    """Устанавливает таймауты на работу с Edge:
      - page load (для driver.get при медленной загрузке страниц)
      - urllib3 client timeout (Selenium ↔ webdriver HTTP-запросы)

    Дефолт Selenium для urllib3 — 120с, для page_load — 300с. Этот хелпер
    унифицирует оба. Пробует несколько вариантов API из-за разных версий
    Selenium 4.x (set_page_load_timeout / command_executor._client_config).
    """
    try:
        driver.set_page_load_timeout(seconds)
    except Exception as e:
        log.debug(f"set_page_load_timeout не сработал: {e}")
    # urllib3 client timeout — пути отличаются между версиями Selenium
    for attr_path in (
        ('command_executor', '_client_config', 'timeout'),
        ('command_executor', '_conn', 'timeout'),
    ):
        try:
            obj = driver
            for a in attr_path[:-1]:
                obj = getattr(obj, a)
            setattr(obj, attr_path[-1], seconds)
            log.debug(f"urllib3 timeout установлен через {'.'.join(attr_path)}")
            break
        except Exception:
            continue


def wait_asud_loaded(driver, max_wait=120):
    """Адаптивное ожидание полной загрузки АСУД."""
    log.info("Жду загрузку АСУД...")
    try:
        WebDriverWait(driver, max_wait).until(
            lambda d: d.execute_script("return document.readyState === 'complete'"))
    except Exception:
        log.warning("readyState не complete")

    try:
        WebDriverWait(driver, max_wait).until(
            EC.element_to_be_clickable((By.ID, "mainscreen-create-button")))
    except Exception:
        log.warning("Кнопка создания не появилась")

    try:
        WebDriverWait(driver, max_wait).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR,
                "tr[class*='GridView-row'], tr[class*='grid-row'], "
                "tr[class*='OSHSGridStyle-row'], tr[class*='obj-list-rec']")) > 0)
    except Exception:
        log.warning("Данные в таблице не появились")

    log.info("АСУД загружен")


def wait_modal_closed(driver, timeout=15):
    """Ждёт пока закроется модальное окно GXT ModalPanel."""
    log.info("Жду закрытия модалки...")
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: not any(
                m.is_displayed() for m in d.find_elements(
                    By.CSS_SELECTOR, "div[class*='ModalPanel'][class*='panel']")))
        log.info("Модалка закрыта")
    except Exception:
        log.warning("Модалка не закрылась — Escape")
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(1)
        except Exception:
            pass


_DUPLICATE_MARKERS = (
    "уже зарегистрирован",
    "уже существует",
    "дубликат",
    "уже создан",
    "повторная регистрация",
)


def is_duplicate_warning(driver):
    """Проверяет показывает ли АСУД сообщение что документ уже зарегистрирован.

    Используется ПОСЛЕ Save: если за 1-2s появилась warning/toast/модалка
    с маркером 'уже зарегистрирован' — пропускаем документ без долгого
    ожидания register-кнопки которая никогда не появится.

    Возвращает True если найдено, False если нет.
    """
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text or ''
        body_lower = body_text.lower()
        return any(m in body_lower for m in _DUPLICATE_MARKERS)
    except Exception:
        return False


def close_open_modals(driver, max_escapes=5):
    """Закрывает все модалки через Escape."""
    log.info("Закрываю модалки...")
    for i in range(max_escapes):
        modals = driver.find_elements(By.CSS_SELECTOR,
            "div[class*='ModalPanel'][class*='panel']")
        visible = [m for m in modals if m.is_displayed()]
        if not visible:
            log.info(f"Модалки закрыты (попыток: {i})")
            return
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)
    log.warning(f"Не все модалки закрылись после {max_escapes} Escape")


def cdp_click(driver, element):
    """Trusted mouse-click через CDP Input.dispatchMouseEvent.

    В headless-режиме synthetic events от Selenium (ActionChains, JS click)
    имеют isTrusted=false. GXT-widget'ы (комбобоксы, FileUploadButton) могут
    отвергать такие события из соображений безопасности.

    CDP-events генерируются на уровне Chromium-движка, у них isTrusted=true —
    GXT не отличит от настоящего пользовательского клика.

    Полная mouse-цепочка: mouseMoved (hover) → mousePressed → mouseReleased.
    GXT-widget'ы часто требуют hover-state перед click чтобы пометить
    target как 'over' / 'highlighted'.

    Возвращает True если клик удался, False если CDP не доступен / упал.
    """
    try:
        # Прокручиваем в viewport — CDP кликает по viewport-абсолютным координатам
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", element)
        # Берём viewport-relative координаты через getBoundingClientRect
        # (не зависят от страничного скролла, нужны для CDP)
        coords = driver.execute_script("""
            const r = arguments[0].getBoundingClientRect();
            return {x: r.left + r.width/2, y: r.top + r.height/2};
        """, element)
        x = int(coords['x'])
        y = int(coords['y'])
        # 1. mouseMoved → GXT помечает элемент как hovered/highlighted
        driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
            "type": "mouseMoved",
            "x": x, "y": y,
            "button": "none",
        })
        # 2. mousePressed → mouseReleased — собственно click
        for evt_type in ('mousePressed', 'mouseReleased'):
            driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
                "type": evt_type,
                "x": x, "y": y,
                "button": "left",
                "clickCount": 1,
            })
        return True
    except Exception as e:
        log.debug(f"cdp_click упал: {e}")
        return False


def wait_pointer_events_auto(driver, element, timeout=2):
    """Ждёт пока CSS pointer-events на элементе станет НЕ 'none'.

    GXT в transition-анимации (например, после Save рендерит кнопку
    'Зарегистрировать') может ставить pointer-events:none на 100-500ms.
    Клик в это время визуально проходит, но handler не вызывается.

    Возвращает True если pointer-events стал auto/inherit/etc,
    False если за timeout не дождались.
    """
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        try:
            pe = driver.execute_script(
                "return window.getComputedStyle(arguments[0]).pointerEvents;", element)
            if pe and pe != 'none':
                return True
        except Exception:
            pass
        time.sleep(0.05)
    return False


def js_set_value(driver, element, value):
    """Устанавливает значение input через JS + dispatch events.
    Подходит для plain-полей (textarea, дата, номер) — без autocomplete."""
    driver.execute_script("""
        arguments[0].focus();
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
        arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
    """, element, value)


_FIND_OPTIONS_JS = """
const text = arguments[0], inp = arguments[1];
const needle = String(text || '').toLocaleLowerCase('ru-RU');
const nodes = document.querySelectorAll('div,span,td,li,a');
const out = [];
for (const r of nodes) {
    if (!r.offsetParent) continue;
    if (r === inp) continue;
    if (r.tagName === 'INPUT') continue;
    if ((r.textContent || '').length > 150) continue;
    if (!(r.textContent || '').toLocaleLowerCase('ru-RU').includes(needle)) continue;
    out.push(r);
}
return out;
"""


def find_dropdown_options(driver, text, anchor_input):
    """Возвращает видимые варианты выпадашки, где text встречается в textContent.
    Один JS-вызов вместо N+1 Selenium round-trips."""
    return driver.execute_script(_FIND_OPTIONS_JS, text, anchor_input)


def js_type_combobox(driver, element, value):
    """Печатает в combobox-autocomplete атомарно через JS.

    Для GXT/GWT-комбобоксов (корреспондент, адресат, исполнитель) —
    выпадашка фильтруется по событиям input/keyup. Просто `value=...`
    не открывает её. Здесь:
      1) фокус
      2) value = пустая строка → input/keyup (на всякий случай — сброс)
      3) value = искомая строка → input/keyup → keypress → change
    Этого хватает чтобы GXT-фильтр перерисовал список вариантов.
    """
    driver.execute_script("""
        const el = arguments[0], v = arguments[1];
        el.focus();
        el.value = '';
        el.dispatchEvent(new Event('input', {bubbles:true}));
        el.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true}));
        el.value = v;
        el.dispatchEvent(new Event('input', {bubbles:true}));
        el.dispatchEvent(new KeyboardEvent('keydown', {bubbles:true}));
        el.dispatchEvent(new KeyboardEvent('keypress', {bubbles:true}));
        el.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
    """, element, value)
