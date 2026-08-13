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
from selenium.common.exceptions import StaleElementReferenceException

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
    """Ждёт кликабельный элемент и кликает. Без post-sleep.

    GWT/GXT может перерисовать найденный узел прямо между ожиданием и
    ``click()``. В таком случае повторно ищем элемент по локатору. Перед JS
    fallback элемент тоже ищется заново: устаревший WebElement повторно
    использовать нельзя.
    """
    log.info(f"Ожидаю: {description or selector}")
    locator = (by, selector)
    deadline = time.monotonic() + timeout
    max_attempts = 3

    def find_fresh_clickable():
        remaining = max(0.0, deadline - time.monotonic())
        return WebDriverWait(driver, remaining).until(
            EC.element_to_be_clickable(locator)
        )

    for attempt in range(1, max_attempts + 1):
        try:
            el = find_fresh_clickable()
            try:
                el.click()
            except StaleElementReferenceException:
                raise
            except Exception:
                # Не используем handle, на котором уже упал native click.
                el = find_fresh_clickable()
                driver.execute_script("arguments[0].click();", el)

            log.info(f"Клик: {description or selector}")
            return el
        except StaleElementReferenceException:
            if attempt == max_attempts:
                raise
            log.debug(
                "Элемент устарел при клике (%s/%s), ищу заново: %s",
                attempt,
                max_attempts,
                description or selector,
            )

    raise AssertionError("unreachable")


_FIND_INPUT_JS = """
const labelText = arguments[0];
const xp = `//*[normalize-space(text())='${labelText}']`;
const snap = document.evaluate(xp, document, null,
    XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
function usable(inp) {
    if (!inp || !inp.isConnected || inp.readOnly || inp.disabled) return false;
    if ((inp.getAttribute('type') || '').toLowerCase() === 'hidden') return false;
    if ((inp.getAttribute('aria-hidden') || '').toLowerCase() === 'true') return false;
    const rect = inp.getBoundingClientRect();
    const style = window.getComputedStyle(inp);
    // GXT keeps a transparent 1x1 focus sink after a value is selected.
    // It has an offsetParent, but it is not an editable field.
    if (rect.width < 4 || rect.height < 4) return false;
    if (style.display === 'none' || style.visibility === 'hidden' ||
            parseFloat(style.opacity || '1') <= 0.05 ||
            parseInt(style.zIndex || '0', 10) < 0 ||
            style.pointerEvents === 'none') return false;
    const pointX = rect.left + Math.min(rect.width / 2, 12);
    const pointY = rect.top + rect.height / 2;
    if (pointX >= 0 && pointX < window.innerWidth &&
            pointY >= 0 && pointY < window.innerHeight) {
        const top = document.elementFromPoint(pointX, pointY);
        // An unrelated TD/button at the click point means Selenium will
        // intercept the click. Off-screen inputs are checked after scrolling.
        if (top && top !== inp && !inp.contains(top) && !top.contains(inp)) return false;
    }
    return true;
}
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
            if (usable(inp)) return inp;
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


_FIND_OPTIONS_JS = r"""
const text = arguments[0], inp = arguments[1];
const needle = String(text || '').toLocaleLowerCase('ru-RU');

function visible(el) {
    if (!el || !el.isConnected) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 &&
        style.display !== 'none' && style.visibility !== 'hidden' &&
        parseFloat(style.opacity || '1') > 0.01;
}

if (!inp || !visible(inp)) {
    return {popup_seen: false, empty_explicit: false, loading: false, options: []};
}

const ir = inp.getBoundingClientRect();
const linkedIds = String(inp.getAttribute('aria-controls') ||
    inp.getAttribute('aria-owns') || '').split(/\s+/).filter(Boolean);
const roots = [];
const rootSeen = new Set();

function popupClass(cls) {
    const value = String(cls || '');
    if (/(boundlist-item|combo-list-item|menu-item|select-option|list-item)/i.test(value)) {
        return false;
    }
    return /(?:^|[\s_-])(popup|dropdown|boundlist|combo-list|listbox|popupmenu)(?:$|[\s_-])/i.test(value) ||
        /(PopupPanel|MenuPanel|BoundList|ComboList)/i.test(value);
}

function addRoot(el, linked) {
    if (!el || rootSeen.has(el) || !visible(el) || el.contains(inp)) return;
    const r = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    const role = (el.getAttribute('role') || '').toLowerCase();
    const cls = (el.className || '').toString();
    const positioned = style.position === 'absolute' || style.position === 'fixed';
    const semantic = role === 'listbox' || role === 'menu' || popupClass(cls);
    if (!linked && !positioned && !semantic) return;

    const overlapPx = Math.max(0,
        Math.min(r.right, ir.right) - Math.max(r.left, ir.left));
    const overlapRatio = overlapPx / Math.max(1, Math.min(r.width, ir.width));
    const widthRatio = r.width / Math.max(1, ir.width);
    const belowGap = r.top - ir.bottom;
    const aboveGap = ir.top - r.bottom;
    const roomBelow = window.innerHeight - ir.bottom;
    const directBelow = belowGap >= -8 && belowGap <= 96;
    const directAbove = roomBelow < Math.max(120, Math.min(r.height, 300)) &&
        aboveGap >= -8 && aboveGap <= 96;
    const anchored = overlapRatio >= 0.55 && widthRatio >= 0.45 &&
        widthRatio <= 2.5 && r.height >= 16 && (directBelow || directAbove);
    const semanticNear = overlapRatio >= 0.35 &&
        Math.min(Math.abs(belowGap), Math.abs(aboveGap)) <= 140;
    if (!linked && !(anchored || (semantic && semanticNear))) return;
    if (!linked && r.width > window.innerWidth * 0.96 &&
            r.height > window.innerHeight * 0.85) return;

    const rootText = (el.textContent || '').toLocaleLowerCase('ru-RU');
    const hasNeedle = needle && rootText.includes(needle);
    const saysEmpty = /(ничего\s+не\s+найдено|нет\s+(данных|результат)|совпадени\w*\s+не\s+найден\w*|запис\w*\s+не\s+найден\w*|данн\w*\s+отсутству\w*|no\s+(data|result))/i.test(rootText);
    const loadingHint = el.getAttribute('aria-busy') === 'true' ||
        !!el.querySelector("[aria-busy='true'],[class*='loading'],[class*='Loading']");
    // Arbitrary page text never turns into a popup merely because it contains
    // the query. An obfuscated list is accepted only by its strict anchor
    // geometry directly below (or, when there is no room, above) the input.
    if (!linked && !semantic && (!anchored ||
            (rootText.trim() && !hasNeedle && !saysEmpty && !loadingHint))) return;
    if (!linked && !semantic &&
            el.querySelector('input,textarea,select') &&
            !hasNeedle && !saysEmpty && !loadingHint) return;
    if (!linked && semantic && rootText.trim() && !hasNeedle && !saysEmpty) return;

    let z = parseInt(style.zIndex, 10);
    if (!Number.isFinite(z)) z = 0;
    const edgeGap = directBelow ? Math.abs(belowGap) : Math.abs(aboveGap);
    const score = (linked ? 1000000 : 0) + (semantic ? 100000 : 0) +
        (directBelow ? 50000 : 0) + (directAbove ? 40000 : 0) +
        (saysEmpty ? 2000 : 0) + (hasNeedle ? 500 : 0) +
        (positioned ? 1000 : 0) + z - edgeGap;
    rootSeen.add(el);
    roots.push({el, score, emptyExplicit: saysEmpty});
}

for (const id of linkedIds) addRoot(document.getElementById(id), true);

const textNodes = document.querySelectorAll('div,span,td,li,a');
for (const node of textNodes) {
    if (!visible(node) || node === inp) continue;
    const value = (node.textContent || '').trim();
    if (!value || value.length > 500 || !value.toLocaleLowerCase('ru-RU').includes(needle)) continue;
    let el = node;
    for (let level = 0; level < 12 && el && el !== document.body;
            level++, el = el.parentElement) {
        const style = window.getComputedStyle(el);
        const role = (el.getAttribute('role') || '').toLowerCase();
        const cls = (el.className || '').toString();
        if (style.position === 'absolute' || style.position === 'fixed' ||
                role === 'listbox' || role === 'menu' || popupClass(cls)) {
            addRoot(el, false);
            break;
        }
    }
}

// Needed for linked/semantic empty lists and explicit "no results" states.
for (const el of document.querySelectorAll('div,ul,ol,table,tbody')) {
    const style = visible(el) ? window.getComputedStyle(el) : null;
    if (style && (style.position === 'absolute' || style.position === 'fixed')) {
        addRoot(el, false);
    }
}

roots.sort((a, b) => b.score - a.score);
const popup = roots.length ? roots[0].el : null;
if (!popup) {
    return {popup_seen: false, empty_explicit: false, loading: false, options: []};
}

function semanticOption(el) {
    const role = (el.getAttribute('role') || '').toLowerCase();
    const cls = (el.className || '').toString();
    return role === 'option' ||
        /(?:^|[\s_-])(option|menu-item|select-option|combo-item|boundlist-item|list-item)(?:$|[\s_-])/i.test(cls) ||
        /gxt-\w*item|x-combo-list-item|x-boundlist-item|ListItem|SelectItem/i.test(cls);
}

function optionFrom(node) {
    let el = node;
    if (semanticOption(popup)) return popup;
    const popupRect = popup.getBoundingClientRect();
    let nearestVisible = node;
    for (let level = 0; level < 10 && el && el !== popup;
            level++, el = el.parentElement) {
        if (semanticOption(el)) return el;
        const er = el.getBoundingClientRect();
        const display = window.getComputedStyle(el).display;
        const value = (el.textContent || '').trim();
        if (visible(el) && value.length <= 500 && er.height <= 160) {
            nearestVisible = el;
            // Obfuscated GXT rows have no semantic class. The nearest
            // block/table-like full-width ancestor is the row. Inline text
            // may span the same width but has no row click handler.
            if (display !== 'inline' && display !== 'contents' &&
                    er.width >= Math.max(80, popupRect.width * 0.45) &&
                    er.height >= 16) {
                return el;
            }
        }
    }
    return nearestVisible;
}

const options = [];
const optionSeen = new Set();
for (const node of popup.querySelectorAll('div,span,td,li,a')) {
    if (!visible(node)) continue;
    const value = (node.textContent || '').trim();
    if (!value || value.length > 500 || !value.toLocaleLowerCase('ru-RU').includes(needle)) continue;
    const matchingChild = Array.from(node.children).some(child =>
        visible(child) && (child.textContent || '').toLocaleLowerCase('ru-RU').includes(needle));
    if (matchingChild) continue;
    const option = optionFrom(node);
    if (!visible(option) || optionSeen.has(option)) continue;
    optionSeen.add(option);
    options.push(option);
}
const popupEntry = roots[0];
const loading = popup.getAttribute('aria-busy') === 'true' ||
    !!popup.querySelector("[aria-busy='true'],[class*='loading'],[class*='Loading']");
window.__asudPopupIds = window.__asudPopupIds || new WeakMap();
window.__asudPopupSeq = window.__asudPopupSeq || 0;
if (!window.__asudPopupIds.has(popup)) {
    window.__asudPopupIds.set(popup, `asud-popup-${++window.__asudPopupSeq}`);
}
const pr = popup.getBoundingClientRect();
let textHash = 0;
const signatureText = String(popup.textContent || '');
for (let i = 0; i < signatureText.length; i++) {
    textHash = ((textHash * 31) + signatureText.charCodeAt(i)) | 0;
}
return {
    popup_seen: true,
    empty_explicit: !!popupEntry.emptyExplicit,
    root_blank: !String(popup.textContent || '').trim(),
    loading,
    popup_key: window.__asudPopupIds.get(popup),
    signature: [Math.round(pr.left), Math.round(pr.top), Math.round(pr.width),
        Math.round(pr.height), popup.childElementCount,
        signatureText.length, textHash,
        (popup.className || '').toString()].join('|'),
    input_value: String(inp.value || '').replace(/\s+/g, ' ').trim(),
    options,
};
"""


class DropdownOptions(list):
    """Список вариантов с признаком, что popup конкретного input распознан."""

    def __init__(self, values=(), *, popup_seen=False,
                 empty_explicit=False, loading=False, popup_key=None,
                 signature=None, input_value="", root_blank=False,
                 input_observed=False):
        super().__init__(values or ())
        self.popup_seen = bool(popup_seen)
        self.empty_explicit = bool(empty_explicit)
        self.loading = bool(loading)
        self.popup_key = popup_key
        self.signature = signature
        self.input_value = str(input_value or "").strip()
        self.input_observed = bool(input_observed)
        self.root_blank = bool(root_blank)
        self.scoped = True


def find_dropdown_options(driver, text, anchor_input):
    """Возвращает только реальные option-контейнеры выпадающего списка.

    Текст документа намеренно не считается вариантом. Это особенно важно для
    feedback-писем: адрес одновременно присутствует в кратком содержании и в
    строке поиска корреспондента.
    """
    state = driver.execute_script(_FIND_OPTIONS_JS, text, anchor_input)
    if isinstance(state, dict):
        return DropdownOptions(
            state.get("options") or (),
            popup_seen=state.get("popup_seen", False),
            empty_explicit=state.get("empty_explicit", False),
            loading=state.get("loading", False),
            popup_key=state.get("popup_key"),
            signature=state.get("signature"),
            input_value=state.get("input_value", ""),
            input_observed="input_value" in state,
            root_blank=state.get("root_blank", False),
        )
    # Совместимость с тестовыми/старыми драйверами, возвращавшими только list.
    return DropdownOptions(state or (), popup_seen=bool(state))


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
