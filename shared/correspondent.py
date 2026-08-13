"""
correspondent.py — Создание нового корреспондента в АСУД.

7 шагов:
  1. Клик '+' у Корреспондент
  2. 'Добавить' в Поиске
  3. Поиск организации → 'Создать организацию'
  4. 'Добавить' в Физические лица
  5. Заполнить карточку (ФИО + Должность=ФЛ)
  6. 'Выбрать физ. лиц'
  7. 'Готово'
"""

import re
import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait

from shared.ui import (click, find_input_near_label, close_open_modals,
                js_type_combobox, find_dropdown_options, cdp_click)

log = logging.getLogger("asud.correspondent")


# Маркеры, после которых следует ФИО корреспондента в теле письма
FIO_MARKER_RE = re.compile(
    r'(?iu)(?:ф\.?\s*и\.?\s*о|фио)\.?\s*'
    r'(?:абонента|заявителя|обратившегося|обращающегося)?\s*:?\s*'
    r'([А-ЯЁ][а-яёА-ЯЁ\s\.\-]{5,100}?)'
    r'(?=\s*(?:\n|$|;|,|\(|\-\s|ЛС|Email|E\-mail|Почт|Телефон|Адрес|Контакт))'
)

# Loose-fallback: 3 заглавных слова подряд с русским отчеством
FIO_LOOSE_RE = re.compile(
    r'\b([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)\s+'
    r'([А-ЯЁ][а-яё]*(?:ович|евич|ич|овна|евна|ична|инична))\b'
)


def _clean_fio(raw):
    """Нормализует извлечённое ФИО: лишние пробелы, NBSP, точки инициалов, хвосты."""
    if not raw:
        return None
    s = raw.replace('\xa0', ' ').strip(' \t\r\n.,;:-')
    s = re.sub(r'\s+', ' ', s)
    # Оставляем только первые 3 слова, если их больше (обрежем хвост "и контакты" и пр.)
    parts = s.split()
    if len(parts) > 3:
        parts = parts[:3]
    s = ' '.join(parts)
    # Должно начинаться с заглавной кириллицы и содержать >=2 слов
    if len(parts) < 2:
        return None
    if not re.match(r'^[А-ЯЁ]', s):
        return None
    return s


def extract_fio_from_text(text):
    """Извлекает ФИО корреспондента из TextBody.

    Возвращает (fio, source) где source ∈ {'marker','loose',None}.
    Если не найдено — (None, None).
    """
    if not text:
        return None, None
    t = str(text).replace('_x000D_', '\n')
    # 1) По маркеру "Ф.И.О. абонента: ..."
    m = FIO_MARKER_RE.search(t)
    if m:
        fio = _clean_fio(m.group(1))
        if fio:
            return fio, 'marker'
    # 2) Loose: 3 слова с русским отчеством — берём первое
    m = FIO_LOOSE_RE.search(t)
    if m:
        fio = _clean_fio(f"{m.group(1)} {m.group(2)} {m.group(3)}")
        if fio:
            return fio, 'loose'
    return None, None


def fio_to_initials(full_name):
    """'Калганова Тамара Алексеевна' → 'Калганова Т А'"""
    parts = full_name.strip().split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]} {parts[2][0]}"
    elif len(parts) == 2:
        return f"{parts[0]} {parts[1][0]}"
    return parts[0] if parts else full_name


def _norm_no_space(s):
    return s.replace('.', '').replace(',', '').replace(' ', '').replace('\xa0', '').lower()


def _norm_keep_space(s):
    import re as _re
    s = s.replace('\xa0', ' ').replace('.', '').replace(',', '')
    return _re.sub(r'\s+', ' ', s).strip().lower()


def match_correspondent(text, full_name):
    """Мягкий матч: полное ФИО / инициалы / фамилия. Для адресатов."""
    text_clean = text.strip()
    if full_name in text_clean:
        return True
    initials = fio_to_initials(full_name)
    if _norm_keep_space(initials) in _norm_keep_space(text_clean):
        return True
    if _norm_no_space(initials) in _norm_no_space(text_clean):
        return True
    if _norm_no_space(full_name) in _norm_no_space(text_clean):
        return True
    surname = full_name.split()[0]
    if text_clean.lower().startswith(surname.lower()):
        return True
    return False


def match_strict(text, full_name):
    """Строгий матч: только полное ФИО или инициалы. Для корреспондентов.
    Нормализует любые пробелы (включая NBSP) и точки/запятые."""
    text_clean = text.strip()
    if full_name in text_clean:
        return True
    if _norm_keep_space(full_name) in _norm_keep_space(text_clean):
        return True
    if _norm_no_space(full_name) in _norm_no_space(text_clean):
        return True
    initials = fio_to_initials(full_name)
    if _norm_keep_space(initials) in _norm_keep_space(text_clean):
        return True
    if _norm_no_space(initials) in _norm_no_space(text_clean):
        return True
    return False


def _is_legal_kind(kind):
    return str(kind or "").strip().casefold() in {
        "legal", "organization", "org", "юл", "юрлицо", "юридическое лицо",
    }


def _is_address_kind(kind):
    return str(kind or "").strip().casefold() in {
        "address", "feedback-address", "feedback_address", "адрес",
    }


def correspondent_card_parts(name, kind="person"):
    """Возвращает значения полей Фамилия/Имя/Отчество для новой карточки."""
    value = " ".join(str(name or "").split())
    if not value:
        return "", "", ""
    if _is_address_kind(kind):
        return value, "", ""
    if _is_legal_kind(kind):
        return value, "-", "-"
    parts = value.split()
    surname = parts[0]
    first_name = parts[1] if len(parts) >= 2 else "Н"
    middle_name = parts[2] if len(parts) >= 3 else "Н"
    return surname, first_name, middle_name


def match_legal_correspondent(text, legal_name):
    """Строго сверяет полное название ЮЛ, игнорируя пунктуацию и пробелы."""
    expected = _norm_no_space(str(legal_name or ""))
    actual = _norm_no_space(str(text or ""))
    return bool(expected) and expected in actual


_FIND_CORRESPONDENT_ADD_BUTTON_JS = r"""
const inp = arguments[0];
if (!inp || !inp.isConnected) return null;
const ir = inp.getBoundingClientRect();
const inputY = ir.top + ir.height / 2;
const candidates = [];
let root = inp;
for (let level = 0; level < 8 && root; level++, root = root.parentElement) {
    const nodes = root.querySelectorAll(
        "[data-marker='select-btn'], button, [role='button'], img.gwt-Image, img"
    );
    for (const el of nodes) {
        if (!el.isConnected) continue;
        const r = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        if (r.width <= 0 || r.height <= 0 || style.display === 'none' ||
                style.visibility === 'hidden') continue;
        const cy = r.top + r.height / 2;
        if (Math.abs(cy - inputY) > Math.max(35, ir.height * 1.75)) continue;
        if (r.right < ir.left - 15 || r.left > ir.right + 180) continue;

        const marker = (el.getAttribute('data-marker') || '').toLowerCase();
        const hint = [el.getAttribute('title'), el.getAttribute('aria-label'),
            el.getAttribute('alt'), el.textContent].filter(Boolean).join(' ').toLowerCase();
        let score = 0;
        if (marker === 'select-btn') score += 1000;
        if (/(добав|выбр|созд|select|add|choose|\+)/i.test(hint)) score += 250;
        if (r.left >= ir.right - 25) score += 100;
        score -= Math.abs(r.left - ir.right);
        score -= level * 5;
        candidates.push({el, score});
    }
}
candidates.sort((a, b) => b.score - a.score);
return candidates.length && candidates[0].score > -100 ? candidates[0].el : null;
"""


def _find_correspondent_add_button(driver, input_element):
    """Находит кнопку выбора/добавления в той же строке, что и поле."""
    try:
        return driver.execute_script(
            _FIND_CORRESPONDENT_ADD_BUTTON_JS, input_element)
    except Exception:
        return None


_DIALOG_ANCESTOR_XPATH = (
    "./ancestor::*[@role='dialog' or @aria-modal='true' or "
    "contains(@class,'ModalPanel') or contains(@class,'DialogBox') or "
    "contains(@class,'dialog') or contains(@class,'window')]"
)

_ELEMENT_LAYER_INFO_JS = r"""
const el = arguments[0];
if (!el || !el.isConnected) return {known: true, exposed: false, score: -1};
const r = el.getBoundingClientRect();
if (r.width <= 0 || r.height <= 0) return {known: true, exposed: false, score: -1};
const x = Math.max(0, Math.min(window.innerWidth - 1, r.left + r.width / 2));
const y = Math.max(0, Math.min(window.innerHeight - 1, r.top + r.height / 2));
const hit = document.elementFromPoint(x, y);
const exposed = !!hit && (hit === el || el.contains(hit) || hit.contains(el));
let z = 0, dialogDepth = 0, node = el;
for (let depth = 0; node && depth < 20; depth++, node = node.parentElement) {
    const style = window.getComputedStyle(node);
    const zi = parseInt(style.zIndex, 10);
    if (Number.isFinite(zi)) z = Math.max(z, zi);
    const role = (node.getAttribute('role') || '').toLowerCase();
    const cls = (node.className || '').toString();
    if (role === 'dialog' || node.getAttribute('aria-modal') === 'true' ||
            /(ModalPanel|DialogBox|dialog|window|popup)/i.test(cls)) {
        dialogDepth = Math.max(dialogDepth, 20 - depth);
    }
}
return {known: true, exposed, score: z * 100 + dialogDepth};
"""


def _pick_topmost_visible(driver, elements):
    """Выбирает видимый элемент верхней модалки, а не контрол под overlay."""
    visible = []
    for index, element in enumerate(elements or ()):
        try:
            if element.is_displayed():
                visible.append((index, element))
        except Exception:
            continue
    if not visible:
        return False

    ranked = []
    layer_info_known = False
    for index, element in visible:
        try:
            info = driver.execute_script(_ELEMENT_LAYER_INFO_JS, element) or {}
            if isinstance(info, dict) and info.get("known"):
                layer_info_known = True
                if info.get("exposed"):
                    ranked.append((float(info.get("score", 0)), index, element))
                continue
        except Exception:
            pass

        # Browserless fallback and compatibility with old Selenium: prefer a
        # candidate that belongs to a dialog over a visible page control.
        try:
            dialogs = element.find_elements(By.XPATH, _DIALOG_ANCESTOR_XPATH)
        except Exception:
            dialogs = []
        ranked.append((1000 if dialogs else 0, index, element))

    if layer_info_known and not ranked:
        return False
    if not ranked:
        return visible[0][1]
    ranked.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return ranked[0][2]


def _wait_visible_xpath(driver, xpath, timeout=15):
    """Ждёт видимый элемент на верхнем UI-слое/в активной модалке."""
    return WebDriverWait(driver, timeout).until(
        lambda d: _pick_topmost_visible(
            d, d.find_elements(By.XPATH, xpath))
    )


def _find_visible(driver, by, selector):
    """Ищет контрол по locator, предпочитая верхнюю модалку."""
    search_supported = callable(getattr(driver, "find_elements", None))
    try:
        if not search_supported:
            raise AttributeError("find_elements unavailable")
        found = _pick_topmost_visible(driver, driver.find_elements(by, selector))
        if found:
            return found
        # Production Selenium search completed and proved that every displayed
        # candidate is covered by a higher UI layer. Never bypass that verdict
        # with find_element(), which would return the underlay control again.
        return None
    except Exception:
        if search_supported:
            return None
    # Некоторые тестовые/старые драйверы реализуют только find_element.
    try:
        element = driver.find_element(by, selector)
        return element if element.is_displayed() else None
    except Exception:
        return None


def create_correspondent(driver, person_name, kind="person"):
    """Создаёт карточку; привязка к документу проверяется вызывающим кодом."""
    legal = _is_legal_kind(kind)
    address_only = _is_address_kind(kind)
    surname, first_name, middle_name = correspondent_card_parts(
        person_name, kind)
    if not surname:
        log.error("Пустое имя корреспондента")
        return False
    if not legal and not address_only and len(person_name.strip().split()) < 3:
        log.warning(f"Неполное ФИО '{person_name}' → недостающие = 'Н'")

    kind_label = "ЮЛ" if legal else ("АДРЕС" if address_only else "ФЛ")
    shown_card_name = ("<адрес>" if address_only else
                       f"{surname} {first_name} {middle_name}")
    log.info(f"Создаю корреспондента ({kind_label}): {shown_card_name}")

    # ШАГ 1: Клик "+" у Корреспондент
    log.info("[1/7] Клик '+' у Корреспондент")
    try:
        corr_input = find_input_near_label(driver, "Корреспондент")
        plus_btn = (_find_correspondent_add_button(driver, corr_input)
                    if corr_input else None)
        if not plus_btn:
            log.error("Кнопка '+' у поля Корреспондент не найдена")
            return False
        if not click(driver, plus_btn, "+ Корреспондент"):
            log.error("Кнопка '+' у поля Корреспондент не нажалась")
            return False
        _wait_visible_xpath(
            driver, "//*[contains(normalize-space(text()),'Поиск корреспондента')]", 15)
    except Exception as e:
        log.error(f"Шаг 1: {e}")
        close_open_modals(driver)
        return False

    # ШАГ 2: 'Добавить' в Поиске
    log.info("[2/7] 'Добавить' в Поиске корреспондента")
    try:
        add_btn = _wait_visible_xpath(
            driver, "//*[normalize-space(text())='Добавить']", 15)
        if not add_btn:
            log.error("Кнопка 'Добавить' не найдена")
            close_open_modals(driver)
            return False
        if not click(driver, add_btn, "Добавить"):
            log.error("Кнопка 'Добавить' не нажалась")
            close_open_modals(driver)
            return False
        _wait_visible_xpath(
            driver,
            "//*[contains(normalize-space(text()),'Редактирование организации') or "
            "contains(normalize-space(text()),'Поиск организации')]",
            15,
        )
    except Exception as e:
        log.error(f"Шаг 2: {e}")
        close_open_modals(driver)
        return False

    # ШАГ 3: Поиск организации → 'Создать организацию'
    log.info("[3/7] Поиск организации")
    try:
        org_heading = _wait_visible_xpath(
            driver, "//*[normalize-space(text())='Поиск организации']", 15)
        # Ввод через JS (атомарно, без stale)
        js_result = driver.execute_script("""
            var heading = arguments[0], surname = arguments[1];
            if (!heading || !heading.isConnected) return 'no-label';
            var parent = heading;
            for (var lvl = 0; lvl < 8; lvl++) {
                parent = parent.parentElement;
                if (!parent) break;
                var inputs = parent.querySelectorAll('input[type="text"]');
                for (var i = 0; i < inputs.length; i++) {
                    var inp = inputs[i];
                    if (inp.offsetParent !== null && !inp.readOnly) {
                        inp.focus(); inp.value = surname;
                        inp.dispatchEvent(new Event('input', {bubbles:true}));
                        inp.dispatchEvent(new Event('keyup', {bubbles:true}));
                        inp.dispatchEvent(new Event('change', {bubbles:true}));
                        return 'ok';
                    }
                }
            }
            return 'no-input';
        """, org_heading, surname)
        log.info(f"JS ввод: {js_result}")
        if js_result != 'ok':
            log.error("Поле 'Поиск организации' не найдено")
            close_open_modals(driver)
            return False

        # Ждём кнопку "Создать организацию"
        create_org_btn = None
        for _ in range(10):
            create_org_btn = _find_visible(
                driver,
                By.CSS_SELECTOR,
                "[id*='create_custom_org'], [id*='custom_org_button']",
            )
            if not create_org_btn:
                create_org_btn = _pick_topmost_visible(
                    driver,
                    driver.find_elements(
                        By.XPATH, "//*[contains(text(),'Создать организацию')]"),
                )
            if create_org_btn:
                break
            time.sleep(1)

        if create_org_btn:
            if not click(driver, create_org_btn, "Создать организацию"):
                log.error("Кнопка 'Создать организацию' не нажалась")
                close_open_modals(driver)
                return False
            time.sleep(1)
        else:
            log.warning("Кнопка 'Создать организацию' не найдена")
            close_open_modals(driver)
            return False
    except Exception as e:
        log.error(f"Шаг 3: {e}")
        close_open_modals(driver)
        return False

    # ШАГ 4: 'Добавить' в Физические лица
    log.info("[4/7] 'Добавить' в Физические лица")
    try:
        _wait_visible_xpath(
            driver, "//*[contains(text(),'Физические лица')]", 15)
        time.sleep(3)
        add_user_btn = None
        for attempt in range(20):
            try:
                btn = _find_visible(
                    driver,
                    By.CSS_SELECTOR,
                    "[id*='header-organization-dialog-add-a-user-button']",
                )
                if not btn:
                    section = _wait_visible_xpath(
                        driver, "//*[contains(text(),'Физические лица')]", 2)
                    parent = section
                    for _ in range(1, 6):
                        parent = parent.find_element(By.XPATH, "..")
                        bs = parent.find_elements(By.XPATH,
                            ".//*[normalize-space(text())='Добавить']")
                        btn = _pick_topmost_visible(driver, bs)
                        if btn:
                            break
                if btn:
                    is_enabled = driver.execute_script("""
                        var el = arguments[0];
                        if (!el.offsetParent) return false;
                        if (el.getAttribute('aria-disabled')==='true') return false;
                        if (el.classList.contains('x-disabled')) return false;
                        var style = window.getComputedStyle(el);
                        if (style.pointerEvents==='none') return false;
                        if (parseFloat(style.opacity)<0.5) return false;
                        return true;
                    """, btn)
                    if is_enabled:
                        add_user_btn = btn
                        break
            except Exception:
                pass
            time.sleep(1)
        if not add_user_btn:
            log.error("Кнопка 'Добавить' не активировалась")
            close_open_modals(driver)
            return False
        if not click(driver, add_user_btn, "Добавить физ. лицо"):
            log.error("Кнопка 'Добавить физ. лицо' не нажалась")
            close_open_modals(driver)
            return False
        person_heading = _wait_visible_xpath(
            driver, "//*[normalize-space(text())='Фамилия']", 15)
    except Exception as e:
        log.error(f"Шаг 4: {e}")
        close_open_modals(driver)
        return False

    # ШАГ 5: Заполнить карточку
    log.info("[5/7] Заполнение карточки")
    try:
        fields = [
            ("Фамилия", surname, "outer_person_dialog-last_name-input"),
            ("Имя", first_name, "outer_person_dialog-first_name-input"),
            ("Отчество", middle_name, "outer_person_dialog-middle_name-input"),
            ("Должность", "ЮЛ" if legal else "ФЛ", "outer_person_dialog-position-input"),
        ]
        for label_text, value, input_id in fields:
            field_element = _find_visible(driver, By.ID, input_id)
            if not field_element:
                log.error(f"Поле '{label_text}' карточки не найдено")
                close_open_modals(driver)
                return False
            result = driver.execute_script("""
                var el = arguments[0]; var value = arguments[1];
                if (!el || !el.isConnected || el.offsetParent === null) return 'no-element';
                el.focus(); el.value = value;
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
                return 'ok:'+el.id;
            """, field_element, value)
            shown_value = ("<пусто>" if not value else
                           ("<адрес>" if address_only and label_text == "Фамилия"
                            else value))
            log.info(f"  {label_text}: {shown_value} [{result}]")
            if not str(result or "").startswith("ok:"):
                log.error(f"Поле '{label_text}' карточки не заполнилось")
                close_open_modals(driver)
                return False

        # Нажать "Добавить" в карточке
        save_btn = _find_visible(
            driver,
            By.CSS_SELECTOR,
            "[id*='Parton_person_dialog_save_button']",
        )
        if not save_btn:
            save_btn = _pick_topmost_visible(
                driver,
                driver.find_elements(
                    By.XPATH, "//*[normalize-space(text())='Добавить']"),
            )
        if not save_btn:
            log.error("Кнопка сохранения карточки корреспондента не найдена")
            close_open_modals(driver)
            return False
        if not click(driver, save_btn, "Сохранить карточку"):
            log.error("Карточка корреспондента не сохранилась")
            close_open_modals(driver)
            return False
        _wait_visible_xpath(driver, "//*[contains(text(),'Выбрать физ')]", 15)
    except Exception as e:
        log.error(f"Шаг 5: {e}")
        close_open_modals(driver)
        return False

    # ШАГ 6: 'Выбрать физ. лиц'
    log.info("[6/7] Выбрать физ. лиц")
    try:
        select_btn = _find_visible(
            driver,
            By.CSS_SELECTOR,
            "[id*='Parton_organization_dialog_select_persons_button']",
        )
        if not select_btn:
            select_btn = _pick_topmost_visible(
                driver,
                driver.find_elements(By.XPATH, "//*[contains(text(),'Выбрать физ')]"),
            )
        if not select_btn:
            log.error("Кнопка 'Выбрать физ. лиц' не найдена")
            close_open_modals(driver)
            return False
        if not click(driver, select_btn, "Выбрать физ. лиц"):
            log.error("Кнопка 'Выбрать физ. лиц' не нажалась")
            close_open_modals(driver)
            return False
        WebDriverWait(driver, 15).until(
            lambda d: _find_visible(d, By.ID, "oshs-select-button"))
    except Exception as e:
        log.error(f"Шаг 6: {e}")
        close_open_modals(driver)
        return False

    # ШАГ 7: 'Готово'
    log.info("[7/7] Готово")
    try:
        done_btn = _find_visible(driver, By.ID, "oshs-select-button")
        if not done_btn:
            done_btn = _pick_topmost_visible(
                driver,
                driver.find_elements(
                    By.XPATH, "//*[normalize-space(text())='Готово']"),
            )
        if not done_btn:
            log.error("Кнопка 'Готово' не найдена")
            close_open_modals(driver)
            return False
        if not click(driver, done_btn, "Готово"):
            log.error("Кнопка 'Готово' не нажалась")
            close_open_modals(driver)
            return False
        from shared.ui import wait_modal_closed
        wait_modal_closed(driver)
        selected = _wait_for_correspondent_value(
            driver, person_name, kind=kind, timeout=5,
            allow_closed_input=True, baseline_input="")
        if selected:
            log.info(f"Корреспондент создан и выбран: {shown_card_name}")
            return True
        # Живой АСУД сохраняет карточку после шага 5, но не всегда переносит
        # созданное физлицо обратно в исходный combobox после «Готово».
        # Это не ошибка создания: вызывающий код повторно найдёт уже созданную
        # запись в справочнике и выберет её, не запуская создание второй раз.
        log.warning(
            "Карточка корреспондента создана, но не подставилась в документ — "
            "выберу созданную запись из справочника"
        )
        return True
    except Exception as e:
        log.error(f"Шаг 7: {e}")
        close_open_modals(driver)
        return False


_READ_CORRESPONDENT_FIELD_JS = r"""
const inp = arguments[0];
if (!inp || !inp.isConnected) {
    return {input_value: '', popup_visible: false, semantic_values: []};
}
const ir = inp.getBoundingClientRect();
const values = [];
const seen = new Set();

function add(value) {
    const text = String(value || '').replace(/\s+/g, ' ').trim();
    if (!text || text === 'Корреспондент' || text.length > 500 || seen.has(text)) return;
    seen.add(text);
    values.push(text);
}

function visible(el) {
    if (!el || !el.isConnected) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.display !== 'none' &&
        style.visibility !== 'hidden' && parseFloat(style.opacity || '1') > 0.01;
}

function explicitPopupAncestor(el) {
    let node = el;
    for (let level = 0; level < 10 && node; level++, node = node.parentElement) {
        const role = (node.getAttribute('role') || '').toLowerCase();
        const cls = (node.className || '').toString();
        if (role === 'option' || role === 'listbox' || role === 'menu' ||
                /(popup|dropdown|boundlist|combo-list|listbox|menu-item|select-option)/i.test(cls)) {
            return true;
        }
    }
    return false;
}

function popupVisible() {
    if ((inp.getAttribute('aria-expanded') || '').toLowerCase() === 'true') return true;
    const linked = String(inp.getAttribute('aria-controls') ||
        inp.getAttribute('aria-owns') || '').split(/\s+/).filter(Boolean);
    for (const id of linked) {
        if (visible(document.getElementById(id))) return true;
    }
    for (const el of document.querySelectorAll(
            "[role='listbox'],[role='menu'],div,ul,ol,table")) {
        if (!visible(el) || el.contains(inp)) continue;
        const style = window.getComputedStyle(el);
        const cls = (el.className || '').toString();
        const role = (el.getAttribute('role') || '').toLowerCase();
        const itemClass = /(boundlist-item|combo-list-item|menu-item|select-option|list-item)/i.test(cls);
        const popupClass = !itemClass &&
            (/(?:^|[\s_-])(popup|dropdown|boundlist|combo-list|listbox|popupmenu)(?:$|[\s_-])/i.test(cls) ||
             /(PopupPanel|MenuPanel|BoundList|ComboList)/i.test(cls));
        const positioned = style.position === 'absolute' || style.position === 'fixed';
        const semantic = role === 'listbox' || role === 'menu' || popupClass;
        if (!positioned && !semantic) continue;
        const r = el.getBoundingClientRect();
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
        const semanticNear = semantic && overlapRatio >= 0.35 &&
            Math.min(Math.abs(belowGap), Math.abs(aboveGap)) <= 140;
        const rootText = String(el.textContent || '').replace(/\s+/g, ' ').trim();
        const query = String(inp.value || '').replace(/\s+/g, ' ').trim();
        const containsQuery = !!query && rootText.toLocaleLowerCase('ru-RU').includes(
            query.toLocaleLowerCase('ru-RU'));
        const saysEmpty = /(ничего\s+не\s+найдено|нет\s+(данных|результат)|совпадени\w*\s+не\s+найден\w*|запис\w*\s+не\s+найден\w*|данн\w*\s+отсутству\w*|no\s+(data|result))/i.test(rootText);
        const loadingHint = el.getAttribute('aria-busy') === 'true' ||
            !!el.querySelector("[aria-busy='true'],[class*='loading'],[class*='Loading']");
        const nonsemanticPopupShape = anchored &&
            (!rootText || containsQuery || saysEmpty || loadingHint) &&
            (!el.querySelector('input,textarea,select') ||
             containsQuery || saysEmpty || loadingHint);
        if ((nonsemanticPopupShape || semanticNear) &&
                !(r.width > window.innerWidth * 0.96 && r.height > window.innerHeight * 0.85)) {
            return true;
        }
    }
    return false;
}

// inp.value является либо поисковой строкой открытой выпадашки, либо итоговым
// display-value после её закрытия. Python принимает его только во втором случае.
for (const attr of ['data-value', 'data-display-value', 'data-selected-value']) {
    add(inp.getAttribute(attr));
}

let root = inp.parentElement;
for (let level = 0; level < 5 && root; level++, root = root.parentElement) {
    for (const el of root.querySelectorAll(
            "input, span, div, td, [data-value], [data-display-value]")) {
        if (el === inp || !el.isConnected) continue;
        if (explicitPopupAncestor(el)) continue;
        if (el.tagName === 'INPUT') {
            const type = (el.getAttribute('type') || '').toLowerCase();
            const inputCls = (el.className || '').toString();
            if (el.readOnly || el.disabled || type === 'hidden' ||
                    /(chip|token|selected|display)/i.test(inputCls + ' ' + el.id)) {
                add(el.value);
            }
        }
        for (const attr of ['data-value', 'data-display-value', 'data-selected-value']) {
            add(el.getAttribute(attr));
        }
        const r = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        if (r.width <= 0 || r.height <= 0 || style.display === 'none' ||
                style.visibility === 'hidden') continue;
        const sameRow = r.bottom >= ir.top - 20 && r.top <= ir.bottom + 20;
        const nearby = r.right >= ir.left - 80 && r.left <= ir.right + 220;
        if (!sameRow || !nearby) continue;
        const semanticHint = [el.className, el.id, el.getAttribute('role'),
            el.getAttribute('data-marker')].filter(Boolean).join(' ');
        if (/(chip|token|selected|selection|display-value|selected-value)/i.test(semanticHint)) {
            add(el.textContent);
        }
    }
}
return {
    input_value: String(inp.value || '').replace(/\s+/g, ' ').trim(),
    popup_visible: popupVisible(),
    semantic_values: values,
};
"""


def _correspondent_field_state(driver):
    """Возвращает model-bound state поля и текущее display-value."""
    try:
        inp = find_input_near_label(driver, "Корреспондент")
        if not inp:
            return {
                "input_value": "", "popup_visible": False,
                "semantic_values": [], "semantic_value": "",
            }
        state = driver.execute_script(_READ_CORRESPONDENT_FIELD_JS, inp)
        if isinstance(state, dict):
            semantic = " | ".join(
                str(value).strip()
                for value in (state.get("semantic_values") or ())
                if str(value).strip()
            )
            state = dict(state)
            state["semantic_value"] = semantic
            return state
        # Backward compatibility for a legacy/mock driver.
        value = str(state or "").strip()
        return {
            "input_value": value, "popup_visible": False,
            "semantic_values": [value] if value else [],
            "semantic_value": value,
        }
    except Exception:
        return {
            "input_value": "", "popup_visible": False,
            "semantic_values": [], "semantic_value": "",
        }


def _correspondent_field_value(driver):
    """Читает только подтверждённое model-bound значение поля."""
    return _correspondent_field_state(driver).get("semantic_value", "")


def _correspondent_value_matches(value, expected, kind="person"):
    if not value:
        return False
    if _is_legal_kind(kind) or _is_address_kind(kind):
        return match_legal_correspondent(value, expected)
    return match_strict(value, expected)


def _wait_for_correspondent_value(driver, expected, kind="person", timeout=3,
                                  allow_closed_input=False,
                                  baseline_input=None):
    """Поллит локальный state и возвращает подтверждённое значение.

    ``input.value`` принимается только после клика по доказанному option и
    закрытия его popup. В остальных местах поисковая строка не является
    доказательством выбора модели.
    """
    attempts = max(1, int(timeout / 0.2))
    for _ in range(attempts):
        state = _correspondent_field_state(driver)
        # An open popup can contain hover/selected-looking option classes;
        # nothing is accepted until that popup has actually closed.
        if state.get("popup_visible") is True:
            time.sleep(0.2)
            continue
        # Keep the narrow value helper as a compatibility/test hook.
        value = _correspondent_field_value(driver)
        if _correspondent_value_matches(value, expected, kind):
            return value
        if allow_closed_input and state.get("popup_visible") is False:
            value = str(state.get("input_value") or "").strip()
            baseline_changed = (
                baseline_input is not None and
                _norm_no_space(value) != _norm_no_space(str(baseline_input))
            )
            if (baseline_changed and
                    _correspondent_value_matches(value, expected, kind)):
                return value
        time.sleep(0.2)
    return ""


_OPTION_ANCESTOR_JS = r"""
let el = arguments[0];
let genericItem = null;
for (let level = 0; level < 8 && el; level++, el = el.parentElement) {
    const role = (el.getAttribute('role') || '').toLowerCase();
    const cls = (el.className || '').toString();
    if (role === 'option' ||
        /(?:^|[\s_-])(option|menu-item|select-option|combo-item|boundlist-item|list-item)(?:$|[\s_-])/i.test(cls) ||
        /gxt-\w*item|x-combo-list-item|x-boundlist-item|ListItem|SelectItem/i.test(cls)) {
        return el;
    }
    if (!genericItem && /(?:^|[\s_-])item(?:$|[\s_-])/i.test(cls)) {
        genericItem = el;
    }
    if (genericItem && (role === 'listbox' || role === 'menu' ||
            /(popup|dropdown|boundlist|combo-list|menu|listbox)/i.test(cls))) {
        return genericItem;
    }
}
return null;
"""

def _option_ancestor(driver, element):
    """Возвращает реальный option-контейнер или None для текста карточки."""
    try:
        return driver.execute_script(_OPTION_ANCESTOR_JS, element)
    except Exception:
        return None


_CLEAR_QUERY_BEFORE_OPTION_CLICK_JS = r"""
const inp = arguments[0], option = arguments[1];
if (!inp || !inp.isConnected || !option || !option.isConnected) return false;
const r = option.getBoundingClientRect();
const style = window.getComputedStyle(option);
if (r.width <= 0 || r.height <= 0 || style.display === 'none' ||
        style.visibility === 'hidden') return false;
// Do not dispatch input/change: the already rendered option and GXT's query
// model stay intact. A real option selection writes the display value back;
// merely closing/redrawing the popup leaves this raw field empty.
const descriptor = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype, 'value');
if (descriptor && descriptor.set) descriptor.set.call(inp, '');
else inp.value = '';
return inp.value === '' && option.isConnected;
"""


def _clear_query_before_option_click(driver, input_element, option_element):
    """Creates a before/after signal without re-filtering the open popup."""
    try:
        return bool(driver.execute_script(
            _CLEAR_QUERY_BEFORE_OPTION_CLICK_JS,
            input_element,
            option_element,
        ))
    except Exception:
        return False


def fill_correspondent_field(driver, person_name, kind="person", *, allow_create=True):
    """Выбирает существующего или создаёт нового корреспондента.

    Возвращает True только когда значение исходного поля подтверждено.
    """
    person_name = " ".join(str(person_name or "").split())
    if not person_name:
        log.error("Пустое значение корреспондента")
        return False

    legal = _is_legal_kind(kind)
    address_only = _is_address_kind(kind)
    kind_label = "ЮЛ" if legal else ("АДРЕС" if address_only else "ФЛ")
    shown_name = "<адрес>" if address_only else person_name
    log.info(f"Корреспондент ({kind_label}): {shown_name}")
    time.sleep(1)

    inp = find_input_near_label(driver, "Корреспондент")
    if not inp:
        log.warning("Поле корреспондента не найдено")
        return False

    exact_name = legal or address_only
    search_value = person_name if exact_name else person_name.split()[0]
    initials = person_name if exact_name else fio_to_initials(person_name)

    inp.click()
    # Combobox-autocomplete: JS-set + dispatch events открывают выпадашку
    js_type_combobox(driver, inp, search_value)
    log.info(f"Введён поиск (JS): {shown_name}")

    all_results = None
    lookup_error = None
    empty_since = None
    empty_samples = 0
    empty_key = None
    empty_signature = None
    result_since = None
    result_samples = 0
    result_key = None
    result_signature = None
    lookup_completed = False
    lookup_timeout = 10 if address_only else 5
    lookup_started = time.monotonic()

    def _lookup_ready(d):
        nonlocal all_results, empty_since, empty_samples, empty_key
        nonlocal empty_signature, result_since, result_samples, result_key
        nonlocal result_signature, lookup_completed
        all_results = find_dropdown_options(d, search_value, inp)
        popup_seen = getattr(all_results, "popup_seen", None)
        if popup_seen is not None:
            now = time.monotonic()
            popup_key = getattr(all_results, "popup_key", None)
            signature = getattr(all_results, "signature", None)
            reported_input = getattr(all_results, "input_value", "")
            input_observed = getattr(all_results, "input_observed", False)
            input_ok = (not input_observed or
                        _norm_keep_space(reported_input) ==
                        _norm_keep_space(search_value))
            if not popup_seen:
                empty_since = None
                empty_samples = 0
                empty_key = None
                empty_signature = None
                result_since = None
                result_samples = 0
                result_key = None
                result_signature = None
                return False
            if getattr(all_results, "loading", False):
                empty_since = None
                empty_samples = 0
                empty_key = None
                empty_signature = None
                result_since = None
                result_samples = 0
                result_key = None
                result_signature = None
                return False
            if not input_ok:
                empty_since = None
                empty_samples = 0
                empty_key = None
                empty_signature = None
                result_since = None
                result_samples = 0
                result_key = None
                result_signature = None
                return False
            if all_results:
                empty_since = None
                empty_samples = 0
                state_key = popup_key or signature
                if (state_key != result_key or
                        signature != result_signature):
                    result_key = state_key
                    result_signature = signature
                    result_since = now
                    result_samples = 1
                else:
                    result_samples += 1
                # A freshly opened GXT list may briefly contain rows for the
                # clear/previous query. Require two identical observations.
                if popup_key is None and signature is None:
                    lookup_completed = True  # browserless/legacy tests
                elif (result_samples >= 2 and result_since is not None and
                      now - result_since >= 0.8):
                    lookup_completed = True
                return lookup_completed

            result_since = None
            result_samples = 0
            state_key = popup_key or signature
            if state_key != empty_key or signature != empty_signature:
                empty_key = state_key
                empty_signature = signature
                empty_since = now
                empty_samples = 1
            else:
                empty_samples += 1

            if getattr(all_results, "empty_explicit", False):
                if popup_key is None and signature is None:
                    lookup_completed = True  # browserless/legacy tests
                elif (empty_samples >= 2 and empty_since is not None and
                      now - empty_since >= 0.8):
                    lookup_completed = True
                return lookup_completed

            # A blank, strictly anchored popup is a valid zero-result state
            # only after it stayed unchanged for practically the whole lookup
            # window. This permits a new feedback address while preventing a
            # transient network/loading blank from creating a duplicate.
            if (getattr(all_results, "root_blank", False) and
                    state_key is not None and empty_samples >= 4 and
                    empty_since is not None and now - empty_since >= 2.0 and
                    now - lookup_started >= lookup_timeout - 0.75):
                lookup_completed = True
                return True
            return False
        # Unit/legacy compatibility: a successful empty list means a completed
        # empty lookup; non-empty unscoped nodes still require validation below.
        lookup_completed = True
        return True

    try:
        WebDriverWait(driver, lookup_timeout).until(_lookup_ready)
    except Exception as exc:
        lookup_error = exc
        try:
            # One final sample may cross the stability threshold at timeout;
            # it must pass the same rules as every earlier poll.
            _lookup_ready(driver)
        except Exception as final_exc:
            lookup_error = final_exc
            all_results = None

    if all_results is None:
        log.error(f"Не удалось определить popup корреспондента: {lookup_error}")
        return False

    if not lookup_completed:
        log.error(
            "Поиск корреспондента не завершился (popup отсутствует/загружается) — "
            "создание отменено во избежание дубля"
        )
        return False

    popup_seen = getattr(all_results, "popup_seen", None)
    legacy_lookup = popup_seen is None
    scoped_options = bool(getattr(all_results, "scoped", False))
    if popup_seen is None:
        # Plain list exists only in compatibility/tests. Empty is an explicit
        # completed result; matching page nodes are an unknown unsafe state.
        popup_seen = not bool(all_results)

    if not popup_seen and not all_results:
        log.error("Popup корреспондента не распознан — создание отменено во избежание дубля")
        return False

    # Защитный фильтр поверх find_dropdown_options: текст карточки документа
    # не является вариантом, даже если содержит тот же адрес/ФИО.
    option_results = []
    seen_options = set()
    for result in all_results:
        option = result if scoped_options else _option_ancestor(driver, result)
        if option is None:
            continue
        key = getattr(option, "id", None) or id(option)
        if key in seen_options:
            continue
        seen_options.add(key)
        option_results.append((result, option))

    log.info(f"Вариантов справочника: {len(option_results)}")

    if legacy_lookup and option_results:
        # A real option ancestor is sufficient evidence for legacy list-only
        # drivers; production uses the explicit popup_seen state.
        popup_seen = True

    if all_results and not option_results:
        log.error(
            "Popup не подтверждён: совпадения оказались элементами карточки, "
            "создание отменено во избежание дубля"
        )
        return False

    if not popup_seen:
        log.error("Popup корреспондента не подтверждён — создание отменено")
        return False

    # Строгий матч по инициалам
    target = None
    target_option = None
    target_desc = ""
    option_read_error = False
    for idx, (r, option) in enumerate(option_results, 1):
        try:
            raw = (getattr(option, "text", None) or r.text or "")
            ok = (match_legal_correspondent(raw, person_name) if exact_name
                  else match_strict(raw, person_name))
            preview = ("<адрес скрыт>" if address_only
                       else raw.strip().replace('\n', ' ')[:80])
            # Логируем tag + class для диагностики что именно за элемент
            try:
                tag = r.tag_name
                cls = (r.get_attribute('class') or '')[:40]
                meta = f"{tag}.{cls}" if cls else tag
            except Exception:
                meta = "?"
            log.info(f"  [{idx}] {'OK' if ok else '--'} <{meta}> | {preview!r}")
            if ok and target is None:
                target = r
                target_option = option
                target_desc = f"[{idx}] <{meta}> {preview!r}"
        except Exception as e:
            option_read_error = True
            log.info(f"  [{idx}] ERR читаю text: {e}")
            continue

    if target:
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", target_option)
        except Exception:
            pass

        raw_query_cleared = _clear_query_before_option_click(
            driver, inp, target_option)
        click_ok = cdp_click(driver, target_option)
        selected = (_wait_for_correspondent_value(
            driver, person_name, kind=kind, timeout=3,
            allow_closed_input=raw_query_cleared,
            baseline_input="" if raw_query_cleared else None) if click_ok else "")
        if selected:
            log.info(f"Корреспондент выбран из справочника: {shown_name}")
            return True

        # Повторный клик по тому же реальному option — страховка для CDP.
        try:
            ActionChains(driver).move_to_element(
                target_option).pause(0.2).click().perform()
        except Exception:
            pass
        selected = _wait_for_correspondent_value(
            driver, person_name, kind=kind, timeout=2,
            allow_closed_input=raw_query_cleared,
            baseline_input="" if raw_query_cleared else None)
        if selected:
            log.info(f"Корреспондент выбран из справочника (повтор): {shown_name}")
            return True

        # Запись уже существует: создание новой породит дубль. Останавливаем
        # документ и даём daemon повторить его после восстановления страницы.
        log.error(
            f"Корреспондент {target_desc} найден, но выбор не подтверждён")
        return False

    if option_read_error:
        log.error(
            "Список корреспондентов перерисовался во время чтения — "
            "создание отменено во избежание дубля"
        )
        return False

    # Нет совпадения — создаём нового. После успешного создания этот же
    # поиск вызывается повторно с allow_create=False: так мы привязываем
    # карточку к документу, но никогда не создаём дубль при задержке индекса.
    if not allow_create:
        log.warning("Созданный корреспондент не найден в этой попытке выбора")
        return False

    shown_initials = "<адрес>" if address_only else initials
    log.info(f"'{shown_initials}' не найден — создаю нового")
    from selenium.webdriver.common.keys import Keys as _Keys
    input_cleared = False
    try:
        inp.send_keys(_Keys.ESCAPE)
        time.sleep(0.5)
        input_cleared = bool(driver.execute_script("""
            const el = arguments[0];
            if (el && el.isConnected) {
                el.value = '';
                el.dispatchEvent(new Event('input', {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                return el.value === '';
            }
            return false;
        """, inp))
    except Exception:
        input_cleared = False
    if not input_cleared:
        log.error("Поле Корреспондент не очистилось перед созданием")
        return False
    created = create_correspondent(driver, person_name, kind=kind)
    if not created:
        log.error("Новый корреспондент не создан или не выбран")
        return False
    selected = _wait_for_correspondent_value(
        driver, person_name, kind=kind, timeout=2,
        allow_closed_input=True, baseline_input="")
    if not selected:
        for attempt in range(1, 3):
            log.info(
                "Ищу созданного корреспондента и привязываю к документу "
                f"(попытка {attempt}/2)"
            )
            if fill_correspondent_field(
                    driver, person_name, kind=kind, allow_create=False):
                return True
            if attempt < 2:
                time.sleep(1)
        log.error(
            "Созданный корреспондент пока не появился в справочнике — "
            "повторите документ позже: новая карточка уже существует"
        )
        return False
    return True
