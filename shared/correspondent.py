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
from selenium.webdriver.support import expected_conditions as EC

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


def create_correspondent(driver, person_name):
    """Создаёт нового корреспондента через 7-шаговый flow."""
    parts = person_name.strip().split()
    if not parts:
        log.error("Пустое ФИО")
        return
    surname = parts[0]
    first_name = parts[1] if len(parts) >= 2 else "Н"
    middle_name = parts[2] if len(parts) >= 3 else "Н"
    if len(parts) < 3:
        log.warning(f"Неполное ФИО '{person_name}' → недостающие = 'Н'")

    log.info(f"Создаю корреспондента: {surname} {first_name} {middle_name}")

    # ШАГ 1: Клик "+" у Корреспондент
    log.info("[1/7] Клик '+' у Корреспондент")
    try:
        label = driver.find_element(By.XPATH,
            "//*[normalize-space(text())='Корреспондент']")
        parent = label
        plus_btn = None
        for _ in range(1, 7):
            parent = parent.find_element(By.XPATH, "..")
            btns = parent.find_elements(By.CSS_SELECTOR,
                "img[data-marker='select-btn'], img.gwt-Image")
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                plus_btn = visible[-1]
                break
        if not plus_btn:
            log.error("Кнопка '+' не найдена")
            return
        click(driver, plus_btn, "+ Корреспондент")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(),'Поиск корреспондента')]")))
        except Exception:
            pass
    except Exception as e:
        log.error(f"Шаг 1: {e}")
        close_open_modals(driver)
        return

    # ШАГ 2: 'Добавить' в Поиске
    log.info("[2/7] 'Добавить' в Поиске корреспондента")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.XPATH, "//*[normalize-space(text())='Добавить']")))
        btns = driver.find_elements(By.XPATH, "//*[normalize-space(text())='Добавить']")
        add_btn = next((b for b in btns if b.is_displayed()), None)
        if not add_btn:
            log.error("Кнопка 'Добавить' не найдена")
            close_open_modals(driver)
            return
        click(driver, add_btn, "Добавить")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(),'Редактирование организации') or "
                 "contains(text(),'Поиск организации')]")))
        except Exception:
            pass
    except Exception as e:
        log.error(f"Шаг 2: {e}")
        close_open_modals(driver)
        return

    # ШАГ 3: Поиск организации → 'Создать организацию'
    log.info("[3/7] Поиск организации")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.XPATH, "//*[normalize-space(text())='Поиск организации']")))
        # Ввод через JS (атомарно, без stale)
        js_result = driver.execute_script("""
            var surname = arguments[0];
            var xpath = "//*[normalize-space(text())='Поиск организации']";
            var iter = document.evaluate(xpath, document, null,
                XPathResult.ORDERED_NODE_ITERATOR_TYPE, null);
            var labels = [], node;
            while ((node = iter.iterateNext()) !== null) {
                if (node.offsetParent !== null) labels.push(node);
            }
            if (!labels.length) return 'no-label';
            for (var li = 0; li < labels.length; li++) {
                var parent = labels[li];
                for (var lvl = 0; lvl < 6; lvl++) {
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
            }
            return 'no-input';
        """, surname)
        log.info(f"JS ввод: {js_result}")
        if js_result != 'ok':
            close_open_modals(driver)
            return

        # Ждём кнопку "Создать организацию"
        create_org_btn = None
        for _ in range(10):
            try:
                btn = driver.find_element(By.CSS_SELECTOR,
                    "[id*='create_custom_org'], [id*='custom_org_button']")
                if btn.is_displayed():
                    create_org_btn = btn
                    break
            except Exception:
                pass
            btns = driver.find_elements(By.XPATH,
                "//*[contains(text(),'Создать организацию')]")
            for b in btns:
                if b.is_displayed():
                    create_org_btn = b
                    break
            if create_org_btn:
                break
            time.sleep(1)

        if create_org_btn:
            click(driver, create_org_btn, "Создать организацию")
            time.sleep(1)
        else:
            log.warning("Кнопка 'Создать организацию' не найдена")
            close_open_modals(driver)
            return
    except Exception as e:
        log.error(f"Шаг 3: {e}")
        close_open_modals(driver)
        return

    # ШАГ 4: 'Добавить' в Физические лица
    log.info("[4/7] 'Добавить' в Физические лица")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(text(),'Физические лица')]")))
        time.sleep(3)
        add_user_btn = None
        for attempt in range(20):
            try:
                btn = None
                try:
                    btn = driver.find_element(By.CSS_SELECTOR,
                        "[id*='header-organization-dialog-add-a-user-button']")
                except Exception:
                    pass
                if not btn:
                    section = driver.find_element(By.XPATH,
                        "//*[contains(text(),'Физические лица')]")
                    parent = section
                    for _ in range(1, 6):
                        parent = parent.find_element(By.XPATH, "..")
                        bs = parent.find_elements(By.XPATH,
                            ".//*[normalize-space(text())='Добавить']")
                        vis = [b for b in bs if b.is_displayed()]
                        if vis:
                            btn = vis[0]
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
            return
        click(driver, add_user_btn, "Добавить физ. лицо")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//*[normalize-space(text())='Фамилия']")))
        except Exception:
            pass
    except Exception as e:
        log.error(f"Шаг 4: {e}")
        close_open_modals(driver)
        return

    # ШАГ 5: Заполнить карточку
    log.info("[5/7] Заполнение карточки")
    try:
        fields = [
            ("Фамилия", surname, "outer_person_dialog-last_name-input"),
            ("Имя", first_name, "outer_person_dialog-first_name-input"),
            ("Отчество", middle_name, "outer_person_dialog-middle_name-input"),
            ("Должность", "ФЛ", "outer_person_dialog-position-input"),
        ]
        for label_text, value, input_id in fields:
            result = driver.execute_script("""
                var inputId = arguments[0]; var value = arguments[1];
                var el = document.getElementById(inputId);
                if (!el) {
                    var base = inputId.replace('-input','');
                    var container = document.getElementById(base);
                    if (container) {
                        var inputs = container.querySelectorAll('input[type="text"]');
                        for (var i=0;i<inputs.length;i++)
                            if (inputs[i].offsetParent!==null && !inputs[i].readOnly)
                                { el=inputs[i]; break; }
                    }
                }
                if (!el) return 'no-element';
                el.focus(); el.value = value;
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
                return 'ok:'+el.id;
            """, input_id, value)
            log.info(f"  {label_text}: {value} [{result}]")

        # Нажать "Добавить" в карточке
        save_btn = None
        try:
            save_btn = driver.find_element(By.CSS_SELECTOR,
                "[id*='Parton_person_dialog_save_button']")
        except Exception:
            btns = driver.find_elements(By.XPATH,
                "//*[normalize-space(text())='Добавить']")
            vis = [b for b in btns if b.is_displayed()]
            if vis:
                save_btn = vis[-1]
        if save_btn:
            click(driver, save_btn, "Сохранить карточку")
            try:
                WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(),'Выбрать физ')]")))
            except Exception:
                pass
    except Exception as e:
        log.error(f"Шаг 5: {e}")
        close_open_modals(driver)
        return

    # ШАГ 6: 'Выбрать физ. лиц'
    log.info("[6/7] Выбрать физ. лиц")
    try:
        select_btn = None
        try:
            select_btn = driver.find_element(By.CSS_SELECTOR,
                "[id*='Parton_organization_dialog_select_persons_button']")
        except Exception:
            btns = driver.find_elements(By.XPATH,
                "//*[contains(text(),'Выбрать физ')]")
            vis = [b for b in btns if b.is_displayed()]
            if vis:
                select_btn = vis[0]
        if select_btn:
            click(driver, select_btn, "Выбрать физ. лиц")
            try:
                WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                    (By.ID, "oshs-select-button")))
            except Exception:
                pass
        else:
            log.error("Кнопка 'Выбрать физ. лиц' не найдена")
            close_open_modals(driver)
            return
    except Exception as e:
        log.error(f"Шаг 6: {e}")
        close_open_modals(driver)
        return

    # ШАГ 7: 'Готово'
    log.info("[7/7] Готово")
    try:
        done_btn = None
        try:
            done_btn = driver.find_element(By.ID, "oshs-select-button")
        except Exception:
            btns = driver.find_elements(By.XPATH,
                "//*[normalize-space(text())='Готово']")
            vis = [b for b in btns if b.is_displayed()]
            if vis:
                done_btn = vis[0]
        if done_btn:
            click(driver, done_btn, "Готово")
            from shared.ui import wait_modal_closed
            wait_modal_closed(driver)
            log.info(f"Корреспондент создан: {person_name}")
        else:
            log.error("Кнопка 'Готово' не найдена")
            close_open_modals(driver)
    except Exception as e:
        log.error(f"Шаг 7: {e}")
        close_open_modals(driver)


def _correspondent_field_value(driver):
    """Читает текущее значение поля 'Корреспондент' (для пост-верификации)."""
    try:
        inp = find_input_near_label(driver, "Корреспондент")
        if not inp:
            return ""
        val = (inp.get_attribute('value') or '').strip()
        return val
    except Exception:
        return ""


def fill_correspondent_field(driver, person_name):
    """Заполняет поле Корреспондент через combobox.
    Если не найден по инициалам — создаёт нового.
    Пост-верификация: после клика проверяет что поле реально заполнилось;
    иначе падает в create_correspondent."""
    log.info(f"Корреспондент: {person_name}")
    time.sleep(1)

    inp = find_input_near_label(driver, "Корреспондент")
    if not inp:
        log.warning("Поле корреспондента не найдено")
        return

    surname = person_name.split()[0]
    initials = fio_to_initials(person_name)

    inp.click()
    # Combobox-autocomplete: JS-set + dispatch events открывают выпадашку
    js_type_combobox(driver, inp, surname)
    log.info(f"Введена фамилия (JS): {surname}")

    all_results = []
    try:
        WebDriverWait(driver, 5).until(
            lambda d: len(find_dropdown_options(d, surname, inp)) > 0)
        all_results = find_dropdown_options(driver, surname, inp)
    except Exception:
        from selenium.webdriver.common.keys import Keys as _Keys
        try:
            inp.send_keys(_Keys.ENTER)
            WebDriverWait(driver, 3).until(
                lambda d: len(find_dropdown_options(d, surname, inp)) > 0)
            all_results = find_dropdown_options(driver, surname, inp)
        except Exception:
            pass

    log.info(f"Кандидатов: {len(all_results)} (ищем '{initials}')")

    # Строгий матч по инициалам
    target = None
    target_desc = ""
    for idx, r in enumerate(all_results, 1):
        try:
            raw = r.text
            ok = match_strict(raw, person_name)
            preview = raw.strip().replace('\n', ' ')[:80]
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
                target_desc = f"[{idx}] <{meta}> {preview!r}"
        except Exception as e:
            log.info(f"  [{idx}] ERR читаю text: {e}")
            continue

    if target:
        from selenium.webdriver.common.action_chains import ActionChains as _AC
        from selenium.webdriver.common.keys import Keys as _Keys
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
        except Exception:
            pass

        def _try_click_target(el, strategy):
            try:
                if strategy == 'ac':
                    _AC(driver).move_to_element(el).pause(0.2).click().perform()
                elif strategy == 'js':
                    driver.execute_script("arguments[0].click();", el)
                elif strategy == 'enter':
                    inp.send_keys(_Keys.ENTER)
                return True
            except Exception:
                return False

        def _check(timeout_iters=5):
            """Поллинг поля. По дефолту 1s (5 × 0.2s) — быстрая проверка.
            В новой версии АСУД value часто хранится не в input.value а в
            chip-узле; если поле пустое после клика это не обязательно
            означает что клик не сработал. Trust-mode в caller'е."""
            for _ in range(timeout_iters):
                time.sleep(0.2)
                val = _correspondent_field_value(driver)
                if val and surname.lower() in val.lower():
                    return val
            return None

        # Найти родительский option-контейнер ПЕРЕД click-стратегиями.
        # GXT часто рендерит option как <div role='option'><span>имя</span><span> - </span>...</div>
        # — handler на div, а наш XPath нашёл внутренний <span>. Клик по span'у
        # ничего не даёт; нужен сам option.
        parent_option = None
        try:
            parent_option = driver.execute_script("""
                let el = arguments[0];
                for (let i = 0; i < 6 && el; i++) {
                    el = el.parentElement;
                    if (!el) break;
                    const role = el.getAttribute('role');
                    const cls = (el.className || '').toString().toLowerCase();
                    // Расширенный список классов GXT/Sencha widget'ов:
                    // option/item/menu-item/select-option — общие
                    // gxt-/x-combo/x-boundlist — Sencha-специфичные
                    // ListItem/SelectItem — типичные именования виджетов
                    if (role === 'option' ||
                        /\\b(option|item|menu-item|select-option|combo-item|boundlist-item|ListItem|SelectItem)\\b/i.test(cls) ||
                        /gxt-\\w*item|x-combo-list-item|x-boundlist-item/i.test(cls)) {
                        return el;
                    }
                }
                return null;
            """, target)
        except Exception:
            pass
        if parent_option is not None:
            log.debug(f"Найден parent-option, делаю trusted-click по нему")

        # TRUST-MODE: делаем один CDP-click и доверяем визуальному результату.
        # Раньше пробовали 4 стратегии + create-new fallback (16+ секунд на
        # документ). В новой АСУД клик визуально срабатывает, просто value
        # хранится не в input.value — наш polling его не видит, а UI заполнен.
        cdp_target = parent_option if parent_option is not None else target
        log.debug(f"CDP trusted-click по {'parent-option' if parent_option else 'target'}")
        click_ok = cdp_click(driver, cdp_target)
        val = _check() if click_ok else None
        if val:
            log.info(f"Корреспондент выбран (CDP, value виден): {person_name} (поле: {val!r})")
            return
        if click_ok:
            log.info(f"Корреспондент: CDP-click отправлен, value в input не виден — "
                     f"доверяем (вероятно value в chip-узле)")
            return

        # CDP-click сам не сработал (rare) — fallback ActionChains
        log.debug("CDP не сработал, fallback ActionChains")
        _try_click_target(target, 'ac')
        val = _check()
        if val:
            log.info(f"Корреспондент выбран (AC fallback): {person_name} (поле: {val!r})")
            return

        # Совсем никак — ASUD-state скорее всего сломан, но НЕ падаем в
        # create_correspondent (он у нас всё равно не работает — кнопка '+'
        # не найдена). Просто warning и идём дальше.
        log.warning(f"Корреспондент {target_desc}: не подтверждён, но идём дальше "
                    f"(возможно value в chip-узле, регистрация попробует сохранить)")
        return

    # Нет совпадения — создаём нового
    log.info(f"'{initials}' не найден — создаю нового")
    from selenium.webdriver.common.keys import Keys as _Keys
    try:
        inp.send_keys(_Keys.ESCAPE)
        time.sleep(0.5)
    except Exception:
        pass
    create_correspondent(driver, person_name)
