"""
attachments.py — Поиск и прикрепление .msg файлов.

Ищет файл по Link из Excel в outlook_dir (рекурсивно по подпапкам),
прикрепляет через CDP file chooser intercept + send_keys на скрытый input
GXT FileUploadButton. Работает в headless-режиме, не требует Windows GUI.
"""

import os
import re
import time
import logging
from datetime import datetime, date
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from shared.ui import click, wait_modal_closed, close_open_modals

log = logging.getLogger("asud.attach")


def _link_to_variants(link):
    """Генерирует все возможные имена файлов для link.
    Форматы: DD.MM.YYYY / YYYY-MM-DD, с/без ведущего нуля, суффиксы _1-_9."""
    variants = set()

    if isinstance(link, (datetime, date)):
        try:
            variants.add(link.strftime("%d.%m.%Y %H-%M-%S"))
            variants.add(link.strftime("%Y-%m-%d %H-%M-%S"))
            with_lead = link.strftime("%d.%m.%Y %H-%M-%S")
            m = re.match(r'^(\d{2}\.\d{2}\.\d{4})\s+0(\d)-(\d{2})-(\d{2})$', with_lead)
            if m:
                variants.add(f"{m.group(1)} {m.group(2)}-{m.group(3)}-{m.group(4)}")
        except Exception:
            pass
    else:
        s = str(link).strip()
        if s:
            variants.add(s)
            # DD.MM.YYYY → ISO (+ с/без ведущего нуля в часе)
            m = re.match(r'^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{1,2})-(\d{2})-(\d{2})$', s)
            if m:
                dd, mm, yyyy, h, mn, sec = m.groups()
                h_lead = f"{int(h):02d}"
                h_no = str(int(h))
                variants.add(f"{dd}.{mm}.{yyyy} {h_lead}-{mn}-{sec}")
                variants.add(f"{dd}.{mm}.{yyyy} {h_no}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_lead}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_no}-{mn}-{sec}")
            # ISO → DD.MM.YYYY
            m = re.match(r'^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2})-(\d{2})-(\d{2})$', s)
            if m:
                yyyy, mm, dd, h, mn, sec = m.groups()
                h_lead = f"{int(h):02d}"
                h_no = str(int(h))
                variants.add(f"{dd}.{mm}.{yyyy} {h_lead}-{mn}-{sec}")
                variants.add(f"{dd}.{mm}.{yyyy} {h_no}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_lead}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_no}-{mn}-{sec}")

    variants.discard('')
    return list(variants)


def find_msg_by_link(link, outlook_dir, fallback_path=None):
    """Ищет .msg файл в outlook_dir (рекурсивно по подпапкам) по link.
    Fallback → пустышка."""
    log.info(f"link={link!r} (тип: {type(link).__name__})")

    if link is None or (isinstance(link, str) and not link.strip()):
        log.warning("link пустой — пустышка")
        return fallback_path

    if not os.path.isdir(outlook_dir):
        log.warning(f"Папка {outlook_dir} не найдена — пустышка")
        return fallback_path

    variants = _link_to_variants(link)
    log.info(f"Варианты: {variants}")

    # Один проход по дереву — собираем все .msg
    all_msg = []  # [(full_path, filename_no_ext)]
    try:
        for root, _dirs, files in os.walk(outlook_dir):
            for f in files:
                if f.lower().endswith('.msg'):
                    name_no_ext = os.path.splitext(f)[0]
                    all_msg.append((os.path.join(root, f), name_no_ext))
    except Exception as e:
        log.error(f"Ошибка обхода {outlook_dir}: {e}")
        return fallback_path

    if not all_msg:
        log.warning(f"В {outlook_dir} нет .msg файлов")
        return fallback_path

    def _rel(full):
        try:
            return os.path.relpath(full, outlook_dir)
        except Exception:
            return os.path.basename(full)

    # Фаза 1: точное совпадение с вариантом
    variants_set = set(variants)
    for full, name in all_msg:
        if name in variants_set:
            log.info(f"Нашёл: {_rel(full)}")
            return full

    # Фаза 2: variant + _1.._9
    suffix_set = {f"{v}_{i}" for v in variants for i in range(1, 10)}
    for full, name in all_msg:
        if name in suffix_set:
            log.info(f"Нашёл (суффикс): {_rel(full)}")
            return full

    # Фаза 3: подстрока
    for full, name in all_msg:
        for v in variants:
            if v in name:
                log.info(f"Нашёл (подстрока): {_rel(full)}")
                return full

    log.warning("Файл не найден — пустышка")
    return fallback_path


def get_dummy_msg(base_dir):
    """Ищет .msg-пустышку рядом с exe."""
    msg_files = [f for f in os.listdir(base_dir) if f.lower().endswith('.msg')]
    if msg_files:
        return os.path.join(base_dir, msg_files[0])
    return None


def move_to_done(file_path, outlook_dir, done_dirname="Завершено"):
    """Перемещает обработанный .msg в <outlook_dir>/Завершено/.
    Папка создаётся если нет. Конфликт имён → суффикс _HHMMSS.
    Ничего не делает если file_path пустой/не существует.
    """
    if not file_path or not os.path.isfile(file_path):
        return
    if not outlook_dir or not os.path.isdir(outlook_dir):
        log.warning(f"outlook_dir '{outlook_dir}' не существует — "
                    f"не перемещаю {os.path.basename(file_path)}")
        return

    done_dir = os.path.join(outlook_dir, done_dirname)
    try:
        os.makedirs(done_dir, exist_ok=True)
    except Exception as e:
        log.warning(f"Не удалось создать {done_dir}: {e}")
        return

    name = os.path.basename(file_path)
    dest = os.path.join(done_dir, name)
    if os.path.exists(dest):
        base, ext = os.path.splitext(name)
        ts = datetime.now().strftime("%H%M%S")
        dest = os.path.join(done_dir, f"{base}_{ts}{ext}")

    try:
        import shutil
        shutil.move(file_path, dest)
        log.info(f"→ Завершено/{os.path.basename(dest)}")
    except Exception as e:
        log.warning(f"Не удалось переместить {name}: {e}")


def move_to_drafts(file_path, base_dir, drafts_dirname="Черновики"):
    """Перемещает .msg в <base_dir>/Черновики/ — для daemon-режима, чтобы
    не подхватывать снова в следующей итерации. Юзер при необходимости
    вернёт обратно после ручной правки.
    """
    if not file_path or not os.path.isfile(file_path):
        return
    if not base_dir or not os.path.isdir(base_dir):
        return
    drafts_dir = os.path.join(base_dir, drafts_dirname)
    try:
        os.makedirs(drafts_dir, exist_ok=True)
    except Exception as e:
        log.warning(f"Не удалось создать {drafts_dir}: {e}")
        return
    name = os.path.basename(file_path)
    dest = os.path.join(drafts_dir, name)
    if os.path.exists(dest):
        base, ext = os.path.splitext(name)
        ts = datetime.now().strftime("%H%M%S")
        dest = os.path.join(drafts_dir, f"{base}_{ts}{ext}")
    try:
        import shutil
        shutil.move(file_path, dest)
        log.info(f"→ Черновики/{os.path.basename(dest)}")
    except Exception as e:
        log.warning(f"Не удалось переместить {name} в Черновики: {e}")


def move_to_errors(file_path, base_dir, reason="", err_dirname="Ошибки"):
    """Перемещает .msg в <base_dir>/Ошибки/ и пишет рядом .txt с причиной.
    Папка создаётся если нет. Конфликт имён → суффикс _HHMMSS.
    Ничего не делает если file_path пустой/не существует.
    """
    if not file_path or not os.path.isfile(file_path):
        return
    if not base_dir or not os.path.isdir(base_dir):
        log.warning(f"base_dir '{base_dir}' не существует — "
                    f"не перемещаю {os.path.basename(file_path)}")
        return

    err_dir = os.path.join(base_dir, err_dirname)
    try:
        os.makedirs(err_dir, exist_ok=True)
    except Exception as e:
        log.warning(f"Не удалось создать {err_dir}: {e}")
        return

    name = os.path.basename(file_path)
    dest = os.path.join(err_dir, name)
    if os.path.exists(dest):
        base, ext = os.path.splitext(name)
        ts = datetime.now().strftime("%H%M%S")
        dest = os.path.join(err_dir, f"{base}_{ts}{ext}")

    try:
        import shutil
        shutil.move(file_path, dest)
        log.info(f"→ Ошибки/{os.path.basename(dest)}")
        if reason:
            sidecar = os.path.splitext(dest)[0] + ".txt"
            try:
                with open(sidecar, "w", encoding="utf-8") as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n{reason}\n")
            except Exception as e:
                log.debug(f"sidecar не записан: {e}")
    except Exception as e:
        log.warning(f"Не удалось переместить {name} в Ошибки: {e}")


_DIAG_BUTTONS_JS = r"""
// Диагностика: вернуть видимые кандидаты на confirm-кнопку
const out = [];
for (const el of document.querySelectorAll(
        "[id*='SetContent'], [id*='Send'], [id*='Submit'], button, div, span")) {
    if (!el.offsetParent) continue;
    const t = (el.textContent || '').trim();
    if (t.length > 80) continue;
    if (!/(Присоединить|Подтвердить|Сохранить|Загрузить)/i.test(t)
        && !/(SetContent|Send|Submit|Save)/i.test(el.id || '')) continue;
    out.push((el.id || el.tagName) + '|' + t.slice(0, 60));
}
return out;
"""


_DIAG_FILE_INPUTS_JS = r"""
// Диагностика: переписать все <input type="file"> в DOM
// со всеми атрибутами и контекстом. Используется при таймауте
// _attach_via_input для понимания что в DOM пошло не так.
const out = [];
for (const inp of document.querySelectorAll("input[type='file']")) {
    const r = inp.getBoundingClientRect();
    const cs = window.getComputedStyle(inp);
    const parentTag = inp.parentElement ? inp.parentElement.tagName : null;
    const parentCls = inp.parentElement ? (inp.parentElement.className || '').toString().slice(0, 80) : null;
    out.push({
        id: inp.id || null,
        name: inp.name || null,
        accept: inp.getAttribute('accept'),
        multiple: inp.multiple,
        disabled: inp.disabled,
        offsetParent: !!inp.offsetParent,  // true если отображается
        rect: {w: r.width, h: r.height, x: r.left, y: r.top},
        display: cs.display,
        visibility: cs.visibility,
        opacity: cs.opacity,
        position: cs.position,
        parent_tag: parentTag,
        parent_class: parentCls,
        outerHTML_preview: inp.outerHTML.slice(0, 200),
    });
}
return {count: out.length, inputs: out};
"""


def _diag_file_inputs(driver, label):
    """Дампит в лог состояние всех <input type='file'> на странице."""
    try:
        diag = driver.execute_script(_DIAG_FILE_INPUTS_JS) or {}
        log.info(f"[diag {label}] input[type=file] count={diag.get('count', 0)}")
        for i, inp in enumerate(diag.get('inputs', []) or []):
            log.info(f"  [diag {label}] input[{i}]: id={inp.get('id')!r} name={inp.get('name')!r} "
                     f"disabled={inp.get('disabled')} offsetParent={inp.get('offsetParent')} "
                     f"display={inp.get('display')} visibility={inp.get('visibility')} "
                     f"opacity={inp.get('opacity')}")
            log.info(f"  [diag {label}]    rect={inp.get('rect')} parent={inp.get('parent_tag')}.{inp.get('parent_class')}")
            log.info(f"  [diag {label}]    html: {inp.get('outerHTML_preview')!r}")
    except Exception as e:
        log.warning(f"[diag {label}] dump упал: {e}")


def _attach_via_input(driver, file_path):
    """Стратегия через CDP file chooser intercept + send_keys на скрытый input.

    Найдено диагностикой стадии 0:
      • АСУД использует GXT FileUploadButton — скрытый <input type='file'> с
        name='SetContentDialog' создаётся через ~0.5-0.8s после клика
        'Присоединить содержимое'.
      • input невидим (display:none, rect 0×0), но functional.

    Шаги:
      1. Включаем Page.setInterceptFileChooserDialog → native picker не открывается
      2. Кликаем 'Присоединить содержимое' → input создаётся в DOM, picker
         перехвачен CDP'ом и не показан
      3. Ждём появления input[name='SetContentDialog'] (до 3s)
      4. Делаем input visible через JS (чтобы send_keys прошёл visibility-check)
      5. input.send_keys(file_path) — Selenium через WebDriver-protocol устанавливает
         значение файлового input'а без участия native picker
      6. Выключаем intercept
      7. Ждём confirm-кнопку, кликаем

    Возвращает True если успех, False если что-то упало (caller сделает retry).
    """
    intercept_enabled = False
    try:
        # 1. Включаем перехват native file picker
        try:
            driver.execute_cdp_cmd("Page.setInterceptFileChooserDialog",
                                    {"enabled": True})
            intercept_enabled = True
            log.debug("CDP file chooser intercept включён")
        except Exception as e:
            log.warning(f"CDP intercept не доступен: {e}")
            return False

        # 2. Клик по 'Присоединить содержимое' — picker не откроется (intercept)
        try:
            btn = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH,
                    "//div[contains(text(),'Присоединить содержимое')]")))
            click(driver, btn, "Присоединить содержимое (CDP)")
        except Exception as e:
            log.error(f"Кнопка 'Присоединить содержимое' не найдена: {e}")
            return False

        # 3. Ждём появления input[name='SetContentDialog']
        end = time.monotonic() + 3
        inp = None
        while time.monotonic() < end:
            try:
                inp = driver.find_element(By.CSS_SELECTOR,
                    "input[type='file'][name='SetContentDialog']")
                break
            except Exception:
                pass
            time.sleep(0.1)
        if not inp:
            log.warning("input[name='SetContentDialog'] не появился за 3s")
            return False
        log.debug(f"input найден: id={inp.get_attribute('id')}")

        # 4. Делаем visible — иначе send_keys жалуется на visibility
        driver.execute_script("""
            var el = arguments[0];
            el.style.display='block';
            el.style.visibility='visible';
            el.style.opacity='1';
            el.style.position='absolute';
            el.style.left='0px';
            el.style.top='0px';
            el.style.width='1px';
            el.style.height='1px';
            el.removeAttribute('hidden');
            el.disabled = false;
        """, inp)

        # 5. send_keys устанавливает значение файлового input напрямую через
        # WebDriver-protocol (без участия native picker)
        inp.send_keys(file_path)

        # 5а. ВАЖНО: дождаться пока input.files реально содержит файл.
        # Без этого confirm-клик случается до того как upload дошёл, модалка
        # остаётся открытой, файл по факту не прикреплён.
        # 25с (раньше было 10) — в новой версии АСУД иногда долго отрисовывается.
        try:
            WebDriverWait(driver, 25, poll_frequency=0.1).until(
                lambda d: d.execute_script(
                    "return arguments[0].files && arguments[0].files.length > 0;",
                    inp))
            log.debug("input.files подтверждён")
        except Exception:
            log.warning("input.files пустой за 25s — upload скорее всего не дошёл")
            return False

        # Триггерим change на случай если автоматический dispatch не сработал
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", inp)

        # 5б. Даём АСУД 4s на серверный upload. На некоторых документах
        # 1.5s было мало — upload не успевал, confirm кликался впустую,
        # ждали 30s+ retry-cycle. 4s покрывает ~95% случаев одной попыткой.
        time.sleep(4)

        log.info(f"Файл отправлен через CDP-intercept + send_keys")
        return True

    except Exception as e:
        log.warning(f"CDP-стратегия упала: {e}")
        return False
    # NB: intercept НЕ выключаем здесь — caller (attach_content) сам отключит
    # после завершения всего flow включая confirm-click. Иначе на retry
    # _wait_confirm_and_click может случайно нажать главную кнопку
    # 'Присоединить содержимое' (имеет тот же текст что confirm), и АСУД
    # откроет native picker потому что intercept выключен.


def _wait_confirm_and_click(driver, timeout=30):
    """Поллинг + клик confirm-кнопки 'Присоединить' с retry если модалка не закрылась.

    На большинстве документов АСУД-server upload занимает >5s, и confirm-кнопка
    в это время хоть и видима, но имеет pointer-events:none или data-disabled='1'
    (как register-кнопка после Save). Кликаем только когда:
      • кнопка реально clickable (pointer-events != 'none', data-disabled != '1')
    Логика:
      1. Найти confirm-кнопку
      2. Подождать пока она станет clickable (poll 200ms, до 15s)
      3. Кликнуть
      4. Подождать 3s что модалка закрылась
      5. Если ещё видна — снова ждать clickable + клик. До 5 попыток.
    """
    end = time.time() + timeout

    def _find_btn():
        # Только confirm-кнопки С НЕПУСТЫМ id-маркером SetContent/Send/Submit.
        # Без id или с пустым id — это может быть главная кнопка 'Присоединить
        # содержимое' на карточке документа, кликать её — открывать picker (если
        # бы intercept был выкл) или по крайней мере бессмысленно.
        try:
            btns = driver.find_elements(By.CSS_SELECTOR,
                "[id*='SetContent'], [id*='Send'], [id*='Submit']")
            for b in btns:
                if b.is_displayed() and b.get_attribute('id'):
                    return b
        except Exception:
            pass
        # Fallback: text-search ВНУТРИ открытой модалки (не на карточке)
        try:
            modals = driver.find_elements(By.CSS_SELECTOR,
                "div[class*='ModalPanel'][class*='panel']")
            for modal in modals:
                if not modal.is_displayed():
                    continue
                btns = modal.find_elements(By.XPATH,
                    ".//button[contains(text(),'Присоединить')] | .//div[contains(text(),'Присоединить')]")
                for b in btns:
                    if b.is_displayed():
                        return b
        except Exception:
            pass
        return None

    def _is_btn_clickable(btn):
        """clickable = visible + pointer-events != 'none' + data-disabled != '1' + не в class*='disabled'"""
        try:
            if not btn.is_displayed():
                return False
            data_dis = btn.get_attribute('data-disabled')
            aria_dis = btn.get_attribute('aria-disabled')
            cls = (btn.get_attribute('class') or '').lower()
            if data_dis == '1' or aria_dis == 'true' or 'disabled' in cls:
                return False
            pe = driver.execute_script(
                "return window.getComputedStyle(arguments[0]).pointerEvents;", btn)
            if pe == 'none':
                return False
            return True
        except Exception:
            return False

    def _wait_clickable(btn, max_wait):
        """Ждём пока кнопка станет clickable. Возвращает True/False."""
        deadline = time.time() + max_wait
        last_state = None
        while time.time() < deadline:
            if _is_btn_clickable(btn):
                return True
            try:
                state = (
                    f"d={btn.get_attribute('data-disabled')} "
                    f"a={btn.get_attribute('aria-disabled')} "
                    f"pe={driver.execute_script('return window.getComputedStyle(arguments[0]).pointerEvents;', btn)}"
                )
                if state != last_state:
                    log.debug(f"  confirm не clickable: {state}")
                    last_state = state
            except Exception:
                pass
            time.sleep(0.2)
        return False

    def _modal_closed():
        return _find_btn() is None

    attempts = 0
    while time.time() < end and attempts < 5:
        btn = _find_btn()
        if not btn:
            log.info("Файл присоединён!")
            return True

        # Дожидаемся пока кнопка реально станет clickable (upload завершился)
        log.debug(f"Жду пока confirm станет clickable (try #{attempts + 1})")
        if not _wait_clickable(btn, max_wait=15):
            log.warning(f"Confirm-кнопка не стала clickable за 15s (try #{attempts + 1})")
            attempts += 1
            time.sleep(1)
            continue

        attempts += 1
        click(driver, btn, f"confirm by-id {btn.get_attribute('id')} (try #{attempts})")

        # Ждём 3s что модалка закрылась
        sub_end = time.time() + 3
        while time.time() < sub_end:
            if _modal_closed():
                log.info("Файл присоединён!")
                return True
            time.sleep(0.2)

        log.debug(f"Модалка ещё открыта после try #{attempts}, жду 2s и retry")
        time.sleep(2)

    log.warning(f"Confirm не сработал за {attempts} попыток — диагностика:")
    try:
        candidates = driver.execute_script(_DIAG_BUTTONS_JS) or []
        log.warning(f"  Видимых кандидатов: {len(candidates)}")
        for c in candidates[:10]:
            log.warning(f"    - {c}")
    except Exception:
        pass
    return False


def attach_content(driver, file_path, max_attempts=3):
    """Прикрепляет файл к карточке документа через CDP intercept + send_keys.

    Возвращает True если файл успешно приложен, False если все retry'и
    провалились. Caller'ы могут использовать результат для лога в
    высокоуровневом логгере 'asud' (виден в консоли).

    До max_attempts попыток. Между попытками закрываем висящие модалки
    и даём АСУД пару секунд отдохнуть.

    CDP intercept ВКЛЮЧАЕТСЯ внутри _attach_via_input и не выключается там —
    выключается ТОЛЬКО в finally этой функции, после всех retry'ев и confirm-click'ов.
    Иначе случайный клик на похожую кнопку 'Присоединить содержимое' между
    попытками открыл бы native picker.
    """
    log.info(f"Прикрепление: {os.path.basename(file_path)}")

    try:
        for attempt in range(1, max_attempts + 1):
            if attempt > 1:
                # Проверка: может быть первый attempt по факту УЖЕ прикрепил файл,
                # просто input.files не отрисовался вовремя? Если confirm-кнопка
                # SetContent/Send уже видна — значит АСУД принял файл, повторный
                # _attach_via_input приведёт к ДУБЛИРОВАНИЮ вложения.
                try:
                    confirm_btns = driver.find_elements(By.CSS_SELECTOR,
                        "[id*='SetContent'], [id*='Send'], [id*='Submit']")
                    has_visible_confirm = any(
                        b.is_displayed() and b.get_attribute('id')
                        for b in confirm_btns)
                    if has_visible_confirm:
                        log.info(f"Перед попыткой {attempt}: confirm-кнопка уже видна — "
                                 f"файл по факту приложен. Пропускаю _attach_via_input, "
                                 f"иду на confirm.")
                        if _wait_confirm_and_click(driver, timeout=20):
                            return True
                        # confirm не сработал — закрываем и пробуем сначала
                        log.warning("confirm не сработал даже когда кнопка была видна — закрываю всё")
                except Exception:
                    pass
                log.warning(f"Прикрепление: попытка {attempt}/{max_attempts}")
                try:
                    close_open_modals(driver)
                except Exception:
                    pass
                time.sleep(2)

            if _attach_via_input(driver, file_path):
                if _wait_confirm_and_click(driver, timeout=20):
                    return True  # успех
                log.warning(f"Попытка {attempt}: файл загружен но confirm не сработал")
            else:
                log.warning(f"Попытка {attempt}: _attach_via_input вернул False")

        log.error(f"Прикрепление НЕ удалось за {max_attempts} попыток — файл не приложен")
        return False
    finally:
        # Выключаем CDP intercept в самом конце — гарантия что между retry'ями
        # native picker не откроется случайно от чужих clikks.
        try:
            driver.execute_cdp_cmd("Page.setInterceptFileChooserDialog",
                                    {"enabled": False})
        except Exception:
            pass
