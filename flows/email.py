"""
flows/email.py — Email-direct: создание Входящих документов прямо из .msg файлов
в указанной папке (без xlsx-реестра).

Два режима:
  • main()        — однопроходный: пробежать все .msg в корне папки и выйти
  • daemon_main() — непрерывный мониторинг: опрос папки раз в N сек,
                     обработка новых .msg, автоперемещение по результату
                     (Завершено / Ошибки / Черновики), Ctrl+C для остановки

Запуск через app.py с --mode=email (одноразовый) или --mode=email --watch (daemon).
"""

import os
import re
import signal
import sys
import time
import logging
from datetime import date, datetime, timedelta

import openpyxl
import extract_msg
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService

from shared import config as cfg
from shared.ui import wait_asud_loaded, set_driver_timeout
from shared.correspondent import extract_fio_from_text
from shared.okrug_parser import okrug_from_textbody
from shared.deadline import (
    parse_ddmmyyyy as _parse_ddmmyyyy,
    add_working_days as _add_working_days,
    compute_deadline as _compute_zhkh_deadline,
)
from shared.attachments import move_to_done, move_to_errors, move_to_drafts
from shared.colors import green, yellow, red, status_colored
from shared.classifier import classify_doc_type
from shared.zhkh_parser import parse_zhkh_body
from shared.feedback_parser import parse_feedback_body
from shared.xlsx_lock import xlsx_lock

# Переиспользуем создание/регистрацию/output из mix
from flows import mix as mix_flow

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger("asud")
start_time = time.monotonic()

settings = {}


# ================= ENCODING-FIX =================

def _fix_mojibake(s):
    """Эвристика: если строка похожа на CP1251 байты, прочитанные как Latin-1
    (типичная ошибка extract_msg на некоторых .msg) — реkodируем обратно.

    Triggers (все вместе):
      - 3+ символов из диапазона U+00C0..U+00FF (Latin-1 «cyrillic-looking»)
      - реальной кириллицы (U+0400..U+04FF) меньше чем «бракованных»
      - после re-encode latin-1 → cp1251 ≥80% бракованных символов
        стали кириллицей (раньше сравнивали с len(candidate)//3 — ломалось
        на строках с большим количеством ASCII: «55-2026-33646 24-я Северная,
        д. 168, к. 1 до 14.06» — 13 кириллических из 50 символов, порог 16,
        не проходило, кракозябры оставались)
    Иначе возвращаем строку как есть.
    """
    if not s or not isinstance(s, str):
        return s
    bad = sum(1 for c in s if 0x00C0 <= ord(c) <= 0x00FF)
    if bad < 3:
        return s
    cyrillic = sum(1 for c in s if 0x0400 <= ord(c) <= 0x04FF)
    if cyrillic > bad:
        return s
    try:
        candidate = s.encode('latin-1').decode('cp1251')
        new_cyr = sum(1 for c in candidate if 0x0400 <= ord(c) <= 0x04FF)
        # 0xC0-0xFF в CP1251 — вся кириллица. True mojibake → каждый bad
        # становится кириллической буквой. Допуск 80% — на случай если
        # пара байт оказалась знаком/латиницей.
        if new_cyr >= bad * 0.8:
            return candidate
    except Exception:
        pass
    return s


# ================= FOLDER CHECK WITH RETRY =================

def _normalize_unc(folder):
    """Чинит частый косяк в JSON-конфиге: путь к UNC-шаре указан с одним
    бэкслешем (\\\\server в JSON → \\server в строке) вместо двух
    (\\\\\\\\server в JSON → \\\\server в строке).

    Симптом: log.error выводит repr типа '\\\\interrao.ru\\\\oms10\\\\...' —
    repr показывает каждый реальный \\ как \\\\, и видно что в начале только
    один настоящий \\. На os.path.isdir такая «полу-UNC» отдаёт False.

    Эвристика: путь начинается с \\ + хост.домен.* (т.е. с точкой в первой
    после \\ части) и os.path.isdir(folder) == False — добавляем ведущий \\.
    """
    if not folder or os.path.isdir(folder):
        return folder
    if folder.startswith('\\') and not folder.startswith('\\\\'):
        # Берём первую часть после ведущего \\ — если в ней есть точка,
        # похоже на FQDN (interrao.ru, server.local, etc.)
        head = folder[1:].split('\\', 1)[0]
        if '.' in head:
            fixed = '\\' + folder  # дописываем второй ведущий \\
            if os.path.isdir(fixed):
                log.warning(f"Путь {folder!r} был не-UNC (один \\ в начале) — "
                            f"автоисправлено на {fixed!r}. ИСПРАВЬ settings.json: "
                            f"в начале UNC-пути должно быть \\\\\\\\ (четыре \\\\), "
                            f"они декодируются JSON'ом в два.")
                return fixed
    return folder


def _wait_for_folder(folder, retries=4, delays=(0, 2, 5, 10)):
    """os.path.isdir с retry-backoff. Нужно для UNC-путей (\\\\server\\share)
    у которых первый доступ может уйти в timeout (DNS resolve + Kerberos
    handshake + проверка прав занимают несколько секунд после холодного
    старта процесса).

    Дополнительно: если путь похож на UNC но с одним \\ в начале вместо
    двух (распространённый косяк в settings.json) — авто-исправляет.

    Возвращает (True, normalized_path) если папка доступна, (False, folder)
    после всех попыток.
    """
    import time as _time
    # Сразу пробуем нормализовать (быстрая проверка без retry)
    folder = _normalize_unc(folder)
    for attempt, delay in enumerate(delays[:retries]):
        if delay:
            log.info(f"Папка {folder!r}: попытка {attempt+1}/{retries} "
                     f"через {delay}с (UNC холодный старт?)")
            _time.sleep(delay)
        if os.path.isdir(folder):
            if attempt > 0:
                log.info(f"Папка появилась после {attempt} ретраев")
            return True, folder
    return False, folder


# ================= EMAIL → DOC_DATA =================

# Имя .msg-файла начинается с даты-времени: '2026-05-06 10-58-43.msg'
_FILENAME_DATE_RE = re.compile(r'^(\d{4}-\d{2}-\d{2})')


def _msg_date_prefix(file_path):
    """Извлекает 'YYYY-MM-DD' из имени .msg. None если не распарсилось."""
    m = _FILENAME_DATE_RE.match(os.path.basename(file_path))
    return m.group(1) if m else None


def _msg_link(msg, file_path):
    """Генерирует Link для doc_data.

    Приоритет:
      1) Дата письма из .msg (msg.date)
      2) Имя файла без расширения
    Формат — как в mix-flow реестре: 'DD.MM.YYYY HH-MM-SS'.
    """
    try:
        if msg.date:
            return msg.date.strftime("%d.%m.%Y %H-%M-%S")
    except Exception:
        pass
    return os.path.splitext(os.path.basename(file_path))[0]


# ================= PER-DATE XLSX REGISTRY =================

# Колонки реестра (под per-date).
# Колонки 1-5 — данные документа, 6 — статус/дата обработки, 7-8 —
# заполняются позже («Отписано Халецкой» — после второго прохода ГИСЖКХ,
# «Отписано в округ» — после прогона clean-resolutions).
_REGISTRY_HEADERS = ["Номер", "Link", "Округ", "Subject", "Body",
                     "Дата получения", "Планируемая дата", "Статус",
                     "Отписано Халецкой", "Отписано в округ"]
_REGISTRY_WIDTHS = {1: 18, 2: 22, 3: 8, 4: 50, 5: 80, 6: 16, 7: 16, 8: 28,
                    9: 18, 10: 18}


def _xlsx_path(base_dir, suffix=None, target_folder=None):
    """Путь к накопительному реестру (один файл, все записи append'ятся).

    Раньше был per-date (<YYYY-MM-DD>_<suffix>_резолюции.xlsx) — на каждый день
    свой файл. Теперь один файл-накопитель: <suffix>_резолюции.xlsx. Так удобнее:
      • daemon мониторит один файл, не glob
      • вся история в одном месте, фильтр в Excel по дате через колонку «Статус»
      • меньше мусора в папке

    Куда кладём:
      • Если target_folder задан И существует — прямо туда (рядом с .msg).
      • Иначе — fallback на <base_dir>/Registered/ (легаси).

    Без suffix: <where>/резолюции.xlsx
    С suffix:   <where>/<suffix>_резолюции.xlsx
    """
    name = f"{suffix}_резолюции.xlsx" if suffix else "резолюции.xlsx"

    if target_folder and os.path.isdir(target_folder):
        return os.path.join(target_folder, name)

    registered_dir = os.path.join(base_dir, "Registered")
    os.makedirs(registered_dir, exist_ok=True)
    return os.path.join(registered_dir, name)


# Backward-compat alias — старое имя ещё может встречаться в каком-то коде.
def _dated_xlsx_path(base_dir, date_prefix, suffix=None, target_folder=None):
    """DEPRECATED: используй _xlsx_path. date_prefix теперь игнорируется."""
    return _xlsx_path(base_dir, suffix=suffix, target_folder=target_folder)


def _ensure_dated_xlsx(path):
    """Создаёт per-date xlsx с шапкой если не существует."""
    if os.path.isfile(path):
        return
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Резолюции"
        ws.append(_REGISTRY_HEADERS)
        for c in range(1, len(_REGISTRY_HEADERS) + 1):
            ws.cell(row=1, column=c).font = openpyxl.styles.Font(bold=True)
        for col, w in _REGISTRY_WIDTHS.items():
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(col)].width = w
        ws.freeze_panes = "A2"
        wb.save(path)
    except Exception as e:
        log.warning(f"Не удалось создать {path}: {e}")


def _append_dated_row(path, doc, asud_id, status="OK"):
    """Дописывает строку в накопительный xlsx.

    Адаптивен к схеме файла: читает шапку, выбирает значения по именам
    колонок. Если в xlsx нет какой-то колонки (например, старый файл
    без «Планируемая дата») — её просто пропускаем. Если в xlsx новая
    колонка которой нет в этой функции — пишем пусто. Обратная
    совместимость гарантирована.

    status — internal код результата обработки. В колонку «Статус» пишем:
      OK        → «Зарегистрирован DD.MM.YYYY HH:MM»
      DRAFT     → «Черновик DD.MM.YYYY HH:MM»
      DUPLICATE → «Дубликат DD.MM.YYYY HH:MM»
    """
    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    status_text = {
        "OK":        f"Зарегистрирован {ts}",
        "DRAFT":     f"Черновик {ts}",
        "DUPLICATE": f"Дубликат {ts}",
    }.get(status, f"{status} {ts}")

    # Знакомые колонки → значения. ЛЮБОЕ имя колонки в шапке xlsx, которое
    # есть в этом словаре, будет заполнено. Остальные останутся пустыми.
    # «Дата получения» — чистая DD.MM.YYYY из тела ГИС ЖКХ (без времени).
    # Нужна P3 (clean-resolutions) чтобы пересчитать срок если «Планируемая
    # дата» в реестре пустая. Для не-ГИС писем — пусто.
    _recv = _parse_ddmmyyyy(doc.get("дата_обращения"))
    recv_str = _recv.strftime("%d.%m.%Y") if _recv else ""

    values = {
        "Номер":             asud_id or "",
        "Link":              doc.get("link") or "",
        "Округ":             doc.get("округ_прогноз") or "",
        "Subject":           doc.get("тема") or "",
        "Body":              doc.get("содержание") or "",  # уже _clean_body
        "Дата получения":    recv_str,
        "Планируемая дата":  doc.get("планируемая_дата") or "",
        "Статус":            status_text,
    }
    try:
        with xlsx_lock(path, timeout=60):
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            # Шапка существующего файла
            headers = [str(c.value or '').strip() for c in next(ws.iter_rows(max_row=1))]
            row = [values.get(h, "") for h in headers]
            ws.append(row)
            wb.save(path)
    except Exception as e:
        log.warning(f"Не удалось записать строку в {path}: {e}")


def _list_root_msgs(folder_path):
    """Возвращает sorted list абсолютных путей к .msg в корне folder_path
    (без рекурсии). Пустой list если папки нет / I/O ошибка."""
    try:
        out = []
        for f in os.listdir(folder_path):
            full = os.path.join(folder_path, f)
            if os.path.isfile(full) and f.lower().endswith('.msg'):
                out.append(full)
        return sorted(out)
    except OSError as e:
        log.error(f"Не могу прочитать папку {folder_path}: {e}")
        return []


def _parse_one_msg(msg_path, process_mode="mix"):
    """Парсит один .msg в doc_data dict. None если не получилось/пустое/брак.

    process_mode:
      'mix'   — классификатор НЕ вызывается, тип всегда из settings
                ('default_type_idx', по умолчанию 8). Корреспондент из ФИО в теле.
      'smart' — вызывается classify_doc_type (для определения тип_индекс).
                cat -1 → пропуск. cat 0 → тип 8 + force_draft.
                Корреспондент всё равно будет затёрт в _process_doc на «Неизвестный...».

    Использует module-global settings.
    """
    unknown = settings.get("unknown_correspondent",
                            cfg.DEFAULTS["unknown_correspondent"])

    try:
        msg = extract_msg.openMsg(msg_path)
        subject = _fix_mojibake(msg.subject or "")
        body = _fix_mojibake(msg.body or "")
        link = _msg_link(msg, msg_path)
        try:
            msg.close()
        except Exception:
            pass
    except Exception as e:
        log.warning(f"Не удалось прочитать {os.path.basename(msg_path)}: {e}")
        return None

    if not body and not subject:
        log.warning(f"Пустое письмо {os.path.basename(msg_path)} — пропускаю")
        return None

    # ГИС ЖКХ — спец-парсер табличного body. Запускаем ДО классификатора —
    # для ZHKH-писем тип всегда 8 (письма граждан), классификатор не трогаем.
    zhkh = parse_zhkh_body(body)
    # Feedback с сайта Омск РТС («Новый вопрос с сайта Омск РТС») — структурный
    # парсер. ФИО заявителя указано явно в теле → авто-регистрация, не черновик.
    # Парсим ТОЛЬКО если не ZHKH (zhkh > feedback по приоритету).
    feedback = None if zhkh else parse_feedback_body(body, subject=subject)

    force_draft = False
    if zhkh:
        # ZHKH-письмо: тип 8 железно, классификатор скипаем
        type_idx = 8
        log.info(f"{os.path.basename(msg_path)}: ГИС ЖКХ → {zhkh.get('фио') or '(аноним)'}, "
                 f"№{zhkh.get('номер_обращения')} от {zhkh.get('дата_обращения')}")
    elif feedback:
        # Feedback-письмо: тип 8 (письма граждан), регистрация а не черновик
        type_idx = 8
        log.info(f"{os.path.basename(msg_path)}: feedback Омск РТС → "
                 f"{feedback['фио']}, ЛС {feedback.get('лицевой_счет') or '—'}")
    elif process_mode == "smart":
        # Не ZHKH: для smart-режима пускаем общий классификатор
        type_idx = classify_doc_type(body)
        if type_idx == -1:
            log.warning(f"{os.path.basename(msg_path)}: помечено классификатором "
                        f"как 'случайно отправил' — пропускаю")
            return None
        if type_idx == 0:
            type_idx = 8
            force_draft = True
            log.info(f"{os.path.basename(msg_path)}: тип не определился → 8, "
                     f"оставляю в черновике для ручной проверки")
    else:
        # mix-режим без ZHKH: дефолтный тип из settings
        type_idx = settings.get("default_type_idx", 8)
    type_name = cfg.DOC_TYPE_MAP.get(
        type_idx, "Письма, заявления и жалобы граждан, акционеров")

    clean_subject = re.sub(r'^(FW:|RE:|Fwd:)\s*', '',
                            str(subject).strip(), flags=re.IGNORECASE)
    # ГИС ЖКХ: если тема приехала без префикса (начинается сразу с номера
    # обращения вроде «55-2026-36415 ул. ...») — допишем «ГИС ЖКХ» вперёд
    # для единообразия в реестре и в карточке АСУД.
    if zhkh and not re.match(r'^\s*ГИС\s*ЖКХ', clean_subject, re.IGNORECASE):
        clean_subject = f"ГИС ЖКХ {clean_subject}"
    body_clean = mix_flow._clean_body(body) if body else clean_subject

    if zhkh:
        # ФИО может быть None для анонимных ГИС ЖКХ-писем (нет поля «Заявитель»)
        # — тогда корреспондент = «Неизвестный», но номер/тема/адрес из zhkh
        # всё равно используются.
        correspondent = zhkh.get('фио') or unknown
        corr_found = bool(zhkh.get('фио'))
        fio_src = "zhkh"
    elif feedback:
        correspondent = feedback['фио']
        corr_found = True
        fio_src = "feedback"
    elif force_draft:
        # Smart + cat-0: принудительно черновик с фикс. корреспондентом
        # (чтобы письмо точно ушло в DRAFT-ветку mix.create_one_document).
        corr_found = False
        correspondent = unknown
        fio_src = ""
    else:
        fio, fio_src = extract_fio_from_text(body_clean)
        correspondent = fio if fio else unknown
        corr_found = bool(fio)

    try:
        okrug = okrug_from_textbody(body_clean, base_dir_fn=cfg.get_base_dir)
    except Exception as e:
        log.warning(f"okrug_parser упал для {os.path.basename(msg_path)}: {e}")
        okrug = None

    # Для ZHKH и feedback-писем краткое содержание в АСУД = тема (subject),
    # а не полный body (там длинный шаблонный текст с подписью).
    # Для feedback дополнительно приклеиваем адрес из тела: «<subject> — <адрес>»
    # (subject обычно «Новый вопрос с сайта Омск РТС» — оставляем как есть,
    # юзеру удобнее видеть и его, и адрес).
    # Срок отработки ГИС ЖКХ: РТС даёт 17 РАБОЧИХ дней с даты получения,
    # но если срок по ГИС ЖКХ раньше — берём его (min из двух).
    zhkh_deadline = None
    if zhkh:
        zhkh_deadline = _compute_zhkh_deadline(
            zhkh.get('дата_обращения'), zhkh.get('планируемая_дата'))

    if feedback:
        addr = (feedback.get('адрес') or '').strip()
        soderzhanie = f"{clean_subject} — {addr}" if addr else clean_subject
    elif zhkh:
        # В краткое содержание дописываем срок исполнения (если посчитался)
        soderzhanie = clean_subject
        if zhkh_deadline:
            soderzhanie = f"{clean_subject} (срок до {zhkh_deadline})"
    else:
        soderzhanie = body_clean

    return {
        "row_idx": 1,
        "содержание": soderzhanie,
        "корреспондент": correspondent,
        "корр_найден": corr_found,
        "корр_источник": fio_src,
        "тема": clean_subject,
        "тип_индекс": type_idx,
        "тип_название": type_name,
        "link": link,
        "файл": msg_path,
        "округ_прогноз": okrug,
        "msg_date_prefix": _msg_date_prefix(msg_path),
        "force_draft": force_draft,
        # ZHKH-специфичные поля (None если письмо не из ГИС ЖКХ).
        # Дальше передадутся в mix.create_one_document для заполнения
        # доп. полей АСУД-карточки.
        "номер_обращения":  zhkh.get('номер_обращения')  if zhkh else None,
        "дата_обращения":   zhkh.get('дата_обращения')   if zhkh else None,
        # Планируемая дата = вычисленный срок отработки (min(получ.+17 раб.дн,
        # срок ГИС)). Этот же срок уходит в контрольную дату резолюции.
        # Fallback на сырую дату ГИС если вычисление не удалось.
        "планируемая_дата": (zhkh_deadline or zhkh.get('планируемая_дата')) if zhkh else None,
    }


def load_emails(folder_path, process_mode="mix"):
    """Парсит все .msg в корне folder_path. Возвращает list of doc dicts.
    process_mode пробрасывается в _parse_one_msg (см. его docstring)."""
    msg_files = _list_root_msgs(folder_path)
    log.info(f"Найдено .msg файлов: {len(msg_files)}")

    rows, skipped = [], 0
    for idx, msg_path in enumerate(msg_files, 1):
        doc = _parse_one_msg(msg_path, process_mode)
        if doc is None:
            skipped += 1
            continue
        doc["row_idx"] = idx
        rows.append(doc)

    log.info(f"Загружено: {len(rows)} писем, пропущено: {skipped}")
    return rows


# ================= PROCESSING ONE DOC =================

def _print_doc_line(index, total, status, info=""):
    """Одна строка результата по документу — с цветным статусом в консоль.
    OK — зелёный, DRAFT — жёлтый, FAILED — красный, DUPLICATE — без цвета."""
    label = status_colored(status)
    suffix = f" — {info}" if info else ""
    print(f"  [{index}/{total}] {label}{suffix}")


def _process_doc(driver, doc, base_dir, folder, index, total, in_daemon,
                 process_mode="mix", output_suffix=None):
    """Обрабатывает один doc: create_one_document → ветвление по статусу
    → запись в xlsx + перенос .msg.

    process_mode:
      'mix'   — текущая логика: ФИО найдено → register, нет → DRAFT
      'smart' — всегда черновик: forсируем корр_найден=False + фикс. корреспондент;
                 DRAFT считается УСПЕХОМ → пишем в реестр и переносим в Завершено/

    output_suffix — суффикс в имени per-date xlsx (для разделения реестров
    при параллельных запусках двух пресетов).

    Возвращает финальный статус: 'OK' | 'DUPLICATE' | 'DRAFT' | 'FAILED'.
    in_daemon=True (mix-режим): DRAFT → перенос в Черновики/.
    """
    msg_path = doc.get("файл")
    written_xlsx = None  # путь к xlsx куда записали (для второго прохода)

    if process_mode == "smart" and doc.get("корр_источник") not in ("zhkh", "feedback"):
        # Smart-пресет: каждый .msg создаётся как черновик с фикс. корреспондентом.
        # Исключение — структурно распознанные письма (ZHKH / feedback): у них
        # ФИО заявителя в теле явно, оно надёжное → пускаем в регистрацию.
        doc["корр_найден"] = False
        doc["корреспондент"] = settings.get("unknown_correspondent",
                                             cfg.DEFAULTS["unknown_correspondent"])

    asud_id = mix_flow.create_one_document(driver, doc, index, total)
    status = mix_flow._last_result.get("status", "FAILED")

    if status == "OK":
        written_xlsx = _xlsx_path(base_dir, output_suffix, target_folder=folder)
        _ensure_dated_xlsx(written_xlsx)
        _append_dated_row(written_xlsx, doc, asud_id, status="OK")
        move_to_done(msg_path, folder)
    elif status == "DUPLICATE":
        log.info(f"Документ {index}: уже зарегистрирован — .msg в Завершено/")
        move_to_done(msg_path, folder)
    elif status == "DRAFT":
        if process_mode == "smart":
            # Smart: черновик — это нормальный исход. Пишем в реестр (без АСУД-ID)
            # и переносим .msg в Завершено/.
            written_xlsx = _xlsx_path(base_dir, output_suffix, target_folder=folder)
            _ensure_dated_xlsx(written_xlsx)
            _append_dated_row(written_xlsx, doc, asud_id or "", status="DRAFT")
            move_to_done(msg_path, folder)
            log.info(f"Документ {index}: создан как черновик (smart) — .msg в Завершено/")
        elif in_daemon:
            log.info(f"Документ {index}: ФИО не найдено — .msg в Черновики/")
            move_to_drafts(msg_path, folder)
        # one-shot mix: оставляем в корне как и было
    else:  # FAILED — caller сам решает что делать (retry / move-to-errors)
        pass

    return status, asud_id, written_xlsx


# ================= MAIN =================

def main():
    global settings
    settings = cfg.load()
    cfg.setup_file_logger("email")
    cfg.keep_system_awake(True)

    # Читаем логику обработки сразу — нужно для корректного превью
    process_mode = os.environ.get('ASUD_EMAIL_PROCESS_MODE', 'mix')

    log.info("=" * 50)
    log.info("АСУД ИК — Email-direct (создание из .msg-писем)")
    log.info("=" * 50)

    base_dir = cfg.get_base_dir()

    # Запрос пути к папке с .msg — спрашиваем всегда, даже при пресете.
    # Дефолтом подставляется ASUD_EMAIL_FOLDER (от пресета) или email_folder из конфига.
    default = os.environ.get('ASUD_EMAIL_FOLDER') \
              or settings.get("email_folder", "")
    print(f"\nПапка с .msg-письмами (только из корня папки, подпапки игнорируются).")
    if default:
        print(f"Enter — использовать: {default}")
    user_dir = input("Путь: ").strip().strip('"').strip("'")
    folder = user_dir or default

    if folder:
        ok, folder = _wait_for_folder(folder)
    else:
        ok = False
    if not ok:
        log.error(f"Папка не найдена (после ретраев): {folder!r}")
        input("Enter...")
        sys.exit(1)

    log.info(f"Папка писем: {folder}")

    # Парсим письма (process_mode определяет, использовать ли классификатор)
    docs = load_emails(folder, process_mode)
    if not docs:
        log.error("Нет .msg файлов или все пропущены")
        input("Enter...")
        sys.exit(1)

    # Превью — зависит от process_mode
    if process_mode == "smart":
        print(f"\nПервые 5 (все будут созданы как ЧЕРНОВИКИ):")
        for i, d in enumerate(docs[:5], 1):
            print(f"  {i}. [тип {d['тип_индекс']}] {d['тема'][:60]}")
        print(f"\nВсего: {len(docs)}  (тип определён классификатором, "
              f"корреспондент будет «Неизвестный...»)")
        print("режим: EMAIL/SMART  —  создание ТОЛЬКО как черновик + .msg "
              "(без регистрации)")
    else:  # mix
        known = sum(1 for d in docs if d["корр_найден"])
        unknown_n = len(docs) - known
        print(f"\nПервые 5:")
        for i, d in enumerate(docs[:5], 1):
            flag = 'OK' if d["корр_найден"] else '!!'
            print(f"  {i}. [тип {d['тип_индекс']}] {flag} "
                  f"{d['корреспондент'][:30]} | {d['тема'][:50]}")
        print(f"\nВсего: {len(docs)}  (ФИО найдено: {known}, без ФИО: {unknown_n})")
        print("режим: EMAIL/MIX  —  создание + регистрация + На резолюцию + .msg")

    if input("Начать? (да/нет): ").strip().lower() not in ("да", "д", "y", "yes", ""):
        sys.exit(0)

    # === Запуск браузера и обработки (повторяем mix-loop, но с нашими docs)
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

    # Настраиваем mix_flow.settings — он использует module-level global
    mix_flow.settings = settings

    output_suffix = os.environ.get('ASUD_OUTPUT_SUFFIX') or None
    log.info(f"Логика обработки: {process_mode}"
             + (f", суффикс реестра: {output_suffix}" if output_suffix else ""))

    try:
        url = settings.get("asud_url", cfg.DEFAULTS["asud_url"])
        log.info(f"Открываю {url}")
        driver.get(url)
        wait_asud_loaded(driver)

        # Per-date реестры: Registered/YYYY-MM-DD[_<suffix>]_резолюции.xlsx.
        # Каждый doc пишется в xlsx своей даты (из имени .msg).
        log.info(f"Per-date реестры в: {os.path.join(base_dir, 'Registered')}")

        done_count, dup_count, draft_count, err_count = 0, 0, 0, 0
        # Просто счётчик зарегистрированных ГИСЖКХ — для итогового лога.
        # Сама отписка делается отдельным процессом (asud_zhkh_daemon.bat).
        registered_docs = []
        for i, doc in enumerate(docs, 1):
            msg_path = doc.get("файл")
            try:
                status, asud_id, written_xlsx = _process_doc(
                                       driver, doc, base_dir, folder,
                                       i, len(docs), in_daemon=False,
                                       process_mode=process_mode,
                                       output_suffix=output_suffix)
                if status == "OK":
                    done_count += 1
                    if asud_id:
                        registered_docs.append(asud_id)
                elif status == "DUPLICATE":
                    dup_count += 1
                elif status == "DRAFT":
                    draft_count += 1
                else:  # FAILED
                    move_to_errors(msg_path, folder,
                                   f"Регистрация не удалась (status={status})")
                    err_count += 1
                _print_doc_line(i, len(docs), status,
                                 doc.get("тема", "")[:60])
            except Exception as e:
                log.error(f"ОШИБКА документ {i}: {e}")
                move_to_errors(msg_path, folder, f"Exception: {e}")
                err_count += 1
                _print_doc_line(i, len(docs), "FAILED", str(e)[:80])
                try:
                    driver.get(url)
                    wait_asud_loaded(driver)
                except Exception:
                    pass
                continue

        elapsed_seconds = time.monotonic() - start_time
        elapsed = timedelta(seconds=int(elapsed_seconds))
        avg = (timedelta(seconds=int(elapsed_seconds / done_count))
               if done_count else None)
        # Лог-файл — плоский (без ANSI), консоль — цветной
        plain = [
            "",
            "=" * 60,
            "ГОТОВО!",
            f"  Обработано:   {done_count} / {len(docs)}  (→ Завершено/)",
            f"  Дубликаты:    {dup_count}  (уже были в АСУД, → Завершено/)",
            f"  В черновиках: {draft_count}  (ФИО не найдено, .msg остался в корне)",
            f"  Ошибок:       {err_count}  (→ Ошибки/)",
            f"  Затрачено:    {elapsed}" + (f"  (в среднем {avg}/док)" if avg else ""),
            "=" * 60,
        ]
        for line in plain:
            log.info(line)
        print("")
        print("=" * 60)
        print("ГОТОВО!")
        print(f"  Обработано:   {green(str(done_count))} / {len(docs)}  (→ Завершено/)")
        print(f"  Дубликаты:    {dup_count}  (уже были в АСУД, → Завершено/)")
        print(f"  В черновиках: {yellow(str(draft_count))}  (ФИО не найдено, .msg в корне)")
        print(f"  Ошибок:       {red(str(err_count))}  (→ Ошибки/)")
        print(f"  Затрачено:    {elapsed}" + (f"  (в среднем {avg}/док)" if avg else ""))
        print("=" * 60)

        # Раньше тут запускался ZHKH-второй проход (Басманов → резолюции
        # Халецкой). Теперь это отдельный процесс — zhkh_daemon, который
        # читает xlsx-реестр и отписывает в фоне.
        # Запускается отдельно: asud_zhkh_daemon.bat
        if output_suffix == "ГИСЖКХ" and registered_docs:
            log.info(f"Зарегистрировано {len(registered_docs)} ГИСЖКХ-документов. "
                     f"Для отписки запусти asud_zhkh_daemon.bat (или он уже работает фоном).")

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


# ================= DAEMON MODE =================

# Sigint handler — устанавливает флаг, текущий документ доделывается, потом выход.
_stop_flag = False


def _on_sigint(signum, frame):
    global _stop_flag
    if _stop_flag:
        log.warning("Повторный Ctrl+C — выход немедленно")
        sys.exit(130)
    _stop_flag = True
    log.info("Ctrl+C получен — закончу текущий документ и выйду из мониторинга")


def _interruptible_sleep(seconds):
    """Sleep с пробуждением по _stop_flag (проверка раз в секунду)."""
    for _ in range(int(seconds)):
        if _stop_flag:
            return
        time.sleep(1)


def daemon_main():
    """Непрерывный мониторинг папки: опрос раз в N сек, обработка новых .msg.
    Ctrl+C для остановки (graceful — после текущего документа)."""
    global settings
    settings = cfg.load()
    cfg.setup_file_logger("email_daemon")
    cfg.keep_system_awake(True)

    log.info("=" * 50)
    log.info("АСУД ИК — Email-DAEMON (непрерывный мониторинг)")
    log.info("=" * 50)

    base_dir = cfg.get_base_dir()
    interval = int(settings.get("email_watch_interval_sec",
                                cfg.DEFAULTS["email_watch_interval_sec"]))
    max_retries = int(settings.get("email_max_retries",
                                    cfg.DEFAULTS["email_max_retries"]))
    process_mode = os.environ.get('ASUD_EMAIL_PROCESS_MODE', 'mix')
    output_suffix = os.environ.get('ASUD_OUTPUT_SUFFIX') or None

    # Multi-folder режим — пресет с "folders" списком вместо "folder" строки.
    # Если ASUD_EMAIL_FOLDERS_JSON задан, обрабатываем все эти папки
    # (по очереди если round_robin=true, либо все за тик если false).
    import json as _json
    folders_list = []
    round_robin_mode = False
    raw_json = os.environ.get('ASUD_EMAIL_FOLDERS_JSON')
    if raw_json:
        try:
            raw_list = _json.loads(raw_json)
            # Нормализуем элементы: string → {dir, output_suffix=basename},
            # dict → как есть
            for e in raw_list:
                if isinstance(e, dict):
                    if e.get("dir"):
                        folders_list.append({
                            "dir": e["dir"],
                            "output_suffix": e.get("output_suffix") or os.path.basename(e["dir"]),
                        })
                elif isinstance(e, str):
                    folders_list.append({
                        "dir": e,
                        "output_suffix": os.path.basename(e),
                    })
            round_robin_mode = os.environ.get('ASUD_EMAIL_ROUND_ROBIN') == '1'
        except Exception as e:
            log.error(f"ASUD_EMAIL_FOLDERS_JSON не распарсился: {e}")
            folders_list = []

    log.info(f"Логика обработки: {process_mode}"
             + (f", суффикс реестра: {output_suffix}" if output_suffix and not folders_list else ""))

    if folders_list:
        # Multi-folder: показываем список, валидируем (но не падаем если что-то
        # не нашлось — daemon будет переопрашивать)
        log.info(f"Multi-folder режим: {len(folders_list)} папок"
                 + (" (round-robin)" if round_robin_mode else " (все за тик)"))
        for e in folders_list:
            ok, new_dir = _wait_for_folder(e["dir"])
            e["dir"] = new_dir  # сохраним возможный UNC-fix
            marker = "✓" if ok else "⚠"
            log.info(f"  {marker} {e['dir']} (suffix={e['output_suffix']})")
        # folder для legacy-вызовов (move_to_errors и т.п. перед стартом loop'а)
        # просто берём первую папку. Дальше в loop'е current_folder per-tick.
        folder = folders_list[0]["dir"]
    else:
        # Single-folder (legacy): спрашиваем папку у юзера
        default = os.environ.get('ASUD_EMAIL_FOLDER') \
                  or settings.get("email_folder", "")
        print(f"\nПапка с .msg-письмами для непрерывного мониторинга.")
        if default:
            print(f"Enter — использовать: {default}")
        user_dir = input("Путь: ").strip().strip('"').strip("'")
        folder = user_dir or default
        if folder:
            ok, folder = _wait_for_folder(folder)
        else:
            ok = False
        if not ok:
            log.error(f"Папка не найдена (после ретраев): {folder!r}")
            input("Enter...")
            sys.exit(1)
        log.info(f"Папка: {folder}")

    log.info(f"Опрос: каждые {interval} сек, макс retry: {max_retries}")
    print(f"\nМониторинг включён. Ctrl+C для остановки.")

    # Browser
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
    mix_flow.settings = settings

    signal.signal(signal.SIGINT, _on_sigint)

    url = settings.get("asud_url", cfg.DEFAULTS["asud_url"])
    log.info(f"Открываю {url}")
    driver.get(url)
    wait_asud_loaded(driver)

    # Счётчики и retry-state
    retry_count = {}  # basename → int (фейлов подряд)
    totals = {"OK": 0, "DUPLICATE": 0, "DRAFT": 0, "FAILED": 0, "ITER": 0}
    rr_idx = 0  # для round-robin между папками

    def _process_folder(current_folder, current_suffix):
        """Обрабатывает все .msg из current_folder. Логика как раньше,
        просто вынесена в функцию чтобы вызываться для каждой папки в
        multi-folder режиме. Использует closure: driver, base_dir и т.д."""
        queue = _list_root_msgs(current_folder)
        if not queue:
            return 0  # ничего не было
        log.info(f"[итер. {totals['ITER']}] {os.path.basename(current_folder)}: "
                 f"в очереди {len(queue)}")
        for idx, msg_path in enumerate(queue, 1):
            if _stop_flag:
                return idx
            name = os.path.basename(msg_path)

            doc = _parse_one_msg(msg_path, process_mode)
            if doc is None:
                move_to_errors(msg_path, current_folder,
                               "Не удалось распарсить или пустое")
                totals["FAILED"] += 1
                retry_count.pop(name, None)
                _print_doc_line(idx, len(queue), "FAILED",
                                 "не распарсилось / пустое")
                continue

            try:
                status, _asud_id, _xlsx = _process_doc(
                                       driver, doc, base_dir, current_folder,
                                       idx, len(queue), in_daemon=True,
                                       process_mode=process_mode,
                                       output_suffix=current_suffix)
                if status == "FAILED":
                    retry_count[name] = retry_count.get(name, 0) + 1
                    if retry_count[name] >= max_retries:
                        move_to_errors(msg_path, current_folder,
                            f"Регистрация не удалась за {max_retries} попыток")
                        retry_count.pop(name, None)
                        totals["FAILED"] += 1
                        _print_doc_line(idx, len(queue), "FAILED",
                                         f"max_retries ({max_retries}) → Ошибки/")
                    else:
                        log.warning(f"{name}: фейл {retry_count[name]}/{max_retries} "
                                    f"— оставляю в корне на следующую итерацию")
                        _print_doc_line(idx, len(queue), "FAILED",
                                         f"retry {retry_count[name]}/{max_retries}")
                    try:
                        driver.get(url)
                        wait_asud_loaded(driver)
                    except Exception:
                        pass
                else:
                    totals[status] = totals.get(status, 0) + 1
                    retry_count.pop(name, None)
                    _print_doc_line(idx, len(queue), status,
                                     doc.get("тема", "")[:60])
            except Exception as e:
                log.error(f"Exception на {name}: {e}")
                retry_count[name] = retry_count.get(name, 0) + 1
                if retry_count[name] >= max_retries:
                    move_to_errors(msg_path, current_folder, f"Exception: {e}")
                    retry_count.pop(name, None)
                    totals["FAILED"] += 1
                _print_doc_line(idx, len(queue), "FAILED", str(e)[:80])
                try:
                    driver.get(url)
                    wait_asud_loaded(driver)
                except Exception:
                    pass
        return len(queue)

    try:
        while not _stop_flag:
            totals["ITER"] += 1

            # Какие папки обрабатываем на этом тике
            if folders_list and round_robin_mode:
                entry = folders_list[rr_idx % len(folders_list)]
                rr_idx += 1
                log.info(f"[итер. {totals['ITER']}] round-robin: "
                         f"{os.path.basename(entry['dir'])}")
                tick_folders = [entry]
            elif folders_list:
                tick_folders = folders_list
            else:
                tick_folders = [{"dir": folder, "output_suffix": output_suffix}]

            processed = 0
            for entry in tick_folders:
                if _stop_flag:
                    break
                processed += _process_folder(entry["dir"], entry.get("output_suffix"))

            if processed == 0:
                log.info(f"[итер. {totals['ITER']}] очереди пусты — sleep {interval}s")

            if _stop_flag:
                break
            log.info(f"  итог итер. {totals['ITER']}: "
                     f"OK={totals['OK']} DUP={totals['DUPLICATE']} "
                     f"DRAFT={totals['DRAFT']} FAIL={totals['FAILED']}")
            print(f"  итог итер. {totals['ITER']}: "
                  f"OK={green(str(totals['OK']))} DUP={totals['DUPLICATE']} "
                  f"DRAFT={yellow(str(totals['DRAFT']))} "
                  f"FAIL={red(str(totals['FAILED']))}")
            _interruptible_sleep(interval)

        log.info("=" * 60)
        log.info("МОНИТОРИНГ ОСТАНОВЛЕН")
        log.info(f"  Итераций:   {totals['ITER']}")
        log.info(f"  Обработано: {totals['OK']}")
        log.info(f"  Дубликаты:  {totals['DUPLICATE']}")
        log.info(f"  Черновики:  {totals['DRAFT']}")
        log.info(f"  Ошибки:     {totals['FAILED']}")
        log.info("=" * 60)
        print("=" * 60)
        print("МОНИТОРИНГ ОСТАНОВЛЕН")
        print(f"  Итераций:   {totals['ITER']}")
        print(f"  Обработано: {green(str(totals['OK']))}")
        print(f"  Дубликаты:  {totals['DUPLICATE']}")
        print(f"  Черновики:  {yellow(str(totals['DRAFT']))}")
        print(f"  Ошибки:     {red(str(totals['FAILED']))}")
        print("=" * 60)

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        cfg.keep_system_awake(False)


if __name__ == "__main__":
    if os.environ.get("ASUD_WATCH") == "1":
        daemon_main()
    else:
        main()
