"""
shared/feedback_parser.py — Парсер писем «Новый вопрос с сайта Омск РТС»
(обратная связь с сайта customer.omskrts.ru → пересылка через feedback@).

Эти письма имеют структурный body с прямыми лейблами:
  Фамилия, имя, отчество: <ФИО>  (в новых письмах может отсутствовать)
  Адрес электронной почты: <email>
  Адрес: <адрес>
  Лицевой счет: <ЛС>
  Телефон: <телефон>
  Тема вопроса: <тема>
  -------------------------------------------------------
  Вопрос:
  <текст вопроса>
  -------------------------------------------------------

В отличие от ZHKH (табличный, с табами как разделителями) — тут плоский
текст с «label: value» парами. Намного проще распарсить и точнее чем
extract_fio_from_text по всему телу.

Если body не совпадает с шаблоном — возвращается None, и вызывающий код
использует общую логику.
"""

import re
import html
import logging
import unicodedata
from collections.abc import Mapping

log = logging.getLogger("asud.feedback")


DEFAULT_FEEDBACK_ADDRESSEE_ROUTES = {
    "Качество коммунальных услуг": ["Жуков Иван Сергеевич"],
}

QUALITY_FEEDBACK_TOPIC = "Качество коммунальных услуг"

# Письма ОЭК не имеют поля «Тема вопроса», как feedback. Поэтому для них
# используем только структурный блок «Суть обращения» и небольшой набор
# сильных признаков качества ГВС. Обычное слово «перерасчёт» намеренно не
# является признаком: оно встречается и в расчётных вопросах.
_OEK_SOURCE_RE = re.compile(
    r"(?:[\w.+-]+@(?:[\w-]+\.)*omesc\.ru\b|"
    r"омск\w*\s+энергосбыт\w*\s+компан\w*)",
    re.IGNORECASE,
)
_OEK_FOLDER_HINT_RE = re.compile(
    r"(?:^|[\\/\s_-])оэк(?:$|[\\/\s_.-])",
    re.IGNORECASE,
)
_OEK_REQUEST_RE = re.compile(r"\bсуть\s+обращения\s*:\s*", re.IGNORECASE)
_MAIL_BOUNDARY_RE = re.compile(
    r"(?im)^\s*(?:-{3,}\s*(?:original message|исходное сообщение).*|"
    r"от\s*:|from\s*:|sent\s*:|кому\s*:|to\s*:|subject\s*:|тема\s*:)"
)

_OEK_QUALITY_RULES = (
    (
        "отсутствует циркуляция ГВС",
        re.compile(
            r"(?:\b(?:отсутств\w*|нет)\b.{0,35}\bциркуляц\w*\b|"
            r"\bциркуляц\w*\b.{0,35}\b(?:отсутств\w*|нет)\b)",
            re.IGNORECASE | re.DOTALL,
        ),
    ),
    (
        "отсутствует подача ГВС",
        re.compile(
            r"(?:\bподач\w*\b.{0,35}\bгвс\b.{0,35}"
            r"\b(?:отсутств\w*|нет|не\s+(?:появ\w*|поступ\w*|"
            r"возобнов\w*|восстанов\w*))\b|"
            r"\bгвс\b.{0,35}\bподач\w*\b.{0,35}"
            r"\b(?:отсутств\w*|нет|не\s+(?:появ\w*|поступ\w*|"
            r"возобнов\w*|восстанов\w*))\b)",
            re.IGNORECASE | re.DOTALL,
        ),
    ),
    (
        "нет горячей воды",
        re.compile(
            # Отрицание должно относиться непосредственно к услуге. Нельзя
            # использовать произвольное расстояние между словами: иначе
            # «нет строки с показаниями ГВС» и «нет перерасчёта по горячей
            # воде» ошибочно превращаются в отсутствие горячей воды.
            r"(?:"
            r"\b(?:нет|отсутств\w*)\s+"
            r"(?:(?:подач\w*|водоснабжен\w*)\s+)?"
            r"(?:гор(?:яч|ч)\w*\s+вод\w*|гвс)\b|"
            r"\b(?:гор(?:яч|ч)\w*\s+вод\w*|гвс)\b"
            r"\s*[,;:—-]?\s*"
            r"(?:(?:полностью|совсем|фактически)\s+)?"
            r"(?:нет|отсутств\w*)\b"
            r"(?!\s+(?:возможност\w*|данн\w*|показан\w*|"
            r"строк\w*|перерасч\w*))|"
            r"\b(?:гор(?:яч|ч)\w*\s+вод\w*|гвс)\b\s+"
            r"не\s+(?:появ\w*|поступ\w*|возобнов\w*|восстанов\w*)\b|"
            r"\bне\s+(?:появ\w*|поступ\w*|возобнов\w*|восстанов\w*)"
            r"\s+(?:гор(?:яч|ч)\w*\s+вод\w*|гвс)\b"
            r")",
            re.IGNORECASE,
        ),
    ),
    (
        "вода отсутствует до сих пор",
        re.compile(
            r"\bвод[аы]\s+(?:нет|не\s+появ\w*)\s+"
            r"(?:до\s+сих\s+пор|все\s+еще|по\s+настоящ\w+\s+врем\w*)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "непредоставление ГВС",
        re.compile(
            r"\bне\s*предостав\w*\b.{0,45}"
            r"\b(?:гвс|гор(?:яч|ч)\w*\s+вод\w*)\b",
            re.IGNORECASE | re.DOTALL,
        ),
    ),
)

_HOT_WATER_CONTEXT_RE = re.compile(
    r"\b(?:гвс|гор(?:яч|ч)\w*\s+вод\w*)\b", re.IGNORECASE
)
_OEK_CONTEXT_QUALITY_RULES = (
    (
        "слабый напор горячей воды",
        re.compile(
            r"\b(?:слаб\w*\s+напор\w*|низк\w*\s+давлен\w*)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "низкая температура горячей воды",
        re.compile(
            r"\b(?:(?:еле|чуть|едва)\s+тепл\w*|"
            r"низк\w*\s+температур\w*|"
            r"не\s+соответств\w*.{0,25}температур\w*)\b",
            re.IGNORECASE | re.DOTALL,
        ),
    ),
)

# Маркеры: либо subject «Новый вопрос с сайта», либо явный feedback@-from,
# либо комбинация лейблов в теле.
_BODY_MARKERS = (
    "Новый вопрос с сайта",
    "feedback@customer.omskrts.ru",
)

# Лейблы и regex для значения после двоеточия до конца строки.
_FIELD_RE_TEMPLATE = r'^\s*{label}\s*:\s*([^\r\n]+?)\s*$'


def normalize_feedback_topic(value):
    """Нормализует название категории feedback для точного сопоставления.

    Сайт обычно дописывает служебный код в конце: ``[82]``. Код не является
    частью названия категории и не должен мешать маршрутизации.
    """
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = re.sub(r"\s*\[\s*\d+\s*\]\s*$", "", text)
    text = re.sub(r"\s+", " ", text).strip().rstrip(".").strip()
    return text.casefold().replace("ё", "е")


def match_feedback_addressees(topic, routes=None):
    """Возвращает адресатов для точной категории feedback либо ``None``."""
    routes = DEFAULT_FEEDBACK_ADDRESSEE_ROUTES if routes is None else routes
    if not isinstance(routes, Mapping):
        raise ValueError("feedback_addressee_routes должен быть JSON-объектом")

    normalized_topic = normalize_feedback_topic(topic)
    if not normalized_topic:
        return None

    for configured_topic, raw_addressees in routes.items():
        if normalize_feedback_topic(configured_topic) != normalized_topic:
            continue
        if isinstance(raw_addressees, str):
            raw_addressees = [raw_addressees]
        if not isinstance(raw_addressees, (list, tuple)):
            raise ValueError(
                f"Маршрут feedback {configured_topic!r} должен содержать список адресатов"
            )
        addressees = [str(value).strip() for value in raw_addressees
                      if str(value).strip()]
        if not addressees:
            raise ValueError(
                f"Маршрут feedback {configured_topic!r} не содержит адресатов"
            )
        return addressees
    return None


def _plain_text(value):
    """Приводит plain/HTML-фрагмент письма к тексту для правил."""
    text = html.unescape(unicodedata.normalize("NFKC", str(value or "")))
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("\xa0", " ").replace("ё", "е").replace("Ё", "Е")
    return re.sub(r"[ \t]+", " ", text)


def _extract_oek_request_blocks(body):
    """Возвращает только блоки после «Суть обращения», без цепочки пересылки."""
    text = _plain_text(body)
    markers = list(_OEK_REQUEST_RE.finditer(text))
    blocks = []
    for index, marker in enumerate(markers):
        end = markers[index + 1].start() if index + 1 < len(markers) else len(text)
        block = text[marker.end():end]
        boundary = _MAIL_BOUNDARY_RE.search(block)
        if boundary:
            block = block[:boundary.start()]
        block = block.strip(" \t\r\n-–—")
        if block:
            # Защита от очень длинной истории, если в письме нет стандартного
            # разделителя пересылки. Суть обращения обычно короче 1 страницы.
            blocks.append(block[:4000])
    return blocks


def detect_oek_quality_issue(body, subject=None, source_hint=None):
    """Возвращает объяснение, если ОЭК-письмо точно относится к качеству ГВС.

    ``None`` означает, что уверенного признака нет и штатный адресат меняться
    не должен. Одного слова «перерасчёт» для срабатывания недостаточно.
    """
    source_text = _plain_text(
        "\n".join(str(value or "") for value in (subject, body))
    )
    source_is_oek = bool(_OEK_SOURCE_RE.search(source_text))
    if not source_is_oek and source_hint:
        source_is_oek = bool(_OEK_FOLDER_HINT_RE.search(str(source_hint)))
    if not source_is_oek:
        return None

    for block in _extract_oek_request_blocks(body):
        for reason, pattern in _OEK_QUALITY_RULES:
            if pattern.search(block):
                return reason

        if _HOT_WATER_CONTEXT_RE.search(block):
            for reason, pattern in _OEK_CONTEXT_QUALITY_RULES:
                if pattern.search(block):
                    return reason
    return None


def _find_field(body, label):
    """Ищет значение поля по лейблу. None если не нашли."""
    pat = _FIELD_RE_TEMPLATE.replace('{label}', re.escape(label))
    m = re.search(pat, body, re.MULTILINE | re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        # Иногда после email-адреса идёт <mailto:...> — обрезаем.
        val = re.sub(r'\s*<mailto:[^>]+>\s*$', '', val).strip()
        if val and val.lower() not in ('—', '-', 'нет', 'не указан', 'не указано'):
            return val
    return None


def _extract_question(body):
    """Достаёт текст после «Вопрос:» до следующей сплошной строки дефисов
    или до конца body. None если не нашли."""
    m = re.search(
        r'Вопрос\s*:\s*\r?\n([\s\S]+?)(?:\r?\n-{5,}|\Z)',
        body, re.IGNORECASE
    )
    if m:
        q = m.group(1).strip()
        return q if q else None
    return None


def parse_feedback_body(body, subject=None):
    """Парсит body «Новый вопрос с сайта Омск РТС». Возвращает dict или None.

    Структура dict:
      {
        'фио':              'Токарева Лариса Михайловна',  # или None
        'фамилия':          'Токарева',     # первое слово ФИО
        'имя':              'Лариса',       # второе
        'отчество':         'Михайловна',   # третье (если есть)
        'email':            'tokareva.l.1976@gmail.com',
        'адрес':            'Комкова 8/1 кв.166',
        'лицевой_счет':     '97011811660',
        'телефон':          '8 (796) 204-78-06',
        'тема_вопроса':     'Перерасчёт [81]',
        'вопрос':           'Ошибка в передачи данных...',
        'корреспондент':     'Токарева Лариса Михайловна',  # либо адрес
        'корреспондент_тип': 'person',       # либо address
      }

    None если body не похоже на feedback-формат.
    """
    if not body:
        return None
    text = body
    # Маркер либо в subject, либо в body
    has_marker = any(m in text for m in _BODY_MARKERS)
    if subject and not has_marker:
        has_marker = any(m in subject for m in _BODY_MARKERS)
    if not has_marker:
        # Дополнительный тригер — наличие лейбла «Фамилия, имя, отчество:» в теле
        if 'Фамилия, имя, отчество' not in text:
            return None

    fio = _find_field(text, 'Фамилия, имя, отчество')
    parts = fio.split() if fio else []
    surname = parts[0] if parts else None
    name = parts[1] if len(parts) > 1 else None
    patronymic = parts[2] if len(parts) > 2 else None

    # В старом шаблоне было ФИО, в новом его может не быть совсем. Неполное
    # значение не используем как корреспондента, но не отбрасываем всё письмо:
    # адрес всё равно нужен для краткого содержания и ручной обработки.
    if fio and not (surname and name):
        log.debug(f"feedback-парсер: некорректное ФИО {fio!r} (нужно ≥2 слов)")
        fio = surname = name = patronymic = None

    address = _find_field(text, 'Адрес')
    if fio:
        correspondent = fio
        correspondent_kind = 'person'
    elif address:
        # Новый шаблон feedback не передаёт ФИО. Для таких писем адрес
        # становится «Фамилией», а обязательные Имя/Отчество получают прочерки.
        correspondent = address
        correspondent_kind = 'address'
    else:
        correspondent = None
        correspondent_kind = 'person'

    return {
        'фио':           fio,
        'фамилия':       surname,
        'имя':           name,
        'отчество':      patronymic,
        'email':         _find_field(text, 'Адрес электронной почты'),
        'адрес':         address,
        'лицевой_счет':  _find_field(text, 'Лицевой счет') or _find_field(text, 'Лицевой счёт'),
        'телефон':       _find_field(text, 'Телефон'),
        'тема_вопроса':  _find_field(text, 'Тема вопроса'),
        'вопрос':        _extract_question(text),
        'корреспондент': correspondent,
        'корреспондент_тип': correspondent_kind,
    }
