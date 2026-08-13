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
import logging

log = logging.getLogger("asud.feedback")

# Маркеры: либо subject «Новый вопрос с сайта», либо явный feedback@-from,
# либо комбинация лейблов в теле.
_BODY_MARKERS = (
    "Новый вопрос с сайта",
    "feedback@customer.omskrts.ru",
)

# Лейблы и regex для значения после двоеточия до конца строки.
_FIELD_RE_TEMPLATE = r'^\s*{label}\s*:\s*([^\r\n]+?)\s*$'


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
        # становится значением поля «Фамилия», а Имя/Отчество не заполняются.
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
