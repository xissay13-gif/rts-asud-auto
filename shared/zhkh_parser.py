"""
shared/zhkh_parser.py — Парсер писем-обращений через ГИС ЖКХ (рассылка ТГК-11).

Эти письма имеют строго табличный body «Информация о заявителе / Информация об
обращении». Парсер вытаскивает структурно: ФИО заявителя, номер обращения,
дату получения. Намного точнее чем общий extract_fio_from_text.

Если body не совпадает с шаблоном ГИС ЖКХ — возвращается None, и вызывающий
код использует общую логику.
"""

import re
import logging

log = logging.getLogger("asud.zhkh")

# Маркер что это вообще ГИС ЖКХ письмо: заголовок «Информация о заявителе»
_ZHKH_MARKER = "Информация о заявителе"

# Поля формата «Название_поля\t значение\t». Label может начинаться с начала
# строки ИЛИ после таба (на одной строке бывает несколько полей подряд:
# «Номер обращения\t XXX\t Дата получения обращения\t YYY\t»).
_FIELD_RE_TEMPLATE = r'(?:(?<=\t)|^)\s*{label}\s*\t\s*([^\t\r\n]+?)\s*(?=\t|\r|\n|$)'


def _find_field(body, label):
    """Ищет значение поля по названию в табличном body. None если не нашли."""
    pat = _FIELD_RE_TEMPLATE.replace('{label}', re.escape(label))
    m = re.search(pat, body, re.MULTILINE)
    if m:
        val = m.group(1).strip()
        if val and val.lower() not in ('информация не указана', '-', '—'):
            return val
    return None


def parse_zhkh_body(body):
    """Парсит body-обращение ГИС ЖКХ. Возвращает dict или None.

    Структура dict:
      {
        'фамилия':     'Айданова',
        'имя':         'Сатича',
        'отчество':    'Айсатулловна',
        'фио':         'Айданова Сатича Айсатулловна',
        'номер_обращения': '55-2026-31734',
        'дата_обращения':  '19.05.2026',  # как строка, формат DD.MM.YYYY
        'тема_обращения':  '...',           # если есть, опционально
        'адрес':       '...',               # если есть, опционально
        'телефон':     '+7 (913) ...',     # опционально
        'email':       '...',               # опционально
      }

    None если body не похоже на ГИС ЖКХ-формат.
    """
    if not body or _ZHKH_MARKER not in body:
        return None

    surname = _find_field(body, 'Фамилия')
    name = _find_field(body, 'Имя')
    patronymic = _find_field(body, 'Отчество')

    # ФИО — минимум 2 части должны быть. Иначе парсер промахнулся.
    parts = [p for p in (surname, name, patronymic) if p]
    if len(parts) < 2:
        log.debug(f"ZHKH-парсер: не хватает ФИО-частей ({surname=!r} {name=!r} {patronymic=!r})")
        return None

    fio = ' '.join(parts)

    return {
        'фамилия':         surname,
        'имя':             name,
        'отчество':        patronymic,
        'фио':             fio,
        'номер_обращения': _find_field(body, 'Номер обращения'),
        'дата_обращения':  _find_field(body, 'Дата получения обращения'),
        'тема_обращения':  _find_field(body, 'Тема обращения'),
        'адрес':           _find_field(body, 'Адрес дома/территории'),
        'телефон':         _find_field(body, 'Номер телефона'),
        'email':           _find_field(body, 'E-mail'),
    }
