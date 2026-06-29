"""
shared/zhkh_parser.py — Парсер писем-обращений через ГИС ЖКХ (рассылка ТГК-11).

Поддерживает ТРИ исторических формата тела письма:
  1) Старый табличный «Информация о заявителе»: Фамилия/Имя/Отчество раздельно,
     разделитель — таб.
  2) Блочный (июнь 2026): label на своей строке, пустая строка, значение
     на следующей («Фамилия\\n\\nАрнгольд\\n\\n»).
  3) Новый колон-табличный (конец июня 2026): «Заявитель:\\t<ФИО>\\t» —
     ФИО одним полем, лейблы с двоеточием. Маркер «Информация о заявителе»
     отсутствует. Поля: Номер обращения / Дата получения / Тема / Заявитель /
     Email заявителя / Телефон / Адрес / Срок исполнения / Текст обращения.

Парсер вытаскивает структурно: ФИО заявителя, номер обращения, дату,
планируемую дату. Намного точнее чем общий extract_fio_from_text.

Если body не совпадает ни с одним шаблоном — возвращается None, и вызывающий
код использует общую логику.
"""

import re
import logging

log = logging.getLogger("asud.zhkh")

# Маркер старого/блочного формата: заголовок «Информация о заявителе»
_ZHKH_MARKER = "Информация о заявителе"

# Маркеры нового колон-табличного формата (формат 3). Любой из них +
# отсутствие старого маркера → формат 3.
_NEW_MARKERS = ("Номер обращения:", "Текст обращения:", "Заявитель:")

# Шаблоны раскладки полей:
#   BLOCK — label / пустая строка / значение
#   TAB   — «Label\t Value\t» (несколько полей на строке)
#   COLON — «Label:\t Value\t» (формат 3, лейбл с двоеточием)
_BLOCK_RE_TEMPLATE = r'(?ms)^\s*{label}\s*\r?\n\s*\r?\n\s*([^\r\n]+?)\s*\r?$'
_TAB_RE_TEMPLATE   = r'(?m)(?:(?<=\t)|^)\s*{label}\s*\t\s*([^\t\r\n]+?)\s*(?=\t|\r|\n|$)'
_COLON_RE_TEMPLATE = r'(?m)^\s*{label}\s*:\s*\t?\s*([^\t\r\n]+?)\s*(?=\t|\r|\n|$)'


def _find_field(body, label, templates=None):
    """Ищет значение поля по названию. Пробует переданные шаблоны по очереди."""
    if templates is None:
        templates = (_BLOCK_RE_TEMPLATE, _TAB_RE_TEMPLATE)
    for tmpl in templates:
        pat = tmpl.replace('{label}', re.escape(label))
        m = re.search(pat, body)
        if m:
            val = m.group(1).strip()
            if val and val.lower() not in ('информация не указана', '-', '—', 'нет'):
                return val
    return None


def _date_only(val):
    """Из «21.07.2026 23:59» → «21.07.2026». Если времени нет — как есть."""
    if not val:
        return val
    m = re.match(r'(\d{2}\.\d{2}\.\d{4})', val.strip())
    return m.group(1) if m else val.strip()


def _parse_new_format(body):
    """Формат 3: колон-табличный, ФИО одним полем «Заявитель».
    Возвращает dict или None.

    Для анонимных писем (поля «Заявитель» нет) фио=None — вызывающий код
    подставит «Неизвестный». Остальные поля (номер/тема/адрес) всё равно
    извлекаются и используются.
    """
    C = (_COLON_RE_TEMPLATE,)
    fio = _find_field(body, 'Заявитель', templates=C)
    num = _find_field(body, 'Номер обращения', templates=C)

    # Без номера обращения это не ГИС ЖКХ-письмо нового формата — отдаём None.
    if not num:
        return None

    return {
        'фамилия':         (fio.split()[0] if fio else None),
        'имя':             (fio.split()[1] if fio and len(fio.split()) > 1 else None),
        'отчество':        (fio.split()[2] if fio and len(fio.split()) > 2 else None),
        'фио':             fio,  # может быть None (анонимное) — caller подставит «Неизвестный»
        'номер_обращения': num,
        'дата_обращения':  _find_field(body, 'Дата получения', templates=C),
        'планируемая_дата': _date_only(_find_field(body, 'Срок исполнения', templates=C)),
        'тема_обращения':  _find_field(body, 'Тема', templates=C),
        'адрес':           _find_field(body, 'Адрес', templates=C),
        'телефон':         _find_field(body, 'Телефон', templates=C),
        'email':           _find_field(body, 'Email заявителя', templates=C),
    }


def parse_zhkh_body(body):
    """Парсит body-обращение ГИС ЖКХ. Возвращает dict или None.

    Структура dict:
      {
        'фамилия':     'Айданова',
        'имя':         'Сатича',
        'отчество':    'Айсатулловна',
        'фио':         'Айданова Сатича Айсатулловна',  # None если анонимное
        'номер_обращения': '55-2026-31734',
        'дата_обращения':  '19.05.2026',  # как строка, формат DD.MM.YYYY
        'планируемая_дата': '21.07.2026',  # опционально
        'тема_обращения':  '...',           # если есть, опционально
        'адрес':       '...',               # если есть, опционально
        'телефон':     '+7 (913) ...',     # опционально
        'email':       '...',               # опционально
      }

    None если body не похоже ни на один ГИС ЖКХ-формат.
    """
    if not body:
        return None

    # Формат 1/2 — есть маркер «Информация о заявителе»
    if _ZHKH_MARKER in body:
        surname = _find_field(body, 'Фамилия')
        name = _find_field(body, 'Имя')
        patronymic = _find_field(body, 'Отчество')
        parts = [p for p in (surname, name, patronymic) if p]
        if len(parts) >= 2:
            return {
                'фамилия':         surname,
                'имя':             name,
                'отчество':        patronymic,
                'фио':             ' '.join(parts),
                'номер_обращения': _find_field(body, 'Номер обращения'),
                'дата_обращения':  _find_field(body, 'Дата получения обращения'),
                'планируемая_дата': _find_field(body, 'Планируемая дата исполнения'),
                'тема_обращения':  _find_field(body, 'Тема обращения'),
                'адрес':           _find_field(body, 'Адрес дома/территории'),
                'телефон':         _find_field(body, 'Номер телефона'),
                'email':           _find_field(body, 'E-mail'),
            }
        log.debug(f"ZHKH (формат 1/2): не хватает ФИО-частей "
                  f"({surname=!r} {name=!r} {patronymic=!r})")
        return None

    # Формат 3 — колон-табличный (маркеры «Номер обращения:» / «Заявитель:»)
    if any(m in body for m in _NEW_MARKERS):
        result = _parse_new_format(body)
        if result:
            log.debug(f"ZHKH (формат 3): фио={result.get('фио')!r}, "
                      f"№{result.get('номер_обращения')}")
            return result

    return None
