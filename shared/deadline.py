"""
shared/deadline.py — Расчёт срока отработки обращений ГИС ЖКХ.

Общий модуль: используется и при регистрации (P1, email-flow), и при
отписывании (P3, clean-resolutions — там лежит идентичная копия deadline.py).
Держать обе копии БАЙТ-В-БАЙТ одинаковыми, чтобы срок считался одинаково
на обоих этапах.

Правило (от пользователя АО «ОмскРТС»):
  РТС даёт 17 КАЛЕНДАРНЫХ дней на отработку с даты получения обращения.
  Если конец срока выпал на выходной (сб/вс) — переносим на ближайший
  рабочий день (понедельник).
  Но если срок по ГИС ЖКХ (планируемая/крайняя дата) РАНЬШE — берём его.
  Итог = min(сдвиг_с_выходных(дата_получения + 17 кал.дней), дата_ГИС).
"""

import re
from datetime import date, timedelta

CALENDAR_DAYS_DEFAULT = 17


def parse_ddmmyyyy(s):
    """DD.MM.YYYY (с временем после или без) → date. None если не распарсилось."""
    if not s:
        return None
    m = re.match(r'\s*(\d{2})\.(\d{2})\.(\d{4})', str(s))
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def shift_off_weekend(d):
    """Если дата выпала на сб/вс — сдвигает вперёд на ближайший рабочий день."""
    while d.weekday() >= 5:  # 5=сб, 6=вс
        d += timedelta(days=1)
    return d


def add_working_days(start, n):
    """start + n рабочих дней (сб/вс пропускаются). Возвращает date.
    Используется для overdue-кейсов (ближайший рабочий день и т.п.)."""
    cur = start
    added = 0
    while added < n:
        cur += timedelta(days=1)
        if cur.weekday() < 5:  # 0-4 = пн-пт
            added += 1
    return cur


def compute_deadline(receipt_str, gis_date_str, calendar_days=CALENDAR_DAYS_DEFAULT):
    """Срок отработки = min(сдвиг_с_выходных(дата_получения + N кал.дней), дата_ГИС).

    receipt_str — «Дата получения» (может быть с временем).
    gis_date_str — планируемая/крайняя дата из письма (DD.MM.YYYY) или None.

    N календарных дней; если итог выпал на выходной — переносится на
    ближайший рабочий день. Если срок ГИС раньше — берётся он (как есть).

    Возвращает строку DD.MM.YYYY или None если дату получения не распарсили.
    """
    receipt = parse_ddmmyyyy(receipt_str)
    if not receipt:
        return None
    rts = shift_off_weekend(receipt + timedelta(days=calendar_days))
    gis = parse_ddmmyyyy(gis_date_str)
    chosen = min(rts, gis) if gis else rts
    return chosen.strftime("%d.%m.%Y")
