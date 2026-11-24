"""
okrug_parser.py — определение округа Омска по адресу из текста.

Используется и в clean-mix (для записи в output-таблицу), и в clean-resolutions
(для сопоставления с DEFAULT_OKRUG_MAP). Источник — addresses.csv с колонками
street;house;okrug, который запекается в exe через PyInstaller --add-data.

Стратегия поиска адреса в TextBody:
  1. Блок "Суть обращения: ..." (там часто реальный адрес проблемы)
  2. Строка "Почтовый адрес: ..." (адрес регистрации, fallback)
  3. Весь TextBody (последний fallback)

Матчинг — поиск по списку известных улиц (~465). Сначала длинные,
чтобы не зацепить "Ленина" в "100 лет Ленина". Дом ищется в окне
50 символов после улицы.

На Почта_ТЭС.xlsx даёт ~57% точного автоматического определения,
~1% ошибочно, остальное — улица отсутствует в БД или нестандартный
формат записи.
"""

import os
import re
import sys
import csv
import logging
from collections import defaultdict

log = logging.getLogger("asud.okrug")

_PREFIX_RE = re.compile(
    r'^(?:г\s*омск\s*,?\s*)?'
    r'(?:улица|ул\.?|проспект|пр[-]?т\.?|пр[-]?кт\.?|пр\.?|переулок|пер\.?|'
    r'бульвар|б[-]?р\.?|площадь|пл\.?|шоссе|ш\.?|набережная|наб\.?|'
    r'линия|тупик|проезд|пр[-]?д\.?|микрорайон|мкр\.?|'
    r'пос[её]лок|пос\.?|пгт\.?|п\.?|городок|г[-\s]*к\.?)\s*',
    re.IGNORECASE)

_HOUSE_RE = re.compile(r'(\d+[а-я]?)', re.IGNORECASE)


def _norm_text(s):
    if not s:
        return ''
    s = str(s).lower().replace('ё', 'е')
    s = re.sub(r'[«»"\'`]', '', s)
    s = re.sub(r'[.,;:()\\/]+', ' ', s)
    # "3-я молодежная" / "3-й разъезд" → "3 молодежная" / "3 разъезд".
    s = re.sub(r'(\d+)[\s-]*(?:я|й|ая|ой|ый)(?=\s)', r'\1', s)
    # Частая склейка в письмах: "13линия" → "13 линия".
    s = re.sub(r'(\d+)(?=(?:линия|разъезд)\b)', r'\1 ', s)
    return re.sub(r'\s+', ' ', s).strip()


def _norm_street_name(s):
    return _PREFIX_RE.sub('', _norm_text(s)).strip()


def _norm_house(s):
    if not s:
        return ''
    m = _HOUSE_RE.search(str(s).lower().replace('ё', 'е'))
    return m.group(1) if m else ''


_index = None
_streets_sorted = None


def _addresses_csv_path(base_dir_fn=None):
    """Возвращает путь к addresses.csv: внутри exe (через _MEIPASS),
    либо рядом с exe / py-скриптом."""
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        path = os.path.join(meipass, 'addresses.csv')
        if os.path.exists(path):
            return path
    if base_dir_fn:
        path = os.path.join(base_dir_fn(), 'addresses.csv')
        if os.path.exists(path):
            return path
    # last resort — current dir
    if os.path.exists('addresses.csv'):
        return 'addresses.csv'
    return None


def _build_index(base_dir_fn=None):
    global _index, _streets_sorted
    if _index is not None:
        return _index, _streets_sorted
    path = _addresses_csv_path(base_dir_fn)
    if not path:
        log.warning("addresses.csv не найден — авто-определение округа отключено")
        _index = {}
        _streets_sorted = []
        return _index, _streets_sorted
    try:
        idx = defaultdict(set)
        with open(path, encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=';')
            header = next(reader, None)
            if not header:
                _index = {}; _streets_sorted = []
                return _index, _streets_sorted
            cols = {h.lower(): i for i, h in enumerate(header)}
            if 'street' not in cols or 'house' not in cols or 'okrug' not in cols:
                log.warning(f"Неподдерживаемый формат addresses.csv: {header}")
                _index = {}; _streets_sorted = []
                return _index, _streets_sorted
            for row in reader:
                if not row or len(row) <= max(cols['street'], cols['house'], cols['okrug']):
                    continue
                s = _norm_street_name(row[cols['street']])
                h = _norm_house(row[cols['house']])
                o = row[cols['okrug']].strip()
                if s and h and o:
                    aliases = {s}
                    # Номерные улицы встречаются в обоих порядках:
                    # "Молодёжная 3-я" ↔ "3-я Молодёжная".
                    trailing = re.match(r'^(.+?)\s+(\d+)$', s)
                    leading = re.match(r'^(\d+)\s+(.+)$', s)
                    if trailing:
                        aliases.add(f"{trailing.group(2)} {trailing.group(1).strip()}")
                    if leading:
                        aliases.add(f"{leading.group(2).strip()} {leading.group(1)}")
                    for alias in aliases:
                        idx[alias].add((h, o))
        _index = idx
        _streets_sorted = sorted(idx.keys(), key=lambda x: -len(x))
        log.info(f"Street index: {len(idx)} улиц")
    except Exception as e:
        log.warning(f"Ошибка построения street index: {e}")
        _index = {}
        _streets_sorted = []
    return _index, _streets_sorted


def _street_candidates(text, idx, sorted_streets):
    """Все известные улицы в тексте с ближайшим номером дома, если он есть."""
    norm = _norm_text(text)
    candidates = []
    for street in sorted_streets:
        if len(street) < 3:
            continue
        pos = 0
        while True:
            i = norm.find(street, pos)
            if i < 0:
                break
            left_ok = i == 0 or not norm[i - 1].isalnum()
            end = i + len(street)
            right_ok = end == len(norm) or not norm[end].isalnum()
            if left_ok and right_ok:
                tail = norm[end:end + 50]
                # «Омская обл.» — административный регион, а не улица Омская.
                if re.match(r'\s*(?:обл(?:асть|асти)?|район|р-н)\b', tail):
                    pos = i + 1
                    continue
                # «Омская энергосбытовая компания» — тоже не адрес. Для улицы
                # Омской нужен явный префикс «ул.» либо номер сразу после неё.
                if street == 'омская':
                    prefix = norm[max(0, i - 12):i]
                    explicit_street = bool(re.search(r'(?:ул|улица)\s*$', prefix))
                    direct_house = bool(re.match(r'\s*(?:д(?:ом)?\s*)?\d', tail))
                    if not explicit_street and not direct_house:
                        pos = i + 1
                        continue
                m = re.search(r'\b(\d+[а-я]?)', tail)
                # В свободном тексте встречаются слова «мира», «труда» и т.п.
                # Считаем их адресом только при явном типе улицы либо когда
                # номер дома расположен рядом, а не через полпредложения.
                prefix = norm[max(0, i - 24):i]
                explicit_street = bool(re.search(
                    r'(?:улица|ул|проспект|пр т|пр кт|переулок|пер|бульвар|'
                    r'б р|площадь|пл|шоссе|наб|линия|тупик|проезд|микрорайон|'
                    r'мкр|поселок|пос|пгт|городок|г к)\s*$', prefix))
                house_is_near = bool(m and m.start() <= 18)
                if not explicit_street and not house_is_near:
                    pos = i + 1
                    continue
                candidates.append((i, street, m.group(1) if m else None))
            pos = i + 1
    candidates.sort(key=lambda item: (item[0], -len(item[1])))
    return candidates


def _find_street_house(text, idx, sorted_streets):
    """Обратная совместимость: первый найденный кандидат улица/дом."""
    candidates = _street_candidates(text, idx, sorted_streets)
    return candidates[0][1:] if candidates else (None, None)


def _match_fragment(fragment, idx, sorted_streets):
    """Возвращает точное и безопасное резервное совпадение для фрагмента."""
    exact = []
    fallback = []
    for _, street, house in _street_candidates(fragment, idx, sorted_streets):
        street_rows = idx.get(street, ())
        districts = {district for _, district in street_rows}
        if house:
            house_num = re.match(r'\d+', house)
            house_num = house_num.group(0) if house_num else house
            for known_house, district in street_rows:
                known_num = re.match(r'\d+', known_house)
                known_num = known_num.group(0) if known_num else known_house
                if known_house == house or known_num == house_num:
                    exact.append((district, street, house))
        # Резерв разрешён только когда вся улица в БД относится к одному округу.
        if len(districts) == 1:
            fallback.append((next(iter(districts)), street, house))
    exact_districts = {item[0] for item in exact}
    if len(exact_districts) == 1:
        return exact[0], None
    fallback_districts = {item[0] for item in fallback}
    if not exact and len(fallback_districts) == 1:
        return None, fallback[0]
    return None, None


def okrug_from_textbody(textbody, base_dir_fn=None, subject=None,
                        return_reason=False):
    """Извлекает округ из тела/темы письма.

    ``return_reason=True`` возвращает ``(округ, причина)`` для диагностики.
    Резерв по одной улице применяется только если все её дома в адресной БД
    находятся в одном округе.
    """
    if not textbody and not subject:
        return (None, "нет текста") if return_reason else None
    idx, sorted_streets = _build_index(base_dir_fn)
    if not idx:
        return (None, "адресная база недоступна") if return_reason else None
    text = str(textbody or '')
    fragments = []
    m = re.search(r'суть\s+обращени[яе]\s*:?\s*([\s\S]+?)(?:\n\s*\n|$)',
                  text, re.IGNORECASE)
    if m:
        fragments.append(('суть', m.group(1)))
    m = re.search(r'почтов[а-я]+\s+адрес[а-я]*\s*:\s*([^\n]+)',
                  text, re.IGNORECASE)
    if m:
        fragments.append(('почт', m.group(1)))
    if subject:
        fragments.append(('тема', str(subject)))
    if text:
        fragments.append(('тело', text))

    fallback_hits = []
    for name, frag in fragments:
        exact, fallback = _match_fragment(frag, idx, sorted_streets)
        if exact:
            district, street, house = exact
            reason = f"точно, {name}: {street} {house}"
            log.info(f"Округ {district} ({reason})")
            return (district, reason) if return_reason else district
        if fallback:
            fallback_hits.append((name, fallback))

    fallback_districts = {item[1][0] for item in fallback_hits}
    if len(fallback_districts) == 1:
        name, (district, street, house) = fallback_hits[0]
        suffix = f" {house}" if house else ""
        reason = f"по однозначной улице, {name}: {street}{suffix}"
        log.info(f"Округ {district} ({reason})")
        return (district, reason) if return_reason else district

    reason = ("несколько возможных округов" if fallback_hits
              else "адрес не найден в базе")
    log.info(f"Округ не определён: {reason}")
    return (None, reason) if return_reason else None
