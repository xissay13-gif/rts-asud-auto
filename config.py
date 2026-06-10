"""
config.py — Настройки скрипта выдачи резолюций (clean-resolutions).

Читает config.json рядом с exe (если есть), иначе дефолты.
"""

import os
import sys
import json
import logging

log = logging.getLogger("asud.config")

DEFAULTS = {
    "asud_url": "https://asud.interrao.ru/asudik/",
    "timeout": 20,
    # Интервал опроса xlsx-реестров в --watch (daemon-режим), в секундах.
    # Переопределяется флагом --poll-interval.
    "poll_interval_sec": 300,
    # Если True — на каждом тике проверяется только ОДИН реестр из preset.watch
    # в порядке round-robin (ОЭК → ТЭС → ГИСЖКХ → ОЭК → ...). Удобно когда
    # хочется чёткое расписание. Если False (по умолчанию) — каждый тик
    # проверяются все реестры, но только изменённые перечитываются.
    "watch_round_robin": False,
    # Учётка под которой работает Халецкая (на которую переключаемся при старте).
    # Поиск пункта в выпадашке профиля по подстроке этого ФИО.
    "target_account": "Халецкая",
    # Лист реестра с обращениями (колонки Link, Subject, TextBody, Тема, To, LS, ao, fio).
    "sheet_name": "Лист2",
    # Что писать в "Содержание" резолюции (выбор из выпадашки).
    "resolution_content": "Подготовить ответ",
    # На сколько РАБОЧИХ дней дать срок исполнения от сегодня.
    "workdays": 7,
    # Пункт левого сайдбара куда заходим (ровно как написано в АСУД).
    # Документы для выдачи резолюций у Халецкой попадают в раздел "Исполнение"
    # (id: CABINET_MENU__RECEIVED__ALL_ACTIVE__TO_EXECUTION).
    "sidebar_section": "Исполнение",
    # Включать ли тоггл "Требуется отчёт".
    "require_report": True,
    # Включать ли тоггл "Контрольная резолюция".
    "control_resolution": True,
    # Если задано — подменяет исполнителя для ВСЕХ строк реестра (используется
    # пресетом «ГИСЖКХ → Халецкая», где executor берётся не из округа, а одинаков).
    "force_executor": "",
    # Режим расчёта даты контрольного этапа: 'workdays' (по умолчанию, today + N рабочих)
    # или 'calendar' (today + N календарных). Преcет ZHKH-Халецкая использует 'calendar'.
    "stage_date_mode": "workdays",
    # Кол-во дней для контрольного этапа. Если пусто — используется workdays.
    "stage_date_days": 0,
    # Пресеты сценариев для меню выбора. Каждый пресет может перекрыть любые поля
    # выше (target_account, sidebar_section, force_executor, stage_date_mode/days,
    # resolution_content и т.д.). Меню показывается при старте если presets не пуст.
    #
    # Поле "watch" (опционально): список папок с реестрами для multi-folder
    # daemon-режима (--watch). Без него exe смотрит xlsx рядом с собой.
    # Формат: [{"dir": "путь", "xlsx_pattern": "имя или маска"}, ...].
    "presets": [
        {
            "name": "Округа → начальницам (под Халецкой)",
            "target_account": "Халецкая",
            "sidebar_section": "Исполнение",
            "force_executor": "",
            "resolution_content": "Подготовить ответ",
            "stage_date_mode": "workdays",
            "stage_date_days": 7,
        },
        {
            "name": "Басманов → Халецкая (все реестры)",
            "target_account": "Басманов",
            "sidebar_section": "На резолюцию",
            "force_executor": "Халецкая Юлия Владимировна",
            "resolution_content": "Для работы",
            "stage_date_mode": "calendar",
            "stage_date_days": 3,
        },
    ],
}

# Округ → ФИО начальника. Используется как fallback, если в реестре
# колонка fio пустая, но колонка ao заполнена.
DEFAULT_OKRUG_MAP = {
    "САО": "Гренц Екатерина Александровна",
    "ЦАО": "Емельянова Татьяна Николаевна",
    "ОАО": "Рендюк Юлия Павловна",
    "ЛАО": "Вырва Елена Анатольевна",
    "КАО": "Кравец Татьяна Александровна",
    # Полные названия — на случай если в реестре пишут так
    "Советский": "Гренц Екатерина Александровна",
    "Центральный": "Емельянова Татьяна Николаевна",
    "Октябрьский": "Рендюк Юлия Павловна",
    "Ленинский": "Вырва Елена Анатольевна",
    "Кировский": "Кравец Татьяна Александровна",
}


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load():
    cfg = dict(DEFAULTS)
    path = os.path.join(get_base_dir(), "config.json")
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                user_cfg = json.load(f)
            cfg.update(user_cfg)
            log.info(f"Конфиг загружен: {path}")
        except Exception as e:
            log.warning(f"Ошибка чтения config.json: {e}, используем дефолты")
    return cfg


# ============================================================
# Анти-сон Windows
# ============================================================

import threading

_keep_awake_thread = None
_keep_awake_stop = threading.Event()


def _keep_awake_heartbeat():
    """Daemon-поток: раз в 30с переподтверждает SetThreadExecutionState."""
    import ctypes
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    while not _keep_awake_stop.wait(30):
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(flags)
        except Exception:
            break


def keep_system_awake(enabled=True):
    """Блокирует Windows-таймер автосна на время процесса.

    Использует SetThreadExecutionState с ES_CONTINUOUS | ES_SYSTEM_REQUIRED
    | ES_DISPLAY_REQUIRED. Дополнительно поднимает daemon-поток который
    переподтверждает запрос каждые 30 секунд (на некоторых корпоративных
    машинах флаг сбрасывался при долгих Selenium-операциях).

    На Linux/macOS — no-op. Реверт автоматический при выходе процесса.
    """
    global _keep_awake_thread
    if not sys.platform.startswith('win'):
        return
    log_a = logging.getLogger("asud")
    try:
        import ctypes
        ES_CONTINUOUS = 0x80000000
        ES_SYSTEM_REQUIRED = 0x00000001
        ES_DISPLAY_REQUIRED = 0x00000002
        if enabled:
            flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
            ctypes.windll.kernel32.SetThreadExecutionState(flags)
            log_a.info("Сон Windows и таймер дисплея заблокированы (+heartbeat 30с)")
            if _keep_awake_thread is None or not _keep_awake_thread.is_alive():
                _keep_awake_stop.clear()
                _keep_awake_thread = threading.Thread(
                    target=_keep_awake_heartbeat,
                    daemon=True,
                    name="keep-awake-hb",
                )
                _keep_awake_thread.start()
        else:
            _keep_awake_stop.set()
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
            log_a.info("Сон Windows разблокирован")
    except Exception as e:
        log_a.debug(f"keep_system_awake не сработал: {e}")
