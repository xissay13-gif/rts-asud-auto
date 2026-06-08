"""
config.py — Настройки АСУД-скрипта.

Читает config.json рядом с exe (если есть), иначе дефолты.
Менять настройки можно без пересборки — просто правишь JSON.
"""

import os
import sys
import json
import logging

log = logging.getLogger("asud.config")

# Дефолтные значения
DEFAULTS = {
    "asud_url": "https://asud.interrao.ru/asudik/",
    "timeout": 20,
    "outlook_dir": r"D:\OutlookSubjects\ТЭС",
    "addressees": [
        "Басманов Александр Владимирович",
        "Халецкая Юлия Владимировна",
    ],
    "unknown_correspondent": "Неизвестный Неизвестный Неизвестный",
    "delivery_method": "Электронная почта",
    "sheet_name": "Лист2",
}

# Маппинг индекса из Excel → название вида в АСУД
DOC_TYPE_MAP = {
    1: "Указы, распоряжения Президента Российской Федерации",
    2: "Документы Администрации Президента",
    3: "Документы Правительства Российской Федерации",
    4: "Документы Федеральных органов исполнительной и законодательной власти",
    5: "Письма юридических лиц",
    6: "Письма компаний Топливно-энергетического комплекса",
    7: "Документы органов законодательной и исполнительной власти субъектов",
    8: "Письма, заявления и жалобы граждан, акционеров",
}


def get_base_dir():
    """Папка где лежит exe/скрипт."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load():
    """Загружает config.json, мержит с дефолтами."""
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


def setup_file_logger(mode_name="asud"):
    """Подключает FileHandler с DEBUG-уровнем рядом с exe.
    Имя файла: asud_<mode>_<YYYYMMDD_HHMMSS>.txt

    Консоль остаётся на INFO (как раньше — только описания действий),
    в файл идёт DEBUG (всё подробно для разбора зависаний).

    Возвращает путь к лог-файлу (или None если упало).
    Вызывается из main() каждого flow.
    """
    from datetime import datetime
    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(get_base_dir(), f"asud_{mode_name}_{ts}.txt")
        fh = logging.FileHandler(path, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter(
            '%(asctime)s.%(msecs)03d [%(levelname)s] %(name)s: %(message)s',
            datefmt='%H:%M:%S'))
        # Корневой логгер — DEBUG (чтобы файл получал всё). Консольные
        # хендлеры остаются на INFO — пользователь видит только описания
        # действий, без debug-шума.
        root = logging.getLogger()
        root.setLevel(logging.DEBUG)
        for h in root.handlers:
            if not isinstance(h, logging.FileHandler):
                h.setLevel(logging.INFO)
        root.addHandler(fh)
        logging.getLogger("asud").info(f"Подробный лог пишется в: {path}")
        return path
    except Exception as e:
        logging.getLogger("asud").warning(f"Не удалось создать файл лога: {e}")
        return None


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
    переподтверждает запрос каждые 30 секунд (без heartbeat'а флаг
    иногда «забывался» при долгих Selenium-операциях на корпоративных
    машинах с агрессивной энергосберегайкой).

    На Linux/macOS — no-op. Реверт автоматический при выходе процесса.
    """
    global _keep_awake_thread
    if not sys.platform.startswith('win'):
        return
    log = logging.getLogger("asud")
    try:
        import ctypes
        ES_CONTINUOUS = 0x80000000
        ES_SYSTEM_REQUIRED = 0x00000001
        ES_DISPLAY_REQUIRED = 0x00000002
        if enabled:
            flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
            ctypes.windll.kernel32.SetThreadExecutionState(flags)
            log.info("Сон Windows и таймер дисплея заблокированы (+heartbeat 30с)")
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
            log.info("Сон Windows разблокирован")
    except Exception as e:
        log.debug(f"keep_system_awake не сработал: {e}")
