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
    "outlook_dir": "",  # пусто = не искать .msg, пользователь укажет интерактивно
    "addressees": [
        "Басманов Александр Владимирович",
    ],
    "unknown_correspondent": "Неизвестный Неизвестный Неизвестный",
    "delivery_method": "Электронная почта",
    "sheet_name": "Лист2",
    # email-daemon (--watch)
    "email_watch_interval_sec": 30,
    "email_max_retries": 3,
    # Если True — обработанные .msg удаляются вместо переноса в Завершено/.
    # По умолчанию True — в daemon-режиме Завершено/ разрастается до тысяч
    # файлов, и юзер всё равно туда не лазит. Можно явно поставить False
    # глобально или в конкретном пресете если нужен бэкап.
    "delete_after_done": True,
    # Таймаут загрузки страниц ASUD (page_load + urllib3 client). 300с = 5 минут.
    # Поднят с дефолтных 120с т.к. ASUD иногда долго инициализируется.
    "asud_load_timeout_sec": 300,
    # Пресеты сценариев — список словарей. Поля:
    #   name          — отображение в меню
    #   mode          — 'mix' / 'smart' / 'auto-create' / 'email'
    #   folder        — источник .msg-писем для email-flow
    #   xlsx          — путь к реестру для mix/auto-create/smart (опциональный)
    #   outlook_dir   — папка где искать .msg по Link для mix-вложений
    #                    (опциональный; если не задан — спрашиваем интерактивно)
    #   output_suffix — добавляется в имя per-date реестра, чтобы
    #                    Registered/2026-05-12_ОЭК_резолюции.xlsx и
    #                    Registered/2026-05-12_ТЭС_резолюции.xlsx не конфликтовали.
    # Чтобы переопределить без пересборки — создать settings.json или
    # config.json рядом с exe (см. settings.json.example).
    "presets": [
        {
            "name": "ОЭК — Smart (черновики)",
            "mode": "smart",
            "folder": "D:\\OutlookSubjects\\ОЭК",
            "outlook_dir": "D:\\OutlookSubjects\\ОЭК",
            "output_suffix": "ОЭК",
        },
        {
            "name": "ТЭС — Mix (с регистрацией)",
            "mode": "mix",
            "folder": "D:\\OutlookSubjects\\ТЭС",
            "outlook_dir": "D:\\OutlookSubjects\\ТЭС",
            "output_suffix": "ТЭС",
        },
        {
            "name": "ГИС ЖКХ — Mix (с регистрацией)",
            "mode": "mix",
            "folder": "D:\\OutlookSubjects\\ГИСЖКХ",
            "outlook_dir": "D:\\OutlookSubjects\\ГИСЖКХ",
            "output_suffix": "ГИСЖКХ",
        },
    ],
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
    """Загружает config.json + settings.json, мержит с дефолтами.

    Порядок применения (поздние перекрывают ранние):
      1. DEFAULTS (зашиты в код)
      2. config.json (legacy, для обратной совместимости)
      3. settings.json (рекомендуемый)

    Можно использовать любой из двух файлов рядом с exe. Если оба
    присутствуют — settings.json побеждает.
    """
    cfg = dict(DEFAULTS)
    base = get_base_dir()
    for fname in ("config.json", "settings.json"):
        path = os.path.join(base, fname)
        if os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    user_cfg = json.load(f)
                cfg.update(user_cfg)
                log.info(f"Конфиг загружен: {path}")
            except Exception as e:
                log.warning(f"Ошибка чтения {fname}: {e}, пропускаем")
    return cfg


def setup_file_logger(mode_name="asud"):
    """Подключает FileHandler с DEBUG-уровнем в папку Logs/ рядом с exe.
    Имя файла: Logs/asud_<mode>_<YYYYMMDD_HHMMSS>.txt

    Консоль показывает только высокоуровневый flow (логгер 'asud') —
    шумные под-логгеры (asud.ui, asud.correspondent, asud.attach, asud.okrug)
    идут ТОЛЬКО в txt-файл. Это убирает из консоли «Клик: ...», «Ожидаю: ...»,
    «Файл присоединён», «Корреспондент выбран» — оставляет понятный flow:

    12:01:15 [INFO] ДОКУМЕНТ 1/50: ...
    12:01:32 [INFO] Документ 1/50 ЗАРЕГИСТРИРОВАН: ОРТС/8/...
    12:01:34 [INFO] Документ 1/50 НА РЕЗОЛЮЦИИ

    Возвращает путь к лог-файлу (или None если упало).
    Вызывается из main() каждого flow.
    """
    from datetime import datetime
    try:
        logs_dir = os.path.join(get_base_dir(), "Logs")
        os.makedirs(logs_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(logs_dir, f"asud_{mode_name}_{ts}.txt")
        fh = logging.FileHandler(path, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter(
            '%(asctime)s.%(msecs)03d [%(levelname)s] %(name)s: %(message)s',
            datefmt='%H:%M:%S'))
        # Корневой логгер — DEBUG (чтобы файл получал всё). Консольные
        # хендлеры остаются на INFO.
        root = logging.getLogger()
        root.setLevel(logging.DEBUG)
        for h in root.handlers:
            if not isinstance(h, logging.FileHandler):
                h.setLevel(logging.INFO)
        root.addHandler(fh)

        # Шумные под-логгеры: НЕ propagate в root (значит не попадают в консоль),
        # но получают свой file-handler чтобы по-прежнему писаться в txt.
        for sub_name in ('asud.ui', 'asud.correspondent', 'asud.attach',
                          'asud.okrug', 'asud.config'):
            sub = logging.getLogger(sub_name)
            sub.setLevel(logging.DEBUG)
            sub.propagate = False
            sub.addHandler(fh)

        logging.getLogger("asud").info(f"Подробный лог пишется в: {path}")
        return path
    except Exception as e:
        logging.getLogger("asud").warning(f"Не удалось создать файл лога: {e}")
        return None


def build_edge_options():
    """Собирает EdgeOptions с учётом ASUD_HEADLESS env-переменной.

    Если ASUD_HEADLESS=1 — добавляется --headless=new + фиксированный
    window-size (для правильного рендера форм АСУД). Иначе обычный
    --start-maximized.
    """
    from selenium.webdriver.edge.options import Options as EdgeOptions
    options = EdgeOptions()
    options.page_load_strategy = "eager"
    if os.environ.get('ASUD_HEADLESS') == '1':
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        logging.getLogger("asud").info("Edge запущен в HEADLESS режиме")
    else:
        options.add_argument("--start-maximized")
    options.add_argument("--auth-server-whitelist=*.interrao.ru")
    options.add_argument("--auth-negotiate-delegate-whitelist=*.interrao.ru")
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    return options


import threading

_keep_awake_thread = None
_keep_awake_stop = threading.Event()


def _keep_awake_heartbeat():
    """Фоновый daemon-поток: раз в 30с переподтверждает SetThreadExecutionState.
    Без heartbeat'а флаг иногда «забывался» если основной поток долго висел
    внутри Selenium-операции (особенно на корпоративных лэптопах с агрессивной
    энергосберегайкой). Демон сам умирает при выходе процесса."""
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
    """Блокирует/разблокирует Windows-таймер автосна на время процесса.

    Флаги:
      ES_CONTINUOUS        — постоянный (а не одноразовый) запрос
      ES_SYSTEM_REQUIRED   — не уходить в системный сон/гибернацию
      ES_DISPLAY_REQUIRED  — дополнительно не отключать дисплей. Часто
                              сильнее ловит overzealous энергосбережение
                              на корпоративных машинах, где одного
                              ES_SYSTEM_REQUIRED бывает недостаточно

    Дополнительно поднимает daemon-поток, который переподтверждает запрос
    каждые 30 секунд (некоторые версии Windows сбрасывали флаг при долгих
    блокирующих операциях).

    На Linux/macOS — no-op. Реверт автоматический при выходе процесса
    (можно явно вызвать с enabled=False).
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
