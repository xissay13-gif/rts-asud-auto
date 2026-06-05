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
    # Таймаут загрузки страниц ASUD (page_load + urllib3 client). 300с = 5 минут.
    # Поднят с дефолтных 120с т.к. ASUD иногда долго инициализируется.
    "asud_load_timeout_sec": 300,
    # ZHKH-второй проход (после ГИСЖКХ-пресета):
    # zhkh_complete_user      — фрагмент ФИО для переключения учётки
    # zhkh_initial_user       — фрагмент ФИО для возврата (пусто = не возвращать)
    # zhkh_sidebar_section    — пункт sidebar после переключения
    # zhkh_executor_fio       — ФИО исполнителя для резолюции
    # zhkh_content_template   — шаблон «Содержание»
    # zhkh_require_report     — включать тоггл «Требуется отчёт»
    # zhkh_control_resolution — включать тоггл «Контрольная резолюция»
    # zhkh_fallback_days      — если планируемая дата в прошлом,
    #                            ставим today + N календарных дней
    "zhkh_complete_user": "Басманов",
    "zhkh_initial_user": "",
    "zhkh_sidebar_section": "На резолюцию",
    "zhkh_executor_fio": "Халецкая Юлия Владимировна",
    "zhkh_content_template": "Для работы",
    "zhkh_require_report": True,
    "zhkh_control_resolution": True,
    "zhkh_fallback_days": 3,
    # Пресеты сценариев — список словарей с name + mode + folder.
    # Зашиты дефолтные пути для ОЭК (smart) и ТЭС (mix). Чтобы переопределить
    # без пересборки — создать settings.json рядом с exe (см. settings.json.example).
    # output_suffix — добавляется в имя реестра, чтобы при параллельном
    # запуске двух пресетов файлы не конфликтовали:
    #   Registered/2026-05-12_ОЭК_резолюции.xlsx
    #   Registered/2026-05-12_ТЭС_резолюции.xlsx
    "presets": [
        {
            "name": "ОЭК — Smart (черновики)",
            "mode": "smart",
            "folder": "D:\\OutlookSubjects\\ОЭК",
            "output_suffix": "ОЭК",
        },
        {
            "name": "ТЭС — Mix (с регистрацией)",
            "mode": "mix",
            "folder": "D:\\OutlookSubjects\\ТЭС",
            "output_suffix": "ТЭС",
        },
        {
            "name": "ГИС ЖКХ — Mix (с регистрацией)",
            "mode": "mix",
            "folder": "D:\\OutlookSubjects\\ГИСЖКХ",
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


def keep_system_awake(enabled=True):
    """Блокирует/разблокирует Windows-таймер автосна на время процесса.
    На Linux/macOS — no-op. Реверт автоматический при выходе процесса
    (можно явно вызвать с enabled=False).
    """
    if not sys.platform.startswith('win'):
        return
    try:
        import ctypes
        ES_CONTINUOUS = 0x80000000
        ES_SYSTEM_REQUIRED = 0x00000001
        flags = (ES_CONTINUOUS | ES_SYSTEM_REQUIRED) if enabled else ES_CONTINUOUS
        ctypes.windll.kernel32.SetThreadExecutionState(flags)
        if enabled:
            logging.getLogger("asud").info("Сон Windows заблокирован на время работы")
    except Exception as e:
        logging.getLogger("asud").debug(f"keep_system_awake не сработал: {e}")
