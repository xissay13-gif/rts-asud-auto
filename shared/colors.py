"""
shared/colors.py — ANSI-цвета для консольного вывода (Windows-safe через colorama).

Используется только для print() в консоль — лог-файл остаётся плоским.
"""

try:
    from colorama import init as _colorama_init, Fore, Style
    _colorama_init()
    _HAS_COLOR = True
except ImportError:
    # Без colorama — все helpers возвращают строку без изменений
    class _NoColor:
        def __getattr__(self, _): return ""
    Fore = _NoColor()
    Style = _NoColor()
    _HAS_COLOR = False


def green(s):  return f"{Fore.GREEN}{s}{Style.RESET_ALL}" if _HAS_COLOR else str(s)
def yellow(s): return f"{Fore.YELLOW}{s}{Style.RESET_ALL}" if _HAS_COLOR else str(s)
def red(s):    return f"{Fore.RED}{s}{Style.RESET_ALL}" if _HAS_COLOR else str(s)


def status_colored(status):
    """Возвращает раскрашенную метку статуса фиксированной ширины (8 симв)."""
    label = f"{status:<8}"
    if status == "OK":
        return green(label)
    if status in ("DRAFT", "EXCLUDED"):
        return yellow(label)
    if status == "FAILED" or status == "FAIL":
        return red(label)
    # DUPLICATE и прочее — без цвета
    return label
