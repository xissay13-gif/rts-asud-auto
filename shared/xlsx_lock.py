"""
shared/xlsx_lock.py — Кооперативный file-lock для xlsx-реестров.

Используется когда несколько процессов читают/пишут один и тот же реестр:
  • P1 (регистратор) дописывает новые строки
  • P2 (zhkh-daemon под Басмановым) обновляет «Отписано Халецкой»
  • P3 (clean-resolutions под Халецкой) обновляет «Отписано в округ»

Простая стратегия: рядом с xlsx создаётся `<xlsx>.lock` через атомарный
exclusive-create. Если файл уже есть — ждём 200мс и пробуем. Защита от
stale-lock'а (если процесс упал не сняв lock): если .lock старше
60 секунд — считаем брошенным и забираем.

PermissionError (юзер открыл xlsx в Excel) lock не покрывает — там нужен
свой retry на уровне openpyxl-вызовов.

Используется как контекст-менеджер:

    from shared.xlsx_lock import xlsx_lock
    with xlsx_lock(path, timeout=30):
        # ... openpyxl read+modify+save ...
"""

import os
import time
import logging
from contextlib import contextmanager

log = logging.getLogger("asud")

LOCK_STALE_SECONDS = 60
LOCK_POLL_INTERVAL = 0.2


def is_xlsx_busy(xlsx_path):
    """Не-блокирующая проверка: занят ли xlsx другим процессом прямо сейчас.

    Возвращает True если рядом с xlsx есть свежий .lock-файл (моложе
    LOCK_STALE_SECONDS). Используется для приоритизации регистрации:
    daemon-resolutions видит что регистратор пишет в реестр, пропускает
    его на этом тике и берётся за другие файлы.

    NB: это подсказка. Между проверкой и реальной попыткой взять
    lock остаётся race-window — если другой процесс пробрался в этот
    промежуток, наша попытка acquire будет ждать обычным образом.
    """
    if not xlsx_path:
        return False
    lock_path = xlsx_path + ".lock"
    try:
        age = time.time() - os.path.getmtime(lock_path)
        return age < LOCK_STALE_SECONDS
    except OSError:
        return False


@contextmanager
def xlsx_lock(xlsx_path, timeout=30):
    """Кооперативный exclusive lock на xlsx через сосед-файл `<path>.lock`.

    Внутри ждёт пока удастся взять lock (до timeout секунд). На выходе из
    контекста — снимает. Если процесс упадёт внутри — lock останется до
    LOCK_STALE_SECONDS (60s), потом следующий захотевший его заберёт.

    Бросает TimeoutError если за timeout не удалось взять.
    """
    if not xlsx_path:
        # Нет пути — просто отдаём управление без локов
        yield
        return

    lock_path = xlsx_path + ".lock"
    end = time.monotonic() + max(timeout, 0)
    acquired = False

    while not acquired:
        try:
            # Atomic exclusive create: успех = мы первые
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, str(os.getpid()).encode())
            finally:
                os.close(fd)
            acquired = True
            break
        except FileExistsError:
            # Lock уже есть — проверим не stale ли
            try:
                age = time.time() - os.path.getmtime(lock_path)
                if age > LOCK_STALE_SECONDS:
                    log.warning(f"[xlsx_lock] {lock_path}: stale ({age:.0f}с), удаляю")
                    try:
                        os.remove(lock_path)
                    except OSError:
                        pass
                    continue  # сразу попробуем взять
            except OSError:
                pass
        except OSError as e:
            log.warning(f"[xlsx_lock] неожиданная ошибка при создании {lock_path}: {e}")

        if time.monotonic() >= end:
            raise TimeoutError(
                f"Не удалось взять lock {lock_path} за {timeout}с")
        time.sleep(LOCK_POLL_INTERVAL)

    try:
        yield
    finally:
        try:
            os.remove(lock_path)
        except OSError as e:
            log.debug(f"[xlsx_lock] не удалось удалить {lock_path}: {e}")
