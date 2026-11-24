"""Folder → execution task → reply file(s) in a saved report."""

import logging
import os
import time
from pathlib import Path

from selenium.webdriver.edge.webdriver import WebDriver as Edge
from selenium.webdriver.edge.service import Service as EdgeService

from shared import config as cfg
from shared.asud_reports import (
    DEFAULT_REPORT, account_identity, add_reply_report, close_card, open_execution,
    wait_reports_home,
)
from shared.asud_resolution import switch_account
from shared.reply_files import ReplyJournal, folder_lock, scan_folder
from shared.ui import set_driver_timeout

log = logging.getLogger('asud.replies')


def process_job(driver, journal, account, job, text, timeout=60):
    fingerprints = job.fingerprints()
    previous = journal.previous(account, job.number)
    if previous and previous['status'] != 'failed':
        if previous['status'] == 'saved' and previous['files'] == fingerprints:
            log.info('%s: отчёт уже сохранён, пропускаю', job.number)
            return 'skip'
        log.error('%s: требуется проверка в АСУД: предыдущий статус %s; '
                  'файлы не отправлены повторно', job.number, previous['status'])
        return 'review'

    resolved = None
    saving = False
    try:
        if account_identity(driver)[0] != account:
            raise RuntimeError('Учётная запись изменилась')
        resolved = open_execution(driver, job.number)

        def before_save():
            nonlocal saving
            if job.fingerprints() != fingerprints:
                raise RuntimeError('Файлы изменились во время загрузки; сохранение отменено')
            journal.record(account, job, 'pending', fingerprints, resolved_number=resolved)
            saving = True

        add_reply_report(driver, account, resolved, job.files, text, before_save, timeout)
        journal.record(account, job, 'saved', fingerprints, resolved_number=resolved)
    except Exception as exc:
        if not saving:
            journal.record(account, job, 'failed', fingerprints, str(exc), resolved)
        # After a save attempt leave pending intact even when the confirmation
        # or journal write fails. Never turn an ambiguous save into a retry.
        raise
    log.info('%s: отчёт сохранён, вложений: %s', resolved, len(job.files))
    close_card(driver)
    return 'saved'


def run_folder(driver, folder, account, text, *, timeout=60, stable_seconds=5):
    jobs, errors = scan_folder(folder)
    totals = {'saved': 0, 'skip': 0, 'review': 0, 'waiting': 0, 'invalid': len(errors)}
    for error in errors:
        log.error('%s', error)
    journal = ReplyJournal(folder)
    for job in jobs:
        if any(time.time() - file.stat().st_mtime < stable_seconds for file in job.files):
            log.info('%s: файлы ещё копируются; обработка при следующем запуске', job.number)
            totals['waiting'] += 1
            continue
        status = process_job(driver, journal, account, job, text, timeout)
        totals[status] += 1
    return totals


def main():
    settings = cfg.load()
    options = settings.get('reply_reports') or {}
    folder = (os.environ.get('ASUD_REPLY_FOLDER') or options.get('folder') or '').strip()
    if not folder:
        folder = input('Папка с готовыми ответами: ').strip().strip('"')
    folder = Path(folder).expanduser().resolve(strict=True)
    if not folder.is_dir():
        raise ValueError(f'Это не папка: {folder}')
    text = str(options.get('report_text', DEFAULT_REPORT))
    if not text.strip() or len(text.encode('utf-16-le')) // 2 > 2000:
        raise ValueError('Текст отчёта должен содержать от 1 до 2000 символов')
    target_account = os.environ.get('ASUD_REPLY_ACCOUNT') or options.get('switch_to_user', '')
    watch = os.environ.get('ASUD_REPLY_WATCH') == '1'
    interval = max(5, int(options.get('poll_interval_sec', 300)))
    timeout = max(5, int(options.get('upload_timeout_sec', 90)))
    stable_seconds = max(0, int(options.get('file_stable_seconds', 5)))
    cfg.setup_file_logger('attach_replies')
    log.info('Ответы из папки: %s', folder)
    with folder_lock(folder):
        ReplyJournal(folder)  # Fail on a corrupt journal before starting Edge.
        driver_path = Path(cfg.get_base_dir()) / 'msedgedriver.exe'
        if not driver_path.is_file():
            raise FileNotFoundError(f'Поместите msedgedriver.exe рядом с программой: {driver_path}')
        # Explicit import is required for PyInstaller: webdriver.Edge is lazy
        # in Selenium and an otherwise valid EXE can omit its implementation.
        driver = Edge(service=EdgeService(str(driver_path)),
                      options=cfg.build_edge_options())
        try:
            set_driver_timeout(driver, settings.get('asud_load_timeout_sec', 120))
            driver.get(settings.get('asud_url', cfg.DEFAULTS['asud_url']))
            wait_reports_home(driver, settings.get('asud_load_timeout_sec', 120))
            if target_account and not switch_account(driver, target_account):
                raise RuntimeError(f'Не удалось выбрать учётную запись {target_account}')
            if target_account:
                wait_reports_home(driver, settings.get('asud_load_timeout_sec', 120))
            account, name = account_identity(driver)
            log.info('Исполнитель: %s (%s)', name, account)
            while True:
                totals = run_folder(driver, folder, account, text,
                                    timeout=timeout, stable_seconds=stable_seconds)
                log.info('Сохранено: %(saved)s; ранее обработано: %(skip)s; '
                         'требует проверки: %(review)s; ожидает копирования: %(waiting)s; '
                         'неверные имена/файлы: %(invalid)s', totals)
                if not watch:
                    if totals['review'] or totals['invalid'] or totals['waiting']:
                        raise RuntimeError('Обработаны не все ответы. Причины указаны в журнале.')
                    break
                time.sleep(interval)
        except KeyboardInterrupt:
            log.info('Обработка остановлена. Повторный запуск учитывает журнал.')
        finally:
            driver.quit()
