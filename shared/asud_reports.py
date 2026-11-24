"""Add a reply report through the department head's execution lifecycle row.

Selectors follow the supplied five ASUD DOM snapshots. Saving is a single
attempt; a closed dialog alone is not evidence that the report was saved.
"""

import logging
import re
import time

from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from shared.asud_resolution import click_sidebar_section, filter_by_number
from shared.reply_files import number_matches
from shared.ui import click, dismiss_system_notifications


FULL_NUMBER = re.compile(r'(?<![\w/])[А-ЯЁA-Z]+/\d+/\d+(?![\w/])', re.IGNORECASE)
DEFAULT_REPORT = 'Ответ прикреплен.'
DIALOG_SELECTOR = 'div[class*="Css3WindowStyle-panel"]'
UPLOAD_SELECTOR = 'input[type="file"][name="_reportMultiFileUploadComponent"]'
log = logging.getLogger('asud.replies')


def report_wait(driver, timeout):
    """Only re-poll observations/text entry when GXT replaces DOM nodes."""
    return WebDriverWait(driver, timeout, poll_frequency=0.2,
                         ignored_exceptions=(StaleElementReferenceException,))


def wait_reports_home(driver, timeout=120):
    """An executor needs an account and execution menu, not document creation."""
    log.info('Жду учётную запись и меню заданий АСУД...')
    report_wait(driver, timeout).until(lambda d:
        d.execute_script("return document.readyState === 'complete'") and
        visible(d, '#app-acc-info[x-account-name]') and
        visible(d, '#CABINET_MENU__RECEIVED__ALL_ACTIVE__TO_EXECUTION'))
    if not dismiss_system_notifications(driver):
        raise RuntimeError('Системное уведомление перекрывает меню заданий')
    account_identity(driver)
    log.info('АСУД готов к прикреплению отчётов')


def visible(root, selector):
    result = []
    for element in root.find_elements(By.CSS_SELECTOR, selector):
        try:
            if element.is_displayed():
                result.append(element)
        except StaleElementReferenceException:
            continue
    return result


def unique(elements, description):
    if len(elements) > 1:
        raise RuntimeError(f'Неоднозначный выбор: {description} ({len(elements)})')
    return elements[0] if elements else None


def account_identity(driver):
    elements = driver.find_elements(By.XPATH,
        "//span[@id='app-acc-info' and not(ancestor::div[contains(@style,'z-index')])]")
    element = unique([el for el in elements if el.is_displayed()], 'учётная запись')
    if element is None or not element.get_attribute('x-account-name'):
        raise RuntimeError('Не удалось определить текущую учётную запись')
    return element.get_attribute('x-account-name'), element.text.strip()


def execution_candidates(rows, requested):
    """ASUD groups a document header and its tasks as sibling table rows."""
    current = None
    result = []
    for element, kind, is_document in rows:
        if is_document:
            numbers = FULL_NUMBER.findall(kind)
            current = numbers[0] if len(numbers) == 1 else None
        elif current and kind.strip() == 'Исполнение' and number_matches(current, requested):
            result.append((element, current))
    return result


def find_execution_task(driver, number):
    rows = []
    for row in visible(driver, 'tr.obj-list-rec.obj-list-task'):
        columns = row.find_elements(By.XPATH, './td[@cellindex="3"]')
        if not columns:
            continue
        is_document = 'containerGridTaskRow' in (row.get_attribute('class') or '')
        rows.append((row, columns[0].text, is_document))
    return unique(execution_candidates(rows, number), f'поручение {number}')


def card_number(driver):
    numbers = [element.text.strip().upper()
               for element in visible(driver, '[class*="headerInfoContainer"] b')
               if FULL_NUMBER.fullmatch(element.text.strip())]
    return unique(numbers, 'номер в карточке')


def open_execution(driver, number, timeout=25):
    if not click_sidebar_section(driver, 'Исполнение'):
        raise RuntimeError('Не удалось открыть полученные задания «Исполнение»')
    # Clear the date interval seen in snapshot 1 so old tasks can be found too.
    for field in visible(driver, '#FCPC_Дата_получения-input'):
        driver.execute_script("""
            arguments[0].value = '';
            for (const type of ['input', 'keyup', 'change'])
                arguments[0].dispatchEvent(new Event(type, {bubbles:true}));
        """, field)
    if not filter_by_number(driver, number):
        raise RuntimeError('Недоступен фильтр номера')
    time.sleep(1.5)  # Existing GXT filter debounce.
    task, resolved = report_wait(driver, timeout).until(
        lambda d: find_execution_task(d, number))
    # Do not use open_doc_card: its fallback replaces the selected task with
    # the first grid row, which can be the document header or a notification.
    ActionChains(driver).double_click(task).perform()
    report_wait(driver, timeout).until(lambda d: card_number(d) == resolved)
    header = unique(visible(driver, '[data-marker="ScreenHeader1"]'), 'шапка задания')
    if header is None or not header.text.strip().startswith('Исполнение:'):
        raise RuntimeError('Открыта карточка без задания на исполнение')
    click_lifecycle(driver)
    return resolved


def click_lifecycle(driver):
    tab = unique(visible(driver, '[data-marker="tab-lifecycle"]'), 'Жизненный цикл')
    if tab is None or not click(driver, tab, 'Жизненный цикл'):
        raise RuntimeError('Не удалось открыть жизненный цикл')


def own_lifecycle_row(driver, account):
    rows = [row for row in visible(driver, 'tr[class*="lfce-stage-grid-row_"]')
            if f'lfce-stage-grid-row_{account}' in (row.get_attribute('class') or '').split()]
    return unique(rows, 'строка исполнителя в жизненном цикле')


def report_dialog(driver):
    dialogs = [dialog for dialog in visible(driver, DIALOG_SELECTOR)
               if dialog.find_elements(By.CSS_SELECTOR, UPLOAD_SELECTOR)]
    return unique(dialogs, 'окно отчёта')


def enabled(driver, element):
    return element.is_enabled() and driver.execute_script("""
        for (let el = arguments[0]; el; el = el.parentElement) {
            if (el.getAttribute('aria-disabled') === 'true' ||
                el.getAttribute('data-disabled') === '1' ||
                /disabled/i.test(el.className || '') ||
                getComputedStyle(el).pointerEvents === 'none') return false;
        }
        return true;
    """, element)


def attachment_names(root):
    """Names acknowledged in the attachment list, never input.files."""
    names = set()
    for listing in visible(root, '[data-marker="combobox-list-view"]'):
        for element in [listing, *listing.find_elements(By.CSS_SELECTOR, '*')]:
            if not element.is_displayed():
                continue
            names.update(line.strip() for line in element.text.splitlines() if line.strip())
            for attribute in ('title', 'download'):
                value = element.get_attribute(attribute)
                if value:
                    names.add(value.strip())
    return names


def row_values(row):
    text_values = {line.strip() for line in row.text.splitlines() if line.strip()}
    for element in row.find_elements(By.CSS_SELECTOR, '[title], [download]'):
        if element.is_displayed():
            text_values.update(filter(None, (element.get_attribute('title'),
                                            element.get_attribute('download'))))
    return text_values


def report_visible(driver, account, text, filenames):
    row = own_lifecycle_row(driver, account)
    if row is None:
        return False
    # Match only the executor's row, including multiline report text.
    return (' '.join(text.split()) in ' '.join(row.text.split()) and
            set(filenames).issubset(row_values(row)))


def dialog_control(driver, selector, description):
    dialog = report_dialog(driver)
    return unique(visible(dialog, selector), description) if dialog is not None else None


def fill_report_text(driver, text, timeout):
    def enter_and_verify(d):
        textarea = dialog_control(d, 'textarea', 'текст отчёта')
        if textarea is None:
            return False
        if textarea.get_attribute('value') != text:
            # Retrying text entry is safe: replace the value, never append.
            textarea.clear()
            textarea.send_keys(text)
        textarea.send_keys(Keys.TAB)
        # TAB commits GXT's value and can replace the textarea or its dialog.
        # Every subsequent read must find the live control again.
        current = dialog_control(d, 'textarea', 'текст отчёта')
        return current is not None and current.get_attribute('value') == text

    report_wait(driver, timeout).until(enter_and_verify,
        message='Не удалось заполнить и подтвердить текст отчёта')
    log.info('Текст отчёта заполнен и проверен после обновления формы')


def uploaded_files_visible(driver, filenames):
    dialog = report_dialog(driver)
    return dialog is not None and filenames.issubset(attachment_names(dialog))


def ready_save_button(driver, account, number, text, filenames):
    dialog = report_dialog(driver)
    if dialog is None:
        return None
    textarea = unique(visible(dialog, 'textarea'), 'текст отчёта')
    if textarea is None or textarea.get_attribute('value') != text:
        return None
    if not filenames.issubset(attachment_names(dialog)):
        return None
    if any(el.text.strip() for el in visible(dialog, '[id="error-message-container"]')):
        raise RuntimeError('АСУД показывает ошибку заполнения отчёта')
    if card_number(driver) != number or own_lifecycle_row(driver, account) is None:
        raise RuntimeError('Перед сохранением изменились карточка или исполнитель')
    button = unique(visible(dialog, '[data-marker="dialog-ok-button"]'), 'Сохранить отчет')
    return button if button is not None and enabled(driver, button) else None


def add_reply_report(driver, account, number, files, text, before_save, timeout=60):
    if not text.strip() or len(text.encode('utf-16-le')) // 2 > 2000:
        raise ValueError('Текст отчёта должен содержать от 1 до 2000 символов')
    if not files:
        raise ValueError('Не заданы файлы ответа')
    if card_number(driver) != number:
        raise RuntimeError('Номер открытой карточки изменился')
    row = report_wait(driver, timeout).until(lambda d: own_lifecycle_row(d, account))
    if 'Принято к исполнению' not in row.text:
        raise RuntimeError('Поручение текущего исполнителя не принято к исполнению')
    if {path.name for path in files}.intersection(row_values(row)):
        raise RuntimeError('В строке исполнителя уже есть одноимённое вложение; проверьте отчёт')
    menu_button = unique(visible(row, '[data-name="MSFO_CNTX_MENU_IMG"]'), 'меню исполнителя')
    if menu_button is None or not click(driver, menu_button, 'меню исполнителя'):
        raise RuntimeError('Меню исполнителя недоступно')
    add = report_wait(driver, timeout).until(lambda d: unique(
        visible(d, '[data-marker="cntx-menu-list"] #cntx-menu-item-add_resolution_report'),
        'Добавить отчет'))
    if not enabled(driver, add):
        raise RuntimeError('Добавление отчёта недоступно; проверьте существующий отчёт')
    click(driver, add, 'Добавить отчет')
    fill_report_text(driver, text, timeout)
    # File inputs may be hidden by GXT. They must be reacquired from the live
    # dialog but do not need to pass the visibility test for send_keys.
    def upload_input(d):
        dialog = report_dialog(d)
        return dialog.find_element(By.CSS_SELECTOR, UPLOAD_SELECTOR) if dialog is not None else None

    upload = report_wait(driver, timeout).until(upload_input)
    log.info('Загружаю файлы ответа: %s', ', '.join(path.name for path in files))
    upload.send_keys('\n'.join(str(path.resolve(strict=True)) for path in files))
    filenames = {path.name for path in files}
    # send_keys dispatches change once. A second synthetic change can upload twice.
    report_wait(driver, timeout).until(
        lambda d: uploaded_files_visible(d, filenames),
        message='Файлы не появились в списке вложений отчёта')
    log.info('Все файлы появились в списке вложений')
    save = report_wait(driver, timeout).until(
        lambda d: ready_save_button(d, account, number, text, filenames),
        message='Форма отчёта не готова к сохранению')
    before_save()  # Durable pending journal before the first save attempt.
    log.info('Сохраняю отчёт по документу %s', number)
    save.click()  # Exactly one attempt; timeout/transport error is ambiguous.
    report_wait(driver, timeout).until(lambda d: report_dialog(d) is None)
    report_wait(driver, timeout).until(
        lambda d: report_visible(d, account, text, filenames))
    # Confirm persistence by reopening the task from the server-backed list.
    close_card(driver)
    reopened = open_execution(driver, number)
    if reopened != number:
        raise RuntimeError('После открытия изменился номер документа')
    report_wait(driver, timeout).until(
        lambda d: report_visible(d, account, text, filenames))


def close_card(driver):
    button = unique(visible(driver, '#header-close-btn'), 'закрыть карточку')
    if button is None or not click(driver, button, 'закрыть карточку'):
        raise RuntimeError('Не удалось закрыть карточку')
    report_wait(driver, 20).until(lambda d: not visible(d, '#header-close-btn'))
