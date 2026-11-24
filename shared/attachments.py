"""
attachments.py — Поиск и прикрепление .msg файлов.

Ищет файл по Link из Excel в outlook_dir (рекурсивно по подпапкам),
прикрепляет через CDP file chooser intercept + send_keys на скрытый input
GXT FileUploadButton. Работает в headless-режиме, не требует Windows GUI.
"""

import os
import re
import time
import logging
import mimetypes
import tempfile
import zipfile
from email.message import EmailMessage
from email.policy import SMTP
from io import BytesIO
from datetime import datetime, date

import extract_msg
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from shared.ui import click, wait_modal_closed, close_open_modals

log = logging.getLogger("asud.attach")


def _safe_msg_value(message, field, default=""):
    """Читает поле MSG, не ломая упаковку на повреждённой метаинформации."""
    if message is None:
        return default
    try:
        value = getattr(message, field, default)
    except Exception:
        return default
    if value is None:
        return default
    return str(value)


def _safe_archive_stem(file_path):
    """Делает короткое Windows-совместимое имя временного ZIP."""
    stem = os.path.splitext(os.path.basename(file_path))[0]
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', stem).strip(' .')
    return (stem or "Письмо")[:120]


def _docx_text(value):
    """Очищает текст от недопустимых в WordprocessingML управляющих символов."""
    text = str(value or "")
    text = text.encode("utf-8", errors="replace").decode("utf-8")
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)


def _set_cell_margins(cell, *, top=90, start=110, bottom=90, end=110):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (
            ("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_table_borders(table, color="D9D9D9", size="6"):
    table_pr = table._tbl.tblPr
    borders = table_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = borders.find(qn(f"w:{edge}"))
        if border is None:
            border = OxmlElement(f"w:{edge}")
            borders.append(border)
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)


def _set_style_font(style, name):
    """Закрепляет шрифт стиля для кириллицы во всех Word-рендерерах."""
    style.font.name = name
    r_pr = style._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.set(qn(f"w:{key}"), name)


def _build_letter_docx(metadata, body):
    """Создаёт компактный DOCX с реквизитами и полным текстом письма."""
    document = Document()
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    normal = document.styles["Normal"]
    _set_style_font(normal, "Arial")
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.08

    title_style = document.styles["Title"]
    _set_style_font(title_style, "Arial")
    title_style.font.size = Pt(16)
    title_style.font.bold = True
    title_style.font.color.rgb = RGBColor(0, 0, 0)
    title_p_pr = title_style._element.get_or_add_pPr()
    title_border = title_p_pr.find(qn("w:pBdr"))
    if title_border is not None:
        title_p_pr.remove(title_border)

    heading_style = document.styles["Heading 1"]
    _set_style_font(heading_style, "Arial")
    heading_style.font.size = Pt(12)
    heading_style.font.bold = True
    heading_style.font.color.rgb = RGBColor(0, 0, 0)

    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title.paragraph_format.space_after = Pt(12)
    title.add_run("Обращение")

    rows = [(str(label), _docx_text(value)) for label, value in metadata]
    table = document.add_table(rows=len(rows), cols=2)
    table.autofit = False
    table.columns[0].width = Inches(1.0)
    table.columns[1].width = Inches(5.95)
    _set_table_borders(table)
    for index, (label, value) in enumerate(rows):
        label_cell, value_cell = table.rows[index].cells
        label_cell.width = Inches(1.0)
        value_cell.width = Inches(5.95)
        label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        value_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _set_cell_margins(label_cell)
        _set_cell_margins(value_cell)
        label_paragraph = label_cell.paragraphs[0]
        label_paragraph.paragraph_format.space_after = Pt(0)
        label_paragraph.add_run(label).bold = True
        value_paragraph = value_cell.paragraphs[0]
        value_paragraph.paragraph_format.space_after = Pt(0)
        value_paragraph.add_run(value)

    spacer = document.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)
    heading = document.add_paragraph(style="Heading 1")
    heading.paragraph_format.space_before = Pt(4)
    heading.paragraph_format.space_after = Pt(6)
    heading.add_run("Текст обращения")

    cleaned_body = _docx_text(body).replace("\r\n", "\n").replace("\r", "\n")
    if cleaned_body:
        for line in cleaned_body.split("\n"):
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.space_after = Pt(3 if line else 0)
            if line:
                paragraph.add_run(line)
    else:
        document.add_paragraph("Текст письма отсутствует")

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def build_msg_zip(msg_path, output_path, *, subject=None, body=None):
    """Упаковывает письмо в ZIP: ``Письмо.docx`` + все MSG-вложения.

    ZIP сначала записывается во временный ``.partial`` и только после
    успешного чтения всех вложений атомарно становится ``output_path``.
    Поэтому в АСУД никогда не уйдёт недособранный архив.
    """
    source = os.path.abspath(os.fspath(msg_path))
    target = os.path.abspath(os.fspath(output_path))
    partial = f"{target}.partial"
    os.makedirs(os.path.dirname(target), exist_ok=True)

    message = None
    try:
        message = extract_msg.openMsg(source)
        effective_subject = (
            str(subject) if subject is not None
            else _safe_msg_value(message, "subject")
        )
        effective_body = (
            str(body) if body is not None
            else _safe_msg_value(message, "body")
        )
        metadata = [
            ("Тема", effective_subject),
            ("От", _safe_msg_value(message, "sender")),
            ("Кому", _safe_msg_value(message, "to")),
            ("Копия", _safe_msg_value(message, "cc")),
            ("Дата", _safe_msg_value(message, "date")),
        ]
        letter_docx = _build_letter_docx(metadata, effective_body)

        if os.path.exists(partial):
            os.remove(partial)
        with zipfile.ZipFile(
                partial, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("Письмо.docx", letter_docx)
            for attachment in (getattr(message, "attachments", None) or []):
                # extract-msg сам очищает имя и устраняет конфликты имён.
                # Вложенное письмо сохраняется как .msg, а не теряется.
                attachment.save(
                    zip=archive,
                    customPath="Вложения",
                    extractEmbedded=True,
                )

        os.replace(partial, target)
        return target
    except Exception:
        try:
            if os.path.exists(partial):
                os.remove(partial)
        except OSError:
            pass
        raise
    finally:
        if message is not None:
            try:
                message.close()
            except Exception:
                pass


def _save_msg_attachments(message, target_dir):
    """Извлекает все MSG-вложения во временную папку и возвращает пути.

    Используем штатный ``extract-msg.save``: так обычные файлы и вложенные
    письма сохраняются одинаково надёжно, а имена очищаются самой библиотекой.
    """
    root = os.path.abspath(os.fspath(target_dir))
    os.makedirs(root, exist_ok=True)
    saved = []
    for attachment in (getattr(message, "attachments", None) or []):
        result = attachment.save(
            customPath=root,
            extractEmbedded=True,
        )
        path = result[1] if isinstance(result, tuple) and len(result) > 1 else None
        if not path:
            continue
        full_path = os.path.abspath(os.fspath(path))
        try:
            inside_root = os.path.commonpath((root, full_path)) == root
        except ValueError:
            inside_root = False
        if not inside_root or not os.path.isfile(full_path):
            raise ValueError(
                f"extract-msg вернул небезопасный путь вложения: {full_path}"
            )
        saved.append(full_path)
    return saved


def build_msg_eml(msg_path, output_path, *, subject=None, body=None):
    """Создаёт безопасный для АСУД EML-контейнер без внешнего ZIP.

    Внутри EML находятся ``Обращение.docx`` и все исходные вложения письма.
    Это сохраняет комплектность обращения, но не заставляет АСУД обрабатывать
    ZIP, на скачивании которого его веб-интерфейс может зависнуть.
    """
    source = os.path.abspath(os.fspath(msg_path))
    target = os.path.abspath(os.fspath(output_path))
    partial = f"{target}.partial"
    os.makedirs(os.path.dirname(target), exist_ok=True)

    message = None
    try:
        message = extract_msg.openMsg(source)
        effective_subject = (
            str(subject) if subject is not None
            else _safe_msg_value(message, "subject")
        )
        effective_body = (
            str(body) if body is not None
            else _safe_msg_value(message, "body")
        )
        metadata = [
            ("Тема", effective_subject),
            ("От", _safe_msg_value(message, "sender")),
            ("Кому", _safe_msg_value(message, "to")),
            ("Копия", _safe_msg_value(message, "cc")),
            ("Дата", _safe_msg_value(message, "date")),
        ]
        letter_docx = _build_letter_docx(metadata, effective_body)

        bundle = EmailMessage(policy=SMTP)
        safe_subject = re.sub(r"[\r\n]+", " ", effective_subject).strip()
        bundle["Subject"] = safe_subject or "Обращение"
        bundle["X-ASUD-Container"] = "letter-docx-with-original-attachments"
        bundle.set_content(
            effective_body or "Текст письма отсутствует",
            subtype="plain",
            charset="utf-8",
        )
        bundle.add_attachment(
            letter_docx,
            maintype="application",
            subtype=(
                "vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            filename="Обращение.docx",
        )

        with tempfile.TemporaryDirectory(prefix="asud_msg_parts_") as parts_dir:
            for attachment_path in _save_msg_attachments(message, parts_dir):
                mime_type, _encoding = mimetypes.guess_type(attachment_path)
                if mime_type and "/" in mime_type:
                    maintype, subtype = mime_type.split("/", 1)
                else:
                    maintype, subtype = "application", "octet-stream"
                with open(attachment_path, "rb") as attachment_file:
                    bundle.add_attachment(
                        attachment_file.read(),
                        maintype=maintype,
                        subtype=subtype,
                        filename=os.path.basename(attachment_path),
                    )

        if os.path.exists(partial):
            os.remove(partial)
        with open(partial, "wb") as eml_file:
            eml_file.write(bundle.as_bytes(policy=SMTP))
        os.replace(partial, target)
        return target
    except Exception:
        try:
            if os.path.exists(partial):
                os.remove(partial)
        except OSError:
            pass
        raise
    finally:
        if message is not None:
            try:
                message.close()
            except Exception:
                pass


def _link_to_variants(link):
    """Генерирует все возможные имена файлов для link.
    Форматы: DD.MM.YYYY / YYYY-MM-DD, с/без ведущего нуля, суффиксы _1-_9."""
    variants = set()

    if isinstance(link, (datetime, date)):
        try:
            variants.add(link.strftime("%d.%m.%Y %H-%M-%S"))
            variants.add(link.strftime("%Y-%m-%d %H-%M-%S"))
            with_lead = link.strftime("%d.%m.%Y %H-%M-%S")
            m = re.match(r'^(\d{2}\.\d{2}\.\d{4})\s+0(\d)-(\d{2})-(\d{2})$', with_lead)
            if m:
                variants.add(f"{m.group(1)} {m.group(2)}-{m.group(3)}-{m.group(4)}")
        except Exception:
            pass
    else:
        s = str(link).strip()
        if s:
            variants.add(s)
            # DD.MM.YYYY → ISO (+ с/без ведущего нуля в часе)
            m = re.match(r'^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{1,2})-(\d{2})-(\d{2})$', s)
            if m:
                dd, mm, yyyy, h, mn, sec = m.groups()
                h_lead = f"{int(h):02d}"
                h_no = str(int(h))
                variants.add(f"{dd}.{mm}.{yyyy} {h_lead}-{mn}-{sec}")
                variants.add(f"{dd}.{mm}.{yyyy} {h_no}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_lead}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_no}-{mn}-{sec}")
            # ISO → DD.MM.YYYY
            m = re.match(r'^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2})-(\d{2})-(\d{2})$', s)
            if m:
                yyyy, mm, dd, h, mn, sec = m.groups()
                h_lead = f"{int(h):02d}"
                h_no = str(int(h))
                variants.add(f"{dd}.{mm}.{yyyy} {h_lead}-{mn}-{sec}")
                variants.add(f"{dd}.{mm}.{yyyy} {h_no}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_lead}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_no}-{mn}-{sec}")

    variants.discard('')
    return list(variants)


def find_msg_by_link(link, outlook_dir, fallback_path=None):
    """Ищет .msg файл в outlook_dir (рекурсивно по подпапкам) по link.
    Fallback → пустышка."""
    log.info(f"link={link!r} (тип: {type(link).__name__})")

    if link is None or (isinstance(link, str) and not link.strip()):
        log.warning("link пустой — пустышка")
        return fallback_path

    if not os.path.isdir(outlook_dir):
        log.warning(f"Папка {outlook_dir} не найдена — пустышка")
        return fallback_path

    variants = _link_to_variants(link)
    log.info(f"Варианты: {variants}")

    # Один проход по дереву — собираем все .msg
    all_msg = []  # [(full_path, filename_no_ext)]
    try:
        for root, _dirs, files in os.walk(outlook_dir):
            for f in files:
                if f.lower().endswith('.msg'):
                    name_no_ext = os.path.splitext(f)[0]
                    all_msg.append((os.path.join(root, f), name_no_ext))
    except Exception as e:
        log.error(f"Ошибка обхода {outlook_dir}: {e}")
        return fallback_path

    if not all_msg:
        log.warning(f"В {outlook_dir} нет .msg файлов")
        return fallback_path

    def _rel(full):
        try:
            return os.path.relpath(full, outlook_dir)
        except Exception:
            return os.path.basename(full)

    # Фаза 1: точное совпадение с вариантом
    variants_set = set(variants)
    for full, name in all_msg:
        if name in variants_set:
            log.info(f"Нашёл: {_rel(full)}")
            return full

    # Фаза 2: variant + _1.._9
    suffix_set = {f"{v}_{i}" for v in variants for i in range(1, 10)}
    for full, name in all_msg:
        if name in suffix_set:
            log.info(f"Нашёл (суффикс): {_rel(full)}")
            return full

    # Фаза 3: подстрока
    for full, name in all_msg:
        for v in variants:
            if v in name:
                log.info(f"Нашёл (подстрока): {_rel(full)}")
                return full

    log.warning("Файл не найден — пустышка")
    return fallback_path


def get_dummy_msg(base_dir):
    """Ищет .msg-пустышку рядом с exe."""
    msg_files = [f for f in os.listdir(base_dir) if f.lower().endswith('.msg')]
    if msg_files:
        return os.path.join(base_dir, msg_files[0])
    return None


def move_to_done(file_path, outlook_dir, done_dirname="Завершено", delete=None):
    """Завершает работу с обработанным .msg.

    delete — режим финализации:
      None (default) → читать из env ASUD_DELETE_AFTER_DONE ('1' = удалять)
      True  → удалить файл
      False → переместить в <outlook_dir>/Завершено/

    Если delete=True или env='1' — файл удаляется. Иначе переезжает в
    Завершено/ (как раньше). Папка создаётся если нет. Конфликт имён → _HHMMSS.
    """
    if not file_path or not os.path.isfile(file_path):
        return

    if delete is None:
        delete = os.environ.get('ASUD_DELETE_AFTER_DONE', '0') == '1'

    name = os.path.basename(file_path)

    if delete:
        try:
            os.remove(file_path)
            log.info(f"× удалён: {name}")
        except Exception as e:
            log.warning(f"Не удалось удалить {name}: {e}")
        return

    if not outlook_dir or not os.path.isdir(outlook_dir):
        log.warning(f"outlook_dir '{outlook_dir}' не существует — "
                    f"не перемещаю {name}")
        return

    done_dir = os.path.join(outlook_dir, done_dirname)
    try:
        os.makedirs(done_dir, exist_ok=True)
    except Exception as e:
        log.warning(f"Не удалось создать {done_dir}: {e}")
        return

    dest = os.path.join(done_dir, name)
    if os.path.exists(dest):
        base, ext = os.path.splitext(name)
        ts = datetime.now().strftime("%H%M%S")
        dest = os.path.join(done_dir, f"{base}_{ts}{ext}")

    try:
        import shutil
        shutil.move(file_path, dest)
        log.info(f"→ Завершено/{os.path.basename(dest)}")
    except Exception as e:
        log.warning(f"Не удалось переместить {name}: {e}")


def move_to_drafts(file_path, base_dir, drafts_dirname="Черновики"):
    """Перемещает .msg в <base_dir>/Черновики/ — для daemon-режима, чтобы
    не подхватывать снова в следующей итерации. Юзер при необходимости
    вернёт обратно после ручной правки.
    """
    if not file_path or not os.path.isfile(file_path):
        return
    if not base_dir or not os.path.isdir(base_dir):
        return
    drafts_dir = os.path.join(base_dir, drafts_dirname)
    try:
        os.makedirs(drafts_dir, exist_ok=True)
    except Exception as e:
        log.warning(f"Не удалось создать {drafts_dir}: {e}")
        return
    name = os.path.basename(file_path)
    dest = os.path.join(drafts_dir, name)
    if os.path.exists(dest):
        base, ext = os.path.splitext(name)
        ts = datetime.now().strftime("%H%M%S")
        dest = os.path.join(drafts_dir, f"{base}_{ts}{ext}")
    try:
        import shutil
        shutil.move(file_path, dest)
        log.info(f"→ Черновики/{os.path.basename(dest)}")
    except Exception as e:
        log.warning(f"Не удалось переместить {name} в Черновики: {e}")


def move_to_errors(file_path, base_dir, reason="", err_dirname="Ошибки"):
    """Перемещает .msg в <base_dir>/Ошибки/ и пишет рядом .txt с причиной.
    Папка создаётся если нет. Конфликт имён → суффикс _HHMMSS.
    Ничего не делает если file_path пустой/не существует.
    """
    if not file_path:
        return False
    if not os.path.isfile(file_path):
        # Источник уже исчез (например, его успел переместить другой процесс).
        # Для вызывающего это безопасный terminal-исход: повторять нечего.
        return not os.path.exists(file_path)
    if not base_dir or not os.path.isdir(base_dir):
        log.warning(f"base_dir '{base_dir}' не существует — "
                    f"не перемещаю {os.path.basename(file_path)}")
        return False

    err_dir = os.path.join(base_dir, err_dirname)
    try:
        os.makedirs(err_dir, exist_ok=True)
    except Exception as e:
        log.warning(f"Не удалось создать {err_dir}: {e}")
        return False

    name = os.path.basename(file_path)
    dest = os.path.join(err_dir, name)
    if os.path.exists(dest):
        base, ext = os.path.splitext(name)
        ts = datetime.now().strftime("%H%M%S")
        dest = os.path.join(err_dir, f"{base}_{ts}{ext}")

    try:
        import shutil
        shutil.move(file_path, dest)
        log.info(f"→ Ошибки/{os.path.basename(dest)}")
        if reason:
            sidecar = os.path.splitext(dest)[0] + ".txt"
            try:
                with open(sidecar, "w", encoding="utf-8") as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n{reason}\n")
            except Exception as e:
                log.debug(f"sidecar не записан: {e}")
        return not os.path.exists(file_path) and os.path.isfile(dest)
    except Exception as e:
        log.warning(f"Не удалось переместить {name} в Ошибки: {e}")
        return False


_DIAG_BUTTONS_JS = r"""
// Диагностика: вернуть видимые кандидаты на confirm-кнопку
const out = [];
for (const el of document.querySelectorAll(
        "[id*='SetContent'], [id*='Send'], [id*='Submit'], button, div, span")) {
    if (!el.offsetParent) continue;
    const t = (el.textContent || '').trim();
    if (t.length > 80) continue;
    if (!/(Присоединить|Подтвердить|Сохранить|Загрузить)/i.test(t)
        && !/(SetContent|Send|Submit|Save)/i.test(el.id || '')) continue;
    out.push((el.id || el.tagName) + '|' + t.slice(0, 60));
}
return out;
"""


_DIAG_FILE_INPUTS_JS = r"""
// Диагностика: переписать все <input type="file"> в DOM
// со всеми атрибутами и контекстом. Используется при таймауте
// _attach_via_input для понимания что в DOM пошло не так.
const out = [];
for (const inp of document.querySelectorAll("input[type='file']")) {
    const r = inp.getBoundingClientRect();
    const cs = window.getComputedStyle(inp);
    const parentTag = inp.parentElement ? inp.parentElement.tagName : null;
    const parentCls = inp.parentElement ? (inp.parentElement.className || '').toString().slice(0, 80) : null;
    out.push({
        id: inp.id || null,
        name: inp.name || null,
        accept: inp.getAttribute('accept'),
        multiple: inp.multiple,
        disabled: inp.disabled,
        offsetParent: !!inp.offsetParent,  // true если отображается
        rect: {w: r.width, h: r.height, x: r.left, y: r.top},
        display: cs.display,
        visibility: cs.visibility,
        opacity: cs.opacity,
        position: cs.position,
        parent_tag: parentTag,
        parent_class: parentCls,
        outerHTML_preview: inp.outerHTML.slice(0, 200),
    });
}
return {count: out.length, inputs: out};
"""


def _diag_file_inputs(driver, label):
    """Дампит в лог состояние всех <input type='file'> на странице."""
    try:
        diag = driver.execute_script(_DIAG_FILE_INPUTS_JS) or {}
        log.info(f"[diag {label}] input[type=file] count={diag.get('count', 0)}")
        for i, inp in enumerate(diag.get('inputs', []) or []):
            log.info(f"  [diag {label}] input[{i}]: id={inp.get('id')!r} name={inp.get('name')!r} "
                     f"disabled={inp.get('disabled')} offsetParent={inp.get('offsetParent')} "
                     f"display={inp.get('display')} visibility={inp.get('visibility')} "
                     f"opacity={inp.get('opacity')}")
            log.info(f"  [diag {label}]    rect={inp.get('rect')} parent={inp.get('parent_tag')}.{inp.get('parent_class')}")
            log.info(f"  [diag {label}]    html: {inp.get('outerHTML_preview')!r}")
    except Exception as e:
        log.warning(f"[diag {label}] dump упал: {e}")


def _attach_via_input(driver, file_path):
    """Стратегия через CDP file chooser intercept + send_keys на скрытый input.

    Найдено диагностикой стадии 0:
      • АСУД использует GXT FileUploadButton — скрытый <input type='file'> с
        name='SetContentDialog' создаётся через ~0.5-0.8s после клика
        'Присоединить содержимое'.
      • input невидим (display:none, rect 0×0), но functional.

    Шаги:
      1. Включаем Page.setInterceptFileChooserDialog → native picker не открывается
      2. Кликаем 'Присоединить содержимое' → input создаётся в DOM, picker
         перехвачен CDP'ом и не показан
      3. Ждём появления input[name='SetContentDialog'] (до 3s)
      4. Делаем input visible через JS (чтобы send_keys прошёл visibility-check)
      5. input.send_keys(file_path) — Selenium через WebDriver-protocol устанавливает
         значение файлового input'а без участия native picker
      6. Выключаем intercept
      7. Ждём confirm-кнопку, кликаем

    Возвращает True если успех, False если что-то упало (caller сделает retry).
    """
    intercept_enabled = False
    try:
        # 1. Включаем перехват native file picker
        try:
            driver.execute_cdp_cmd("Page.setInterceptFileChooserDialog",
                                    {"enabled": True})
            intercept_enabled = True
            log.debug("CDP file chooser intercept включён")
        except Exception as e:
            log.warning(f"CDP intercept не доступен: {e}")
            return False

        # 2. Клик по 'Присоединить содержимое' — picker не откроется (intercept)
        try:
            btn = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH,
                    "//div[contains(text(),'Присоединить содержимое')]")))
            click(driver, btn, "Присоединить содержимое (CDP)")
        except Exception as e:
            log.error(f"Кнопка 'Присоединить содержимое' не найдена: {e}")
            return False

        # 3. Ждём появления input[name='SetContentDialog']
        end = time.monotonic() + 3
        inp = None
        while time.monotonic() < end:
            try:
                inp = driver.find_element(By.CSS_SELECTOR,
                    "input[type='file'][name='SetContentDialog']")
                break
            except Exception:
                pass
            time.sleep(0.1)
        if not inp:
            log.warning("input[name='SetContentDialog'] не появился за 3s")
            return False
        log.debug(f"input найден: id={inp.get_attribute('id')}")

        # 4. Делаем visible — иначе send_keys жалуется на visibility
        driver.execute_script("""
            var el = arguments[0];
            el.style.display='block';
            el.style.visibility='visible';
            el.style.opacity='1';
            el.style.position='absolute';
            el.style.left='0px';
            el.style.top='0px';
            el.style.width='1px';
            el.style.height='1px';
            el.removeAttribute('hidden');
            el.disabled = false;
        """, inp)

        # 5. send_keys устанавливает значение файлового input напрямую через
        # WebDriver-protocol (без участия native picker)
        inp.send_keys(file_path)

        # 5а. ВАЖНО: дождаться пока input.files реально содержит файл.
        # Без этого confirm-клик случается до того как upload дошёл, модалка
        # остаётся открытой, файл по факту не прикреплён.
        # 25с (раньше было 10) — в новой версии АСУД иногда долго отрисовывается.
        try:
            WebDriverWait(driver, 25, poll_frequency=0.1).until(
                lambda d: d.execute_script(
                    "return arguments[0].files && arguments[0].files.length > 0;",
                    inp))
            log.debug("input.files подтверждён")
        except Exception:
            log.warning("input.files пустой за 25s — upload скорее всего не дошёл")
            return False

        # Триггерим change на случай если автоматический dispatch не сработал
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", inp)

        # 5б. Даём АСУД 4s на серверный upload. На некоторых документах
        # 1.5s было мало — upload не успевал, confirm кликался впустую,
        # ждали 30s+ retry-cycle. 4s покрывает ~95% случаев одной попыткой.
        time.sleep(4)

        log.info(f"Файл отправлен через CDP-intercept + send_keys")
        return True

    except Exception as e:
        log.warning(f"CDP-стратегия упала: {e}")
        return False
    # NB: intercept НЕ выключаем здесь — caller (attach_content) сам отключит
    # после завершения всего flow включая confirm-click. Иначе на retry
    # _wait_confirm_and_click может случайно нажать главную кнопку
    # 'Присоединить содержимое' (имеет тот же текст что confirm), и АСУД
    # откроет native picker потому что intercept выключен.


def _wait_confirm_and_click(driver, timeout=30):
    """Поллинг + клик confirm-кнопки 'Присоединить' с retry если модалка не закрылась.

    На большинстве документов АСУД-server upload занимает >5s, и confirm-кнопка
    в это время хоть и видима, но имеет pointer-events:none или data-disabled='1'
    (как register-кнопка после Save). Кликаем только когда:
      • кнопка реально clickable (pointer-events != 'none', data-disabled != '1')
    Логика:
      1. Найти confirm-кнопку
      2. Подождать пока она станет clickable (poll 200ms, до 15s)
      3. Кликнуть
      4. Подождать 3s что модалка закрылась
      5. Если ещё видна — снова ждать clickable + клик. До 5 попыток.
    """
    end = time.time() + timeout

    def _find_btn():
        # Только confirm-кнопки С НЕПУСТЫМ id-маркером SetContent/Send/Submit.
        # Без id или с пустым id — это может быть главная кнопка 'Присоединить
        # содержимое' на карточке документа, кликать её — открывать picker (если
        # бы intercept был выкл) или по крайней мере бессмысленно.
        try:
            btns = driver.find_elements(By.CSS_SELECTOR,
                "[id*='SetContent'], [id*='Send'], [id*='Submit']")
            for b in btns:
                if b.is_displayed() and b.get_attribute('id'):
                    return b
        except Exception:
            pass
        # Fallback: text-search ВНУТРИ открытой модалки (не на карточке)
        try:
            modals = driver.find_elements(By.CSS_SELECTOR,
                "div[class*='ModalPanel'][class*='panel']")
            for modal in modals:
                if not modal.is_displayed():
                    continue
                btns = modal.find_elements(By.XPATH,
                    ".//button[contains(text(),'Присоединить')] | .//div[contains(text(),'Присоединить')]")
                for b in btns:
                    if b.is_displayed():
                        return b
        except Exception:
            pass
        return None

    def _is_btn_clickable(btn):
        """clickable = visible + pointer-events != 'none' + data-disabled != '1' + не в class*='disabled'"""
        try:
            if not btn.is_displayed():
                return False
            data_dis = btn.get_attribute('data-disabled')
            aria_dis = btn.get_attribute('aria-disabled')
            cls = (btn.get_attribute('class') or '').lower()
            if data_dis == '1' or aria_dis == 'true' or 'disabled' in cls:
                return False
            pe = driver.execute_script(
                "return window.getComputedStyle(arguments[0]).pointerEvents;", btn)
            if pe == 'none':
                return False
            return True
        except Exception:
            return False

    def _wait_clickable(btn, max_wait):
        """Ждём пока кнопка станет clickable. Возвращает True/False."""
        deadline = time.time() + max_wait
        last_state = None
        while time.time() < deadline:
            if _is_btn_clickable(btn):
                return True
            try:
                state = (
                    f"d={btn.get_attribute('data-disabled')} "
                    f"a={btn.get_attribute('aria-disabled')} "
                    f"pe={driver.execute_script('return window.getComputedStyle(arguments[0]).pointerEvents;', btn)}"
                )
                if state != last_state:
                    log.debug(f"  confirm не clickable: {state}")
                    last_state = state
            except Exception:
                pass
            time.sleep(0.2)
        return False

    def _modal_closed():
        return _find_btn() is None

    attempts = 0
    while time.time() < end and attempts < 5:
        btn = _find_btn()
        if not btn:
            log.info("Файл присоединён!")
            return True

        # Дожидаемся пока кнопка реально станет clickable (upload завершился)
        log.debug(f"Жду пока confirm станет clickable (try #{attempts + 1})")
        if not _wait_clickable(btn, max_wait=15):
            log.warning(f"Confirm-кнопка не стала clickable за 15s (try #{attempts + 1})")
            attempts += 1
            time.sleep(1)
            continue

        attempts += 1
        click(driver, btn, f"confirm by-id {btn.get_attribute('id')} (try #{attempts})")

        # Ждём 3s что модалка закрылась
        sub_end = time.time() + 3
        while time.time() < sub_end:
            if _modal_closed():
                log.info("Файл присоединён!")
                return True
            time.sleep(0.2)

        log.debug(f"Модалка ещё открыта после try #{attempts}, жду 2s и retry")
        time.sleep(2)

    log.warning(f"Confirm не сработал за {attempts} попыток — диагностика:")
    try:
        candidates = driver.execute_script(_DIAG_BUTTONS_JS) or []
        log.warning(f"  Видимых кандидатов: {len(candidates)}")
        for c in candidates[:10]:
            log.warning(f"    - {c}")
    except Exception:
        pass
    return False


def _attach_content_file(driver, file_path, max_attempts=3):
    """Прикрепляет файл к карточке документа через CDP intercept + send_keys.

    Возвращает True если файл успешно приложен, False если все retry'и
    провалились. Caller'ы могут использовать результат для лога в
    высокоуровневом логгере 'asud' (виден в консоли).

    До max_attempts попыток. Между попытками закрываем висящие модалки
    и даём АСУД пару секунд отдохнуть.

    CDP intercept ВКЛЮЧАЕТСЯ внутри _attach_via_input и не выключается там —
    выключается ТОЛЬКО в finally этой функции, после всех retry'ев и confirm-click'ов.
    Иначе случайный клик на похожую кнопку 'Присоединить содержимое' между
    попытками открыл бы native picker.
    """
    log.info(f"Прикрепление: {os.path.basename(file_path)}")

    try:
        for attempt in range(1, max_attempts + 1):
            if attempt > 1:
                # Проверка: может быть первый attempt по факту УЖЕ прикрепил файл,
                # просто input.files не отрисовался вовремя? Если confirm-кнопка
                # SetContent/Send уже видна — значит АСУД принял файл, повторный
                # _attach_via_input приведёт к ДУБЛИРОВАНИЮ вложения.
                try:
                    confirm_btns = driver.find_elements(By.CSS_SELECTOR,
                        "[id*='SetContent'], [id*='Send'], [id*='Submit']")
                    has_visible_confirm = any(
                        b.is_displayed() and b.get_attribute('id')
                        for b in confirm_btns)
                    if has_visible_confirm:
                        log.info(f"Перед попыткой {attempt}: confirm-кнопка уже видна — "
                                 f"файл по факту приложен. Пропускаю _attach_via_input, "
                                 f"иду на confirm.")
                        if _wait_confirm_and_click(driver, timeout=20):
                            return True
                        # confirm не сработал — закрываем и пробуем сначала
                        log.warning("confirm не сработал даже когда кнопка была видна — закрываю всё")
                except Exception:
                    pass
                log.warning(f"Прикрепление: попытка {attempt}/{max_attempts}")
                try:
                    close_open_modals(driver)
                except Exception:
                    pass
                time.sleep(2)

            if _attach_via_input(driver, file_path):
                if _wait_confirm_and_click(driver, timeout=20):
                    return True  # успех
                log.warning(f"Попытка {attempt}: файл загружен но confirm не сработал")
            else:
                log.warning(f"Попытка {attempt}: _attach_via_input вернул False")

        log.error(f"Прикрепление НЕ удалось за {max_attempts} попыток — файл не приложен")
        return False
    finally:
        # Выключаем CDP intercept в самом конце — гарантия что между retry'ями
        # native picker не откроется случайно от чужих clikks.
        try:
            driver.execute_cdp_cmd("Page.setInterceptFileChooserDialog",
                                    {"enabled": False})
        except Exception:
            pass


def attach_content(driver, file_path, max_attempts=3, *,
                   msg_attachment_mode="original", email_subject=None,
                   email_body=None):
    """Прикрепляет файл; для MSG собирает временный EML или legacy-ZIP.

    ``eml`` создаёт один неархивный EML: внутри ``Обращение.docx`` и исходные
    вложения. Старое значение ``zip`` автоматически мигрирует на ``eml``, чтобы
    уже существующие settings.json перестали отправлять проблемные ZIP.
    ``zip-legacy`` оставлен только для ручной диагностики.
    """
    mode = str(msg_attachment_mode or "original").strip().casefold()
    if mode == "zip":
        log.warning(
            "Режим zip отключён: АСУД зависает при скачивании архивов. "
            "Автоматически использую безопасный EML-контейнер."
        )
        mode = "eml"
    if mode not in {"original", "eml", "zip-legacy"}:
        log.warning(
            f"Неизвестный msg_attachment_mode={msg_attachment_mode!r}; "
            "прикрепляю исходный файл"
        )
        mode = "original"

    source = os.path.abspath(os.fspath(file_path))
    if mode == "original" or not source.lower().endswith(".msg"):
        return _attach_content_file(driver, source, max_attempts=max_attempts)

    try:
        with tempfile.TemporaryDirectory(prefix="asud_msg_bundle_") as temp_dir:
            suffix = ".zip" if mode == "zip-legacy" else ".eml"
            bundle_path = os.path.join(
                temp_dir, f"{_safe_archive_stem(source)}{suffix}"
            )
            builder = build_msg_zip if mode == "zip-legacy" else build_msg_eml
            builder(
                source,
                bundle_path,
                subject=email_subject,
                body=email_body,
            )
            log.info(
                f"MSG преобразован: {os.path.basename(source)} → "
                f"{os.path.basename(bundle_path)}"
            )
            return _attach_content_file(
                driver, bundle_path, max_attempts=max_attempts
            )
    except Exception as exc:
        log.exception(
            f"Не удалось собрать вложение из {os.path.basename(source)}: {exc}. "
            "Исходный MSG сохранён, вложение не отправлено."
        )
        return False
