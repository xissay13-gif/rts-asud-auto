"""
flows/sbis.py — регистрация в АСУД файлов, выгруженных из СБИС.

Каждый найденный файл становится отдельным входящим документом:
  • краткое содержание = имя файла (без расширения по умолчанию);
  • ФЛ = полное ФИО; ЮЛ = полное название / «-» / «-»;
  • номер у корреспондента = «б/н (N)»;
  • исходный файл обязательно прикрепляется к карточке;
  • далее используется обычный mix-flow: регистрация + «На резолюцию».

Новый sbis_downloader сохраняет PDF в формате
«[ФЛ][<полное ФИО>] <содержание>.pdf» или
«[ЮЛ][<полное название>] <содержание>.pdf». Поэтому регистратор получает
данные из имени, не открывая PDF. Старые имена поддерживаются для совместимости.

Рядом с исходными файлами хранится .asud_sbis_state.json. Он фиксирует
стабильный номер «б/н (N)» и не даёт повторно зарегистрировать уже
обработанный файл после перезапуска.
"""

import json
import logging
import os
import re
import sys
import time
from datetime import timedelta
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService

from flows import mix as mix_flow
from shared import config as cfg
from shared.ui import set_driver_timeout, wait_asud_loaded


log = logging.getLogger("asud")

STATE_NAME = ".asud_sbis_state.json"
DOWNLOAD_MANIFEST_NAME = ".manifest.json"
DONE_STATUSES = {"OK", "DUPLICATE"}
DEFAULT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".rtf", ".txt",
    ".xls", ".xlsx", ".xml", ".zip",
    ".jpg", ".jpeg", ".png", ".tif", ".tiff",
}
SKIP_DIRS = {
    "logs", "registered", "завершено", "ошибки", "черновики",
    "__pycache__", "build", "dist",
}

_DATE_PREFIX_RE = re.compile(r"^\s*\d{2}\.\d{2}\.\d{4}\s+")
_SURNAME_TOKEN_RE = re.compile(
    r"[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9.\-]*")
_PROTOCOL_V2_RE = re.compile(
    r"^\[(?P<kind>ФЛ|ЮЛ)\]\[(?P<correspondent>[^\]]+)\]\s*(?P<content>.+)$",
    re.IGNORECASE,
)
_PROTOCOL_V1_RE = re.compile(
    r"^\[SBIS-(?P<sbis_id>[^\]]+)\]\[(?P<surname>[^\]]+)\]\s*(?P<content>.+)$",
    re.IGNORECASE,
)


def _options(settings):
    """Возвращает настройки режима с безопасными дефолтами."""
    result = {
        "input_dir": "",
        "surname_override": "",
        "doc_type_index": 8,
        "person_doc_type_index": 8,
        "legal_doc_type_index": 5,
        "include_extension_in_content": False,
        "extensions": sorted(DEFAULT_EXTENSIONS),
    }
    user = settings.get("sbis") or {}
    if isinstance(user, dict):
        result.update(user)
    return result


def _normalise_extensions(raw):
    if isinstance(raw, str):
        raw = raw.split(",")
    result = set()
    for value in raw or []:
        ext = str(value).strip().lower()
        if not ext:
            continue
        result.add(ext if ext.startswith(".") else "." + ext)
    return result or set(DEFAULT_EXTENSIONS)


def scan_files(root, extensions=None):
    """Рекурсивно находит поддерживаемые файлы в стабильном порядке."""
    root = Path(root).resolve()
    extensions = _normalise_extensions(extensions or DEFAULT_EXTENSIONS)
    found = []
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        rel = path.relative_to(root)
        if any(part.casefold() in SKIP_DIRS for part in rel.parts[:-1]):
            continue
        name_lower = path.name.casefold()
        if (name_lower == STATE_NAME.casefold()
                or name_lower.startswith("~$")
                or name_lower.endswith("_резолюции.xlsx")):
            continue
        if path.suffix.casefold() not in extensions:
            continue
        try:
            if path.stat().st_size <= 0:
                log.warning(f"Пустой файл пропущен: {rel}")
                continue
        except OSError:
            continue
        found.append(path)
    return sorted(found, key=lambda p: str(p.relative_to(root)).casefold())


def _first_surname_token(value):
    value = str(value or "").strip().strip("_.,;:()[]{}\"'")
    match = _SURNAME_TOKEN_RE.search(value)
    return match.group(0).strip(".") if match else ""


def parse_protocol_filename(path):
    """Разбирает служебное имя PDF, созданное sbis_downloader."""
    stem = Path(path).stem.strip()
    match = _PROTOCOL_V2_RE.match(stem)
    if match:
        raw = {key: value.strip() for key, value in match.groupdict().items()}
        kind = "person" if raw["kind"].casefold() == "фл" else "legal"
        return {
            "protocol": "v2",
            "sbis_id": "",
            "correspondent_type": kind,
            "correspondent": raw["correspondent"],
            "surname": _first_surname_token(raw["correspondent"]),
            "content": raw["content"],
        }

    match = _PROTOCOL_V1_RE.match(stem)
    if match:
        raw = {key: value.strip() for key, value in match.groupdict().items()}
        if all(raw.values()):
            return {
                "protocol": "v1",
                "sbis_id": raw["sbis_id"],
                "correspondent_type": "legacy",
                "correspondent": raw["surname"],
                "surname": raw["surname"],
                "content": raw["content"],
            }
    return None


def derive_surname(path, root, override=""):
    """Извлекает фамилию из override, имени файла или старой папки СБИС."""
    if override:
        return _first_surname_token(override)

    path = Path(path)
    root = Path(root).resolve()
    protocol = parse_protocol_filename(path)
    if protocol:
        return _first_surname_token(protocol["correspondent"])
    candidate = path.stem

    # Старая раскладка: <тип>/<месяц>/<DD.MM.YYYY контрагент название>/<файл>.
    parent = path.parent
    while parent != root and root in parent.parents:
        if _DATE_PREFIX_RE.match(parent.name):
            candidate = _DATE_PREFIX_RE.sub("", parent.name)
            break
        parent = parent.parent

    candidate = _DATE_PREFIX_RE.sub("", candidate)
    # В новой files-раскладке: «контрагент - имя вложения».
    if " - " in candidate:
        candidate = candidate.split(" - ", 1)[0]
    return _first_surname_token(candidate)


def correspondent_from_file(path, root, override=""):
    """Возвращает (тип, полное имя) для поиска/создания в АСУД."""
    protocol = parse_protocol_filename(path)
    if protocol and protocol["protocol"] == "v2":
        return protocol["correspondent_type"], protocol["correspondent"]
    if override:
        return "legacy", _first_surname_token(override)
    if protocol:
        return "legacy", protocol["correspondent"]
    return "legacy", derive_surname(path, root)


def content_from_file(path, include_extension=False):
    path = Path(path)
    protocol = parse_protocol_filename(path)
    if protocol:
        value = protocol["content"] + (path.suffix if include_extension else "")
    else:
        value = path.name if include_extension else path.stem
    return re.sub(r"\s+", " ", value).strip()


def _state_path(root):
    return Path(root) / STATE_NAME


def load_manifest_ids(root):
    """Связывает имя PDF с внутренним ID СБИС без вывода ID в имя/АСУД."""
    path = Path(root) / DOWNLOAD_MANIFEST_NAME
    if not path.is_file():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        log.warning(f"Не удалось прочитать {DOWNLOAD_MANIFEST_NAME}: {exc}")
        return {}

    result = {}
    for manifest_key, entry in raw.items():
        if not isinstance(entry, dict):
            continue
        sbis_id = str(entry.get("sbis_id") or str(manifest_key).split("|", 1)[0]).strip()
        if not sbis_id:
            continue
        names = list(entry.get("files") or [])
        if entry.get("archive"):
            names.append(entry["archive"])
        for name in names:
            key = Path(str(name)).as_posix().casefold()
            result[key] = sbis_id
    return result


def load_state(root):
    path = _state_path(root)
    if not path.exists():
        return {"version": 1, "next_number": 1, "files": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data.get("files"), dict):
            raise ValueError("поле files отсутствует")
        data.setdefault("version", 1)
        data.setdefault("next_number", 1)
        return data
    except Exception as exc:
        log.warning(f"Не удалось прочитать {path.name}: {exc}. Начинаю новый state.")
        return {"version": 1, "next_number": 1, "files": {}}


def save_state(root, state):
    path = _state_path(root)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    os.replace(tmp, path)


def _file_signature(path):
    stat = Path(path).stat()
    return f"{stat.st_size}:{stat.st_mtime_ns}"


def prepare_items(root, files, state, surname_override="",
                  include_extension=False, doc_type_index=8,
                  person_doc_type_index=8, legal_doc_type_index=5,
                  manifest_ids=None):
    """Назначает стабильные номера и готовит doc_data для mix-flow."""
    root = Path(root).resolve()
    records = state.setdefault("files", {})
    try:
        next_number = max(1, int(state.get("next_number", 1)))
    except (TypeError, ValueError):
        next_number = 1

    items = []
    seen_state_keys = set()
    manifest_ids = manifest_ids or {}
    for path in files:
        path = Path(path).resolve()
        display_key = path.relative_to(root).as_posix()
        protocol = parse_protocol_filename(path)
        sbis_id = protocol["sbis_id"] if protocol else ""
        if not sbis_id:
            sbis_id = manifest_ids.get(display_key.casefold(), "")
        state_key = f"sbis:{sbis_id.casefold()}" if sbis_id else display_key
        if state_key in seen_state_keys:
            log.warning(f"Повтор одного ID СБИС в партии пропущен: {display_key}")
            continue
        seen_state_keys.add(state_key)

        entry = records.setdefault(state_key, {})
        if not isinstance(entry.get("number"), int) or entry["number"] < 1:
            entry["number"] = next_number
            next_number += 1
        entry["source_path"] = display_key
        if sbis_id:
            entry["sbis_id"] = sbis_id

        correspondent_type, correspondent = correspondent_from_file(
            path, root, surname_override)
        surname = _first_surname_token(correspondent)
        content = content_from_file(path, include_extension)
        signature = _file_signature(path)
        already_done = (entry.get("status") in DONE_STATUSES
                        and (bool(sbis_id) or entry.get("signature") == signature))
        legacy_type_index = int(doc_type_index)
        if correspondent_type == "person":
            item_type_index = int(person_doc_type_index)
        elif correspondent_type == "legal":
            item_type_index = int(legal_doc_type_index)
        else:
            item_type_index = legacy_type_index
        if correspondent_type in {"person", "legal"}:
            correspondent_field = correspondent
            correspondent_kind = correspondent_type
        else:
            correspondent_field = f"{correspondent} - -"
            correspondent_kind = "person"

        items.append({
            "key": state_key,
            "display_key": display_key,
            "path": path,
            "number": entry["number"],
            "sbis_id": sbis_id,
            "surname": surname,
            "correspondent": correspondent,
            "correspondent_type": correspondent_type,
            "content": content,
            "signature": signature,
            "already_done": already_done,
            "doc_data": {
                "row_idx": entry["number"],
                "содержание": content,
                "корреспондент": correspondent_field,
                "корреспондент_тип": correspondent_kind,
                "корр_найден": True,
                "корр_источник": "полные данные из имени файла СБИС",
                "тема": content,
                "тип_индекс": item_type_index,
                "тип_название": cfg.DOC_TYPE_MAP[item_type_index],
                "link": None,
                "файл": str(path),
                "sbis_id": sbis_id,
                "номер_обращения": f"б/н ({entry['number']})",
                "require_attachment": True,
            },
        })

    state["next_number"] = next_number
    return items


def _resolve_driver(base_dir):
    candidates = []
    env_driver = os.environ.get("ASUD_EDGE_DRIVER", "").strip()
    if env_driver:
        candidates.append(Path(env_driver))
    candidates.extend([
        Path(base_dir) / "msedgedriver.exe",
        Path(__file__).resolve().parent.parent / "msedgedriver.exe",
        Path.cwd() / "msedgedriver.exe",
    ])
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    return None


def _output_path(root):
    safe = re.sub(r"[^A-Za-zА-Яа-яЁё0-9_.\-]+", "_", Path(root).name).strip("_")
    return mix_flow._output_xlsx_path(f"СБИС_{safe or 'письма'}.xlsx")


def main():
    settings = cfg.load()
    options = _options(settings)
    cfg.setup_file_logger("sbis")
    cfg.keep_system_awake(True)
    started = time.monotonic()

    log.info("=" * 55)
    log.info("АСУД ИК — регистрация файлов, выгруженных из СБИС")
    log.info("=" * 55)

    raw_root = (os.environ.get("ASUD_SBIS_DIR", "").strip()
                or str(options.get("input_dir") or "").strip())
    if not raw_root:
        raw_root = input("Папка с выгруженными файлами СБИС: ").strip().strip('"')
    root = Path(raw_root).expanduser().resolve()
    if not root.is_dir():
        log.error(f"Папка не найдена: {root}")
        cfg.keep_system_awake(False)
        return

    if os.environ.get("ASUD_SBIS_RESET_STATE") == "1":
        state = {"version": 1, "next_number": 1, "files": {}}
        log.warning("State нумерации сброшен по флагу --reset-state")
    else:
        state = load_state(root)

    surname_override = (os.environ.get("ASUD_SBIS_SURNAME", "").strip()
                        or str(options.get("surname_override") or "").strip())
    extensions = _normalise_extensions(options.get("extensions"))
    files = scan_files(root, extensions)
    if not files:
        log.error(f"В {root} не найдено поддерживаемых файлов")
        cfg.keep_system_awake(False)
        return

    try:
        doc_type_index = int(options.get("doc_type_index", 8))
        person_doc_type_index = int(options.get("person_doc_type_index", 8))
        legal_doc_type_index = int(options.get("legal_doc_type_index", 5))
        if any(index not in cfg.DOC_TYPE_MAP for index in (
                doc_type_index, person_doc_type_index, legal_doc_type_index)):
            raise ValueError
    except (TypeError, ValueError):
        log.error("Индексы типов документов sbis должны быть числами от 1 до 8")
        cfg.keep_system_awake(False)
        return

    items = prepare_items(
        root, files, state,
        surname_override=surname_override,
        include_extension=bool(options.get("include_extension_in_content", False)),
        doc_type_index=doc_type_index,
        person_doc_type_index=person_doc_type_index,
        legal_doc_type_index=legal_doc_type_index,
        manifest_ids=load_manifest_ids(root),
    )
    save_state(root, state)  # фиксируем номера до первого запуска браузера

    invalid = [item for item in items if not item["correspondent"]]
    if invalid:
        log.error("Не удалось определить корреспондента для файлов:")
        for item in invalid[:20]:
            log.error(f"  {item['display_key']}")
        log.error("Переименуйте файлы по формату [ФЛ]/[ЮЛ] или задайте фамилию для старых файлов.")
        cfg.keep_system_awake(False)
        return

    pending = [item for item in items if not item["already_done"]]
    skipped = len(items) - len(pending)
    print(f"\nПапка: {root}")
    print(f"Тип ФЛ: [{person_doc_type_index}] "
          f"{cfg.DOC_TYPE_MAP[person_doc_type_index]}")
    print(f"Тип ЮЛ: [{legal_doc_type_index}] "
          f"{cfg.DOC_TYPE_MAP[legal_doc_type_index]}")
    print(f"Найдено: {len(items)}; уже обработано: {skipped}; к регистрации: {len(pending)}")
    if surname_override:
        print(f"Фамилия для всей партии: {derive_surname(Path('x'), Path.cwd(), surname_override)}")
    print("\nПроверка первых файлов:")
    for item in pending[:20]:
        type_label = {
            "person": "ФЛ",
            "legal": "ЮЛ",
            "legacy": "старый формат",
        }.get(item["correspondent_type"], item["correspondent_type"])
        corr_preview = item["correspondent"]
        if item["correspondent_type"] == "legal":
            corr_preview += " / - / -"
        elif item["correspondent_type"] == "legacy":
            corr_preview += " - -"
        print(f"  б/н ({item['number']}) | {type_label}: {corr_preview} | "
              f"{item['content'][:70]}")
    if len(pending) > 20:
        print(f"  ... ещё {len(pending) - 20}")

    if not pending:
        log.info("Все файлы уже зарегистрированы по state-файлу.")
        cfg.keep_system_awake(False)
        return

    answer = input("\nНачать регистрацию? (да/нет) [да]: ").strip().lower()
    if answer not in ("", "да", "д", "y", "yes"):
        print("Отменено.")
        cfg.keep_system_awake(False)
        return

    base_dir = cfg.get_base_dir()
    driver_path = _resolve_driver(base_dir)
    if driver_path is None:
        log.error("msedgedriver.exe не найден рядом с asud.exe/скриптом")
        cfg.keep_system_awake(False)
        return

    mix_flow.settings = settings
    driver = webdriver.Edge(
        service=EdgeService(executable_path=str(driver_path)),
        options=cfg.build_edge_options())
    set_driver_timeout(driver, settings.get(
        "asud_load_timeout_sec", cfg.DEFAULTS["asud_load_timeout_sec"]))

    output_path = _output_path(root)
    mix_flow._ensure_output_xlsx(output_path)
    done_count = duplicate_count = error_count = 0

    try:
        url = settings.get("asud_url", cfg.DEFAULTS["asud_url"])
        log.info(f"Открываю {url}")
        driver.get(url)
        wait_asud_loaded(driver)
        log.info(f"Реестр резолюций: {output_path}")

        total = len(pending)
        for position, item in enumerate(pending, 1):
            entry = state["files"][item["key"]]
            try:
                asud_id = mix_flow.create_one_document(
                    driver, item["doc_data"], position, total)
                status = mix_flow._last_result.get("status", "FAILED")
                if status not in DONE_STATUSES:
                    raise RuntimeError(f"регистрация завершилась со статусом {status}")
                if status == "DUPLICATE":
                    duplicate_count += 1
                    mix_flow.close_card_and_wait_main(driver)
                else:
                    done_count += 1
                entry.update({
                    "signature": item["signature"],
                    "status": status,
                    "surname": item["surname"],
                    "correspondent": item["correspondent"],
                    "correspondent_type": item["correspondent_type"],
                    "content": item["content"],
                    "sbis_id": item["sbis_id"],
                    "source_path": item["display_key"],
                    "asud_id": asud_id or "",
                    "last_error": "",
                })
                save_state(root, state)
                mix_flow._append_output_row(
                    output_path, item["doc_data"], asud_id, status=status)
            except Exception as exc:
                error_count += 1
                entry["status"] = "FAILED"
                entry["last_error"] = str(exc)[:500]
                save_state(root, state)
                log.error(f"ОШИБКА файл {item['display_key']}: {exc}")
                try:
                    driver.get(url)
                    wait_asud_loaded(driver)
                except Exception as recovery_exc:
                    log.error(f"Не удалось восстановить главную страницу: {recovery_exc}")

        elapsed = timedelta(seconds=int(time.monotonic() - started))
        log.info("=" * 55)
        log.info(f"ГОТОВО: зарегистрировано {done_count}, "
                 f"дубликатов {duplicate_count}, ошибок {error_count}, "
                 f"ранее обработано {skipped}; время {elapsed}")
        log.info(f"State: {_state_path(root)}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        cfg.keep_system_awake(False)


if __name__ == "__main__":
    main()
