"""Folder input and durable journal for department reply reports."""

import hashlib
import json
import os
import re
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path


EXTENSIONS = {'.pdf', '.doc', '.docx', '.rtf', '.txt', '.xls', '.xlsx',
              '.jpg', '.jpeg', '.png', '.tif', '.tiff', '.zip', '.msg'}
NUMBER = re.compile(r'(?:[А-ЯЁA-Z]+/\d+/)?\d+', re.IGNORECASE)


def normalize_number(value):
    value = str(value).strip().upper().replace('_', '/')
    if not NUMBER.fullmatch(value):
        raise ValueError(f'Некорректный номер АСУД: {value!r}')
    return value


def number_from_filename(filename):
    """Only an explicit number prefix; never guess from arbitrary digits."""
    stem = Path(filename).stem
    prefix, separator, description = stem.partition('__')
    if separator and not description.strip():
        raise ValueError('После __ должно быть название ответа')
    return normalize_number(prefix)


def number_matches(actual, requested):
    actual, requested = normalize_number(actual), normalize_number(requested)
    return actual == requested if '/' in requested else actual.split('/')[-1] == requested


@dataclass(frozen=True)
class ReplyJob:
    number: str
    files: tuple[Path, ...]

    def fingerprints(self):
        result = []
        for path in self.files:
            with path.open('rb') as stream:
                digest = hashlib.file_digest(stream, 'sha256').hexdigest()
            result.append({'name': path.name, 'sha256': digest})
        return result


def scan_folder(folder):
    """One report per number. Reject aliases that could target the same report."""
    folder = Path(folder).resolve(strict=True)
    if not folder.is_dir():
        raise ValueError(f'Это не папка: {folder}')
    grouped, errors = {}, []
    for path in sorted(folder.iterdir(), key=lambda p: p.name.casefold()):
        if not path.is_file() or path.name.startswith(('.', '~')):
            continue
        if path.suffix.lower() not in EXTENSIONS:
            errors.append(f'{path.name}: неподдерживаемое расширение')
            continue
        try:
            number = number_from_filename(path.name)
            if not path.stat().st_size:
                raise ValueError('пустой файл')
            grouped.setdefault(number, []).append(path)
        except ValueError as exc:
            errors.append(f'{path.name}: {exc}')
    ambiguous = set()
    for number in grouped:
        if '/' not in number:
            aliases = {other for other in grouped
                       if other != number and other.split('/')[-1] == number}
            if aliases:
                ambiguous.update({number, *aliases})
    for number in sorted(ambiguous):
        errors.append(f'{number}: одновременно есть короткий и полный номер; '
                      'используйте одинаковый полный номер во всех именах')
    return [ReplyJob(number, tuple(paths)) for number, paths in grouped.items()
            if number not in ambiguous], errors


@contextmanager
def folder_lock(folder):
    """An OS lock releases on process exit, including an interrupted upload."""
    path = Path(folder) / '.asud_replies.lock'
    with path.open('a+b') as stream:
        stream.seek(0, os.SEEK_END)
        if stream.tell() == 0:
            stream.write(b'0')
            stream.flush()
        stream.seek(0)
        if os.name == 'nt':
            import msvcrt
            lock = lambda: msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            unlock = lambda: msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            lock = lambda: fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
            unlock = lambda: fcntl.flock(stream, fcntl.LOCK_UN)
        try:
            lock()
        except OSError as exc:
            raise RuntimeError('Эта папка уже обрабатывается другим запуском') from exc
        try:
            yield
        finally:
            unlock()


class ReplyJournal:
    def __init__(self, folder):
        self.path = Path(folder) / '.asud_replies.json'
        self.data = {'version': 1, 'reports': {}}
        if self.path.exists():
            self.data = json.loads(self.path.read_text(encoding='utf-8'))
            if (self.data.get('version') != 1 or
                    not isinstance(self.data.get('reports'), dict) or
                    any(not isinstance(row, dict) or row.get('status') not in
                        {'pending', 'saved', 'failed'}
                        for row in self.data['reports'].values())):
                raise ValueError(f'Повреждён журнал: {self.path}')

    def previous(self, account, number):
        # Short and full names must share the duplicate guard across runs.
        matches = [entry for entry in self.data['reports'].values()
                   if entry.get('account') == account and
                   (number_matches(entry['number'], number) or
                    number_matches(number, entry['number']))]
        protected = [entry for entry in matches if entry['status'] != 'failed']
        return (protected or matches or [None])[0]

    def record(self, account, job, status, fingerprints, detail='', resolved_number=None):
        number = resolved_number or job.number
        key = f'{account}|{number}'
        self.data['reports'][key] = {
            'account': account, 'number': number, 'status': status,
            'files': fingerprints, 'detail': detail,
            'updated_at': datetime.now(timezone.utc).isoformat(),
        }
        temporary = self.path.with_suffix('.json.tmp')
        with temporary.open('w', encoding='utf-8') as stream:
            json.dump(self.data, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, self.path)
