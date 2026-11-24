"""Exercise the packaged reply flow through a real Edge session, offline.

Usage: python scripts/smoke_frozen_edge.py --exe dist/asud_reports.exe --driver path/to/msedgedriver.exe
Only local test pages and generated attachments are opened. User settings,
ASUD, and production reply files are never used by this check.
"""

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile


HTML = '''<!doctype html><meta charset="utf-8"><title>Frozen Edge smoke</title>
<table id="CABINET_MENU__RECEIVED__ALL_ACTIVE__TO_EXECUTION"><tr><td>Local execution menu</td></tr></table>
<span id="app-acc-info" x-account-name="frozen-edge-smoke-account">Frozen Edge smoke</span>
<table><tr class="obj-list-rec"><td>Local test data</td></tr></table>
'''


def check(exe, driver, *, exercise_report=False):
    exe, driver = Path(exe).resolve(strict=True), Path(driver).resolve(strict=True)
    # TemporaryDirectory owns exactly this directory and releases it on exit.
    with tempfile.TemporaryDirectory(prefix='asud-frozen-edge-') as temporary:
        root = Path(temporary).resolve()
        shutil.copy2(exe, root / 'asud_reports.exe')
        shutil.copy2(driver, root / 'msedgedriver.exe')
        replies = root / 'replies'
        replies.mkdir()
        page = root / 'local.html'
        if exercise_report:
            fixture = Path(__file__).resolve().parent.parent / 'tests/fixtures/asud_reply_flow.html'
            html = fixture.read_text(encoding='utf-8')
            html = html.replace('window.rerenderOnBlur = false;', 'window.rerenderOnBlur = true;')
            html = html.replace('window.rerenderAfterUpload = false;', 'window.rerenderAfterUpload = true;')
            page.write_text(html, encoding='utf-8')
            (replies / '41359__Ответ.msg').write_bytes(b'Local smoke test attachment')
            account = b'Department_Head'
        else:
            page.write_text(HTML, encoding='utf-8')
            account = b'frozen-edge-smoke-account'
        (root / 'settings.json').write_text(json.dumps({
            'asud_url': page.as_uri(),
            'reply_reports': {'folder': str(replies), 'switch_to_user': '', 'file_stable_seconds': 0},
        }), encoding='utf-8')
        environment = {key: value for key, value in os.environ.items()
                       if not key.startswith('ASUD_')}
        environment['PYTHONIOENCODING'] = 'utf-8'
        result = subprocess.run(
            [str(root / 'asud_reports.exe'), '--headless'],
            cwd=root, env=environment, capture_output=True, timeout=60,
        )
        output = result.stdout + result.stderr
        if result.returncode or account not in output:
            raise RuntimeError(
                f'Frozen Edge smoke failed (exit {result.returncode}):\n'
                + output.decode('utf-8', errors='replace'))
        if exercise_report:
            journal = json.loads((replies / '.asud_replies.json').read_text(encoding='utf-8'))
            reports = list(journal['reports'].values())
            assert len(reports) == 1 and reports[0]['status'] == 'saved', journal
            assert reports[0]['number'] == 'ОРТС/8/41359', journal
            print('PASS: packaged EXE uploaded a local MSG, handled textarea and dialog '
                  'replacement, saved the report, reopened it, and recorded verified success.')
        else:
            print('PASS: packaged reply flow started Edge, read the local page, '
                  'processed an empty folder, and exited successfully.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--exe', required=True)
    parser.add_argument('--driver', help='Existing EdgeDriver; otherwise resolve through Selenium Manager')
    parser.add_argument('--exercise-report', action='store_true',
                        help='Exercise a complete upload with a generated file and local UI fixture')
    args = parser.parse_args()
    if not args.driver:
        from selenium.webdriver.common.selenium_manager import SeleniumManager
        args.driver = SeleniumManager().binary_paths(['--browser', 'MicrosoftEdge'])['driver_path']
    check(args.exe, args.driver, exercise_report=args.exercise_report)
