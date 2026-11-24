import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import app
from flows.attach_replies import process_job
from shared.asud_reports import execution_candidates, unique
from shared.reply_files import (
    ReplyJob, ReplyJournal, folder_lock, number_from_filename, number_matches, scan_folder,
)


class ReplyFileTests(unittest.TestCase):
    def test_explicit_number_formats(self):
        for filename, expected in (
            ('41359__Ответ.pdf', '41359'), ('41359.docx', '41359'),
            ('ОРТС_8_41359__Приложение.pdf', 'ОРТС/8/41359'),
        ):
            self.assertEqual(number_from_filename(filename), expected)

    def test_does_not_guess_numbers_in_addresses_or_dates(self):
        for filename in ('Ответ Иванову 41359.pdf', '07.09.2026 Ответ.pdf',
                         'Архитекторов 10-54.pdf', '41359__.pdf', '41359-ответ.pdf'):
            with self.subTest(filename=filename), self.assertRaises(ValueError):
                number_from_filename(filename)

    def test_exact_number_not_substring(self):
        self.assertTrue(number_matches('ОРТС/8/41359', '41359'))
        self.assertFalse(number_matches('ОРТС/8/141359', '41359'))
        self.assertFalse(number_matches('ОРТС/5/41359', 'ОРТС/8/41359'))

    def test_groups_attachments_but_rejects_alias_collision(self):
        with tempfile.TemporaryDirectory() as folder:
            for name in ('41359__Ответ.pdf', '41359__Приложение.pdf', 'Ответ.pdf'):
                (Path(folder) / name).write_bytes(b'file')
            jobs, errors = scan_folder(folder)
            self.assertEqual(len(jobs), 1)
            self.assertEqual(len(jobs[0].files), 2)
            self.assertEqual(len(errors), 1)
            (Path(folder) / 'ОРТС_8_41359.pdf').write_bytes(b'file')
            self.assertEqual(scan_folder(folder)[0], [])

    def test_corrupt_journal_blocks_processing(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / '.asud_replies.json'
            path.write_text('{broken', encoding='utf-8')
            with self.assertRaises(ValueError):
                ReplyJournal(folder)
            path.write_text(json.dumps({'version': 1, 'reports': {'x': {'status': '???'}}}))
            with self.assertRaises(ValueError):
                ReplyJournal(folder)

    def test_folder_lock_releases_on_exception(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaisesRegex(ValueError, 'stop'):
                with folder_lock(folder):
                    with self.assertRaises(RuntimeError):
                        with folder_lock(folder):
                            self.fail('second process should not acquire lock')
                    raise ValueError('stop')
            with folder_lock(folder):
                pass


class ExecutionSelectionTests(unittest.TestCase):
    def test_selects_execution_below_matching_document_not_notification(self):
        rows = [('doc', 'Входящий № ОРТС/8/41359 26.08.2026', True),
                ('notice', 'Подходит контрольный срок', False),
                ('task', 'Исполнение', False),
                ('other_doc', 'Входящий № ОРТС/8/141359', True),
                ('other_task', 'Исполнение', False)]
        self.assertEqual(execution_candidates(rows, '41359'), [('task', 'ОРТС/8/41359')])

    def test_duplicate_tasks_are_not_automatically_selected(self):
        rows = [('doc', 'Входящий № ОРТС/8/41359', True),
                ('task1', 'Исполнение', False), ('task2', 'Исполнение', False)]
        with self.assertRaises(RuntimeError):
            unique(execution_candidates(rows, '41359'), 'поручение')

    def test_unparseable_document_resets_group(self):
        rows = [('doc', 'Входящий № ОРТС/8/41359', True),
                ('other_doc', 'Черновик', True), ('task', 'Исполнение', False)]
        self.assertEqual(execution_candidates(rows, '41359'), [])


class ReplyProcessingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.file = Path(self.temp.name) / '41359__Ответ.pdf'
        self.file.write_bytes(b'first version')
        self.job = ReplyJob('41359', (self.file,))
        self.journal = ReplyJournal(self.temp.name)
        self.driver = Mock()
        self.identity = patch('flows.attach_replies.account_identity', return_value=('head', 'Иванов'))
        self.open = patch('flows.attach_replies.open_execution', return_value='ОРТС/8/41359')
        self.close = patch('flows.attach_replies.close_card')
        for patcher in (self.identity, self.open, self.close):
            patcher.start()
            self.addCleanup(patcher.stop)

    def process(self):
        return process_job(self.driver, self.journal, 'head', self.job, 'Ответ прикреплен.')

    def test_success_skips_repeat_and_protects_changed_content(self):
        def save(*args):
            args[5]()
        with patch('flows.attach_replies.add_reply_report', side_effect=save) as report:
            self.assertEqual(self.process(), 'saved')
            self.journal = ReplyJournal(self.temp.name)
            self.assertEqual(self.process(), 'skip')
            self.file.write_bytes(b'corrected reply')
            self.assertEqual(self.process(), 'review')
            report.assert_called_once()

    def test_uncertain_save_survives_restart_and_is_not_retried(self):
        def save(*args):
            args[5]()
            raise TimeoutError('server response lost')
        with patch('flows.attach_replies.add_reply_report', side_effect=save) as report:
            with self.assertRaises(TimeoutError):
                self.process()
            self.journal = ReplyJournal(self.temp.name)
            self.assertEqual(self.process(), 'review')
            self.assertEqual(self.journal.previous('head', '41359')['status'], 'pending')
            report.assert_called_once()

    def test_upload_failure_can_retry_without_pending_save(self):
        with patch('flows.attach_replies.add_reply_report', side_effect=TimeoutError('upload')):
            with self.assertRaises(TimeoutError):
                self.process()
        self.assertEqual(self.journal.previous('head', '41359')['status'], 'failed')

    def test_file_changed_during_upload_prevents_save(self):
        def save(*args):
            self.file.write_bytes(b'changed during upload')
            args[5]()
            self.fail('must not save a changed file')
        with patch('flows.attach_replies.add_reply_report', side_effect=save):
            with self.assertRaisesRegex(RuntimeError, 'изменились'):
                self.process()
        self.assertEqual(self.journal.previous('head', '41359')['status'], 'failed')

    def test_pending_matches_renamed_full_number(self):
        self.journal.record('head', self.job, 'pending', self.job.fingerprints())
        self.assertIsNotNone(self.journal.previous('head', 'ОРТС/8/41359'))
        self.assertIsNone(self.journal.previous('another_head', '41359'))

    def test_account_change_prevents_opening_document(self):
        with patch('flows.attach_replies.account_identity', return_value=('other', 'Другой')):
            with self.assertRaisesRegex(RuntimeError, 'Учётная запись'):
                self.process()


class ReplyDispatchTests(unittest.TestCase):
    def test_folder_mode_does_not_dispatch_email_and_preserves_cli_folder(self):
        module = types.ModuleType('flows.attach_replies')
        module.main = Mock()
        preset = {'id': 'replies', 'mode': 'attach-replies', 'folder': 'preset-folder'}
        with patch.object(sys, 'argv', ['app.py', '--preset=replies', '--folder=chosen', '--watch']), \
                patch.object(app.cfg, 'load', return_value={'presets': [preset]}), \
                patch.dict(sys.modules, {'flows.attach_replies': module}), \
                patch.dict(os.environ, {}, clear=False), \
                patch.object(app, 'pick_source', side_effect=AssertionError('wrong dispatch')):
            app.main()
            self.assertEqual(os.environ['ASUD_REPLY_FOLDER'], 'chosen')
            self.assertEqual(os.environ['ASUD_REPLY_WATCH'], '1')
            module.main.assert_called_once()


if __name__ == '__main__':
    unittest.main()
