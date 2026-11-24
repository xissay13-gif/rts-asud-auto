"""Opt-in local fixture checks: ASUD_REPLY_BROWSER_TESTS=1, no ASUD connection."""
import os
import tempfile
import unittest
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.edge.service import Service

from shared.asud_reports import account_identity, add_reply_report, open_execution, wait_reports_home


@unittest.skipUnless(os.environ.get('ASUD_REPLY_BROWSER_TESTS') == '1', 'local browser tests opt-in')
class ReplyBrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        options = webdriver.EdgeOptions()
        options.add_argument('--headless=new')
        options.add_argument('--window-size=1400,1000')
        driver_path = os.environ.get('ASUD_TEST_EDGEDRIVER')
        service = Service(driver_path) if driver_path else Service()
        cls.driver = webdriver.Edge(service=service, options=options)

    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()

    def setUp(self):
        self.driver.get((Path(__file__).parent / 'fixtures/asud_reply_flow.html').resolve().as_uri())
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.files = tuple(Path(self.temporary.name) / name
                           for name in ('41359__Ответ.pdf', '41359__Приложение.txt'))
        for path in self.files:
            path.write_bytes(b'Local test attachment')
        self.pending = False

    def before_save(self):
        self.pending = True

    def run_report(self):
        number = open_execution(self.driver, '41359', timeout=3)
        add_reply_report(self.driver, 'Department_Head', number, self.files,
                         'Ответ прикреплен.', self.before_save, timeout=2)

    def test_full_flow_uses_own_task_uploads_once_and_reopens(self):
        self.assertEqual(account_identity(self.driver)[0], 'Department_Head')
        self.run_report()
        self.assertTrue(self.pending)
        self.assertEqual(self.driver.execute_script('return window.saveCount'), 1)
        self.assertEqual(self.driver.execute_script('return window.uploadCount'), 1)
        self.assertEqual(self.driver.execute_script('return window.openCount'), 2)

    def test_files_in_input_do_not_count_as_successful_upload(self):
        self.driver.execute_script('window.failUpload = true')
        with self.assertRaises(TimeoutException):
            self.run_report()
        self.assertFalse(self.pending)
        self.assertEqual(self.driver.execute_script('return window.saveCount'), 0)

    def test_closed_dialog_does_not_count_as_saved(self):
        self.driver.execute_script('window.failSave = true')
        with self.assertRaises(TimeoutException):
            self.run_report()
        self.assertTrue(self.pending)
        self.assertEqual(self.driver.execute_script('return window.saveCount'), 1)

    def test_saved_text_must_survive_reopening(self):
        self.driver.execute_script('window.failReopen = true')
        with self.assertRaises(TimeoutException):
            self.run_report()
        self.assertTrue(self.pending)
        self.assertEqual(self.driver.execute_script('return window.saveCount'), 1)

    def test_existing_attachment_blocks_duplicate_even_without_journal(self):
        self.run_report()
        with self.assertRaisesRegex(RuntimeError, 'одноимённое вложение'):
            add_reply_report(self.driver, 'Department_Head', 'ОРТС/8/41359', self.files,
                             'Ответ прикреплен.', self.before_save, timeout=2)
        self.assertEqual(self.driver.execute_script('return window.saveCount'), 1)

    def test_multiline_report_is_verified(self):
        number = open_execution(self.driver, '41359', timeout=3)
        add_reply_report(self.driver, 'Department_Head', number, self.files,
                         'Ответ прикреплен.\nПриложен расчёт.', self.before_save, timeout=2)
        self.assertTrue(self.pending)

    def test_textarea_replaced_on_tab_is_reacquired(self):
        self.driver.execute_script('window.rerenderOnBlur = true')
        self.run_report()
        self.assertGreater(self.driver.execute_script('return window.blurRenders'), 0)
        self.assertEqual(self.driver.execute_script('return window.uploadCount'), 1)
        self.assertEqual(self.driver.execute_script('return window.saveCount'), 1)

    def test_executor_home_needs_neither_create_button_nor_existing_documents(self):
        self.driver.execute_script("document.querySelectorAll('tr.obj-list-rec').forEach(el => el.remove())")
        wait_reports_home(self.driver, timeout=2)
        self.assertEqual(account_identity(self.driver)[0], 'Department_Head')

    def test_whole_dialog_replaced_after_upload_is_reacquired_without_reupload(self):
        self.driver.execute_script('window.rerenderAfterUpload = true')
        self.run_report()
        self.assertEqual(self.driver.execute_script('return window.uploadRenders'), 1)
        self.assertEqual(self.driver.execute_script('return window.uploadCount'), 1)
        self.assertEqual(self.driver.execute_script('return window.saveCount'), 1)


if __name__ == '__main__':
    unittest.main()
