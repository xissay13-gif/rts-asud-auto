import contextlib
import io
import os
import unittest
from unittest.mock import patch

import reports_app


class StandaloneReportsTests(unittest.TestCase):
    def test_runs_uploader_directly_without_source_or_mode_menu(self):
        with patch('flows.attach_replies.main') as upload, patch.dict(os.environ, {}, clear=True):
            result = reports_app.main(['--folder', 'D:\\Ответы', '--watch', '--headless'])
            self.assertEqual(result, 0)
            upload.assert_called_once_with()
            self.assertEqual(os.environ['ASUD_REPLY_FOLDER'], 'D:\\Ответы')
            self.assertEqual(os.environ['ASUD_REPLY_WATCH'], '1')
            self.assertEqual(os.environ['ASUD_HEADLESS'], '1')

    def test_default_run_leaves_folder_to_settings(self):
        with patch('flows.attach_replies.main') as upload, patch.dict(os.environ, {}, clear=True):
            self.assertEqual(reports_app.main([]), 0)
            self.assertNotIn('ASUD_REPLY_FOLDER', os.environ)
            self.assertEqual(os.environ['ASUD_REPLY_WATCH'], '0')
            upload.assert_called_once_with()

    def test_other_modes_are_not_available(self):
        with patch('flows.attach_replies.main') as upload, contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as error:
                reports_app.main(['--mode=email'])
            self.assertEqual(error.exception.code, 2)
            upload.assert_not_called()

    def test_failure_returns_nonzero_status(self):
        with patch('flows.attach_replies.main', side_effect=RuntimeError('startup failure')), \
                patch('reports_app.logging.getLogger'):
            self.assertEqual(reports_app.main([]), 1)


if __name__ == '__main__':
    unittest.main()
