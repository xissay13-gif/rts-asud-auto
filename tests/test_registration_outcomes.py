import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from flows import email as email_flow
from flows import mix


class _Element:
    def is_displayed(self):
        return True


class _ImmediateWait:
    """Browserless stand-in: every requested ASUD control is available."""

    def __init__(self, _driver, _timeout):
        pass

    def until(self, _predicate):
        return _Element()


class _DaemonDriver:
    def __init__(self):
        self.visited = []
        self.quit_called = False

    def get(self, url):
        self.visited.append(url)

    def quit(self):
        self.quit_called = True


def _document(dummy_path):
    return {
        "тема": "Обращение с сайта",
        "корреспондент": "644021, г. Омск, ул. Тестовая, д. 1",
        "корр_источник": "feedback-address",
        "корр_найден": True,
        "тип_название": "Письма, заявления и жалобы граждан, акционеров",
        "содержание": "Тестовое обращение",
        "link": "feedback-message-link",
        "row_idx": 1,
        "файл": str(dummy_path),
        "require_attachment": True,
    }


class MixRegistrationOutcomeTests(unittest.TestCase):
    def _run_create(self, registered, resolved=None, submission_uncertain=False):
        if resolved is None:
            resolved = registered
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        dummy_path = root / "dummy.msg"
        real_path = root / "feedback.msg"
        dummy_path.touch()
        real_path.touch()
        driver = object()
        doc = _document(dummy_path)

        error = None
        with (
            patch.object(mix, "settings", {
                "addressees": [],
                "outlook_dir": str(root),
            }),
            patch.object(mix, "WebDriverWait", _ImmediateWait),
            patch.object(mix.time, "sleep", return_value=None),
            patch.object(mix, "wait_and_click", return_value=True),
            patch.object(mix, "click", return_value=True),
            patch.object(mix, "fill_text"),
            patch.object(mix, "fill_correspondent_field", return_value=True),
            patch.object(mix, "fill_corr_number"),
            patch.object(mix, "fill_corr_date"),
            patch.object(mix, "fill_delivery_method"),
            patch.object(mix, "is_duplicate_warning", return_value=False),
            patch.object(mix, "find_msg_by_link", return_value=str(real_path)),
            patch.object(mix, "attach_content", return_value=True),
            patch.object(mix, "wait_modal_closed"),
            patch.object(
                mix, "register_and_resolve",
                return_value=(
                    registered, resolved, None, submission_uncertain
                )
            ),
            patch.object(mix, "move_to_done") as move_mock,
            patch.object(mix, "close_card_and_wait_main"),
        ):
            try:
                result = mix.create_one_document(driver, doc, index=1, total=1)
            except RuntimeError as exc:
                result = None
                error = exc

        return result, error, real_path, root, move_mock

    def test_registered_without_captured_asud_id_is_ok_and_moves_real_msg(self):
        result, error, real_path, root, move_mock = self._run_create(
            registered=True
        )

        self.assertIsNone(result)
        self.assertIsNone(error)
        self.assertEqual(mix._last_result["status"], "OK")
        move_mock.assert_called_once_with(str(real_path), str(root))

    def test_failed_registration_does_not_move_real_msg(self):
        result, error, _real_path, _root, move_mock = self._run_create(
            registered=False
        )

        self.assertIsNone(result)
        self.assertIsInstance(error, RuntimeError)
        self.assertEqual(str(error), "Регистрация документа не подтверждена")
        self.assertEqual(mix._last_result["status"], "FAILED")
        move_mock.assert_not_called()

    def test_registered_without_resolution_is_terminal_manual_status(self):
        result, error, _real_path, _root, move_mock = self._run_create(
            registered=True, resolved=False
        )

        self.assertIsNone(result)
        self.assertIsNone(error)
        self.assertEqual(mix._last_result["status"], "REGISTERED_ONLY")
        move_mock.assert_not_called()

    def test_uncertain_submission_is_terminal_and_never_retried_as_failed(self):
        result, error, _real_path, _root, move_mock = self._run_create(
            registered=False,
            resolved=False,
            submission_uncertain=True,
        )

        self.assertIsNone(result)
        self.assertIsNone(error)
        self.assertEqual(mix._last_result["status"], "SUBMISSION_UNKNOWN")
        move_mock.assert_not_called()


class EmailDaemonRetryKeyTests(unittest.TestCase):
    def test_terminal_marker_excludes_msg_when_move_to_errors_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            mailbox = Path(tmp)
            msg_path = mailbox / "uncertain.msg"
            msg_path.touch()

            with patch.object(
                    email_flow, "move_to_errors", return_value=False):
                quarantined = email_flow._quarantine_terminal_msg(
                    str(msg_path),
                    str(mailbox),
                    "SUBMISSION_UNKNOWN",
                    "результат регистрации не определён",
                )

            marker = Path(str(msg_path) + ".asud_terminal.json")
            self.assertTrue(quarantined)
            self.assertTrue(msg_path.exists())
            self.assertTrue(marker.exists())
            self.assertEqual(email_flow._list_root_msgs(str(mailbox)), [])

    def test_submission_unknown_moves_to_errors_immediately_without_retry(self):
        old_stop_flag = email_flow._stop_flag
        old_settings = email_flow.settings
        self.addCleanup(setattr, email_flow, "_stop_flag", old_stop_flag)
        self.addCleanup(setattr, email_flow, "settings", old_settings)

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            mailbox = root / "mailbox"
            mailbox.mkdir()
            msg_path = mailbox / "submission-unknown.msg"
            msg_path.touch()
            (root / "msedgedriver.exe").touch()

            folders_json = json.dumps([
                {"dir": str(mailbox), "output_suffix": "submission-unknown"},
            ])
            driver = _DaemonDriver()
            sleep_ticks = 0

            def stop_after_empty_followup_tick(_seconds):
                nonlocal sleep_ticks
                sleep_ticks += 1
                if sleep_ticks >= 2:
                    email_flow._stop_flag = True

            email_flow._stop_flag = False
            daemon_settings = {
                "email_watch_interval_sec": 1,
                "email_max_retries": 50,
                "asud_url": "https://asud.test/",
            }
            original_move_to_errors = email_flow.move_to_errors

            with (
                patch.dict(os.environ, {
                    "ASUD_EMAIL_FOLDERS_JSON": folders_json,
                    "ASUD_EMAIL_ROUND_ROBIN": "0",
                    "ASUD_EMAIL_PROCESS_MODE": "mix",
                    "ASUD_OUTPUT_SUFFIX": "",
                }, clear=False),
                patch.object(email_flow.cfg, "load", return_value=daemon_settings),
                patch.object(email_flow.cfg, "setup_file_logger"),
                patch.object(email_flow.cfg, "keep_system_awake"),
                patch.object(email_flow.cfg, "get_base_dir", return_value=str(root)),
                patch.object(email_flow.cfg, "build_edge_options", return_value=object()),
                patch.object(
                    email_flow, "_wait_for_folder",
                    side_effect=lambda folder: (True, folder),
                ),
                patch.object(email_flow.webdriver, "Edge", return_value=driver),
                patch.object(email_flow.signal, "signal"),
                patch.object(email_flow, "set_driver_timeout"),
                patch.object(email_flow, "wait_asud_loaded"),
                patch.object(
                    email_flow, "_parse_one_msg",
                    return_value={"тема": "test", "файл": str(msg_path)},
                ),
                patch.object(
                    email_flow, "_process_doc",
                    return_value=("SUBMISSION_UNKNOWN", "", None),
                ) as process_mock,
                patch.object(
                    email_flow, "move_to_errors",
                    wraps=original_move_to_errors,
                ) as errors_mock,
                patch.object(email_flow, "_print_doc_line"),
                patch.object(
                    email_flow, "_interruptible_sleep",
                    side_effect=stop_after_empty_followup_tick,
                ),
            ):
                email_flow.daemon_main()

            self.assertFalse(msg_path.exists())
            self.assertTrue((mailbox / "Ошибки" / msg_path.name).exists())

        process_mock.assert_called_once()
        errors_mock.assert_called_once()
        self.assertTrue(driver.quit_called)

    def test_registered_only_moves_to_errors_immediately_without_retry(self):
        old_stop_flag = email_flow._stop_flag
        old_settings = email_flow.settings
        self.addCleanup(setattr, email_flow, "_stop_flag", old_stop_flag)
        self.addCleanup(setattr, email_flow, "settings", old_settings)

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            mailbox = root / "mailbox"
            mailbox.mkdir()
            msg_path = mailbox / "registered-only.msg"
            msg_path.touch()
            (root / "msedgedriver.exe").touch()

            folders_json = json.dumps([
                {"dir": str(mailbox), "output_suffix": "registered-only"},
            ])
            driver = _DaemonDriver()
            sleep_ticks = 0

            def stop_after_empty_followup_tick(_seconds):
                nonlocal sleep_ticks
                sleep_ticks += 1
                if sleep_ticks >= 2:
                    email_flow._stop_flag = True

            email_flow._stop_flag = False
            daemon_settings = {
                "email_watch_interval_sec": 1,
                # A terminal registered document must not consume even the
                # first one of these otherwise available retry attempts.
                "email_max_retries": 50,
                "asud_url": "https://asud.test/",
            }
            original_move_to_errors = email_flow.move_to_errors

            with (
                patch.dict(os.environ, {
                    "ASUD_EMAIL_FOLDERS_JSON": folders_json,
                    "ASUD_EMAIL_ROUND_ROBIN": "0",
                    "ASUD_EMAIL_PROCESS_MODE": "mix",
                    "ASUD_OUTPUT_SUFFIX": "",
                }, clear=False),
                patch.object(email_flow.cfg, "load", return_value=daemon_settings),
                patch.object(email_flow.cfg, "setup_file_logger"),
                patch.object(email_flow.cfg, "keep_system_awake"),
                patch.object(email_flow.cfg, "get_base_dir", return_value=str(root)),
                patch.object(email_flow.cfg, "build_edge_options", return_value=object()),
                patch.object(
                    email_flow, "_wait_for_folder",
                    side_effect=lambda folder: (True, folder),
                ),
                patch.object(email_flow.webdriver, "Edge", return_value=driver),
                patch.object(email_flow.signal, "signal"),
                patch.object(email_flow, "set_driver_timeout"),
                patch.object(email_flow, "wait_asud_loaded"),
                patch.object(
                    email_flow, "_parse_one_msg",
                    return_value={"тема": "test", "файл": str(msg_path)},
                ),
                patch.object(
                    email_flow, "_process_doc",
                    return_value=("REGISTERED_ONLY", "ASUD/1", None),
                ) as process_mock,
                patch.object(
                    email_flow, "move_to_errors",
                    wraps=original_move_to_errors,
                ) as errors_mock,
                patch.object(email_flow, "_print_doc_line") as line_mock,
                patch.object(
                    email_flow, "_interruptible_sleep",
                    side_effect=stop_after_empty_followup_tick,
                ),
            ):
                email_flow.daemon_main()

            error_copy = mailbox / "Ошибки" / msg_path.name
            self.assertFalse(msg_path.exists())
            self.assertTrue(error_copy.exists())

        process_mock.assert_called_once()
        errors_mock.assert_called_once()
        self.assertIn(
            "резолюция не подтверждена",
            errors_mock.call_args.args[2],
        )
        self.assertEqual(
            line_mock.call_args.args[2:4],
            (
                "FAILED",
                "зарегистрирован без подтверждённой резолюции → Ошибки/",
            ),
        )
        self.assertTrue(driver.quit_called)

    def test_same_basename_in_different_folders_has_independent_retry_count(self):
        old_stop_flag = email_flow._stop_flag
        old_settings = email_flow.settings
        self.addCleanup(setattr, email_flow, "_stop_flag", old_stop_flag)
        self.addCleanup(setattr, email_flow, "settings", old_settings)

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            folder_a = root / "mailbox-a"
            folder_b = root / "mailbox-b"
            folder_a.mkdir()
            folder_b.mkdir()
            (folder_a / "same-name.msg").touch()
            (folder_b / "same-name.msg").touch()
            (root / "msedgedriver.exe").touch()

            folders_json = json.dumps([
                {"dir": str(folder_a), "output_suffix": "a"},
                {"dir": str(folder_b), "output_suffix": "b"},
            ])
            driver = _DaemonDriver()

            def stop_after_first_tick(_seconds):
                email_flow._stop_flag = True

            email_flow._stop_flag = False
            daemon_settings = {
                "email_watch_interval_sec": 1,
                "email_max_retries": 2,
                "asud_url": "https://asud.test/",
            }

            with (
                patch.dict(os.environ, {
                    "ASUD_EMAIL_FOLDERS_JSON": folders_json,
                    "ASUD_EMAIL_ROUND_ROBIN": "0",
                    "ASUD_EMAIL_PROCESS_MODE": "mix",
                    "ASUD_OUTPUT_SUFFIX": "",
                }, clear=False),
                patch.object(email_flow.cfg, "load", return_value=daemon_settings),
                patch.object(email_flow.cfg, "setup_file_logger"),
                patch.object(email_flow.cfg, "keep_system_awake"),
                patch.object(email_flow.cfg, "get_base_dir", return_value=str(root)),
                patch.object(email_flow.cfg, "build_edge_options", return_value=object()),
                patch.object(
                    email_flow, "_wait_for_folder",
                    side_effect=lambda folder: (True, folder),
                ),
                patch.object(email_flow.webdriver, "Edge", return_value=driver),
                patch.object(email_flow.signal, "signal"),
                patch.object(email_flow, "set_driver_timeout"),
                patch.object(email_flow, "wait_asud_loaded"),
                patch.object(email_flow, "_parse_one_msg", return_value={"тема": "test"}),
                patch.object(
                    email_flow, "_process_doc",
                    return_value=("FAILED", None, None),
                ) as process_mock,
                patch.object(email_flow, "move_to_errors") as errors_mock,
                patch.object(email_flow, "_print_doc_line"),
                patch.object(
                    email_flow, "_interruptible_sleep",
                    side_effect=stop_after_first_tick,
                ),
            ):
                email_flow.daemon_main()

        self.assertEqual(process_mock.call_count, 2)
        errors_mock.assert_not_called()
        self.assertTrue(driver.quit_called)


if __name__ == "__main__":
    unittest.main()
