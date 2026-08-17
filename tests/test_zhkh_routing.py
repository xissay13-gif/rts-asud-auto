import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from flows import email as email_flow
from flows import mix
from shared import config as cfg
from shared.zhkh_routing import normalize_topic_title


BASMANOV = "Басманов Александр Владимирович"

EXCLUDED_TOPICS = {
    "2.1": "Отсутствие отопления",
    "2.2": "Отсутствие водоснабжения",
    "2.3": "Нарушение температурного режима подачи воды",
    "2.16": "Порыв труб",
    "2.25": "Затопление подвала",
    "5.1": "Качество оказания коммунальных услуг",
    "5.2": "Сроки оказания коммунальных услуг",
    "23": "Некачественная поставка ресурса",
}

# The config still carries the business-table code/title mapping. Its action
# for an exact match is now terminal exclusion from ASUD, not reassignment.
TOPIC_RULES = dict(EXCLUDED_TOPICS)


class _DummyMessage:
    def close(self):
        pass


class _Element:
    def is_displayed(self):
        return True


class _ImmediateWait:
    """Browserless wait: every requested ASUD control exists."""

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


def _zhkh_payload(topic):
    return {
        "фамилия": "Иванов",
        "имя": "Иван",
        "отчество": "Иванович",
        "фио": "Иванов Иван Иванович",
        "номер_обращения": "55-2026-1",
        "дата_обращения": "17.08.2026",
        "планируемая_дата": None,
        "тема_обращения": topic,
        "адрес": "г. Омск, ул. Тестовая, д. 1",
        "телефон": None,
        "email": None,
    }


def _parse_zhkh_email(topic, routes=TOPIC_RULES, msg_path="zhkh-routing.msg"):
    values = {
        "subject": f"55-2026-1 — {topic or 'Без тематики'}",
        "body": "Структурированное обращение ГИС ЖКХ",
    }
    old_settings = email_flow.settings
    email_flow.settings = {
        "unknown_correspondent": "Неизвестный Неизвестный Неизвестный",
        "default_type_idx": 8,
        "addressees": [BASMANOV],
        "zhkh_excluded_topics": routes,
    }
    try:
        with (
            patch.object(
                email_flow.extract_msg,
                "openMsg",
                return_value=_DummyMessage(),
            ),
            patch.object(
                email_flow,
                "_safe_field",
                side_effect=lambda _msg, attr, *_args: values[attr],
            ),
            patch.object(email_flow, "_msg_link", return_value="zhkh-1"),
            patch.object(
                email_flow, "_msg_date_prefix", return_value="2026-08-17"
            ),
            patch.object(
                email_flow,
                "parse_zhkh_body",
                return_value=_zhkh_payload(topic),
            ),
            patch.object(email_flow, "okrug_from_textbody", return_value=None),
            patch.object(email_flow, "_compute_zhkh_deadline", return_value=None),
        ):
            return email_flow._parse_one_msg(msg_path, process_mode="mix")
    finally:
        email_flow.settings = old_settings


def _excluded_doc(msg_path, code="5.2"):
    return {
        "row_idx": 1,
        "тема": "ГИС ЖКХ 55-2026-1",
        "содержание": "ГИС ЖКХ 55-2026-1, тестовый адрес",
        "корреспондент": "Иванов Иван Иванович",
        "корреспондент_тип": "person",
        "корр_источник": "zhkh",
        "корр_найден": True,
        "тип_название": "Письма, заявления и жалобы граждан, акционеров",
        "link": "zhkh-1",
        "файл": str(msg_path),
        "тема_обращения": EXCLUDED_TOPICS[code],
        "zhkh_topic_code": code,
        "skip_asud_registration": True,
        "addressees_override": None,
    }


class ZhkhExclusionPolicyTests(unittest.TestCase):
    def test_normalization_is_unicode_case_and_whitespace_stable(self):
        self.assertEqual(
            normalize_topic_title(
                "  НЕКАЧЕСТВЕННАЯ\u00a0\u00a0ПОСТАВКА   РЕСУРСА  "
            ),
            "некачественная поставка ресурса",
        )
        self.assertEqual(normalize_topic_title("  Ёлка  "), "елка")
        self.assertEqual(normalize_topic_title(None), "")

    def test_all_eight_exact_topics_become_terminal_asud_exclusions(self):
        for code, title in EXCLUDED_TOPICS.items():
            with self.subTest(code=code, title=title):
                doc = _parse_zhkh_email(title)

                self.assertIsNotNone(doc)
                self.assertEqual(doc["тема_обращения"], title)
                self.assertEqual(doc["zhkh_topic_code"], code)
                self.assertTrue(doc["skip_asud_registration"])
                self.assertIsNone(doc.get("addressees_override"))

    def test_normalized_exact_title_is_still_excluded(self):
        doc = _parse_zhkh_email(
            "  СРОКИ\u00a0  ОКАЗАНИЯ  КОММУНАЛЬНЫХ УСЛУГ "
        )

        self.assertTrue(doc["skip_asud_registration"])
        self.assertEqual(doc["zhkh_topic_code"], "5.2")

    def test_empty_or_old_rc9_user_config_cannot_disable_builtin_exclusion(self):
        # RC9 settings may still contain the former Zhukov route and no useful
        # exclusion override. The built-in safety policy must win regardless.
        for user_exclusions in ({}, None):
            with self.subTest(user_exclusions=user_exclusions):
                doc = _parse_zhkh_email(
                    EXCLUDED_TOPICS["2.1"], routes=user_exclusions
                )
                self.assertTrue(doc["skip_asud_registration"])
                self.assertEqual(doc["zhkh_topic_code"], "2.1")

    def test_similar_missing_or_code_prefixed_titles_are_not_excluded(self):
        nonmatches = (
            "Порыв трубы",
            "Качество оказания коммунальных услуг населению",
            "Срок оказания коммунальных услуг",
            "Отсутствие отопления и водоснабжения",
            "2.1 Отсутствие отопления",
            "2.10 Отсутствие отопления",
            "23.1 Некачественная поставка ресурса",
            "Некачественная поставка ресурсов",
            "Другая тема",
            "",
            None,
        )
        for title in nonmatches:
            with self.subTest(title=title):
                doc = _parse_zhkh_email(title)
                self.assertIsNotNone(doc)
                self.assertFalse(doc["skip_asud_registration"])
                self.assertIsNone(doc["zhkh_topic_code"])
                self.assertIsNone(doc.get("addressees_override"))

    def test_shipped_default_and_example_exclude_the_same_eight_titles(self):
        sources = [cfg.DEFAULTS["zhkh_excluded_topics"]]
        example_path = Path(__file__).resolve().parents[1] / "settings.json.example"
        with example_path.open(encoding="utf-8") as stream:
            sources.append(json.load(stream)["zhkh_excluded_topics"])

        for source_name, rules in zip(("DEFAULTS", "settings example"), sources):
            for code, title in EXCLUDED_TOPICS.items():
                with self.subTest(source=source_name, code=code):
                    doc = _parse_zhkh_email(title, routes=rules)
                    self.assertTrue(doc["skip_asud_registration"])
                    self.assertEqual(doc["zhkh_topic_code"], code)


class UnmatchedZhkhDefaultRouteTests(unittest.TestCase):
    def test_unmatched_topic_still_uses_basmanov_and_registers_normally(self):
        parsed = _parse_zhkh_email("Цены (тарифы) на коммунальные услуги")
        self.assertFalse(parsed["skip_asud_registration"])

        driver = object()
        parsed.update({
            "row_idx": 1,
            "link": None,
            "файл": None,
        })
        with (
            patch.object(
                mix,
                "settings",
                {"addressees": [BASMANOV], "outlook_dir": ""},
            ),
            patch.object(mix, "WebDriverWait", _ImmediateWait),
            patch.object(mix.time, "sleep", return_value=None),
            patch.object(mix, "wait_and_click", return_value=True),
            patch.object(mix, "click", return_value=True),
            patch.object(mix, "fill_text"),
            patch.object(mix, "fill_correspondent_field", return_value=True),
            patch.object(mix, "fill_corr_number"),
            patch.object(mix, "fill_corr_date"),
            patch.object(mix, "add_addressee", return_value=True) as add_mock,
            patch.object(mix, "fill_delivery_method"),
            patch.object(mix, "is_duplicate_warning", return_value=False),
            patch.object(mix, "find_msg_by_link", return_value=None),
            patch.object(
                mix,
                "register_and_resolve",
                return_value=(True, True, "ОРТС/8/2026/1", False),
            ) as register_mock,
            patch.object(mix, "close_card_and_wait_main"),
        ):
            result = mix.create_one_document(driver, parsed, index=1, total=1)

        self.assertEqual(result, "ОРТС/8/2026/1")
        add_mock.assert_called_once_with(driver, BASMANOV)
        register_mock.assert_called_once()


class TerminalExclusionProcessingTests(unittest.TestCase):
    def test_exclusion_keeps_msg_and_writes_terminal_marker_without_side_effects(self):
        with tempfile.TemporaryDirectory() as tmp:
            mailbox = Path(tmp)
            msg_path = mailbox / "excluded.msg"
            msg_path.touch()
            doc = _excluded_doc(msg_path)

            with (
                patch.dict(
                    os.environ,
                    {"ASUD_DELETE_AFTER_DONE": "1"},
                    clear=False,
                ),
                patch.object(
                    email_flow.mix_flow,
                    "create_one_document",
                ) as create_mock,
                patch.object(email_flow, "_xlsx_path") as xlsx_path_mock,
                patch.object(email_flow, "_ensure_dated_xlsx") as ensure_mock,
                patch.object(email_flow, "_append_dated_row") as append_mock,
                patch.object(email_flow, "move_to_done") as done_mock,
                patch.object(email_flow, "move_to_errors") as errors_mock,
                patch.object(email_flow, "move_to_drafts") as drafts_mock,
            ):
                result = email_flow._process_doc(
                    object(),
                    doc,
                    str(mailbox),
                    str(mailbox),
                    1,
                    1,
                    in_daemon=True,
                    process_mode="mix",
                    output_suffix="ГИСЖКХ",
                )

            marker = Path(str(msg_path) + ".asud_terminal.json")
            self.assertEqual(result, ("EXCLUDED", None, None))
            # The downloader deduplicates by the MSG itself, so it must stay
            # in the root. The sidecar alone hides it from the ASUD scanner.
            self.assertTrue(msg_path.exists())
            self.assertTrue(marker.exists())
            self.assertFalse(Path(str(marker) + ".tmp").exists())
            with marker.open(encoding="utf-8") as stream:
                payload = json.load(stream)
            self.assertEqual(payload["status"], "EXCLUDED")
            self.assertEqual(email_flow._list_root_msgs(str(mailbox)), [])
            create_mock.assert_not_called()
            xlsx_path_mock.assert_not_called()
            ensure_mock.assert_not_called()
            append_mock.assert_not_called()
            done_mock.assert_not_called()
            errors_mock.assert_not_called()
            drafts_mock.assert_not_called()

    def test_existing_exclusion_marker_makes_the_next_scan_empty(self):
        with tempfile.TemporaryDirectory() as tmp:
            mailbox = Path(tmp)
            msg_path = mailbox / "excluded-second-scan.msg"
            msg_path.touch()
            doc = _excluded_doc(msg_path)

            with (
                patch.object(
                    email_flow.mix_flow,
                    "create_one_document",
                ) as create_mock,
                patch.object(email_flow, "_xlsx_path") as xlsx_path_mock,
                patch.object(email_flow, "move_to_done") as done_mock,
                patch.object(email_flow, "move_to_errors") as errors_mock,
            ):
                result = email_flow._process_doc(
                    object(),
                    doc,
                    str(mailbox),
                    str(mailbox),
                    1,
                    1,
                    in_daemon=True,
                    process_mode="mix",
                    output_suffix="ГИСЖКХ",
                )

            marker = Path(str(msg_path) + ".asud_terminal.json")
            self.assertEqual(result, ("EXCLUDED", None, None))
            self.assertTrue(msg_path.exists())
            self.assertTrue(marker.exists())
            with marker.open(encoding="utf-8") as stream:
                payload = json.load(stream)
            self.assertEqual(payload["status"], "EXCLUDED")
            self.assertEqual(email_flow._list_root_msgs(str(mailbox)), [])
            create_mock.assert_not_called()
            xlsx_path_mock.assert_not_called()
            done_mock.assert_not_called()
            errors_mock.assert_not_called()

    def test_daemon_processes_excluded_message_once_without_retry_or_error(self):
        old_stop_flag = email_flow._stop_flag
        old_settings = email_flow.settings
        self.addCleanup(setattr, email_flow, "_stop_flag", old_stop_flag)
        self.addCleanup(setattr, email_flow, "settings", old_settings)

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            mailbox = root / "mailbox"
            mailbox.mkdir()
            msg_path = mailbox / "excluded-daemon.msg"
            msg_path.touch()
            (root / "msedgedriver.exe").touch()

            folders_json = json.dumps([
                {"dir": str(mailbox), "output_suffix": "ГИСЖКХ"},
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
            parsed_doc = _excluded_doc(msg_path)

            with (
                patch.dict(os.environ, {
                    "ASUD_EMAIL_FOLDERS_JSON": folders_json,
                    "ASUD_EMAIL_ROUND_ROBIN": "0",
                    "ASUD_EMAIL_PROCESS_MODE": "mix",
                    "ASUD_OUTPUT_SUFFIX": "",
                    # Normal successes may be deleted; exclusions must not be.
                    "ASUD_DELETE_AFTER_DONE": "1",
                }, clear=False),
                patch.object(email_flow.cfg, "load", return_value=daemon_settings),
                patch.object(email_flow.cfg, "setup_file_logger"),
                patch.object(email_flow.cfg, "keep_system_awake"),
                patch.object(email_flow.cfg, "get_base_dir", return_value=str(root)),
                patch.object(email_flow.cfg, "build_edge_options", return_value=object()),
                patch.object(
                    email_flow,
                    "_wait_for_folder",
                    side_effect=lambda folder: (True, folder),
                ),
                patch.object(email_flow.webdriver, "Edge", return_value=driver),
                patch.object(email_flow.signal, "signal"),
                patch.object(email_flow, "set_driver_timeout"),
                patch.object(email_flow, "wait_asud_loaded"),
                patch.object(
                    email_flow,
                    "_parse_one_msg",
                    return_value=parsed_doc,
                ) as parse_mock,
                patch.object(
                    email_flow.mix_flow,
                    "create_one_document",
                ) as create_mock,
                patch.object(email_flow, "move_to_errors") as errors_mock,
                patch.object(email_flow, "_print_doc_line") as line_mock,
                patch.object(
                    email_flow,
                    "_interruptible_sleep",
                    side_effect=stop_after_empty_followup_tick,
                ),
            ):
                email_flow.daemon_main()

            marker = Path(str(msg_path) + ".asud_terminal.json")
            self.assertTrue(msg_path.exists())
            self.assertTrue(marker.exists())
            self.assertEqual(email_flow._list_root_msgs(str(mailbox)), [])

        parse_mock.assert_called_once()
        create_mock.assert_not_called()
        errors_mock.assert_not_called()
        self.assertEqual(line_mock.call_args.args[2], "EXCLUDED")
        self.assertTrue(driver.quit_called)


if __name__ == "__main__":
    unittest.main()
