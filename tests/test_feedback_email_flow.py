import unittest
from unittest.mock import patch

from flows import email as email_flow


class _DummyMessage:
    def close(self):
        pass


class FeedbackEmailFlowTests(unittest.TestCase):
    def test_feedback_without_fio_registers_address_as_correspondent(self):
        body = """Адрес электронной почты: a.mannik@yandex.ru
Адрес: 644021, г. Омск, ул. 4-я Транспортная 15 кв 8
Тема вопроса: Документы о поверке счётчика [82]
-------------------------------------------------------
Вопрос:
Прошу внести данные о прохождении поверки счетчика ГВС
-------------------------------------------------------"""
        values = {
            "subject": "Новый вопрос с сайта Омск РТС",
            "body": body,
        }
        old_settings = email_flow.settings
        email_flow.settings = {
            "unknown_correspondent": "Неизвестный Неизвестный Неизвестный",
            "default_type_idx": 8,
        }
        try:
            with (
                patch.object(email_flow.extract_msg, "openMsg",
                             return_value=_DummyMessage()),
                patch.object(email_flow, "_safe_field",
                             side_effect=lambda _msg, attr, *_: values[attr]),
                patch.object(email_flow, "_msg_link", return_value="feedback-1"),
                patch.object(email_flow, "_msg_date_prefix",
                             return_value="2026-08-06"),
                patch.object(email_flow, "parse_zhkh_body", return_value=None),
                patch.object(email_flow, "okrug_from_textbody",
                             return_value=None),
            ):
                doc = email_flow._parse_one_msg(
                    "feedback-without-fio.msg", process_mode="mix")
        finally:
            email_flow.settings = old_settings

        address = "644021, г. Омск, ул. 4-я Транспортная 15 кв 8"
        self.assertEqual(doc["корреспондент"], address)
        self.assertEqual(doc["корреспондент_тип"], "address")
        self.assertTrue(doc["корр_найден"])
        self.assertEqual(doc["тип_индекс"], 8)
        self.assertEqual(
            doc["содержание"],
            f"Новый вопрос с сайта Омск РТС — {address}",
        )


if __name__ == "__main__":
    unittest.main()
