import unittest
from unittest.mock import patch

from flows import email as email_flow


class _DummyMessage:
    def close(self):
        pass


class FeedbackEmailFlowTests(unittest.TestCase):
    def _parse(self, body, *, subject="Новый вопрос с сайта Омск РТС",
               msg_path="feedback-without-fio.msg", settings=None):
        values = {
            "subject": subject,
            "body": body,
        }
        old_settings = email_flow.settings
        email_flow.settings = settings or {
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
                return email_flow._parse_one_msg(
                    msg_path, process_mode="mix")
        finally:
            email_flow.settings = old_settings

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
        self.assertIsNone(doc["addressees_override"])

    def test_quality_feedback_overrides_addressee_with_zhukov_only(self):
        body = """Адрес: Омск, ул. Тестовая, 1
Тема вопроса: Качество коммунальных услуг [84]
-------------------------------------------------------
Вопрос:
Жалоба на качество услуги
-------------------------------------------------------"""

        doc = self._parse(body)

        self.assertEqual(
            doc["addressees_override"], ["Жуков Иван Сергеевич"]
        )

    def test_other_feedback_keeps_global_addressee_route(self):
        body = """Адрес: Омск, ул. Тестовая, 1
Тема вопроса: Документы о поверке счётчика [82]
Вопрос:
Прошу внести сведения
-------------------------------------------------------"""

        doc = self._parse(body)

        self.assertIsNone(doc["addressees_override"])

    def test_oek_quality_request_overrides_addressee_with_zhukov_only(self):
        body = """От: noreply@omesc.ru
Тема: Перенаправление обращения абонента
Суть обращения: После ремонта подача ГВС отсутствует до настоящего времени.
Прошу произвести перерасчёт."""

        doc = self._parse(
            body,
            subject="FW: Перенаправление обращения абонента",
            msg_path=r"D:\OutlookSubjects\ОЭК\quality.msg",
        )

        self.assertEqual(
            doc["addressees_override"], ["Жуков Иван Сергеевич"]
        )
        self.assertEqual(
            doc["addressee_route_reason"], "ОЭК: отсутствует подача ГВС"
        )

    def test_oek_billing_request_keeps_global_addressee_route(self):
        body = """От: operator@omesc.ru
Суть обращения: Прошу проверить начисления и произвести перерасчёт."""

        doc = self._parse(
            body,
            subject="FW: Обращение абонента",
            msg_path=r"D:\OutlookSubjects\ОЭК\billing.msg",
        )

        self.assertIsNone(doc["addressees_override"])
        self.assertIsNone(doc["addressee_route_reason"])

    def test_service_only_body_falls_back_to_subject(self):
        body = "ВНИМАНИЕ! Письмо было отправлено внешним отправителем"

        doc = self._parse(
            body,
            subject="для работы и ответа потребителю",
            msg_path=r"D:\OutlookSubjects\ОЭК\service-only.msg",
        )

        self.assertEqual(doc["содержание"], "для работы и ответа потребителю")
        self.assertEqual(
            doc["корреспондент"], "Неизвестный Неизвестный Неизвестный"
        )
        self.assertFalse(doc["корр_найден"])


if __name__ == "__main__":
    unittest.main()
