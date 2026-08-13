import unittest

from shared.feedback_parser import parse_feedback_body


class FeedbackParserTests(unittest.TestCase):
    def test_new_template_without_fio_keeps_address(self):
        body = """Адрес электронной почты: a.mannik@yandex.ru<mailto:a.mannik@yandex.ru>
Адрес: 644021, г. Омск, ул. 4-я Транспортная 15 кв 8
Тема вопроса: Документы о поверке счётчика [82]
-------------------------------------------------------
Вопрос:
Прошу внести данные о прохождении поверки счетчика ГВС
-------------------------------------------------------
Письмо сгенерировано автоматически."""

        parsed = parse_feedback_body(
            body, subject="Новый вопрос с сайта Омск РТС")

        self.assertIsNotNone(parsed)
        self.assertIsNone(parsed["фио"])
        self.assertEqual(
            parsed["адрес"],
            "644021, г. Омск, ул. 4-я Транспортная 15 кв 8",
        )
        self.assertEqual(parsed["email"], "a.mannik@yandex.ru")
        self.assertEqual(
            parsed["корреспондент"],
            "644021, г. Омск, ул. 4-я Транспортная 15 кв 8",
        )
        self.assertEqual(parsed["корреспондент_тип"], "address")
        self.assertEqual(
            parsed["вопрос"],
            "Прошу внести данные о прохождении поверки счетчика ГВС",
        )

    def test_old_template_with_fio_is_unchanged(self):
        body = """Фамилия, имя, отчество: Токарева Лариса Михайловна
Адрес: Комкова 8/1 кв.166
Тема вопроса: Перерасчёт [81]
Вопрос:
Прошу выполнить перерасчёт
-------------------------------------------------------"""

        parsed = parse_feedback_body(body)

        self.assertEqual(parsed["фио"], "Токарева Лариса Михайловна")
        self.assertEqual(parsed["фамилия"], "Токарева")
        self.assertEqual(parsed["адрес"], "Комкова 8/1 кв.166")
        self.assertEqual(
            parsed["корреспондент"], "Токарева Лариса Михайловна")
        self.assertEqual(parsed["корреспондент_тип"], "person")

    def test_unrelated_message_is_not_feedback(self):
        self.assertIsNone(parse_feedback_body(
            "Адрес: Омск, улица Ленина, 1",
            subject="Обычное письмо",
        ))


if __name__ == "__main__":
    unittest.main()
