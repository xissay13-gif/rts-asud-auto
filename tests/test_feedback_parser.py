import unittest

from shared.feedback_parser import (
    detect_oek_quality_issue,
    parse_feedback_body,
    match_feedback_addressees,
    normalize_feedback_topic,
)


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

    def test_quality_category_routes_to_zhukov_with_optional_code(self):
        expected = ["Жуков Иван Сергеевич"]
        self.assertEqual(
            match_feedback_addressees("Качество коммунальных услуг [80]"),
            expected,
        )
        self.assertEqual(
            match_feedback_addressees("  КАЧЕСТВО  коммунальных услуг.  "),
            expected,
        )

    def test_feedback_route_is_exact_not_free_text_contains(self):
        self.assertIsNone(match_feedback_addressees(
            "Качество коммунальных услуг населению"
        ))
        self.assertIsNone(match_feedback_addressees(
            "Некачественная поставка ресурса"
        ))

    def test_topic_normalization_removes_only_trailing_numeric_code(self):
        self.assertEqual(
            normalize_feedback_topic("Качество коммунальных услуг [84]"),
            "качество коммунальных услуг",
        )

    def test_oek_absent_hot_water_is_quality_even_with_recalculation(self):
        body = """From: noreply@omesc.ru
Суть обращения: После планового отключения подача ГВС отсутствует.
Прошу произвести перерасчёт в связи с непредоставлением ГВС.
-----Original Message-----
Старый текст цепочки."""

        reason = detect_oek_quality_issue(body, subject="Перенаправление")

        self.assertEqual(reason, "отсутствует подача ГВС")

    def test_oek_weak_pressure_is_quality(self):
        body = """Суть обращения:
В квартире очень слабый напор горчей воды уже несколько дней."""

        reason = detect_oek_quality_issue(
            body,
            subject="Обращение абонента",
            source_hint=r"D:\OutlookSubjects\ОЭК\request.msg",
        )

        self.assertEqual(reason, "слабый напор горячей воды")

    def test_oek_water_still_missing_is_quality(self):
        body = """От: operator@omesc.ru
Суть обращения: После ремонта воды нет до сих пор."""

        self.assertEqual(
            detect_oek_quality_issue(body),
            "вода отсутствует до сих пор",
        )

    def test_oek_explicit_absence_of_hot_water_is_quality(self):
        cases = (
            "В доме нет горячей воды третий день.",
            "Горячая вода отсутствует с утра.",
            "Горячей воды нет уже неделю.",
            "ГВС не восстановлено после ремонта.",
        )

        for request in cases:
            with self.subTest(request=request):
                body = f"""От: noreply@omesc.ru
Суть обращения: {request}"""
                self.assertEqual(
                    detect_oek_quality_issue(body),
                    "нет горячей воды",
                )

    def test_oek_billing_and_metering_absence_is_not_hot_water_absence(self):
        cases = (
            (
                "Была проведена поверка счётчика горячей воды. "
                "В квитанции НЕТ СТРОКИ с показаниями счётчика ГВС."
            ),
            (
                "В графе «Горячая вода» последние показания отсутствовали, "
                "хотя ранее они были переданы."
            ),
            (
                "Подать данные приборов учёта горячей воды нет возможности, "
                "так как начальные показания не внесены."
            ),
            "Почему нет перерасчёта в квитанции по горячей воде?",
        )

        for request in cases:
            with self.subTest(request=request):
                body = f"""От: noreply@omesc.ru
Суть обращения: {request}"""
                self.assertIsNone(detect_oek_quality_issue(body))

    def test_oek_plain_recalculation_does_not_trigger_quality_route(self):
        body = """От: operator@omesc.ru
Суть обращения: Прошу проверить начисления и произвести перерасчёт платы."""

        self.assertIsNone(detect_oek_quality_issue(body))

    def test_quality_words_outside_request_block_are_ignored(self):
        body = """От: operator@omesc.ru
Суть обращения: Прошу предоставить копию квитанции.
From: old@example.org
В старой переписке: нет горячей воды."""

        self.assertIsNone(detect_oek_quality_issue(body))

    def test_non_oek_letter_does_not_trigger_from_phrase_alone(self):
        body = """Суть обращения: Подача ГВС отсутствует второй день."""

        self.assertIsNone(detect_oek_quality_issue(body))


if __name__ == "__main__":
    unittest.main()
