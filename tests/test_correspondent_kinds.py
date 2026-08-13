import unittest

from shared.correspondent import (
    correspondent_card_parts,
    match_correspondent,
    match_legal_correspondent,
    match_strict,
)


class CorrespondentKindTests(unittest.TestCase):
    def test_person_card_uses_full_fio_parts(self):
        self.assertEqual(
            correspondent_card_parts("Иванов Иван Иванович", "person"),
            ("Иванов", "Иван", "Иванович"),
        )

    def test_legal_card_puts_full_name_in_surname_and_dashes(self):
        self.assertEqual(
            correspondent_card_parts("Тепловая Компания", "legal"),
            ("Тепловая Компания", "-", "-"),
        )

    def test_feedback_address_card_puts_address_in_surname_and_dashes(self):
        address = "644021, г. Омск, ул. 4-я Транспортная 15 кв 8"

        for kind in ("address", "feedback-address", "feedback_address"):
            with self.subTest(kind=kind):
                self.assertEqual(
                    correspondent_card_parts(address, kind),
                    (address, "-", "-"),
                )

    def test_legal_match_requires_full_name(self):
        self.assertTrue(match_legal_correspondent(
            "Тепловая Компания - -", "Тепловая Компания"))
        self.assertFalse(match_legal_correspondent(
            "Тепловая Сеть - -", "Тепловая Компания"))

    def test_person_match_requires_exact_surname_token(self):
        expected = "Иванов Иван Иванович"

        self.assertTrue(match_strict("Иванов И И", expected))
        self.assertTrue(match_strict(
            "Иванов И И Иванов Иван Иванович / ФЛ", expected))
        self.assertFalse(match_strict("Сиванов И И", expected))
        self.assertFalse(match_strict("Переиванов И И", expected))

        # Live regression: ранее нормализованная строка искалась как
        # подстрока и «Гудов А А» ошибочно совпадал с «Огудов А А».
        self.assertFalse(match_strict(
            "Огудов А А", "Гудов Александр Анатольевич"))

    def test_addressee_soft_match_does_not_use_surname_substring(self):
        expected = "Иванов Иван Иванович"

        self.assertTrue(match_correspondent("Иванов", expected))
        self.assertFalse(match_correspondent("Сиванов И И", expected))


if __name__ == "__main__":
    unittest.main()
