import unittest

from shared.correspondent import (
    correspondent_card_parts,
    match_legal_correspondent,
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

    def test_address_card_puts_full_address_in_surname_only(self):
        self.assertEqual(
            correspondent_card_parts(
                "644021, г. Омск, ул. 4-я Транспортная 15 кв 8",
                "address",
            ),
            ("644021, г. Омск, ул. 4-я Транспортная 15 кв 8", "", ""),
        )

    def test_legal_match_requires_full_name(self):
        self.assertTrue(match_legal_correspondent(
            "Тепловая Компания - -", "Тепловая Компания"))
        self.assertFalse(match_legal_correspondent(
            "Тепловая Сеть - -", "Тепловая Компания"))


if __name__ == "__main__":
    unittest.main()
