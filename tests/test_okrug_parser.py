import tempfile
import unittest
from pathlib import Path

from shared import okrug_parser


class OkrugParserTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        Path(self.tmp.name, "addresses.csv").write_text(
            "street;house;okrug\n"
            "3 разъезд;27;САО\n"
            "пгт биофабрика;20;ОАО\n"
            "70 лет октября;18;КАО\n"
            "однозначная;1;ЦАО\n"
            "однозначная;3;ЦАО\n"
            "пограничная;1;ЦАО\n"
            "пограничная;2;ЛАО\n"
            "омская;1;ЦАО\n"
            "мира;1;САО\n",
            encoding="utf-8",
        )
        okrug_parser._index = None
        okrug_parser._streets_sorted = None

    def tearDown(self):
        okrug_parser._index = None
        okrug_parser._streets_sorted = None
        self.tmp.cleanup()

    def parse(self, body="", subject=None, reason=False):
        return okrug_parser.okrug_from_textbody(
            body,
            subject=subject,
            base_dir_fn=lambda: self.tmp.name,
            return_reason=reason,
        )

    def test_uses_address_from_subject(self):
        self.assertEqual(self.parse("Без адреса", "70 лет Октября 18"), "КАО")

    def test_normalizes_masculine_numbered_street(self):
        self.assertEqual(self.parse("ул. 3-й Разъезд, д. 27А"), "САО")

    def test_normalizes_settlement_alias(self):
        self.assertEqual(self.parse("п. Биофабрика, д. 20"), "ОАО")

    def test_falls_back_only_for_single_district_street(self):
        district, reason = self.parse("ул. Однозначная, д. 99", reason=True)
        self.assertEqual(district, "ЦАО")
        self.assertIn("однозначной улице", reason)
        self.assertIsNone(self.parse("ул. Пограничная, д. 99"))

    def test_region_name_is_not_treated_as_omskaya_street(self):
        self.assertIsNone(self.parse(
            "Омская обл., г. Омск, ул. Владимира Жоги, д. 1"
        ))
        self.assertIsNone(self.parse(
            "Интернет-приёмная ООО Омская энергосбытовая компания, "
            "обращение 5423832"
        ))
        self.assertIsNone(self.parse(
            "Желаю мира и сообщаю номер обращения 5423832"
        ))


if __name__ == "__main__":
    unittest.main()
