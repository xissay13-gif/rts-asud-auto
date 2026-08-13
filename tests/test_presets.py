import unittest

from app import find_preset


class PresetTests(unittest.TestCase):
    def setUp(self):
        self.presets = [
            {
                "id": "sbis",
                "name": "SBIS incoming",
                "mode": "sbis",
                "folder": r"D:\SBIS\Incoming",
            },
            {
                "name": "Email flow",
                "mode": "email",
                "folder": r"D:\Mail",
            },
        ]

    def test_find_by_stable_id_case_insensitive(self):
        preset = find_preset(self.presets, "SBIS")
        self.assertEqual(preset["folder"], r"D:\SBIS\Incoming")

    def test_find_by_exact_display_name(self):
        preset = find_preset(self.presets, "email flow")
        self.assertEqual(preset["mode"], "email")

    def test_unknown_or_invalid_preset_returns_none(self):
        self.assertIsNone(find_preset(self.presets, "missing"))
        self.assertIsNone(find_preset([None, "bad"], "sbis"))


if __name__ == "__main__":
    unittest.main()
