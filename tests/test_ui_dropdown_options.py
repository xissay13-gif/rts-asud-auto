import unittest

from shared.ui import DropdownOptions, find_dropdown_options


class _Driver:
    def __init__(self, state):
        self.state = state
        self.calls = []

    def execute_script(self, script, *args):
        self.calls.append((script, args))
        return self.state


class DropdownOptionsTests(unittest.TestCase):
    def test_explicit_empty_popup_is_distinct_from_unknown_popup(self):
        driver = _Driver({"popup_seen": True, "options": []})

        options = find_dropdown_options(driver, "Иванов", object())

        self.assertIsInstance(options, DropdownOptions)
        self.assertTrue(options.popup_seen)
        self.assertEqual(options, [])

    def test_missing_popup_is_not_treated_as_empty_lookup(self):
        driver = _Driver({"popup_seen": False, "options": []})

        options = find_dropdown_options(driver, "Иванов", object())

        self.assertFalse(options.popup_seen)
        self.assertEqual(options, [])

    def test_scoped_options_are_preserved(self):
        row = object()
        driver = _Driver({"popup_seen": True, "options": [row]})

        options = find_dropdown_options(driver, "Иванов", object())

        self.assertTrue(options.scoped)
        self.assertEqual(options, [row])


if __name__ == "__main__":
    unittest.main()
