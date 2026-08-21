import os
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import app
from shared import config as config_module


class ApiPresetTests(unittest.TestCase):
    @staticmethod
    def _preset(enabled=False):
        return {
            "id": "gis-api-test",
            "name": "GIS API test",
            "mode": "mix",
            "folder": r"D:\OutlookSubjects\GIS",
            "registration_backend": "asud_api",
            "asud_api_enabled": enabled,
        }

    def _run(self, preset, initial_env=None):
        email_main = Mock()
        email_module = types.ModuleType("flows.email")
        email_module.main = email_main

        with patch.object(
                sys, "argv", ["app.py", "--preset=gis-api-test"]), \
             patch.object(app.cfg, "get_base_dir", return_value=os.getcwd()), \
             patch.object(
                 app.cfg, "load", return_value={"presets": [preset]}), \
             patch.dict(sys.modules, {"flows.email": email_module}), \
             patch.dict(os.environ, initial_env or {}, clear=True):
            app.main()
            result = {
                "backend": os.environ.get("ASUD_EMAIL_REGISTRATION_BACKEND"),
                "enabled": os.environ.get("ASUD_API_ENABLED"),
            }

        email_main.assert_called_once_with()
        return result

    def test_api_preset_exports_safe_disabled_backend(self):
        result = self._run(self._preset(enabled=False))

        self.assertEqual(result["backend"], "asud_api")
        self.assertEqual(result["enabled"], "0")

    def test_explicit_bat_environment_wins_over_preset_default(self):
        result = self._run(
            self._preset(enabled=False),
            {"ASUD_API_ENABLED": "1"},
        )

        self.assertEqual(result["backend"], "asud_api")
        self.assertEqual(result["enabled"], "1")

    def test_api_preset_overrides_ambient_selenium_backend(self):
        result = self._run(
            self._preset(enabled=False),
            {
                "ASUD_EMAIL_REGISTRATION_BACKEND": "selenium",
                "ASUD_API_ENABLED": "1",
            },
        )

        self.assertEqual(result["backend"], "asud_api")

    def test_default_api_configuration_is_fail_closed_and_secret_free(self):
        api = config_module.DEFAULTS["asud_api"]

        self.assertFalse(api["enabled"])
        self.assertEqual(api["mode"], "dry-run")
        self.assertFalse(api["allow_mutations"])
        self.assertTrue(all(not value for value in api["endpoints"].values()))
        for key in ("base_url", "lis", "user", "branch_id",
                    "incoming_type_path", "addressee_id"):
            self.assertEqual(api[key], "")
        self.assertEqual(api["attachment"], {
            "confirm_msg_supported": False,
            "max_bytes": 0,
        })
        self.assertFalse({
            "password", "token", "auth", "auth_header_value"
        }.intersection(api))

    def test_live_confirmation_cannot_be_inherited_from_environment(self):
        bat = Path("asud_gis_api_test.bat").read_text(encoding="utf-8")
        prompt = 'set /p "ASUD_API_CONFIRM=Для запуска введите LIVE-ONE: "'

        self.assertIn('set "ASUD_API_CONFIRM="\n' + prompt, bat)


if __name__ == "__main__":
    unittest.main()
