import unittest
from unittest.mock import Mock, patch

from flows import zhkh_daemon


class ZhkhDaemonRecoveryTests(unittest.TestCase):
    def test_recovery_replaces_only_current_driver(self):
        old_driver = Mock()
        new_driver = Mock()

        with (
            patch.object(zhkh_daemon, "_start_driver", return_value=new_driver),
            patch.object(zhkh_daemon, "wait_asud_loaded"),
            patch.object(zhkh_daemon, "switch_account", return_value=True),
            patch.object(zhkh_daemon, "click_sidebar_section", return_value=True),
        ):
            result = zhkh_daemon._recover_driver(
                old_driver,
                "C:/app",
                300,
                "https://asud.example/",
                "Басманов",
                "На резолюцию",
            )

        self.assertIs(result, new_driver)
        old_driver.quit.assert_called_once_with()
        new_driver.get.assert_called_once_with("https://asud.example/")
        new_driver.quit.assert_not_called()

    def test_failed_recovery_closes_new_driver(self):
        old_driver = Mock()
        new_driver = Mock()

        with (
            patch.object(zhkh_daemon, "_start_driver", return_value=new_driver),
            patch.object(zhkh_daemon, "wait_asud_loaded"),
            patch.object(zhkh_daemon, "switch_account", return_value=False),
        ):
            result = zhkh_daemon._recover_driver(
                old_driver,
                "C:/app",
                300,
                "https://asud.example/",
                "Басманов",
                "На резолюцию",
            )

        self.assertIsNone(result)
        new_driver.quit.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
