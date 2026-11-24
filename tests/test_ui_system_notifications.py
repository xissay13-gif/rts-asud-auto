import unittest
from unittest.mock import patch

from selenium.webdriver.common.by import By

from shared import ui


class _CloseButton:
    def __init__(self, dialog):
        self.dialog = dialog

    def is_displayed(self):
        return True


class _Dialog:
    def __init__(self):
        self.visible = True
        self.close_button = _CloseButton(self)

    def is_displayed(self):
        return self.visible

    def find_elements(self, by, selector):
        if by == By.CSS_SELECTOR and selector == "[data-marker='dialog-close-btn']":
            return [self.close_button]
        return []


class _Driver:
    def __init__(self, dialog=None):
        self.dialog = dialog

    def find_elements(self, by, selector):
        if (by == By.CSS_SELECTOR and selector == "#notifications-dialog" and
                self.dialog is not None):
            return [self.dialog]
        return []

    def execute_script(self, _script, element):
        element.dialog.visible = False


class SystemNotificationTests(unittest.TestCase):
    def test_absent_notification_is_noop(self):
        self.assertTrue(ui.dismiss_system_notifications(_Driver()))

    def test_visible_notification_is_closed_and_verified(self):
        dialog = _Dialog()
        driver = _Driver(dialog)

        def close_dialog(_driver, element, _description):
            element.dialog.visible = False
            return True

        with patch.object(ui, "click", side_effect=close_dialog) as click_mock:
            result = ui.dismiss_system_notifications(driver)

        self.assertTrue(result)
        self.assertFalse(dialog.visible)
        click_mock.assert_called_once_with(
            driver, dialog.close_button, "Закрыть системное уведомление"
        )

    def test_reports_failure_when_overlay_stays_visible(self):
        dialog = _Dialog()
        driver = _Driver(dialog)

        with (
            patch.object(ui, "click", return_value=True),
            patch.object(ui.time, "monotonic", side_effect=[0.0, 0.0, 9.0, 9.0]),
            patch.object(ui, "WebDriverWait") as wait_mock,
        ):
            wait_mock.return_value.until.side_effect = TimeoutError
            result = ui.dismiss_system_notifications(driver, timeout=8)

        self.assertFalse(result)
        self.assertTrue(dialog.visible)


if __name__ == "__main__":
    unittest.main()
