import unittest

from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
)
from selenium.webdriver.common.by import By

from shared.ui import wait_and_click


class _Element:
    def __init__(self, click_error=None):
        self.click_error = click_error
        self.click_calls = 0

    def is_displayed(self):
        return True

    def is_enabled(self):
        return True

    def click(self):
        self.click_calls += 1
        if self.click_error is not None:
            raise self.click_error


class _Driver:
    def __init__(self, elements):
        self.elements = list(elements)
        self.find_calls = []
        self.script_calls = []

    def find_element(self, by, selector):
        self.find_calls.append((by, selector))
        return self.elements.pop(0)

    def execute_script(self, script, element):
        self.script_calls.append((script, element))


class WaitAndClickTests(unittest.TestCase):
    def test_refinds_element_after_stale_native_click(self):
        stale = _Element(StaleElementReferenceException("rerendered"))
        fresh = _Element()
        driver = _Driver([stale, fresh])

        clicked = wait_and_click(
            driver, By.ID, "incoming-document", timeout=1
        )

        self.assertIs(clicked, fresh)
        self.assertEqual(stale.click_calls, 1)
        self.assertEqual(fresh.click_calls, 1)
        self.assertEqual(
            driver.find_calls,
            [(By.ID, "incoming-document"), (By.ID, "incoming-document")],
        )
        self.assertEqual(driver.script_calls, [])

    def test_normal_click_uses_single_lookup(self):
        element = _Element()
        driver = _Driver([element])

        clicked = wait_and_click(driver, By.CSS_SELECTOR, ".create", timeout=1)

        self.assertIs(clicked, element)
        self.assertEqual(element.click_calls, 1)
        self.assertEqual(driver.find_calls, [(By.CSS_SELECTOR, ".create")])
        self.assertEqual(driver.script_calls, [])

    def test_js_fallback_uses_freshly_located_element(self):
        intercepted = _Element(ElementClickInterceptedException("overlay"))
        fresh = _Element()
        driver = _Driver([intercepted, fresh])

        clicked = wait_and_click(driver, By.ID, "register", timeout=1)

        self.assertIs(clicked, fresh)
        self.assertEqual(intercepted.click_calls, 1)
        self.assertEqual(fresh.click_calls, 0)
        self.assertEqual(len(driver.find_calls), 2)
        self.assertEqual(len(driver.script_calls), 1)
        self.assertIs(driver.script_calls[0][1], fresh)


if __name__ == "__main__":
    unittest.main()
