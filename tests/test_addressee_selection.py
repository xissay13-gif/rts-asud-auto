import unittest
from unittest.mock import patch

from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
)
from selenium.webdriver.common.by import By

from flows import auto_create, mix, smart
from shared import addressee
from shared.ui import DropdownOptions


FULL_NAME = "Иванов Иван Иванович"
INITIALS = "Иванов И.И."


class _TextElement:
    def __init__(self, text, *, css_class="", tag_name="div"):
        self.text = text
        self.tag_name = tag_name
        self._css_class = css_class

    def is_displayed(self):
        return True

    def is_enabled(self):
        return True

    def get_attribute(self, name):
        if name == "class":
            return self._css_class
        if name == "readonly":
            return None
        return None

    def click(self):
        return None

    def send_keys(self, *_keys):
        return None


class _Grid:
    """Small Selenium-like view of the real #addressee_grid_id."""

    def __init__(self, driver):
        self.driver = driver

    @property
    def text(self):
        snapshot = self.driver.next_snapshot()
        return "\n".join(snapshot)

    def find_elements(self, by, selector):
        if by not in (By.CSS_SELECTOR, By.XPATH):
            return []
        snapshot = self.driver.next_snapshot()
        return [
            _TextElement(
                value,
                css_class=(
                    "GxtGridStyle-row addressees-row"
                    if "row" in selector.casefold() or "tr" in selector.casefold()
                    else "contentTd"
                ),
            )
            for value in snapshot
        ]


class _GridDriver:
    """Supports either find_elements or one-shot JS DOM readers."""

    def __init__(self, snapshots, *, popup_texts=()):
        self.snapshots = list(snapshots)
        self.popup_texts = list(popup_texts)
        self.snapshot_calls = 0
        self.grid = _Grid(self)

    def next_snapshot(self):
        index = min(self.snapshot_calls, max(0, len(self.snapshots) - 1))
        self.snapshot_calls += 1
        value = self.snapshots[index] if self.snapshots else []
        if isinstance(value, BaseException):
            raise value
        return list(value)

    def find_element(self, by, selector):
        if by == By.ID and selector == "addressee_grid_id":
            return self.grid
        raise NoSuchElementException(selector)

    def find_elements(self, _by, _selector):
        # Deliberately expose matching popup text outside the committed grid.
        # A verifier that scans the whole document would fail this contract.
        return [_TextElement(value, css_class="popup-item")
                for value in self.popup_texts]

    def execute_script(self, script, *args):
        if "scrollIntoView" in script:
            return None
        if "parentElement" in script and args:
            return args[0]
        # The shared verifier may read the grid in one browser-side pass.
        return self.next_snapshot()


class _ImmediateWait:
    def __init__(self, driver, _timeout):
        self.driver = driver

    def until(self, predicate):
        value = predicate(self.driver)
        if not value:
            raise RuntimeError("predicate is false")
        return value


class _BlindWait:
    """Form-flow wait double; Selenium predicates are outside this contract."""

    def __init__(self, _driver, _timeout):
        pass

    def until(self, _predicate):
        return _TextElement("form-control")


class AddresseeCommittedTests(unittest.TestCase):
    def test_dom_reader_is_scoped_to_semantic_committed_row_markers(self):
        script = addressee._READ_COMMITTED_ADDRESSEES_JS

        self.assertIn("#addressee_grid_id", script)
        self.assertIn("[data-attr='agents-content']", script)
        self.assertIn("img[data-attr='remove-trigger']", script)
        self.assertNotIn("document.body.innerText", script)
        self.assertNotIn("chip", script.casefold())

    def test_empty_grid_stays_false_even_when_popup_contains_target(self):
        driver = _GridDriver([[]], popup_texts=[FULL_NAME])

        self.assertFalse(addressee.addressee_committed(driver, FULL_NAME))

    def test_exact_full_name_or_initials_in_committed_row_is_true(self):
        for display in (
            FULL_NAME,
            f"{INITIALS} Организация / Подразделение / Должность",
        ):
            with self.subTest(display_form=display[:12]):
                driver = _GridDriver([[display]])
                self.assertTrue(
                    addressee.addressee_committed(driver, FULL_NAME)
                )

    def test_unrelated_committed_row_is_false(self):
        driver = _GridDriver([["Иваненко И.И. Организация / Должность"]])

        self.assertFalse(addressee.addressee_committed(driver, FULL_NAME))

    def test_wait_survives_stale_rerender_and_delayed_committed_row(self):
        driver = _GridDriver([
            StaleElementReferenceException("GXT replaced the grid"),
            [],
            [f"{INITIALS} Организация / Должность"],
        ])

        with patch.object(addressee.time, "sleep", return_value=None):
            selected = addressee.wait_addressee_committed(
                driver,
                FULL_NAME,
                timeout=1.0,
                poll_interval=0.01,
            )

        self.assertTrue(selected)
        self.assertGreaterEqual(driver.snapshot_calls, 3)


class AddresseeSelectionFlowTests(unittest.TestCase):
    def test_already_committed_addressee_is_idempotent(self):
        driver = _GridDriver([[f"{INITIALS} Организация / Должность"]])

        with patch.object(addressee, "cdp_click") as cdp_mock:
            selected = addressee.add_addressee(driver, FULL_NAME)

        self.assertTrue(selected)
        cdp_mock.assert_not_called()

    def test_exact_option_becomes_confirmed_after_grid_rerender(self):
        driver = _GridDriver([
            [],
            [f"{INITIALS} Организация / Должность"],
        ])
        input_element = _TextElement("", tag_name="input")
        option = _TextElement(
            INITIALS,
            css_class="Css3ListViewAppearance-Css3ListViewStyle-item",
        )
        options = DropdownOptions(
            [option], popup_seen=True,
            input_value="Иванов", input_observed=True,
        )

        with (
            patch.object(
                addressee, "find_input_near_label",
                return_value=input_element,
            ),
            patch.object(addressee, "js_type_combobox"),
            patch.object(
                addressee, "find_dropdown_options", return_value=options
            ),
            patch.object(addressee, "WebDriverWait", _ImmediateWait),
            patch.object(addressee, "cdp_click", return_value=True),
            patch.object(addressee.time, "sleep", return_value=None),
        ):
            selected = addressee.add_addressee(driver, FULL_NAME)

        self.assertTrue(selected)

    def test_cdp_dispatch_without_committed_row_returns_false(self):
        driver = _GridDriver([[], [], []])
        input_element = _TextElement("", tag_name="input")
        option = _TextElement(
            FULL_NAME,
            css_class="Css3ListViewAppearance-Css3ListViewStyle-item",
        )
        options = DropdownOptions(
            [option], popup_seen=True,
            input_value="Иванов", input_observed=True,
        )

        with (
            patch.object(
                addressee, "find_input_near_label",
                return_value=input_element,
            ),
            patch.object(addressee, "js_type_combobox"),
            patch.object(
                addressee, "find_dropdown_options", return_value=options
            ),
            patch.object(addressee, "WebDriverWait", _ImmediateWait),
            patch.object(addressee, "cdp_click", return_value=True) as cdp_mock,
            patch.object(addressee.time, "sleep", return_value=None),
        ):
            selected = addressee.add_addressee(driver, FULL_NAME)

        self.assertFalse(selected)
        cdp_mock.assert_called_once()

    def test_mix_aborts_before_save_when_addressee_is_not_confirmed(self):
        driver = object()
        doc = {
            "тема": "Тест",
            "корреспондент": "Тестовый корреспондент",
            "корреспондент_тип": "person",
            "корр_источник": "test",
            "корр_найден": True,
            "тип_название": "Письма, заявления и жалобы граждан, акционеров",
            "содержание": "Тестовое обращение",
        }

        with (
            patch.object(mix, "settings", {"addressees": [FULL_NAME]}),
            patch.object(mix, "WebDriverWait", _BlindWait),
            patch.object(mix, "wait_and_click", return_value=True),
            patch.object(mix, "click", return_value=True) as click_mock,
            patch.object(mix, "fill_text"),
            patch.object(mix, "fill_correspondent_field", return_value=True),
            patch.object(mix, "fill_corr_number"),
            patch.object(mix, "fill_corr_date"),
            patch.object(mix, "add_addressee", return_value=False),
            patch.object(mix, "fill_delivery_method") as delivery_mock,
            patch.object(mix, "find_msg_by_link") as attachment_lookup,
            patch.object(mix, "attach_content") as attach_mock,
            patch.object(mix, "register_and_resolve") as register_mock,
            patch.object(mix, "close_open_modals"),
            patch.object(mix, "close_card_and_wait_main"),
        ):
            with self.assertRaisesRegex(RuntimeError, "Адресат"):
                mix.create_one_document(driver, doc, index=1, total=1)

        self.assertNotIn(
            "Сохранить", [call.args[2] for call in click_mock.call_args_list]
        )
        delivery_mock.assert_not_called()
        attachment_lookup.assert_not_called()
        attach_mock.assert_not_called()
        register_mock.assert_not_called()
        self.assertEqual(mix._last_result["status"], "FAILED")

    def test_auto_create_aborts_before_save_attach_and_register(self):
        driver = object()
        doc = {
            "содержание": "Тестовое обращение",
            "корреспондент": "Тестовый корреспондент",
            "тема_индекс": 8,
            "файл": "test.msg",
        }

        with (
            patch.object(
                auto_create, "settings", {"addressees": [FULL_NAME]}
            ),
            patch.object(auto_create, "WebDriverWait", _BlindWait),
            patch.object(auto_create, "wait_and_click", return_value=True),
            patch.object(
                auto_create, "click", return_value=True
            ) as click_mock,
            patch.object(auto_create, "fill_text"),
            patch.object(
                auto_create, "fill_correspondent_field", return_value=True
            ),
            patch.object(auto_create, "fill_corr_number"),
            patch.object(auto_create, "fill_corr_date"),
            patch.object(auto_create, "add_addressee", return_value=False),
            patch.object(
                auto_create, "fill_delivery_method"
            ) as delivery_mock,
            patch.object(auto_create, "attach_content") as attach_mock,
            patch.object(
                auto_create, "register_and_resolve"
            ) as register_mock,
            patch.object(auto_create, "close_open_modals"),
        ):
            with self.assertRaisesRegex(RuntimeError, "Адресат"):
                auto_create.create_one_document(
                    driver, doc, index=1, total=1
                )

        self.assertNotIn(
            "Сохранить", [call.args[2] for call in click_mock.call_args_list]
        )
        delivery_mock.assert_not_called()
        attach_mock.assert_not_called()
        register_mock.assert_not_called()

    def test_smart_aborts_before_save_and_attachment_lookup(self):
        driver = object()
        doc = {
            "тема": "Тест",
            "корреспондент": "Тестовый корреспондент",
            "тип_название": "Письма, заявления и жалобы граждан, акционеров",
            "содержание": "Тестовое обращение",
            "link": "test-link",
            "файл": "test.msg",
        }

        with (
            patch.object(smart, "settings", {"addressee": FULL_NAME}),
            patch.object(smart, "WebDriverWait", _BlindWait),
            patch.object(smart, "wait_and_click", return_value=True),
            patch.object(smart, "click", return_value=True) as click_mock,
            patch.object(smart, "fill_text"),
            patch.object(
                smart, "fill_correspondent_field", return_value=True
            ),
            patch.object(smart, "fill_corr_number"),
            patch.object(smart, "fill_corr_date"),
            patch.object(smart, "add_addressee", return_value=False),
            patch.object(smart, "fill_delivery_method") as delivery_mock,
            patch.object(smart, "find_msg_by_link") as attachment_lookup,
            patch.object(smart, "attach_content") as attach_mock,
            patch.object(smart, "move_to_done") as move_mock,
            patch.object(smart, "close_open_modals"),
        ):
            with self.assertRaisesRegex(RuntimeError, "Адресат"):
                smart.create_one_document(driver, doc, index=1, total=1)

        self.assertNotIn(
            "Сохранить", [call.args[2] for call in click_mock.call_args_list]
        )
        delivery_mock.assert_not_called()
        attachment_lookup.assert_not_called()
        attach_mock.assert_not_called()
        move_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
