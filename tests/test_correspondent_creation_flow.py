import unittest
from unittest.mock import patch

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from flows import mix
from shared import correspondent, ui


ADDRESS = "644021, г. Омск, ул. 4-я Транспортная, д. 15, кв. 8"


class _Element:
    def __init__(self, name):
        self.name = name

    def is_displayed(self):
        return True


class _ScopedElement(_Element):
    """Visible node that may belong to the currently open dialog."""

    def __init__(self, name, dialog=None):
        super().__init__(name)
        self.dialog = dialog

    def find_elements(self, by, selector):
        if by == By.XPATH and "ancestor::" in selector:
            return [self.dialog] if self.dialog is not None else []
        return []


class _ScopeDriver:
    def __init__(self, candidates):
        self.candidates = candidates

    def find_elements(self, by, _selector):
        if by == By.XPATH:
            return self.candidates
        return []


class _CoveredDriver:
    def __init__(self, candidate):
        self.candidate = candidate

    def find_elements(self, _by, _selector):
        return [self.candidate]

    def find_element(self, _by, _selector):
        return self.candidate

    def execute_script(self, _script, _element):
        return {"known": True, "exposed": False, "score": 0}


class _ImmediateWait:
    """Browserless replacement for Selenium waits in the happy-path unit test."""

    def __init__(self, _driver, _timeout):
        pass

    def until(self, _predicate):
        return _Element("wait-result")


class _PredicateWait:
    def __init__(self, driver, _timeout):
        self.driver = driver

    def until(self, predicate):
        result = predicate(self.driver)
        if not result:
            raise TimeoutException("predicate is false")
        return result


class _CreationDriver:
    """Records card writes while exposing only the DOM hooks the flow needs."""

    def __init__(self):
        self.field_writes = []
        self.create_org = _Element("create-organization")
        self.add_person = _Element("add-person")
        self.save_person = _Element("save-person")
        self.select_person = _Element("select-person")
        self.done = _Element("done")

    def execute_script(self, script, *args):
        if "var el = arguments[0]; var value = arguments[1]" in script:
            element, value = args
            input_id = element.name
            self.field_writes.append((input_id, value))
            return f"ok:{input_id}"
        if "var heading = arguments[0], surname = arguments[1]" in script:
            return "ok"
        if "style.pointerEvents" in script:
            return True
        raise AssertionError("Unexpected JavaScript in creation flow")

    def find_element(self, by, selector):
        if by == By.CSS_SELECTOR:
            if "create_custom_org" in selector:
                return self.create_org
            if "header-organization-dialog-add-a-user-button" in selector:
                return self.add_person
            if "Parton_person_dialog_save_button" in selector:
                return self.save_person
            if "Parton_organization_dialog_select_persons_button" in selector:
                return self.select_person
        if by == By.ID:
            if selector == "oshs-select-button":
                return self.done
            if selector.startswith("outer_person_dialog-"):
                return _Element(selector)
        raise AssertionError(f"Unexpected locator: {by!r}, {selector!r}")

    def find_elements(self, _by, _selector):
        try:
            return [self.find_element(_by, _selector)]
        except AssertionError:
            return []


class CorrespondentCreationFlowTests(unittest.TestCase):
    def test_visible_xpath_prefers_candidate_in_topmost_dialog(self):
        """A visible control underneath a modal must never receive the click."""
        active_dialog = _Element("topmost-dialog")
        underlying = _ScopedElement("underlying-add-button")
        active = _ScopedElement("active-dialog-add-button", active_dialog)
        driver = _ScopeDriver([underlying, active])

        with patch.object(correspondent, "WebDriverWait", _PredicateWait):
            selected = correspondent._wait_visible_xpath(
                driver, "//*[normalize-space(text())='Добавить']", timeout=1
            )

        self.assertIs(selected, active)

    def test_find_visible_never_falls_back_to_covered_underlay(self):
        covered = _Element("covered")
        driver = _CoveredDriver(covered)

        selected = correspondent._find_visible(driver, By.ID, "covered")

        self.assertIsNone(selected)

    def test_address_creation_writes_full_address_and_confirms_all_seven_steps(self):
        driver = _CreationDriver()
        corr_input = _Element("correspondent-input")
        plus = _Element("plus")

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(
                correspondent,
                "find_input_near_label",
                return_value=corr_input,
            ),
            patch.object(
                correspondent,
                "_find_correspondent_add_button",
                return_value=plus,
            ),
            patch.object(
                correspondent,
                "_wait_visible_xpath",
                return_value=_Element("xpath-result"),
            ),
            patch.object(correspondent, "click", return_value=True) as click_mock,
            patch.object(
                correspondent,
                "_wait_for_correspondent_value",
                return_value=ADDRESS,
            ) as confirm_mock,
            patch.object(correspondent, "close_open_modals") as close_mock,
            patch.object(ui, "wait_modal_closed") as modal_wait_mock,
        ):
            result = correspondent.create_correspondent(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, True)
        self.assertEqual(
            driver.field_writes,
            [
                ("outer_person_dialog-last_name-input", ADDRESS),
                ("outer_person_dialog-first_name-input", ""),
                ("outer_person_dialog-middle_name-input", ""),
                ("outer_person_dialog-position-input", "ФЛ"),
            ],
        )
        self.assertEqual(
            [call.args[2] for call in click_mock.call_args_list],
            [
                "+ Корреспондент",
                "Добавить",
                "Создать организацию",
                "Добавить физ. лицо",
                "Сохранить карточку",
                "Выбрать физ. лиц",
                "Готово",
            ],
        )
        confirm_mock.assert_called_once_with(
            driver, ADDRESS, kind="address", timeout=5,
            allow_closed_input=True, baseline_input="",
        )
        modal_wait_mock.assert_called_once_with(driver)
        close_mock.assert_not_called()

    def test_missing_plus_button_returns_false_before_opening_dialogs(self):
        driver = object()
        corr_input = _Element("correspondent-input")

        with (
            patch.object(
                correspondent,
                "find_input_near_label",
                return_value=corr_input,
            ),
            patch.object(
                correspondent,
                "_find_correspondent_add_button",
                return_value=None,
            ),
            patch.object(correspondent, "click") as click_mock,
            patch.object(
                correspondent, "_wait_visible_xpath"
            ) as visible_wait_mock,
        ):
            result = correspondent.create_correspondent(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, False)
        click_mock.assert_not_called()
        visible_wait_mock.assert_not_called()

    def test_missing_search_dialog_transition_returns_false(self):
        """A successful click is insufficient when the next modal never opens."""
        driver = object()
        corr_input = _Element("correspondent-input")
        plus = _Element("plus")

        with (
            patch.object(
                correspondent,
                "find_input_near_label",
                return_value=corr_input,
            ),
            patch.object(
                correspondent,
                "_find_correspondent_add_button",
                return_value=plus,
            ),
            patch.object(correspondent, "click", return_value=True) as click_mock,
            patch.object(
                correspondent,
                "_wait_visible_xpath",
                side_effect=TimeoutException("search dialog did not open"),
            ),
            patch.object(correspondent, "close_open_modals") as close_mock,
        ):
            result = correspondent.create_correspondent(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, False)
        click_mock.assert_called_once_with(driver, plus, "+ Корреспондент")
        close_mock.assert_called_once_with(driver)

    def test_mix_aborts_before_save_attach_and_register_when_correspondent_fails(self):
        driver = object()
        doc = {
            "тема": "Обращение с сайта",
            "корреспондент": ADDRESS,
            "корреспондент_тип": "address",
            "корр_источник": "feedback-address",
            "корр_найден": True,
            "тип_название": "Письма, заявления и жалобы граждан, акционеров",
            "содержание": "Тестовое обращение",
        }

        with (
            patch.object(mix, "WebDriverWait", _ImmediateWait),
            patch.object(mix, "wait_and_click", return_value=True),
            patch.object(mix, "click", return_value=True) as click_mock,
            patch.object(mix, "fill_text"),
            patch.object(
                mix, "fill_correspondent_field", return_value=False
            ) as fill_correspondent_mock,
            patch.object(mix, "fill_corr_number") as number_mock,
            patch.object(mix, "fill_corr_date") as date_mock,
            patch.object(mix, "add_addressee") as addressee_mock,
            patch.object(mix, "fill_delivery_method") as delivery_mock,
            patch.object(mix, "find_msg_by_link") as find_attachment_mock,
            patch.object(mix, "attach_content") as attach_mock,
            patch.object(mix, "register_and_resolve") as register_mock,
            patch.object(mix, "close_open_modals") as close_modals_mock,
            patch.object(
                mix, "close_card_and_wait_main"
            ) as close_card_mock,
        ):
            with self.assertRaisesRegex(
                RuntimeError, "Корреспондент не подтвержден"
            ):
                mix.create_one_document(driver, doc, index=1, total=1)

        fill_correspondent_mock.assert_called_once_with(
            driver, ADDRESS, kind="address"
        )
        self.assertNotIn(
            "Сохранить", [call.args[2] for call in click_mock.call_args_list]
        )
        number_mock.assert_not_called()
        date_mock.assert_not_called()
        addressee_mock.assert_not_called()
        delivery_mock.assert_not_called()
        find_attachment_mock.assert_not_called()
        attach_mock.assert_not_called()
        register_mock.assert_not_called()
        close_modals_mock.assert_called_once_with(driver)
        close_card_mock.assert_called_once_with(driver)
        self.assertEqual(mix._last_result["status"], "FAILED")


if __name__ == "__main__":
    unittest.main()
