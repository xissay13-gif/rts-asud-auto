import unittest
from unittest.mock import patch

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from flows import mix
from shared import correspondent, ui


ADDRESS = "644021, г. Омск, ул. 4-я Транспортная, д. 15, кв. 8"
FIO = "Гудов Александр Анатольевич"


class _Element:
    def __init__(self, name, text=None):
        self.name = name
        self.text = text if text is not None else name

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


class _CreationWait:
    def __init__(self, driver, _timeout):
        self.driver = driver

    def until(self, predicate):
        result = predicate(self.driver)
        if not result:
            raise TimeoutException("predicate is false")
        return result


class _SaveTransitionTimeoutWait(_CreationWait):
    def until(self, predicate):
        result = predicate(self.driver)
        if result is self.driver.select_person:
            raise TimeoutException("saved card transition did not finish")
        if not result:
            raise TimeoutException("predicate is false")
        return result


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
        self.created_person_selected = False
        self.add_search = _Element("add-search")
        self.create_org = _Element("create-organization")
        self.add_person = _Element("add-person")
        self.save_person = _Element("save-person")
        self.person_row = _Element("person-row")
        self.person_checker = _Element("person-checker")
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
        if script == correspondent._CONTROL_ENABLED_JS:
            element = args[0]
            if element is self.select_person:
                return self.created_person_selected
            return True
        if script == correspondent._FIND_OUTER_PERSON_ROWS_JS:
            return [{
                "row": self.person_row,
                "text": ADDRESS,
                "checker": self.person_checker,
            }]
        if script == correspondent._ROW_SELECTED_JS:
            return self.created_person_selected
        raise AssertionError("Unexpected JavaScript in creation flow")

    def find_element(self, by, selector):
        if by == By.CSS_SELECTOR:
            if "outer_organisation_dialog-create_new_org_label" in selector:
                return self.create_org
            if "outer_organisation_dialog-add_person_button" in selector:
                return self.add_person
            if "outer_person_dialog-save_button" in selector:
                return self.save_person
            if "outer_organisation_dialog-select_persons_button" in selector:
                return self.select_person
        if by == By.ID:
            if selector == "outer_org_person_add_button":
                return self.add_search
            if selector == "oshs-select-button":
                return self.done
            if selector == "outer_organisation_dialog-add_person_button":
                return self.add_person
            if selector == "outer_organisation_dialog-select_persons_button":
                return self.select_person
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

    def test_disabled_gxt_control_is_not_enabled(self):
        button = _Element("disabled-select")

        class _Driver:
            @staticmethod
            def execute_script(script, _element):
                return script != correspondent._CONTROL_ENABLED_JS

        self.assertFalse(correspondent._control_enabled(_Driver(), button))

    def test_row_selector_uses_row_checker_not_header_checkbox(self):
        self.assertIn("td[cellindex='1']", correspondent._FIND_OUTER_PERSON_ROWS_JS)
        self.assertIn("td[cellindex='0'] .check-cell", correspondent._FIND_OUTER_PERSON_ROWS_JS)
        self.assertNotIn("FCPC_selection_column", correspondent._FIND_OUTER_PERSON_ROWS_JS)

    def test_selected_state_supports_real_gxt_camel_case_classes(self):
        self.assertIn("row(?:checker)?checked", correspondent._ROW_SELECTED_JS)
        self.assertIn("rowselected", correspondent._ROW_SELECTED_JS)

    def test_selects_only_exact_person_row_and_waits_for_enabled_button(self):
        driver = object()
        button = _Element("select-persons")
        wrong_checker = _Element("wrong-checker")
        exact_checker = _Element("exact-checker")
        records = [
            {"row": _Element("wrong-row"), "text": "Огудов А А",
             "checker": wrong_checker},
            {"row": _Element("exact-row"), "text": "Гудов А А",
             "checker": exact_checker},
        ]

        with (
            patch.object(correspondent, "WebDriverWait", _PredicateWait),
            patch.object(correspondent, "_find_visible", return_value=button),
            patch.object(correspondent, "_outer_person_rows", return_value=records),
            patch.object(
                correspondent, "_control_enabled",
                side_effect=[False, False, True],
            ),
            patch.object(correspondent, "click", return_value=True) as click_mock,
        ):
            result = correspondent._select_created_person_row(
                driver, FIO, kind="person", timeout=1
            )

        self.assertIs(result, True)
        click_mock.assert_called_once_with(
            driver, exact_checker, "Отметить созданное физ. лицо"
        )

    def test_stale_row_click_refetches_record_before_retry(self):
        driver = object()
        button = _Element("select-persons")
        first_checker = _Element("stale-checker")
        fresh_checker = _Element("fresh-checker")
        records = [
            [{"row": _Element("row-1"), "text": "Гудов А А",
              "checker": first_checker}],
            [{"row": _Element("row-2"), "text": "Гудов А А",
              "checker": fresh_checker}],
        ]

        with (
            patch.object(correspondent, "WebDriverWait", _PredicateWait),
            patch.object(correspondent, "_find_visible", return_value=button),
            patch.object(
                correspondent, "_outer_person_rows",
                side_effect=lambda *_args: records.pop(0) if len(records) > 1 else records[0],
            ) as rows_mock,
            patch.object(
                correspondent, "_control_enabled",
                side_effect=[False, False, False, True],
            ),
            patch.object(
                correspondent, "click", side_effect=[False, True]
            ) as click_mock,
        ):
            result = correspondent._select_created_person_row(
                driver, FIO, kind="person", timeout=1
            )

        self.assertIs(result, True)
        self.assertGreaterEqual(rows_mock.call_count, 3)
        self.assertEqual(
            [call.args[1] for call in click_mock.call_args_list],
            [fresh_checker, fresh_checker],
        )

    def test_duplicate_exact_person_rows_fail_closed(self):
        driver = object()
        button = _Element("select-persons")
        records = [
            {"row": _Element("row-1"), "text": "Гудов А А",
             "checker": _Element("checker-1")},
            {"row": _Element("row-2"), "text": FIO,
             "checker": _Element("checker-2")},
        ]

        with (
            patch.object(correspondent, "WebDriverWait", _PredicateWait),
            patch.object(correspondent, "_find_visible", return_value=button),
            patch.object(correspondent, "_outer_person_rows", return_value=records),
            patch.object(correspondent, "click") as click_mock,
        ):
            result = correspondent._select_created_person_row(
                driver, FIO, kind="person", timeout=1
            )

        self.assertIs(result, False)
        click_mock.assert_not_called()

    def test_address_creation_selects_exact_created_row_and_confirms_all_eight_steps(self):
        driver = _CreationDriver()
        corr_input = _Element("correspondent-input")
        plus = _Element("plus")

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _CreationWait),
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
            patch.object(
                correspondent,
                "_outer_person_rows",
                return_value=[{
                    "row": driver.person_row,
                    "text": ADDRESS,
                    "checker": driver.person_checker,
                }],
            ),
            patch.object(
                correspondent,
                "click",
                side_effect=lambda _driver, _element, description: (
                    setattr(driver, "created_person_selected", True)
                    if description == "Отметить созданное физ. лицо" else None
                ) is None,
            ) as click_mock,
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
                ("outer_person_dialog-first_name-input", "-"),
                ("outer_person_dialog-middle_name-input", "-"),
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
                "Отметить созданное физ. лицо",
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

    def test_saved_card_without_auto_binding_is_recoverable(self):
        """Creation success and binding success are separate ASUD transitions."""
        driver = _CreationDriver()

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(
                correspondent,
                "find_input_near_label",
                return_value=_Element("correspondent-input"),
            ),
            patch.object(
                correspondent,
                "_find_correspondent_add_button",
                return_value=_Element("plus"),
            ),
            patch.object(
                correspondent,
                "_wait_visible_xpath",
                return_value=_Element("xpath-result"),
            ),
            patch.object(
                correspondent,
                "_select_created_person_row",
                return_value=True,
            ),
            patch.object(correspondent, "click", return_value=True),
            patch.object(
                correspondent, "_wait_for_correspondent_value", return_value=""
            ),
            patch.object(correspondent, "close_open_modals") as close_mock,
            patch.object(ui, "wait_modal_closed"),
        ):
            result = correspondent.create_correspondent(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, True)
        close_mock.assert_not_called()

    def test_saved_card_row_selection_failure_uses_no_create_recovery(self):
        """A saved card is never reported as absent and recreated as a duplicate."""
        driver = _CreationDriver()

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(
                correspondent,
                "find_input_near_label",
                return_value=_Element("correspondent-input"),
            ),
            patch.object(
                correspondent,
                "_find_correspondent_add_button",
                return_value=_Element("plus"),
            ),
            patch.object(
                correspondent,
                "_wait_visible_xpath",
                return_value=_Element("xpath-result"),
            ),
            patch.object(
                correspondent, "_select_created_person_row", return_value=False
            ),
            patch.object(correspondent, "click", return_value=True) as click_mock,
            patch.object(correspondent, "close_open_modals") as close_mock,
        ):
            result = correspondent.create_correspondent(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, True)
        close_mock.assert_called_once_with(driver)
        self.assertNotIn(
            "Выбрать физ. лиц",
            [call.args[2] for call in click_mock.call_args_list],
        )

    def test_save_transition_timeout_never_authorizes_duplicate_creation(self):
        driver = _CreationDriver()

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(
                correspondent, "WebDriverWait", _SaveTransitionTimeoutWait
            ),
            patch.object(
                correspondent,
                "find_input_near_label",
                return_value=_Element("correspondent-input"),
            ),
            patch.object(
                correspondent,
                "_find_correspondent_add_button",
                return_value=_Element("plus"),
            ),
            patch.object(
                correspondent,
                "_wait_visible_xpath",
                return_value=_Element("xpath-result"),
            ),
            patch.object(correspondent, "click", return_value=True),
            patch.object(correspondent, "close_open_modals") as close_mock,
        ):
            result = correspondent.create_correspondent(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, True)
        close_mock.assert_called_once_with(driver)

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
