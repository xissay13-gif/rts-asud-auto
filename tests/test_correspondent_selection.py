import unittest
from unittest.mock import Mock, patch

from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException

from shared import correspondent
from shared.ui import DropdownOptions


ADDRESS = "644021, г. Омск, ул. 4-я Транспортная, д. 15, кв. 8"
FIO = "Иванов Иван Иванович"


class _ImmediateWait:
    """Runs Selenium predicates once so these tests never open a browser."""

    def __init__(self, driver, _timeout):
        self.driver = driver

    def until(self, predicate):
        result = predicate(self.driver)
        if not result:
            raise TimeoutException("predicate is false")
        return result


class _FourPollWait:
    def __init__(self, driver, _timeout):
        self.driver = driver

    def until(self, predicate):
        result = False
        for _ in range(4):
            result = predicate(self.driver)
        if not result:
            raise TimeoutException("predicate is false")
        return result


class _Input:
    def __init__(self):
        self.keys = []

    def click(self):
        return None

    def send_keys(self, *keys):
        self.keys.extend(keys)


class _Candidate:
    def __init__(self, text, css_class="document-card-text"):
        self.text = text
        self.tag_name = "span"
        self.css_class = css_class

    def get_attribute(self, name):
        if name == "class":
            return self.css_class
        return None


class _StaleCandidate:
    id = "stale-option"

    @property
    def text(self):
        raise RuntimeError("stale option")


class _Driver:
    def __init__(self, parent_option=None):
        self.parent_option = parent_option

    def execute_script(self, script, *_args):
        if "el.value === ''" in script:
            return True
        if "parentElement" in script:
            return self.parent_option
        return None


class CorrespondentFieldConfirmationTests(unittest.TestCase):
    def _read_state(self, state):
        driver = _Driver()
        driver.execute_script = Mock(side_effect=lambda script, *_args: (
            [] if script == correspondent._READ_COMMITTED_CORRESPONDENT_JS
            else state
        ))
        inp = _Input()
        with patch.object(
            correspondent, "find_input_near_label", return_value=inp
        ):
            value = correspondent._correspondent_field_value(driver)
        driver.execute_script.assert_any_call(
            correspondent._READ_COMMITTED_CORRESPONDENT_JS, "Корреспондент"
        )
        driver.execute_script.assert_any_call(
            correspondent._READ_CORRESPONDENT_FIELD_JS, inp
        )
        return value

    def test_committed_row_is_read_when_editable_input_was_replaced(self):
        driver = _Driver()
        driver.execute_script = Mock(return_value=[FIO])

        with patch.object(
            correspondent, "find_input_near_label", return_value=None
        ):
            state = correspondent._correspondent_field_state(driver)

        self.assertEqual(state["semantic_values"], [FIO])
        self.assertEqual(state["semantic_value"], FIO)
        self.assertEqual(state["input_value"], "")

    def test_plain_input_equal_to_expected_is_not_selected_while_popup_is_visible(self):
        value = self._read_state({
            "input_value": ADDRESS,
            "popup_visible": True,
            "semantic_values": [],
        })

        self.assertEqual(value, "")
        self.assertFalse(correspondent._correspondent_value_matches(
            value, ADDRESS, kind="address"
        ))

    def test_plain_input_equal_to_expected_is_not_selection_after_popup_closes(self):
        value = self._read_state({
            "input_value": ADDRESS,
            "popup_visible": False,
            "semantic_values": [],
        })

        self.assertEqual(value, "")
        self.assertFalse(correspondent._correspondent_value_matches(
            value, ADDRESS, kind="address"
        ))

    def test_unchanged_raw_query_is_not_a_confirmed_selection(self):
        with patch.object(
            correspondent, "_correspondent_field_state", return_value={
                "input_value": ADDRESS,
                "popup_visible": False,
                "semantic_values": [],
                "semantic_value": "",
            }
        ):
            selected = correspondent._wait_for_correspondent_value(
                object(), ADDRESS, kind="address", timeout=0.2,
                allow_closed_input=True, baseline_input=ADDRESS,
            )

        self.assertEqual(selected, "")

    def test_created_value_is_confirmed_after_empty_baseline(self):
        with patch.object(
            correspondent, "_correspondent_field_state", return_value={
                "input_value": ADDRESS,
                "popup_visible": False,
                "semantic_values": [],
                "semantic_value": "",
            }
        ):
            selected = correspondent._wait_for_correspondent_value(
                object(), ADDRESS, kind="address", timeout=0.2,
                allow_closed_input=True, baseline_input="",
            )

        self.assertEqual(selected, ADDRESS)

    def test_exact_address_is_confirmed_after_cleared_query_changes(self):
        with patch.object(
            correspondent, "_correspondent_field_state", return_value={
                "input_value": ADDRESS,
                "popup_visible": False,
                "semantic_values": [],
                "semantic_value": "",
            }
        ):
            selected = correspondent._wait_for_correspondent_value(
                object(), ADDRESS, kind="address", timeout=0.2,
                allow_closed_input=True, baseline_input="",
            )

        self.assertEqual(selected, ADDRESS)

    def test_semantic_chip_is_selected_even_if_popup_is_visible(self):
        value = self._read_state({
            "input_value": "",
            "popup_visible": True,
            "semantic_values": [ADDRESS],
        })

        self.assertEqual(value, ADDRESS)
        with patch.object(
            correspondent, "_correspondent_field_state", return_value={
                "input_value": "",
                "popup_visible": True,
                "semantic_values": [ADDRESS],
                "semantic_value": ADDRESS,
            }
        ):
            selected = correspondent._wait_for_correspondent_value(
                object(), ADDRESS, kind="address", timeout=0.2
            )
        self.assertEqual(selected, "")


class CorrespondentSelectionTests(unittest.TestCase):
    def _patch_flow(self, *, driver, candidates, field_value, create_result):
        inp = _Input()
        def field_state(*_args, **_kwargs):
            value = (field_value(*_args, **_kwargs)
                     if callable(field_value) else field_value)
            return {
                "input_value": "",
                "popup_visible": False,
                "semantic_values": [value] if value else [],
                "semantic_value": value or "",
            }
        create_patch = (
            patch.object(correspondent, "create_correspondent", side_effect=create_result)
            if callable(create_result)
            else patch.object(correspondent, "create_correspondent", return_value=create_result)
        )
        return (
            inp,
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(correspondent, "find_input_near_label", return_value=inp),
            patch.object(correspondent, "js_type_combobox", return_value=None),
            patch.object(correspondent, "find_dropdown_options", return_value=candidates),
            patch.object(correspondent, "_correspondent_field_state", side_effect=field_state),
            patch.object(
                correspondent, "_clear_query_before_option_click", return_value=True
            ),
            patch.object(correspondent, "cdp_click", return_value=True),
            create_patch,
        )

    def test_committed_exact_row_finishes_without_clicking_hidden_input(self):
        driver = object()

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(
                correspondent, "_committed_correspondent_values",
                return_value=[FIO],
            ),
            patch.object(
                correspondent, "find_input_near_label"
            ) as input_mock,
            patch.object(correspondent, "create_correspondent") as create_mock,
        ):
            result = correspondent.fill_correspondent_field(
                driver, FIO, kind="person"
            )

        self.assertIs(result, True)
        input_mock.assert_not_called()
        create_mock.assert_not_called()

    def test_different_committed_row_fails_closed_without_replacing_it(self):
        driver = object()

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(
                correspondent, "_committed_correspondent_values",
                return_value=["Петров П П"],
            ),
            patch.object(
                correspondent, "find_input_near_label"
            ) as input_mock,
            patch.object(correspondent, "create_correspondent") as create_mock,
        ):
            result = correspondent.fill_correspondent_field(
                driver, FIO, kind="person"
            )

        self.assertIs(result, False)
        input_mock.assert_not_called()
        create_mock.assert_not_called()

    def test_intercepted_editor_click_accepts_row_committed_during_rerender(self):
        driver = object()
        inp = _Input()
        inp.click = Mock(side_effect=ElementClickInterceptedException("GXT rerender"))

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(
                correspondent, "_committed_correspondent_values",
                side_effect=[[], [FIO]],
            ),
            patch.object(
                correspondent, "find_input_near_label", return_value=inp
            ),
            patch.object(correspondent, "js_type_combobox") as type_mock,
            patch.object(correspondent, "create_correspondent") as create_mock,
        ):
            result = correspondent.fill_correspondent_field(
                driver, FIO, kind="person"
            )

        self.assertIs(result, True)
        type_mock.assert_not_called()
        create_mock.assert_not_called()

    def test_address_text_from_document_card_does_not_trigger_creation(self):
        """Unscoped matching text is unknown popup state, not an empty lookup."""
        driver = _Driver(parent_option=None)
        false_candidate = _Candidate(f"Тема письма\nАдрес: {ADDRESS}")
        _inp, *patchers = self._patch_flow(
            driver=driver,
            candidates=[false_candidate],
            field_value=lambda *_args, **_kwargs: "",
            create_result=True,
        )

        with self._enter_patchers(patchers) as mocks:
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        create_mock = mocks[-1]
        cdp_click_mock = mocks[-2]
        self.assertIs(result, False)
        create_mock.assert_not_called()
        cdp_click_mock.assert_not_called()

    def test_unknown_popup_state_does_not_create_possible_duplicate(self):
        inp = _Input()
        driver = _Driver(parent_option=None)

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(
                correspondent, "find_input_near_label", return_value=inp
            ),
            patch.object(correspondent, "js_type_combobox", return_value=None),
            patch.object(
                correspondent,
                "find_dropdown_options",
                side_effect=RuntimeError("popup state is unknown"),
            ),
            patch.object(
                correspondent, "create_correspondent", return_value=True
            ) as create_mock,
        ):
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, False)
        create_mock.assert_not_called()

    def test_real_option_with_confirmed_field_does_not_create_duplicate(self):
        option = object()
        driver = _Driver(parent_option=option)
        candidate = _Candidate(ADDRESS, css_class="x-boundlist-item")
        inp, *patchers = self._patch_flow(
            driver=driver,
            candidates=[candidate],
            field_value=lambda *_args, **_kwargs: ADDRESS,
            create_result=True,
        )

        with self._enter_patchers(patchers) as mocks:
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        cdp_click_mock = mocks[-2]
        create_mock = mocks[-1]
        self.assertIs(result, True)
        cdp_click_mock.assert_called_once_with(driver, option)
        create_mock.assert_not_called()
        self.assertFalse(inp.keys)

    def test_successful_cdp_click_without_confirmation_is_failure(self):
        """Sending a mouse event is not proof that ASUD accepted the selection.

        A real option already exists, so an unconfirmed click must not create a
        duplicate; the document should fail and be retried instead.
        """
        driver = _Driver(parent_option=object())
        candidate = _Candidate(ADDRESS, css_class="x-boundlist-item")
        _inp, *patchers = self._patch_flow(
            driver=driver,
            candidates=[candidate],
            field_value=lambda *_args, **_kwargs: "",
            create_result=True,
        )

        with self._enter_patchers(patchers) as mocks:
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        cdp_click_mock = mocks[-2]
        create_mock = mocks[-1]
        self.assertIs(result, False)
        cdp_click_mock.assert_called_once()
        create_mock.assert_not_called()

    def test_cdp_click_does_not_promote_unchanged_exact_query_to_selection(self):
        """The typed address is still only a query, even after CDP dispatch.

        Address lookup types the full expected value.  If ASUD closes (or we
        fail to detect) the popup without committing a row, that unchanged
        input must not be accepted as proof of selection.
        """
        option = object()
        driver = _Driver(parent_option=option)
        candidate = _Candidate(ADDRESS, css_class="x-boundlist-item")
        inp = _Input()
        raw_query_only = {
            "input_value": ADDRESS,
            "popup_visible": False,
            "semantic_values": [],
            "semantic_value": "",
        }

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(
                correspondent, "find_input_near_label", return_value=inp
            ),
            patch.object(correspondent, "js_type_combobox"),
            patch.object(
                correspondent, "find_dropdown_options",
                return_value=[candidate],
            ),
            patch.object(
                correspondent, "_correspondent_field_state",
                return_value=raw_query_only,
            ),
            patch.object(
                correspondent, "_clear_query_before_option_click",
                return_value=False,
            ),
            patch.object(
                correspondent, "cdp_click", return_value=True
            ) as cdp_click_mock,
            patch.object(
                correspondent, "create_correspondent", return_value=True
            ) as create_mock,
        ):
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, False)
        cdp_click_mock.assert_called_once_with(driver, option)
        create_mock.assert_not_called()

    def test_cdp_selection_is_accepted_only_after_popup_closes(self):
        """A committed row under an open popup is not final UI state yet."""
        option = object()
        driver = _Driver(parent_option=option)
        candidate = _Candidate(ADDRESS, css_class="x-boundlist-item")
        inp = _Input()
        states = [
            {
                "input_value": "",
                "popup_visible": True,
                "semantic_values": [ADDRESS],
                "semantic_value": ADDRESS,
            },
            {
                "input_value": "",
                "popup_visible": False,
                "semantic_values": [ADDRESS],
                "semantic_value": ADDRESS,
            },
        ]

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(
                correspondent, "find_input_near_label", return_value=inp
            ),
            patch.object(correspondent, "js_type_combobox"),
            patch.object(
                correspondent, "find_dropdown_options",
                return_value=[candidate],
            ),
            patch.object(
                correspondent, "_correspondent_field_state",
                side_effect=states,
            ) as state_mock,
            patch.object(
                correspondent, "_clear_query_before_option_click",
                return_value=True,
            ),
            patch.object(
                correspondent, "cdp_click", return_value=True
            ) as cdp_click_mock,
            patch.object(correspondent, "create_correspondent") as create_mock,
        ):
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, True)
        self.assertEqual(state_mock.call_count, 2)
        cdp_click_mock.assert_called_once_with(driver, option)
        create_mock.assert_not_called()

    def test_stale_nonempty_options_never_trigger_duplicate_creation(self):
        driver = _Driver(parent_option=None)
        candidates = DropdownOptions(
            [_StaleCandidate()], popup_seen=True, input_value=ADDRESS,
            input_observed=True,
        )
        _inp, *patchers = self._patch_flow(
            driver=driver,
            candidates=candidates,
            field_value=lambda *_args, **_kwargs: "",
            create_result=True,
        )

        with self._enter_patchers(patchers) as mocks:
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, False)
        mocks[-1].assert_not_called()

    def test_failed_creation_returns_false(self):
        driver = _Driver(parent_option=None)
        _inp, *patchers = self._patch_flow(
            driver=driver,
            candidates=DropdownOptions(
                [], popup_seen=True, empty_explicit=True
            ),
            field_value=lambda *_args, **_kwargs: "",
            create_result=False,
        )

        with self._enter_patchers(patchers) as mocks:
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        create_mock = mocks[-1]
        self.assertIs(result, False)
        create_mock.assert_called_once_with(driver, ADDRESS, kind="address")

    def test_created_card_is_requeried_and_selected_without_second_creation(self):
        """ASUD may save a person card without binding it to the document."""
        driver = _Driver(parent_option=None)
        inp = _Input()
        candidate = _Candidate(FIO, css_class="x-boundlist-item")
        lookups = [
            DropdownOptions([], popup_seen=True, empty_explicit=True),
            DropdownOptions([candidate], popup_seen=True, input_value="Иванов"),
        ]

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(
                correspondent, "find_input_near_label", return_value=inp
            ),
            patch.object(correspondent, "js_type_combobox", return_value=None),
            patch.object(
                correspondent, "find_dropdown_options", side_effect=lookups
            ) as lookup_mock,
            patch.object(
                correspondent, "create_correspondent", return_value=True
            ) as create_mock,
            patch.object(
                correspondent,
                "_wait_for_correspondent_value",
                side_effect=["", FIO],
            ),
            patch.object(
                correspondent, "_clear_query_before_option_click", return_value=True
            ),
            patch.object(correspondent, "cdp_click", return_value=True),
        ):
            result = correspondent.fill_correspondent_field(
                driver, FIO, kind="person"
            )

        self.assertIs(result, True)
        create_mock.assert_called_once_with(driver, FIO, kind="person")
        self.assertEqual(lookup_mock.call_count, 2)

    def test_created_card_not_yet_indexed_never_creates_a_duplicate(self):
        driver = _Driver(parent_option=None)
        inp = _Input()
        empty = DropdownOptions([], popup_seen=True, empty_explicit=True)

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(correspondent, "WebDriverWait", _ImmediateWait),
            patch.object(
                correspondent, "find_input_near_label", return_value=inp
            ),
            patch.object(correspondent, "js_type_combobox", return_value=None),
            patch.object(
                correspondent, "find_dropdown_options", return_value=empty
            ),
            patch.object(
                correspondent, "create_correspondent", return_value=True
            ) as create_mock,
            patch.object(
                correspondent, "_wait_for_correspondent_value", return_value=""
            ),
        ):
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, False)
        create_mock.assert_called_once_with(driver, ADDRESS, kind="address")

    def test_stable_anchored_blank_popup_creates_after_full_wait(self):
        driver = _Driver(parent_option=None)
        inp = _Input()
        blank = DropdownOptions(
            [], popup_seen=True, popup_key="popup-1",
            signature="same-empty-popup", input_value=ADDRESS,
            root_blank=True, input_observed=True,
        )

        with (
            patch.object(correspondent.time, "sleep", return_value=None),
            patch.object(
                correspondent.time, "monotonic",
                side_effect=[0.0, 0.0, 3.0, 6.0, 9.5],
            ),
            patch.object(correspondent, "WebDriverWait", _FourPollWait),
            patch.object(
                correspondent, "find_input_near_label", return_value=inp
            ),
            patch.object(correspondent, "js_type_combobox", return_value=None),
            patch.object(
                correspondent, "find_dropdown_options", return_value=blank
            ),
            patch.object(
                correspondent, "create_correspondent", return_value=True
            ) as create_mock,
            patch.object(
                correspondent, "_wait_for_correspondent_value",
                return_value=ADDRESS,
            ),
        ):
            result = correspondent.fill_correspondent_field(
                driver, ADDRESS, kind="address"
            )

        self.assertIs(result, True)
        create_mock.assert_called_once_with(driver, ADDRESS, kind="address")

    def test_blank_or_loading_popup_never_creates_possible_duplicate(self):
        driver = _Driver(parent_option=None)

        for candidates in (
            DropdownOptions([], popup_seen=True),
            DropdownOptions([], popup_seen=True, loading=True),
            DropdownOptions([], popup_seen=False),
            DropdownOptions(
                [], popup_seen=True, root_blank=True, popup_key="popup",
                signature="blank", input_value="", input_observed=True,
            ),
        ):
            with self.subTest(
                popup_seen=candidates.popup_seen,
                loading=candidates.loading,
            ):
                _inp, *patchers = self._patch_flow(
                    driver=driver,
                    candidates=candidates,
                    field_value=lambda *_args, **_kwargs: "",
                    create_result=True,
                )

                with self._enter_patchers(patchers) as mocks:
                    result = correspondent.fill_correspondent_field(
                        driver, ADDRESS, kind="address"
                    )

                self.assertIs(result, False)
                mocks[-1].assert_not_called()

    @staticmethod
    def _enter_patchers(patchers):
        class _PatchStack:
            def __enter__(self):
                self.started = [p.start() for p in patchers]
                return self.started

            def __exit__(self, exc_type, exc, tb):
                for p in reversed(patchers):
                    p.stop()

        return _PatchStack()


if __name__ == "__main__":
    unittest.main()
