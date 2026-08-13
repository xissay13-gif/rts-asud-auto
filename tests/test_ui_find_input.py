import re
import unittest
from dataclasses import dataclass

from shared.ui import find_input_near_label


@dataclass
class _Input:
    """Small DOM input double used by the browserless script harness."""

    name: str
    connected: bool = True
    read_only: bool = False
    disabled: bool = False
    aria_hidden: str | None = None
    tab_index: str | None = None
    width: float = 240
    height: float = 24
    display: str = "block"
    visibility: str = "visible"
    opacity: float = 1
    z_index: int = 0
    offset_parent: bool = True
    covered: bool = False


class _BrowserlessDomDriver:
    """Runs the input-selection contract without starting Edge.

    ``find_input_near_label`` deliberately keeps its DOM walk in JavaScript.
    This test double interprets the visibility guards used by that script and
    applies them to inputs in DOM order.  It therefore catches the regression
    where the first text input is a GXT 1x1 focus/service input rather than the
    editor a user can click.
    """

    def __init__(self, inputs):
        self.inputs = inputs
        self.calls = []

    def execute_script(self, script, label_text):
        self.calls.append((script, label_text))
        for candidate in self.inputs:
            if self._script_accepts(script, candidate):
                return candidate
        return None

    @staticmethod
    def _script_accepts(script, candidate):
        if "inp.offsetParent" in script and not candidate.offset_parent:
            return False
        if "inp.isConnected" in script and not candidate.connected:
            return False
        if "inp.readOnly" in script and candidate.read_only:
            return False
        if "inp.disabled" in script and candidate.disabled:
            return False

        if "aria-hidden" in script and (candidate.aria_hidden or "").lower() == "true":
            return False

        rejects_minus_one = re.search(
            r"getAttribute\(['\"]tabindex['\"]\)\s*={2,3}\s*['\"]-1['\"]",
            script,
        )
        if rejects_minus_one and candidate.tab_index == "-1":
            return False

        width_limit = re.search(r"rect\.width\s*<\s*([0-9.]+)", script)
        height_limit = re.search(r"rect\.height\s*<\s*([0-9.]+)", script)
        if width_limit and candidate.width < float(width_limit.group(1)):
            return False
        if height_limit and candidate.height < float(height_limit.group(1)):
            return False

        if "style.display" in script and candidate.display == "none":
            return False
        if "style.visibility" in script and candidate.visibility in {"hidden", "collapse"}:
            return False
        if "style.opacity" in script and candidate.opacity <= 0.05:
            return False
        if "style.zIndex" in script and candidate.z_index < 0:
            return False
        if "elementFromPoint" in script and candidate.covered:
            return False
        return True


def _hidden_gxt_service_input():
    # Mirrors the element from the Edge error: it is laid out, but is a 1x1,
    # transparent, negative-z-index focus sink and must never be clicked.
    return _Input(
        "hidden service input",
        aria_hidden="true",
        tab_index="-1",
        width=1,
        height=1,
        opacity=0,
        z_index=-1,
        covered=True,
    )


class FindInputNearLabelTests(unittest.TestCase):
    def test_skips_hidden_service_input_and_returns_visible_editor(self):
        hidden = _hidden_gxt_service_input()
        visible = _Input("visible correspondent editor")
        driver = _BrowserlessDomDriver([hidden, visible])

        result = find_input_near_label(driver, "Корреспондент")

        self.assertIs(result, visible)
        self.assertEqual(len(driver.calls), 1)
        self.assertEqual(driver.calls[0][1], "Корреспондент")

    def test_selected_chip_with_only_hidden_service_input_returns_none(self):
        hidden = _hidden_gxt_service_input()
        driver = _BrowserlessDomDriver([hidden])

        result = find_input_near_label(driver, "Корреспондент")

        self.assertIsNone(result)

    def test_visible_editable_input_can_have_tabindex_minus_one(self):
        visible = _Input("visible editor outside tab order", tab_index="-1")
        driver = _BrowserlessDomDriver([visible])

        result = find_input_near_label(driver, "Корреспондент")

        self.assertIs(result, visible)


if __name__ == "__main__":
    unittest.main()
