import unittest
from unittest.mock import patch

from shared import registration


class _Clock:
    """Deterministic clock for polling state-machine tests."""

    def __init__(self):
        self.now = 0.0

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        # Ensure even a zero poll interval advances a timeout-bound loop.
        self.now += max(float(seconds), 0.001)


class _RegistrationDriver:
    """Small browserless model of the controls rendered by ASUD/GXT."""

    def __init__(
        self,
        *,
        stale_first_register=False,
        registration_progress_polls=0,
        silent_registration_polls=0,
        registration_never_finishes=False,
        main_after_confirm=True,
        register_delivery=registration.ClickDelivery.DISPATCHED,
        register_dispatched_no_transition=False,
        resolution_delivery=registration.ClickDelivery.DISPATCHED,
        resolution_skips_confirm=False,
        confirm_delivery=registration.ClickDelivery.DISPATCHED,
    ):
        self.phase = "draft"
        self.stale_first_register = stale_first_register
        self.registration_progress_polls = registration_progress_polls
        self.silent_registration_polls = silent_registration_polls
        self.registration_never_finishes = registration_never_finishes
        self.main_after_confirm = main_after_confirm
        self.register_delivery = register_delivery
        self.register_dispatched_no_transition = register_dispatched_no_transition
        self.resolution_delivery = resolution_delivery
        self.resolution_skips_confirm = resolution_skips_confirm
        self.confirm_delivery = confirm_delivery

        self.register_clicks = 0
        self.resolution_clicks = 0
        self.confirm_clicks = 0
        self.registration_polls = 0

    def observe(self):
        if self.phase == "draft":
            return registration.RegistrationSnapshot(register_actionable=True)

        if self.phase == "registering":
            self.registration_polls += 1
            if self.registration_never_finishes:
                # A changed document state/number proves that the first request
                # is in flight, but does not by itself prove registration.
                return registration.RegistrationSnapshot(
                    register_actionable=False,
                    progress=True,
                    asud_id="TEST/1/999",
                )
            if self.registration_polls <= self.registration_progress_polls:
                # Deliberately leave the register control actionable: a blind
                # timer-based retry would submit the document twice here.
                return registration.RegistrationSnapshot(
                    register_actionable=True,
                    progress=True,
                )
            if self.registration_polls <= self.silent_registration_polls:
                # Server accepted the click but the old actionable control has
                # not been rerendered yet. A timer retry would be a duplicate.
                return registration.RegistrationSnapshot(
                    register_actionable=True,
                    progress=False,
                )
            self.phase = "registered"

        if self.phase == "registered":
            return registration.RegistrationSnapshot(
                resolution_actionable=True,
                asud_id="TEST/1/999",
            )

        if self.phase == "confirm":
            return registration.RegistrationSnapshot(
                confirm_actionable=True,
                asud_id="TEST/1/999",
            )

        if self.phase == "after_confirm":
            return registration.RegistrationSnapshot(
                main_visible=self.main_after_confirm,
                asud_id="TEST/1/999",
            )

        raise AssertionError(f"unknown fake phase: {self.phase}")

    def click_register(self):
        self.register_clicks += 1
        if self.stale_first_register and self.register_clicks == 1:
            # Models the GXT node becoming stale between lookup and click.
            # The next attempt must locate a fresh control through the seam.
            return registration.ClickDelivery.NOT_ATTEMPTED
        if (self.register_delivery is registration.ClickDelivery.DISPATCHED and
                not self.register_dispatched_no_transition):
            self.phase = "registering"
        return self.register_delivery

    def click_resolution(self):
        self.resolution_clicks += 1
        if self.phase != "registered":
            return registration.ClickDelivery.NOT_ATTEMPTED
        if self.resolution_delivery is registration.ClickDelivery.DISPATCHED:
            self.phase = (
                "after_confirm" if self.resolution_skips_confirm else "confirm"
            )
        return self.resolution_delivery

    def click_confirm(self):
        self.confirm_clicks += 1
        if self.phase != "confirm":
            return registration.ClickDelivery.NOT_ATTEMPTED
        if self.confirm_delivery is registration.ClickDelivery.DISPATCHED:
            self.phase = "after_confirm"
        return self.confirm_delivery


class RegistrationStateMachineTests(unittest.TestCase):
    def _run(self, driver, *, timeout=1.0, retry_grace=0.05):
        clock = _Clock()
        with (
            patch.object(registration, "_observe", side_effect=lambda d: d.observe()),
            patch.object(
                registration,
                "_click_register",
                side_effect=lambda d: d.click_register(),
            ),
            patch.object(
                registration,
                "_click_resolution",
                side_effect=lambda d: d.click_resolution(),
            ),
            patch.object(
                registration,
                "_click_confirm",
                side_effect=lambda d: d.click_confirm(),
            ),
            patch.object(registration.time, "monotonic", side_effect=clock.monotonic),
            patch.object(registration.time, "sleep", side_effect=clock.sleep),
        ):
            return registration.run_registration(
                driver,
                timeout=timeout,
                retry_grace=retry_grace,
                poll_interval=0.01,
                capture_id=None,
                logger=None,
            )

    def test_stale_first_register_control_is_refetched_and_retried_once(self):
        driver = _RegistrationDriver(stale_first_register=True)

        outcome = self._run(driver)

        self.assertTrue(outcome.ok)
        self.assertTrue(outcome.registered)
        self.assertTrue(outcome.resolved)
        self.assertEqual(driver.register_clicks, 2)
        self.assertEqual(driver.resolution_clicks, 1)
        self.assertEqual(driver.confirm_clicks, 1)

    def test_slow_registration_progress_never_causes_duplicate_click(self):
        driver = _RegistrationDriver(registration_progress_polls=12)

        outcome = self._run(driver, retry_grace=0.02)

        self.assertTrue(outcome.ok)
        self.assertEqual(driver.register_clicks, 1)
        self.assertGreater(driver.registration_polls, 2)

    def test_silent_slow_server_never_causes_duplicate_click(self):
        driver = _RegistrationDriver(silent_registration_polls=12)

        outcome = self._run(driver, retry_grace=0.02)

        self.assertTrue(outcome.ok)
        self.assertEqual(driver.register_clicks, 1)
        self.assertGreater(driver.registration_polls, 2)

    def test_unknown_register_delivery_is_never_retried(self):
        driver = _RegistrationDriver(
            register_delivery=registration.ClickDelivery.UNKNOWN_AFTER_ATTEMPT,
        )

        outcome = self._run(driver, timeout=0.15)

        self.assertFalse(outcome.registered)
        self.assertFalse(outcome.resolved)
        self.assertTrue(outcome.submission_uncertain)
        self.assertEqual(driver.register_clicks, 1)
        self.assertEqual(driver.resolution_clicks, 0)

    def test_dispatched_register_without_transition_is_submission_uncertain(self):
        driver = _RegistrationDriver(register_dispatched_no_transition=True)

        outcome = self._run(driver, timeout=0.15)

        self.assertFalse(outcome.registered)
        self.assertFalse(outcome.resolved)
        self.assertTrue(outcome.submission_uncertain)
        self.assertEqual(driver.register_clicks, 1)
        self.assertEqual(driver.resolution_clicks, 0)

    def test_unknown_resolution_delivery_is_not_false_success(self):
        driver = _RegistrationDriver(
            resolution_delivery=registration.ClickDelivery.UNKNOWN_AFTER_ATTEMPT,
        )

        outcome = self._run(driver)

        self.assertTrue(outcome.registered)
        self.assertFalse(outcome.resolved)
        self.assertFalse(outcome.ok)
        self.assertEqual(driver.resolution_clicks, 1)
        self.assertEqual(driver.confirm_clicks, 0)

    def test_unknown_confirm_delivery_is_not_retried_or_false_success(self):
        driver = _RegistrationDriver(
            confirm_delivery=registration.ClickDelivery.UNKNOWN_AFTER_ATTEMPT,
        )

        outcome = self._run(driver)

        self.assertTrue(outcome.registered)
        self.assertFalse(outcome.resolved)
        self.assertFalse(outcome.ok)
        self.assertEqual(driver.confirm_clicks, 1)

    def test_asud_id_and_progress_are_not_registration_success_without_resolution(self):
        driver = _RegistrationDriver(registration_never_finishes=True)

        outcome = self._run(driver, timeout=0.15, retry_grace=0.02)

        self.assertFalse(outcome.registered)
        self.assertFalse(outcome.resolved)
        self.assertFalse(outcome.ok)
        self.assertEqual(driver.register_clicks, 1)
        self.assertEqual(driver.resolution_clicks, 0)
        self.assertEqual(driver.confirm_clicks, 0)

    def test_confirm_disappearing_without_main_screen_is_not_resolution_success(self):
        driver = _RegistrationDriver(main_after_confirm=False)

        outcome = self._run(driver, timeout=0.15)

        self.assertTrue(outcome.registered)
        self.assertFalse(outcome.resolved)
        self.assertFalse(outcome.ok)
        self.assertEqual(driver.register_clicks, 1)
        self.assertEqual(driver.resolution_clicks, 1)
        self.assertEqual(driver.confirm_clicks, 1)

    def test_resolution_succeeds_only_after_confirm_is_gone_and_main_is_visible(self):
        driver = _RegistrationDriver(main_after_confirm=True)

        outcome = self._run(driver)

        self.assertTrue(outcome.registered)
        self.assertTrue(outcome.resolved)
        self.assertTrue(outcome.ok)
        self.assertEqual(outcome.asud_id, "TEST/1/999")
        self.assertEqual(driver.register_clicks, 1)
        self.assertEqual(driver.resolution_clicks, 1)
        self.assertEqual(driver.confirm_clicks, 1)

    def test_build_without_confirm_requires_stable_main_and_gone_actions(self):
        driver = _RegistrationDriver(resolution_skips_confirm=True)

        outcome = self._run(driver)

        self.assertTrue(outcome.ok)
        self.assertEqual(driver.register_clicks, 1)
        self.assertEqual(driver.resolution_clicks, 1)
        self.assertEqual(driver.confirm_clicks, 0)


if __name__ == "__main__":
    unittest.main()
