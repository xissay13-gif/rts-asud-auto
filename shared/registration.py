"""Проверяемая регистрация документа и отправка на резолюцию.

Клик Selenium/CDP означает лишь доставку mouse event. Этот модуль считает
операцию успешной только после наблюдаемого перехода интерфейса АСУД:

* регистрация — появилась активная кнопка ``На резолюцию``;
* резолюция — исчез диалог подтверждения и стабильно виден главный экран.
"""

from dataclasses import dataclass
from enum import Enum
import logging
import time

from selenium.webdriver.common.action_chains import ActionChains


log = logging.getLogger("asud.registration")


class RegistrationPhase(str, Enum):
    WAIT_REGISTER = "wait_register"
    REGISTER_SUBMITTED = "register_submitted"
    REGISTERED = "registered"
    RESOLUTION_SUBMITTED = "resolution_submitted"
    CONFIRM_VISIBLE = "confirm_visible"
    CONFIRM_SUBMITTED = "confirm_submitted"
    RESOLVED = "resolved"
    FAILED = "failed"


class ClickDelivery(str, Enum):
    NOT_ATTEMPTED = "not_attempted"
    DISPATCHED = "dispatched"
    UNKNOWN_AFTER_ATTEMPT = "unknown_after_attempt"


@dataclass(frozen=True)
class RegistrationSnapshot:
    register_actionable: bool = False
    resolution_actionable: bool = False
    confirm_actionable: bool = False
    main_visible: bool = False
    progress: bool = False
    asud_id: str | None = None
    register_present: bool = False
    resolution_present: bool = False
    confirm_present: bool = False
    busy: bool = False


@dataclass(frozen=True)
class RegistrationOutcome:
    registered: bool = False
    resolved: bool = False
    asud_id: str | None = None
    phase: RegistrationPhase = RegistrationPhase.FAILED
    reason: str = ""
    submission_uncertain: bool = False

    @property
    def ok(self):
        return self.registered and self.resolved


_OBSERVE_JS = r"""
function visible(el) {
    if (!el || !el.isConnected) return false;
    const r = el.getBoundingClientRect();
    const s = window.getComputedStyle(el);
    return r.width > 0 && r.height > 0 &&
        s.display !== 'none' && s.visibility !== 'hidden' &&
        parseFloat(s.opacity || '1') > 0.01;
}

function exposed(el) {
    if (!visible(el)) return false;
    const r = el.getBoundingClientRect();
    const x = r.left + r.width / 2;
    const y = r.top + r.height / 2;
    if (x < 0 || y < 0 || x >= innerWidth || y >= innerHeight) return false;
    const hit = document.elementFromPoint(x, y);
    return !!hit && (hit === el || el.contains(hit) || hit.contains(el));
}

function disabled(el) {
    if (!el) return true;
    const cls = String(el.className || '').toLowerCase();
    const s = window.getComputedStyle(el);
    return el.disabled === true || el.getAttribute('data-disabled') === '1' ||
        String(el.getAttribute('aria-disabled') || '').toLowerCase() === 'true' ||
        /(?:^|[\s_-])disabled(?:$|[\s_-])/.test(cls) ||
        s.pointerEvents === 'none';
}

function actionable(el) {
    return exposed(el) && !disabled(el);
}

function control(selectors) {
    let present = null;
    for (const selector of selectors) {
        for (const el of document.querySelectorAll(selector)) {
            if (!visible(el)) continue;
            if (!present) present = el;
            // GXT may keep an older visible-looking copy under an overlay.
            // Never let that first copy hide the later topmost control.
            if (actionable(el)) return {present, actionable: el};
        }
    }
    return {present, actionable: null};
}

const registerState = control([
    '#header-action-btn-register', "[id*='header-action-btn-register']"
]);
const resolutionState = control([
    '#header-action-btn-send_on_resolution',
    "[id*='header-action-btn-send_on_resolution']"
]);
const confirmState = control([
    '#confirm-dialog-btn-yes', '#confirm_dialog_btn_yes',
    "[id*='confirm-dialog-btn-yes']", "[id*='confirm_dialog_btn_yes']",
    "[id*='confirm'][id*='yes']"
]);
const mainState = control(['#mainscreen-create-button']);

let busy = false;
for (const el of document.querySelectorAll(
        "[aria-busy='true'],[class*='loading'],[class*='Loading']," +
        "[class*='loadmask'],[class*='LoadMask']")) {
    if (visible(el)) { busy = true; break; }
}

const body = document.body ? (document.body.innerText || '') : '';
let asudId = null;
let match = body.match(/№\s+([А-ЯЁ]{2,5}(?:\/[А-ЯЁ0-9.\-]+){2,})\s+от/u);
if (match) asudId = match[1];

const registerPresent = !!registerState.present;
const registerActionable = !!registerState.actionable;
const resolutionPresent = !!resolutionState.present;
const resolutionActionable = !!resolutionState.actionable;
const confirmPresent = !!confirmState.present;
const confirmActionable = !!confirmState.actionable;
const mainVisible = !!mainState.actionable;

return {
    register_actionable: registerActionable,
    resolution_actionable: resolutionActionable,
    confirm_actionable: confirmActionable,
    main_visible: mainVisible,
    register_present: registerPresent,
    resolution_present: resolutionPresent,
    confirm_present: confirmPresent,
    busy,
    asud_id: asudId,
    // These are progress signals only. Registration success still requires
    // the actionable resolution control.
    progress: busy || resolutionPresent || confirmPresent ||
        (registerPresent && !registerActionable) || !!asudId,
};
"""


_FIND_CONTROL_JS = r"""
const kind = arguments[0];
const selectors = {
    register: ['#header-action-btn-register', "[id*='header-action-btn-register']"],
    resolution: ['#header-action-btn-send_on_resolution',
        "[id*='header-action-btn-send_on_resolution']"],
    confirm: ['#confirm-dialog-btn-yes', '#confirm_dialog_btn_yes',
        "[id*='confirm-dialog-btn-yes']", "[id*='confirm_dialog_btn_yes']",
        "[id*='confirm'][id*='yes']"],
}[kind] || [];

function actionable(el) {
    if (!el || !el.isConnected) return false;
    const r = el.getBoundingClientRect();
    const s = window.getComputedStyle(el);
    if (r.width <= 0 || r.height <= 0 || s.display === 'none' ||
            s.visibility === 'hidden' || parseFloat(s.opacity || '1') <= 0.01 ||
            s.pointerEvents === 'none') return false;
    const cls = String(el.className || '').toLowerCase();
    if (el.disabled === true || el.getAttribute('data-disabled') === '1' ||
            String(el.getAttribute('aria-disabled') || '').toLowerCase() === 'true' ||
            /(?:^|[\s_-])disabled(?:$|[\s_-])/.test(cls)) return false;
    const x = r.left + r.width / 2, y = r.top + r.height / 2;
    if (x < 0 || y < 0 || x >= innerWidth || y >= innerHeight) return false;
    const hit = document.elementFromPoint(x, y);
    return !!hit && (hit === el || el.contains(hit) || hit.contains(el));
}

for (const selector of selectors) {
    for (const el of document.querySelectorAll(selector)) {
        if (actionable(el)) return el;
    }
}
return null;
"""


def _observe(driver):
    try:
        raw = driver.execute_script(_OBSERVE_JS) or {}
    except Exception:
        return RegistrationSnapshot()
    if not isinstance(raw, dict):
        return RegistrationSnapshot()
    allowed = RegistrationSnapshot.__dataclass_fields__.keys()
    return RegistrationSnapshot(**{key: raw.get(key) for key in allowed})


def _find_actionable(driver, kind):
    try:
        return driver.execute_script(_FIND_CONTROL_JS, kind)
    except Exception:
        return None


def _click_control(driver, kind):
    element = _find_actionable(driver, kind)
    if not element:
        return ClickDelivery.NOT_ATTEMPTED
    try:
        ActionChains(driver).move_to_element(element).pause(0.15).click().perform()
        return ClickDelivery.DISPATCHED
    except Exception:
        # Selenium may raise after POST /actions was already accepted. This
        # state must never authorize another potentially duplicate request.
        return ClickDelivery.UNKNOWN_AFTER_ATTEMPT


def _click_register(driver):
    return _click_control(driver, "register")


def _click_resolution(driver):
    return _click_control(driver, "resolution")


def _click_confirm(driver):
    return _click_control(driver, "confirm")


def _sleep(interval):
    time.sleep(max(float(interval), 0.001))


def _capture_id(capture_id, driver, fallback):
    if capture_id is None:
        return fallback
    try:
        return capture_id(driver) or fallback
    except Exception:
        return fallback


def run_registration(driver, *, timeout=20.0, retry_grace=2.5,
                     poll_interval=0.1, capture_id=None, logger=None):
    """Регистрирует документ и отправляет его на резолюцию.

    Повтор ``Зарегистрировать`` допускается ровно один раз только когда helper
    сообщил, что событие вообще не было доставлено (stale/intercepted). После
    доставленного клика таймерный повтор запрещён: медленный ответ сервера
    невозможно надёжно отличить от no-op без риска двойной регистрации.
    """
    current_log = logger or log
    timeout = max(float(timeout), 0.01)
    # Аргумент оставлен для совместимости RC/tests. Таймерный retry больше не
    # используется: повтор допустим только после явной ошибки доставки клика.
    _ = retry_grace

    deadline = time.monotonic() + timeout
    initial = RegistrationSnapshot()
    while time.monotonic() < deadline:
        initial = _observe(driver)
        if initial.resolution_actionable or initial.register_actionable:
            break
        _sleep(poll_interval)

    asud_id = initial.asud_id
    if initial.resolution_actionable:
        registered = True
        phase = RegistrationPhase.REGISTERED
    elif not initial.register_actionable:
        return RegistrationOutcome(
            asud_id=asud_id, phase=RegistrationPhase.FAILED,
            reason="активная кнопка регистрации не появилась",
        )
    else:
        current_log.info("Регистрирую...")
        first_delivery = _click_register(driver)

        # A stale/intercepted first control is refetched immediately once.
        if first_delivery is ClickDelivery.NOT_ATTEMPTED:
            _sleep(poll_interval)
            snap = _observe(driver)
            if snap.register_actionable:
                first_delivery = _click_register(driver)

        registered = False
        phase = RegistrationPhase.REGISTER_SUBMITTED
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            snap = _observe(driver)
            if snap.asud_id:
                asud_id = snap.asud_id
            if snap.resolution_actionable:
                registered = True
                phase = RegistrationPhase.REGISTERED
                break

            _sleep(poll_interval)

        if not registered:
            return RegistrationOutcome(
                asud_id=asud_id, phase=RegistrationPhase.FAILED,
                reason=(
                    "результат отправленного клика регистрации не определён"
                    if first_delivery is not ClickDelivery.NOT_ATTEMPTED else
                    "регистрация не была отправлена"
                ),
                submission_uncertain=(
                    first_delivery is not ClickDelivery.NOT_ATTEMPTED
                ),
            )

    asud_id = _capture_id(capture_id, driver, asud_id)
    if asud_id:
        current_log.info("Документ зарегистрирован: %s", asud_id)
    else:
        current_log.info("Документ зарегистрирован (номер не захвачен)")

    resolution_delivery = _click_resolution(driver)
    if resolution_delivery is ClickDelivery.NOT_ATTEMPTED:
        # Если событие точно не начиналось, разрешён один поиск свежей GXT-
        # кнопки. UNKNOWN_AFTER_ATTEMPT повторять нельзя: POST /actions мог
        # успеть уйти на сервер до исключения Selenium.
        _sleep(poll_interval)
        snap = _observe(driver)
        if snap.resolution_actionable:
            resolution_delivery = _click_resolution(driver)
    if resolution_delivery is ClickDelivery.NOT_ATTEMPTED:
        return RegistrationOutcome(
            registered=True, asud_id=asud_id,
            phase=RegistrationPhase.REGISTERED,
            reason="не удалось нажать «На резолюцию»",
        )

    phase = RegistrationPhase.RESOLUTION_SUBMITTED
    deadline = time.monotonic() + timeout
    confirm_seen = False
    main_samples = 0
    while time.monotonic() < deadline:
        snap = _observe(driver)
        if snap.asud_id:
            asud_id = snap.asud_id
        # Some ASUD builds skip the confirmation dialog, but main must be
        # stable and both action controls gone before that can be success.
        if (snap.main_visible and not snap.confirm_present and
                not snap.register_present and not snap.resolution_present):
            main_samples += 1
            if main_samples >= 2:
                current_log.info("Документ отправлен на резолюцию")
                return RegistrationOutcome(
                    registered=True, resolved=True, asud_id=asud_id,
                    phase=RegistrationPhase.RESOLVED,
                )
        else:
            main_samples = 0
        if snap.confirm_actionable:
            confirm_seen = True
            phase = RegistrationPhase.CONFIRM_VISIBLE
            break
        _sleep(poll_interval)

    if not confirm_seen:
        return RegistrationOutcome(
            registered=True, asud_id=asud_id, phase=phase,
            reason="диалог подтверждения резолюции не появился",
        )

    confirm_delivery = _click_confirm(driver)
    if confirm_delivery is ClickDelivery.NOT_ATTEMPTED:
        _sleep(poll_interval)
        snap = _observe(driver)
        if snap.confirm_actionable:
            confirm_delivery = _click_confirm(driver)
    if confirm_delivery is ClickDelivery.NOT_ATTEMPTED:
        return RegistrationOutcome(
            registered=True, asud_id=asud_id,
            phase=RegistrationPhase.CONFIRM_VISIBLE,
            reason="не удалось нажать «Да»",
        )

    phase = RegistrationPhase.CONFIRM_SUBMITTED
    deadline = time.monotonic() + timeout
    main_samples = 0
    while time.monotonic() < deadline:
        snap = _observe(driver)
        if snap.asud_id:
            asud_id = snap.asud_id
        if (not snap.confirm_present and snap.main_visible
                and not snap.register_present
                and not snap.resolution_present):
            main_samples += 1
            if main_samples >= 2:
                current_log.info("Документ отправлен на резолюцию")
                return RegistrationOutcome(
                    registered=True, resolved=True, asud_id=asud_id,
                    phase=RegistrationPhase.RESOLVED,
                )
        else:
            main_samples = 0
        _sleep(poll_interval)

    return RegistrationOutcome(
        registered=True, asud_id=asud_id, phase=phase,
        reason="после «Да» главный экран не подтверждён",
    )
