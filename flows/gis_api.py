"""Experimental, fail-closed API registration flow for GIS ЖКХ messages.

This module is intentionally independent from Selenium.  It accepts the
already parsed document dictionary produced by :mod:`flows.email`, performs an
exact correspondent lookup, creates one incoming document with its MSG file,
registers it and (when configured) sends it on resolution.

The live path is write-ahead and at-most-once.  A persistent ``.claim``
sidecar is created with ``O_EXCL`` and is never automatically removed.  Before
every mutating request a JSON state sidecar is flushed and atomically replaced.
Consequently an interrupted mutation is surfaced for manual review instead of
being submitted a second time.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import json
import logging
import os
import re
import unicodedata
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, MutableMapping, Optional, Tuple

from shared.asud_api import (
    AsudApiClient,
    AsudApiConfig,
    AsudApiConfigError,
    AsudApiError,
    response_succeeded,
)


_LOG = logging.getLogger("asud.gis_api")
_UUID_NAMESPACE = uuid.uuid5(uuid.NAMESPACE_DNS, "rts-asud-auto.gis-zhkh")
_FIO_PART = re.compile(r"[А-ЯЁ][А-Яа-яЁё]*(?:-[А-ЯЁ][А-Яа-яЁё]*)*\Z")
_OUTCOME_STATUSES = {
    "OK",
    "DRY_RUN",
    "PROBE",
    "FAILED",
    "MANUAL_REVIEW",
    "SUBMISSION_UNKNOWN",
    "REGISTERED_ONLY",
}
_RESERVED_ATTRIBUTES = {
    "typePath",
    "description",
    "addresseesIds",
    "deliveryType",
    "correspondentId",
    "correspondentsRegDate",
    "correspondentsRegNumber",
    "extExaminationDate",
}


@dataclass(frozen=True)
class GisApiOutcome:
    status: str
    object_id: Optional[str] = None
    registration_number: Optional[str] = None
    external_guid: Optional[str] = None
    state_path: Optional[str] = None
    message: str = ""

    def __post_init__(self) -> None:
        if self.status not in _OUTCOME_STATUSES:
            raise ValueError(f"Unsupported GIS API outcome: {self.status}")


def _normalized(value: Any) -> str:
    return " ".join(
        unicodedata.normalize("NFKC", "" if value is None else str(value))
        .strip()
        .split()
    )


def _comparison(value: Any) -> str:
    return _normalized(value).casefold().replace("ё", "е")


def gis_external_guid(appeal_number: Any, branch_id: Any) -> str:
    """Return a stable UUIDv5 derived only from GIS number and ASUD branch."""

    appeal = _comparison(appeal_number)
    branch = _comparison(branch_id)
    if not appeal or not branch:
        raise ValueError("appeal number and branch are required for external GUID")
    return str(uuid.uuid5(_UUID_NAMESPACE, f"{branch}\x00{appeal}"))


def extract_strict_fio(doc: Mapping[str, Any]) -> Optional[Tuple[str, str, str]]:
    """Extract an exact three-part Cyrillic FIO, otherwise return ``None``.

    The boolean parser flag is part of the trust boundary: a placeholder or an
    anonymous GIS appeal must go to manual review and must never be used to
    select a similarly named ASUD correspondent.
    """

    if doc.get("корр_найден") is not True:
        return None
    if _comparison(doc.get("корреспондент_тип", "person")) != "person":
        return None
    parts = _normalized(doc.get("корреспондент")).split(" ")
    if len(parts) != 3 or any(not _FIO_PART.fullmatch(part) for part in parts):
        return None
    if any(part in ("-", "—") for part in parts):
        return None
    return parts[0], parts[1], parts[2]


def should_use_gis_api(doc: Mapping[str, Any], settings: Mapping[str, Any]) -> bool:
    """Return true only for a structured, non-excluded GIS document.

    Selection of the email backend (``ASUD_EMAIL_REGISTRATION_BACKEND``) is an
    outer integration concern.  This seam enforces the inner API enable switch
    and refuses non-GIS documents even if a caller routes one here by mistake.
    """

    if not isinstance(doc, Mapping) or doc.get("skip_asud_registration"):
        return False
    try:
        config = AsudApiConfig.from_settings(settings)
    except AsudApiConfigError:
        return False
    if not config.enabled:
        return False
    appeal = _normalized(doc.get("номер_обращения"))
    # This parser provenance is the authoritative discriminator.  Subject and
    # topic text are user-controlled and must never route an ordinary message
    # into the API mutation path merely because it contains "ГИС ЖКХ".
    if not appeal or _comparison(doc.get("корр_источник")) != "zhkh":
        return False
    file_path = _normalized(doc.get("файл"))
    return bool(
        file_path
        and file_path.casefold().endswith(".msg")
    )


def build_exact_correspondent_query(
    config: AsudApiConfig,
    fio: Tuple[str, str, str],
) -> Dict[str, Any]:
    surname, first_name, middle_name = fio
    return {
        "userLogin": config.user,
        "queryType": "ddv_outer_org_person",
        "requiredAttributes": {"attribute": ["distinct *"]},
        "criterion": [
            {
                "name": "r_object_type",
                "operation": "EQUALS",
                "value1": "ddt_person",
            },
            {
                "name": "dss_relation_type",
                "operation": "EQUALS",
                "value1": "CORRESPONDENT_RELATION",
            },
            {
                "name": "dsid_branch",
                "operation": "EQUALS",
                "value1": config.branch_id,
            },
        {
            "name": "dss_organisation",
            "operation": "EQUALS",
            "value1": surname,
        },
        {
            "name": "dss_full_name",
            "operation": "STARTS_WITH",
            "value1": surname,
        },
        {
            "name": "dss_first_name",
            "operation": "STARTS_WITH",
            "value1": first_name,
        },
        {
            "name": "dss_middle_name",
            "operation": "STARTS_WITH",
            "value1": middle_name,
        },
        ],
    }


def _iter_attribute_pairs(value: Any) -> Iterable[Tuple[str, Any]]:
    if isinstance(value, Mapping):
        if "key" in value and "value" in value:
            yield _normalized(value.get("key")), value.get("value")
            return
        for key in ("attribute", "attributes"):
            if key in value:
                yield from _iter_attribute_pairs(value[key])
        # Some REST serializers return attributes as a plain object.
        for key, item in value.items():
            if key not in ("attribute", "attributes") and not isinstance(
                item, (Mapping, list, tuple)
            ):
                yield _normalized(key), item
    elif isinstance(value, (list, tuple)):
        for item in value:
            yield from _iter_attribute_pairs(item)


def _attribute_map(value: Any) -> Optional[Dict[str, str]]:
    result: Dict[str, str] = {}
    for key, raw_value in _iter_attribute_pairs(value):
        if not key:
            continue
        normalized_value = _normalized(raw_value)
        if key in result:
            if _comparison(result[key]) != _comparison(normalized_value):
                # Flattened serializers can repeat an attribute.  Conflicting
                # identity values make the whole candidate ambiguous.
                return None
            continue
        result[key] = normalized_value
    return result


def extract_exact_correspondent_id(
    response: Mapping[str, Any],
    fio: Tuple[str, str, str],
    branch_id: str,
) -> Optional[str]:
    """Return an ID only when exactly one response candidate fully matches."""

    if not response_succeeded(response):
        return None
    raw_objects = response.get("objects", [])
    if isinstance(raw_objects, Mapping):
        objects = [raw_objects]
    elif isinstance(raw_objects, list):
        objects = raw_objects
    else:
        return None

    expected = {
        "r_object_type": "ddt_person",
        "dss_relation_type": "CORRESPONDENT_RELATION",
        "dsid_branch": branch_id,
        # The documented query uses this field as the organisation/name
        # discriminator.  Require the response to echo it so an ignored server
        # criterion cannot silently select a different correspondent.
        "dss_organisation": fio[0],
    }
    exact_ids = []
    for candidate in objects:
        attrs = _attribute_map(candidate)
        if attrs is None:
            continue
        core_exact = all(
            key in attrs and _comparison(attrs[key]) == _comparison(value)
            for key, value in expected.items()
        )
        separate_fio = all(
            key in attrs
            and _comparison(attrs[key]) == _comparison(value)
            for key, value in {
                "dss_last_name": fio[0],
                "dss_first_name": fio[1],
                "dss_middle_name": fio[2],
            }.items()
        )
        full_fio = (
            "dss_full_name" in attrs
            and _comparison(attrs["dss_full_name"])
            == _comparison(" ".join(fio))
        )
        if core_exact and (separate_fio or full_fio):
            object_id = _normalized(attrs.get("r_object_id"))
            if object_id:
                exact_ids.append(object_id)
    unique_ids = list(dict.fromkeys(exact_ids))
    return unique_ids[0] if len(unique_ids) == 1 else None


def _strict_fio_value(value: Any) -> Optional[Tuple[str, str, str]]:
    parts = _normalized(value).split(" ")
    if len(parts) != 3 or any(not _FIO_PART.fullmatch(part) for part in parts):
        return None
    return parts[0], parts[1], parts[2]


def build_exact_appointment_query(
    config: AsudApiConfig,
    fio: Tuple[str, str, str],
) -> Dict[str, Any]:
    """Build the literal documented getOshsAppontmentByCriteria request."""

    surname, first_name, middle_name = fio
    return {
        "criterion": [
            {
                "name": "dss_user_last_name",
                "operation": "EQUALS",
                "value1": surname,
            },
            {
                "name": "dss_user_first_name",
                "operation": "EQUALS",
                "value1": first_name,
            },
            {
                "name": "dss_user_middle_name",
                "operation": "EQUALS",
                "value1": middle_name,
            },
            {
                "name": "ddt_branch.dss_name",
                "operation": "EQUALS",
                "value1": config.branch_name,
            },
        ]
    }


def extract_exact_appointment_id(
    response: Mapping[str, Any],
    fio: Tuple[str, str, str],
    branch_id: str,
    branch_name: str,
) -> Optional[str]:
    """Return one active appointment ID matching exact FIO and branch."""

    if not response_succeeded(response):
        return None
    raw_objects = response.get("typedObjects", [])
    if isinstance(raw_objects, Mapping):
        objects = [raw_objects]
    elif isinstance(raw_objects, list):
        objects = raw_objects
    else:
        return None

    matches = []
    expected = {
        "dss_user_last_name": fio[0],
        "dss_user_first_name": fio[1],
        "dss_user_middle_name": fio[2],
        "dsid_branch": branch_id,
        "branch_name": branch_name,
        "dsb_deleted": "0",
    }
    for candidate in objects:
        if not isinstance(candidate, Mapping):
            continue
        if _comparison(candidate.get("type")) != "ddt_appointment":
            continue
        attrs = _attribute_map(candidate.get("attributes", candidate))
        if attrs is None or not all(
            key in attrs and _comparison(attrs[key]) == _comparison(value)
            for key, value in expected.items()
        ):
            continue
        object_id = _normalized(attrs.get("r_object_id"))
        global_id = _normalized(attrs.get("dsid_global_id"))
        if not object_id or (global_id and global_id != object_id):
            continue
        matches.append(object_id)
    # The API contract requires exactly one active appointment object.  Even
    # two byte-for-byte duplicate rows are ambiguous server data and must not
    # be collapsed into an apparent unique result.
    return matches[0] if len(matches) == 1 else None


def _asud_midnight(value: Any) -> Optional[str]:
    if isinstance(value, _dt.datetime):
        day = value.date()
    elif isinstance(value, _dt.date):
        day = value
    else:
        raw = _normalized(value)
        if not raw:
            return None
        day = None
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S"):
            try:
                day = _dt.datetime.strptime(raw, fmt).date()
                break
            except ValueError:
                continue
        if day is None:
            return None
    return day.isoformat() + "T00:00:00"


def build_incoming_payload(
    doc: Mapping[str, Any],
    config: AsudApiConfig,
    correspondent_id: str,
    external_guid: str,
) -> Dict[str, Any]:
    """Build the documented ``handleObject`` incoming JSON payload."""

    msg_path = Path(str(doc.get("файл") or ""))
    if not msg_path.is_file() or msg_path.suffix.casefold() != ".msg":
        raise ValueError("GIS MSG attachment is missing")
    attachment_size = msg_path.stat().st_size
    if (
        not config.confirm_msg_supported
        or config.max_attachment_bytes <= 0
        or attachment_size <= 0
        or attachment_size > config.max_attachment_bytes
    ):
        raise ValueError("GIS MSG attachment is not explicitly allowed")
    appeal_number = _normalized(doc.get("номер_обращения"))
    correspondent_date = _asud_midnight(doc.get("дата_обращения"))
    if not appeal_number or not correspondent_date:
        raise ValueError("GIS appeal number/date is missing or invalid")
    description = _normalized(doc.get("содержание"))
    if not description:
        raise ValueError("GIS description is missing")

    configured_identity_attributes = [
        config.branch_attribute,
        config.author_attribute,
    ]
    normalized_identity_attributes = [
        value.casefold() for value in configured_identity_attributes if value
    ]
    reserved_casefold = {value.casefold() for value in _RESERVED_ATTRIBUTES}
    if (
        len(set(normalized_identity_attributes))
        != len(normalized_identity_attributes)
        or any(value in reserved_casefold
               for value in normalized_identity_attributes)
    ):
        raise AsudApiConfigError(
            "branch/author attribute names must be distinct non-core fields"
        )
    reserved = reserved_casefold.union(normalized_identity_attributes)
    collisions = reserved.intersection(
        str(key).casefold() for key in config.incoming_attributes
    )
    if collisions:
        raise AsudApiConfigError(
            "incoming_attributes cannot override core incoming fields"
        )

    attributes = [
        {"key": "typePath", "value": config.incoming_type_path},
        {"key": "description", "value": description},
        {"key": "addresseesIds", "value": config.addressee_id},
        {"key": "correspondentId", "value": correspondent_id},
        {"key": "correspondentsRegDate", "value": correspondent_date},
        {"key": "correspondentsRegNumber", "value": appeal_number},
    ]
    if config.delivery_type:
        attributes.append({"key": "deliveryType", "value": config.delivery_type})
    examination_date = _asud_midnight(doc.get("планируемая_дата"))
    if examination_date:
        attributes.append({"key": "extExaminationDate", "value": examination_date})
    if config.branch_attribute:
        attributes.append({
            "key": config.branch_attribute,
            "value": config.branch_id,
        })
    if config.author_attribute:
        attributes.append({
            "key": config.author_attribute,
            "value": config.author_id,
        })
    for key in sorted(config.incoming_attributes):
        attributes.append({
            "key": key,
            "value": str(config.incoming_attributes[key]),
        })

    content = msg_path.read_bytes()
    if not content or len(content) > config.max_attachment_bytes:
        raise ValueError("GIS MSG attachment exceeds the configured limit")
    return {
        "lis": config.lis,
        "user": config.user,
        "typedObject": {
            "type": "incoming",
            "guid": external_guid,
            "attributes": attributes,
            "file": {
                "name": msg_path.name,
                "checkSum": hashlib.md5(content).hexdigest(),  # nosec B303
                "content": list(content),
            },
            "fullCreate": True,
        },
    }


def extract_created_object_id(response: Mapping[str, Any]) -> Optional[str]:
    if not response_succeeded(response):
        return None
    document = response.get("doc")
    if not isinstance(document, Mapping):
        return None
    return _normalized(document.get("objectId") or document.get("docId")) or None


def extract_registration_number(response: Mapping[str, Any]) -> Optional[str]:
    """Extract one unambiguous registration number from ``getObject`` JSON."""

    if not response_succeeded(response):
        return None
    typed = response.get("typedObject")
    if not isinstance(typed, Mapping):
        return None
    values = []
    for key, value in _iter_attribute_pairs(typed.get("attributes", typed)):
        if key.casefold() == "dss_reg_number" and _normalized(value):
            values.append(_normalized(value))
    unique = list(dict.fromkeys(values))
    return unique[0] if len(unique) == 1 else None


def extract_registration_date(response: Mapping[str, Any]) -> Optional[str]:
    """Extract one unambiguous registration date from ``getObject`` JSON."""

    if not response_succeeded(response):
        return None
    typed = response.get("typedObject")
    if not isinstance(typed, Mapping):
        return None
    values = []
    for key, value in _iter_attribute_pairs(typed.get("attributes", typed)):
        if key.casefold() == "dsdt_reg_date" and _normalized(value):
            values.append(_normalized(value))
    unique = list(dict.fromkeys(values))
    return unique[0] if len(unique) == 1 else None


def _atomic_write_json(path: Path, state: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(
        f".{path.name}.{os.getpid()}.{uuid.uuid4().hex}.tmp"
    )
    encoded = json.dumps(
        state, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    try:
        with open(temporary, "xb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
        # Best-effort directory flush (not supported for directories on all
        # Windows/Python combinations); file flush + replace remains atomic.
        try:
            directory_fd = os.open(str(path.parent), os.O_RDONLY)
        except OSError:
            directory_fd = None
        if directory_fd is not None:
            try:
                os.fsync(directory_fd)
            except OSError:
                pass
            finally:
                os.close(directory_fd)
    finally:
        try:
            temporary.unlink()
        except FileNotFoundError:
            pass


def _claim_message(claim_path: Path, external_guid: str) -> bool:
    payload = json.dumps({
        "version": 1,
        "external_guid": external_guid,
        "claimed_at": _dt.datetime.now(_dt.timezone.utc).isoformat(),
    }, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
    try:
        fd = os.open(str(claim_path), flags, 0o600)
    except FileExistsError:
        return False
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
    except Exception:
        try:
            claim_path.unlink()
        except OSError:
            pass
        raise
    return True


def _read_state(path: Path) -> Optional[Mapping[str, Any]]:
    try:
        with open(path, "r", encoding="utf-8") as stream:
            state = json.load(stream)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError):
        return None
    return state if isinstance(state, Mapping) else None


def _state_record(
    *,
    phase: str,
    external_guid: str,
    status: str = "IN_PROGRESS",
    object_id: Optional[str] = None,
    registration_number: Optional[str] = None,
) -> Dict[str, Any]:
    return {
        "version": 1,
        "status": status,
        "phase": phase,
        "external_guid": external_guid,
        "object_id": object_id,
        "registration_number": registration_number,
        "updated_at": _dt.datetime.now(_dt.timezone.utc).isoformat(),
    }


def _outcome_from_existing_state(
    state: Optional[Mapping[str, Any]],
    state_path: Path,
    external_guid: str,
) -> GisApiOutcome:
    if not state:
        return GisApiOutcome(
            "SUBMISSION_UNKNOWN",
            external_guid=external_guid,
            state_path=str(state_path),
            message="claim_exists_without_readable_state",
        )
    object_id = _normalized(state.get("object_id")) or None
    registration_number = _normalized(state.get("registration_number")) or None
    saved_guid = _normalized(state.get("external_guid"))
    if not saved_guid or saved_guid != external_guid:
        return GisApiOutcome(
            "MANUAL_REVIEW",
            external_guid=external_guid,
            state_path=str(state_path),
            message="existing_state_guid_mismatch",
        )
    status = _normalized(state.get("status"))
    if status in _OUTCOME_STATUSES:
        return GisApiOutcome(
            status,
            object_id=object_id,
            registration_number=registration_number,
            external_guid=saved_guid,
            state_path=str(state_path),
            message="existing_claim",
        )
    phase = _normalized(state.get("phase")).upper()
    if phase in ("REGISTRATION_SUCCEEDED",):
        recovered_status = "REGISTERED_ONLY"
    elif phase in ("HANDLE_OBJECT_SUCCEEDED",):
        recovered_status = "MANUAL_REVIEW"
    elif phase in ("RESOLUTION_SUCCEEDED",):
        recovered_status = "OK"
    else:
        # Every BEFORE_* state is deliberately uncertain: the process may
        # have crashed just after submission and before recording the result.
        recovered_status = "SUBMISSION_UNKNOWN"
    return GisApiOutcome(
        recovered_status,
        object_id=object_id,
        registration_number=registration_number,
        external_guid=saved_guid,
        state_path=str(state_path),
        message="existing_incomplete_claim",
    )


def _log(logger: Any, level: str, message: str) -> None:
    target = logger or _LOG
    method = getattr(target, level, None)
    if callable(method):
        method(message)


def _finish_claimed(
    state_path: Path,
    status: str,
    external_guid: str,
    *,
    object_id: Optional[str] = None,
    registration_number: Optional[str] = None,
    message: str = "",
) -> GisApiOutcome:
    try:
        _atomic_write_json(state_path, _state_record(
            phase="TERMINAL",
            status=status,
            external_guid=external_guid,
            object_id=object_id,
            registration_number=registration_number,
        ))
    except OSError:
        # The persistent claim still prevents a retry.  If an ASUD object may
        # exist, failure to persist the terminal state is itself uncertain.
        if object_id:
            status = "SUBMISSION_UNKNOWN"
            message = "terminal_state_write_failed"
    return GisApiOutcome(
        status,
        object_id=object_id,
        registration_number=registration_number,
        external_guid=external_guid,
        state_path=str(state_path),
        message=message,
    )


def process_gis_document(
    doc: MutableMapping[str, Any],
    settings: Mapping[str, Any],
    transport: Optional[Any] = None,
    logger: Optional[Any] = None,
) -> GisApiOutcome:
    """Process one parsed GIS message according to the configured safety mode."""

    if not isinstance(doc, MutableMapping):
        return GisApiOutcome("FAILED", message="document_must_be_mutable")
    try:
        config = AsudApiConfig.from_settings(settings)
    except AsudApiConfigError:
        return GisApiOutcome("FAILED", message="invalid_api_configuration")
    if not should_use_gis_api(doc, settings):
        return GisApiOutcome("FAILED", message="not_an_enabled_gis_document")

    appeal_number = _normalized(doc.get("номер_обращения"))
    try:
        external_guid = gis_external_guid(appeal_number, config.branch_id)
    except ValueError:
        return GisApiOutcome("FAILED", message="missing_guid_inputs")
    # Stable integration seam: the XLSX writer can persist this even for
    # dry-run/probe outcomes without learning the GUID algorithm.
    doc["asud_api_external_guid"] = external_guid

    fio = extract_strict_fio(doc)
    if fio is None:
        return GisApiOutcome(
            "MANUAL_REVIEW",
            external_guid=external_guid,
            message="strict_fio_required",
        )
    msg_path = Path(str(doc.get("файл") or ""))
    if not msg_path.is_file() or msg_path.suffix.casefold() != ".msg":
        return GisApiOutcome(
            "MANUAL_REVIEW",
            external_guid=external_guid,
            message="gis_msg_missing",
        )
    if _asud_midnight(doc.get("дата_обращения")) is None:
        return GisApiOutcome(
            "MANUAL_REVIEW",
            external_guid=external_guid,
            message="gis_appeal_date_invalid",
        )

    # A previous live claim is terminal for automatic processing.  Consult it
    # before even a read-only lookup so a restart is deterministic and does not
    # depend on current directory data or API availability.
    claim_path = Path(str(msg_path) + ".asud_api.claim")
    state_path = Path(str(msg_path) + ".asud_api_state.json")
    if config.mode == "live-one" and (
            claim_path.exists() or state_path.exists()):
        return _outcome_from_existing_state(
            _read_state(state_path), state_path, external_guid
        )

    if config.mode == "dry-run":
        _log(logger, "info", "GIS ASUD API dry-run validation completed")
        return GisApiOutcome(
            "DRY_RUN",
            external_guid=external_guid,
            message="no_network_or_mutation",
        )

    try:
        config.validate_probe()
    except AsudApiConfigError:
        return GisApiOutcome(
            "FAILED",
            external_guid=external_guid,
            message="probe_configuration_incomplete",
        )
    client = AsudApiClient(config, transport=transport, logger=logger)
    _log(logger, "info", "GIS ASUD API exact correspondent lookup started")
    try:
        search_response = client.find_objects2(
            build_exact_correspondent_query(config, fio)
        )
    except AsudApiError:
        return GisApiOutcome(
            "FAILED",
            external_guid=external_guid,
            message="exact_correspondent_lookup_failed",
        )
    correspondent_id = extract_exact_correspondent_id(
        search_response, fio, config.branch_id
    )
    if not correspondent_id:
        return GisApiOutcome(
            "MANUAL_REVIEW",
            external_guid=external_guid,
            message="exact_correspondent_not_unique",
        )

    addressee_fio = _strict_fio_value(config.addressee_name)
    if addressee_fio is None:
        return GisApiOutcome(
            "FAILED",
            external_guid=external_guid,
            message="strict_addressee_fio_required",
        )
    _log(logger, "info", "GIS ASUD API exact addressee lookup started")
    try:
        appointment_response = client.find_appointment(
            build_exact_appointment_query(config, addressee_fio)
        )
    except AsudApiError:
        return GisApiOutcome(
            "FAILED",
            external_guid=external_guid,
            message="exact_addressee_lookup_failed",
        )
    appointment_id = extract_exact_appointment_id(
        appointment_response,
        addressee_fio,
        config.branch_id,
        config.branch_name,
    )
    if not appointment_id or appointment_id != config.addressee_id:
        return GisApiOutcome(
            "MANUAL_REVIEW",
            external_guid=external_guid,
            message="exact_addressee_not_unique_or_id_mismatch",
        )
    if config.mode == "probe":
        _log(logger, "info", "GIS ASUD API probe completed without mutation")
        return GisApiOutcome(
            "PROBE",
            external_guid=external_guid,
            message="exact_correspondent_found",
        )

    try:
        config.validate_live()
        handle_payload = build_incoming_payload(
            doc, config, correspondent_id, external_guid
        )
    except (AsudApiConfigError, OSError, ValueError):
        return GisApiOutcome(
            "MANUAL_REVIEW",
            external_guid=external_guid,
            message="live_payload_validation_failed",
        )

    try:
        claimed = _claim_message(claim_path, external_guid)
    except OSError:
        return GisApiOutcome(
            "FAILED",
            external_guid=external_guid,
            state_path=str(state_path),
            message="claim_write_failed",
        )
    if not claimed:
        return _outcome_from_existing_state(
            _read_state(state_path), state_path, external_guid
        )

    # Barrier 1: creation.  If the process dies after this replace, restart
    # reports SUBMISSION_UNKNOWN and never resends handleObject.
    try:
        _atomic_write_json(state_path, _state_record(
            phase="BEFORE_HANDLE_OBJECT",
            external_guid=external_guid,
        ))
    except OSError:
        return GisApiOutcome(
            "MANUAL_REVIEW",
            external_guid=external_guid,
            state_path=str(state_path),
            message="write_ahead_state_failed",
        )
    _log(logger, "info", "GIS ASUD API handleObject submission started")
    try:
        handle_response = client.handle_object(handle_payload)
    except AsudApiError:
        return _finish_claimed(
            state_path,
            "SUBMISSION_UNKNOWN",
            external_guid,
            message="handle_object_submission_unknown",
        )
    if not response_succeeded(handle_response):
        return _finish_claimed(
            state_path,
            "MANUAL_REVIEW",
            external_guid,
            message="handle_object_rejected",
        )
    object_id = extract_created_object_id(handle_response)
    if not object_id:
        return _finish_claimed(
            state_path,
            "SUBMISSION_UNKNOWN",
            external_guid,
            message="created_object_id_missing",
        )
    try:
        _atomic_write_json(state_path, _state_record(
            phase="HANDLE_OBJECT_SUCCEEDED",
            external_guid=external_guid,
            object_id=object_id,
        ))
    except OSError:
        return GisApiOutcome(
            "SUBMISSION_UNKNOWN",
            object_id=object_id,
            external_guid=external_guid,
            state_path=str(state_path),
            message="post_create_state_failed",
        )

    # Barrier 2: registration.
    try:
        _atomic_write_json(state_path, _state_record(
            phase="BEFORE_REGISTRATION",
            external_guid=external_guid,
            object_id=object_id,
        ))
    except OSError:
        return _finish_claimed(
            state_path,
            "MANUAL_REVIEW",
            external_guid,
            object_id=object_id,
            message="registration_barrier_failed",
        )
    _log(logger, "info", "GIS ASUD API registration submission started")
    try:
        registration_response = client.execute_action(
            object_id, config.registration_action
        )
    except AsudApiError:
        return _finish_claimed(
            state_path,
            "SUBMISSION_UNKNOWN",
            external_guid,
            object_id=object_id,
            message="registration_submission_unknown",
        )
    if not response_succeeded(registration_response):
        return _finish_claimed(
            state_path,
            "MANUAL_REVIEW",
            external_guid,
            object_id=object_id,
            message="registration_rejected",
        )

    registration_number = None
    registration_date = None
    try:
        object_response = client.get_object(object_id)
        registration_number = extract_registration_number(object_response)
        registration_date = extract_registration_date(object_response)
    except AsudApiError:
        _log(logger, "warning", "GIS ASUD API registration lookup unavailable")
    if not registration_number or not registration_date:
        # Registration was submitted and returned success, but the required
        # postcondition is absent.  Do not submit registration again and do not
        # advance to resolution without a confirmed number/date.
        return _finish_claimed(
            state_path,
            "SUBMISSION_UNKNOWN",
            external_guid,
            object_id=object_id,
            registration_number=registration_number,
            message="registration_postcondition_missing",
        )
    try:
        _atomic_write_json(state_path, _state_record(
            phase="REGISTRATION_SUCCEEDED",
            external_guid=external_guid,
            object_id=object_id,
            registration_number=registration_number,
        ))
    except OSError:
        return GisApiOutcome(
            "SUBMISSION_UNKNOWN",
            object_id=object_id,
            registration_number=registration_number,
            external_guid=external_guid,
            state_path=str(state_path),
            message="post_registration_state_failed",
        )

    if not config.resolution_action:
        return _finish_claimed(
            state_path,
            "REGISTERED_ONLY",
            external_guid,
            object_id=object_id,
            registration_number=registration_number,
            message="resolution_action_not_configured",
        )

    # Barrier 3: send on resolution.
    try:
        _atomic_write_json(state_path, _state_record(
            phase="BEFORE_RESOLUTION",
            external_guid=external_guid,
            object_id=object_id,
            registration_number=registration_number,
        ))
    except OSError:
        return _finish_claimed(
            state_path,
            "REGISTERED_ONLY",
            external_guid,
            object_id=object_id,
            registration_number=registration_number,
            message="resolution_barrier_failed",
        )
    _log(logger, "info", "GIS ASUD API resolution submission started")
    try:
        resolution_response = client.execute_action(
            object_id, config.resolution_action
        )
    except AsudApiError:
        return _finish_claimed(
            state_path,
            "SUBMISSION_UNKNOWN",
            external_guid,
            object_id=object_id,
            registration_number=registration_number,
            message="resolution_submission_unknown",
        )
    if not response_succeeded(resolution_response):
        return _finish_claimed(
            state_path,
            "REGISTERED_ONLY",
            external_guid,
            object_id=object_id,
            registration_number=registration_number,
            message="resolution_rejected",
        )
    return _finish_claimed(
        state_path,
        "OK",
        external_guid,
        object_id=object_id,
        registration_number=registration_number,
        message="registered_and_sent_on_resolution",
    )


__all__ = [
    "GisApiOutcome",
    "build_exact_correspondent_query",
    "build_exact_appointment_query",
    "build_incoming_payload",
    "extract_created_object_id",
    "extract_exact_correspondent_id",
    "extract_exact_appointment_id",
    "extract_registration_date",
    "extract_registration_number",
    "extract_strict_fio",
    "gis_external_guid",
    "process_gis_document",
    "should_use_gis_api",
]
