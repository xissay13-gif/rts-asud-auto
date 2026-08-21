import hashlib
import json
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from flows.gis_api import (
    build_incoming_payload,
    extract_exact_correspondent_id,
    gis_external_guid,
    process_gis_document,
    should_use_gis_api,
)
from shared.asud_api import (
    AsudApiClient,
    AsudApiConfig,
    AsudApiConfigError,
    AsudApiHttpError,
    AsudApiTransportError,
)


FIO = "Иванов Иван Иванович"
ADDRESSEE_FIO = "Басманов Александр Владимирович"
BRANCH_NAME = "Тестовый филиал"


def _settings(*, mode="live-one", resolution_action="send_on_resolution"):
    return {
        "asud_api": {
            "enabled": True,
            "mode": mode,
            "allow_mutations": mode == "live-one",
            "max_documents": 1,
            "base_url": "https://asud-api.test/",
            "endpoints": {
                "find_objects2": "findObjects2",
                "find_appointment": "getOshsAppontmentByCriteria",
                "handle_object": "handleObject",
                "get_object": "getObject",
                "execute_action": "executeAction",
            },
            "timeout_sec": 3,
            "verify_tls": True,
            "lis": "TEST-LIS",
            "user": "api-test-user",
            "branch_id": "branch-1",
            "branch_name": BRANCH_NAME,
            "incoming_type_path": "Входящий документ/Письма",
            "addressee_id": "appointment-basmanov",
            "addressee_name": ADDRESSEE_FIO,
            "author_id": "appointment-author",
            "author_attribute": "authorId",
            "delivery_type": "Электронная почта",
            "registration_action": "registration",
            "resolution_action": resolution_action,
            "branch_attribute": "dsid_branch",
            "attachment": {
                "confirm_msg_supported": True,
                "max_bytes": 1024 * 1024,
            },
        }
    }


def _doc(msg_path):
    return {
        "файл": str(msg_path),
        "корр_источник": "zhkh",
        "корр_найден": True,
        "корреспондент": FIO,
        "корреспондент_тип": "person",
        "номер_обращения": "55-2026-12345",
        "дата_обращения": "21.08.2026",
        "планируемая_дата": "28.08.2026",
        "тема": "ГИС ЖКХ 55-2026-12345",
        "тема_обращения": "Другая тема",
        "содержание": "Проверочное обращение ГИС ЖКХ",
        "skip_asud_registration": False,
    }


def _attributes(**values):
    return {
        "attributes": {
            "attribute": [
                {"key": key, "value": value}
                for key, value in values.items()
            ]
        }
    }


def _correspondent_response(*object_ids):
    objects = []
    for object_id in object_ids:
        objects.append(_attributes(
            r_object_id=object_id,
            r_object_type="ddt_person",
            dss_relation_type="CORRESPONDENT_RELATION",
            dsid_branch="branch-1",
            dss_organisation="Иванов",
            dss_last_name="Иванов",
            dss_first_name="Иван",
            dss_middle_name="Иванович",
        ))
    return {"returnCode": "EA.200", "objects": objects}


def _appointment_candidate(
    *,
    object_id="appointment-basmanov",
    global_id=None,
    object_type="ddt_appointment",
    surname="Басманов",
    first_name="Александр",
    middle_name="Владимирович",
    branch_id="branch-1",
    branch_name=BRANCH_NAME,
    deleted="0",
):
    if global_id is None:
        global_id = object_id
    return {
        "type": object_type,
        "guid": None,
        "objectId": None,
        "docId": None,
        "attributes": [
            {"key": "r_object_id", "value": object_id},
            {"key": "dsid_global_id", "value": global_id},
            {"key": "dss_user_last_name", "value": surname},
            {"key": "dss_user_first_name", "value": first_name},
            {"key": "dss_user_middle_name", "value": middle_name},
            {"key": "dsid_branch", "value": branch_id},
            {"key": "branch_name", "value": branch_name},
            {"key": "dsb_deleted", "value": deleted},
        ],
        "relationObjects": [],
        "childObjects": [],
        "file": None,
        "fullCreate": False,
    }


def _appointment_response(*candidates):
    return {
        "returnCode": "EA.200",
        "returnMessage": "Операция выполнена успешно",
        "typedObjects": list(candidates),
    }


def _valid_appointment_response():
    return _appointment_response(_appointment_candidate())


def _get_object_response(*, number="АСУД/2026/1", date="2026-08-21T09:00:00"):
    return {
        "returnCode": "EA.200",
        "typedObject": _attributes(
            dss_reg_number=number,
            dsdt_reg_date=date,
        ),
    }


class ScriptedTransport:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def post_json(self, **kwargs):
        self.calls.append(kwargs)
        if not self.responses:
            raise AssertionError("unexpected API call")
        response = self.responses.pop(0)
        if isinstance(response, BaseException):
            raise response
        return response


class AsudApiContractTests(unittest.TestCase):
    def _assert_appointment_rejected_before_handle(self, response):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "appointment-rejected.msg"
            msg.write_bytes(b"message")
            transport = ScriptedTransport([
                _correspondent_response("person-1"),
                response,
            ])

            result = process_gis_document(
                _doc(msg), _settings(), transport=transport
            )

            self.assertEqual(result.status, "MANUAL_REVIEW")
            self.assertEqual(len(transport.calls), 2)
            self.assertTrue(
                transport.calls[1]["url"].endswith(
                    "getOshsAppontmentByCriteria"
                )
            )
            self.assertFalse(any(
                call["url"].endswith("handleObject")
                for call in transport.calls
            ))
            self.assertFalse(Path(str(msg) + ".asud_api.claim").exists())

    def test_api_scope_requires_structural_zhkh_marker(self):
        settings = _settings(mode="dry-run")
        suspicious = {
            "файл": "letter.msg",
            "тема": "ГИС ЖКХ 55-1",
            "тема_обращения": "Другая тема",
            "номер_обращения": "55-1",
        }

        self.assertFalse(should_use_gis_api(suspicious, settings))
        suspicious["корр_источник"] = "zhkh"
        self.assertTrue(should_use_gis_api(suspicious, settings))

    def test_external_guid_is_stable_and_branch_scoped(self):
        first = gis_external_guid(" 55-2026-12345 ", "BRANCH-1")
        second = gis_external_guid("55-2026-12345", "branch-1")

        self.assertEqual(first, second)
        self.assertNotEqual(
            first,
            gis_external_guid("55-2026-12345", "branch-2"),
        )

    def test_exact_correspondent_lookup_rejects_ambiguity(self):
        fio = ("Иванов", "Иван", "Иванович")

        self.assertEqual(
            extract_exact_correspondent_id(
                _correspondent_response("person-1"), fio, "branch-1"
            ),
            "person-1",
        )
        self.assertIsNone(extract_exact_correspondent_id(
            _correspondent_response("person-1", "person-2"),
            fio,
            "branch-1",
        ))

    def test_appointment_query_uses_exact_documented_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "appointment-probe.msg"
            msg.write_bytes(b"message")
            transport = ScriptedTransport([
                _correspondent_response("person-1"),
                _valid_appointment_response(),
            ])

            result = process_gis_document(
                _doc(msg), _settings(mode="probe"), transport=transport
            )

        self.assertEqual(result.status, "PROBE")
        self.assertEqual(len(transport.calls), 2)
        self.assertTrue(
            transport.calls[1]["url"].endswith(
                "getOshsAppontmentByCriteria"
            )
        )
        self.assertEqual(
            transport.calls[1]["payload"],
            {
                "criterion": [
                    {
                        "name": "dss_user_last_name",
                        "operation": "EQUALS",
                        "value1": "Басманов",
                    },
                    {
                        "name": "dss_user_first_name",
                        "operation": "EQUALS",
                        "value1": "Александр",
                    },
                    {
                        "name": "dss_user_middle_name",
                        "operation": "EQUALS",
                        "value1": "Владимирович",
                    },
                    {
                        "name": "ddt_branch.dss_name",
                        "operation": "EQUALS",
                        "value1": BRANCH_NAME,
                    },
                ]
            },
        )

    def test_appointment_mismatch_is_rejected_before_handle_object(self):
        cases = {
            "wrong_type": {"object_type": "ddt_person"},
            "wrong_surname": {"surname": "Петров"},
            "wrong_first_name": {"first_name": "Иван"},
            "wrong_middle_name": {"middle_name": "Иванович"},
            "wrong_branch_id": {"branch_id": "branch-2"},
            "wrong_branch_name": {"branch_name": "Другой филиал"},
            "wrong_configured_id": {"object_id": "appointment-other"},
        }
        for label, overrides in cases.items():
            with self.subTest(label=label):
                self._assert_appointment_rejected_before_handle(
                    _appointment_response(
                        _appointment_candidate(**overrides)
                    )
                )

    def test_ambiguous_appointments_are_rejected_before_handle_object(self):
        self._assert_appointment_rejected_before_handle(
            _appointment_response(
                _appointment_candidate(),
                _appointment_candidate(),
            )
        )

    def test_deleted_appointment_is_rejected_before_handle_object(self):
        self._assert_appointment_rejected_before_handle(
            _appointment_response(_appointment_candidate(deleted="1"))
        )

    def test_conflicting_appointment_ids_are_rejected_before_handle_object(self):
        self._assert_appointment_rejected_before_handle(
            _appointment_response(_appointment_candidate(
                object_id="appointment-basmanov",
                global_id="appointment-conflict",
            ))
        )

    def test_correspondent_query_uses_documented_find_objects2_criteria(self):
        from flows.gis_api import build_exact_correspondent_query

        config = AsudApiConfig.from_settings(_settings(), environ={})
        query = build_exact_correspondent_query(
            config, ("Иванов", "Иван", "Иванович")
        )
        criteria = {
            item["name"]: (item["operation"], item["value1"])
            for item in query["criterion"]
        }

        self.assertEqual(
            criteria["dss_organisation"], ("EQUALS", "Иванов")
        )
        self.assertEqual(
            criteria["dss_full_name"], ("STARTS_WITH", "Иванов")
        )
        self.assertEqual(
            criteria["dss_first_name"], ("STARTS_WITH", "Иван")
        )
        self.assertEqual(
            criteria["dss_middle_name"], ("STARTS_WITH", "Иванович")
        )
        self.assertNotIn("dss_last_name", criteria)

    def test_exact_correspondent_rejects_conflicting_duplicate_attributes(self):
        fio = ("Иванов", "Иван", "Иванович")
        pairs = [
            ("r_object_id", "person-1"),
            ("r_object_type", "ddt_person"),
            ("dss_relation_type", "CORRESPONDENT_RELATION"),
            ("dsid_branch", "branch-1"),
            ("dss_last_name", "Иванов"),
            ("dss_last_name", "Петров"),
            ("dss_first_name", "Иван"),
            ("dss_middle_name", "Иванович"),
        ]
        response = {
            "returnCode": "EA.200",
            "objects": [{
                "attributes": {
                    "attribute": [
                        {"key": key, "value": value}
                        for key, value in pairs
                    ]
                }
            }],
        }

        self.assertIsNone(extract_exact_correspondent_id(
            response, fio, "branch-1"
        ))

    def test_payload_contains_exact_attachment_bytes_and_md5(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "binary.msg"
            content = b"\x00\x01GIS\xff"
            msg.write_bytes(content)
            doc = _doc(msg)
            config = AsudApiConfig.from_settings(_settings())
            guid = gis_external_guid(doc["номер_обращения"], config.branch_id)

            payload = build_incoming_payload(
                doc, config, "person-1", guid
            )

        typed = payload["typedObject"]
        self.assertEqual(typed["guid"], guid)
        self.assertTrue(typed["fullCreate"])
        self.assertEqual(typed["file"]["name"], "binary.msg")
        self.assertEqual(typed["file"]["content"], list(content))
        self.assertEqual(
            typed["file"]["checkSum"], hashlib.md5(content).hexdigest()
        )
        attributes = {
            item["key"]: item["value"] for item in typed["attributes"]
        }
        attribute_keys = set(attributes)
        self.assertNotIn("dss_reg_number", attribute_keys)
        self.assertNotIn("dsdt_reg_date", attribute_keys)
        self.assertEqual(attributes["addresseesIds"], "appointment-basmanov")
        self.assertEqual(attributes["deliveryType"], "Электронная почта")
        self.assertEqual(attributes["dsid_branch"], "branch-1")
        self.assertEqual(attributes["authorId"], "appointment-author")

    def test_payload_rejects_duplicate_or_core_identity_attribute_names(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "identity-collision.msg"
            msg.write_bytes(b"msg")
            doc = _doc(msg)
            settings = _settings()
            settings["asud_api"]["author_attribute"] = "dsid_branch"
            config = AsudApiConfig.from_settings(settings, environ={})
            guid = gis_external_guid(doc["номер_обращения"], config.branch_id)

            with self.assertRaisesRegex(
                AsudApiConfigError, "distinct non-core fields"
            ):
                build_incoming_payload(doc, config, "person-1", guid)

            settings = _settings()
            settings["asud_api"]["branch_attribute"] = "description"
            config = AsudApiConfig.from_settings(settings, environ={})
            with self.assertRaisesRegex(
                AsudApiConfigError, "distinct non-core fields"
            ):
                build_incoming_payload(doc, config, "person-1", guid)

    def test_dry_run_never_calls_transport_or_creates_claim(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "dry.msg"
            msg.write_bytes(b"msg")
            transport = ScriptedTransport([])

            result = process_gis_document(
                _doc(msg), _settings(mode="dry-run"), transport=transport
            )

            self.assertEqual(result.status, "DRY_RUN")
            self.assertEqual(transport.calls, [])
            self.assertFalse(Path(str(msg) + ".asud_api.claim").exists())
            self.assertFalse(
                Path(str(msg) + ".asud_api_state.json").exists()
            )

    def test_live_success_registers_and_sends_resolution_once(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "live.msg"
            msg.write_bytes(b"message")
            transport = ScriptedTransport([
                _correspondent_response("person-1"),
                _valid_appointment_response(),
                {"returnCode": "EA.200", "doc": {"objectId": "doc-1"}},
                {"returnCode": "EA.200"},
                _get_object_response(),
                {"returnCode": "EA.200"},
            ])

            result = process_gis_document(
                _doc(msg), _settings(), transport=transport
            )

            self.assertEqual(result.status, "OK")
            self.assertEqual(result.object_id, "doc-1")
            self.assertEqual(result.registration_number, "АСУД/2026/1")
            self.assertEqual(len(transport.calls), 6)
            self.assertTrue(
                transport.calls[1]["url"].endswith(
                    "getOshsAppontmentByCriteria"
                )
            )
            self.assertTrue(transport.calls[2]["url"].endswith("handleObject"))
            actions = [
                call["payload"].get("action")
                for call in transport.calls
                if call["url"].endswith("executeAction")
            ]
            self.assertEqual(actions, ["registration", "send_on_resolution"])
            self.assertTrue(Path(str(msg) + ".asud_api.claim").exists())
            state_path = Path(str(msg) + ".asud_api_state.json")
            with state_path.open(encoding="utf-8") as stream:
                state = json.load(stream)
            self.assertEqual(state["status"], "OK")
            self.assertNotIn(FIO, json.dumps(state, ensure_ascii=False))

    def test_handle_timeout_is_terminal_and_never_submitted_twice(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "timeout.msg"
            msg.write_bytes(b"message")
            first_transport = ScriptedTransport([
                _correspondent_response("person-1"),
                _valid_appointment_response(),
                AsudApiTransportError("timeout"),
            ])
            doc = _doc(msg)

            first = process_gis_document(
                doc, _settings(), transport=first_transport
            )
            second_transport = ScriptedTransport([])
            second = process_gis_document(
                _doc(msg), _settings(), transport=second_transport
            )

            self.assertEqual(first.status, "SUBMISSION_UNKNOWN")
            self.assertEqual(second.status, "SUBMISSION_UNKNOWN")
            self.assertEqual(len(first_transport.calls), 3)
            self.assertEqual(second_transport.calls, [])

    def test_registration_without_strong_get_object_confirmation_stops(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "unconfirmed-registration.msg"
            msg.write_bytes(b"message")
            transport = ScriptedTransport([
                _correspondent_response("person-1"),
                _valid_appointment_response(),
                {"returnCode": "EA.200", "doc": {"objectId": "doc-1"}},
                {"returnCode": "EA.200"},
                _get_object_response(number="", date=""),
            ])

            result = process_gis_document(
                _doc(msg), _settings(), transport=transport
            )

            self.assertEqual(result.status, "SUBMISSION_UNKNOWN")
            self.assertEqual(len(transport.calls), 5)
            self.assertNotIn(
                "send_on_resolution",
                [call["payload"].get("action") for call in transport.calls],
            )

    def test_live_mode_requires_explicit_msg_support_and_size_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "guarded.msg"
            msg.write_bytes(b"message")
            settings = _settings()
            settings["asud_api"]["attachment"] = {
                "confirm_msg_supported": False,
                "max_bytes": 1,
            }
            transport = ScriptedTransport([
                _correspondent_response("person-1"),
                _valid_appointment_response(),
            ])

            result = process_gis_document(
                _doc(msg), settings, transport=transport
            )

            self.assertEqual(result.status, "MANUAL_REVIEW")
            # Read-only correspondent and addressee lookups are allowed, but
            # handleObject is not.
            self.assertEqual(len(transport.calls), 2)
            self.assertTrue(transport.calls[0]["url"].endswith("findObjects2"))
            self.assertTrue(
                transport.calls[1]["url"].endswith(
                    "getOshsAppontmentByCriteria"
                )
            )
            self.assertFalse(Path(str(msg) + ".asud_api.claim").exists())

    def test_definite_handle_rejection_is_not_reported_as_retryable_failed(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "rejected-create.msg"
            msg.write_bytes(b"message")
            transport = ScriptedTransport([
                _correspondent_response("person-1"),
                _valid_appointment_response(),
                {"returnCode": "EA.500", "returnMessage": "rejected"},
            ])

            result = process_gis_document(
                _doc(msg), _settings(), transport=transport
            )

            self.assertEqual(result.status, "MANUAL_REVIEW")
            self.assertEqual(len(transport.calls), 3)
            self.assertTrue(Path(str(msg) + ".asud_api.claim").exists())

    def test_registration_only_never_claims_full_ok(self):
        with tempfile.TemporaryDirectory() as tmp:
            msg = Path(tmp) / "registration-only.msg"
            msg.write_bytes(b"message")
            transport = ScriptedTransport([
                _correspondent_response("person-1"),
                _valid_appointment_response(),
                {"returnCode": "EA.200", "doc": {"objectId": "doc-1"}},
                {"returnCode": "EA.200"},
                _get_object_response(),
            ])

            result = process_gis_document(
                _doc(msg),
                _settings(resolution_action=""),
                transport=transport,
            )

            self.assertEqual(result.status, "REGISTERED_ONLY")
            self.assertEqual(result.registration_number, "АСУД/2026/1")
            self.assertEqual(len(transport.calls), 5)

    def test_settings_file_must_not_contain_authentication_secrets(self):
        settings = _settings(mode="dry-run")
        settings["asud_api"]["password"] = "must-not-live-in-json"

        with self.assertRaises(AsudApiConfigError):
            AsudApiConfig.from_settings(settings, environ={})

    def test_probe_and_live_require_verified_https(self):
        insecure = _settings()
        insecure["asud_api"]["base_url"] = "http://asud-api.test/"
        config = AsudApiConfig.from_settings(insecure, environ={})
        with self.assertRaisesRegex(AsudApiConfigError, "HTTPS"):
            config.validate_live()

        unverified = _settings(mode="probe")
        unverified["asud_api"]["verify_tls"] = False
        config = AsudApiConfig.from_settings(unverified, environ={})
        with self.assertRaisesRegex(
            AsudApiConfigError, "certificate verification"
        ):
            config.validate_probe()

    def test_live_endpoints_must_share_one_origin(self):
        settings = _settings()
        settings["asud_api"]["endpoints"]["execute_action"] = (
            "https://other-api.test/executeAction"
        )
        config = AsudApiConfig.from_settings(settings, environ={})

        with self.assertRaisesRegex(AsudApiConfigError, "one HTTPS origin"):
            config.validate_live()

    def test_transport_blocks_redirect_without_forwarding_authorization(self):
        hits = []
        authorization = []

        class RedirectHandler(BaseHTTPRequestHandler):
            def do_POST(self):
                length = int(self.headers.get("Content-Length", "0"))
                self.rfile.read(length)
                hits.append(self.path)
                authorization.append(self.headers.get("Authorization"))
                if self.path == "/start":
                    self.send_response(302)
                    self.send_header(
                        "Location",
                        f"http://127.0.0.1:{self.server.server_port}/sink",
                    )
                    self.end_headers()
                else:
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(b'{"returnCode":"EA.200"}')

            def log_message(self, _format, *args):
                pass

        server = ThreadingHTTPServer(("127.0.0.1", 0), RedirectHandler)
        worker = threading.Thread(target=server.serve_forever, daemon=True)
        worker.start()
        try:
            config = AsudApiConfig(
                endpoints={
                    "find_objects2": (
                        f"http://127.0.0.1:{server.server_port}/start"
                    )
                },
                timeout_sec=2,
                auth_type="basic",
                auth_username="test-user",
                auth_password="test-password",
            )
            client = AsudApiClient(config)

            with self.assertRaisesRegex(AsudApiHttpError, "302"):
                client.find_objects2({"test": True})
        finally:
            server.shutdown()
            server.server_close()
            worker.join(timeout=2)

        self.assertEqual(hits, ["/start"])
        self.assertEqual(len(authorization), 1)
        self.assertTrue(authorization[0].startswith("Basic "))


if __name__ == "__main__":
    unittest.main()
