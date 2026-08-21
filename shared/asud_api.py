"""Small, fail-closed REST/JSON client for the experimental ASUD API flow.

The production application historically talks to ASUD through Selenium.  This
module deliberately does not know any ASUD URL layout: every endpoint must be
provided as either a full URL or an explicit path paired with ``base_url``.
Authentication material is accepted from environment variables only.

Only Python's standard library is imported so the module remains usable in the
single-file Windows build.
"""

from __future__ import annotations

import base64
import json
import os
import socket
import ssl
from dataclasses import dataclass, field
from typing import Any, Dict, Mapping, Optional
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest


OPERATIONS = (
    "find_objects2",
    "find_appointment",
    "handle_object",
    "get_object",
    "execute_action",
)
SUCCESS_CODE = "EA.200"


class AsudApiError(RuntimeError):
    """Base exception for the experimental API client."""


class AsudApiConfigError(AsudApiError):
    """Configuration is missing or unsafe."""


class AsudApiHttpError(AsudApiError):
    """The server returned a definite non-2xx HTTP response."""

    def __init__(self, status: int):
        super().__init__(f"ASUD API HTTP status {status}")
        self.status = status


class AsudApiUncertainError(AsudApiError):
    """A request may have reached the server but no usable response exists."""


class AsudApiTransportError(AsudApiUncertainError):
    """Network/transport failure with uncertain delivery."""


class AsudApiResponseError(AsudApiUncertainError):
    """A successful HTTP response could not be interpreted as JSON."""


def _bool_value(value: Any, *, name: str) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, int) and value in (0, 1):
        return bool(value)
    if isinstance(value, str):
        normalized = value.strip().casefold()
        if normalized in ("1", "true", "yes", "on", "да"):
            return True
        if normalized in ("0", "false", "no", "off", "нет", ""):
            return False
    raise AsudApiConfigError(f"{name} must be a boolean")


def _int_value(value: Any, *, name: str, minimum: int = 0) -> int:
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise AsudApiConfigError(f"{name} must be an integer") from exc
    if result < minimum:
        raise AsudApiConfigError(f"{name} must be >= {minimum}")
    return result


def _float_value(value: Any, *, name: str, minimum: float = 0.1) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise AsudApiConfigError(f"{name} must be a number") from exc
    if result < minimum:
        raise AsudApiConfigError(f"{name} must be >= {minimum}")
    return result


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


@dataclass(frozen=True)
class AsudApiConfig:
    """Validated configuration assembled from ``settings['asud_api']``.

    Environment variables override non-secret settings.  Credentials and
    authorization headers have *no* settings-file equivalent by design.
    """

    enabled: bool = False
    mode: str = "dry-run"
    allow_mutations: bool = False
    max_documents: int = 0
    base_url: str = ""
    endpoints: Mapping[str, str] = field(default_factory=dict)
    timeout_sec: float = 30.0
    verify_tls: bool = True
    lis: str = ""
    user: str = ""
    branch_id: str = ""
    branch_name: str = ""
    incoming_type_path: str = ""
    addressee_id: str = ""
    addressee_name: str = ""
    author_id: str = ""
    author_attribute: str = ""
    delivery_type: str = ""
    registration_action: str = ""
    resolution_action: str = ""
    incoming_attributes: Mapping[str, Any] = field(default_factory=dict)
    branch_attribute: str = ""
    confirm_msg_supported: bool = False
    max_attachment_bytes: int = 0
    auth_type: str = "none"
    auth_username: str = field(default="", repr=False)
    auth_password: str = field(default="", repr=False)
    bearer_token: str = field(default="", repr=False)
    auth_header_name: str = field(default="", repr=False)
    auth_header_value: str = field(default="", repr=False)

    @classmethod
    def from_settings(
        cls,
        settings: Mapping[str, Any],
        environ: Optional[Mapping[str, str]] = None,
    ) -> "AsudApiConfig":
        if not isinstance(settings, Mapping):
            raise AsudApiConfigError("settings must be an object")
        raw = settings.get("asud_api", {})
        if raw is None:
            raw = {}
        if not isinstance(raw, Mapping):
            raise AsudApiConfigError("settings.asud_api must be an object")
        env = os.environ if environ is None else environ

        # Secret-bearing values in a JSON settings file are an easy accidental
        # leak into Git and logs.  Reject them instead of silently using them.
        forbidden = {
            "auth", "auth_type", "username", "password", "token",
            "bearer_token", "authorization", "authorization_header",
            "auth_header_name", "auth_header_value", "secret",
        }
        if any(str(key).casefold() in forbidden for key in raw):
            raise AsudApiConfigError(
                "ASUD API authentication is allowed through environment only"
            )

        def pick(key: str, env_name: str, default: Any = "") -> Any:
            if env_name in env:
                return env[env_name]
            return raw.get(key, default)

        endpoint_settings = raw.get("endpoints", {})
        if endpoint_settings is None:
            endpoint_settings = {}
        if not isinstance(endpoint_settings, Mapping):
            raise AsudApiConfigError("asud_api.endpoints must be an object")
        endpoint_env = {
            "find_objects2": "ASUD_API_FIND_OBJECTS2_URL",
            "find_appointment": "ASUD_API_FIND_APPOINTMENT_URL",
            "handle_object": "ASUD_API_HANDLE_OBJECT_URL",
            "get_object": "ASUD_API_GET_OBJECT_URL",
            "execute_action": "ASUD_API_EXECUTE_ACTION_URL",
        }
        endpoints: Dict[str, str] = {}
        for operation, env_name in endpoint_env.items():
            if env_name in env:
                value = env[env_name]
            else:
                value = endpoint_settings.get(
                    operation, raw.get(f"{operation}_url", "")
                )
            endpoints[operation] = _text(value)

        extra_attributes = raw.get("incoming_attributes", {})
        if extra_attributes is None:
            extra_attributes = {}
        if not isinstance(extra_attributes, Mapping):
            raise AsudApiConfigError(
                "asud_api.incoming_attributes must be an object"
            )
        extra_attributes = {
            _text(key): value for key, value in extra_attributes.items()
            if _text(key) and value is not None
        }

        attachment = raw.get("attachment", {})
        if attachment is None:
            attachment = {}
        if not isinstance(attachment, Mapping):
            raise AsudApiConfigError("asud_api.attachment must be an object")

        def pick_attachment(key: str, env_name: str, default: Any) -> Any:
            if env_name in env:
                return env[env_name]
            if key in attachment:
                return attachment[key]
            # Flat aliases keep environment-generated configs easy to consume,
            # while the documented settings format remains nested.
            return raw.get(key, default)

        enabled = _bool_value(
            pick("enabled", "ASUD_API_ENABLED", False),
            name="asud_api.enabled",
        )
        allow_mutations = _bool_value(
            pick("allow_mutations", "ASUD_API_ALLOW_MUTATIONS", False),
            name="asud_api.allow_mutations",
        )
        verify_tls = _bool_value(
            pick("verify_tls", "ASUD_API_VERIFY_TLS", True),
            name="asud_api.verify_tls",
        )
        mode = _text(pick("mode", "ASUD_API_MODE", "dry-run")).casefold()
        if mode not in ("dry-run", "probe", "live-one"):
            raise AsudApiConfigError(
                "asud_api.mode must be dry-run, probe, or live-one"
            )

        auth_type = _text(env.get("ASUD_API_AUTH_TYPE", "none")).casefold()
        if auth_type not in ("none", "basic", "bearer", "custom"):
            raise AsudApiConfigError(
                "ASUD_API_AUTH_TYPE must be none, basic, bearer, or custom"
            )

        config = cls(
            enabled=enabled,
            mode=mode,
            allow_mutations=allow_mutations,
            max_documents=_int_value(
                pick("max_documents", "ASUD_API_MAX_DOCUMENTS", 0),
                name="asud_api.max_documents",
            ),
            base_url=_text(pick("base_url", "ASUD_API_BASE_URL", "")),
            endpoints=endpoints,
            timeout_sec=_float_value(
                pick("timeout_sec", "ASUD_API_TIMEOUT_SEC", 30),
                name="asud_api.timeout_sec",
            ),
            verify_tls=verify_tls,
            lis=_text(pick("lis", "ASUD_API_LIS", "")),
            user=_text(pick("user", "ASUD_API_USER", "")),
            branch_id=_text(pick("branch_id", "ASUD_API_BRANCH_ID", "")),
            branch_name=_text(pick(
                "branch_name", "ASUD_API_BRANCH_NAME", ""
            )),
            incoming_type_path=_text(pick(
                "incoming_type_path", "ASUD_API_INCOMING_TYPE_PATH", ""
            )),
            addressee_id=_text(pick(
                "addressee_id", "ASUD_API_ADDRESSEE_ID", ""
            )),
            addressee_name=_text(pick(
                "addressee_name", "ASUD_API_ADDRESSEE_NAME", ""
            )),
            author_id=_text(pick(
                "author_id", "ASUD_API_AUTHOR_ID", ""
            )),
            author_attribute=_text(pick(
                "author_attribute", "ASUD_API_AUTHOR_ATTRIBUTE", ""
            )),
            delivery_type=_text(pick(
                "delivery_type", "ASUD_API_DELIVERY_TYPE", ""
            )),
            registration_action=_text(pick(
                "registration_action", "ASUD_API_REGISTRATION_ACTION", ""
            )),
            resolution_action=_text(pick(
                "resolution_action", "ASUD_API_RESOLUTION_ACTION", ""
            )),
            incoming_attributes=extra_attributes,
            branch_attribute=_text(pick(
                "branch_attribute", "ASUD_API_BRANCH_ATTRIBUTE", ""
            )),
            confirm_msg_supported=_bool_value(
                pick_attachment(
                    "confirm_msg_supported",
                    "ASUD_API_CONFIRM_MSG_SUPPORTED",
                    False,
                ),
                name="asud_api.attachment.confirm_msg_supported",
            ),
            max_attachment_bytes=_int_value(
                pick_attachment(
                    "max_bytes", "ASUD_API_MAX_ATTACHMENT_BYTES", 0
                ),
                name="asud_api.attachment.max_bytes",
            ),
            auth_type=auth_type,
            auth_username=_text(env.get("ASUD_API_BASIC_USERNAME", "")),
            auth_password=_text(env.get("ASUD_API_BASIC_PASSWORD", "")),
            bearer_token=_text(env.get("ASUD_API_BEARER_TOKEN", "")),
            auth_header_name=_text(env.get("ASUD_API_AUTH_HEADER_NAME", "")),
            auth_header_value=_text(env.get("ASUD_API_AUTH_HEADER_VALUE", "")),
        )
        config._validate_urls_that_are_present()
        config._validate_auth()
        return config

    @property
    def mutations_enabled(self) -> bool:
        return (
            self.enabled
            and self.allow_mutations
            and self.mode == "live-one"
            and self.max_documents == 1
        )

    def endpoint(self, operation: str) -> str:
        if operation not in OPERATIONS:
            raise AsudApiConfigError(f"Unknown ASUD API operation: {operation}")
        configured = _text(self.endpoints.get(operation, ""))
        if not configured:
            raise AsudApiConfigError(
                f"Explicit ASUD API endpoint is required for {operation}"
            )
        parsed = urlparse.urlsplit(configured)
        if parsed.scheme:
            if parsed.scheme.casefold() not in ("http", "https") or not parsed.netloc:
                raise AsudApiConfigError(
                    f"Invalid ASUD API endpoint for {operation}"
                )
            if parsed.username or parsed.password or parsed.fragment:
                raise AsudApiConfigError(
                    f"Credentials/fragments are forbidden in endpoint {operation}"
                )
            return configured
        if not self.base_url:
            raise AsudApiConfigError(
                f"Relative endpoint for {operation} requires base_url"
            )
        base = urlparse.urlsplit(self.base_url)
        if base.scheme.casefold() not in ("http", "https") or not base.netloc:
            raise AsudApiConfigError("Invalid ASUD API base_url")
        resolved = urlparse.urljoin(
            self.base_url.rstrip("/") + "/", configured
        )
        parsed_resolved = urlparse.urlsplit(resolved)
        if (parsed_resolved.username or parsed_resolved.password
                or parsed_resolved.fragment):
            raise AsudApiConfigError(
                f"Credentials/fragments are forbidden in endpoint {operation}"
            )
        return resolved

    def _validate_secure_network(
        self,
        operations: tuple[str, ...],
        *,
        same_origin: bool,
    ) -> None:
        if not self.verify_tls:
            raise AsudApiConfigError(
                "ASUD API probe/live requires TLS certificate verification"
            )
        origins = set()
        for operation in operations:
            endpoint = urlparse.urlsplit(self.endpoint(operation))
            if endpoint.scheme.casefold() != "https":
                raise AsudApiConfigError(
                    f"ASUD API {operation} endpoint must use HTTPS"
                )
            origins.add((
                endpoint.scheme.casefold(),
                (endpoint.hostname or "").casefold(),
                endpoint.port or 443,
            ))
        if same_origin and len(origins) != 1:
            raise AsudApiConfigError(
                "All mutating ASUD API endpoints must use one HTTPS origin"
            )

    def validate_probe(self) -> None:
        if not self.enabled:
            raise AsudApiConfigError("ASUD API is disabled")
        if self.mode not in ("probe", "live-one"):
            raise AsudApiConfigError("Probe requires probe or live-one mode")
        required = {
            "user": self.user,
            "branch_id": self.branch_id,
            "branch_name": self.branch_name,
            "addressee_id": self.addressee_id,
            "addressee_name": self.addressee_name,
        }
        missing = [key for key, value in required.items() if not value]
        if missing:
            raise AsudApiConfigError(
                "Missing ASUD API probe settings: " + ", ".join(missing)
            )
        self._validate_secure_network(
            ("find_objects2", "find_appointment"), same_origin=True
        )

    def validate_live(self) -> None:
        if not self.mutations_enabled:
            raise AsudApiConfigError(
                "Mutations require enabled=true, allow_mutations=true, "
                "mode=live-one, and max_documents=1"
            )
        required = {
            "lis": self.lis,
            "user": self.user,
            "branch_id": self.branch_id,
            "branch_name": self.branch_name,
            "incoming_type_path": self.incoming_type_path,
            "addressee_id": self.addressee_id,
            "addressee_name": self.addressee_name,
            "author_id": self.author_id,
            "author_attribute": self.author_attribute,
            "delivery_type": self.delivery_type,
            "branch_attribute": self.branch_attribute,
            "registration_action": self.registration_action,
        }
        missing = [key for key, value in required.items() if not value]
        if missing:
            raise AsudApiConfigError(
                "Missing ASUD API live settings: " + ", ".join(missing)
            )
        if not self.confirm_msg_supported or self.max_attachment_bytes <= 0:
            raise AsudApiConfigError(
                "Live MSG upload requires attachment.confirm_msg_supported=true "
                "and attachment.max_bytes > 0"
            )
        self._validate_secure_network(OPERATIONS, same_origin=True)

    def auth_headers(self) -> Dict[str, str]:
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
        }
        if self.auth_type == "basic":
            raw = f"{self.auth_username}:{self.auth_password}".encode("utf-8")
            headers["Authorization"] = (
                "Basic " + base64.b64encode(raw).decode("ascii")
            )
        elif self.auth_type == "bearer":
            headers["Authorization"] = "Bearer " + self.bearer_token
        elif self.auth_type == "custom":
            headers[self.auth_header_name] = self.auth_header_value
        return headers

    def _validate_urls_that_are_present(self) -> None:
        if self.base_url:
            parsed = urlparse.urlsplit(self.base_url)
            if parsed.scheme.casefold() not in ("http", "https") or not parsed.netloc:
                raise AsudApiConfigError("Invalid ASUD API base_url")
        for operation, configured in self.endpoints.items():
            if configured:
                # endpoint() also validates relative/base combinations.
                self.endpoint(operation)

    def _validate_auth(self) -> None:
        values = (
            self.auth_username,
            self.auth_password,
            self.bearer_token,
            self.auth_header_name,
            self.auth_header_value,
        )
        if any("\r" in value or "\n" in value for value in values):
            raise AsudApiConfigError("Newlines are not allowed in API auth values")
        if self.auth_type == "basic" and not (
            self.auth_username and self.auth_password
        ):
            raise AsudApiConfigError("Basic auth environment values are incomplete")
        if self.auth_type == "bearer" and not self.bearer_token:
            raise AsudApiConfigError("Bearer token environment value is missing")
        if self.auth_type == "custom" and not (
            self.auth_header_name and self.auth_header_value
        ):
            raise AsudApiConfigError("Custom auth environment values are incomplete")


class _UrllibJsonTransport:
    """One-shot JSON POST transport.  It intentionally contains no retry."""

    class _NoRedirectHandler(urlrequest.HTTPRedirectHandler):
        """Never forward API credentials/body to a redirect target."""

        def redirect_request(
            self, req, fp, code, msg, headers, newurl
        ):  # pragma: no cover - exercised through urllib's handler chain
            raise urlerror.HTTPError(
                req.full_url,
                code,
                "ASUD API redirect blocked",
                headers,
                fp,
            )

    def post_json(
        self,
        *,
        url: str,
        payload: Mapping[str, Any],
        headers: Mapping[str, str],
        timeout: float,
        verify_tls: bool,
    ) -> Mapping[str, Any]:
        body = json.dumps(
            payload, ensure_ascii=False, separators=(",", ":")
        ).encode("utf-8")
        req = urlrequest.Request(
            url, data=body, headers=dict(headers), method="POST"
        )
        context = (
            ssl.create_default_context()
            if verify_tls else ssl._create_unverified_context()
        )
        opener = urlrequest.build_opener(
            self._NoRedirectHandler(),
            urlrequest.HTTPSHandler(context=context),
        )
        try:
            with opener.open(req, timeout=timeout) as response:
                response_bytes = response.read()
        except urlerror.HTTPError as exc:
            # Do not include the response body: server diagnostics can echo
            # document data or authentication details.
            raise AsudApiHttpError(exc.code) from exc
        except (urlerror.URLError, TimeoutError, socket.timeout, OSError) as exc:
            raise AsudApiTransportError("ASUD API transport failed") from exc
        try:
            decoded = json.loads(response_bytes.decode("utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise AsudApiResponseError("ASUD API returned invalid JSON") from exc
        if not isinstance(decoded, Mapping):
            raise AsudApiResponseError("ASUD API JSON response is not an object")
        return decoded


class AsudApiClient:
    """REST/JSON client with injectable transport for tests.

    A custom transport is either an object with ``post_json`` or a callable;
    both receive keyword-only ``url``, ``payload``, ``headers``, ``timeout``
    and ``verify_tls``.  The callable is invoked exactly once.
    """

    def __init__(
        self,
        config: AsudApiConfig,
        transport: Optional[Any] = None,
        logger: Optional[Any] = None,
    ) -> None:
        self.config = config
        self.transport = transport or _UrllibJsonTransport()
        self.logger = logger

    def request(self, operation: str, payload: Mapping[str, Any]) -> Mapping[str, Any]:
        url = self.config.endpoint(operation)
        kwargs = {
            "url": url,
            "payload": payload,
            "headers": self.config.auth_headers(),
            "timeout": self.config.timeout_sec,
            "verify_tls": self.config.verify_tls,
        }
        try:
            if hasattr(self.transport, "post_json"):
                response = self.transport.post_json(**kwargs)
            elif callable(self.transport):
                response = self.transport(**kwargs)
            else:
                raise AsudApiConfigError(
                    "transport must be callable or provide post_json"
                )
        except AsudApiError:
            raise
        except Exception as exc:
            # A custom transport failure has the same delivery ambiguity as a
            # socket failure.  Do not call it again with a fallback signature.
            raise AsudApiTransportError("ASUD API transport failed") from exc
        if not isinstance(response, Mapping):
            raise AsudApiResponseError("ASUD API response is not an object")
        return response

    def find_objects2(self, payload: Mapping[str, Any]) -> Mapping[str, Any]:
        return self.request("find_objects2", payload)

    def find_appointment(
        self, payload: Mapping[str, Any]
    ) -> Mapping[str, Any]:
        return self.request("find_appointment", payload)

    def handle_object(self, payload: Mapping[str, Any]) -> Mapping[str, Any]:
        return self.request("handle_object", payload)

    def get_object(self, object_id: str) -> Mapping[str, Any]:
        return self.request("get_object", {
            "lis": self.config.lis,
            "user": self.config.user,
            "objectId": object_id,
        })

    def execute_action(self, object_id: str, action: str) -> Mapping[str, Any]:
        return self.request("execute_action", {
            "userLogin": self.config.user,
            "docOrTaskId": object_id,
            "targetObjectId": object_id,
            "action": action,
        })


def response_succeeded(response: Mapping[str, Any]) -> bool:
    return _text(response.get("returnCode")) == SUCCESS_CODE


__all__ = [
    "AsudApiClient",
    "AsudApiConfig",
    "AsudApiConfigError",
    "AsudApiError",
    "AsudApiHttpError",
    "AsudApiResponseError",
    "AsudApiTransportError",
    "AsudApiUncertainError",
    "response_succeeded",
]
