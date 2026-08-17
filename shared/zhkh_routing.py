"""Business routing for GIS ЖКХ document addressees.

The GIS exporter currently stores only the human-readable ``appealName`` in
the generated MSG.  Topic numbers from the internal routing table (``2.1``,
``5.2`` and so on) are therefore kept here as audit identifiers while matching
is performed against the exact normalized title.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Mapping, Sequence


ZHUKOV_ADDRESSEE = "Жуков Иван Сергеевич"

ZHUKOV_TOPIC_TITLES = {
    "2.1": "Отсутствие отопления",
    "2.2": "Отсутствие водоснабжения",
    "2.3": "Нарушение температурного режима подачи воды",
    "2.16": "Порыв труб",
    "2.25": "Затопление подвала",
    "5.1": "Качество оказания коммунальных услуг",
    "5.2": "Сроки оказания коммунальных услуг",
    "23": "Некачественная поставка ресурса",
}

DEFAULT_ZHKH_ADDRESSEE_ROUTES = [
    {
        "id": "zhukov_priority_topics",
        "topics": dict(ZHUKOV_TOPIC_TITLES),
        "addressees": [ZHUKOV_ADDRESSEE],
    }
]


def normalize_topic_title(value) -> str:
    """Normalize a GIS topic title without weakening exact matching."""
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = text.replace("ё", "е").replace("Ё", "Е").casefold()
    text = re.sub(r"\s+", " ", text).strip()
    # A terminal dot is presentation punctuation, not part of the topic name.
    return text.rstrip(".").strip()


def _clean_addressees(raw) -> list[str]:
    if isinstance(raw, str) or not isinstance(raw, Sequence):
        return []
    cleaned = []
    for index, value in enumerate(raw):
        if not isinstance(value, str) or not value.strip():
            raise ValueError(
                f"route addressees[{index}] must be a non-empty string"
            )
        name = value.strip()
        if name not in cleaned:
            cleaned.append(name)
    return cleaned


def resolve_addressee_override(topic_title, routes):
    """Return ``(addressees, topic_code)`` for an exact configured match.

    ``(None, None)`` means that no rule matched and the caller must retain the
    ordinary global addressee list.  A matched but malformed rule raises
    ``ValueError`` so a GIS document can never silently fall back to the wrong
    person.
    """
    normalized = normalize_topic_title(topic_title)
    if not normalized:
        return None, None

    if routes is None:
        routes = DEFAULT_ZHKH_ADDRESSEE_ROUTES
    if isinstance(routes, (str, bytes)) or not isinstance(routes, Sequence):
        raise ValueError("zhkh_addressee_routes must be a list")

    matches = []
    for rule_index, rule in enumerate(routes):
        if not isinstance(rule, Mapping):
            raise ValueError(
                f"zhkh_addressee_routes[{rule_index}] must be an object"
            )
        topics = rule.get("topics")
        if not isinstance(topics, Mapping):
            raise ValueError(
                f"zhkh_addressee_routes[{rule_index}].topics must be an object"
            )
        for raw_code, raw_title in topics.items():
            if normalize_topic_title(raw_title) == normalized:
                matches.append((rule_index, str(raw_code), rule))

    if not matches:
        return None, None
    if len(matches) > 1:
        codes = ", ".join(code for _, code, _ in matches)
        raise ValueError(
            f"GIS ЖКХ topic matches multiple addressee rules: {codes}"
        )

    rule_index, topic_code, rule = matches[0]
    addressees = _clean_addressees(rule.get("addressees"))
    if not addressees:
        raise ValueError(
            f"zhkh_addressee_routes[{rule_index}] matched topic "
            f"{topic_code}, but addressees is empty"
        )
    return addressees, topic_code
