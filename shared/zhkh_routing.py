"""Business policy for GIS ЖКХ topics.

The GIS exporter currently stores only the human-readable ``appealName`` in
the generated MSG.  Topic numbers from the internal routing table (``2.1``,
``5.2`` and so on) are therefore kept here as audit identifiers while all
matching is performed against the exact normalized title.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Mapping


EXCLUDED_TOPIC_TITLES = {
    "2.1": "Отсутствие отопления",
    "2.2": "Отсутствие водоснабжения",
    "2.3": "Нарушение температурного режима подачи воды",
    "2.16": "Порыв труб",
    "2.25": "Затопление подвала",
    "5.1": "Качество оказания коммунальных услуг",
    "5.2": "Сроки оказания коммунальных услуг",
    "23": "Некачественная поставка ресурса",
}

# These topics must not be registered in ASUD at all.  The downloader may also
# filter them upstream, but this built-in guard is authoritative in case an MSG
# still reaches the monitored folder.
DEFAULT_ZHKH_EXCLUDED_TOPICS = dict(EXCLUDED_TOPIC_TITLES)

def normalize_topic_title(value) -> str:
    """Normalize a GIS topic title without weakening exact matching."""
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = text.replace("ё", "е").replace("Ё", "Е").casefold()
    text = re.sub(r"\s+", " ", text).strip()
    # A terminal dot is presentation punctuation, not part of the topic name.
    return text.rstrip(".").strip()


def match_excluded_topic(topic_title, additional_topics=None):
    """Return the topic code when an exact title must be skipped.

    The eight built-in exclusions are mandatory and cannot be disabled by an
    empty user config. ``additional_topics`` may only add exclusions. ``None``
    means no match. Malformed configuration and duplicate normalized titles
    fail closed with ``ValueError`` rather than allowing a potentially excluded
    document to reach ASUD.
    """
    normalized = normalize_topic_title(topic_title)
    if not normalized:
        return None

    def matching_codes(source):
        matches = []
        for raw_code, raw_title in source.items():
            if not isinstance(raw_title, str) or not raw_title.strip():
                raise ValueError(
                    f"zhkh_excluded_topics[{raw_code!r}] must be a non-empty string"
                )
            if normalize_topic_title(raw_title) == normalized:
                matches.append(str(raw_code))
        return matches

    # Mandatory built-ins always win, even if a stale RC9 user config contains
    # an unrelated/malformed routing section or an empty exclusion override.
    matches = matching_codes(DEFAULT_ZHKH_EXCLUDED_TOPICS)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:  # defensive: the shipped table must stay unambiguous
        raise ValueError(
            "GIS ЖКХ topic matches multiple exclusion entries: "
            + ", ".join(matches)
        )

    if additional_topics is None:
        return None
    if not isinstance(additional_topics, Mapping):
        raise ValueError("zhkh_excluded_topics must be an object")
    matches = matching_codes(additional_topics)
    if len(matches) > 1:
        raise ValueError(
            "GIS ЖКХ topic matches multiple additional exclusion entries: "
            + ", ".join(matches)
        )
    return matches[0] if matches else None
