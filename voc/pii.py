"""Personal-data handling.

Two layers:
- Column level: identifying columns are dropped at ingest, before any analysis,
  caching, or model call. Spend is kept only as a coarse band.
- Text level: emails, phone numbers, URLs, social handles, and order numbers are
  masked inside free-text answers, because customers type them in anyway.

Names typed inside free text cannot be caught reliably with rules; reports
therefore never print verbatims unless `--quotes` is passed for internal use.
"""
from __future__ import annotations

import re
from typing import Iterable

from .textfix import fold

# Column headers (accent-folded, lowercase) that identify a person or a device.
PII_COLUMN_PATTERNS = [
    r"^(full |first |last )?name$", r"^nombre", r"apellido", r"^reviewer",
    r"e-?mail", r"correo", r"phone", r"telefono", r"celular",
    r"^city$", r"ciudad", r"province", r"estado", r"address", r"direccion",
    r"zip", r"postal", r"ip ?address", r"user ?agent", r"customer ?id",
    r"^id$", r"order ?(id|number)", r"referrer", r"landing", r"^url$", r"^href$",
    r"survey metadata",
]
_PII_COL_RE = re.compile("|".join(PII_COLUMN_PATTERNS))

SPEND_COLUMN_RE = re.compile(r"total ?spent|gasto")

_TEXT_RULES = [
    (re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"), "[email]"),
    (re.compile(r"https?://\S+|www\.\S+"), "[url]"),
    (re.compile(r"(?<!\w)@[A-Za-z0-9_.]{2,}"), "[handle]"),
    (re.compile(r"#\s?\d{4,}"), "[order]"),
    (re.compile(r"(?<!\d)(?:\+?\d[\s-]?){10,13}(?!\d)"), "[phone]"),
]


def is_pii_column(header: str) -> bool:
    return bool(_PII_COL_RE.search(fold(str(header)).strip()))


def pii_columns(headers: Iterable[str]) -> list[str]:
    return [h for h in headers if is_pii_column(h)]


def scrub_text(text: str) -> str:
    for pattern, token in _TEXT_RULES:
        text = pattern.sub(token, text)
    return text


def spend_band(value: object) -> str:
    """Bucket a spend amount so no exact customer value leaves ingest."""
    try:
        amount = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return "unknown"
    if amount != amount:  # NaN
        return "unknown"
    for upper, label in ((500, "<500"), (1000, "500-999"), (2500, "1000-2499"), (5000, "2500-4999")):
        if amount < upper:
            return label
    return "5000+"
