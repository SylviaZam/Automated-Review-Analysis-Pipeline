"""Text repair for survey exports.

Shopify survey apps export CSVs in whatever encoding the analyst's spreadsheet
tool last saved them in. In practice we saw three failure modes:

1. UTF-8 files that open fine.
2. Files re-saved as Mac Roman / Windows-1252, where "¿Cómo" arrives as
   "ÀC\x97mo" when read as Latin-1.
3. UTF-8 text that was decoded twice ("Cu√©ntanos" instead of "Cuéntanos").

`decode_bytes` handles (1) and (2) at file level, `fix_cell` handles (3) at cell level.
"""
from __future__ import annotations

import re
import unicodedata

try:
    import ftfy  # type: ignore
except ImportError:  # pragma: no cover - ftfy is a hard dependency, but stay importable
    ftfy = None

_CANDIDATES = ("cp1252", "mac_roman", "latin-1")
_GOOD = set("áéíóúüñÁÉÍÓÚÜÑ¿¡")
_BAD_RE = re.compile(r"[\x80-\x9f]|[ÀÃÂ][^\s]|√|¬|Ž")


def _score(text: str) -> int:
    good = sum(ch in _GOOD for ch in text)
    bad = len(_BAD_RE.findall(text))
    return good - 3 * bad


def decode_bytes(raw: bytes) -> tuple[str, str]:
    """Return (text, encoding_used). Prefers UTF-8, otherwise picks the legacy
    encoding that produces the most plausible Spanish text."""
    try:
        return raw.decode("utf-8-sig"), "utf-8"
    except UnicodeDecodeError:
        pass
    best = max(_CANDIDATES, key=lambda enc: _score(raw.decode(enc, errors="replace")))
    return raw.decode(best, errors="replace"), best


# Mac Roman renderings of UTF-8 punctuation that survive ftfy when a cell mixes
# repaired and broken text.
_MAC_ROMAN_PAIRS = {
    "¬ø": "¿", "¬°": "¡", "√°": "á", "√©": "é", "√≠": "í", "√≥": "ó",
    "√∫": "ú", "√±": "ñ", "√ë": "Ñ", "√º": "ü", "√Å": "Á", "√â": "É",
}


def fix_cell(value: object) -> object:
    """Repair double-decoded UTF-8 and normalise whitespace in a single cell."""
    if not isinstance(value, str):
        return value
    text = ftfy.fix_text(value) if ftfy is not None else value
    for broken, good in _MAC_ROMAN_PAIRS.items():
        text = text.replace(broken, good)
    text = unicodedata.normalize("NFC", text)
    return re.sub(r"\s+", " ", text).strip()


def fold(text: str) -> str:
    """Lowercase and strip accents, for keyword matching only."""
    nfkd = unicodedata.normalize("NFKD", text.lower())
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))
