"""Answer cleaning: junk detection and typo repair.

Runs after ingest and before classification.

Junk answers (keyboard mashing, "test", lone punctuation or digits, one-letter
replies) are excluded from every base and counted in the run manifest, so
shares are computed on real answers only. "No", "nada", "N/A" are *not* junk:
they are meaningful "no concern" answers.

Typo repair works in two passes:
1. Shorthand expansion ("q" -> "que", "xq" -> "porque", "tmb" -> "también").
2. Corpus-aware spelling correction. A token is only touched when it is not a
   known word form (wordfreq Spanish + English frequency lists, which include
   conjugations and unaccented spellings) and is one *physical* slip away from a
   known word: a missing
   letter, an extra letter, two swapped letters, or a neighbouring key. The
   export itself also acts as a dictionary, so brand and product names that
   customers spell consistently are learned rather than "corrected".
The original answer is always kept alongside the cleaned one.
"""
from __future__ import annotations

import re
from collections import Counter
from functools import lru_cache
from dataclasses import dataclass
from pathlib import Path

from .textfix import fold

SHORTHAND = {
    "q": "que", "k": "que", "xq": "porque", "pq": "porque", "porq": "porque", "xk": "porque",
    "tmb": "también", "tb": "también", "tambn": "también", "x": "por", "xfa": "por favor", "porfa": "por favor",
    "pls": "please", "msj": "mensaje", "info": "información", "dnd": "donde", "cdo": "cuando", "qdo": "cuando",
    "bn": "bien", "mjr": "mejor", "d": "de", "sta": "esta", "toy": "estoy", "n/a": "n/a",
}
TEST_WORDS = {"test", "testing", "prueba", "pruebas", "asdf", "qwerty", "hola", "hi", "hello", "ok", "okay", "xd", "lol", "jaja", "jajaja"}
VOWELS = set("aeiouy")
KEYBOARD_ROWS = ("qwertyuiop", "asdfghjkl", "zxcvbnm")
_WORD = re.compile(r"[a-záéíóúüñ]+", re.I)


def _load_vocab() -> set[str]:
    path = Path(__file__).with_name("vocab_es_en.txt")
    words = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            words.update(fold(w) for w in line.split())
    return words


try:
    from wordfreq import zipf_frequency  # type: ignore
except ImportError:  # pragma: no cover - fall back to the built-in vocabulary only
    zipf_frequency = None

VOCAB = _load_vocab()
MIN_ZIPF = 1.5  # roughly "appears at least a few times per hundred million words"


_ACCENT = {"a": "á", "e": "é", "i": "í", "o": "ó", "u": "ú", "n": "ñ"}


@lru_cache(maxsize=200_000)
def word_score(token: str, lang: str = "es") -> float:
    """Zipf frequency of a word (0 when unknown or wordfreq is missing).

    Spanish scores take the best single-accent spelling, because customers
    type "padecia" for "padecía" and the unaccented form is rare in print.
    """
    if zipf_frequency is None:
        return 0.0
    best = zipf_frequency(token, lang)
    if lang == "es":
        for i, ch in enumerate(token):
            if ch in _ACCENT:
                best = max(best, zipf_frequency(token[:i] + _ACCENT[ch] + token[i + 1:], "es"))
    return best


_ROWS = ("qwertyuiop", "asdfghjkl", "zxcvbnm")
_WORD = re.compile(r"[a-záéíóúüñ]+", re.I)


def _load_vocab() -> set[str]:
    path = Path(__file__).with_name("vocab_es_en.txt")
    words = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            words.update(fold(w) for w in line.split())
    return words


def _load_dictionary() -> set[str]:
    """Accent-folded Spanish + English word lists from pyspellchecker (optional)."""
    try:
        from spellchecker import SpellChecker  # type: ignore
    except ImportError:  # pragma: no cover - falls back to the built-in vocabulary
        return set()
    words: set[str] = set()
    for lang in ("es", "en"):
        words.update(fold(w) for w in SpellChecker(language=lang, distance=1).word_frequency.dictionary)
    return words


VOCAB = _load_vocab()
DICTIONARY = _load_dictionary() | VOCAB

_ROWS = ("qwertyuiop", "asdfghjklñ", "zxcvbnm")
ADJACENT: dict[str, set[str]] = {}
for _r, _row in enumerate(_ROWS):
    for _c, _ch in enumerate(_row):
        near = set()
        for _dr in (-1, 0, 1):
            if 0 <= _r + _dr < len(_ROWS):
                other = _ROWS[_r + _dr]
                near.update(other[max(0, _c - 1):_c + 2])
        near.discard(_ch)
        ADJACENT[_ch] = near


def is_word(token: str) -> bool:
    """A real Spanish or English word form (accents optional)."""
    return token in VOCAB or word_score(token, "es") >= MIN_ZIPF or word_score(token, "en") >= MIN_ZIPF


@dataclass
class Cleaned:
    original: str
    text: str
    junk: str | None  # reason, or None when the answer is usable
    corrections: int = 0


def junk_reason(text: str) -> str | None:
    raw = text.strip()
    folded = re.sub(r"([a-zñ])\1{2,}", r"\1", fold(raw))  # "fabulosossss" is enthusiasm, not mashing
    letters = re.sub(r"[^a-zñ]", "", folded)
    if not raw:
        return "empty"
    if folded in {"no", "na", "n/a", "nada", "none", "ninguno", "ninguna"}:
        return None
    if re.fullmatch(r"\d{1,2}\s?/\s?\d{1,2}", raw.strip(" .!")):
        return None  # "10/10" is a rating, not junk
    if not letters:
        # emoji-only answers still carry sentiment; digits/punctuation do not
        return None if any(ord(ch) > 0x2600 for ch in raw) else "no_letters"
    if re.fullmatch(r"(.)\1{2,}", re.sub(r"[^a-zñ]", "", fold(raw))):
        return "repeated_character"
    if len(letters) == 1:
        return "single_letter"
    if folded.strip(" .!?") in TEST_WORDS:
        return "test_or_greeting"
    if any(row[i:i + 5] in letters for row in KEYBOARD_ROWS for i in range(len(row) - 4)):
        return "keyboard_mash"
    if len(letters) >= 5:
        vowel_ratio = sum(ch in VOWELS for ch in letters) / len(letters)
        tokens = folded.split()
        known = sum(is_word(tok.strip(".,!?")) for tok in tokens)
        if (vowel_ratio < 0.15 or re.search(r"[^aeiouy\s]{6,}", folded)) and known == 0:
            return "keyboard_mash"
    return None


def _slips(word: str) -> set[str]:
    """Words one physical typing slip away: extra letter, missing letter,
    swapped neighbours, or a key next to the intended one."""
    letters = "abcdefghijklmnopqrstuvwxyzñ"
    splits = [(word[:i], word[i:]) for i in range(len(word) + 1)]
    deletes = {a + b[1:] for a, b in splits if b}
    transposes = {a + b[1] + b[0] + b[2:] for a, b in splits if len(b) > 1}
    replaces = {a + c + b[1:] for a, b in splits if b for c in ADJACENT.get(b[0], ())}
    inserts = {a + c + b for a, b in splits for c in letters}
    return deletes | transposes | replaces | inserts


class Cleaner:
    """Fit on all open answers of one export, then clean each answer."""

    def __init__(self, min_frequent: int = 5, max_rare: int = 2):
        self.min_frequent = min_frequent
        self.max_rare = max_rare
        self.counts: Counter = Counter()

    def fit(self, answers: list[str]) -> "Cleaner":
        for answer in answers:
            self.counts.update(fold(tok) for tok in _WORD.findall(answer))
        return self

    def _known(self, word: str) -> bool:
        return is_word(word) or self.counts[word] >= self.min_frequent

    def _correct(self, token: str) -> str:
        folded = fold(token)
        if len(folded) < 4 or is_word(folded):
            return token
        candidates = [
            w for w in _slips(folded)
            if w[:1] == folded[:1]  # first letters are rarely mistyped; changing them invents words
            and (w in VOCAB or word_score(w, "es") >= MIN_ZIPF or self.counts[w] >= self.min_frequent)
            # An edit at the end of a word is usually an inflection or enclitic
            # ("tomandolos", "batallaba"), not a typo.
            and not (w.startswith(folded[:-1]) and folded.startswith(w[:-1]))
        ]
        if not candidates:
            return token
        # Prefer the word customers use most; general Spanish frequency breaks ties.
        best = max(candidates, key=lambda w: (self.counts[w], word_score(w, "es")))
        own = self.counts[folded]
        # A token is a typo if it is rare, or if a one-edit neighbour is used far more
        # often (the same slip repeated by several customers).
        if own > self.max_rare and self.counts[best] < 10 * own:
            return token
        return best if not token[:1].isupper() else best.capitalize()

    @staticmethod
    def _collapse(match: re.Match) -> str:
        """'amoooo' -> 'amo', 'muuuy' -> 'muy', 'cooperar' untouched (only runs of 3+)."""
        word = match.group(0)
        single = re.sub(r"([a-záéíóúñ])\1{2,}", r"\1", word)
        if is_word(fold(single)):
            return single
        return re.sub(r"([a-záéíóúñ])\1{2,}", r"\1\1", word)

    def clean(self, answer: str) -> Cleaned:
        reason = junk_reason(answer)
        if reason:
            return Cleaned(answer, answer, reason)
        fixes = 0

        def expand(match: re.Match) -> str:
            nonlocal fixes
            word = match.group(0)
            repl = SHORTHAND.get(word.lower())
            if repl and repl != word.lower():
                fixes += 1
                return repl
            fixed = self._correct(word)
            if fixed != word:
                fixes += 1
            return fixed

        text = re.sub(r"([!?.])\1{2,}", r"\1", answer)        # "!!!!" -> "!"
        text = re.sub(r"[A-Za-záéíóúüñÁÉÍÓÚÜÑ]*([a-záéíóúñ])\1{2,}[A-Za-záéíóúüñ]*", self._collapse, text)
        text = re.sub(r"(?<![\w/])[A-Za-záéíóúüñÁÉÍÓÚÜÑ/]+(?![\w/])", expand, text)
        return Cleaned(answer, text, None, fixes)
