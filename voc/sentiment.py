"""Bilingual (Spanish-first) rule-based sentiment.

v1 used VADER, which only knows English. On Spanish product reviews it labelled
most five-star reviews Neutral or Negative. This lexicon is small on purpose:
it is readable, editable by a researcher, and evaluated in docs/evaluation.md.
"""
from __future__ import annotations

import re

from .textfix import fold

POSITIVE = [
    "excelente", "encant", "buen", "buenisim", "genial", "perfect", "recomiend", "feliz", "content",
    "maravill", "increibl", "rapid", "bonit", "hermos", "lind", "delicios", "rico", "rica",
    "fantastic", "lo mejor", "gust", "funciona", "ayudo", "ayudado", "amo", "satisf", "agradec",
    "facil", "comod", "suave", "fresc", "me fascina", "10/10", "vale la pena", "buenos resultados", "excelentes resultados",
    "favorit", "recomendad", "antes de lo esperado", "llego bien", "a tiempo",
    "love", "great", "good", "excellent", "amazing", "perfect", "happy", "recommend", "fast", "works",
    "best", "awesome", "nice", "worth", "comfortable", "easy",
]
NEGATIVE = [
    "mal", "pesim", "horribl", "terribl", "decepcion", "defect", "roto", "rota", "rompi", "caro",
    "estafa", "fraude", "queja", "feo", "fea", "peor", "triste", "molest", "irrit", "alergi", "tard",
    "demor", "nunca llego", "no llego", "incomod", "problema", "enojad", "lastima", "desperdicio",
    "bad", "poor", "awful", "worst", "broken", "late", "disappoint", "refund", "scam", "waste",
    "terrible", "horrible", "problem", "never arrived", "itchy", "cheap quality",
    "dura poco", "no dura", "lento", "lenta", "regular", "no se parece", "no huele", "doesnt last",
]
NEGATORS = {"no", "nunca", "ni", "sin", "ningun", "ninguna", "ningunos", "tampoco", "jamas", "not", "never", "dont", "didnt", "isnt", "wasnt", "doesnt"}
CONTRAST = re.compile(r"\b(pero|aunque|sin embargo|solo que|but|however|although)\b")
POS_EMOJI = set("😍❤💕👍🙌✨🥰😊💖💯🔥")
NEG_EMOJI = set("😡👎😞🙁😢😠💔😤")

_POS = [re.compile(r"\b" + re.escape(w)) for w in POSITIVE]
_NEG = [re.compile(r"\b" + re.escape(w)) for w in NEGATIVE]
_TOKEN = re.compile(r"[a-z0-9/']+")

LABELS = ("Positive", "Neutral", "Negative", "Mixed")


def _negated(folded: str, start: int) -> bool:
    before = _TOKEN.findall(folded[:start].replace("'", ""))[-3:]
    return any(tok in NEGATORS for tok in before)


def score(text: str) -> tuple[int, int]:
    folded = fold(text)
    pos = neg = 0
    for patterns, polarity in ((_POS, 1), (_NEG, -1)):
        for p in patterns:
            for m in p.finditer(folded):
                effective = -polarity if _negated(folded, m.start()) else polarity
                if effective > 0:
                    pos += 1
                else:
                    neg += 1
    pos += sum(ch in POS_EMOJI for ch in text)
    neg += sum(ch in NEG_EMOJI for ch in text)
    return pos, neg


def classify(text: str) -> str:
    pos, neg = score(text)
    if pos and neg:
        if CONTRAST.search(fold(text)):
            return "Mixed"
        return "Positive" if pos > neg else ("Negative" if neg > pos else "Mixed")
    if pos:
        return "Positive"
    if neg:
        return "Negative"
    return "Neutral"


def from_rating(rating: object) -> str | None:
    try:
        value = float(rating)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return None
    if value != value:
        return None
    if value >= 4:
        return "Positive"
    if value <= 2:
        return "Negative"
    return "Neutral"
