"""Question detection and role assignment.

Every brand wrote the post-purchase survey slightly differently, so each column
is mapped to a *role* (what the question is for). Roles decide how the answers
are analysed: open-text roles are coded against the codebook, choice roles are
tallied, and only evaluative roles get a sentiment label.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

from .textfix import fold

# (role, accent-folded regex) - first match wins, so specific patterns go first.
ROLE_PATTERNS: list[tuple[str, str]] = [
    ("pdp_gate", r"tienes alguna (inquietud|duda)|anything stopping you\?$"),
    ("pdp_blocker", r"inquietud|duda que te dificulte|cuentanos,? .?cual es|what'?s stopping you|questions? before (you )?buy"),
    ("ad_location", r"donde (estaba|viste) el anuncio|enteraste del anuncio|anuncio\?|where (was|did you see) the ad"),
    ("visit_trigger", r"que te trajo|what brought you"),
    ("discovery", r"como (nos |la |lo )?(conociste|te enteraste)|how did you (first )?(hear|find|discover)"),
    ("awareness_time", r"hace cuanto tiempo|how long (have you|did you) know"),
    ("recipient", r"para quien|who (is this|did you buy this) for"),
    ("competitor", r"otra marca .*en mente|en mente .*otra marca|other brands? in mind|consider(ing)? other"),
    ("hesitation", r"incertidumbre|miedo|preocup|hesitat|worr(y|ied)|concern|unsure"),
    ("choice_driver", r"antes que otra marca|en lugar de otra|instead of|over other brands|why (did you )?choose"),
    ("purchase_reason", r"razon principal|problemas? .*resolver|main reason|what problem"),
    ("review", r"^(body|review|review_text|review text|comment|comentario|title)$"),
]

TEXT_ROLES = {"pdp_blocker", "hesitation", "choice_driver", "purchase_reason", "review", "open_feedback"}
SENTIMENT_ROLES = {"review", "open_feedback"}

META_PATTERNS = {
    "product": r"^(line items|products?|product_handle|handle|producto)$",
    "date": r"started at|submitted at|created ?(date|at)|last updated|fecha",
    "rating": r"^(rating|stars|calificacion|estrellas)$",
    "spend": r"total ?spent",
    "orders": r"orders count|customer purchase order",
    "time_to_purchase": r"first visit to purchase",
}

YES_NO = {"si", "sí", "no", "yes"}


@dataclass
class Question:
    column: str
    text: str
    role: str
    kind: str  # "text" | "choice" | "multi" | "gate"


def meta_role(header: str) -> str | None:
    h = fold(str(header)).strip()
    for role, pattern in META_PATTERNS.items():
        if re.search(pattern, h):
            return role
    return None


def _clean_header(header: str) -> str:
    text = re.sub(r"^slide:\s*", "", str(header), flags=re.I)
    text = re.sub(r"\s*\|\s*id:.*$", "", text, flags=re.I)
    return re.sub(r"^\d+\s+", "", text).strip()


def looks_like_question(header: str) -> bool:
    h = str(header)
    return (
        "?" in h
        or h.lower().startswith("slide:")
        or bool(re.fullmatch(r"q\d+", h.strip().lower()))
        or fold(h).strip() in {"response", "body", "review_text", "review text", "comment", "comentario"}
    )


def role_for(header: str) -> str:
    h = fold(_clean_header(header))
    for role, pattern in ROLE_PATTERNS:
        if re.search(pattern, h):
            return role
    return "open_feedback"


def multi_select_share(values: pd.Series) -> float:
    """Share of answers built entirely from options that recur across the column.

    Several stores asked "why did you buy?" as a multi-select; the export joins the
    chosen options with commas ("Precio Justo, Entrega Rapida"), which looks like
    free text until you notice the segments repeat.
    """
    segments = values.str.split(",").explode().str.strip()
    segments = segments[segments != ""]
    if segments.empty:
        return 0.0
    counts = segments.map(fold).value_counts()
    options = {opt for opt, n in counts.items() if n >= max(3, 0.02 * len(values)) and len(opt) <= 60}
    if not options:
        return 0.0
    def from_options(answer: str) -> bool:
        parts = [fold(p.strip()) for p in answer.split(",") if p.strip()]
        return bool(parts) and all(p in options for p in parts)
    return float(values.map(from_options).mean())


def answer_kind(series: pd.Series, role: str) -> str:
    """Decide whether a column holds free text, one pick, several picks, or yes/no."""
    values = series.dropna().astype(str).str.strip()
    values = values[values != ""]
    if values.empty:
        return "text"
    folded = values.map(fold)
    if folded.isin(YES_NO).mean() >= 0.9:
        return "gate"
    if len(values) >= 20 and folded.str.contains(",").mean() > 0.15 and multi_select_share(values) >= 0.9:
        return "multi"
    if role in TEXT_ROLES and role != "open_feedback":
        # Some brands used multiple choice for "why did you buy"; detect it.
        top = folded.value_counts()
        if len(values) >= 20 and top.head(8).sum() / len(values) >= 0.8 and len(top) <= 25:
            return "choice"
        return "text"
    uniq_ratio = folded.nunique() / len(folded)
    if role in {"discovery", "visit_trigger", "ad_location", "recipient", "awareness_time", "competitor"}:
        # "- Other, please specify" columns are free text even for choice questions.
        return "text" if (uniq_ratio > 0.5 and len(values) >= 10) else "choice"
    return "choice" if (len(values) >= 20 and uniq_ratio < 0.15) else "text"


def detect_questions(df: pd.DataFrame, overrides: dict[str, str] | None = None) -> list[Question]:
    overrides = overrides or {}
    questions: list[Question] = []
    for col in df.columns:
        if meta_role(col) and col not in overrides:
            continue
        if col not in overrides and not looks_like_question(col):
            continue
        role = overrides.get(col) or role_for(col)
        kind = answer_kind(df[col], role)
        if role == "pdp_gate" and kind != "gate":
            role = "pdp_blocker"  # a single open "any questions?" prompt
        elif role == "pdp_blocker" and kind == "gate":
            role = "pdp_gate"
        text = _clean_header(col)
        if _is_copy(df, col, [c.column for c in questions if c.text == text]):
            continue
        questions.append(Question(column=col, text=text, role=role, kind=kind))
    return questions


def _is_copy(df: pd.DataFrame, col: str, same_text_cols: list[str]) -> bool:
    """True when an earlier column with the same question holds the same answers.

    Analysts sometimes pasted a second export of the same survey next to the first.
    The rows do not always line up, so compare row by row *and* as a bag of values:
    the same 84 answers offset by a few rows are still the same 84 answers.
    """
    values = df[col].dropna().astype(str).str.strip()
    values = values[values != ""]
    for other in same_text_cols:
        both = df[col].notna() & df[other].notna()
        if both.sum() and (df.loc[both, col].astype(str) == df.loc[both, other].astype(str)).mean() >= 0.5:
            return True
        theirs = df[other].dropna().astype(str).str.strip()
        theirs = theirs[theirs != ""]
        if len(values) and len(theirs):
            shared = (values.value_counts() - theirs.value_counts().reindex(values.value_counts().index, fill_value=0)).clip(lower=0).sum()
            if 1 - shared / len(values) >= 0.8:
                return True
    return False
