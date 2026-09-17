"""Load a survey or review export and return a PII-free, analysable frame.

Supported shapes (auto-detected by headers):
- Post-purchase exports (KNO, Fairing, Zigpoll): one column per question.
- Product-page (PDP) poll exports: a yes/no gate plus an open "what is it?" column.
- Review exports (Judge.me-style): title, body, rating, product_handle.
- v1 format: Email, Name, Products, Q1..Qn.
"""
from __future__ import annotations

import hashlib
import io
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from . import pii
from .questions import Question, detect_questions, meta_role
from .textfix import decode_bytes, fix_cell


@dataclass
class Dataset:
    label: str
    frame: pd.DataFrame
    questions: list[Question]
    meta: dict[str, str]  # meta role -> column
    dropped_columns: int
    encoding: str
    source_sha256: str
    notes: list[str] = field(default_factory=list)


def _read(path: Path, sheet: str | int | None) -> tuple[pd.DataFrame, str]:
    raw = path.read_bytes()
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(io.BytesIO(raw), sheet_name=sheet if sheet is not None else 0, dtype=object), "xlsx"
    text, encoding = decode_bytes(raw)
    return pd.read_csv(io.StringIO(text), dtype=object, keep_default_na=True), encoding


def load(path: str | Path, label: str | None = None, sheet: str | int | None = None,
         role_overrides: dict[str, str] | None = None) -> Dataset:
    path = Path(path)
    df, encoding = _read(path, sheet)
    notes: list[str] = []

    # Repair headers and cells, drop empty/unnamed filler columns.
    df.columns = [fix_cell(str(c)) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if not c.lower().startswith("unnamed") and c.strip()]]
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df = df.apply(lambda col: col.map(fix_cell))
    df = df.loc[:, ~df.columns.duplicated()]

    # Reviews: merge title + body into one answer so the headline is not lost.
    lower = {c.lower(): c for c in df.columns}
    if "body" in lower and "title" in lower:
        t, b = lower["title"], lower["body"]
        df[b] = (df[t].fillna("").astype(str) + ". " + df[b].fillna("").astype(str)).str.strip(". ")
        df = df.drop(columns=[t])
        notes.append("review title merged into body")

    questions = detect_questions(df, role_overrides)
    if not questions:
        raise ValueError(
            f"{path.name}: no question columns found. Expected survey headers ending in '?', "
            "a review 'body' column, or Q1..Qn. Pass role overrides for custom exports."
        )
    question_cols = {q.column for q in questions}

    # Metadata we keep (product, date, rating...), everything identifying goes.
    meta: dict[str, str] = {}
    for col in df.columns:
        if col in question_cols:
            continue
        role = meta_role(col)
        if role and role not in meta:
            meta[role] = col
    drop = [c for c in df.columns if c not in question_cols and c not in meta.values()]
    dropped = len([c for c in drop if pii.is_pii_column(c)])
    df = df.drop(columns=drop)

    if "spend" in meta:
        df[meta["spend"]] = df[meta["spend"]].map(pii.spend_band)
    for q in questions:
        if q.kind == "text":
            df[q.column] = df[q.column].map(lambda v: pii.scrub_text(v) if isinstance(v, str) else v)

    df = df.reset_index(drop=True)
    df.insert(0, "row_id", range(1, len(df) + 1))
    sha = hashlib.sha256(path.read_bytes()).hexdigest()
    return Dataset(label or path.stem, df, questions, meta, dropped, encoding, sha, notes)
