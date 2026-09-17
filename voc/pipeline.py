"""Run a dataset through classification and build the summary tables."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

import pandas as pd

from . import __version__
from . import sentiment as senti
from .classify import Item, Result
from .clean import Cleaner
from .codebook import NO_CONCERN, OTHER, Codebook
from .ingest import Dataset
from .textfix import fold

MIN_N_DEFAULT = 30       # smallest base for which shares are published
MIN_CELL_DEFAULT = 5     # smaller counts print as "<5" in public outputs


@dataclass
class Analysis:
    dataset: Dataset
    codebook: Codebook
    backend: str
    coded: pd.DataFrame            # one row per (respondent, text question)
    themes: pd.DataFrame           # question x theme counts and shares
    choices: pd.DataFrame          # question x option counts and shares
    gates: pd.DataFrame            # yes/no gate rates
    sentiment: pd.DataFrame        # question x sentiment
    products: pd.DataFrame         # product x primary theme (reviews / text)
    manifest: dict = field(default_factory=dict)


def _is_answer(value: object) -> bool:
    return isinstance(value, str) and fold(value).strip(" .-") not in {"", "nan", "none", "null"}


def run(ds: Dataset, classifier, codebook: Codebook, min_n: int = MIN_N_DEFAULT) -> Analysis:
    df = ds.frame
    product_col = ds.meta.get("product")
    rating_col = ds.meta.get("rating")

    raw = []
    for q in ds.questions:
        if q.kind != "text":
            continue
        for idx, value in df[q.column].items():
            if _is_answer(value):
                raw.append((idx, q, value))

    # Clean: flag junk, repair typos. The cleaner learns vocabulary from this export.
    cleaner = Cleaner().fit([v for _, _, v in raw])
    cleaned = [cleaner.clean(v) for _, _, v in raw]
    usable = [i for i, c in enumerate(cleaned) if c.junk is None]
    items = []
    for i in usable:
        idx, q, _ = raw[i]
        items.append(Item(q.role, q.text, cleaned[i].text, df.at[idx, rating_col] if rating_col else None))
    classified = classifier.classify(items) if items else []
    results: list[Result] = [Result([OTHER], "n/a", status="junk", sentiment_source="n/a") for _ in raw]
    for i, res in zip(usable, classified):
        results[i] = res

    rows = []
    for (idx, q, value), clean, res in zip(raw, cleaned, results):
        rows.append({
            "row_id": df.at[idx, "row_id"],
            "question": q.text,
            "role": q.role,
            "product": df.at[idx, product_col] if product_col else None,
            "answer": value,
            "answer_clean": clean.text,
            "quality_flag": clean.junk or ("typo_fixed" if clean.corrections else ""),
            "primary_theme": res.primary,
            "themes": ";".join(res.themes),
            "sentiment": res.sentiment,
            "rating_sentiment": senti.from_rating(df.at[idx, rating_col]) if rating_col else None,
            "status": res.status,
        })
    coded = pd.DataFrame(rows, columns=["row_id", "question", "role", "product", "answer", "answer_clean", "quality_flag",
                                        "primary_theme", "themes", "sentiment", "rating_sentiment", "status"])

    # Theme table: multi-label mentions, shares of answered respondents.
    theme_rows = []
    for (question, role), grp in coded.groupby(["question", "role"], sort=False):
        ok = grp[grp.status == "ok"]
        base = len(ok)
        exploded = ok.assign(theme=ok.themes.str.split(";")).explode("theme")
        counts = exploded.theme.value_counts()
        primary = ok.primary_theme.value_counts()
        for theme_id in codebook.ids:
            n = int(counts.get(theme_id, 0))
            if n == 0:
                continue
            theme_rows.append({
                "question": question, "role": role, "theme_id": theme_id, "theme": codebook.label(theme_id),
                "mentions": n, "primary": int(primary.get(theme_id, 0)), "base": base,
                "share": n / base if base else 0.0, "reportable": base >= min_n,
            })
    themes = pd.DataFrame(theme_rows, columns=["question", "role", "theme_id", "theme", "mentions", "primary",
                                               "base", "share", "reportable"])
    if not themes.empty:
        themes = themes.sort_values(["question", "mentions"], ascending=[True, False], kind="stable")

    choice_rows, gate_rows = [], []
    for q in ds.questions:
        series = df[q.column][df[q.column].map(_is_answer)]
        base = len(series)
        if q.kind == "choice":
            # Merge spelling variants ("Instagram", "instagram ") under the most common spelling.
            canon = series.str.strip()
            groups = canon.groupby(canon.map(fold))
            for _, grp in sorted(groups, key=lambda kv: -len(kv[1])):
                choice_rows.append({"question": q.text, "role": q.role, "option": grp.value_counts().index[0],
                                    "count": len(grp), "base": base, "share": len(grp) / base,
                                    "reportable": base >= min_n})
        elif q.kind == "gate":
            yes = int(series.map(fold).isin({"si", "yes"}).sum())
            gate_rows.append({"question": q.text, "shown": base, "yes": yes,
                              "yes_rate": yes / base if base else 0.0, "reportable": base >= min_n})
    choices = pd.DataFrame(choice_rows, columns=["question", "role", "option", "count", "base", "share", "reportable"])
    gates = pd.DataFrame(gate_rows, columns=["question", "shown", "yes", "yes_rate", "reportable"])

    sent = coded[(coded.sentiment != "n/a") & (coded.status == "ok")]
    if sent.empty:
        sentiment = pd.DataFrame(columns=["question"] + list(senti.LABELS) + ["base"])
    else:
        sentiment = (sent.groupby("question").sentiment.value_counts().unstack(fill_value=0)
                     .reindex(columns=list(senti.LABELS), fill_value=0))
        sentiment["base"] = sentiment.sum(axis=1)
        sentiment = sentiment.reset_index()

    products = pd.DataFrame()
    if product_col and not coded.empty:
        valid = coded[(coded.status == "ok") & ~coded.primary_theme.isin([OTHER, NO_CONCERN])]
        products = pd.crosstab(valid["product"], valid["primary_theme"])
        products = products.loc[products.sum(axis=1).sort_values(ascending=False).index]

    manifest = {
        "tool": f"voc-pipeline {__version__}",
        "run_at": dt.datetime.now().isoformat(timespec="seconds"),
        "label": ds.label,
        "source_sha256": ds.source_sha256,
        "source_encoding": ds.encoding,
        "rows": int(len(df)),
        "pii_columns_dropped": ds.dropped_columns,
        "questions": [{"text": q.text, "role": q.role, "kind": q.kind} for q in ds.questions],
        "backend": getattr(classifier, "name", "?"),
        "model": getattr(classifier, "model", None),
        "codebook_version": codebook.version,
        "text_answers_received": int(len(coded)),
        "text_answers_coded": int((coded.status != "junk").sum()) if not coded.empty else 0,
        "junk_removed": {k: int(v) for k, v in coded[coded.status == "junk"].quality_flag.value_counts().items()} if not coded.empty else {},
        "answers_with_typo_fixes": int((coded.quality_flag == "typo_fixed").sum()) if not coded.empty else 0,
        "failed": int((coded.status == "failed").sum()) if not coded.empty else 0,
        "min_n": min_n,
        "notes": ds.notes,
    }
    return Analysis(ds, codebook, manifest["backend"], coded, themes, choices, gates, sentiment, products, manifest)


def public_count(n: int, min_cell: int = MIN_CELL_DEFAULT) -> str:
    return f"<{min_cell}" if 0 < n < min_cell else str(n)
