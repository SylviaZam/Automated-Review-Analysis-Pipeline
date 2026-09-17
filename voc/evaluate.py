"""Score classifiers against a hand-labelled gold set.

Gold CSV columns:
    id, role, question, answer, gold_theme, gold_sentiment [, gold_theme_2]

`gold_sentiment` may be blank for non-opinion roles. `gold_theme_2`, when
present, is a second coder's primary theme and is used for Cohen's kappa
(how consistent the human labels are - the ceiling for any classifier).
"""
from __future__ import annotations

import importlib.util
from collections import Counter
from pathlib import Path

import pandas as pd

from .classify import Item, Result
from .clean import Cleaner
from .codebook import OTHER, Codebook
from .questions import SENTIMENT_ROLES

V1_CATEGORY_MAP = {
    "Price": "price_value", "Shipping": "shipping_delivery", "Quality": "quality",
    "Fit": "size_fit", "Design": "design_style", "Support": "customer_service_returns", "General": OTHER,
}
LEGACY_PATH = Path(__file__).resolve().parent.parent / "legacy" / "v1_survey_analysis.py"


class V1DemoClassifier:
    """The v1 offline analyser (VADER + six English/Spanish keyword buckets), for comparison."""

    name = "v1-demo"

    def __init__(self):
        spec = importlib.util.spec_from_file_location("v1_survey_analysis", LEGACY_PATH)
        module = importlib.util.module_from_spec(spec)  # type: ignore[arg-type]
        spec.loader.exec_module(module)  # type: ignore[union-attr]
        self._analyze = module.demo_analyze_answer

    def classify(self, items: list[Item]) -> list[Result]:
        out = []
        for item in items:
            sentiment, category = self._analyze(item.answer)
            out.append(Result([V1_CATEGORY_MAP.get(category, OTHER)], sentiment))
        return out


def cohen_kappa(a: list[str], b: list[str]) -> float:
    n = len(a)
    if n == 0:
        return float("nan")
    observed = sum(x == y for x, y in zip(a, b)) / n
    ca, cb = Counter(a), Counter(b)
    expected = sum(ca[k] * cb.get(k, 0) for k in ca) / (n * n)
    return (observed - expected) / (1 - expected) if expected < 1 else 1.0


def prf(gold: list[str], pred: list[str]) -> pd.DataFrame:
    labels = sorted(set(gold) | set(pred))
    rows = []
    for label in labels:
        tp = sum(g == label and p == label for g, p in zip(gold, pred))
        fp = sum(g != label and p == label for g, p in zip(gold, pred))
        fn = sum(g == label and p != label for g, p in zip(gold, pred))
        precision = tp / (tp + fp) if tp + fp else 0.0
        recall = tp / (tp + fn) if tp + fn else 0.0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
        rows.append({"label": label, "support": tp + fn, "precision": precision, "recall": recall, "f1": f1})
    return pd.DataFrame(rows)


def score(gold: pd.DataFrame, classifier, clean: bool = True, corpus: list[str] | None = None) -> dict:
    """Score one backend. With `clean`, answers go through the same cleaning step
    as `voc run` (junk answers become `other`/Neutral); v1 had no cleaning step."""
    answers = gold.answer.astype(str).tolist()
    junk = [False] * len(answers)
    if clean:
        cleaner = Cleaner().fit(corpus or answers)
        cleaned = [cleaner.clean(a) for a in answers]
        answers = [c.text for c in cleaned]
        junk = [c.junk is not None for c in cleaned]
    items = [Item(r.role, r.question, a) for r, a in zip(gold.itertuples(), answers)]
    keep = [i for i, j in enumerate(junk) if not j]
    classified = classifier.classify([items[i] for i in keep])
    results = [Result([OTHER], "Neutral", status="junk") for _ in items]
    for i, res in zip(keep, classified):
        results[i] = res
    pred_primary = [r.primary for r in results]
    gold_theme = gold.gold_theme.astype(str).tolist()
    any_hit = [g in r.themes for g, r in zip(gold_theme, results)]
    theme_table = prf(gold_theme, pred_primary)
    supported = theme_table[theme_table.support > 0]

    out = {
        "backend": getattr(classifier, "name", "?"),
        "n": len(gold),
        "theme_accuracy": sum(g == p for g, p in zip(gold_theme, pred_primary)) / len(gold),
        "theme_any_match": sum(any_hit) / len(gold),
        "theme_macro_f1": float(supported.f1.mean()) if len(supported) else float("nan"),
        "theme_table": theme_table,
        "predictions": pd.DataFrame({"id": gold["id"], "gold_theme": gold_theme, "pred_theme": pred_primary,
                                     "pred_themes": [";".join(r.themes) for r in results],
                                     "pred_sentiment": [r.sentiment for r in results]}),
    }

    mask = gold.role.isin(SENTIMENT_ROLES) & gold.gold_sentiment.notna() & (gold.gold_sentiment.astype(str) != "")
    if mask.any():
        g = gold.loc[mask, "gold_sentiment"].astype(str).tolist()
        p = [results[i].sentiment for i in range(len(gold)) if mask.iloc[i]]
        table = prf(g, p)
        out.update({
            "sentiment_n": len(g),
            "sentiment_accuracy": sum(x == y for x, y in zip(g, p)) / len(g),
            "sentiment_macro_f1": float(table[table.support > 0].f1.mean()),
            "false_negative_rate_on_positive": (
                sum(x == "Positive" and y == "Negative" for x, y in zip(g, p)) / max(1, g.count("Positive"))
            ),
            "sentiment_table": table,
            "sentiment_confusion": pd.crosstab(pd.Series(g, name="gold"), pd.Series(p, name="predicted")),
        })

    if "gold_theme_2" in gold.columns and gold.gold_theme_2.notna().any():
        both = gold[gold.gold_theme_2.notna()]
        out["human_kappa"] = cohen_kappa(both.gold_theme.astype(str).tolist(), both.gold_theme_2.astype(str).tolist())
        out["human_kappa_n"] = len(both)
    return out


def markdown(scores: list[dict], title: str) -> str:
    lines = [f"# {title}", "", "| Backend | n | Theme accuracy (primary) | Theme any-match | Theme macro-F1 | Sentiment n | Sentiment accuracy | Sentiment macro-F1 | Positives called Negative |",
             "|---|---|---|---|---|---|---|---|---|"]
    for s in scores:
        def f(key):
            v = s.get(key)
            return "-" if v is None else f"{v:.1%}"
        lines.append(f"| {s['backend']} | {s['n']} | {f('theme_accuracy')} | {f('theme_any_match')} | {s['theme_macro_f1']:.2f} | "
                     f"{s.get('sentiment_n', '-')} | {f('sentiment_accuracy')} | "
                     f"{('%.2f' % s['sentiment_macro_f1']) if 'sentiment_macro_f1' in s else '-'} | {f('false_negative_rate_on_positive')} |")
    for s in scores:
        if "human_kappa" in s:
            lines += ["", f"Human coder agreement (Cohen's kappa, n={s['human_kappa_n']}): {s['human_kappa']:.2f}"]
            break
    for s in scores:
        lines += ["", f"## {s['backend']}: per-theme results", "",
                  s["theme_table"].to_markdown(index=False, floatfmt=".2f")]
        if "sentiment_confusion" in s:
            lines += ["", f"## {s['backend']}: sentiment confusion (rows = gold)", "", s["sentiment_confusion"].to_markdown()]
    return "\n".join(lines) + "\n"


def load_gold(path: str | Path) -> pd.DataFrame:
    gold = pd.read_csv(path, dtype=str)
    missing = {"id", "role", "question", "answer", "gold_theme"} - set(gold.columns)
    if missing:
        raise ValueError(f"gold file missing columns: {sorted(missing)}")
    if "gold_sentiment" not in gold.columns:
        gold["gold_sentiment"] = ""
    unknown = set(gold.gold_theme) - set(Codebook.load().ids)
    if unknown:
        raise ValueError(f"gold_theme values not in the codebook: {sorted(unknown)}")
    return gold.reset_index(drop=True)
