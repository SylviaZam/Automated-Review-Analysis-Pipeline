"""Diagnose one brand's export before coding or labelling it.

The shared codebook is deliberately generic, so every brand needs a check:
what is this export actually asking, what does the codebook fail to cover, and
which columns should not be coded at all.

`voc diagnose` reports:
- question roles and kinds, including multi-select columns that look like free
  text until you notice the options repeat;
- columns that appear twice with different answers (an analyst pasted a second
  export beside the first);
- coverage per question: the share of answers the codebook cannot place;
- candidate vocabulary: the phrases that recur inside those uncovered answers,
  ready for a human to accept into a brand keyword file.

Nothing is applied automatically. The report is a proposal for a researcher.
"""
from __future__ import annotations

import json
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .classify import Item, RulesClassifier
from .clean import Cleaner
from .codebook import NO_CONCERN, OTHER, Codebook
from .ingest import Dataset
from .questions import multi_select_share
from .textfix import fold

# Words that carry no theme on their own.
STOPWORDS = set("""
a al algo alguna algunas alguno algunos ante antes aqui asi aun aunque bien cada como con contra cual cuales cuando
de del desde donde dos el ella ellas ellos en entre era eran es esa ese eso esta estaba estan estar estas este esto
estos fue fueron ha habia han hasta hay la las le les lo los mas me mi mis mucho muy nada ni no nos o os otra otras
otro otros para pero poco por porque pues que se sea ser si sin sobre solo son su sus tambien tan tanto te tener
tengo tiene todo todos tu tus un una uno unos y ya yo mucha muchas les lo cual sino tenia hacer hace hacen ver
the a an and or but not no yes it is was are this that my for with to of in on at from very i you we they have has
had do does did be been so if then than there here what when which who how why all some any more most other
""".split())
_WORD = re.compile(r"[a-z0-9ñ]+")


@dataclass
class Finding:
    level: str  # "issue" | "note"
    message: str


def _phrases(answers: list[str], min_count: int) -> list[tuple[str, int]]:
    counts: Counter = Counter()
    for answer in answers:
        words = [w for w in _WORD.findall(fold(answer)) if w not in STOPWORDS and len(w) > 2]
        counts.update(words)
        counts.update(f"{a} {b}" for a, b in zip(words, words[1:]))
    # Prefer the longer phrase when a bigram is as frequent as its parts.
    ranked = [(p, n) for p, n in counts.items() if n >= min_count]
    ranked.sort(key=lambda kv: (-kv[1], -len(kv[0])))
    kept: list[tuple[str, int]] = []
    for phrase, n in ranked:
        if any(phrase in bigger and n <= bigger_n for bigger, bigger_n in kept):
            continue
        kept.append((phrase, n))
    return kept


def diagnose(ds: Dataset, codebook: Codebook, min_count: int = 3, examples: int = 3) -> dict:
    findings: list[Finding] = []
    rules = RulesClassifier(codebook)

    # Columns that hold the same question but different answers.
    by_text: dict[str, list] = {}
    for q in ds.questions:
        by_text.setdefault(q.text, []).append(q)
    for text, group in by_text.items():
        if len(group) > 1:
            findings.append(Finding("issue", (
                f'"{text[:60]}" appears in {len(group)} columns with different answers. This usually means two '
                "exports were pasted side by side. Split them into separate files before reporting rates, or the "
                "base for that question mixes two populations.")))

    questions = []
    all_uncovered: list[str] = []
    for q in ds.questions:
        values = ds.frame[q.column].dropna().astype(str).str.strip()
        values = values[values != ""]
        entry = {"question": q.text, "role": q.role, "kind": q.kind, "answers": int(len(values))}
        if q.kind == "multi":
            findings.append(Finding("note", (
                f'"{q.text[:60]}" is a multi-select question: answers are combinations of a few options. '
                "It is tallied, not coded, and should not be hand-labelled.")))
        elif q.kind == "text" and len(values) >= 10:
            share = multi_select_share(values)
            if share >= 0.6:
                findings.append(Finding("issue", (
                    f'"{q.text[:60]}" looks partly like a pick-list ({share:.0%} of answers are combinations of '
                    "repeated options), but is being coded as free text. Check the survey setup for this brand.")))
            cleaner = Cleaner().fit(values.tolist())
            cleaned = [cleaner.clean(v) for v in values]
            usable = [c for c in cleaned if c.junk is None]
            results = rules.classify([Item(q.role, q.text, c.text) for c in usable])
            uncovered = [c.text for c, r in zip(usable, results) if r.primary == OTHER]
            entry.update({
                "junk": len(cleaned) - len(usable),
                "coded": len(usable),
                "uncovered": len(uncovered),
                "uncovered_share": round(len(uncovered) / len(usable), 3) if usable else 0.0,
                "no_concern_share": round(sum(r.primary == NO_CONCERN for r in results) / len(usable), 3) if usable else 0.0,
            })
            all_uncovered += uncovered
            if usable and len(uncovered) / len(usable) >= 0.25:
                findings.append(Finding("issue", (
                    f'"{q.text[:60]}": the codebook cannot place {len(uncovered) / len(usable):.0%} of answers '
                    f"({len(uncovered)} of {len(usable)}). Review the candidate vocabulary below before labelling "
                    "or reporting this question.")))
        questions.append(entry)

    candidates = []
    for phrase, n in _phrases(all_uncovered, min_count)[:25]:
        sample = [a for a in all_uncovered if phrase in fold(a)][:examples]
        candidates.append({"phrase": phrase, "count": n, "examples": sample})

    return {
        "label": ds.label,
        "rows": int(len(ds.frame)),
        "questions": questions,
        "findings": [f.__dict__ for f in findings],
        "uncovered_answers": len(all_uncovered),
        "candidates": candidates,
        "codebook_version": codebook.version,
        "themes": codebook.ids,
    }


def markdown(report: dict, show_examples: bool = True) -> str:
    lines = [f"# Diagnosis: {report['label']}", "",
             f"{report['rows']} rows · codebook v{report['codebook_version']} · "
             f"{report['uncovered_answers']} answers the codebook could not place", ""]
    issues = [f for f in report["findings"] if f["level"] == "issue"]
    notes = [f for f in report["findings"] if f["level"] == "note"]
    if issues:
        lines += ["## Fix before labelling or reporting", ""] + [f"- {f['message']}" for f in issues] + [""]
    if notes:
        lines += ["## Notes", ""] + [f"- {f['message']}" for f in notes] + [""]

    lines += ["## Questions", "", "| Question | Role | Kind | Answers | Junk | Not placed |", "|---|---|---|---|---|---|"]
    for q in report["questions"]:
        share = f"{q['uncovered']} ({q['uncovered_share']:.0%})" if "uncovered" in q else "-"
        lines.append(f"| {q['question'][:60]} | {q['role']} | {q['kind']} | {q['answers']} | "
                     f"{q.get('junk', '-')} | {share} |")

    if report["candidates"]:
        lines += ["", "## Candidate vocabulary", "",
                  "Phrases that recur in answers the codebook could not place. Assign each one to a theme id, or "
                  "decide the brand needs a theme the shared codebook does not have. Nothing here is applied "
                  "automatically.", "",
                  "| Phrase | Times | Example answer |", "|---|---|---|"]
        for c in report["candidates"]:
            example = c["examples"][0].replace("|", "/")[:70] if (show_examples and c["examples"]) else ""
            lines.append(f"| `{c['phrase']}` | {c['count']} | {example} |")
        lines += ["", "Accept the useful ones into a brand keyword file, then re-run with it:", "",
                  "```bash", "python -m voc run <export> --extra-keywords codebooks/brand-<name>.json", "```"]
    return "\n".join(lines) + "\n"


def keyword_stub(report: dict) -> str:
    """A starting keyword file: candidate phrases parked under a placeholder theme."""
    return json.dumps({
        "_instructions": (
            "Move each phrase under the theme id it belongs to and delete this key. "
            f"Theme ids: {', '.join(report['themes'])}"
        ),
        "_unassigned": [c["phrase"] for c in report["candidates"]],
    }, ensure_ascii=False, indent=2) + "\n"


def write(report: dict, out_dir: str | Path, stem: str, show_examples: bool = True) -> list[Path]:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    paths = [out / f"{stem}_diagnosis.md", out / f"{stem}_keywords_draft.json"]
    paths[0].write_text(markdown(report, show_examples), encoding="utf-8")
    paths[1].write_text(keyword_stub(report), encoding="utf-8")
    return paths
