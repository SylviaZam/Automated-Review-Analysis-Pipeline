"""Classifier backends.

Both backends return the same shape, so reports and evaluation do not care
which one ran:

    Result(themes=[primary, ...], sentiment="Positive" | ... | "n/a", status="ok" | "failed")

`rules`  - offline, free, deterministic. Codebook keywords + bilingual lexicon.
`openai` - optional. The model must pick theme ids from the codebook enum
           (JSON-schema structured output), so categories stay comparable
           across brands. Failed calls are marked `failed` and counted in the
           run manifest instead of silently becoming "Neutral".
"""
from __future__ import annotations

import hashlib
import json
import os
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

from . import sentiment as senti
from .codebook import OTHER, Codebook
from .questions import SENTIMENT_ROLES


@dataclass
class Item:
    role: str
    question: str
    answer: str
    rating: object = None


@dataclass
class Result:
    themes: list[str]
    sentiment: str
    status: str = "ok"
    sentiment_source: str = "text"
    extra: dict = field(default_factory=dict)

    @property
    def primary(self) -> str:
        return self.themes[0] if self.themes else OTHER


class RulesClassifier:
    name = "rules"

    def __init__(self, codebook: Codebook):
        self.codebook = codebook

    def classify(self, items: list[Item]) -> list[Result]:
        out = []
        for item in items:
            themes = self.codebook.match(item.answer)
            if item.role in SENTIMENT_ROLES:
                out.append(Result(themes, senti.classify(item.answer)))
            else:
                out.append(Result(themes, "n/a", sentiment_source="n/a"))
        return out


class OpenAIClassifier:
    """Batched structured-output classification with an on-disk cache.

    The cache is keyed by a SHA-256 of (codebook version, role, question, answer)
    and stores only labels, never answer text.
    """

    name = "openai"
    batch_size = 20

    def __init__(self, codebook: Codebook, industry: str, model: str | None = None, cache_path: str | None = ".voc_cache.json"):
        from openai import OpenAI  # imported lazily so the offline mode has no API dependency

        if not os.getenv("OPENAI_API_KEY"):
            raise RuntimeError("OPENAI_API_KEY is not set; use --backend rules for offline mode.")
        self.client = OpenAI()
        self.model = model or os.getenv("VOC_OPENAI_MODEL", "gpt-4o-mini")
        self.codebook = codebook
        self.industry = industry
        self.cache_path = Path(cache_path) if cache_path else None
        self.cache: dict[str, dict] = {}
        if self.cache_path and self.cache_path.exists():
            self.cache = json.loads(self.cache_path.read_text(encoding="utf-8"))

    def _key(self, item: Item) -> str:
        raw = "|".join([self.codebook.version, self.model, item.role, item.question, item.answer])
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()

    def _schema(self) -> dict:
        return {
            "type": "object",
            "additionalProperties": False,
            "required": ["results"],
            "properties": {
                "results": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["index", "themes", "sentiment"],
                        "properties": {
                            "index": {"type": "integer"},
                            "themes": {"type": "array", "items": {"type": "string", "enum": self.codebook.ids}},
                            "sentiment": {"type": "string", "enum": list(senti.LABELS) + ["n/a"]},
                        },
                    },
                }
            },
        }

    def _call(self, batch: list[Item]) -> dict[int, Result]:
        system = (
            f"You code customer feedback for a {self.industry} e-commerce brand. Answers are mostly Mexican Spanish.\n"
            "For each answer, return 1-3 theme ids from this codebook, most important first:\n"
            f"{self.codebook.prompt_table()}\n"
            "Use no_concern only when the customer says they had no concern or nothing to add.\n"
            "Read each answer in the context of its question. Give a sentiment only when the question asks "
            "for an opinion (role review or open_feedback); otherwise return n/a."
        )
        payload = [{"index": i, "role": it.role, "question": it.question, "answer": it.answer[:600]} for i, it in enumerate(batch)]
        resp = self.client.chat.completions.create(
            model=self.model,
            temperature=0,
            messages=[{"role": "system", "content": system}, {"role": "user", "content": json.dumps(payload, ensure_ascii=False)}],
            response_format={"type": "json_schema", "json_schema": {"name": "coding", "strict": True, "schema": self._schema()}},
        )
        parsed = json.loads(resp.choices[0].message.content or "{}")
        results = {}
        for row in parsed.get("results", []):
            themes = [t for t in row["themes"] if t in self.codebook.ids] or [OTHER]
            results[int(row["index"])] = Result(themes, row["sentiment"], sentiment_source="model" if row["sentiment"] != "n/a" else "n/a")
        return results

    def classify(self, items: list[Item]) -> list[Result]:
        out: list[Result | None] = [None] * len(items)
        pending = []
        for i, item in enumerate(items):
            cached = self.cache.get(self._key(item))
            if cached:
                out[i] = Result(**cached)
            else:
                pending.append(i)
        for start in range(0, len(pending), self.batch_size):
            idx = pending[start:start + self.batch_size]
            batch = [items[i] for i in idx]
            got: dict[int, Result] = {}
            for attempt in range(4):
                try:
                    got = self._call(batch)
                    break
                except Exception as exc:  # network, rate limit, malformed output
                    if attempt == 3:
                        print(f"[warn] batch failed after retries: {type(exc).__name__}", file=sys.stderr)
                    time.sleep(2 ** attempt)
            for local, i in enumerate(idx):
                res = got.get(local)
                if res is None:
                    out[i] = Result([OTHER], "n/a", status="failed", sentiment_source="n/a")
                    continue
                if items[i].role not in SENTIMENT_ROLES:
                    res.sentiment, res.sentiment_source = "n/a", "n/a"
                out[i] = res
                self.cache[self._key(items[i])] = {"themes": res.themes, "sentiment": res.sentiment, "sentiment_source": res.sentiment_source}
        if self.cache_path:
            self.cache_path.write_text(json.dumps(self.cache), encoding="utf-8")
        return [r for r in out if r is not None]


def make_classifier(backend: str, codebook: Codebook, industry: str, **kwargs):
    if backend == "rules":
        return RulesClassifier(codebook)
    if backend == "openai":
        return OpenAIClassifier(codebook, industry, **kwargs)
    raise ValueError(f"unknown backend {backend!r}")
