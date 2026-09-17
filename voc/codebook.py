"""Codebook loading and keyword matching."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from .textfix import fold

DEFAULT_PATH = Path(__file__).with_name("codebook.json")
OTHER = "other"
NO_CONCERN = "no_concern"


@dataclass
class Theme:
    id: str
    label: str
    label_es: str
    definition: str
    patterns: list[re.Pattern] = field(default_factory=list)


class Codebook:
    def __init__(self, themes: list[Theme], version: str):
        self.themes = themes
        self.version = version
        self.by_id = {t.id: t for t in themes}

    @property
    def ids(self) -> list[str]:
        return [t.id for t in self.themes] + [OTHER]

    def label(self, theme_id: str) -> str:
        return self.by_id[theme_id].label if theme_id in self.by_id else "Other"

    @classmethod
    def load(cls, path: str | Path | None = None, extra_keywords: str | Path | None = None) -> "Codebook":
        data = json.loads(Path(path or DEFAULT_PATH).read_text(encoding="utf-8"))
        extra: dict[str, list[str]] = {}
        if extra_keywords:
            extra = json.loads(Path(extra_keywords).read_text(encoding="utf-8"))
        themes = []
        for t in data["themes"]:
            keywords = list(t["keywords"]) + [fold(k) for k in extra.get(t["id"], [])]
            patterns = []
            for kw in keywords:
                # Anchored patterns (no_concern) must match the whole answer;
                # everything else matches at the start of a word.
                patterns.append(re.compile(kw if kw.startswith("^") else r"\b" + kw))
            themes.append(Theme(t["id"], t["label"], t["label_es"], t["definition"], patterns))
        unknown = set(extra) - {t.id for t in themes}
        if unknown:
            raise ValueError(f"extra keywords reference unknown theme ids: {sorted(unknown)}")
        return cls(themes, data.get("version", "?"))

    def match(self, text: str) -> list[str]:
        """Return matched theme ids ordered by where they first appear in the text.

        The first id is the primary theme. Returns ["other"] when nothing matches.
        """
        folded = fold(text).strip(" .!¡?¿,;:-")
        no_concern = self.by_id.get(NO_CONCERN)
        bare = re.sub(r"[^\w\s/]", "", folded).strip()  # ignore emoji and punctuation
        if no_concern and any(p.search(bare) for p in no_concern.patterns):
            return [NO_CONCERN]
        hits: list[tuple[int, int, str]] = []
        for order, theme in enumerate(self.themes):
            if theme.id == NO_CONCERN:
                continue
            positions = [m.start() for p in theme.patterns for m in [p.search(folded)] if m]
            if positions:
                hits.append((min(positions), order, theme.id))
        if not hits:
            return [OTHER]
        return [theme_id for _, _, theme_id in sorted(hits)]

    def prompt_table(self) -> str:
        """Theme list for LLM prompts."""
        rows = [f"- {t.id}: {t.definition}" for t in self.themes]
        rows.append(f"- {OTHER}: none of the above")
        return "\n".join(rows)
