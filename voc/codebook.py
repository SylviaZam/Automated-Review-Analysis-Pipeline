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


LOCAL_PREFIX = "local_"


@dataclass
class Theme:
    id: str
    label: str
    label_es: str
    definition: str
    patterns: list[re.Pattern] = field(default_factory=list)
    local: bool = False  # proposed for one brand, not part of the shared codebook


class Codebook:
    """The shared themes, optionally plus brand-specific ones.

    Shared themes keep their ids across brands so results stay comparable.
    Brand-specific themes are prefixed `local_` and are always reported as such,
    so nobody mistakes one brand's invention for a cross-brand measure.
    """

    def __init__(self, themes: list[Theme], version: str):
        self.themes = themes
        self.version = version
        self.by_id = {t.id: t for t in themes}

    @property
    def ids(self) -> list[str]:
        return [t.id for t in self.themes] + [OTHER]

    @property
    def local_ids(self) -> list[str]:
        return [t.id for t in self.themes if t.local]

    def label(self, theme_id: str) -> str:
        return self.by_id[theme_id].label if theme_id in self.by_id else "Other"

    @staticmethod
    def _compile(keywords: list[str]) -> list[re.Pattern]:
        # Anchored patterns (no_concern) must match the whole answer;
        # everything else matches at the start of a word.
        return [re.compile(kw if kw.startswith("^") else r"\b" + kw) for kw in keywords]

    @classmethod
    def load(cls, path: str | Path | None = None, extra_keywords: str | Path | None = None,
             extra_themes: str | Path | None = None) -> "Codebook":
        """Load the shared codebook.

        `extra_keywords`: {theme_id: [words]} - more vocabulary for existing themes.
        `extra_themes`: {"themes": [{id, label, definition, keywords}]} - brand-specific
        themes, usually written by `voc diagnose --propose-themes` and reviewed by a
        human. Their ids are forced to start with `local_`.
        """
        data = json.loads(Path(path or DEFAULT_PATH).read_text(encoding="utf-8"))
        extra: dict[str, list[str]] = {}
        if extra_keywords:
            extra = json.loads(Path(extra_keywords).read_text(encoding="utf-8"))
            extra = {k: v for k, v in extra.items() if not k.startswith("_")}
        themes = []
        for t in data["themes"]:
            keywords = list(t["keywords"]) + [fold(k) for k in extra.get(t["id"], [])]
            themes.append(Theme(t["id"], t["label"], t["label_es"], t["definition"], cls._compile(keywords)))
        unknown = set(extra) - {t.id for t in themes}
        if unknown:
            raise ValueError(f"extra keywords reference unknown theme ids: {sorted(unknown)}")

        if extra_themes:
            payload = json.loads(Path(extra_themes).read_text(encoding="utf-8"))
            themes += cls._local_themes(payload.get("themes", []), {t.id for t in themes})
        return cls(themes, data.get("version", "?"))

    @classmethod
    def _local_themes(cls, entries: list[dict], taken: set[str]) -> list[Theme]:
        out = []
        for entry in entries:
            theme_id = str(entry["id"])
            if not theme_id.startswith(LOCAL_PREFIX):
                theme_id = LOCAL_PREFIX + theme_id
            if theme_id in taken:
                raise ValueError(f"brand theme id collides with an existing theme: {theme_id}")
            taken.add(theme_id)
            keywords = [fold(k) for k in entry.get("keywords", [])]
            if not keywords:
                raise ValueError(f"brand theme {theme_id} has no keywords")
            out.append(Theme(theme_id, entry.get("label", theme_id), entry.get("label_es", ""),
                             entry.get("definition", ""), cls._compile(keywords), local=True))
        return out

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
