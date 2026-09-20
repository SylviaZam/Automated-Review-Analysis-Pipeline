"""Ask Claude to name brand-specific themes from the answers nothing matched.

The offline proposer in `diagnose.py` counts distinctive phrases. It works when a
brand's gap has an obvious word behind it ("brenvita"), and produces fragments when
the gap is about meaning rather than vocabulary. This module reads a sample of the
unplaced answers and proposes named themes instead.

Two rules keep it honest:
- It may not propose a theme that duplicates one already in the shared codebook.
- Every proposal's answer count is measured afterwards by matching its keywords
  against the real answers, never taken from the model's own claim.

Nothing is applied. The output is a file a person edits and passes to
`voc run --extra-themes`.
"""
from __future__ import annotations

import json
import os
import random
import re

from .codebook import Codebook
from .textfix import fold

DEFAULT_MODEL = "claude-opus-5"
SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["themes"],
    "properties": {
        "themes": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["id", "label", "definition", "keywords"],
                "properties": {
                    "id": {"type": "string", "description": "snake_case, no spaces"},
                    "label": {"type": "string", "description": "short English label"},
                    "definition": {"type": "string", "description": "one sentence, what belongs in this theme"},
                    "keywords": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "words or short phrases as customers actually wrote them",
                    },
                },
            },
        }
    },
}


def _client(client=None):
    import anthropic

    if client is not None:
        return client
    try:
        return anthropic.Anthropic()
    except Exception as exc:  # no credential source configured
        raise RuntimeError(
            "No Anthropic credentials found. Set ANTHROPIC_API_KEY in .env (see .env.example), "
            "or drop --propose-backend claude to use the offline proposer."
        ) from exc


def propose(uncovered: list[str], codebook: Codebook, industry: str, max_themes: int = 6,
            sample_size: int = 150, model: str | None = None, client=None, seed: int = 11) -> list[dict]:
    if not uncovered:
        return []
    model = model or os.getenv("VOC_CLAUDE_MODEL", DEFAULT_MODEL)
    sample = random.Random(seed).sample(uncovered, min(sample_size, len(uncovered)))

    system = (
        f"You are helping a UX researcher extend a coding scheme for a {industry} brand.\n"
        "These customer answers could not be placed in the shared codebook below.\n\n"
        f"Existing themes (do not propose anything that means the same as one of these):\n"
        f"{codebook.prompt_table()}\n\n"
        f"Propose at most {max_themes} new themes that would cover a real, recurring part of these answers. "
        "A theme must be about what customers are talking about, not about tone. Skip one-off answers. "
        "Keywords must be words or short phrases that actually appear in the answers, in the language they "
        "were written in, lowercase and without accents. Do not invent vocabulary."
    )
    payload = json.dumps([{"answer": a[:300]} for a in sample], ensure_ascii=False)

    kwargs = {
        "model": model,
        "max_tokens": 8000,
        "system": [{"type": "text", "text": system}],
        "messages": [{"role": "user", "content": payload}],
        "output_config": {"format": {"type": "json_schema", "schema": SCHEMA}},
    }
    if model in ("claude-opus-5", "claude-fable-5-1"):
        kwargs["betas"] = ["server-side-fallback-2026-07-01"]
        kwargs["fallbacks"] = "default"

    response = _client(client).beta.messages.create(**kwargs)
    if response.stop_reason in ("refusal", "max_tokens"):
        raise RuntimeError(f"Claude could not complete the proposal (stop reason: {response.stop_reason}).")
    text = "".join(b.text for b in response.content if b.type == "text")
    proposed = json.loads(text or "{}").get("themes", [])

    existing = set(codebook.ids)
    out = []
    for entry in proposed:
        theme_id = re.sub(r"[^a-z0-9_]+", "_", str(entry.get("id", "")).lower()).strip("_")
        keywords = [fold(k).strip() for k in entry.get("keywords", []) if str(k).strip()]
        if not theme_id or not keywords or theme_id in existing or f"local_{theme_id}" in existing:
            continue
        # Count how many answers this actually covers, rather than trusting the model.
        patterns = [re.compile(r"\b" + re.escape(k)) for k in keywords]
        members = [a for a in uncovered if any(p.search(fold(a)) for p in patterns)]
        out.append({
            "id": theme_id,
            "label": entry.get("label", theme_id),
            "definition": entry.get("definition", ""),
            "keywords": keywords,
            "answers": len(members),
            "examples": members[:3],
            "source": "claude",
        })
    out.sort(key=lambda t: -t["answers"])
    return out[:max_themes]
