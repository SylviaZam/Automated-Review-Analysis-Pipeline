"""Claude backend (Anthropic Messages API).

Answers are sent in batches of 20 with the codebook as a JSON schema
(`output_config.format`), so every theme is a valid codebook id. The codebook
instructions sit in a cached system prompt, since they are identical on every
request.

Credentials: ANTHROPIC_API_KEY, or any credential source the Anthropic SDK
resolves (for example an `ant auth login` profile).
Model: `--model`, else $VOC_CLAUDE_MODEL, else claude-opus-5.
Effort: $VOC_CLAUDE_EFFORT, else "low" (short classification rarely needs more).
"""
from __future__ import annotations

import json
import os

from .classify import Item, LLMClassifier, Result, batch_payload, coding_instructions, coding_schema
from .codebook import Codebook

DEFAULT_MODEL = "claude-opus-5"
# Models that accept `output_config.effort` (Haiku 4.5 does not).
EFFORT_MODELS = ("claude-opus-", "claude-sonnet-5", "claude-fable-", "claude-mythos-")
# Models that support the server-side refusal fallback with the "default" routing.
FALLBACK_MODELS = ("claude-opus-5", "claude-fable-5-1")
FALLBACK_BETA = "server-side-fallback-2026-07-01"


class ClaudeClassifier(LLMClassifier):
    name = "claude"

    def __init__(self, codebook: Codebook, industry: str, model: str | None = None,
                 cache_path: str | None = ".voc_cache.json", client=None):
        super().__init__(codebook, industry, cache_path)
        import anthropic  # imported lazily so the offline mode has no API dependency

        self._anthropic = anthropic
        self.model = model or os.getenv("VOC_CLAUDE_MODEL", DEFAULT_MODEL)
        self.effort = os.getenv("VOC_CLAUDE_EFFORT", "low")
        if client is not None:
            self.client = client
        else:
            try:
                self.client = anthropic.Anthropic()
            except Exception as exc:  # no credential source configured
                raise RuntimeError(
                    "Could not create an Anthropic client. Set ANTHROPIC_API_KEY (see .env.example) "
                    "or log in with the `ant` CLI; use --backend rules for offline mode."
                ) from exc

    def _request(self, batch: list[Item]) -> dict:
        output_config: dict = {"format": {"type": "json_schema", "schema": coding_schema(self.codebook)}}
        if self.model.startswith(EFFORT_MODELS):
            output_config["effort"] = self.effort
        kwargs: dict = {
            "model": self.model,
            "max_tokens": 16000,
            "system": [{
                "type": "text",
                "text": coding_instructions(self.codebook, self.industry),
                "cache_control": {"type": "ephemeral"},
            }],
            "messages": [{"role": "user", "content": batch_payload(batch)}],
            "output_config": output_config,
        }
        if self.model in FALLBACK_MODELS:
            # If the model declines a batch for policy reasons, the API retries it on
            # Anthropic's recommended fallback model inside the same call.
            kwargs["betas"] = [FALLBACK_BETA]
            kwargs["fallbacks"] = "default"
        return kwargs

    def _call(self, batch: list[Item]) -> dict[int, Result]:
        try:
            response = self.client.beta.messages.create(**self._request(batch))
        except TypeError as exc:
            if "authentication" in str(exc).lower():
                raise RuntimeError(
                    "No Anthropic credentials found. Set ANTHROPIC_API_KEY in .env (see .env.example) "
                    "or log in with the `ant` CLI; use --backend rules for offline mode."
                ) from exc
            raise
        if response.stop_reason in ("refusal", "max_tokens"):
            # A declined or truncated batch cannot be parsed; its items are marked failed.
            return {}
        text = "".join(block.text for block in response.content if block.type == "text")
        parsed = json.loads(text or "{}")
        return self._results_from_rows(parsed.get("results", []))

    def _is_fatal(self, exc: Exception) -> bool:
        a = self._anthropic
        # Configuration problems: retrying cannot fix them.
        return isinstance(exc, (a.AuthenticationError, a.PermissionDeniedError, a.NotFoundError,
                                a.BadRequestError, RuntimeError, TypeError))
