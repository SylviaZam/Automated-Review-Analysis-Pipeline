# Changelog

## 2.0.0 (2026-09)

Rebuilt after auditing v1 on real Spanish-language exports.

- New `voc` package and CLI: `run`, `eval`, `label-kit`, `synth`.
- Shared 18-theme codebook; multi-label coding with a primary theme.
- Question roles; sentiment only for opinion questions.
- Spanish-first sentiment lexicon with negation, contrast and emoji handling; rating-based sentiment kept for reviews.
- Cleaning step: junk-answer detection, shorthand expansion, typo repair with wordfreq plus a physical-slip model, and the original answer kept alongside the cleaned one.
- Encoding repair for Mac Roman / Windows-1252 exports and double-decoded UTF-8.
- Privacy: identifying columns dropped at ingest, contact details masked, spend banded, small bases suppressed, hash-only cache, and a CI guard against real data, secrets and client names.
- Optional OpenAI backend with strict JSON-schema output constrained to the codebook, batching, and flagged failures.
- Excel workbook, one-page HTML summary, and a run manifest.
- Evaluation harness (accuracy, macro-F1, confusion, Cohen's kappa), a synthetic twin dataset with an answer key, 99 tests, and CI on Python 3.9 and 3.12.
- Removed the committed virtual environment (5,883 files) and the analysis cache; fixed the misnamed `.gitignore`.
- v1 moved to `legacy/` for comparison.

## 1.0 (2025-08)

- `survey_analysis.py`: VADER sentiment with keyword categories (demo), or OpenAI classification; Excel output with per-product sheets and charts.
