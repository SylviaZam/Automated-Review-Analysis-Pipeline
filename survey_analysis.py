#!/usr/bin/env python3
"""
Analyze customer survey answers.

Wide-format output:
- One row per (Response × Product)
- For each question column (detected dynamically after the first three),
  the Excel includes Q*_Answer, Q*_Sentiment, Q*_Category side by side.

Modes:
- Demo Mode (no OPENAI_API_KEY): offline analyzer using VADER and keyword categories.
- OpenAI Mode (with OPENAI_API_KEY): calls OpenAI for sentiment & category per answer.

Input CSV (assumed):
[Email, Name, Products, Q1, Q2, Q3, Q4, Q5, ...]
Products may be comma-separated.

Output Excel:
- One worksheet per Product with wide columns per question.
- "Summary" sheet: sentiment counts per Product × Question (Positive/Neutral/Negative/Mixed).
"""

import argparse
import json
import os
import re
import sys
import time
from typing import Dict, List, Optional, Tuple

import pandas as pd
from dotenv import load_dotenv

# Optional dependencies
try:
    from langdetect import detect  # type: ignore
except Exception:
    detect = None

try:
    from openai import OpenAI  # type: ignore
except Exception:
    OpenAI = None

try:
    from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer  # type: ignore
    _VADER_ANALYZER = SentimentIntensityAnalyzer()
except Exception:
    _VADER_ANALYZER = None

# Values that should be treated as "no feedback"
FILLER_VALUES = {
    "", "n/a", "na", "no", "none", "null", "nan", "sin comentarios", "ninguno", "-", " "
}

# Demo-mode keyword hints (EN + ES) to produce a short category
DEMO_KEYWORDS = [
    ("Price",    ["price", "expensive", "too expensive", "cheap", "cost", "pricing", "value", "caro", "barato", "precio"]),
    ("Shipping", ["ship", "shipping", "delivery", "arrive", "delay", "delayed", "late", "envío", "envio", "tarde", "demor", "entrega"]),
    ("Quality",  ["quality", "material", "durable", "break", "defect", "defecto", "calidad"]),
    ("Fit",      ["fit", "size", "sizing", "tight", "loose", "talla", "ajuste", "grande", "chico"]),
    ("Design",   ["design", "style", "color", "look", "diseño", "estilo", "colores", "colór"]),
    ("Support",  ["support", "help", "service", "refund", "return", "soporte", "atención", "atencion", "reembolso", "devolución", "devolucion"]),
]

def clean_text(s: str) -> str:
    """Trim, remove astral-plane emojis, collapse whitespace."""
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"[\U00010000-\U0010ffff]", "", s)  # strip emoji/high codepoints
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def is_filler(s: str) -> bool:
    t = (s or "").strip().lower()
    return t in FILLER_VALUES

def get_question_columns(df: pd.DataFrame) -> List[str]:
    """Assume first three columns are Email, Name, Products; everything after = questions."""
    return list(df.columns[3:]) if df.shape[1] > 3 else []

def detect_language(sample_answers: List[str]) -> Optional[str]:
    for a in sample_answers:
        a = clean_text(a)
        if a and detect:
            try:
                return detect(a)
            except Exception:
                pass
    return None

def normalize_sentiment(s: str) -> str:
    mapping = {"positive": "Positive", "neutral": "Neutral", "negative": "Negative", "mixed": "Mixed"}
    t = (s or "").strip().lower()
    return mapping.get(t, "Neutral")

def demo_category(answer_lower: str) -> str:
    for cat, kws in DEMO_KEYWORDS:
        if any(k in answer_lower for k in kws):
            return cat
    return "General"

def demo_sentiment(answer_text: str, answer_lower: str) -> str:
    # Prefer VADER if available
    if _VADER_ANALYZER is not None:
        try:
            score = _VADER_ANALYZER.polarity_scores(answer_text)["compound"]
            if score >= 0.35:
                return "Positive"
            if score <= -0.35:
                return "Negative"
            if any(w in answer_lower for w in ["but", "aunque", "pero"]) and abs(score) < 0.35:
                return "Mixed"
            return "Neutral"
        except Exception:
            pass
    # Tiny lexicon fallback
    pos = ["love", "loved", "great", "good", "excellent", "amazing", "encanta", "bueno", "genial", "excelente"]
    neg = ["bad", "poor", "terrible", "awful", "hate", "malo", "expensive", "too expensive", "caro", "tarde", "defecto", "delay", "delayed", "late"]
    p = sum(w in answer_lower for w in pos)
    n = sum(w in answer_lower for w in neg)
    if p and n:
        return "Mixed"
    if p:
        return "Positive"
    if n:
        return "Negative"
    return "Neutral"

def demo_analyze_answer(answer: str) -> Tuple[str, str]:
    """Offline analysis for Demo Mode. Returns (sentiment, category)."""
    text = (answer or "").strip()
    low = text.lower()
    cat = demo_category(low)
    sent = demo_sentiment(text, low)
    return sent, cat

def call_openai_analyze(industry: str, question: str, answer: str, client: "OpenAI", model: str = "gpt-4o-mini") -> Tuple[str, str]:
    """
    Ask OpenAI for JSON {sentiment, category}. Returns (sentiment, category).
    Retries on transient errors and normalizes sentiment.
    """
    sys_prompt = "You are an assistant that analyzes customer feedback."
    user_prompt = (
        "Respond ONLY as JSON with keys 'sentiment' and 'category'.\n"
        f"Industry: {industry}\nQuestion: {question}\nAnswer: {answer}\n"
        "Sentiment must be one of: Positive, Neutral, Negative, Mixed. Category should be 1 to 3 words."
    )
    delay = 1.0
    for attempt in range(4):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.2,
            )
            content = resp.choices[0].message.content or "{}"
            m = re.search(r"\{.*\}", content, re.S)
            payload = json.loads(m.group(0) if m else content)
            sentiment = normalize_sentiment(str(payload.get("sentiment", "Neutral")))
            category = (payload.get("category") or "No Feedback").strip()
            if sentiment not in {"Positive", "Neutral", "Negative", "Mixed"}:
                sentiment = "Neutral"
            if not category:
                category = "No Feedback"
            return sentiment, category
        except Exception as e:
            if attempt == 3:
                print(f"[warn] OpenAI failed: {e}. Defaulting to Neutral/No Feedback.", file=sys.stderr)
                return "Neutral", "No Feedback"
            time.sleep(delay)
            delay *= 2
    return "Neutral", "No Feedback"

def analyze_dataframe_wide(df: pd.DataFrame, industry: str, client: Optional["OpenAI"]) -> pd.DataFrame:
    """
    Build a wide-format DataFrame with columns for each question:
    Qx_Answer, Qx_Sentiment, Qx_Category.
    One row per (Response × Product).
    """
    results: List[Dict[str, str]] = []
    qcols = get_question_columns(df)

    # Optional language info
    if qcols:
        samples = []
        for q in qcols:
            s = df[q].dropna()
            if not s.empty:
                samples.append(str(s.iloc[0]))
        lang = detect_language(samples)
        if lang:
            print(f"[info] Detected language: {lang}")

    # Iterate responses
    for idx, row in df.iterrows():
        products_raw = str(row.get(df.columns[2], "")).strip()
        products = [p.strip() for p in products_raw.split(",") if p.strip()] or ["Unspecified"]

        # Precompute per-question triplets for this response
        q_triplets: Dict[str, Tuple[str, str, str]] = {}
        for q in qcols:
            ans = clean_text(str(row.get(q, "")))
            if is_filler(ans):
                sent, cat = "Neutral", "No Feedback"
            else:
                if client is None:
                    sent, cat = demo_analyze_answer(ans)
                else:
                    sent, cat = call_openai_analyze(industry, q, ans, client)
            q_triplets[q] = (ans, sent, cat)

        # Emit one wide row per product
        for prod in products:
            out: Dict[str, str] = {
                "ResponseID": str(idx + 1),
                "Product": prod[:100],
            }
            for q in qcols:
                ans, sent, cat = q_triplets[q]
                base = re.sub(r"\s+", "_", str(q).strip())
                out[f"{base}_Answer"] = ans
                out[f"{base}_Sentiment"] = sent
                out[f"{base}_Category"] = cat
            results.append(out)

    if not results:
        return pd.DataFrame(columns=["Product"])  # empty fallback

    # Order columns: ResponseID, Product, then grouped triplets by question
    cols = ["ResponseID", "Product"]
    for q in qcols:
        base = re.sub(r"\s+", "_", str(q).strip())
        cols += [f"{base}_Answer", f"{base}_Sentiment", f"{base}_Category"]

    wide = pd.DataFrame(results)
    # Keep stable column order if all exist
    existing = [c for c in cols if c in wide.columns]
    remainder = [c for c in wide.columns if c not in existing]
    return wide[existing + remainder]

def build_summary_from_wide(wide: pd.DataFrame) -> pd.DataFrame:
    """
    Produce sentiment counts per Product × Question from the wide table.
    We reconstruct a long view just for aggregation.
    """
    if wide.empty:
        return pd.DataFrame()

    # Identify questions by looking for *_Sentiment columns
    q_names = []
    for c in wide.columns:
        if c.endswith("_Sentiment"):
            q_names.append(c[:-len("_Sentiment")])

    records: List[Dict[str, str]] = []
    for _, row in wide.iterrows():
        product = row.get("Product", "Unspecified")
        for base in q_names:
            sentiment = str(row.get(f"{base}_Sentiment", "")).strip() or "Neutral"
            question_label = base  # already normalized
            records.append({"Product": product, "Question": question_label, "Sentiment": sentiment})

    long_df = pd.DataFrame(records)
    if long_df.empty:
        return pd.DataFrame()

    grouped = (
        long_df.groupby(["Product", "Question", "Sentiment"])
        .size()
        .reset_index(name="Count")
    )
    pivot = grouped.pivot_table(
        index=["Product", "Question"],
        columns="Sentiment",
        values="Count",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    # Ensure all sentiment columns exist
    for s in ["Positive", "Neutral", "Negative", "Mixed"]:
        if s not in pivot.columns:
            pivot[s] = 0

    # Order columns
    cols = ["Product", "Question", "Positive", "Neutral", "Negative", "Mixed"]
    existing = [c for c in cols if c in pivot.columns]
    remainder = [c for c in pivot.columns if c not in existing]
    return pivot[existing + remainder]

def sanitize_sheet_name(name: str) -> str:
    """Excel sheet name constraints: <=31 chars, no []:*?/\\."""
    s = re.sub(r"[:\\/?*\[\]]", " ", name)
    return s[:31] or "Sheet"

def write_excel_wide(wide: pd.DataFrame, out_path: str) -> None:
    summary = build_summary_from_wide(wide)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Per-product sheets
        if not wide.empty:
            for prod, sub in wide.groupby("Product"):
                sheet = sanitize_sheet_name(str(prod))
                sub_sorted = sub.sort_values("ResponseID")
                sub_sorted.to_excel(writer, index=False, sheet_name=sheet)
        # Summary sheet
        if summary is not None and not summary.empty:
            summary.to_excel(writer, index=False, sheet_name="Summary")

    print(f"[ok] Wrote Excel report to {out_path}")

def main():
    load_dotenv()

    parser = argparse.ArgumentParser(description="Analyze customer survey answers and export a wide-format Excel report.")
    parser.add_argument("--input", required=True, help="Path to input CSV.")
    parser.add_argument("--industry", required=True, help="Industry context, for example 'Fashion'.")
    parser.add_argument("--output", required=False, help="Output Excel path. Defaults to <input>_analysis.xlsx")
    args = parser.parse_args()

    # Client selection (Demo vs OpenAI)
    api_key = os.getenv("OPENAI_API_KEY")
    client = None
    if api_key:
        if OpenAI is None:
            print("[error] openai package not installed. See requirements.txt", file=sys.stderr)
            sys.exit(1)
        client = OpenAI(api_key=api_key)
    else:
        print("[info] OPENAI_API_KEY not set. Running in Demo Mode with offline analyzer.", file=sys.stderr)

    # Read CSV
    try:
        df = pd.read_csv(args.input)
    except FileNotFoundError:
        print(f"[error] File not found: {args.input}", file=sys.stderr); sys.exit(1)
    except Exception as e:
        print(f"[error] Could not read CSV: {e}", file=sys.stderr); sys.exit(1)

    if df.shape[1] < 4:
        print("[error] Need at least 4 columns: Email, Name, Products, and one question column.", file=sys.stderr)
        sys.exit(1)

    out_path = args.output or re.sub(r"\.csv$", "", args.input) + "_analysis.xlsx"

    wide = analyze_dataframe_wide(df, args.industry, client)
    write_excel_wide(wide, out_path)

if __name__ == "__main__":
    main()