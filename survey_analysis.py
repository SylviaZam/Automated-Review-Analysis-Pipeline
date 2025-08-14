#!/usr/bin/env python3
"""
Analyze customer survey answers.

Demo-friendly:
- If OPENAI_API_KEY is not set, the script runs without API calls and marks results as Neutral/Needs Review.
- Reads CSV with columns [Email, Name, Products, Q*...].
- Writes an Excel report: one sheet per product plus a Summary sheet.
"""
import argparse
import json
import os
import re
import sys
import time
from typing import List, Optional, Tuple

import pandas as pd
from dotenv import load_dotenv

# Optional dependency
try:
    from langdetect import detect  # type: ignore
except Exception:
    detect = None

# OpenAI SDK (optional)
try:
    from openai import OpenAI  # type: ignore
except Exception:
    OpenAI = None

FILLER_VALUES = {"", "n/a", "na", "no", "none", "null", "nan", "sin comentarios", "ninguno", "-", " "}

def clean_text(s: str) -> str:
    """Trim, drop emojis and very high code points, collapse whitespace."""
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"[\U00010000-\U0010ffff]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def is_filler(s: str) -> bool:
    t = (s or "").strip().lower()
    return t in FILLER_VALUES

def get_question_columns(df: pd.DataFrame) -> List[str]:
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

def call_openai_analyze(industry: str, question: str, answer: str, client: "OpenAI", model: str = "gpt-4o-mini") -> Tuple[str, str]:
    """Return (sentiment, category). Robust parsing with retries."""
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

def analyze_dataframe(df: pd.DataFrame, industry: str, client: Optional["OpenAI"]) -> pd.DataFrame:
    rows = []
    qcols = get_question_columns(df)

    # Optional language info
    if qcols:
        first_nonempty = []
        for q in qcols:
            series = df[q].dropna()
            if not series.empty:
                first_nonempty.append(str(series.iloc[0]))
        lang = detect_language(first_nonempty)
        if lang:
            print(f"[info] Detected language: {lang}")

    for _, r in df.iterrows():
        products_raw = str(r.get(df.columns[2], "")).strip()
        products = [p.strip() for p in products_raw.split(",") if p.strip()] or ["Unspecified"]

        for q in qcols:
            qtext = q
            answer = clean_text(str(r.get(q, "")))
            if is_filler(answer):
                sentiment, category = "Neutral", "No Feedback"
            else:
                if client is None:
                    sentiment, category = "Neutral", "Needs Review"
                else:
                    sentiment, category = call_openai_analyze(industry, qtext, answer, client)

            for prod in products:
                rows.append(
                    {
                        "Product": prod[:100],
                        "Question": qtext,
                        "Answer": answer,
                        "Sentiment": sentiment,
                        "Category": category,
                    }
                )

    return pd.DataFrame(rows)

def write_excel(results: pd.DataFrame, out_path: str) -> None:
    grouped = results.groupby(["Product", "Question", "Category", "Sentiment"]).size().reset_index(name="Count")
    pivot = grouped.pivot_table(
        index=["Product", "Question", "Category"],
        columns="Sentiment",
        values="Count",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    for col in ["Positive", "Neutral", "Negative", "Mixed"]:
        if col not in pivot.columns:
            pivot[col] = 0

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for prod, sub in results.groupby("Product"):
            sheet = re.sub(r"[:\\/?*\[\]]", " ", prod)[:31] or "Product"
            sub[["Question", "Answer", "Sentiment", "Category"]].to_excel(writer, index=False, sheet_name=sheet)
        pivot.to_excel(writer, index=False, sheet_name="Summary")

    print(f"[ok] Wrote Excel report to {out_path}")

def main():
    load_dotenv()

    ap = argparse.ArgumentParser(description="Analyze customer survey answers and export an Excel report.")
    ap.add_argument("--input", required=True, help="Path to input CSV.")
    ap.add_argument("--industry", required=True, help="Industry context, for example 'Fashion'.")
    ap.add_argument("--output", required=False, help="Output Excel path. Defaults to <input>_analysis.xlsx")
    args = ap.parse_args()

    api_key = os.getenv("OPENAI_API_KEY")
    client = None
    if api_key:
        if OpenAI is None:
            print("[error] openai package not installed. See requirements.txt", file=sys.stderr)
            sys.exit(1)
        client = OpenAI(api_key=api_key)
    else:
        print("[info] OPENAI_API_KEY not set. Running in Demo Mode with no API calls.", file=sys.stderr)

    try:
        df = pd.read_csv(args.input)
    except FileNotFoundError:
        print(f"[error] File not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"[error] Could not read CSV: {e}", file=sys.stderr)
        sys.exit(1)

    if df.shape[1] < 4:
        print("[error] Need at least 4 columns: Email, Name, Products, and one question column.", file=sys.stderr)
        sys.exit(1)

    out_path = args.output or re.sub(r"\.csv$", "", args.input) + "_analysis.xlsx"
    results = analyze_dataframe(df, args.industry, client)
    write_excel(results, out_path)

if __name__ == "__main__":
    main()
