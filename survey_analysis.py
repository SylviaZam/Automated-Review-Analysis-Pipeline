#!/usr/bin/env python3
"""
Analyze survey answers and export a wide-format Excel with per-product sheets
and per-question pie charts.

How it identifies questions:
- The script assumes the first three columns are: Email, Name, Products.
- Every column after those is treated as a question.
- The actual CSV header text for each question is passed to the OpenAI API as context.
  Example: if your header is "Fit and sizing", the model sees that string.
  No separate mapping file is required.

Modes
- Demo Mode: no OPENAI_API_KEY required (VADER + keywords).
- API Mode: structured JSON from OpenAI Chat Completions with response_format='json_object'.

Output
- One worksheet per Product, in wide format with columns like:
  QBase_Answer, QBase_Sentiment, QBase_Category
  where QBase is a sanitized version of the original header (spaces -> underscores).
- Summary sheet with sentiment counts.
- Charts - <Product> sheet with one pie per question (labels + percentages).

Default output filename: data analysis output.xlsx
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

# Optional libs
try:
    from langdetect import detect  # type: ignore
except Exception:
    detect = None

try:
    from openai import OpenAI  # type: ignore
except Exception:
    OpenAI = None

# Demo-mode sentiment
try:
    from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer  # type: ignore
    _VADER_ANALYZER = SentimentIntensityAnalyzer()
except Exception:
    _VADER_ANALYZER = None

# Sentiment constants
SENTIMENT_ORDER = ["Positive", "Neutral", "Negative", "Mixed"]

# Use these as meaning "no feedback" from customer
FILLER_VALUES = {"", "n/a", "na", "no", "none", "null", "nan", "sin comentarios", "ninguno", "-", " "}

# Simple keyword categories for Demo Mode (EN + ES), modify as u need
DEMO_KEYWORDS = [
    ("Price",    ["price", "expensive", "too expensive", "cheap", "cost", "pricing", "value", "caro", "barato", "precio"]),
    ("Shipping", ["ship", "shipping", "delivery", "arrive", "delay", "delayed", "late", "envío", "envio", "tarde", "demor", "entrega"]),
    ("Quality",  ["quality", "material", "durable", "break", "defect", "defecto", "calidad"]),
    ("Fit",      ["fit", "size", "sizing", "tight", "loose", "talla", "ajuste", "grande", "chico"]),
    ("Design",   ["design", "style", "color", "look", "diseño", "estilo", "colores"]),
    ("Support",  ["support", "help", "service", "refund", "return", "soporte", "atención", "atencion", "reembolso", "devolución", "devolucion"]),
]

# -----------------------------------------------------------------------------
# Basic helpers for analyzing the answers
# -----------------------------------------------------------------------------

def clean_text(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"[\U00010000-\U0010ffff]", "", s)  # drop emoji/high codepoints
    return re.sub(r"\s+", " ", s).strip()

def is_filler(s: str) -> bool:
    return (s or "").strip().lower() in FILLER_VALUES

def get_question_columns(df: pd.DataFrame) -> List[str]:
    # Remember first three columns have to be email, name, and product(s) purchased; the rest are questions
    return list(df.columns[3:]) if df.shape[1] > 3 else []

def normalize_sentiment(s: str) -> str:
    return {"positive": "Positive", "neutral": "Neutral", "negative": "Negative", "mixed": "Mixed"}.get(
        (s or "").strip().lower(), "Neutral"
    )

def detect_language(sample_answers: List[str]) -> Optional[str]:
    for a in sample_answers:
        a = clean_text(a)
        if a and detect:
            try:
                return detect(a)
            except Exception:
                pass
    return None

def sanitize_base(header: str) -> str:
    return re.sub(r"\s+", "_", str(header).strip())

# -----------------------------------------------------------------------------
# Demo analyzer (if no OpenAI api key was given)
# -----------------------------------------------------------------------------

def _demo_category(low: str) -> str:
    for cat, kws in DEMO_KEYWORDS:
        if any(k in low for k in kws):
            return cat
    return "General"

def _demo_sentiment(txt: str, low: str) -> str:
    if _VADER_ANALYZER is not None:
        try:
            sc = _VADER_ANALYZER.polarity_scores(txt)["compound"]
            if sc >= 0.35:
                return "Positive"
            if sc <= -0.35:
                return "Negative"
            if any(w in low for w in ["but", "aunque", "pero"]) and abs(sc) < 0.35:
                return "Mixed"
            return "Neutral"
        except Exception:
            pass
    # Tiny fallback lexicon (edit this as needed)
    pos = ["love", "loved", "great", "liked it", "like it", "good", "so good", "excellent", "amazing", "encanta", "muy bueno", "bueno", "me gustó", "gustaron", "genial", "excelente"]
    neg = ["bad", "poor", "terrible", "awful", "hate", "malo", "expensive", "too expensive", "caro", "carísimo", "tarde", "defecto", "delay", "delayed", "late"]
    p = sum(w in low for w in pos)
    n = sum(w in low for w in neg)
    return "Mixed" if (p and n) else ("Positive" if p else ("Negative" if n else "Neutral"))

def demo_analyze_answer(answer: str) -> Tuple[str, str]:
    txt = (answer or "").strip()
    low = txt.lower()
    return _demo_sentiment(txt, low), _demo_category(low)
# -----------------------------------------------------------------------------
# Cache
# -----------------------------------------------------------------------------
def load_cache(path: Optional[str]) -> Dict[str, Tuple[str, str]]:
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
            return {k: (v[0], v[1]) for k, v in raw.items()}
    except Exception:
        return {}

def save_cache(path: Optional[str], cache: Dict[str, Tuple[str, str]]) -> None:
    if not path:
        return
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass

def cache_key(industry: str, question_text: str, answer: str) -> str:
    return f"{industry}|||{question_text}|||{answer}"

# -----------------------------------------------------------------------------
# OpenAI path
# -----------------------------------------------------------------------------

def call_openai_analyze(
    industry: str,
    question_text: str,
    answer: str,
    client: "OpenAI",
    model: str = "gpt-4o-mini",
    max_tokens: int = 40,
) -> Tuple[str, str]:
    """
    Ask OpenAI for JSON {sentiment, category}. Uses response_format='json_object'.
    """
    sys_prompt = "You are an expert CRM assistant that analyzes online customer feedback."
    user_prompt = (
        "Respond ONLY as JSON with keys 'sentiment' and 'category'.\n"
        f"Industry: {industry}\nQuestion: {question_text}\nAnswer: {answer}\n"
        "Sentiment must be one of: Positive, Neutral, Negative, Mixed. Category should be 1 to 3 words."
    )

# backoff retry of up to 5 attempts
    delay = 1.0
    for attempt in range(5):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.1,
                max_tokens=max_tokens,
                response_format={"type": "json_object"},
            )
            content = resp.choices[0].message.content or "{}"
            payload = json.loads(content)
            sentiment = normalize_sentiment(str(payload.get("sentiment", "Neutral")))
            category = (payload.get("category") or "No Feedback").strip()
            if sentiment not in {"Positive", "Neutral", "Negative", "Mixed"}:
                sentiment = "Neutral"
            if not category:
                category = "No Feedback"
            return sentiment, category
        except Exception as e:
            if attempt == 4:
                print(f"[warn] OpenAI failed: {e}. Defaulting to Neutral/No Feedback.", file=sys.stderr)
                return "Neutral", "No Feedback"
            time.sleep(delay)
            delay = min(delay * 2, 8.0)

# -----------------------------------------------------------------------------
# Core analysis
# -----------------------------------------------------------------------------

def analyze_dataframe_wide(
    df: pd.DataFrame,
    industry: str,
    client: Optional["OpenAI"],
    cache_path: Optional[str],
    max_chars: int,
) -> Tuple[pd.DataFrame, Dict[str, str]]:

    results: List[Dict[str, str]] = []
    qcols = get_question_columns(df)

    # Map base -> original header for chart titles
    base_to_display: Dict[str, str] = {}
    for q in qcols:
        raw = str(q).strip()
        base_to_display[sanitize_base(raw)] = raw

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

    # Cache
    cache_dict = load_cache(cache_path)
    dirty = False
    flush_every = 200
    calls_since_flush = 0

    def get_sent_cat(q_header_text: str, ans: str) -> Tuple[str, str]:
        nonlocal dirty, calls_since_flush
        k = cache_key(industry, q_header_text, ans)
        if k in cache_dict:
            return cache_dict[k]
        if client is None:
            sent, cat = demo_analyze_answer(ans)
        else:
            ans_for_api = ans[:max_chars]
            sent, cat = call_openai_analyze(industry, q_header_text, ans_for_api, client)
        cache_dict[k] = (sent, cat)
        dirty = True
        calls_since_flush += 1
        if cache_path and dirty and calls_since_flush >= flush_every:
            save_cache(cache_path, cache_dict)
            calls_since_flush = 0
        return sent, cat

    for idx, row in df.iterrows():
        products_raw = str(row.get(df.columns[2], "")).strip()
        products = [p.strip() for p in products_raw.split(",") if p.strip()] or ["Unspecified"]

        # Analyze each question for this response
        q_triplets: Dict[str, Tuple[str, str, str]] = {}
        for q in qcols:
            q_header_text = str(q).strip()           # full header text as question context
            ans = clean_text(str(row.get(q, "")))
            if is_filler(ans):
                sent, cat = "Neutral", "No Feedback"
            else:
                sent, cat = get_sent_cat(q_header_text, ans)
            q_triplets[q_header_text] = (ans, sent, cat)

        # Emit one row per product
        for prod in products:
            out: Dict[str, str] = {"ResponseID": str(idx + 1), "Product": prod[:100]}
            for q in qcols:
                raw_header = str(q).strip()
                base = sanitize_base(raw_header)
                ans, sent, cat = q_triplets[raw_header]
                out[f"{base}_Answer"] = ans
                out[f"{base}_Sentiment"] = sent
                out[f"{base}_Category"] = cat
            results.append(out)

    if cache_path and dirty:
        save_cache(cache_path, cache_dict)

    if not results:
        return pd.DataFrame(columns=["Product"]), base_to_display

    # Column order
    cols = ["ResponseID", "Product"]
    for q in qcols:
        base = sanitize_base(str(q))
        cols += [f"{base}_Answer", f"{base}_Sentiment", f"{base}_Category"]

    wide = pd.DataFrame(results)
    existing = [c for c in cols if c in wide.columns]
    remainder = [c for c in wide.columns if c not in existing]
    return wide[existing + remainder], base_to_display

# -----------------------------------------------------------------------------
# Summary
# -----------------------------------------------------------------------------

def build_summary_from_wide(wide: pd.DataFrame) -> pd.DataFrame:
    """Aggregate the sentiment counts of each product and question (question is the sanitized version)."""
    if wide.empty:
        return pd.DataFrame()

    q_names = [c[:-len("_Sentiment")] for c in wide.columns if c.endswith("_Sentiment")]

    rows = []
    for _, r in wide.iterrows():
        prod = r.get("Product", "Unspecified")
        for base in q_names:
            s = str(r.get(f"{base}_Sentiment", "")).strip() or "Neutral"
            rows.append({"Product": prod, "Question": base, "Sentiment": s})

    long_df = pd.DataFrame(rows)
    grouped = long_df.groupby(["Product", "Question", "Sentiment"]).size().reset_index(name="Count")
    pivot = grouped.pivot_table(
        index=["Product", "Question"],
        columns="Sentiment",
        values="Count",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    for s in SENTIMENT_ORDER:
        if s not in pivot.columns:
            pivot[s] = 0

    ordered = ["Product", "Question"] + SENTIMENT_ORDER
    existing = [c for c in ordered if c in pivot.columns]
    remainder = [c for c in pivot.columns if c not in existing]
    return pivot[existing + remainder]

# -----------------------------------------------------------------------------
# Excel (XlsxWriter)
# -----------------------------------------------------------------------------

def _column_widths_from_df(df: pd.DataFrame, min_w=12, max_w=60):
    widths = []
    for col in df.columns:
        max_len = max([len(str(col))] + [len(str(x)) for x in df[col].astype(str).values[:1000]])
        widths.append(min(max(min_w, int(max_len * 0.9)), max_w))
    return widths

def sanitize_sheet_name(s: str) -> str:
    return (re.sub(r"[:\\/?*\[\]]", " ", str(s)))[:31] or "Sheet"

def write_excel_wide(wide: pd.DataFrame, out_path: str, base_to_display: Dict[str, str]) -> None:
    """Write data sheets per product + Summary + per-question pie charts using XlsxWriter."""
    summary_all = build_summary_from_wide(wide)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb = writer.book

        # Formats
        wrap_top = wb.add_format({"text_wrap": True, "valign": "top"})
        bold = wb.add_format({"bold": True})

        # Data sheets per product
        if not wide.empty:
            for prod, sub in wide.groupby("Product"):
                sheet = sanitize_sheet_name(prod)
                sub_sorted = sub.sort_values("ResponseID")
                sub_sorted.to_excel(writer, index=False, sheet_name=sheet)
                ws = writer.sheets[sheet]

                # auto width + wrap answers
                widths = _column_widths_from_df(sub_sorted)
                for idx, w in enumerate(widths):
                    header = str(sub_sorted.columns[idx])
                    is_answer = header.endswith("_Answer")
                    ws.set_column(idx, idx, w, wrap_top if is_answer else None)

        # Summary sheet
        if summary_all is not None and not summary_all.empty:
            summary_all.to_excel(writer, index=False, sheet_name="Summary")
            ws_sum = writer.sheets["Summary"]
            wsum = _column_widths_from_df(summary_all, min_w=10, max_w=40)
            for idx, w in enumerate(wsum):
                ws_sum.set_column(idx, idx, w)

        # Charts per product (one pie per question)
        if summary_all is not None and not summary_all.empty:
            charts_per_row = 2
            sentiments = SENTIMENT_ORDER

            for prod, prod_df in summary_all.groupby("Product"):
                sheet = sanitize_sheet_name(f"Charts - {prod}")
                ws = wb.add_worksheet(sheet)
                ws.write(0, 0, f"Sentiment Mix per Question — {prod}", bold)

                for i, (_, row) in enumerate(prod_df.sort_values("Question").iterrows()):
                    q_base = str(row["Question"])               # sanitized base name
                    display_label = base_to_display.get(q_base, q_base)  # original header for title

                    # helper block columns (far right)
                    helper_col_labels = 50  # AY-ish
                    helper_col_values = 51  # AZ-ish
                    start_r = 2 + i * 6

                    for k, snt in enumerate(sentiments):
                        ws.write(start_r + k, helper_col_labels, snt)
                        ws.write(start_r + k, helper_col_values, int(row.get(snt, 0)))

                    # Create pie referencing helper block
                    chart = wb.add_chart({"type": "pie"})
                    chart.add_series({
                        "name": f"{display_label} – Sentiment Mix",
                        "categories": [sheet, start_r, helper_col_labels, start_r + len(sentiments) - 1, helper_col_labels],
                        "values":     [sheet, start_r, helper_col_values, start_r + len(sentiments) - 1, helper_col_values],
                        "data_labels": {"percentage": True, "category": True},
                    })
                    total = sum(int(row.get(s, 0)) for s in sentiments)
                    chart.set_title({"name": f"{display_label} – Sentiment Mix (n={total})"})
                    chart.set_size({"width": 480, "height": 320})

                    # place in grid
                    rblock = i // charts_per_row
                    cblock = i % charts_per_row
                    insert_row = 2 + rblock * 20
                    insert_col = 1 + cblock * 9
                    ws.insert_chart(insert_row, insert_col, chart)

    print(f"[ok] Wrote Excel report to {out_path}")

# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

def main():
    load_dotenv()

    ap = argparse.ArgumentParser(description="Survey analysis -> wide Excel with per-question pie charts (Demo/API).")
    ap.add_argument("--input", required=True, help="Path to input CSV.")
    ap.add_argument("--industry", required=True, help="Industry context, e.g. 'Apparel'.")
    ap.add_argument("--output", help="Output Excel path. Defaults to 'data analysis output.xlsx'")
    ap.add_argument("--cache", default=".analysis_cache.json", help="Path to JSON cache file.")
    ap.add_argument("--max-chars", type=int, default=600, help="Max characters of answer sent to API.")
    args = ap.parse_args()

    # Load CSV
    try:
        df = pd.read_csv(args.input)
    except FileNotFoundError:
        print(f"[error] File not found: {args.input}", file=sys.stderr); sys.exit(1)
    except Exception as e:
        print(f"[error] Could not read CSV: {e}", file=sys.stderr); sys.exit(1)

    if df.shape[1] < 4:
        print("[error] Need at least 4 columns: Email, Name, Products, and one question column.", file=sys.stderr)
        sys.exit(1)

    # Select mode (Demo vs API)
    api_key = os.getenv("OPENAI_API_KEY")
    client = None
    if api_key:
        if OpenAI is None:
            print("[error] 'openai' package not installed. See requirements.txt", file=sys.stderr)
            sys.exit(1)
        client = OpenAI(api_key=api_key)
        print("[info] Using OpenAI API mode.", file=sys.stderr)
    else:
        print("[info] OPENAI_API_KEY not set. Running in Demo Mode (offline analyzer).", file=sys.stderr)

    # Analyze and export
    out_path = args.output or "data analysis output.xlsx"
    wide, base_to_display = analyze_dataframe_wide(
        df=df,
        industry=args.industry,
        client=client,
        cache_path=args.cache,
        max_chars=args.max_chars,
    )
    write_excel_wide(wide, out_path, base_to_display)

if __name__ == "__main__":
    main()
