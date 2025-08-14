#!/usr/bin/env python3
"""
Analyze survey answers and export a wide-format Excel with per-product sheets
and per-question pie charts.

Modes
- Demo Mode: no OPENAI_API_KEY required (VADER + keywords).
- API Mode: structured JSON from OpenAI Chat Completions.

Question context
- If CSV headers are full question texts, they are passed directly to the API.
- If CSV headers are Q1, Q2, ... you can provide a JSON map via --qmap so the API
  sees a friendly label, e.g. {"Q1": "Fit and sizing", "Q2": "Price and value"}.
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

# Treat these as "no feedback"
FILLER_VALUES = {"", "n/a", "na", "no", "none", "null", "nan", "sin comentarios", "ninguno", "-", " "}

# Simple keyword categories for Demo Mode (EN + ES)
DEMO_KEYWORDS = [
    ("Price",    ["price", "expensive", "too expensive", "cheap", "cost", "pricing", "value", "caro", "barato", "precio"]),
    ("Shipping", ["ship", "shipping", "delivery", "arrive", "delay", "delayed", "late", "envío", "envio", "tarde", "demor", "entrega"]),
    ("Quality",  ["quality", "material", "durable", "break", "defect", "defecto", "calidad"]),
    ("Fit",      ["fit", "size", "sizing", "tight", "loose", "talla", "ajuste", "grande", "chico"]),
    ("Design",   ["design", "style", "color", "look", "diseño", "estilo", "colores"]),
    ("Support",  ["support", "help", "service", "refund", "return", "soporte", "atención", "atencion", "reembolso", "devolución", "devolucion"]),
]

# XlsxWriter for charts
# (pandas will use it when engine="xlsxwriter")
# Ensure requirements.txt includes: XlsxWriter
# ---------------------------------------------------------------------------

def clean_text(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"[\U00010000-\U0010ffff]", "", s)  # drop emoji/high codepoints
    return re.sub(r"\s+", " ", s).strip()

def is_filler(s: str) -> bool:
    return (s or "").strip().lower() in FILLER_VALUES

def get_question_columns(df: pd.DataFrame) -> List[str]:
    # Assume first three columns are Email, Name, Products; rest are questions
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

# ---------------- Demo analyzer ----------------

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
    # Tiny fallback lexicon
    pos = ["love", "loved", "great", "good", "excellent", "amazing", "encanta", "bueno", "genial", "excelente"]
    neg = ["bad", "poor", "terrible", "awful", "hate", "malo", "expensive", "too expensive", "caro", "tarde", "defecto", "delay", "delayed", "late"]
    p = sum(w in low for w in pos)
    n = sum(w in low for w in neg)
    return "Mixed" if (p and n) else ("Positive" if p else ("Negative" if n else "Neutral"))

def demo_analyze_answer(answer: str) -> Tuple[str, str]:
    txt = (answer or "").strip()
    low = txt.lower()
    return _demo_sentiment(txt, low), _demo_category(low)

# ---------------- Cache ----------------

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

def cache_key(industry: str, question: str, answer: str) -> str:
    return f"{industry}|||{question}|||{answer}"

# ---------------- Question labels ----------------

def load_qmap(path: Optional[str]) -> Dict[str, str]:
    """
    Load a header -> friendly label map from JSON.
    Keys can be either the raw CSV header (e.g. 'Q1') or its sanitized form ('Q1').
    Values should be the friendly label shown to the model and used in chart titles.
    """
    if not path:
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return {str(k).strip(): str(v).strip() for k, v in data.items()}
            return {}
    except Exception as e:
        print(f"[warn] Could not load qmap file: {e}", file=sys.stderr)
        return {}

def sanitize_base(header: str) -> str:
    return re.sub(r"\s+", "_", str(header).strip())

def build_label_map(question_headers: List[str], qmap: Dict[str, str]) -> Dict[str, str]:
    """
    Returns a map from sanitized header -> display label.
    Priority:
      1) qmap[raw header] if present
      2) qmap[sanitized header] if present
      3) else the raw header itself (works when header is the full question text)
    """
    out = {}
    for h in question_headers:
        raw = str(h).strip()
        san = sanitize_base(raw)
        label = qmap.get(raw) or qmap.get(san) or raw
        out[san] = label
    return out

# ---------------- OpenAI path ----------------

def call_openai_analyze(
    industry: str,
    question_context: str,
    answer: str,
    client: "OpenAI",
    model: str = "gpt-4o-mini",
    max_tokens: int = 40,
) -> Tuple[str, str]:
    """
    Ask OpenAI for JSON {sentiment, category}. Uses response_format='json_object'.
    Retries with exponential backoff on transient errors/rate limits.
    """
    sys_prompt = "You are an assistant that analyzes customer feedback."
    user_prompt = (
        "Respond ONLY as JSON with keys 'sentiment' and 'category'.\n"
        f"Industry: {industry}\nQuestion: {question_context}\nAnswer: {answer}\n"
        "Sentiment must be one of: Positive, Neutral, Negative, Mixed. Category should be 1 to 3 words."
    )

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

# ---------------- Analysis ----------------

def analyze_dataframe_wide(
    df: pd.DataFrame,
    industry: str,
    client: Optional["OpenAI"],
    cache_path: Optional[str],
    max_chars: int,
    label_map: Dict[str, str],
) -> pd.DataFrame:
    """
    Build a wide-format table with Qx_Answer, Qx_Sentiment, Qx_Category.
    Uses on-disk cache to avoid duplicate calls.
    Uses label_map to give the API a better question context when headers are Q1...Qn.
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

    # Cache
    cache_dict = load_cache(cache_path)
    dirty = False
    flush_every = 200
    calls_since_flush = 0

    def get_sent_cat(q_header: str, ans: str) -> Tuple[str, str]:
        nonlocal dirty, calls_since_flush
        san = sanitize_base(q_header)
        question_for_model = label_map.get(san, q_header)  # friendly label or raw header
        k = cache_key(industry, question_for_model, ans)
        if k in cache_dict:
            return cache_dict[k]
        if client is None:
            sent, cat = demo_analyze_answer(ans)
        else:
            ans_for_api = ans[:max_chars]
            sent, cat = call_openai_analyze(industry, question_for_model, ans_for_api, client)
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

        q_triplets: Dict[str, Tuple[str, str, str]] = {}
        for q in qcols:
            raw_header = str(q).strip()
            ans = clean_text(str(row.get(q, "")))
            if is_filler(ans):
                sent, cat = "Neutral", "No Feedback"
            else:
                sent, cat = get_sent_cat(raw_header, ans)
            q_triplets[raw_header] = (ans, sent, cat)

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
        return pd.DataFrame(columns=["Product"])

    # Column order
    cols = ["ResponseID", "Product"]
    for q in qcols:
        base = sanitize_base(str(q))
        cols += [f"{base}_Answer", f"{base}_Sentiment", f"{base}_Category"]

    wide = pd.DataFrame(results)
    existing = [c for c in cols if c in wide.columns]
    remainder = [c for c in wide.columns if c not in existing]
    return wide[existing + remainder]

# ---------------- Summary ----------------

def build_summary_from_wide(wide: pd.DataFrame) -> pd.DataFrame:
    """Aggregate sentiment counts per Product x Question (Question is the sanitized base)."""
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

# ---------------- Excel (XlsxWriter) ----------------

def _column_widths_from_df(df: pd.DataFrame, min_w=12, max_w=60):
    widths = []
    for col in df.columns:
        max_len = max([len(str(col))] + [len(str(x)) for x in df[col].astype(str).values[:1000]])
        widths.append(min(max(min_w, int(max_len * 0.9)), max_w))
    return widths

def sanitize_sheet_name(s: str) -> str:
    return (re.sub(r"[:\\/?*\[\]]", " ", str(s)))[:31] or "Sheet"

def write_excel_wide(wide: pd.DataFrame, out_path: str, label_map: Dict[str, str]) -> None:
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
                    q_base = str(row["Question"])           # sanitized base used in data
                    display_label = label_map.get(q_base, q_base)  # show friendly name if available

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

# ---------------- Main ----------------

def main():
    load_dotenv()

    ap = argparse.ArgumentParser(description="Survey analysis -> wide Excel with per-question pie charts (Demo/API).")
    ap.add_argument("--input", required=True, help="Path to input CSV.")
    ap.add_argument("--industry", required=True, help="Industry context, e.g. 'Apparel'.")
    ap.add_argument("--output", help="Output Excel path. Defaults to 'data analysis output.xlsx'")
    ap.add_argument("--cache", default=".analysis_cache.json", help="Path to JSON cache file.")
    ap.add_argument("--max-chars", type=int, default=600, help="Max characters of answer sent to API.")
    ap.add_argument("--qmap", help="Optional JSON file mapping headers like 'Q1' to friendly labels used for API and charts.")
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

    # Build label map so API sees question context
    q_headers = get_question_columns(df)
    raw_qmap = load_qmap(args.qmap)
    label_map = build_label_map(q_headers, raw_qmap)
    mapped = sum(1 for h in q_headers if label_map.get(sanitize_base(h), h) != h)
    if mapped:
        print(f"[info] Applied {mapped} friendly question label(s) from qmap.", file=sys.stderr)

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
    wide = analyze_dataframe_wide(
        df=df,
        industry=args.industry,
        client=client,
        cache_path=args.cache,
        max_chars=args.max_chars,
        label_map=label_map,
    )
    write_excel_wide(wide, out_path, label_map)

if __name__ == "__main__":
    main()
