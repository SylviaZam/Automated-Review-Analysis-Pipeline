#!/usr/bin/env python3
"""
Analyze customer survey answers and export a wide-format Excel report
with polished per-product pie charts (one pie per question).

Output:
- One worksheet per Product (wide columns: Qx_Answer / Qx_Sentiment / Qx_Category),
  with wrapped text and auto-sized columns.
- One worksheet "Charts - <Product>" containing a grid of pie charts, one per question,
  showing % of Positive / Neutral / Negative / Mixed.
- A "Summary" worksheet aggregating counts per Product × Question × Sentiment.

Modes:
- Demo Mode (no OPENAI_API_KEY): offline analysis via VADER + keyword rules.
- OpenAI Mode (OPENAI_API_KEY set): calls OpenAI for sentiment/category.
"""

import argparse, json, os, re, sys, time
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
try:
    from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer  # type: ignore
    _VADER_ANALYZER = SentimentIntensityAnalyzer()
except Exception:
    _VADER_ANALYZER = None

FILLER_VALUES = {"", "n/a", "na", "no", "none", "null", "nan", "sin comentarios", "ninguno", "-", " "}
DEMO_KEYWORDS = [
    ("Price",    ["price","expensive","too expensive","cheap","cost","pricing","value","caro","barato","precio"]),
    ("Shipping", ["ship","shipping","delivery","arrive","delay","delayed","late","envío","envio","tarde","demor","entrega"]),
    ("Quality",  ["quality","material","durable","break","defect","defecto","calidad"]),
    ("Fit",      ["fit","size","sizing","tight","loose","talla","ajuste","grande","chico"]),
    ("Design",   ["design","style","color","look","diseño","estilo","colores"]),
    ("Support",  ["support","help","service","refund","return","soporte","atención","atencion","reembolso","devolución","devolucion"]),
]

# ---------------- helpers ----------------
def clean_text(s: str) -> str:
    if not isinstance(s, str): return ""
    s = s.strip()
    s = re.sub(r"[\U00010000-\U0010ffff]", "", s)
    return re.sub(r"\s+", " ", s).strip()

def is_filler(s: str) -> bool:
    return (s or "").strip().lower() in FILLER_VALUES

def get_question_columns(df: pd.DataFrame) -> List[str]:
    return list(df.columns[3:]) if df.shape[1] > 3 else []

def detect_language(sample_answers: List[str]) -> Optional[str]:
    for a in sample_answers:
        a = clean_text(a)
        if a and detect:
            try: return detect(a)
            except Exception: pass
    return None

def normalize_sentiment(s: str) -> str:
    return {"positive":"Positive","neutral":"Neutral","negative":"Negative","mixed":"Mixed"}.get((s or "").strip().lower(),"Neutral")

# ---------------- demo analyzer ----------------
def _demo_category(low: str) -> str:
    for cat,kws in DEMO_KEYWORDS:
        if any(k in low for k in kws): return cat
    return "General"

def _demo_sentiment(txt: str, low: str) -> str:
    if _VADER_ANALYZER is not None:
        try:
            sc = _VADER_ANALYZER.polarity_scores(txt)["compound"]
            if sc >= 0.35: return "Positive"
            if sc <= -0.35: return "Negative"
            if any(w in low for w in ["but","aunque","pero"]) and abs(sc) < 0.35: return "Mixed"
            return "Neutral"
        except Exception:
            pass
    pos = ["love","loved","great","good","excellent","amazing","encanta","bueno","genial","excelente"]
    neg = ["bad","poor","terrible","awful","hate","malo","expensive","too expensive","caro","tarde","defecto","delay","delayed","late"]
    p = sum(w in low for w in pos); n = sum(w in low for w in neg)
    return "Mixed" if (p and n) else "Positive" if p else "Negative" if n else "Neutral"

def demo_analyze_answer(ans: str) -> Tuple[str,str]:
    txt = (ans or "").strip(); low = txt.lower()
    return _demo_sentiment(txt,low), _demo_category(low)

# ---------------- OpenAI path ----------------
def call_openai_analyze(industry: str, question: str, answer: str, client: "OpenAI", model: str="gpt-4o-mini") -> Tuple[str,str]:
    sys_prompt = "You are an assistant that analyzes customer feedback."
    user_prompt = ("Respond ONLY as JSON with keys 'sentiment' and 'category'.\n"
                   f"Industry: {industry}\nQuestion: {question}\nAnswer: {answer}\n"
                   "Sentiment must be one of: Positive, Neutral, Negative, Mixed. Category should be 1 to 3 words.")
    delay = 1.0
    for attempt in range(4):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role":"system","content":sys_prompt},{"role":"user","content":user_prompt}],
                temperature=0.2,
            )
            content = resp.choices[0].message.content or "{}"
            m = re.search(r"\{.*\}", content, re.S)
            payload = json.loads(m.group(0) if m else content)
            sentiment = normalize_sentiment(str(payload.get("sentiment","Neutral")))
            category  = (payload.get("category") or "No Feedback").strip()
            if sentiment not in {"Positive","Neutral","Negative","Mixed"}: sentiment = "Neutral"
            if not category: category = "No Feedback"
            return sentiment, category
        except Exception as e:
            if attempt == 3:
                print(f"[warn] OpenAI failed: {e}. Defaulting to Neutral/No Feedback.", file=sys.stderr)
                return "Neutral","No Feedback"
            time.sleep(delay); delay *= 2
    return "Neutral","No Feedback"

# ---------------- shaping ----------------
def analyze_dataframe_wide(df: pd.DataFrame, industry: str, client: Optional["OpenAI"]) -> pd.DataFrame:
    results: List[Dict[str,str]] = []
    qcols = get_question_columns(df)

    if qcols:
        samples = []
        for q in qcols:
            s = df[q].dropna()
            if not s.empty: samples.append(str(s.iloc[0]))
        lang = detect_language(samples)
        if lang: print(f"[info] Detected language: {lang}")

    for idx, row in df.iterrows():
        products_raw = str(row.get(df.columns[2], "")).strip()
        products = [p.strip() for p in products_raw.split(",") if p.strip()] or ["Unspecified"]

        q_triplets: Dict[str,Tuple[str,str,str]] = {}
        for q in qcols:
            ans = clean_text(str(row.get(q,"")))
            if is_filler(ans): sent,cat = "Neutral","No Feedback"
            else: sent,cat = (demo_analyze_answer(ans) if client is None else call_openai_analyze(industry,q,ans,client))
            q_triplets[q] = (ans,sent,cat)

        for prod in products:
            out: Dict[str,str] = {"ResponseID": str(idx+1), "Product": prod[:100]}
            for q in qcols:
                base = re.sub(r"\s+","_",str(q).strip())
                ans,sent,cat = q_triplets[q]
                out[f"{base}_Answer"] = ans
                out[f"{base}_Sentiment"] = sent
                out[f"{base}_Category"] = cat
            results.append(out)

    if not results: return pd.DataFrame(columns=["Product"])

    cols = ["ResponseID","Product"]
    for q in qcols:
        base = re.sub(r"\s+","_",str(q).strip())
        cols += [f"{base}_Answer",f"{base}_Sentiment",f"{base}_Category"]

    wide = pd.DataFrame(results)
    existing = [c for c in cols if c in wide.columns]
    remainder = [c for c in wide.columns if c not in existing]
    return wide[existing + remainder]

def build_summary_from_wide(wide: pd.DataFrame) -> pd.DataFrame:
    if wide.empty: return pd.DataFrame()
    q_names = [c[:-len("_Sentiment")] for c in wide.columns if c.endswith("_Sentiment")]
    rows = []
    for _, r in wide.iterrows():
        prod = r.get("Product","Unspecified")
        for base in q_names:
            rows.append({"Product": prod, "Question": base, "Sentiment": str(r.get(f"{base}_Sentiment","")).strip() or "Neutral"})
    long_df = pd.DataFrame(rows)
    if long_df.empty: return pd.DataFrame()
    grouped = long_df.groupby(["Product","Question","Sentiment"]).size().reset_index(name="Count")
    pivot = grouped.pivot_table(index=["Product","Question"], columns="Sentiment",
                                values="Count", aggfunc="sum", fill_value=0).reset_index()
    for s in ["Positive","Neutral","Negative","Mixed"]:
        if s not in pivot.columns: pivot[s] = 0
    ordered = ["Product","Question","Positive","Neutral","Negative","Mixed"]
    existing = [c for c in ordered if c in pivot.columns]
    remainder = [c for c in pivot.columns if c not in existing]
    return pivot[existing + remainder]

def _column_widths_from_df(df: pd.DataFrame, min_w=12, max_w=60):
    widths = []
    for col in df.columns:
        max_len = max([len(str(col))] + [len(str(x)) for x in df[col].astype(str).values[:1000]])
        widths.append(min(max(min_w, int(max_len * 0.9)), max_w))
    return widths

# ---------------- writer (XlsxWriter) ----------------
def write_excel_wide(wide: pd.DataFrame, out_path: str) -> None:
    summary_all = build_summary_from_wide(wide)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb  = writer.book

        # Format objects
        wrap_top = wb.add_format({"text_wrap": True, "valign": "top"})
        bold     = wb.add_format({"bold": True})

        # Per-product data sheets
        if not wide.empty:
            for prod, sub in wide.groupby("Product"):
                sheet = re.sub(r"[:\\/?*\[\]]", " ", str(prod))[:31] or "Product"
                sub_sorted = sub.sort_values("ResponseID")
                sub_sorted.to_excel(writer, index=False, sheet_name=sheet)
                ws = writer.sheets[sheet]

                # Auto width + wrap answers
                widths = _column_widths_from_df(sub_sorted)
                for idx, w in enumerate(widths):
                    col_letter = idx
                    header = str(sub_sorted.columns[idx])
                    is_answer_col = header.endswith("_Answer")
                    ws.set_column(col_letter, col_letter, w, wrap_top if is_answer_col else None)

        # Summary sheet
        if summary_all is not None and not summary_all.empty:
            summary_all.to_excel(writer, index=False, sheet_name="Summary")
            ws_sum = writer.sheets["Summary"]
            wsum = _column_widths_from_df(summary_all, min_w=10, max_w=40)
            for idx, w in enumerate(wsum):
                ws_sum.set_column(idx, idx, w)

        # Chart sheets (one per product)
        if summary_all is not None and not summary_all.empty:
            sentiments = ["Positive", "Neutral", "Negative", "Mixed"]
            charts_per_row = 2  # layout grid

            for prod, prod_df in summary_all.groupby("Product"):
                sheet = f"Charts - {prod}"
                sheet = re.sub(r"[:\\/?*\[\]]", " ", sheet)[:31] or "Charts"
                ws = wb.add_worksheet(sheet)

                # Title
                ws.write(0, 0, f"Sentiment Mix per Question — {prod}", bold)

                # For each question, write a vertical helper block and insert a pie
                for i, (_, row) in enumerate(prod_df.sort_values("Question").iterrows()):
                    q_label = str(row["Question"])
                    if re.fullmatch(r"\d+", q_label):
                        q_label = f"Q{q_label}"

                    # helper block position (far right, hidden area)
                    helper_col_labels = 50  # AY-ish
                    helper_col_values = 51  # AZ-ish
                    start_r = 2 + i * 6

                    for k, snt in enumerate(sentiments):
                        ws.write(start_r + k, helper_col_labels, snt)
                        ws.write(start_r + k, helper_col_values, int(row.get(snt, 0)))

                    # create pie referencing helper block
                    chart = wb.add_chart({"type": "pie"})
                    chart.add_series({
                        "name":     f"{q_label} – Sentiment Mix",
                        "categories": [sheet, start_r, helper_col_labels, start_r + len(sentiments) - 1, helper_col_labels],
                        "values":     [sheet, start_r, helper_col_values, start_r + len(sentiments) - 1, helper_col_values],
                        "data_labels": {"percentage": True, "category": True},
                    })
                    total = sum(int(row.get(s, 0)) for s in sentiments)
                    chart.set_title({"name": f"{q_label} – Sentiment Mix (n={total})"})
                    chart.set_size({"width": 480, "height": 320})

                    # place in grid
                    rblock = i // charts_per_row
                    cblock = i % charts_per_row
                    insert_row = 2 + rblock * 20
                    insert_col = 1 + cblock * 9
                    ws.insert_chart(insert_row, insert_col, chart)

        # done; writer saves on exit

    print(f"[ok] Wrote Excel report to {out_path}")

# ---------------- main ----------------
def main():
    load_dotenv()
    ap = argparse.ArgumentParser(description="Survey analysis to wide Excel with per-question pie charts (XlsxWriter).")
    ap.add_argument("--input", required=True)
    ap.add_argument("--industry", required=True)
    ap.add_argument("--output", help="Output Excel path. Defaults to 'data analysis output.xlsx'")
    args = ap.parse_args()

    api_key = os.getenv("OPENAI_API_KEY"); client = None
    if api_key:
        if OpenAI is None:
            print("[error] openai package not installed. See requirements.txt", file=sys.stderr); sys.exit(1)
        client = OpenAI(api_key=api_key)
    else:
        print("[info] OPENAI_API_KEY not set. Running in Demo Mode (offline analyzer).", file=sys.stderr)

    try:
        df = pd.read_csv(args.input)
    except FileNotFoundError:
        print(f"[error] File not found: {args.input}", file=sys.stderr); sys.exit(1)
    except Exception as e:
        print(f"[error] Could not read CSV: {e}", file=sys.stderr); sys.exit(1)
    if df.shape[1] < 4:
        print("[error] Need at least 4 columns: Email, Name, Products + ≥1 question.", file=sys.stderr); sys.exit(1)

    out_path = args.output or "data analysis output.xlsx"
    wide = analyze_dataframe_wide(df, args.industry, client)
    write_excel_wide(wide, out_path)

if __name__ == "__main__":
    main()
