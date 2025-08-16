# Automated Review Analysis Pipeline

> Turn survey CSVs into a polished Excel report with product-level insights for VoC/CX analytics.

---

## Real-world usage (Unmade services)

This pipeline was operationalized at **Unmade** and used across **all clients who purchased the Customer Data Analytics report** within our scope of services. It ingests survey exports and produces a stakeholder-ready Excel with per-product tabs and a roll-up summary.

**Data collection (Shopify)**
- Survey answers were collected directly from Shopify via **Zigpoll** and **KNO** apps (post-purchase and CRM prompts).
- Analysts exported CSVs from each app per store, then dropped them into the pipeline’s `--input` path.

**Workflow**
1. **Ingest & clean**: trims/normalizes answers, handles empty responses, and maps **question text from the CSV columns (Q1…Qn)** so each answer is analyzed **in the context of its specific question**.
2. **Classification**:  
   - **Demo mode (free)**: offline rules for sentiment + lightweight topic tags.  
   - **API mode (client-funded)**: OpenAI-based classification, multilingual, higher accuracy.
3. **Outputs**:  
   - Excel with **one sheet per product** (Question, Answer, Sentiment, Category)  
   - **Summary sheet** with aggregated counts by sentiment/category and product  
   - Optional charts for quick readouts

**Where this helped**
- Turned thousands of free-text survey responses into an actionable theme map for **roadmap, CX, ecommerce stores UI redesigns, marketing and merchandising**.
- Made it easy to highlight **top complaint themes** and **drivers of delight** per product and market.

**Privacy / scope**
- Client names and PII are not stored in this repo.  
- API usage (when enabled) was **billed to client engagements**; costs vary by model and volume (see README notes).  
- The public sample here runs in **Demo mode** so reviewers can reproduce end-to-end with zero spend.

**Replicate the pattern**
- Use your own Shopify (Zigpoll/KNO or other) CSV exports with the same column structure (Email, Name, Products, Q1…Qn).  
- For **API mode**, set `OPENAI_API_KEY` and re-run; the script automatically uses the **question text from your CSV** to evaluate each answer with correct context.

---

## Table of Contents
- [What It Does](#what-it-does)
- [CSV Format](#csv-format)
- [Quick Start - Demo Mode](#quick-start---demo-mode)
- [Demo Limitations](#demo-limitations)
- [OpenAI API Mode](#openai-api-mode)
- [What Happens In API Mode](#what-happens-in-api-mode)
- [Approximate Costs and Requirements](#approximate-costs-and-requirements)
- [Output Structure](#output-structure)
- [Notes for Recruiters and Stakeholders](#notes-for-recruiters-and-stakeholders)

---

## What It Does

- Ingests a CSV of survey responses.
- Classifies each answer by:
  - **Sentiment:** Positive, Neutral, Negative, Mixed
  - **Category:** short theme such as Price, Shipping, Quality, Fit, Design, Support
- Exports an Excel workbook for BI/VoC workflows:
  - One sheet per Product in wide format:  
    `<Question>_Answer`, `<Question>_Sentiment`, `<Question>_Category`
  - Summary sheet with counts by **Product × Question × Sentiment**
  - Charts - on each `<Product>` sheet, one pie per question with labels and percentages

In API mode the model reads the exact question text from your CSV headers, so analysis respects each question’s business context (for example, Fit and sizing, Price and value, Shipping and delivery).

---

## CSV Format

**Minimum columns, in this order:**

```
Email, Name, Products, <Question 1>, <Question 2>, ...
```

**Notes**
- All columns after the first three are treated as questions.
- Put your real question text in the headers (recommended).
- `Products` may contain multiple items separated by commas, for example: `Alpha Jacket, Gamma Backpack`.

---

## Quick Start - Demo Mode (zero cost)

Uses VADER for sentiment and simple keywords for category. No API key required.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python3 survey_analysis.py --input example_survey_large.csv --industry "Apparel"
# Output: data analysis output.xlsx
```

Open the Excel:

**macOS**
```bash
open "data analysis output.xlsx"
```

**Windows**
```bat
start "" "data analysis output.xlsx"
```

---

## Demo Limitations

- Sentiment is lexicon-based. Category is keyword-based.
- Great for portfolios, demos, and smoke tests.
- For nuanced, multilingual, or domain-heavy feedback, use API mode.

---

## OpenAI API Mode

Higher-fidelity classification that leverages your CSV question headers as context.

Set your API key (shell or `.env`):

```bash
# Example (Unix shells)
export OPENAI_API_KEY="sk-************************"
```

Run the pipeline:

```bash
source .venv/bin/activate
pip install -r requirements.txt
# Optional: clear cache to force a fresh run
rm -f .analysis_cache.json

python3 survey_analysis.py --input example_survey_large.csv --industry "Apparel"
```

---

## What Happens In API Mode

- Sends the exact question header text to the model for each answer.
- Truncates very long answers and caps `max_tokens` to control spend.
- Uses an on-disk cache so duplicates and re-runs stay inexpensive.

---

## Approximate Costs and Requirements

- **Requirements:** OpenAI account, API key, and the `openai` Python package.
- **Cost:** depends on model and tokens. With a compact classification model, thousands of short answers typically land in low USD totals. Higher-tier models cost more. Always check current pricing and confirm with token logs.

---

## Output Structure

**Per-product sheets (wide layout)**

```
ResponseID | Product | <Question>_Answer | <Question>_Sentiment | <Question>_Category | ...
```

**Summary sheet**
- Sentiment counts per Product × Question.

**Charts**
- On each `<Product>` sheet: one pie per question with sentiment analytics.

---

## Notes

- Built for Voice of Customer, NPS follow-ups, and SKU-level feedback triage.
- Works with English and Spanish responses. PII columns are not used in analysis.
- Output is Excel-ready for quick insight sharing or downstream BI ingestion.
