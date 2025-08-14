# Automated Review Analysis Pipeline

NLP-powered customer survey analysis. This tool reads a survey CSV, classifies each answer’s **sentiment** and **category**, and exports an Excel report with one sheet per product plus a **Summary** sheet.

## Demo Mode (zero cost)
If no `OPENAI_API_KEY` is set, the script runs in Demo Mode and **does not call any API**. It still cleans data and produces the Excel report, filling results with `Neutral / Needs Review`. Provide a key later to enable real AI labeling.

## Features
- Cleans and normalizes free-text answers (multilingual friendly).
- Skips empty/filler answers and logs them as `Neutral / No Feedback`.
- Optional language detection for metadata logging.
- Excel output: one sheet per product + a Summary aggregation with counts.
- Command-line interface with clear errors and helpful messages.

## Installation
```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
