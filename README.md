Automated Review Analysis Pipeline

Turn survey CSVs into a polished Excel report with product-level insights for VoC/CX analytics.

What it does

Ingests a CSV of survey responses.

Classifies each answer by:

Sentiment: Positive, Neutral, Negative, Mixed

Category: short theme such as Price, Shipping, Quality, Fit, Design, Support

Exports an Excel workbook for BI/VoC workflows:

One sheet per Product in wide format: <Question>_Answer, <Question>_Sentiment, <Question>_Category

Summary sheet with counts by Product × Question × Sentiment

Charts – <Product> sheet with one pie per question (labels + percentages)

In API mode the model reads the exact question text from your CSV headers, so analysis respects each question’s business context (e.g., Fit and sizing, Price and value, Shipping and delivery).

CSV format

Minimum columns, in this order:

Email, Name, Products, <Question 1>, <Question 2>, ...


Notes:

All columns after the first three are treated as questions.

Put your real question text in the headers (recommended).

Products may contain multiple items separated by commas, e.g. Alpha Jacket, Gamma Backpack.

Quick start (Demo mode, zero cost)

Uses VADER for sentiment and simple keywords for category. No API key required.

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python3 survey_analysis.py --input example_survey_large.csv --industry "Apparel"
# Output: data analysis output.xlsx


Open the Excel:

macOS: open "data analysis output.xlsx"

Windows: start "" "data analysis output.xlsx"

Demo limitations

Sentiment is lexicon-based; category is keyword-based.

Good for portfolios, demos, and smoke tests.

For nuanced, multilingual, or domain-heavy feedback, use API mode.

OpenAI API mode

Higher-fidelity classification that leverages your CSV question headers as context.

Set your API key (shell or .env):

OPENAI_API_KEY=sk-************************


Run the pipeline:

source .venv/bin/activate
pip install -r requirements.txt
# Optional: clear cache to force a fresh run
rm -f .analysis_cache.json

python3 survey_analysis.py --input example_survey_large.csv --industry "Apparel"


What happens in API mode

Sends the exact question header text to the model for each answer.

Truncates very long answers and caps max_tokens to control spend.

Uses an on-disk cache so duplicates and re-runs stay inexpensive.

Approximate costs and requirements

Requirements: OpenAI account, API key, and the openai Python package.

Cost depends on model and tokens. With a compact classification model, thousands of short answers typically land in low USD totals; higher-tier models cost more. Always check current pricing and confirm with token logs.

Output structure

Per-product sheets (wide layout):

ResponseID | Product | <Question>_Answer | <Question>_Sentiment | <Question>_Category | ...


Summary sheet: sentiment counts per Product × Question.

Charts – <Product>: one pie per question with labels and percentages.

Notes for recruiters and stakeholders

Built for Voice of Customer, NPS follow-ups, and SKU-level feedback triage.

Works with English and Spanish responses. PII columns are not used in analysis.

Output is Excel-ready for quick insight sharing or downstream BI ingestion.
