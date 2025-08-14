Automated Review Analysis Pipeline

Turn a survey CSV into a polished Excel report for CX analytics: sentiment per answer, theme classification, per-product tabs, and per-question pie charts. Works in two modes:

Demo Mode (no API key, zero cost)

API Mode (OpenAI classification with question context)

What this is for

Rapidly analyze customer feedback across products and dimensions (fit and sizing, price and value, shipping and delivery, support and returns, design and aesthetics, quality and durability). Useful for CX, product, ops, and growth teams to spot drivers of churn, NPS shifts, and logistics issues.

CSV format (required)

Columns in this order:

Email, Name, Products, Q1, Q2, Q3, Q4, Q5


Rules:

All columns after the first three are treated as questions.

Use real question text as headers if you want the model to use that context in API Mode. Example:
Email,Name,Products,Fit and sizing,Price and value,Shipping and delivery,Support and returns,Design and aesthetics

Products can be comma-separated. Example: Alpha Jacket, Gamma Backpack.

Output

One worksheet per Product (wide format): Qx_Answer, Qx_Sentiment (Positive, Neutral, Negative, Mixed), Qx_Category (short theme).

A Summary sheet with counts by Product x Question x Sentiment.

A Charts - <Product> sheet with one pie per question showing sentiment proportions.

Install
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\Activate.ps1
pip install -r requirements.txt

Run in Demo Mode (zero cost)

Uses VADER + keyword themes. Good for portfolio demos and local testing.

source .venv/bin/activate
python3 survey_analysis.py --input example_survey_large.csv --industry "Apparel"
# creates: data analysis output.xlsx


Limitations in Demo Mode:

Sentiment is lexicon-based, not model-based.

Category is keyword-driven (EN/ES). Subtle themes may be missed.

Run in API Mode (OpenAI)

Enables model classification with your real question headers as context.

# set your key in .env or export it
# .env: OPENAI_API_KEY=sk-********************************
source .venv/bin/activate
pip install -r requirements.txt
python3 survey_analysis.py --input example_survey_large.csv --industry "Apparel" --cache .analysis_cache.json


Notes:

The script passes each question header to the model, so headers like
"Shipping and delivery" or "Precio y valor" influence the analysis.

Results are cached on disk so repeated texts do not re-call the API.

Long answers are truncated and max_tokens is capped to control spend.

For rate limits and retries, see OpenAI docs. 
OpenAI Platform
+1

Approximate API cost

Pricing is per million tokens and can change. Check OpenAI pricing for current rates. 
OpenAI Platform

Example estimate for 3,000 reviews with 5 questions each (about 15,000 answers), using a cost-efficient model:

Assume about 120 input tokens and 6 output tokens per answer (short JSON).

With gpt-5-mini class pricing, this is typically low single-digit USD for this volume. Always confirm against the live pricing table. 
OpenAI Platform

Requirements for API Mode:

OpenAI account and API key

openai Python package (installed via requirements.txt)

Reasonable token caps and caching enabled (already built in)
