"""Stakeholder Excel workbook."""
from __future__ import annotations

import pandas as pd

from .pipeline import Analysis

PRIVACY_NOTE = (
    "Identifying columns (names, emails, locations, IPs, order ids) were removed at ingest. "
    "Spend is shown only as a band. Emails, phones, URLs and handles inside answers are masked."
)


def _widths(ws, df: pd.DataFrame, start_col: int = 0, max_w: int = 60) -> None:
    for i, col in enumerate(df.columns):
        longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
        ws.set_column(start_col + i, start_col + i, min(max(10, longest + 2), max_w))


def write(an: Analysis, path: str, include_quotes: bool = False) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as xw:
        wb = xw.book
        pct = wb.add_format({"num_format": "0.0%"})
        bold = wb.add_format({"bold": True})
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})

        m = an.manifest
        about = pd.DataFrame([
            ("Dataset", m["label"]), ("Rows", m["rows"]), ("Backend", m["backend"]),
            ("Model", m["model"] or "-"), ("Codebook version", m["codebook_version"]),
            ("Open answers received", m["text_answers_received"]), ("Open answers coded", m["text_answers_coded"]),
            ("Junk answers removed", sum(m["junk_removed"].values())), ("Answers with typo fixes", m["answers_with_typo_fixes"]),
            ("Failed classifications", m["failed"]),
            ("Minimum base for shares", m["min_n"]), ("Run at", m["run_at"]), ("Privacy", PRIVACY_NOTE),
        ], columns=["Field", "Value"])
        about.to_excel(xw, sheet_name="About", index=False)
        xw.sheets["About"].set_column(0, 0, 26, bold)
        xw.sheets["About"].set_column(1, 1, 90, wrap)

        # Themes, with one bar chart per reportable question.
        if not an.themes.empty:
            an.themes.to_excel(xw, sheet_name="Themes", index=False)
            ws = xw.sheets["Themes"]
            _widths(ws, an.themes)
            ws.set_column(7, 7, 10, pct)
            chart_row = 1
            row = 1
            for question, grp in an.themes.groupby("question", sort=False):
                first, last = row, row + len(grp) - 1
                row = last + 1
                if not grp.reportable.iloc[0]:
                    continue
                chart = wb.add_chart({"type": "bar"})
                chart.add_series({
                    "categories": ["Themes", first, 3, last, 3],
                    "values": ["Themes", first, 7, last, 7],
                    "data_labels": {"value": True, "num_format": "0%"},
                })
                chart.set_title({"name": question[:80], "name_font": {"size": 10}})
                chart.set_legend({"none": True})
                chart.set_y_axis({"reverse": True})
                chart.set_x_axis({"num_format": "0%"})
                chart.set_size({"width": 620, "height": 60 + 22 * len(grp)})
                ws.insert_chart(chart_row, 10, chart)
                chart_row += 4 + len(grp) + 2

        if not an.choices.empty:
            an.choices.to_excel(xw, sheet_name="Choice questions", index=False)
            _widths(xw.sheets["Choice questions"], an.choices)
            xw.sheets["Choice questions"].set_column(5, 5, 10, pct)

        if not an.gates.empty:
            an.gates.to_excel(xw, sheet_name="PDP gate", index=False)
            _widths(xw.sheets["PDP gate"], an.gates)
            xw.sheets["PDP gate"].set_column(3, 3, 10, pct)

        if not an.sentiment.empty:
            an.sentiment.to_excel(xw, sheet_name="Sentiment", index=False)
            ws = xw.sheets["Sentiment"]
            _widths(ws, an.sentiment)
            chart = wb.add_chart({"type": "bar", "subtype": "percent_stacked"})
            colors = {"Positive": "#2a9d8f", "Neutral": "#9aa5b1", "Negative": "#d1495b", "Mixed": "#e9c46a"}
            n = len(an.sentiment)
            for j, label in enumerate(["Positive", "Neutral", "Negative", "Mixed"], start=1):
                chart.add_series({"name": label, "categories": ["Sentiment", 1, 0, n, 0],
                                  "values": ["Sentiment", 1, j, n, j], "fill": {"color": colors[label]}})
            chart.set_title({"name": "Sentiment mix", "name_font": {"size": 10}})
            ws.insert_chart(n + 3, 0, chart)

        if not an.products.empty:
            top = an.products.head(30)
            top.to_excel(xw, sheet_name="Products x theme")
            xw.sheets["Products x theme"].set_column(0, 0, 40)

        cols = [c for c in an.coded.columns if include_quotes or c not in {"answer", "answer_clean"}]
        if an.dataset.meta.get("rating") is None:
            cols = [c for c in cols if c != "rating_sentiment"]
        an.coded[cols].to_excel(xw, sheet_name="Coded answers", index=False)
        _widths(xw.sheets["Coded answers"], an.coded[cols])

        cb = pd.DataFrame([(t.id, t.label, t.label_es, t.definition) for t in an.codebook.themes],
                          columns=["id", "label", "label_es", "definition"])
        cb.to_excel(xw, sheet_name="Codebook", index=False)
        _widths(xw.sheets["Codebook"], cb, max_w=80)
