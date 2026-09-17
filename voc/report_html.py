"""One-page, self-contained HTML insight report.

Public by design: aggregates only, no verbatim answers, bases under `min_n`
are flagged instead of charted, and counts under 5 print as "<5".
"""
from __future__ import annotations

import html

from .codebook import NO_CONCERN, OTHER
from .pipeline import Analysis, public_count

ROLE_TITLES = {
    "purchase_reason": "Why they bought",
    "choice_driver": "Why this brand over others",
    "hesitation": "What almost stopped them",
    "pdp_blocker": "Questions blocking purchase on the product page",
    "review": "What reviews talk about",
    "open_feedback": "Open feedback",
    "discovery": "How they first heard of the brand",
    "visit_trigger": "What brought them to the site today",
    "ad_location": "Where they saw the ad",
    "recipient": "Who the purchase was for",
    "awareness_time": "How long they knew the brand before buying",
    "competitor": "Other brands considered",
}

CSS = """
:root{--bg:#f7f6f2;--card:#ffffff;--ink:#1d2327;--muted:#5d6770;--line:#e3e1da;--bar:#2f6f73;--bar2:#c9d9d6;
--pos:#2a9d8f;--neu:#9aa5b1;--neg:#c8553d;--mix:#d9a441;--warn:#8a5a00}
@media (prefers-color-scheme:dark){:root{--bg:#15191c;--card:#1d2226;--ink:#e8eaec;--muted:#a3adb5;--line:#2e353a;
--bar:#6fb3b0;--bar2:#2d3f40;--warn:#e0b25a}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Inter,sans-serif}
main{max-width:980px;margin:0 auto;padding:40px 16px 64px}
h1{font-size:28px;margin:0 0 4px;letter-spacing:-.01em}h2{font-size:17px;margin:0 0 2px}
.sub{color:var(--muted);margin:0 0 28px}.q{color:var(--muted);font-size:13px;margin:0 0 14px}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:28px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px}
.stat b{display:block;font-size:24px;font-variant-numeric:tabular-nums}.stat span{color:var(--muted);font-size:13px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:16px}
@media (max-width:520px){.grid{grid-template-columns:1fr}}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px 18px 12px}
.row{display:grid;grid-template-columns:minmax(120px,190px) 1fr 56px;align-items:center;gap:10px;margin:7px 0;font-size:14px}
.track{background:var(--bar2);border-radius:4px;height:12px;overflow:hidden}.fill{background:var(--bar);height:100%}
.num{text-align:right;font-variant-numeric:tabular-nums;color:var(--muted)}
.base{color:var(--muted);font-size:12px;margin-top:8px}.flag{color:var(--warn);font-size:13px}
.stack{display:flex;height:14px;border-radius:4px;overflow:hidden;margin:8px 0}
.legend{display:flex;flex-wrap:wrap;gap:12px;font-size:12px;color:var(--muted)}
.dot{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:4px}
section{margin-top:28px}footer{margin-top:36px;color:var(--muted);font-size:12.5px;border-top:1px solid var(--line);padding-top:14px}
"""


def _bar_rows(rows: list[tuple[str, float, int]]) -> str:
    out = []
    for label, share, count in rows:
        out.append(
            f'<div class="row"><span>{html.escape(label)}</span>'
            f'<div class="track"><div class="fill" style="width:{share * 100:.1f}%"></div></div>'
            f'<span class="num" title="{public_count(count)} answers">{share:.0%}</span></div>'
        )
    return "".join(out)


def _card(title: str, question: str, body: str, base: int, reportable: bool, min_n: int) -> str:
    if not reportable:
        body = f'<p class="flag">Base of {public_count(base)} is below {min_n}; not charted.</p>'
    return (f'<div class="card"><h2>{html.escape(title)}</h2><p class="q">{html.escape(question)}</p>{body}'
            f'<div class="base">Base: {public_count(base)} answers</div></div>')


def render(an: Analysis, title: str | None = None, top: int = 6) -> str:
    m = an.manifest
    min_n = m["min_n"]
    cards = []

    for question, grp in an.themes.groupby("question", sort=False):
        role = grp.role.iloc[0]
        base = int(grp.base.iloc[0])
        body_rows = grp[~grp.theme_id.isin([OTHER])].head(top)
        rows = [(r.theme, r.share, r.mentions) for r in body_rows.itertuples()]
        cards.append(_card(ROLE_TITLES.get(role, "Open question"), question, _bar_rows(rows), base, bool(grp.reportable.iloc[0]), min_n))

    for question, grp in an.choices.groupby("question", sort=False):
        role = grp.role.iloc[0]
        base = int(grp.base.iloc[0])
        rows = [(str(r.option), r.share, r["count"]) for _, r in grp.head(top).iterrows()]
        title_suffix = " (pick several)" if grp.kind.iloc[0] == "multi" else ""
        cards.append(_card(ROLE_TITLES.get(role, "Choice question") + title_suffix, question,
                           _bar_rows(rows), base, bool(grp.reportable.iloc[0]), min_n))

    gate_html = ""
    for g in an.gates.itertuples():
        if g.reportable:
            gate_html += (f'<div class="stat"><b>{g.yes_rate:.0%}</b><span>of product-page visitors who answered '
                          f'said a question was stopping them</span></div>')

    sent_html = ""
    if not an.sentiment.empty:
        parts = []
        for r in an.sentiment.itertuples():
            base = int(r.base)
            if base < min_n:
                continue
            segs = "".join(
                f'<div style="width:{getattr(r, k) / base * 100:.1f}%;background:var(--{c})" title="{k}"></div>'
                for k, c in (("Positive", "pos"), ("Neutral", "neu"), ("Mixed", "mix"), ("Negative", "neg"))
            )
            pos = r.Positive / base
            parts.append(f'<p class="q">{html.escape(r.question)}: {pos:.0%} positive, base {public_count(base)}</p><div class="stack">{segs}</div>')
        if parts:
            legend = "".join(f'<span><i class="dot" style="background:var(--{c})"></i>{k}</span>'
                             for k, c in (("Positive", "pos"), ("Neutral", "neu"), ("Mixed", "mix"), ("Negative", "neg")))
            sent_html = f'<section><div class="card"><h2>Sentiment (opinion questions only)</h2>{"".join(parts)}<div class="legend">{legend}</div></div></section>'

    none_share = ""
    hes = an.themes[(an.themes.role == "hesitation") & (an.themes.theme_id == NO_CONCERN)]
    if not hes.empty and bool(hes.reportable.iloc[0]):
        none_share = f'<div class="stat"><b>{hes.share.iloc[0]:.0%}</b><span>of buyers said nothing worried them before buying</span></div>'

    stats = (
        f'<div class="stat"><b>{m["rows"]:,}</b><span>responses in this export</span></div>'
        f'<div class="stat"><b>{m["text_answers_coded"]:,}</b><span>open-text answers coded</span></div>'
        f'<div class="stat"><b>{len(m["questions"])}</b><span>questions analysed</span></div>'
        + gate_html + none_share
    )
    heading = html.escape(title or f"{m['label']}: what customers told us")
    junk = sum(m["junk_removed"].values())
    failed = f"; {m['failed']} answers failed classification and are excluded" if m["failed"] else ""
    failed += f"; {junk} junk answers (test entries, keyboard mashing, empty) excluded from bases" if junk else ""
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{heading}</title><style>{CSS}</style></head>
<body><main>
<h1>{heading}</h1>
<p class="sub">Voice-of-customer summary. Shares are of respondents who answered each question; one answer can carry up to three themes, so shares can add to more than 100%.</p>
<div class="stats">{stats}</div>
<div class="grid">{''.join(cards)}</div>
{sent_html}
<footer>Generated by {html.escape(m['tool'])} · backend: {html.escape(m['backend'])} · codebook v{html.escape(m['codebook_version'])} · {html.escape(m['run_at'])}{failed}.<br>
Privacy: identifying columns removed at ingest ({m['pii_columns_dropped']} dropped); no verbatim answers are shown; bases under {min_n} are not charted; counts under 5 print as "&lt;5".</footer>
</main></body></html>"""


def write(an: Analysis, path: str, title: str | None = None) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(render(an, title))
