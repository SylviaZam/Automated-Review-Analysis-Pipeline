import json
from pathlib import Path

import pandas as pd
import pytest

from voc import evaluate, ingest, pipeline, report_excel, report_html, synth
from voc.classify import Item, OpenAIClassifier, RulesClassifier
from voc.cli import main
from voc.codebook import Codebook

CB = Codebook.load()


@pytest.fixture(scope="module")
def synthetic(tmp_path_factory):
    out = tmp_path_factory.mktemp("synthetic")
    return synth.generate(out, seed=7, gold_size=200)


def test_synth_is_deterministic(tmp_path, synthetic):
    again = synth.generate(tmp_path, seed=7, gold_size=200)
    assert again["gold_labels.csv"].read_text() == synthetic["gold_labels.csv"].read_text()


def test_ingest_drops_identifying_columns(synthetic):
    ds = ingest.load(synthetic["wellness_post_purchase.csv"], label="Demo")
    assert ds.dropped_columns == 3
    assert not {"Full Name", "Email", "City"} & set(ds.frame.columns)
    assert set(ds.frame[ds.meta["spend"]]) <= {"<500", "500-999", "1000-2499", "2500-4999", "5000+", "unknown"}
    roles = [q.role for q in ds.questions]
    assert roles == ["purchase_reason", "choice_driver", "hesitation", "discovery", "visit_trigger", "ad_location"]


def test_reviews_merge_title_and_drop_reviewer(synthetic):
    ds = ingest.load(synthetic["fragrance_reviews.csv"])
    assert "title" not in ds.frame.columns
    assert not any("reviewer" in c for c in ds.frame.columns)
    assert ds.meta["rating"] == "rating"


def test_pipeline_outputs_contain_no_pii(tmp_path, synthetic):
    for name in ("wellness_post_purchase.csv", "apparel_pdp_poll.csv", "fragrance_reviews.csv"):
        ds = ingest.load(synthetic[name], label="Demo")
        an = pipeline.run(ds, RulesClassifier(CB), CB)
        xlsx, html_path = tmp_path / f"{name}.xlsx", tmp_path / f"{name}.html"
        report_excel.write(an, str(xlsx), include_quotes=False)
        report_html.write(an, str(html_path))
        text = html_path.read_text(encoding="utf-8")
        sheets = pd.read_excel(xlsx, sheet_name=None)
        dumped = " ".join(df.to_csv() for df in sheets.values())
        for needle in ("@example.com", "Demo Ejemplo", "203.0.113."):
            assert needle not in text
            assert needle not in dumped
        assert "answer" not in sheets["Coded answers"].columns


def test_pipeline_gate_rate_and_bases(synthetic):
    ds = ingest.load(synthetic["apparel_pdp_poll.csv"])
    an = pipeline.run(ds, RulesClassifier(CB), CB)
    gate = an.gates.iloc[0]
    assert gate.shown == 900
    assert 0.12 < gate.yes_rate < 0.24
    top = an.themes[an.themes.role == "pdp_blocker"].iloc[0]
    assert top.theme_id == "size_fit"


def test_small_bases_are_not_reportable(synthetic):
    ds = ingest.load(synthetic["wellness_post_purchase.csv"])
    ds.frame = ds.frame.head(20)
    an = pipeline.run(ds, RulesClassifier(CB), CB, min_n=30)
    assert not an.themes.reportable.any()
    assert "not charted" in report_html.render(an)


def test_junk_is_excluded_from_bases(tmp_path):
    path = tmp_path / "junk.csv"
    pd.DataFrame({"Email": ["a@example.com"] * 4,
                  "¿Qué te preocupaba?": ["Que no llegara", "asdfgh", "test", "Que fuera fraude"]}).to_csv(path, index=False)
    an = pipeline.run(ingest.load(path), RulesClassifier(CB), CB, min_n=1)
    assert an.manifest["junk_removed"] == {"keyboard_mash": 1, "test_or_greeting": 1}
    assert set(an.themes.base) == {2}


def test_evaluate_rules_beats_v1(synthetic):
    gold = evaluate.load_gold(synthetic["gold_labels.csv"])
    v1 = evaluate.score(gold, evaluate.V1DemoClassifier(), clean=False)
    v2 = evaluate.score(gold, RulesClassifier(CB))
    assert v2["theme_accuracy"] > v1["theme_accuracy"] + 0.4
    assert v2["sentiment_accuracy"] >= v1["sentiment_accuracy"]


def test_cohen_kappa():
    assert evaluate.cohen_kappa(["a", "b", "a", "b"], ["a", "b", "a", "b"]) == 1.0
    assert evaluate.cohen_kappa(["a", "a", "b", "b"], ["a", "b", "a", "b"]) == 0.0


def test_load_gold_rejects_unknown_theme(tmp_path):
    path = tmp_path / "gold.csv"
    pd.DataFrame([{"id": 1, "role": "hesitation", "question": "q", "answer": "a", "gold_theme": "nope"}]).to_csv(path, index=False)
    with pytest.raises(ValueError):
        evaluate.load_gold(path)


class _FakeCompletions:
    def __init__(self, payload=None, fail=False):
        self.payload, self.fail, self.calls = payload, fail, 0

    def create(self, **kwargs):
        self.calls += 1
        if self.fail:
            raise RuntimeError("boom")
        schema = kwargs["response_format"]["json_schema"]
        assert schema["strict"] is True
        enum = schema["schema"]["properties"]["results"]["items"]["properties"]["themes"]["items"]["enum"]
        assert "shipping_delivery" in enum and "other" in enum

        class Msg:
            content = json.dumps(self.payload)

        class Choice:
            message = Msg()

        class Resp:
            choices = [Choice()]

        return Resp()


def _fake_openai(monkeypatch, tmp_path, completions):
    monkeypatch.setenv("OPENAI_API_KEY", "test-not-a-real-key")
    clf = OpenAIClassifier(CB, "wellness", cache_path=str(tmp_path / "cache.json"))
    clf.client = type("C", (), {"chat": type("Chat", (), {"completions": completions})()})()
    monkeypatch.setattr("voc.classify.time.sleep", lambda s: None)
    return clf


def test_openai_backend_parses_and_caches(monkeypatch, tmp_path):
    payload = {"results": [{"index": 0, "themes": ["shipping_delivery", "made_up"], "sentiment": "Negative"}]}
    completions = _FakeCompletions(payload)
    clf = _fake_openai(monkeypatch, tmp_path, completions)
    items = [Item("hesitation", "¿Qué te preocupaba?", "que no llegara")]
    (res,) = clf.classify(items)
    assert res.themes == ["shipping_delivery"]
    assert res.sentiment == "n/a"  # not an opinion question
    cache = json.loads((tmp_path / "cache.json").read_text())
    assert "que no llegara" not in json.dumps(cache)  # cache stores hashes and labels only
    clf.classify(items)
    assert completions.calls == 1


def test_openai_failures_are_flagged_not_neutral(monkeypatch, tmp_path):
    clf = _fake_openai(monkeypatch, tmp_path, _FakeCompletions(fail=True))
    (res,) = clf.classify([Item("review", "body", "excelente")])
    assert res.status == "failed"


def test_cli_run_and_eval(tmp_path, synthetic, capsys):
    assert main(["run", str(synthetic["fragrance_reviews.csv"]), "--out", str(tmp_path), "--label", "Demo"]) == 0
    manifest = json.loads((tmp_path / "demo_manifest.json").read_text())
    assert manifest["backend"] == "rules" and manifest["rows"] == 700
    assert main(["eval", str(synthetic["gold_labels.csv"]), "--out", str(tmp_path / "eval.md")]) == 0
    assert "rules" in Path(tmp_path / "eval.md").read_text()
