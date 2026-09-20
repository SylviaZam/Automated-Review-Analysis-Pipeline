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


# --- Claude backend ----------------------------------------------------------------
class _Block:
    def __init__(self, type_, text=""):
        self.type, self.text = type_, text


class _FakeClaudeMessages:
    def __init__(self, payload=None, stop_reason="end_turn", error=None):
        self.payload, self.stop_reason, self.error = payload, stop_reason, error
        self.calls = []

    def create(self, **kwargs):
        self.calls.append(kwargs)
        if self.error is not None:
            raise self.error
        resp = type("Resp", (), {})()
        resp.stop_reason = self.stop_reason
        resp.content = [_Block("thinking"), _Block("text", json.dumps(self.payload or {}))]
        return resp


def _fake_claude(tmp_path, messages, model=None):
    from voc.classify_claude import ClaudeClassifier

    client = type("Client", (), {"beta": type("Beta", (), {"messages": messages})()})()
    return ClaudeClassifier(CB, "wellness", model=model, cache_path=str(tmp_path / "c.json"), client=client)


def test_claude_request_shape_and_parsing(tmp_path):
    payload = {"results": [{"index": 0, "themes": ["efficacy_results"], "sentiment": "n/a"},
                           {"index": 1, "themes": ["quality"], "sentiment": "Positive"}]}
    messages = _FakeClaudeMessages(payload)
    clf = _fake_claude(tmp_path, messages)
    res = clf.classify([Item("hesitation", "¿Qué te preocupaba?", "que no funcione"),
                        Item("review", "body", "excelente calidad")])
    assert [r.themes for r in res] == [["efficacy_results"], ["quality"]]
    assert res[1].sentiment == "Positive"
    (call,) = messages.calls
    assert call["model"] == "claude-opus-5"
    assert call["output_config"]["effort"] == "low"
    fmt = call["output_config"]["format"]
    assert fmt["type"] == "json_schema"
    assert "quality" in fmt["schema"]["properties"]["results"]["items"]["properties"]["themes"]["items"]["enum"]
    assert call["fallbacks"] == "default" and call["betas"] == ["server-side-fallback-2026-07-01"]
    assert call["system"][0]["cache_control"] == {"type": "ephemeral"}
    assert "temperature" not in call


def test_claude_haiku_omits_effort_and_fallbacks(tmp_path):
    messages = _FakeClaudeMessages({"results": [{"index": 0, "themes": ["other"], "sentiment": "n/a"}]})
    _fake_claude(tmp_path, messages, model="claude-haiku-4-5").classify([Item("hesitation", "q", "algo")])
    (call,) = messages.calls
    assert "effort" not in call["output_config"]
    assert "fallbacks" not in call and "betas" not in call


@pytest.mark.parametrize("stop_reason", ["refusal", "max_tokens"])
def test_claude_unparseable_batches_are_failed(tmp_path, stop_reason):
    clf = _fake_claude(tmp_path, _FakeClaudeMessages({}, stop_reason=stop_reason))
    (res,) = clf.classify([Item("review", "body", "excelente")])
    assert res.status == "failed"


def test_claude_auth_errors_are_not_retried(tmp_path, monkeypatch):
    import anthropic

    err = anthropic.AuthenticationError.__new__(anthropic.AuthenticationError)
    messages = _FakeClaudeMessages(error=err)
    monkeypatch.setattr("voc.classify.time.sleep", lambda s: None)
    with pytest.raises(anthropic.AuthenticationError):
        _fake_claude(tmp_path, messages).classify([Item("review", "body", "excelente")])
    assert len(messages.calls) == 1


def test_claude_missing_credentials_is_fatal(tmp_path):
    messages = _FakeClaudeMessages(error=TypeError("Could not resolve authentication method."))
    with pytest.raises(RuntimeError, match="ANTHROPIC_API_KEY"):
        _fake_claude(tmp_path, messages).classify([Item("review", "body", "excelente")])
    assert len(messages.calls) == 1


def test_cli_fails_loudly_when_every_answer_fails(tmp_path, synthetic, monkeypatch, capsys):
    class AlwaysFails:
        name = "broken"

        def classify(self, items):
            from voc.classify import Result
            return [Result(["other"], "n/a", status="failed") for _ in items]

    monkeypatch.setattr("voc.cli.make_classifier", lambda *a, **k: AlwaysFails())
    assert main(["run", str(synthetic["apparel_pdp_poll.csv"]), "--out", str(tmp_path / "o")]) == 1
    assert "every answer failed" in capsys.readouterr().err
    assert not (tmp_path / "o").exists()


# --- diagnose ----------------------------------------------------------------------
def test_diagnose_flags_multi_select_and_gaps(tmp_path):
    from voc import diagnose as diag

    options = ["Precio Justo", "Entrega Rapida", "Servicio al Cliente"]
    path = tmp_path / "brand.csv"
    pd.DataFrame({
        "Email": [f"demo+{i}@example.com" for i in range(40)],
        "¿Cuál fue la razón principal por la que compraste?": [", ".join(options[: 1 + i % 3]) for i in range(40)],
        "¿Qué te causaba incertidumbre antes de comprar?": [f"cocodrilo bailarín numero {i} en el tejado" for i in range(40)],
    }).to_csv(path, index=False)
    report = diag.diagnose(ingest.load(path, label="Brand"), CB)
    kinds = {q["question"][:20]: q["kind"] for q in report["questions"]}
    assert "multi" in kinds.values()
    issues = " ".join(f["message"] for f in report["findings"] if f["level"] == "issue")
    assert "cannot place" in issues
    assert any(c["phrase"] in {"cocodrilo", "bailarin", "cocodrilo bailarin"} for c in report["candidates"])
    md = diag.markdown(report)
    assert "Candidate vocabulary" in md
    assert "unassigned" in diag.keyword_stub(report)


def test_diagnose_flags_two_exports_pasted_together(tmp_path):
    from voc import diagnose as diag

    path = tmp_path / "pasted.csv"
    a = [f"respuesta larga distinta numero {i}" for i in range(30)] + [None] * 30
    b = [None] * 30 + [f"otra respuesta totalmente distinta {i}" for i in range(30)]
    pd.DataFrame({"¿Qué te preocupaba?": a, "Slide: ¿Qué te preocupaba?": b}).to_csv(path, index=False)
    report = diag.diagnose(ingest.load(path, label="Pasted"), CB)
    assert any("pasted side by side" in f["message"] for f in report["findings"])


def test_multi_select_options_are_tallied_per_respondent(tmp_path):
    path = tmp_path / "multi.csv"
    options = ["Precio Justo", "Entrega Rapida", "Servicio al Cliente"]
    pd.DataFrame({"¿Cuál fue la razón principal por la que compraste?": [", ".join(options[: 1 + i % 3]) for i in range(60)]}).to_csv(path, index=False)
    an = pipeline.run(ingest.load(path), RulesClassifier(CB), CB, min_n=1)
    assert an.coded.empty  # nothing was coded as free text
    tally = an.choices.set_index("option")["count"].to_dict()
    assert tally["Precio Justo"] == 60 and tally["Servicio al Cliente"] == 20


def test_eval_excludes_unclear_rows(tmp_path):
    path = tmp_path / "gold.csv"
    pd.DataFrame([
        {"id": "a", "role": "hesitation", "question": "q", "answer": "que no llegara", "gold_theme": "shipping_delivery", "gold_sentiment": ""},
        {"id": "b", "role": "hesitation", "question": "q", "answer": "asdf ???", "gold_theme": "unclear", "gold_sentiment": ""},
        {"id": "c", "role": "hesitation", "question": "q", "answer": "que no funcione", "gold_theme": "", "gold_sentiment": ""},
    ]).to_csv(path, index=False)
    gold = evaluate.load_gold(path)
    result = evaluate.score(gold, RulesClassifier(CB))
    assert result["n"] == 1 and result["coverage_gap"] == 2
    assert result["theme_accuracy"] == 1.0
    assert "unclear or left blank" in evaluate.markdown([result], "t")


# --- brand-specific themes ---------------------------------------------------------
def test_local_themes_are_prefixed_and_marked(tmp_path):
    path = tmp_path / "themes.json"
    path.write_text('{"themes":[{"id":"influencer","label":"Influencer","definition":"d","keywords":["bren"]}]}', encoding="utf-8")
    cb = Codebook.load(extra_themes=path)
    assert cb.local_ids == ["local_influencer"]
    assert cb.match("me inspira bren")[0] == "local_influencer"
    assert cb.match("el precio")[0] == "price_value"  # shared themes still win where they apply


def test_local_theme_cannot_collide_with_a_shared_one(tmp_path):
    path = tmp_path / "themes.json"
    path.write_text('{"themes":[{"id":"price_value","label":"x","keywords":["y"]}]}', encoding="utf-8")
    Codebook.load(extra_themes=path)  # local_price_value does not collide
    path.write_text('{"themes":[{"id":"local_x","keywords":["y"]},{"id":"x","keywords":["z"]}]}', encoding="utf-8")
    with pytest.raises(ValueError, match="collides"):
        Codebook.load(extra_themes=path)


def test_proposer_prefers_distinctive_phrases():
    from voc import diagnose as diag

    # "producto" is everywhere; "brenvita" only in the answers nothing matched.
    uncovered = ["me gusta brenvita y su producto"] * 6 + ["el producto llego"] * 4
    corpus = uncovered + ["buen producto"] * 200
    proposals = diag.propose_themes({"uncovered_examples": uncovered, "corpus_examples": corpus}, max_themes=3)
    ids = [p["id"] for p in proposals]
    assert any("brenvita" in i for i in ids)   # the distinctive word survives
    assert "producto" not in ids               # the word that is everywhere does not


def test_proposed_themes_file_round_trips_into_a_codebook(tmp_path):
    from voc import diagnose as diag

    uncovered = ["me encanta brenvita"] * 5
    proposals = diag.propose_themes({"uncovered_examples": uncovered, "corpus_examples": uncovered}, max_themes=2)
    path = tmp_path / "themes.json"
    path.write_text(diag.themes_file(proposals, "Brand"), encoding="utf-8")
    cb = Codebook.load(extra_themes=path)
    assert cb.local_ids and cb.match("me encanta brenvita")[0].startswith("local_")


def test_claude_proposer_counts_coverage_itself(tmp_path):
    from voc import propose_claude

    payload = {"themes": [
        {"id": "influencer_trust", "label": "Influencer trust", "definition": "d", "keywords": ["bren"]},
        {"id": "price_value", "label": "dup of a shared theme", "definition": "d", "keywords": ["precio"]},
    ]}
    messages = _FakeClaudeMessages(payload)
    client = type("C", (), {"beta": type("B", (), {"messages": messages})()})()
    uncovered = ["confio en bren", "bren lo recomienda", "nada que ver"]
    out = propose_claude.propose(uncovered, CB, "wellness", client=client)
    assert [t["id"] for t in out] == ["influencer_trust"]  # the duplicate is dropped
    assert out[0]["answers"] == 2  # measured from the answers, not claimed by the model
    assert messages.calls[0]["output_config"]["format"]["type"] == "json_schema"


def test_nothing_is_hidden_by_default(tmp_path):
    path = tmp_path / "small.csv"
    pd.DataFrame({"¿Qué te preocupaba?": ["que no llegara", "el precio", "que fuera falso"]}).to_csv(path, index=False)
    an = pipeline.run(ingest.load(path), RulesClassifier(CB), CB)
    assert an.themes.reportable.all()
    html_out = report_html.render(an)
    assert "not charted" not in html_out
    assert "small base" in html_out  # marked, not hidden
