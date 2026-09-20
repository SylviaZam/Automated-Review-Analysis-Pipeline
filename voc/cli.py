"""Command line: `python -m voc {run,eval,synth,label-kit}`."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd

from . import __version__, diagnose as diagnose_mod, evaluate, ingest, pipeline, report_excel, report_html, synth
from .classify import make_classifier
from .codebook import Codebook

try:
    from dotenv import load_dotenv  # type: ignore
except ImportError:  # pragma: no cover
    load_dotenv = None


def _json_default(value):
    return str(value)


def cmd_run(args) -> int:
    codebook = Codebook.load(args.codebook, args.extra_keywords, getattr(args, "extra_themes", None))
    overrides = json.loads(Path(args.roles).read_text(encoding="utf-8")) if args.roles else None
    ds = ingest.load(args.input, label=args.label, sheet=args.sheet, role_overrides=overrides)
    kwargs = {"model": args.model, "cache_path": args.cache} if args.backend != "rules" else {}
    classifier = make_classifier(args.backend, codebook, args.industry, **kwargs)
    try:
        an = pipeline.run(ds, classifier, codebook, min_n=args.min_n)
    except RuntimeError as exc:
        print(f"[error] {exc}", file=sys.stderr)
        return 1
    if an.manifest["text_answers_coded"] and an.manifest["failed"] == an.manifest["text_answers_coded"]:
        print(f"[error] every answer failed to classify with '{an.backend}'; no report written. "
              "Check credentials, model id and network.", file=sys.stderr)
        return 1

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    stem = args.label.lower().replace(" ", "_") if args.label else Path(args.input).stem
    report_excel.write(an, str(out / f"{stem}_report.xlsx"), include_quotes=args.quotes)
    report_html.write(an, str(out / f"{stem}_summary.html"), title=args.title)
    (out / f"{stem}_manifest.json").write_text(json.dumps(an.manifest, indent=2, ensure_ascii=False, default=_json_default), encoding="utf-8")

    print(f"[ok] {ds.label}: {an.manifest['rows']} rows, {len(ds.questions)} questions, "
          f"{an.manifest['text_answers_coded']} open answers coded with '{an.backend}' "
          f"({an.manifest['failed']} failed), {ds.dropped_columns} identifying columns dropped")
    for q in ds.questions:
        print(f"     - {q.role:<16} {q.kind:<6} {q.text[:70]}")
    print(f"[ok] wrote {out}/{stem}_report.xlsx, _summary.html, _manifest.json")
    return 0


def cmd_eval(args) -> int:
    gold = evaluate.load_gold(args.gold)
    corpus = None
    if args.corpus:
        # Let the cleaner learn vocabulary from the full export(s), as `voc run` does.
        corpus = []
        for path in args.corpus:
            frame = pd.read_csv(path, dtype=str)
            corpus += [v for col in frame.columns for v in frame[col].dropna() if isinstance(v, str)]
    codebook = Codebook.load(args.codebook, args.extra_keywords)
    scores = []
    for backend in args.backends:
        if backend == "v1":
            scores.append(evaluate.score(gold, evaluate.V1DemoClassifier(), clean=False))
            continue
        kwargs = {"model": args.model, "cache_path": args.cache} if backend != "rules" else {}
        try:
            clf = make_classifier(backend, codebook, args.industry, **kwargs)
            scores.append(evaluate.score(gold, clf, corpus=corpus))
        except RuntimeError as exc:
            print(f"[error] {backend}: {exc}", file=sys.stderr)
            return 1
    md = evaluate.markdown(scores, args.title)
    if args.out:
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.out).write_text(md, encoding="utf-8")
        if args.predictions:
            for s in scores:
                s["predictions"].to_csv(Path(args.out).with_name(f"predictions_{s['backend']}.csv"), index=False)
        print(f"[ok] wrote {args.out}")
    print(md.split("\n## ")[0])
    return 0


def cmd_diagnose(args) -> int:
    codebook = Codebook.load(args.codebook, args.extra_keywords)
    ds = ingest.load(args.input, label=args.label, sheet=args.sheet)
    report = diagnose_mod.diagnose(ds, codebook, min_count=args.min_count)
    stem = (args.label or Path(args.input).stem).lower().replace(" ", "_")
    if args.propose_themes:
        if args.propose_backend == "claude":
            from . import propose_claude

            try:
                report["proposals"] = propose_claude.propose(
                    report["uncovered_examples"], codebook, args.industry,
                    max_themes=args.propose_themes, model=args.model)
            except RuntimeError as exc:
                print(f"[error] {exc}", file=sys.stderr)
                return 1
        else:
            report["proposals"] = diagnose_mod.propose_themes(report, max_themes=args.propose_themes)
    paths = diagnose_mod.write(report, args.out, stem, show_examples=not args.no_examples)
    if args.propose_themes:
        themes_path = Path(args.out) / f"{stem}_themes_draft.json"
        themes_path.write_text(diagnose_mod.themes_file(report["proposals"], ds.label), encoding="utf-8")
        paths.append(themes_path)
        print(f"[ok] proposed {len(report['proposals'])} brand theme(s) -> {themes_path}")
        print("     Review them, then: python -m voc run <export> --extra-themes " + str(themes_path))
    issues = [f for f in report["findings"] if f["level"] == "issue"]
    print(diagnose_mod.markdown(report, show_examples=not args.no_examples).split("## Questions")[0])
    print(f"[ok] wrote {paths[0]} and {paths[1]}")
    if issues:
        print(f"[warn] {len(issues)} thing(s) to fix before labelling or reporting this brand")
    return 0


def cmd_synth(args) -> int:
    paths = synth.generate(args.out, seed=args.seed, gold_size=args.gold_size)
    for name, path in paths.items():
        print(f"[ok] {path}")
    return 0


def cmd_label_kit(args) -> int:
    """Sample open answers from a real export into a labelling sheet (kept local)."""
    ds = ingest.load(args.input, label=args.label, sheet=args.sheet)
    rows = []
    for q in ds.questions:
        if q.kind != "text":
            continue
        for value in ds.frame[q.column].dropna():
            if isinstance(value, str) and value.strip():
                rows.append({"role": q.role, "question": q.text, "answer": value})
    frame = pd.DataFrame(rows)
    if frame.empty:
        print("no open-text answers found", file=sys.stderr)
        return 1
    frame = frame.sample(n=min(args.n, len(frame)), random_state=args.seed).reset_index(drop=True)
    frame.insert(0, "id", [f"{args.label or 'x'}-{i:04d}".replace(" ", "") for i in range(1, len(frame) + 1)])
    frame["gold_theme"] = ""
    frame["gold_sentiment"] = ""
    frame["gold_theme_2"] = ""
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    frame.to_csv(args.out, index=False)
    print(f"[ok] {len(frame)} answers to label -> {args.out}\n"
          f"     Theme ids: {', '.join(Codebook.load().ids)}\n"
          f"     Keep this file out of git: it contains real (scrubbed) customer text.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="voc", description="Survey and review exports -> coded themes, Excel and HTML reports.")
    p.add_argument("--version", action="version", version=__version__)
    sub = p.add_subparsers(dest="command", required=True)

    r = sub.add_parser("run", help="analyse one export")
    r.add_argument("input", help="CSV or XLSX export")
    r.add_argument("--out", default="output")
    r.add_argument("--label", help="dataset name shown in reports, e.g. 'BFit post-purchase'")
    r.add_argument("--title", help="report heading")
    r.add_argument("--industry", default="e-commerce")
    r.add_argument("--backend", choices=["rules", "claude", "openai"], default="rules")
    r.add_argument("--model", help="model id for a model backend (claude: $VOC_CLAUDE_MODEL or claude-opus-5; openai: $VOC_OPENAI_MODEL or gpt-4o-mini)")
    r.add_argument("--cache", default=".voc_cache.json")
    r.add_argument("--sheet", help="sheet name for XLSX inputs")
    r.add_argument("--roles", help="JSON file mapping column header -> role, for unusual exports")
    r.add_argument("--codebook", help="alternative codebook JSON")
    r.add_argument("--extra-keywords", help="JSON {theme_id: [keywords]} for brand-specific terms")
    r.add_argument("--extra-themes", help="JSON of brand-specific themes from `voc diagnose --propose-themes`")
    r.add_argument("--min-n", type=int, default=pipeline.MIN_N_DEFAULT,
                   help="hold back questions with fewer answers than this (default 0: chart everything)")
    r.add_argument("--mask-small-counts", type=int, default=0, metavar="N",
                   help="print counts below N as '<N' (default 0: show every count)")
    r.add_argument("--quotes", action="store_true", help="include scrubbed answers in the Excel (internal use only)")
    r.set_defaults(func=cmd_run)

    e = sub.add_parser("eval", help="score backends against a gold file")
    e.add_argument("gold")
    e.add_argument("--backends", nargs="+", default=["v1", "rules"], choices=["v1", "rules", "claude", "openai"])
    e.add_argument("--model", help="model id for the model backend(s)")
    e.add_argument("--industry", default="e-commerce")
    e.add_argument("--cache", default=".voc_cache.json")
    e.add_argument("--codebook")
    e.add_argument("--extra-keywords")
    e.add_argument("--corpus", nargs="*", help="CSV exports the gold answers came from (vocabulary for typo repair)")
    e.add_argument("--out", help="write a markdown report here")
    e.add_argument("--predictions", action="store_true", help="also write per-item predictions next to the report")
    e.add_argument("--title", default="Classifier evaluation")
    e.set_defaults(func=cmd_eval)

    g = sub.add_parser("diagnose", help="check one export before coding it: question kinds, coverage gaps, candidate vocabulary")
    g.add_argument("input")
    g.add_argument("--out", default="output")
    g.add_argument("--label")
    g.add_argument("--sheet")
    g.add_argument("--codebook")
    g.add_argument("--extra-keywords")
    g.add_argument("--min-count", type=int, default=3, help="how often a phrase must recur to be a candidate")
    g.add_argument("--propose-themes", type=int, nargs="?", const=6, default=0, metavar="N",
                   help="also propose up to N brand-specific themes from the answers nothing matched")
    g.add_argument("--propose-backend", choices=["phrases", "claude"], default="phrases",
                   help="how to propose them: offline phrase counting (default) or Claude reading a sample")
    g.add_argument("--industry", default="e-commerce", help="context passed to the Claude proposer")
    g.add_argument("--model", help="model id for --propose-backend claude")
    g.add_argument("--no-examples", action="store_true", help="omit example answers (real customer text) from the report")
    g.set_defaults(func=cmd_diagnose)

    s = sub.add_parser("synth", help="generate the synthetic demo datasets and gold labels")
    s.add_argument("--out", default="data/synthetic")
    s.add_argument("--seed", type=int, default=7)
    s.add_argument("--gold-size", type=int, default=400)
    s.set_defaults(func=cmd_synth)

    k = sub.add_parser("label-kit", help="sample real answers into a local labelling sheet")
    k.add_argument("input")
    k.add_argument("--out", required=True)
    k.add_argument("--label")
    k.add_argument("--sheet")
    k.add_argument("-n", type=int, default=100)
    k.add_argument("--seed", type=int, default=11)
    k.set_defaults(func=cmd_label_kit)
    return p


def main(argv: list[str] | None = None) -> int:
    if load_dotenv:
        load_dotenv()
    args = build_parser().parse_args(argv)
    return args.func(args)
