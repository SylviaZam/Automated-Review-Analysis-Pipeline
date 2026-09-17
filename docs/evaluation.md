# Evaluation

## Answer key format

`voc eval gold.csv` expects:

| column | meaning |
|---|---|
| `id` | any unique id |
| `role` | question role (see methodology) |
| `question` | question text |
| `answer` | the customer's answer |
| `gold_theme` | the single most important theme id from the codebook, `other` for a real answer that fits none, or `unclear` when the answer does not answer the question |
| `gold_sentiment` | Positive / Neutral / Negative / Mixed, only for `review` and `open_feedback` rows |
| `gold_theme_2` | optional: a second, independent label of the same answer, used for Cohen's kappa |

## Metrics

- **Theme accuracy (primary)**: the predicted primary theme equals the gold theme.
- **Theme any-match**: the gold theme appears anywhere in the predicted themes.
- **Theme macro-F1**: the average F1 across themes that appear in the key, so rare themes count as much as common ones.
- **Sentiment accuracy and macro-F1**: computed on opinion rows only.
- **Positives called Negative**: the error that hurts a brand most in a report.
- **Cohen's kappa between human labels**: how consistent the labelling is. A classifier cannot be meaningfully better than the humans it is scored against.

Junk answers are removed by the cleaning step before scoring. They count as `other` / Neutral, which is also what a coder would label them.

Rows the coder marked `unclear` (or left blank) are excluded from accuracy and reported as a coverage gap. Forcing a label on an answer that does not answer the question measures the wrong thing; the size of that gap is a finding about the survey instrument, and it belongs in the diagnosis report rather than in the accuracy figure.

## Protocol for the real-data benchmark

1. `voc label-kit <export> --out private/labels.csv -n 200` samples answers (repeat per brand, then combine).
2. Run `voc diagnose` on each export first, and fix what it flags. Multi-select questions must not be in the sample at all.
3. Label `gold_theme` and `gold_sentiment` without looking at any model output. Use `unclear` plus a note rather than guessing.
3. A week later, label the first 50 rows again in `gold_theme_2` without looking at the first labels.
4. Run `voc eval private/labels.csv --backends v1 rules claude --corpus <exports...> --out private/eval.md`.
5. Publish only the aggregate table.

## Results so far

- Synthetic development set: [evaluation_synthetic.md](evaluation_synthetic.md). The numbers are optimistic, because the templates and rules were written by the same person and general gaps were fixed after reading errors.
- Real reviews, using star ratings as a partial key, and a comparison with a manual tally: see the README.
- Real hand-labelled benchmark: pending.
