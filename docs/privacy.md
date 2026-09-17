# Privacy

This pipeline was built for customer data that belongs to someone else: the brands, and above all their customers. The rules below are enforced in code, not left to good intentions.

## What never enters the repository

- Real exports, labelled samples, or reports made from real exports. `.gitignore` blocks `data/*` (except `data/synthetic/`), `private/`, `output/`, and every `.xlsx` except the synthetic examples.
- Secrets. `.env` is ignored, and `scripts/check_no_pii.py` fails CI on API-key shaped strings.
- Client names. The guard compares every word in tracked files against SHA-256 hashes of blocked brand names, so the list itself reveals nothing.
- Contact data. The guard flags any email outside `example.com/.org/.net` and any IP outside the documentation range `203.0.113.0/24`.

## What happens to real data at run time

| Stage | Control |
|---|---|
| Ingest | Name, email, phone, address, city, province, IP, user agent, referrer/landing URLs, and order and customer ids are dropped before analysis. |
| Ingest | Total spent is converted to a band (`<500`, `500-999`, ...). |
| Ingest | Emails, phone numbers, URLs, @handles and order numbers typed inside answers are masked. |
| Model backend | Only the question text and the masked answer (truncated to 600 characters) are sent. The cache stores SHA-256 keys and labels, never text. |
| Reports | `--label` replaces the client name. Bases under 30 are not charted; counts under 5 print as `<5`. The HTML summary has no verbatims. |
| Labelling kit | Masked real answers, kept on the analyst's machine; the file header says so. |

## Limits

Rules cannot reliably catch a person's name typed inside an answer ("le pregunté a Mariana"). That is why verbatims stay out of shared outputs by default. Review any quote by hand before it leaves the team.

## Case studies

Public write-ups use code names ("Brand A, a Mexican wellness brand"), rounded or relative numbers, quarters instead of dates, and composite, paraphrased quotes that are labelled as such.
