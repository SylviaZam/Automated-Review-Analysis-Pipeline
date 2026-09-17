# Method

## 1. The instrument

The client stores asked the same small set of post-purchase questions, worded slightly differently from brand to brand. These questions follow the logic of jobs-to-be-done interviews: what triggered the purchase, what else was considered, and what nearly stopped it.

| Role | Typical wording (ES) | English | Analysed as |
|---|---|---|---|
| `purchase_reason` | Cuéntanos, ¿cuál fue la razón principal por la que decidiste comprar? / ¿Qué problema buscas resolver? | Main reason / problem to solve | codebook themes |
| `choice_driver` | ¿Qué te hizo comprar con nosotros antes que otra marca? | Why us over other brands | codebook themes |
| `hesitation` | ¿Qué te causaba incertidumbre antes de comprar? | What almost stopped you | codebook themes |
| `competitor` | ¿Tenías en mente alguna otra marca? | Other brands considered | tally, or themes if free text |
| `discovery` | ¿Cómo nos conociste por primera vez? | How you first heard of us | tally |
| `visit_trigger` | ¿Y qué te trajo a nuestra página hoy? | What brought you today | tally |
| `ad_location` | ¿Dónde estaba el anuncio? | Where the ad was | tally |
| `recipient` | ¿Para quién compraste esto hoy? | Who it was for | tally |
| `awareness_time` | ¿Hace cuánto tiempo sabías de nosotros? | Time from awareness to purchase | tally |
| `pdp_gate` | ¿Tienes alguna inquietud o duda que te dificulte comprar? | Is anything stopping you? (yes/no) | rate |
| `pdp_blocker` | Cuéntanos, ¿cuál es? | What is it? | codebook themes |
| `review` | review title + body | Product review | themes + sentiment |
| `open_feedback` | anything else, including v1's Q1..Qn | | themes + sentiment |

Roles are assigned from the header wording (`voc/questions.py`). Whether a column is free text, a pick-list, or a yes/no gate is inferred from its answers. A pick-list is detected when a few options cover at least 80% of answers; a gate when at least 90% of answers are yes/no. When an analyst pasted a second export of the same survey beside the first, the duplicate column is dropped.

**Sentiment is only computed for opinion questions** (reviews and open feedback). A worry is not a negative review, and "what almost stopped you?" invites concerns by design. Scoring those answers as negative would make every brand look worse than it is.

## 2. From manual coding to a codebook

The first analyses were done by hand. Open answers were copied into a sheet, grouped by affinity, and tallied per question and brand: "Calidad y duración", "Desconocimiento o inseguridad de la marca", "Precio", "Recomendaciones y reseñas". Themes that recurred across brands became the shared codebook ([codebook.md](codebook.md)). One-off themes stay brand-specific and are added through `--extra-keywords`.

## 3. Cleaning

See the README section on cleaning. Junk answers are excluded from every base and counted in the manifest, so a share always means "of the people who gave a real answer".

## 4. Reporting rules

- Shares are of respondents who gave a real answer to that question. Answers can carry up to three themes, so shares can add up to more than 100%.
- Questions with a base under 30 are listed but not charted (`--min-n`).
- Public outputs never include verbatim answers. The analyst workbook includes them only with `--quotes`.
- Every run writes a manifest: source hash and encoding, rows, identifying columns dropped, question roles, backend and model, codebook version, junk removed, typo fixes, and failed classifications.

## 5. From findings to decisions

In client work, the coded themes fed four kinds of decisions:

| Finding type | Decision it informed |
|---|---|
| Product-page blockers (sizing, stock, shipping, payment) | Product-page content and layout changes; FAQ and size-guide placement |
| Top worries before purchase (efficacy, delivery, trust) | Reassurance in checkout and lifecycle emails: delivery promises, guarantees, reviews near the buy button |
| Why buyers chose the brand | Messaging hierarchy in ads, welcome flows and product pages |
| Discovery and visit triggers | Channel and campaign mix |
