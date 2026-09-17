# Industry packs

Extra keywords for one product category, added to the shared codebook at run time:

```bash
python -m voc run export.csv --extra-keywords codebooks/jewelry.json
```

Each file maps an existing theme id to words customers in that category use. The
shared themes stay the same, so brands remain comparable; only the vocabulary
changes. `voc diagnose` proposes additions for a specific brand, which belong in
a brand file (for example `codebooks/brand-loly.json`), not in these category packs.
