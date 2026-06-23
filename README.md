# The Seventeen — Critical Minerals Primer & Investment Thesis

A self-directed research + investment project on the rare-earth / critical-minerals
supply chain, written as an investment piece (for hedge-fund / merchant-bank
recruiting), not a survey. The deliverables are a written memo, a supporting
basket-economics model, and the code that builds them.

**Last refreshed:** 2026-06-23

> **Status — working draft.** The primer / industry-context sections are written
> and cited. My own investment views, the single-name thesis, company financials,
> and live commodity prices are intentionally left as **clearly-marked placeholders**
> for me to fill in. See the [placeholder checklist](#placeholder-checklist) below.

---

## File map

| Path | What it is |
|---|---|
| `memo/the-seventeen.md` | **The main memo.** Primer (Part 1), Investment Views + single-name thesis (Part 2), tailored framings (Part 3), sources & caveats. |
| `model/basket-economics.xlsx` | **The supporting model.** Inputs → Calc → Sensitivity → Chart → Notes. Computes a basket value per tonne of rare-earth oxide and stress-tests it under a price-floor and an export-control-shock scenario. All figures illustrative until I input real data. |
| `viz/supply-chain-map.html` | **Interactive supply-chain map.** Self-contained (Leaflet via CDN, no build step) — overlays physical assets (mine → separation → magnet → recycling) with the policy layer (export controls, price floors, govt stakes), filterable by stage / element / policy tag. Opens directly in a browser, or served as part of the site. |
| `viz/data/assets.json` | Data behind the map — one object per asset (location, stage, status, elements, policy tags, notes). Edit this to update the map without touching code. All entries are illustrative pending `[[VERIFY]]`. |
| `scripts/build_model.py` | Builds `model/basket-economics.xlsx` from scratch with `openpyxl` (fully formula-driven; re-run to regenerate). |
| `scripts/build_site.py` | Renders the project into a static `site/` (styled memo + downloadable model) for GitHub Pages. |
| `scripts/convert_memo.py` | Renders `memo/the-seventeen.md` to PDF / DOCX via `pandoc`. |
| `.github/workflows/pages.yml` | CI: builds the site and deploys it to GitHub Pages on every push. |
| `README.md` | This file. |

## Placeholder conventions

Everything I still need to supply is flagged with one of three markers, rendered
in **bold double-brackets** so they are easy to grep:

- **`[[MY VIEW: ... ]]`** — my analysis / conclusion to write.
- **`[[VERIFY: ... — as of <date>]]`** — a fact or live price to confirm before circulating.
- **`[[INPUT: ... ]]`** — a hard number to supply (e.g. a deposit's element split).

Find them all:

```bash
grep -rno "\[\[\(MY VIEW\|VERIFY\|INPUT\)" memo/ README.md
```

## Live site

A browsable version is published to **GitHub Pages** and rebuilt on every push
(see `.github/workflows/pages.yml`): the memo as a styled HTML page with the
placeholders colour-coded, plus a download link for the model.

> **URL:** `https://rishi-guntur.github.io/the-seventeen/`
> (First deploy: if Pages isn't on yet, enable it once at **Settings → Pages →
> Source: GitHub Actions**, then re-run the workflow.)

## Build / regenerate

```bash
# 1. the model (requires: pip install openpyxl)
python scripts/build_model.py            # -> model/basket-economics.xlsx

# 2. the static site, locally (requires: pip install markdown)
python scripts/build_site.py             # -> site/index.html  (open in a browser)

# 3. the memo as PDF + DOCX (requires: pandoc, and a LaTeX engine for PDF)
python scripts/convert_memo.py           # -> build/the-seventeen.pdf, build/the-seventeen.docx
```

`build/` and exported `*.pdf` / `*.docx` are git-ignored; regenerate them from source.

## Caveat on data

Rare-earth pricing is opaque, thin, two-regime (China-domestic vs ex-China/CIF),
and goes stale quickly. Every price and deposit figure in the memo and model is a
dated placeholder until refreshed — **do not circulate with stale prices.** See the
memo's "Sources & caveats" section.

---

## Placeholder checklist

The authoritative, per-section checklist lives at the bottom of
[`memo/the-seventeen.md`](memo/the-seventeen.md). In brief, I still owe:

**Memo (`memo/the-seventeen.md`)**
- Executive Summary (write last) — `[[MY VIEW]]`
- Part 1 primer judgment calls + price/fact verifications — several `[[MY VIEW]]` / `[[VERIFY]]`
  (downstream-margin call, demand figures, two-regime prices, Apr-2025 control element list,
  the ~Nov-2026 arrangement terms, MP/DoD floor terms, offtake/stockpile list, CRMA projects,
  company-stage confirmations)
- Part 2 — Views 1–4 (thesis / what-would-have-to-be-true / what-market-misses / what-makes-me-wrong) — `[[MY VIEW]]`
- Part 2 — single-name thesis: name, one-paragraph thesis, cap-structure map, liquidity runway,
  specific catalyst, downside & recovery, what-market-gets-wrong — `[[MY VIEW]]` + `[[VERIFY]]`/`[[INPUT]]` inputs
- Part 3 — Readystate (credit/catalyst) cut and BDT & MSD (owner's-lens) cut — `[[MY VIEW]]`
- Sources — date every citation on final refresh

**Model (`model/basket-economics.xlsx`)**
- `Inputs!D` — real % by oxide for my deposit — `[[INPUT]]`
- `Inputs!E` — a second real deposit (or delete) — `[[INPUT]]`
- `Inputs!F` — live oxide prices, USD/kg, with as-of date — `[[VERIFY]]`
- `Sensitivity!B4:B6` — my own floor / shock multipliers — `[[MY VIEW]]`
