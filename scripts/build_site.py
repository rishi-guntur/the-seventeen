#!/usr/bin/env python3
"""
build_site.py — render the project into a small static site for GitHub Pages.

Produces a `site/` directory containing:
    - index.html                  landing page (project description + file map + links)
    - the-seventeen.html          the memo, rendered as a clean, styled HTML page
    - basket-economics.xlsx       a copy of the model, so it is downloadable from the site

The memo's `[[MY VIEW: ...]]`, `[[VERIFY: ...]]`, and `[[INPUT: ...]]` placeholders
are wrapped in highlighted spans so they are visually obvious on the page.

No external services and no network access required. Only dependency is the
`markdown` package (pip install markdown).

Usage:
    python scripts/build_site.py            # -> site/
    python scripts/build_site.py -o public  # custom output dir
"""

from __future__ import annotations

import argparse
import datetime as _dt
import html
import re
import shutil
import sys
from pathlib import Path

try:
    import markdown as _md
except ImportError:
    sys.exit(
        "Missing dependency: markdown.\n"
        "Install it with:  pip install markdown"
    )

ROOT = Path(__file__).resolve().parent.parent
MEMO = ROOT / "memo" / "the-seventeen.md"
MODEL = ROOT / "model" / "basket-economics.xlsx"
README = ROOT / "README.md"
VIZ_DIR = ROOT / "viz"

# Match a full placeholder, including the surrounding ** bold markers if present,
# across multiple lines. Captures the placeholder type for colour-coding.
PLACEHOLDER_RE = re.compile(
    r"\[\[\s*(MY VIEW|VERIFY|INPUT)\s*:.*?\]\]",
    re.DOTALL,
)

CSS = """
:root {
  --ink: #1a1a1a; --muted: #5b6470; --rule: #e4e7eb; --bg: #ffffff;
  --accent: #14532d; --link: #1d4ed8;
  --myview: #b91c1c; --myview-bg: #fef2f2;
  --verify: #b45309; --verify-bg: #fffbeb;
  --input:  #6d28d9; --input-bg:  #f5f3ff;
}
* { box-sizing: border-box; }
html { -webkit-text-size-adjust: 100%; }
body {
  margin: 0; color: var(--ink); background: var(--bg);
  font: 17px/1.65 "Charter","Georgia","Times New Roman",serif;
}
.wrap { max-width: 820px; margin: 0 auto; padding: 56px 28px 120px; }
.topbar {
  font: 13px/1.5 ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,sans-serif;
  border-bottom: 1px solid var(--rule); padding: 14px 28px; color: var(--muted);
  display: flex; gap: 18px; flex-wrap: wrap; align-items: center;
}
.topbar a { color: var(--link); text-decoration: none; }
.topbar a:hover { text-decoration: underline; }
.topbar .sep { color: var(--rule); }
h1, h2, h3, h4 {
  font-family: ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,sans-serif;
  line-height: 1.25; color: var(--ink); margin: 2em 0 .55em;
}
h1 { font-size: 30px; margin-top: .2em; letter-spacing: -.01em; }
h2 { font-size: 22px; padding-top: .4em; border-top: 1px solid var(--rule); margin-top: 2.2em; }
h3 { font-size: 18px; color: #374151; }
h4 { font-size: 15px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); }
p, li { color: var(--ink); }
em { color: var(--muted); }
a { color: var(--link); }
hr { border: none; border-top: 1px solid var(--rule); margin: 2.4em 0; }
blockquote {
  margin: 1.4em 0; padding: .2em 1.1em; border-left: 3px solid var(--rule);
  color: var(--muted); background: #fafbfc;
}
blockquote code { background: #eef1f4; }
code {
  font: 13.5px/1.5 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;
  background: #f3f4f6; padding: .12em .35em; border-radius: 4px;
}
pre { background: #f6f8fa; border: 1px solid var(--rule); border-radius: 8px;
      padding: 14px 16px; overflow-x: auto; }
pre code { background: none; padding: 0; }
table { border-collapse: collapse; width: 100%; margin: 1.4em 0;
        font: 14px/1.5 ui-sans-serif,system-ui,sans-serif; }
th, td { border: 1px solid var(--rule); padding: 7px 10px; text-align: left;
         vertical-align: top; }
th { background: #f8fafc; font-weight: 600; }
ul, ol { padding-left: 1.4em; }
input[type=checkbox] { margin-right: .5em; }

/* placeholder highlighting */
.ph { border-radius: 5px; padding: .08em .35em; font-weight: 600;
      font-family: ui-sans-serif,system-ui,sans-serif; font-size: .94em;
      border: 1px solid; }
.ph-myview { color: var(--myview); background: var(--myview-bg); border-color: #fecaca; }
.ph-verify { color: var(--verify); background: var(--verify-bg); border-color: #fde68a; }
.ph-input  { color: var(--input);  background: var(--input-bg);  border-color: #ddd6fe; }

/* landing page */
.lede { font-size: 19px; color: var(--muted); margin: .4em 0 1.6em; }
.card { display: block; border: 1px solid var(--rule); border-radius: 10px;
        padding: 18px 20px; margin: 12px 0; text-decoration: none; color: inherit;
        transition: border-color .15s, box-shadow .15s; }
.card:hover { border-color: #c7ccd3; box-shadow: 0 1px 4px rgba(0,0,0,.05); }
.card h3 { margin: 0 0 4px; color: var(--link);
           font-family: ui-sans-serif,system-ui,sans-serif; }
.card p { margin: 0; color: var(--muted); font-size: 15px; }
.legend { display: flex; gap: 10px; flex-wrap: wrap; margin: 1.2em 0; font-size: 14px;
          font-family: ui-sans-serif,system-ui,sans-serif; }
.legend .ph { font-size: 13px; }
.refreshed { color: var(--muted); font-size: 14px;
             font-family: ui-sans-serif,system-ui,sans-serif; }
"""

PAGE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>{css}</style>
</head>
<body>
<nav class="topbar">
  <a href="index.html">The 17 — home</a><span class="sep">|</span>
  <a href="the-seventeen.html">Memo</a><span class="sep">|</span>
  <a href="basket-economics.xlsx">Model (.xlsx)</a><span class="sep">|</span>
  <span>Last refreshed: {refreshed}</span>
</nav>
<div class="wrap">
{body}
</div>
</body>
</html>
"""


def _wrap_placeholders(rendered_html: str) -> str:
    """Wrap each [[TYPE: ...]] occurrence in a colour-coded span."""

    def repl(m: re.Match) -> str:
        kind = m.group(1).lower().replace(" ", "")  # myview / verify / input
        # The inner text was already HTML-escaped by the markdown renderer.
        return f'<span class="ph ph-{kind}">{m.group(0)}</span>'

    return PLACEHOLDER_RE.sub(repl, rendered_html)


def _render_markdown(md_path: Path) -> str:
    text = md_path.read_text(encoding="utf-8")
    body = _md.markdown(
        text,
        extensions=["extra", "tables", "sane_lists", "toc"],
    )
    body = _wrap_placeholders(body)
    # The memo links to ../viz/... (correct from memo/ in the repo); the site
    # root flattens memo/ and viz/ as siblings, so rewrite the link to match.
    body = body.replace('href="../viz/', 'href="viz/')
    return body


def _count_placeholders(md_path: Path) -> dict[str, int]:
    text = md_path.read_text(encoding="utf-8")
    counts = {"MY VIEW": 0, "VERIFY": 0, "INPUT": 0}
    for m in PLACEHOLDER_RE.finditer(text):
        counts[m.group(1)] += 1
    return counts


def build(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    refreshed = _dt.date.today().isoformat()

    # --- memo page ---
    memo_body = _render_markdown(MEMO)
    (out_dir / "the-seventeen.html").write_text(
        PAGE.format(
            title="The 17 — Critical Minerals Primer & Investment Thesis",
            css=CSS, refreshed=refreshed, body=memo_body,
        ),
        encoding="utf-8",
    )

    # --- model copy (downloadable) ---
    if MODEL.exists():
        shutil.copy2(MODEL, out_dir / "basket-economics.xlsx")
    else:
        print(f"warning: {MODEL} not found — run scripts/build_model.py first")

    # --- supply-chain map (standalone HTML + its data file) ---
    if VIZ_DIR.exists():
        out_viz = out_dir / "viz"
        if out_viz.exists():
            shutil.rmtree(out_viz)
        shutil.copytree(VIZ_DIR, out_viz)
    else:
        print(f"warning: {VIZ_DIR} not found — skipping supply-chain map")

    # --- landing page ---
    counts = _count_placeholders(MEMO)
    total = sum(counts.values())
    index_body = f"""
<h1>The 17 — Critical Minerals Primer &amp; Investment Thesis</h1>
<p class="lede">A self-directed research &amp; investment project on the rare-earth /
critical-minerals supply chain — written as an investment piece, not a survey.</p>
<p class="refreshed">Last refreshed: {refreshed} &nbsp;·&nbsp; Working draft.</p>

<a class="card" href="the-seventeen.html">
  <h3>Read the memo &rarr;</h3>
  <p>Primer (Part 1), investment views &amp; single-name thesis (Part 2),
     tailored framings (Part 3), sources &amp; caveats.</p>
</a>
<a class="card" href="basket-economics.xlsx">
  <h3>Download the model (.xlsx) &darr;</h3>
  <p>Basket-economics workbook: Inputs &rarr; Calc &rarr; Sensitivity &rarr; Chart &rarr; Notes.
     Computes basket value per tonne and stress-tests it. All figures illustrative until real data is entered.</p>
</a>
<a class="card" href="viz/supply-chain-map.html">
  <h3>Supply-chain map &rarr;</h3>
  <p>Interactive map overlaying physical assets (mine &rarr; separation &rarr; magnet) with the
     policy layer (export controls, price floors, govt stakes). Filterable by stage, element,
     and policy tag. All locations and statuses are illustrative pending verification.</p>
</a>

<h2>Placeholders still to fill in</h2>
<p>The memo is a scaffold: primer facts are written and cited; my own judgment,
live prices, and deposit inputs are left as marked placeholders. There are
<strong>{total}</strong> in total.</p>
<div class="legend">
  <span class="ph ph-myview">[[MY VIEW]] &times; {counts['MY VIEW']}</span>
  <span class="ph ph-verify">[[VERIFY]] &times; {counts['VERIFY']}</span>
  <span class="ph ph-input">[[INPUT]] &times; {counts['INPUT']}</span>
</div>
<p class="refreshed">A full per-section checklist lives at the bottom of the memo.</p>
"""
    (out_dir / "index.html").write_text(
        PAGE.format(
            title="The 17 — Critical Minerals Primer & Investment Thesis",
            css=CSS, refreshed=refreshed, body=index_body,
        ),
        encoding="utf-8",
    )

    # GitHub Pages: don't run the output through Jekyll.
    (out_dir / ".nojekyll").write_text("", encoding="utf-8")

    print(f"Wrote site to {out_dir}/  ({total} placeholders: {counts})")


def main() -> None:
    ap = argparse.ArgumentParser(description="Build the static site for GitHub Pages.")
    ap.add_argument("-o", "--out", default=str(ROOT / "site"), help="output directory")
    args = ap.parse_args()
    build(Path(args.out))


if __name__ == "__main__":
    main()
