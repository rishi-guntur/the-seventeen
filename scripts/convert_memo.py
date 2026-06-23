#!/usr/bin/env python3
"""
convert_memo.py — render memo/the-seventeen.md to PDF and/or DOCX.

Uses pandoc (the standard, clean route). This script does not invent styling;
it produces a professional, minimal document and preserves the bold
double-bracket placeholders so they remain easy to find in the output.

Prereqs:
    - pandoc            (https://pandoc.org)   — required
    - a LaTeX engine    (e.g. tectonic, xelatex, or wkhtmltopdf) — for PDF only

Usage:
    python scripts/convert_memo.py                 # both docx and pdf
    python scripts/convert_memo.py --docx          # docx only
    python scripts/convert_memo.py --pdf           # pdf only
    python scripts/convert_memo.py -i memo/the-seventeen.md -o build/

If pandoc (or a PDF engine) is missing, the script prints exactly what to
install and exits non-zero rather than producing a half-baked file.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys

DEFAULT_INPUT = os.path.join("memo", "the-seventeen.md")
DEFAULT_OUTDIR = "build"


def _require(tool: str, hint: str) -> str:
    path = shutil.which(tool)
    if not path:
        sys.stderr.write(
            f"[convert_memo] '{tool}' not found on PATH.\n"
            f"               Install it: {hint}\n")
        sys.exit(2)
    return path


def _pick_pdf_engine() -> str | None:
    for engine in ("tectonic", "xelatex", "lualatex", "pdflatex", "wkhtmltopdf"):
        if shutil.which(engine):
            return engine
    return None


def convert(input_md: str, outdir: str, want_docx: bool, want_pdf: bool) -> None:
    _require("pandoc", "https://pandoc.org/installing.html "
                       "(macOS: brew install pandoc).")
    if not os.path.exists(input_md):
        sys.stderr.write(f"[convert_memo] input not found: {input_md}\n")
        sys.exit(1)
    os.makedirs(outdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(input_md))[0]

    common = [
        "pandoc", input_md,
        "--from", "gfm",
        "--standalone",
        "--toc", "--toc-depth=2",
        "--metadata", "title=The Seventeen — Critical Minerals Primer & Investment Thesis",
    ]

    if want_docx:
        out = os.path.join(outdir, base + ".docx")
        subprocess.run(common + ["-o", out], check=True)
        print(f"Wrote {out}")

    if want_pdf:
        engine = _pick_pdf_engine()
        if not engine:
            sys.stderr.write(
                "[convert_memo] No PDF engine found. Install one of: tectonic, "
                "xelatex (TeX Live/MacTeX), or wkhtmltopdf. "
                "Skipping PDF.\n")
            sys.exit(3)
        out = os.path.join(outdir, base + ".pdf")
        cmd = common + ["-o", out]
        if engine == "wkhtmltopdf":
            cmd += ["--pdf-engine=wkhtmltopdf"]
        else:
            cmd += [f"--pdf-engine={engine}",
                    "-V", "geometry:margin=1in",
                    "-V", "fontsize=11pt",
                    "-V", "linkcolor=blue",
                    "-V", "mainfont=Helvetica" if engine in ("xelatex", "lualatex") else "",
                    ]
            cmd = [c for c in cmd if c]  # drop empty mainfont for non-xelatex
        subprocess.run(cmd, check=True)
        print(f"Wrote {out}")


def main() -> None:
    p = argparse.ArgumentParser(description="Render the-seventeen.md to PDF/DOCX via pandoc.")
    p.add_argument("-i", "--input", default=DEFAULT_INPUT)
    p.add_argument("-o", "--outdir", default=DEFAULT_OUTDIR)
    p.add_argument("--docx", action="store_true", help="produce DOCX only")
    p.add_argument("--pdf", action="store_true", help="produce PDF only")
    args = p.parse_args()

    # default: both
    want_docx = args.docx or not (args.docx or args.pdf)
    want_pdf = args.pdf or not (args.docx or args.pdf)
    convert(args.input, args.outdir, want_docx, want_pdf)


if __name__ == "__main__":
    main()
