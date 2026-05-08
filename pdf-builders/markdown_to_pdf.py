#!/usr/bin/env python3
"""Render a Markdown file to a clean A4 PDF via WeasyPrint.

Usage:
    python markdown_to_pdf.py path/to/note.md
    python markdown_to_pdf.py path/to/note.md --out path/to/note.pdf

The intermediate HTML is also written next to the source so you can inspect
or tweak styling before exporting again.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import markdown


CSS = """
@page { size: A4; margin: 18mm 18mm 18mm 18mm; }
html, body {
  font-family: 'Inter', -apple-system, 'Helvetica Neue', sans-serif;
  color: #111;
  line-height: 1.45;
  font-size: 10.5pt;
}
h1 { font-size: 18pt; margin: 0 0 6pt 0; border-bottom: 2px solid #111; padding-bottom: 4pt; }
h2 { font-size: 13pt; margin: 14pt 0 4pt 0; color: #1a1a1a; border-bottom: 1px solid #ccc; padding-bottom: 3pt; }
h3 { font-size: 11pt; margin: 10pt 0 3pt 0; color: #333; }
p { margin: 4pt 0; }
ul, ol { margin: 4pt 0 4pt 18pt; padding: 0; }
li { margin: 2pt 0; }
strong { color: #000; }
em { color: #444; }
blockquote { margin: 6pt 0; padding: 8pt 12pt; background: #f6f6f6; border-left: 3px solid #555; font-style: normal; }
hr { border: 0; border-top: 1px solid #ccc; margin: 14pt 0; }
code { font-family: 'JetBrains Mono', Menlo, Consolas, monospace; font-size: 9.5pt; background: #f0f0f0; padding: 1pt 3pt; border-radius: 2pt; }
a { color: #0a4dad; text-decoration: none; }
table { border-collapse: collapse; width: 100%; margin: 6pt 0; font-size: 10pt; }
th, td { border: 1px solid #888; padding: 5pt 7pt; text-align: left; vertical-align: top; }
th { background: #f0f4fa; }
"""


def build_html(md_text: str) -> str:
    body = markdown.markdown(
        md_text,
        extensions=["extra", "sane_lists", "smarty", "tables"],
    )
    return f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body>{body}</body></html>"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("source", type=Path, help="Markdown source file")
    ap.add_argument("--out", type=Path, help="Output PDF path (default: same dir, .pdf extension)")
    args = ap.parse_args()

    src: Path = args.source
    if not src.exists():
        print(f"error: {src} not found", file=sys.stderr)
        return 1

    html = build_html(src.read_text(encoding="utf-8"))
    html_path = src.with_suffix(".html")
    pdf_path = args.out or src.with_suffix(".pdf")
    html_path.write_text(html, encoding="utf-8")

    try:
        from weasyprint import HTML  # noqa: WPS433
    except ImportError:
        print("weasyprint is required: pip install weasyprint", file=sys.stderr)
        return 2

    HTML(string=html).write_pdf(str(pdf_path))
    print(f"wrote {html_path}")
    print(f"wrote {pdf_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
