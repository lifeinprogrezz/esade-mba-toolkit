# esade-mba-toolkit

Tools and writing produced over an ESADE MBA. Curated portfolio cut — the heavy course-specific notes live elsewhere.

## What's here

- **`slides-generator/`** — Python deck builder on top of `python-pptx` and `reportlab`. Includes `esade_design.py` (a small design-system module: tokens, helpers, layouts) and `velos_deck.py` (a 12-slide reference deck on a fictional logistics consortium pivot, plus a 3-slide presentation cut and a speaker-notes PDF). Build by running `python velos_deck.py`.
- **`pdf-builders/`** — `markdown_to_pdf.py`: tiny Markdown → A4 PDF converter built on `python-markdown` + WeasyPrint. Used to print study notes that read better on paper than on screen.
- **`frameworks/`** — Generic textbook frameworks distilled into one-page sheets. `managerial-accounting-formulas.md` is the cost-accounting CVP/OAR/BE formula set.
- **`writing/`** — Two written pieces that stand on their own:
  - `investment-thesis/` — *Vertical-AI middleware is structurally squeezed; durable edge is time-arbitrage in non-AI-defensible niches.* Submitted assignment in Entrepreneurship Finance.
  - `marketing-summaries/` — Three weekly session summaries (Pearson outcomes, SKF value selling, willingness-to-pay measurement). Written for an MBA marketing course; sanitized for public posting.

## Setup

```bash
# Slides generator
pip install python-pptx reportlab
cd slides-generator && python velos_deck.py

# PDF builder
pip install markdown weasyprint
python pdf-builders/markdown_to_pdf.py path/to/note.md
```

## Note on scope

This is the public bundle. Course-specific exam material (mock exams, professor-attributed notes, study coaches) is intentionally excluded.
