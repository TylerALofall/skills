# Heading extraction to Akoma Ntoso

Use `heading_to_akn.py` to turn a DOCX into a heading-driven Akoma Ntoso outline without any subprocesses.

## What it does
- Reads `word/document.xml` directly from the DOCX zip.
- Detects headings via styles (`Heading1`–`Heading4a`, `H1`–`H4a`).
- Assigns text “possession” to the nearest ancestor heading: content is owned until the next heading of the same or higher level.
- Skips boilerplate like `Case No. XX-XXXX` banners and center page numbers before attaching text.
- Emits nested `<section>` nodes with `<heading>` and `<p>` children inside an `<akomaNtoso><judgment><body>` container.

## Usage
```bash
python docx/heading_to_akn.py source.docx output.xml --toc
```
- `--toc` prints the heading outline (levels H1–H4a) before writing XML.

## Notes and assumptions
- Keep templates read-only; the tool never edits the DOCX in place.
- No shelling out: all work is pure Python (`zipfile` + `defusedxml` + standard `xml.etree` for emission).
- Boilerplate trimming: paragraphs matching `Case No. ...` or simple page numbers (e.g., `3` or `3 of 12`) are ignored so headings own substantive text only.
- Output uses lowercase, dash-separated `id` attributes derived from heading text; empty headings fall back to `section`.
