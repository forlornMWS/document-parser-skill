---
name: document-parser-skill
description: Parse Word (.docx) and Excel (.xlsx) templates into structured JSON capturing every style detail (fonts, alignment, borders, merges, page setup) so an Agent can reproduce the layout in PDF. Use when the user needs to extract format from a document template, convert docx/xlsx to JSON, or generate a style reference for PDF creation.
---

# document-parser-skill

Parse a Word or Excel template into a structured JSON snapshot of its format — every font, margin, border, merge, and page-setup detail — so the layout can be reproduced as PDF.

## When to use

- User provides a `.docx` or `.xlsx` template and wants its format extracted to JSON.
- User wants to reproduce a document's layout in PDF and needs a style reference.
- User asks to "parse the template", "extract the format", or "get the styles" from a Word/Excel file.

## Leading words

- **Snapshot** — the JSON output; the single source of truth for the template's format.
- **Schema** — the reference doc (`references/word-schema.md` or `references/excel-schema.md`) defining every field in the snapshot.

## Instructions

1. Identify the input file's type by extension: `.docx` → Word, `.xlsx` → Excel.
2. Run the matching parser from `scripts/`:
   - Word: `python scripts/parse_word.py <input.docx> [output.json] [options]`
   - Excel: `python scripts/parse_excel.py <input.xlsx> [output.json] [options]`
3. Choose output target:
   - Omit `output` → writes `<input_stem>_format.json` next to the input.
   - `-` → writes to stdout (pipe to other tools).
   - Explicit path → writes there (`.json` appended if missing).
4. Select flags by goal:
   - `--pretty` — indented JSON for human review.
   - `--skip-empty` — drop empty paragraphs (Word) or empty unstyled cells (Excel) to shrink the snapshot.
   - `--extract-images` (Word only) — pull embedded images into `<input_stem>_images/`.
   - `--sheet "A,B"` / `--exclude "C"` (Excel only) — restrict or exclude sheets.
   - `--list-sheets` (Excel only) — list sheet names and exit.
5. Read the matching **Schema** to interpret any field in the snapshot.
6. Hand the snapshot to the PDF-generation step, or report a summary of the extracted format to the user.

## Completion criterion

The snapshot JSON exists at the chosen output path, every sheet (Excel) or body element (Word) is accounted for, and the user has either the file path or a summary of the parsed format.

## Reference

- `references/word-schema.md` — every field in a Word snapshot.
- `references/excel-schema.md` — every field in an Excel snapshot.

## Notes

- Excel formulas: `value` is the cached result (null if Excel never recalculated the file). The `formula` string is always present for formula cells.
- Excel loads the workbook twice (formulas + cached values); large files take longer but produce complete snapshots.
- Word image extraction writes binary files to disk; the snapshot records each image's filename, content type, and size.
