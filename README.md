# document-parser-skill

Parse Word (.docx) and Excel (.xlsx) templates into structured JSON capturing every style detail — fonts, alignment, borders, merges, page setup — so an Agent can reproduce the layout in PDF.

[中文文档](README-ZH.md)

## Installation

Install this skill via the skills CLI:

```bash
npx skills add forlornMWS/document-parser-skill
```

Alternatively, clone the repository and copy it into your skills directory:

```bash
git clone https://github.com/forlornMWS/document-parser-skill.git
cp -r document-parser-skill ~/.codebuddy/skills/
```

### Dependencies

After installing the skill, install the Python dependencies:

```bash
pip install python-docx lxml openpyxl
```

## Usage

### Word

```bash
python scripts/parse_word.py <input.docx> [output.json] [options]
```

| Option | Description |
|--------|-------------|
| `output` | Output path; default `<input>_format.json`; `-` for stdout |
| `--pretty` | Pretty-print with indentation |
| `--skip-empty` | Skip empty paragraphs |
| `--extract-images` | Extract images to `<input>_images/` directory |

Examples:

```bash
# Default output to template_format.json
python scripts/parse_word.py template.docx --pretty

# Extract images
python scripts/parse_word.py template.docx --extract-images

# Output to stdout
python scripts/parse_word.py template.docx - --pretty
```

### Excel

```bash
python scripts/parse_excel.py <input.xlsx> [output.json] [options]
```

| Option | Description |
|--------|-------------|
| `output` | Output path; default `<input>_format.json`; `-` for stdout |
| `--pretty` | Pretty-print with indentation |
| `--skip-empty` | Skip empty unstyled cells |
| `--sheet "A,B"` | Parse only specified sheets |
| `--exclude "C"` | Exclude specified sheets |
| `--list-sheets` | List all sheets and exit |

Examples:

```bash
# Parse only specified sheets
python scripts/parse_excel.py template.xlsx --sheet "Sheet1,Sheet3" --pretty

# List all sheets
python scripts/parse_excel.py template.xlsx --list-sheets
```

## Output Structure

### Word

```json
{
  "metadata": { "source_file", "parsed_at", "parser_version", "statistics" },
  "sections": [ { "page_width_cm", "page_height_cm", "paper_size", "orientation", "margins" } ],
  "body": [ { "type": "paragraph", "text", "style_name", "format", "runs" } ],
  "header": { "linked_to_previous", "paragraphs" },
  "footer": { "linked_to_previous", "paragraphs" }
}
```

### Excel

```json
{
  "metadata": { "source_file", "parsed_at", "parser_version" },
  "sheets": [
    {
      "name", "dimensions", "max_row", "max_column",
      "page_setup": { "paper_size", "orientation", "margins" },
      "columns": [ { "index", "letter", "width", "width_px" } ],
      "rows": [ { "index", "number", "height", "height_px" } ],
      "merged_cells": [ { "range", "start_row", "start_col", "end_row", "end_col" } ],
      "cells": [ { "address", "row", "col", "data_type", "value", "formula", "style" } ]
    }
  ]
}
```

See `references/word-schema.md` and `references/excel-schema.md` for field details.

## Notes

- Excel formula cells output both `formula` (formula string) and `value` (cached result). If `value` is null, the file was never saved in Excel — reopen and save, or use the `formulas` library to compute.
- Excel loads the workbook twice (formulas + cached values); large files take longer.
- Word `--extract-images` writes image files to disk; the JSON records filename, content type, and size.

## Project Structure

```
document-parser-skill/
├── SKILL.md                      # Skill entry (model-invoked)
├── README.md
├── scripts/
│   ├── parse_word.py             # Word template parser
│   └── parse_excel.py            # Excel template parser
└── references/
    ├── word-schema.md            # Word output JSON field reference
    └── excel-schema.md           # Excel output JSON field reference
```
