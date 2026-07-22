# document-parser-skill

将 Word (.docx) 和 Excel (.xlsx) 模板解析为结构化 JSON，捕获所有样式细节——字体、对齐、边框、合并、页面设置——以便 Agent 在 PDF 中复现布局。

[English](README.md)

## 安装

通过 skills CLI 安装本 skill：

```bash
npx skills add forlornMWS/document-parser-skill
```

或者克隆仓库后复制到 skills 目录：

```bash
git clone https://github.com/forlornMWS/document-parser-skill.git
cp -r document-parser-skill ~/.codebuddy/skills/
```

### 依赖

安装 skill 后，安装 Python 依赖：

```bash
pip install python-docx lxml openpyxl
```

## 用法

### Word

```bash
python scripts/parse_word.py <input.docx> [output.json] [options]
```

| 选项 | 说明 |
|------|------|
| `output` | 输出路径，默认 `<input>_format.json`；`-` 输出到 stdout |
| `--pretty` | 美化输出（带缩进） |
| `--skip-empty` | 跳过空段落 |
| `--extract-images` | 提取图片到 `<input>_images/` 目录 |

示例：

```bash
# 默认输出到 template_format.json
python scripts/parse_word.py template.docx --pretty

# 提取图片
python scripts/parse_word.py template.docx --extract-images

# 输出到 stdout
python scripts/parse_word.py template.docx - --pretty
```

### Excel

```bash
python scripts/parse_excel.py <input.xlsx> [output.json] [options]
```

| 选项 | 说明 |
|------|------|
| `output` | 输出路径，默认 `<input>_format.json`；`-` 输出到 stdout |
| `--pretty` | 美化输出（带缩进） |
| `--skip-empty` | 跳过空且无样式的单元格 |
| `--sheet "A,B"` | 只解析指定 sheet |
| `--exclude "C"` | 排除指定 sheet |
| `--list-sheets` | 列出所有 sheet 后退出 |

示例：

```bash
# 只解析指定 sheet
python scripts/parse_excel.py template.xlsx --sheet "Sheet1,Sheet3" --pretty

# 列出所有 sheet
python scripts/parse_excel.py template.xlsx --list-sheets
```

## 输出结构

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

字段详情见 `references/word-schema.md` 和 `references/excel-schema.md`。

## 注意事项

- Excel 公式单元格同时输出 `formula`（公式字符串）和 `value`（缓存计算结果）。若 `value` 为 null，说明文件未在 Excel 中保存过，需打开重存或用 `formulas` 库计算。
- Excel 会加载工作簿两次（公式 + 缓存值），大文件耗时较长。
- Word `--extract-images` 会将图片写入磁盘，JSON 中记录文件名、类型和大小。

## 目录结构

```
document-parser-skill/
├── SKILL.md                      # Skill 入口（model-invoked）
├── README.md
├── scripts/
│   ├── parse_word.py             # Word 模板解析器
│   └── parse_excel.py            # Excel 模板解析器
└── references/
    ├── word-schema.md            # Word 输出 JSON 字段说明
    └── excel-schema.md           # Excel 输出 JSON 字段说明
```
