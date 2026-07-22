# Word Schema Reference

## 顶层结构

```json
{
  "metadata": { "source_file", "parsed_at", "parser_version", "statistics" },
  "sections": [ Section ],
  "body": [ Paragraph | Table ],
  "header": HeaderFooter,
  "footer": HeaderFooter,
  "first_page_header": HeaderFooter,
  "first_page_footer": HeaderFooter
}
```

## Section（页面设置）

| 字段 | 类型 | 说明 |
|------|------|------|
| `page_width_cm` | float | 页面宽度（厘米） |
| `page_height_cm` | float | 页面高度（厘米） |
| `paper_size` | string | 纸张名称（A4/Letter 等） |
| `orientation` | string | `portrait` / `landscape` |
| `margins.top_cm` | float | 上边距 |
| `margins.bottom_cm` | float | 下边距 |
| `margins.left_cm` | float | 左边距 |
| `margins.right_cm` | float | 右边距 |
| `margins.header_cm` | float | 页眉距离 |
| `margins.footer_cm` | float | 页脚距离 |
| `different_first_page` | bool | 首页不同页眉页脚 |

## Paragraph（段落）

| 字段 | 类型 | 说明 |
|------|------|------|
| `type` | string | `"paragraph"` |
| `text` | string | 段落纯文本 |
| `style_name` | string | 样式名称（Normal/Heading 1 等） |
| `format` | ParagraphFormat | 段落格式 |
| `runs` | [Run] | 文本运行列表 |

## ParagraphFormat（段落格式）

| 字段 | 类型 | 说明 |
|------|------|------|
| `alignment` | string | `left` / `center` / `right` / `justify` / `distribute` |
| `left_indent_pt` | float | 左缩进（磅） |
| `right_indent_pt` | float | 右缩进（磅） |
| `first_line_indent_pt` | float | 首行缩进（磅） |
| `line_spacing_rule` | string | `single` / `1.5` / `double` / `at_least` / `exactly` / `multiple` |
| `line_spacing` | float | 倍数行距（single/double/multiple 时） |
| `line_spacing_pt` | float | 固定/最小行距（磅）（at_least/exactly 时） |
| `space_before_pt` | float | 段前间距（磅） |
| `space_after_pt` | float | 段后间距（磅） |
| `page_break_before` | bool | 段前分页 |
| `keep_together` | bool | 保持行 |
| `keep_with_next` | bool | 保持段 |
| `outline_level` | int | 大纲级别 |
| `borders` | Borders | 段落边框 |

## Borders（边框）

| 字段 | 类型 | 说明 |
|------|------|------|
| `top` / `bottom` / `left` / `right` / `between` / `bar` | Border | 各边边框 |

**Border 对象：**

| 字段 | 类型 | 说明 |
|------|------|------|
| `style` | string | 线型：`single` / `double` / `dashed` / `dotted` / `thick` 等 |
| `size` | float | 粗细（磅） |
| `space_pt` | float | 间距（磅） |
| `color` | string | 颜色（十六进制 RGB） |

## Run（文本运行）

| 字段 | 类型 | 说明 |
|------|------|------|
| `text` | string | 文本内容 |
| `font` | Font | 字体样式 |

## Font（字体）

| 字段 | 类型 | 说明 |
|------|------|------|
| `name` | string | 西文字体 |
| `east_asia` | string | 东亚字体 |
| `size_pt` | float | 字号（磅） |
| `bold` | bool | 粗体 |
| `italic` | bool | 斜体 |
| `underline` | bool/string | 下划线 |
| `strike` | bool | 删除线 |
| `all_caps` | bool | 全大写 |
| `small_caps` | bool | 小型大写 |
| `color` | Color | 颜色 |
| `highlight` | string | 高亮色 |
| `superscript` | bool | 上标 |
| `subscript` | bool | 下标 |

## Color（颜色）

| 字段 | 类型 | 说明 |
|------|------|------|
| `rgb` | string | RGB 十六进制（如 `FF0000`） |
| `theme` | string | 主题色名称 |
| `brightness` | float | 亮度调整 |

## Table（表格）

| 字段 | 类型 | 说明 |
|------|------|------|
| `type` | string | `"table"` |
| `alignment` | string | 表格对齐 |
| `style_name` | string | 表格样式 |
| `borders` | Borders | 表格边框 |
| `width` | int | 宽度 |
| `width_type` | string | 宽度类型 |
| `rows` | [Row] | 行列表 |

## Row（行）

| 字段 | 类型 | 说明 |
|------|------|------|
| `height_pt` | float | 行高（磅） |
| `height_rule` | string | 行高规则 |
| `cells` | [Cell] | 单元格列表 |

## Cell（单元格）

| 字段 | 类型 | 说明 |
|------|------|------|
| `text` | string | 单元格文本 |
| `vertical_alignment` | string | `top` / `center` / `bottom` / `justify` |
| `background_color` | string | 背景色 |
| `borders` | Borders | 单元格边框 |
| `colspan` | int | 横向合并 |
| `vmerge_restart` | bool | 纵向合并起始 |
| `vmerge_continue` | bool | 纵向合并延续 |
| `paragraphs` | [Paragraph] | 段落列表 |

## HeaderFooter（页眉页脚）

| 字段 | 类型 | 说明 |
|------|------|------|
| `linked_to_previous` | bool | 链接到上一节 |
| `paragraphs` | [Paragraph] | 段落列表 |
