# Excel Schema Reference

## 顶层结构

```json
{
  "metadata": { "source_file", "parsed_at", "parser_version" },
  "sheets": [ Sheet ]
}
```

## Sheet（工作表）

| 字段 | 类型 | 说明 |
|------|------|------|
| `name` | string | 工作表名称 |
| `dimensions` | string | 数据范围（如 `A1:G14`） |
| `max_row` | int | 最大行号 |
| `max_column` | int | 最大列号 |
| `page_setup` | PageSetup | 页面设置 |
| `sheet_view` | SheetView | 视图设置 |
| `columns` | [Column] | 列定义 |
| `rows` | [Row] | 行定义 |
| `merged_cells` | [MergedCell] | 合并单元格 |
| `cells` | [Cell] | 单元格列表 |

## PageSetup（页面设置）

| 字段 | 类型 | 说明 |
|------|------|------|
| `paper_size` | string | 纸张名称（A4/Letter 等） |
| `orientation` | string | `portrait` / `landscape` |
| `margins.top` | float | 上边距（mm） |
| `margins.bottom` | float | 下边距（mm） |
| `margins.left` | float | 左边距（mm） |
| `margins.right` | float | 右边距（mm） |
| `margins.header` | float | 页眉距离（mm） |
| `margins.footer` | float | 页脚距离（mm） |
| `print_area` | string | 打印区域（如 `A1:F20`） |
| `fit_to_page` | bool | 缩放至页面 |
| `scale` | int | 缩放百分比 |
| `print_title_rows` | string | 打印标题行 |
| `print_title_cols` | string | 打印标题列 |

## SheetView（视图设置）

| 字段 | 类型 | 说明 |
|------|------|------|
| `zoom_scale` | int | 缩放比例（如 85 表示 85%） |
| `zoom_scale_normal` | int | 普通视图缩放 |
| `zoom_scale_page_layout` | int | 页面布局视图缩放 |
| `show_grid_lines` | bool | 显示网格线 |
| `freeze_panes` | string | 冻结窗格（如 `A2`） |
| `show_formulas` | bool | 显示公式 |
| `show_headers` | bool | 显示行列标题 |
| `show_zeros` | bool | 显示零值 |
| `right_to_left` | bool | 从右到左 |
| `tab_selected` | bool | 工作表被选中 |

## Column（列）

| 字段 | 类型 | 说明 |
|------|------|------|
| `index` | int | 列索引（0-based） |
| `letter` | string | 列字母（A/B/C...） |
| `width` | float | 列宽（字符单位） |
| `width_px` | float | 列宽（像素，96 DPI） |
| `hidden` | bool | 隐藏列 |

## Row（行）

| 字段 | 类型 | 说明 |
|------|------|------|
| `index` | int | 行索引（0-based） |
| `number` | int | 行号（1-based） |
| `height` | float | 行高（磅） |
| `height_px` | float | 行高（像素，96 DPI） |
| `hidden` | bool | 隐藏行 |
| `custom_height` | bool | 自定义行高 |

## MergedCell（合并单元格）

| 字段 | 类型 | 说明 |
|------|------|------|
| `range` | string | 合并范围（如 `A1:C3`） |
| `start_row` | int | 起始行（0-based） |
| `start_col` | int | 起始列（0-based） |
| `end_row` | int | 结束行（0-based） |
| `end_col` | int | 结束列（0-based） |
| `rowspan` | int | 跨行数 |
| `colspan` | int | 跨列数 |

## Cell（单元格）

| 字段 | 类型 | 说明 |
|------|------|------|
| `address` | string | 单元格地址（如 `A1`） |
| `row` | int | 行索引（0-based） |
| `col` | int | 列索引（0-based） |
| `data_type` | string | `empty` / `string` / `number` / `datetime` / `date` / `time` / `boolean` / `formula` |
| `value` | any | 单元格值（公式单元格为计算结果，可能为 null） |
| `formula` | string | 公式字符串（仅 data_type 为 formula 时存在） |
| `number_format` | string | 数字格式（如 `#,##0.00`） |
| `style` | CellStyle | 单元格样式 |

**公式单元格说明：**
- 当 `data_type` 为 `formula` 时，会同时输出 `formula`（公式字符串）和 `value`（计算结果）
- `value` 字段依赖 Excel 已缓存的计算结果（文件需在 Excel 中打开保存过）
- 如果 `value` 为 null，说明公式未被计算过，需要：
  - 在 Excel 中打开文件并保存，让 Excel 计算公式
  - 或使用 formulas 库自行计算（需 `pip install formulas`）
  - 或由 Agent 根据公式和依赖单元格自行计算

## CellStyle（单元格样式）

| 字段 | 类型 | 说明 |
|------|------|------|
| `font` | Font | 字体 |
| `alignment` | Alignment | 对齐 |
| `fill` | Fill | 背景填充 |
| `border` | Border | 边框 |

## Font（字体）

| 字段 | 类型 | 说明 |
|------|------|------|
| `name` | string | 字体名称 |
| `size` | float | 字号（磅） |
| `bold` | bool | 粗体 |
| `italic` | bool | 斜体 |
| `underline` | string | 下划线类型 |
| `strikethrough` | bool | 删除线 |
| `color` | Color | 颜色 |

## Alignment（对齐）

| 字段 | 类型 | 说明 |
|------|------|------|
| `horizontal` | string | `left` / `center` / `right` / `justify` / `distributed` |
| `vertical` | string | `top` / `center` / `bottom` / `justify` / `distributed` |
| `wrap_text` | bool | 自动换行 |
| `text_rotation` | int | 文本旋转角度（0-90 逆时针，91-180 顺时针减90，255 竖排） |
| `indent` | float | 缩进级别 |

## Fill（背景填充）

| 字段 | 类型 | 说明 |
|------|------|------|
| `pattern_type` | string | 填充模式（`solid` 等） |
| `fg_color` | Color | 前景色 |
| `bg_color` | Color | 背景色 |

## Border（边框）

| 字段 | 类型 | 说明 |
|------|------|------|
| `top` / `bottom` / `left` / `right` / `diagonal` | BorderSide | 各边边框 |
| `diagonal_down` | bool | 对角线方向（右下） |
| `diagonal_up` | bool | 对角线方向（右上） |

## BorderSide（边框边）

| 字段 | 类型 | 说明 |
|------|------|------|
| `style` | string | 线型：`thin` / `medium` / `thick` / `dashed` / `dotted` / `double` / `hair` 等 13 种 |
| `color` | Color | 颜色 |

## Color（颜色）

| 字段 | 类型 | 说明 |
|------|------|------|
| `rgb` | string | RGB 十六进制（如 `FF0000`） |
| `theme` | int | 主题色索引 |
| `tint` | float | 色调调整（-1.0 到 1.0） |
| `indexed` | int | 索引色 |
