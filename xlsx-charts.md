# xlsxwriter 图表可靠排版

## 原则
xlsxwriter 图表是浮动对象，多图放同一 Sheet 必然互相覆盖。
**唯一稳的方案：一图一 Sheet，数据表在上、图表在下，垂直堆叠。**

---

## 模板

```python
import xlsxwriter

wb = xlsxwriter.Workbook("charts.xlsx")

fmt_hdr  = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
                           'border': 1, 'align': 'center'})
fmt_cell = wb.add_format({'border': 1, 'align': 'center'})
fmt_pct  = wb.add_format({'border': 1, 'align': 'center', 'num_format': '0.0%'})
fmt_num  = wb.add_format({'border': 1, 'align': 'center', 'num_format': '0.00'})

def make_chart_sheet(wb, name, title, headers, data_rows, chart_type, series_list,
                     y_min=None, y_max=None, chart_width=600):
    """一图一 Sheet。图表高度 = 20 * 数据行数，自适应"""
    ws = wb.add_worksheet(name)
    ws.merge_range(0, 0, 0, len(headers)-1, title, fmt_hdr)
    ws.set_row(0, 28)

    # 数据表
    for c, h in enumerate(headers):
        ws.write(2, c, h, fmt_hdr)
    for i, row in enumerate(data_rows):
        for j, val in enumerate(row):
            fmt = fmt_pct if isinstance(val, float) and 0 <= val <= 1 else \
                  fmt_num if isinstance(val, float) else fmt_cell
            ws.write(3+i, j, val, fmt)

    # 图表 — 高度随数据行数自适应
    chart_height = max(260, 20 * len(data_rows))
    chart = wb.add_chart({'type': chart_type})
    for s in series_list:
        chart.add_series(s)
    chart.set_style(10)
    chart.set_size({'width': chart_width, 'height': chart_height})
    if y_min is not None:
        chart.set_y_axis({'min': y_min, 'max': y_max})
    ws.insert_chart(3 + len(data_rows) + 2, 0, chart)
    return ws
```

---

## 图表类型全集

| `chart_type` | 场景 | y 轴建议 |
|---|---|---|
| `'column'` | 分组对比（城乡、年级） | `y_min=0, y_max=5`（Likert） |
| `'bar'` | 水平排序（因素、建议优先级） | 自动 |
| `'pie'` | 占比分布（意向、渠道） | 无 |
| `'line'` | 趋势变化（年级走势） | `y_min=0, y_max=1.0`（比例） |
| `'doughnut'` | 环形占比（与 pie 同场景，更美观） | 无 |
| `'scatter'` | 散点/相关性 | 自动 |
| `'radar'` | 多维度对比（如各维度认知度雷达图） | 需统一量纲 |
| `'area'` | 面积堆积（时间序列） | 自动 |

---

## 中文 Sheet 名：公式必须加单引号

```python
# ❌ 错误 — Sheet 名含中文时公式可能解析失败
'values': '=认知对比!$B$4:$B$7'

# ✅ 正确 — 加单引号
'values': "='认知对比'!$B$4:$B$7"
```

统一写法：
```python
def safe_range(sheet_name, col, start_row, end_row):
    return f"='{sheet_name}'!${col}${start_row}:${col}${end_row}"

# 在 make_chart_sheet 内部自动应用
# 直接在 series_list 的 values/categories 上包一层 safe_range
# 这样调用者永远不用操心引号问题
```

### make_chart_sheet 改造（加入自动引号）
在函数的 series 处理环节里，对每个 series：
```python
for s in series_list:
    if 'values' in s:
        s['values'] = f"='{name}'!{(s['values'].split('!',1)[1] if '!' in s['values'] else s['values'])}"
```

---

## 可选增强

```python
# 轴标题
chart.set_x_axis({'name': '年级'})
chart.set_y_axis({'name': '均值（1-5分）'})

# 图例位置
chart.set_legend({'position': 'bottom'})

# 数据标签（柱状图）
series = {'name': '均值', 'values': ..., 'categories': ...,
          'data_labels': {'value': True, 'num_format': '0.00'}}

# 数据标签（饼图）
series = {'name': '分布', 'values': ..., 'categories': ...,
          'data_labels': {'percentage': True, 'category': True}}
```

---

## 注意事项
- series 的 values/categories 必须用公式引用，不能传 Python list
- 行号从 data_rows 的行数推算，不要硬编码
- 饼图/环形图只支持一个 series
