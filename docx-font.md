# python-docx 可靠设置中文字体

## 问题
`run.font.name = '宋体'` 只设西文。中文仍可能回退到 MS Gothic 或其他默认字体，因为 `w:eastAsia` 属性未被设置。

## 方案
每个 Run / Style 必须同时设三个 XML 属性：`w:ascii`、`w:hAnsi`、`w:eastAsia`。必须用 lxml 的 `find()` + `set()` 操作 XML，不能靠 python-docx 的属性访问（命名空间会丢失）。

---

## Run 级别（正文段落、封面文字、表格单元格）

```python
from docx.shared import Pt
from docx.oxml.ns import qn
from lxml import etree

def set_run_font(run, font_name, size_pt, bold=False):
    """可靠设置 Run 的西文+东亚字体"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
```

## Style 级别（Normal、Heading 1/2/3）

```python
def set_style_font(style, font_name, size_pt, bold=False):
    """可靠设置 Style 的西文+东亚字体"""
    style.font.name = font_name
    style.font.size = Pt(size_pt)
    style.font.bold = bold
    rPr = style.element.find(qn('w:rPr'))
    if rPr is None:
        rPr = etree.SubElement(style.element, qn('w:rPr'))
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
```

---

## 典型用法

### 1. 全局样式（一次设完）
```python
# Normal 设了宋体 12pt → 大部分正文段落会自动继承
# ⚠️ 少数场景（如表单模板、存盘再打开）可能丢失继承
# 保险做法：add_para 里仍调 set_run_font，或通过模板 doc 打开
set_style_font(doc.styles['Normal'], '宋体', 12)
# 标题样式
set_style_font(doc.styles['Heading 1'], '微软雅黑', 22, bold=True)
set_style_font(doc.styles['Heading 2'], '微软雅黑', 16, bold=True)
set_style_font(doc.styles['Heading 3'], '微软雅黑', 14, bold=True)
```

### 2. 封面（不走 Heading 样式，需单独设）
```python
run = doc.add_paragraph().add_run('第四章  调研结果与现状分析')
set_run_font(run, '微软雅黑', 24, bold=True)      # 封面标题
set_run_font(run, '宋体', 16)                      # 封面副标题
set_run_font(run, '宋体', 14)                      # 署名
```

### 3. 表格单元格
```python
for cell in table.rows[0].cells:
    for run in cell.paragraphs[0].runs:
        set_run_font(run, '宋体', 10, bold=True)   # 表头
for row in table.rows[1:]:
    for cell in row.cells:
        for run in cell.paragraphs[0].runs:
            set_run_font(run, '宋体', 10)           # 表体
```
