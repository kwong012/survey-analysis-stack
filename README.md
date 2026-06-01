# Reasonix 技能集 — 问卷调研数据分析

这是一套用于问卷调研数据分析的技能（Skill），覆盖从问卷小程序导出数据到生成 Word 分析报告的全流程。8 个技能可独立使用，也可串联。

---

## 技能清单

### 1. data-clean — 问卷数据清洗
问卷星/腾讯问卷导出的原始数据往往有题目行、元数据列、编码不一致、缺失值、无效卷等问题。本技能自动处理 8 种题型：单选、多选、判断、Likert 量表、填空（数值/混合/文本）、主观题。

### 2. docx-data-driven — Word 报告数据驱动
从 Excel 问卷数据自动生成 Word 分析报告。所有数字实时计算填入，而非手写。A 方案（条件分支模板分析）自动覆盖 ~80% 的分析段落，C 方案（异常检测）标记意外数据模式供人工介入。

### 3. docx-font — python-docx 可靠设中文字体
`run.font.name = '宋体'` 只设西文字体，中文可能回退为 MS Gothic。本技能通过同时设置 `w:ascii`、`w:hAnsi`、`w:eastAsia` 三个 XML 属性解决此问题。支持 Run、Style、表格单元格、封面文字。

### 4. excel-stats — 统计分析工具箱
无需 SPSS 即可完成问卷统计分析。纯 Python 支持频率、描述、Likert、交叉分组、卡方检验、t 检验、Pearson 相关、Cronbach α。对 Excel 原生功能（描述统计/ANOVA/回归）直接调用。对需要 numpy/scipy 的高级方法（多因素 ANOVA/因子分析/Mann-Whitney），询问用户后安装。

### 5. proj-layout — 项目目录规范
标准化的问卷项目目录结构：`data/raw/`（原始只读）、`data/processed/`（清洗后）、`scripts/`（Python 脚本）、`output/`（产物，gitignore）、`docs/`（文档）、`archive/`（归档）。含 .gitignore 和 README 模板。

### 6. py-format — f-string 格式化速查
写数据驱动报告时常遇到格式化问题：`:.1f` vs `:.1%` 的区别、百分比输入必须是小数、None 值会炸 format、百分点差用 `:.1f` 而不用 `:.1%`。附报告常用句子模板和易错点速查。

### 7. survey-pipeline — 整合全流程
以上 7 个技能的串联版本。从问卷导出到 Word 报告一键贯通：Part 0 项目目录 → Part 1 数据清洗 → Part 2 统计分析 + Excel 图表 → Part 3 Word 报告（含 A/C 动态分析）→ Part 4 格式化速查 → Part 5 一键 main.py。

### 8. xlsx-charts — Excel 图表可靠排版
xlsxwriter 多图放同一 Sheet 必重叠。本技能采用一图一 Sheet、数据表在上图表在下的垂直布局，支持 8 种图表类型和自适应尺寸。自动处理中文 Sheet 名公式引号。

---

## 使用建议

| 场景 | 用哪个 |
|---|---|
| 只想清洗数据 | `data-clean` |
| 只想做统计分析 | `excel-stats` |
| 只想生成图表 | `xlsx-charts` |
| 只想写 Word 报告 | `docx-data-driven` + `docx-font` |
| 从头到尾一套走 | `survey-pipeline` |

---

## 依赖

```bash
pip install openpyxl xlsxwriter python-docx lxml
```

高级统计需额外安装（按需询问用户后安装）：
```bash
pip install scipy statsmodels scikit-learn
```
