# Reasonix 技能集 — 问卷调研数据分析

一套问卷调研数据分析技能，覆盖从问卷小程序导出到 Word 报告的全流程。8 个技能可独立使用，也可串联。

---

## 技能索引

| 技能 | 一句话 |
|---|---|
| `data-clean` | 问卷清洗（8 种题型全覆盖） |
| `docx-data-driven` | Word 报告数据驱动（A 模板分析 + C 异常标记） |
| `docx-font` | python-docx 可靠设中文字体 |
| `excel-stats` | 统计分析（Python 原生 + Excel + 询问安装） |
| `proj-layout` | 项目目录规范 |
| `py-format` | f-string 格式化速查 |
| `survey-pipeline` | 整合全流程（清洗→统计→图表→报告） |
| `xlsx-charts` | Excel 图表一图一 Sheet 零重叠 |

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

高级统计按需（先询问用户后安装）：
```bash
pip install scipy statsmodels scikit-learn
```

---

## 详情

### 1. data-clean — 问卷数据清洗
处理问卷星/腾讯问卷导出的原始数据，自动完成：表头识别（跳过题目行）、元数据剥离（IP/时间等列）、编码统一（男/Male/1→男）、无效卷剔除（全选同一选项）、缺失值处理（区分跳题缺失与真缺失）、填空混合路由（数值/混合/文本自动分类）、多选题格式兼容（0/1 多列↔逗号分隔）、矩阵/量表列组映射。

### 2. docx-data-driven — Word 报告数据驱动
所有数字从 Excel 实时读取计算填入，换数据重跑自动刷新。A 方案：条件分支模板分析，覆盖 ~80% 分析段落。C 方案：自动检测异常（城乡倒挂/极端值/逻辑矛盾），标记供人工介入。四层架构：COLUMN_MAP 配置 → load_data 读取 → compute_stats 计算 → render 渲染。

### 3. docx-font — 可靠设中文字体
必须同时设 `w:ascii`、`w:hAnsi`、`w:eastAsia` 三个 XML 属性。支持 Run 级别（正文/封面/表格）、Style 级别（Normal/Heading 1-3）。

### 4. excel-stats — 统计分析
纯 Python：频率、描述、Likert、交叉分组、卡方 + t 检验 + Pearson + Cronbach α。Excel 原生：描述统计/t 检验/ANOVA/回归/CORREL/CHISQ.TEST。需安装：多因素 ANOVA→statsmodels、因子分析→sklearn、Mann-Whitney→scipy（先询问）。

### 5. proj-layout — 目录规范
`data/raw/`（原始只读）+ `data/processed/`（清洗后）+ `scripts/` + `output/`（gitignore）+ `docs/`（变量说明）+ `archive/`（旧版归档）。含 .gitignore 和 README 模板。

### 6. py-format — f-string 速查
`:.1f` vs `:.1%` 的区别、百分比输入必须是小数（0.472）、None 值会炸 format、百分点差用 `:.1f` 不用 `:.1%`、报告常用句子模板。

### 7. survey-pipeline — 整合全流程
前 7 个技能的串联版。Part 0 目录 → Part 1 清洗 → Part 2 统计+图表 → Part 3 报告（含 A/C 动态分析）→ Part 4 格式化 → Part 5 一键 main.py。

### 8. xlsx-charts — Excel 图表
一图一 Sheet 垂直布局，8 种类型（column/bar/pie/line/doughnut/scatter/radar/area），自适应尺寸，中文 Sheet 名公式引号自动处理。
