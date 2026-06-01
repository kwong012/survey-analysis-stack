# python-docx 数据驱动报告生成

## 做什么
从 Excel 问卷数据自动生成 Word 分析报告。换数据 → 重跑脚本 → 数字 + 分析文字全部自动刷新。意外数据模式自动标记，人工介入一次替换。

---

## 四层架构

```
COLUMN_MAP + THRESHOLDS（配置层）  换问卷/调整度只改这里
    ↓
load_data() + validate()（读取层） 自动读 Excel + 质量检查
    ↓
compute_stats()（计算层）          按维度拆分为多个小函数
    ↓
render（渲染层）                   f-string 填入 + 条件分支分析
    ↓
detect_anomalies()（检测层）       自动标记意外 → 报告末尾 ⚠️ 清单
```

---

## 第一层：集中配置 config.py

```python
DATA_FILE = "data.xlsx"

COLUMN_MAP = {
    'gender': 2, 'grade': 3, 'major': 4, 'hukou': 6,
    'aware_gba': 9, 'aware_bqw': 10, 'aware_incub': 11, 'aware_policy': 12,
    'participated': 18, 'will_part': 22,
}

# 阈值配置（换课题时调整，不用改代码）
THRESHOLDS = {
    'gap_severe': 0.60,      # 城乡差距"严重分化"阈值
    'gap_moderate': 0.40,    # "差异明显"阈值
    'gap_mild': 0.20,        # "差异可控"阈值
    'participation_severe': 0.40,  # 意愿-参与缺口"严重"阈值
    'participation_moderate': 0.20,
    'extreme_high': 0.95,    # 极端高值警报
    'extreme_low': 0.05,     # 极端低值警报
    'missing_high': 0.50,    # 缺失率过高警报
}
```

---

## 第二层：数据加载 load.py

```python
import openpyxl

def load_data(path, column_map):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    N = ws.max_row - 1
    data = {'N': N}
    for var, col_idx in column_map.items():
        data[var] = [ws.cell(row=r, column=col_idx).value for r in range(2, N+2)]
    return data

def validate(d, thresholds):
    issues = []
    N = d['N']
    # 清洗：标记无效卷（全选同一选项的）
    for key in ['aware_gba','aware_bqw','aware_incub','aware_policy']:
        vals = [int(x) for x in d[key] if x is not None]
        if any(not (1 <= x <= 5) for x in vals):
            issues.append(f'{key} 存在超出 1-5 范围的值')
    for key in d:
        if key != 'N':
            missing = sum(1 for x in d[key] if x is None) / N
            if missing > thresholds['missing_high']:
                issues.append(f'{key} 缺失率 {missing:.0%}，过高')
    return issues
```

---

## 第三层：统计计算 stats.py — 按维度拆分

```python
from statistics import mean

def compute_awareness_stats(d):
    """认知维度统计"""
    s = {}
    N = d['N']
    for key in ['aware_gba','aware_bqw','aware_incub','aware_policy']:
        clean = [int(x) for x in d[key] if x is not None]
        s[f'{key}_mean'] = mean(clean) if clean else 0
        s[f'{key}_high'] = sum(1 for x in clean if x >= 4) / N if clean else 0
    # 城乡分组：直接在过滤 None 的同时按户籍条件取值，不依赖原始行号
    for key in ['aware_gba','aware_bqw','aware_incub','aware_policy']:
        urban = [int(d[key][i]) for i in range(N) if d['hukou'][i]=='城镇' and d[key][i] is not None]
        rural = [int(d[key][i]) for i in range(N) if d['hukou'][i]=='农村' and d[key][i] is not None]
        s[f'{key}_urban'] = mean(urban) if urban else 0
        s[f'{key}_rural'] = mean(rural) if rural else 0
        s[f'{key}_diff'] = s[f'{key}_urban'] - s[f'{key}_rural']
    return s

def compute_participation_stats(d):
    """参与维度统计"""
    s = {}
    N = d['N']
    s['part_rate'] = sum(1 for x in d['participated'] if x=='是')/N
    s['will_rate'] = sum(1 for x in d['will_part'] if x=='是')/N
    s['gap'] = s['will_rate'] - s['part_rate']
    return s

# 主函数聚合
def compute_all_stats(d, thresholds):
    s = {}
    s.update(compute_awareness_stats(d))
    s.update(compute_participation_stats(d))
    # ... 按需扩展
    return s
```

---

## 第四层：渲染 render.py + A 方案 + C 方案

```python
from docx import Document

def render_report(stats, d, thresholds):
    doc = Document()
    # 样式设置 — 配合 docx-font 技能

    # ── 数字填入 ──
    add_para(f'本次回收有效问卷 {d["N"]} 份。')

    # ── A 方案：阈值来自配置，不硬编码 ──
    diff = stats['aware_gba_diff']
    t = thresholds
    if diff >= t['gap_severe']:
        analysis = f'城乡认知差距高达 {diff:.2f} 分，处于严重分化区间。'
    elif diff >= t['gap_moderate']:
        analysis = f'城乡认知差距为 {diff:.2f} 分，差异明显。'
    elif diff >= t['gap_mild']:
        analysis = f'城乡认知差距仅 {diff:.2f} 分，差异可控。'
    else:
        analysis = f'城乡差距仅 {diff:.2f} 分或倒挂，意外模式。'
    add_para(analysis)

    # ── C 方案：异常标记 ──
    alerts = detect_anomalies(stats, thresholds)
    if alerts:
        add_para('─' * 40)
        add_para('⚠️ 以下段落需人工分析：', bold=True)
        for i, a in enumerate(alerts, 1):
            add_para(f'  {i}. {a}')

    doc.save('report.docx')
```

---

## 检测层：C 方案规则

```python
def detect_anomalies(stats, thresholds):
    t = thresholds
    alerts = []
    for key in ['aware_gba','aware_bqw','aware_incub','aware_policy']:
        d = stats.get(f'{key}_diff', 0)
        if d > t['gap_severe']:
            alerts.append(f'{key} 城乡差距 {d:.2f}，严重分化（>{t["gap_severe"]}）')
        if d < 0:
            alerts.append(f'{key} 农村反超城镇 {abs(d):.2f} 分——意外倒挂')
    if stats.get('part_rate',0) > stats.get('will_rate',0):
        alerts.append('参与率 > 意愿率——逻辑异常')
    for key in ['aware_gba_high','aware_bqw_high','aware_incub_high','aware_policy_high']:
        v = stats.get(key, 0)
        if v > t['extreme_high']:
            alerts.append(f'{key} = {v:.0%}，极端高值，疑有无效样本')
        if v < t['extreme_low']:
            alerts.append(f'{key} = {v:.0%}，极端低值，疑有无效样本')
    return alerts
```

---

## 增量更新策略

先收 100 份跑初版，再补到 320 份跑定版：
```python
# 初版
data = load_data('data/batch1_100.xlsx', COLUMN_MAP)
render(data, 'output/初版.docx')

# 定版 — 换文件，重跑
data = load_data('data/final_320.xlsx', COLUMN_MAP)
render(data, 'output/定版.docx')
```

数据文件命名：`data/YYYY-MM-DD-N份.xlsx`，不覆盖历史。

---

## Word 模板模式（可选）

如果有现成的 .docx 模板（含页眉页脚、校徽），用 `python-docx` 打开模板后只在书签/占位符处填充：

```python
doc = Document('template.docx')  # 模板含页眉页脚
# 找到占位符段落，替换文本
for p in doc.paragraphs:
    if '{{N}}' in p.text:
        p.text = p.text.replace('{{N}}', str(N))
doc.save('output/report.docx')
```

---

## 会出问题的地方

| 问题 | 对策 |
|---|---|
| missing 值炸 mean() | `[x for x in col if x is not None]` |
| 数字类型混 | `int(x)` 统一 |
| 列号漂移 | 读表头按名称匹配 |
| 样本量写死 | `f"回收 {N} 份"` |
| stats 膨胀 | 按维度拆分小函数 |
| 阈值散落代码 | 集中到 `THRESHOLDS` 字典 |
