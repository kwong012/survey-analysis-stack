---
name: survey-pipeline
description: 问卷到报告全流程 — 数据清洗→统计分析→Excel图表→Word报告，数据驱动动态分析，一份配置控制全线
---
# 问卷调研全流程自动化

## 流程

```
问卷小程序导出 (data/raw/)
    ↓
Phase 1: 数据清洗 → data/processed/问卷数据_320份.xlsx
    ↓
Phase 2: 统计分析 + 图表 → output/统计分析结果.xlsx + output/分析图表.xlsx
    ↓
Phase 3: Word 报告 → output/报告.docx（数字+分析全刷新）
```

---

## Part 0: 项目目录

```
project/
├── data/
│   ├── raw/                    ← 原始导出，只读
│   └── processed/              ← 清洗后数据
├── output/                     ← 脚本产出（可删，.gitignore）
├── scripts/
│   ├── config.py               ← 唯一需改的配置
│   └── main.py                 ← 一键入口
├── docs/                       ← 变量说明、方法记录
├── archive/                    ← 旧版归档
├── .gitignore
└── README.md
```

### .gitignore

```gitignore
output/
data/
__pycache__/
*.pyc
.DS_Store
Thumbs.db
```

### README

```markdown
# 项目名

## 快速开始
1. 问卷导出放 `data/raw/`
2. 调整 `scripts/config.py` 的 COLUMN_MAP
3. `python scripts/main.py`
4. 产出在 `output/`，检查报告末尾 ⚠️ 标记

## 依赖
pip install openpyxl xlsxwriter python-docx lxml
```

---

## Part 1: 数据清洗

### 题型总览

| 题型 | 清洗动作 |
|---|---|
| 单选 | 编码归一化 |
| 多选 | 0/1列 或 逗号分隔 → 统一 0/1 |
| 判断 | 是/否 → 1/0 |
| Likert | 范围校验 1-5 |
| 填空数值 | float转换 |
| 填空混合 | 提取数字 + 标记 |
| 填空文本 | 去空格 + 词频 |
| 主观题 | 不处理，直接导出 |

### 核心函数

```python
import openpyxl, re
from statistics import mean, stdev, mode, median

META_KEYWORDS = ['IP','ip','开始时间','结束时间','答题时长','所在地','城市','序号','编号','提交时间','openid','OpenId']

def detect_header_rows(ws):
    for r in range(1, min(5, ws.max_row+1)):
        v = str(ws.cell(row=r, column=2).value or '')
        if v in ('男','女','大一','大二','大三','大四','1','2','是','否') or v.isdigit():
            return (r-1, r)
    return (1, 2)

def strip_meta_columns(ws, header_row):
    cols = {}
    for c in range(1, ws.max_column+1):
        name = str(ws.cell(row=header_row, column=c).value or '')
        if not any(kw in name for kw in META_KEYWORDS):
            cols[c] = name
    return cols

ENCODING_RULES = {
    'gender': {'男':'男','male':'男','Male':'男','1':'男','女':'女','female':'女','Female':'女','2':'女'},
    'yn': {'是':'是','yes':'是','Yes':'是','1':'是','否':'否','no':'否','No':'否','2':'否'},
}

def normalize(raw_data, rules):
    for key, mapping in rules.items():
        if key in raw_data:
            raw_data[key] = [mapping.get(str(v), v) for v in raw_data[key]]
    return raw_data

def normalize_yes_no(data, key):
    yes_set = {'是','yes','Yes','YES','1','对','true','True','√'}
    no_set = {'否','no','No','NO','2','错','false','False','×'}
    data[key] = [1 if str(v).strip() in yes_set else 0 if str(v).strip() in no_set else None for v in data[key]]
    return data

def remove_invalid(data):
    N = data['N']; keep = list(range(N)); removed = []
    for i in range(N):
        likert_keys = [k for k in data if k.startswith('aware_')]
        vals = [data[k][i] for k in likert_keys if data[k][i] is not None]
        if len(vals) >= 3 and len(set(vals)) == 1:
            removed.append({'row': i+2+data.get('_data_start',2), 'reason': '全选同一选项'})
            keep.remove(i)
    return keep, removed

def filter_by_index(data, keep_idx):
    new = {'N': len(keep_idx)}
    for k, v in data.items():
        if k == 'N': continue
        new[k] = [v[i] for i in keep_idx]
    if '_data_start' in data: new['_data_start'] = data['_data_start']
    return new

SKIP_LOGIC = {
    'activity_type': {'dep': 'participated', 'trigger': ['是','有参与','参加']},
}

def handle_missing(data, strategy='mode'):
    N = data['N']; to_drop = set()
    for key in list(data.keys()):
        if key in ('N','_data_start') or key.startswith('_'): continue
        vals = data[key]; none_idx = [i for i in range(N) if vals[i] is None or vals[i]=='']
        if not none_idx: continue
        if key in SKIP_LOGIC:
            rule = SKIP_LOGIC[key]; triggers = rule['trigger']
            real = [i for i in none_idx if data[rule['dep']][i] in triggers]
            if not real: continue
            none_idx = real
        miss_rate = len(none_idx)/N
        if miss_rate > 0.30: to_drop.update(none_idx)
        elif strategy == 'mode':
            non_none = [v for v in vals if v not in (None,'')]
            fill = max(set(non_none), key=non_none.count) if non_none else 0
            for i in none_idx: vals[i] = fill
        elif strategy == 'flag':
            for i in none_idx: vals[i] = -1
    if to_drop:
        keep = [i for i in range(N) if i not in to_drop]
        return keep, list(to_drop)
    return list(range(N)), []

def classify_fill_column(data, key):
    N = data['N']; vals = [str(data[key][i] or '').strip() for i in range(N)]
    pure_num = has_num = pure_text = 0
    for v in vals:
        if not v: continue
        try: float(v); pure_num += 1
        except ValueError:
            nums = re.findall(r'\d+\.?\d*', v)
            if nums: has_num += 1
            else: pure_text += 1
    total = pure_num + has_num + pure_text
    if total == 0: return 'empty', {}
    if pure_num/total >= 0.90: return 'numeric', {}
    if (pure_num+has_num)/total >= 0.50: return 'mixed', {'extracted': has_num, 'total': total}
    return 'text', {}

def clean_fill_by_type(data, key):
    col_type, info = classify_fill_column(data, key)
    N = data['N']; flags = []
    if col_type == 'numeric':
        data[key] = [float(str(v).strip()) if v is not None and str(v).strip() else None for v in data[key]]
    elif col_type == 'mixed':
        new_vals = []
        for i in range(N):
            v = str(data[key][i] or '').strip()
            if not v: new_vals.append(None)
            else:
                try: new_vals.append(float(v))
                except ValueError:
                    nums = re.findall(r'\d+\.?\d*', v)
                    if nums:
                        new_vals.append(float(nums[0]))
                        flags.append((i, f'{key}: "{v}" → {nums[0]}'))
                    else:
                        new_vals.append(None)
                        flags.append((i, f'{key}: "{v}" 无数字 → 置空'))
        data[key] = new_vals
    elif col_type == 'text':
        data[key] = [str(v).strip() if v is not None and str(v).strip() else None for v in data[key]]
    return col_type, flags

def handle_open_ended(data, key):
    N = data['N']; vals = [str(data[key][i] or '').strip() for i in range(N)]
    responded = [v for v in vals if v]
    return {'key': key, '应答率': len(responded)/N,
            '平均字数': sum(len(v) for v in responded)/len(responded) if responded else 0,
            '响应文本': responded}

def normalize_multi_select(data, col_key):
    N = data['N']; vals = [str(data[col_key][i] or '') for i in range(N)]
    sample = [v for v in vals if v and v not in ('0','1','None','')]
    if sample and any(',' in v for v in sample):
        all_opts = set()
        for v in sample: all_opts.update(v.split(','))
        for opt in sorted(all_opts):
            data[f'{col_key}_{opt}'] = [1 if opt in str(data[col_key][i] or '') else 0 for i in range(N)]
        return True
    return False

def load_matrix_columns(raw_data, groups):
    for group_name, col_names in groups.items():
        for i, cn in enumerate(col_names, 1):
            raw_data[f'{group_name}_{i}'] = raw_data[cn]
    return raw_data

def cleaning_report(raw_N, keep_idx, removed, fill_flags, open_stats):
    return {
        '原始样本': raw_N, '有效样本': len(keep_idx),
        '剔除数量': raw_N-len(keep_idx),
        '剔除率': f'{(raw_N-len(keep_idx))/raw_N:.1%}' if raw_N else '0%',
        '剔除明细': removed, '填空提取标记': fill_flags, '主观题摘要': open_stats,
    }
```

---

## Part 2: 统计分析

### 纯 Python 方法

```python
from statistics import mean, stdev, mode, median

def freq(vals):
    d = {}
    for v in vals: d[v] = d.get(v, 0) + 1
    return [(k, v, v/len(vals)) for k, v in sorted(d.items(), key=lambda x:-x[1])]

def describe(col):
    clean = [x for x in col if x is not None]
    return {'n':len(clean),'mean':mean(clean),'median':median(clean),
            'stdev':stdev(clean),'min':min(clean),'max':max(clean),'mode':mode(clean)}

def likert_analysis(col):
    n = len(col)
    return {'top2':sum(1 for x in col if x in(4,5))/n,
            'bottom2':sum(1 for x in col if x in(1,2))/n,
            'net':(sum(1 for x in col if x in(4,5))-sum(1 for x in col if x in(1,2)))/n}

def multi_select(columns, labels):
    n = len(columns[0])
    return [(labels[i], sum(col), sum(col)/n) for i, col in enumerate(columns)]

def group_rate(group_col, condition_fn):
    result = {}
    for g in set(group_col):
        idxs = [i for i, grp in enumerate(group_col) if grp==g]
        if idxs: result[g] = sum(1 for i in idxs if condition_fn(i))/len(idxs)
    return result

def chi_square(crosstab, row_labels, col_labels):
    row_tot = [sum(crosstab.get(r,{}).get(c,0) for c in col_labels) for r in row_labels]
    col_tot = [sum(crosstab.get(r,{}).get(c,0) for r in row_labels) for c in col_labels]
    total = sum(row_tot); chi2 = 0.0
    for i, r in enumerate(row_labels):
        for j, c in enumerate(col_labels):
            o = crosstab.get(r,{}).get(c,0)
            e = row_tot[i]*col_tot[j]/total if total>0 else 1
            if e>0: chi2 += (o-e)**2/e
    df = (len(row_labels)-1)*(len(col_labels)-1)
    return chi2, df

def chi_sq_pvalue(chi2, df, alpha=0.05):
    table = {0.05:{1:3.84,2:5.99,3:7.81,4:9.49,5:11.07,6:12.59,7:14.07,8:15.51,9:16.92,10:18.31},
             0.01:{1:6.63,2:9.21,3:11.34,4:13.28,5:15.09,6:16.81,7:18.48,8:20.09,9:21.67,10:23.21}}
    t = table.get(alpha,{})
    if df in t: return f'<{alpha}' if chi2>t[df] else f'>={alpha}'
    z = (2*chi2)**0.5 - (2*df-1)**0.5
    return f'<{alpha}' if abs(z)>=1.96 else f'>={alpha}'

def pearson_r(x, y):
    n = len(x); mx, my = mean(x), mean(y); sx, sy = stdev(x), stdev(y)
    cov = sum((x[i]-mx)*(y[i]-my) for i in range(n))/(n-1)
    return cov/(sx*sy) if sx and sy else 0

def ttest_ind(group_a, group_b):
    ma, mb = mean(group_a), mean(group_b); na, nb = len(group_a), len(group_b)
    if na<2 or nb<2: return None, '样本量不足'
    va, vb = stdev(group_a)**2, stdev(group_b)**2
    se = ((va/na)+(vb/nb))**0.5
    if se==0: return None, '标准误为0'
    t = (ma-mb)/se
    df_num = ((va/na)+(vb/nb))**2
    df_den = ((va/na)**2/(na-1))+((vb/nb)**2/(nb-1))
    df = df_num/df_den if df_den else 1
    if df>=60: p = '<0.05' if abs(t)>1.96 else '>=0.05'
    else: p = f'需查t表(df={df:.0f})'
    return (t, p)

def cronbach_alpha(items):
    k = len(items); n = len(items[0])
    vt = stdev([sum(items[j][i] for j in range(k)) for i in range(n)])**2
    vi = sum(stdev(col)**2 for col in items)
    return (k/(k-1))*(1-vi/vt) if vt else 0
```

### Excel 原生（分析工具库）

文件 → 选项 → 加载项 → 分析工具库

描述统计 / t检验 / ANOVA / 回归 / `=CORREL()` / `=CHISQ.TEST()`

### 需安装的

先问"需安装 xxx，确认？"。ANOVA 多因素→`statsmodels`、因子分析→`sklearn`、Mann-Whitney→`scipy`

---

## Part 2b: Excel 图表（一图一 Sheet）

```python
import xlsxwriter

def make_chart_sheet(wb, name, title, headers, data_rows, chart_type, series_list,
                     y_min=None, y_max=None, chart_width=600):
    ws = wb.add_worksheet(name)
    fmt_hdr = wb.add_format({'bold':True,'bg_color':'#4472C4','font_color':'white','border':1,'align':'center'})
    fmt_cell = wb.add_format({'border':1,'align':'center'})
    fmt_pct = wb.add_format({'border':1,'align':'center','num_format':'0.0%'})
    fmt_num = wb.add_format({'border':1,'align':'center','num_format':'0.00'})
    ws.merge_range(0, 0, 0, len(headers)-1, title, fmt_hdr); ws.set_row(0, 28)
    for c, h in enumerate(headers): ws.write(2, c, h, fmt_hdr)
    for i, row in enumerate(data_rows):
        for j, val in enumerate(row):
            fmt = fmt_pct if isinstance(val,float) and 0<=val<=1 else fmt_num if isinstance(val,float) else fmt_cell
            ws.write(3+i, j, val, fmt)
    chart = wb.add_chart({'type': chart_type})
    for s in series_list:
        if 'values' in s:
            s['values'] = f"='{name}'!{(s['values'].split('!',1)[1] if '!' in s['values'] else s['values'])}"
        chart.add_series(s)
    chart.set_style(10); chart.set_size({'width':chart_width,'height':max(260,20*len(data_rows))})
    if y_min is not None: chart.set_y_axis({'min':y_min,'max':y_max})
    ws.insert_chart(3+len(data_rows)+2, 0, chart)
    return ws
```

图表类型：`'column'` `'bar'` `'pie'` `'line'` `'doughnut'` `'scatter'` `'radar'` `'area'`

---

## Part 3: Word 报告（数据驱动 + A/C 动态分析）

### Word 字体

```python
from docx.shared import Pt
from docx.oxml.ns import qn
from lxml import etree

def set_run_font(run, font_name, size_pt, bold=False):
    run.font.name = font_name; run.font.size = Pt(size_pt); run.bold = bold
    rPr = run._element.get_or_add_rPr(); rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None: rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    for a in ('w:ascii','w:hAnsi','w:eastAsia'): rFonts.set(qn(a), font_name)

def set_style_font(style, font_name, size_pt, bold=False):
    style.font.name = font_name; style.font.size = Pt(size_pt); style.font.bold = bold
    rPr = style.element.find(qn('w:rPr'))
    if rPr is None: rPr = etree.SubElement(style.element, qn('w:rPr'))
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None: rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    for a in ('w:ascii','w:hAnsi','w:eastAsia'): rFonts.set(qn(a), font_name)
```

### 四层架构

```
config.py (COLUMN_MAP + THRESHOLDS)
    ↓
load_data() → 读 Excel
    ↓
compute_stats() → 按维度计算
    ↓
render() → f-string 填入 + 条件分支分析 + 异常标记
```

### 配置文件

```python
COLUMN_MAP = {
    'gender':2,'grade':3,'major':4,'hukou':6,
    'aware_gba':9,'aware_bqw':10,'aware_incub':11,'aware_policy':12,
    'participated':18,'will_part':22,
}
THRESHOLDS = {
    'gap_severe':0.60,'gap_moderate':0.40,'gap_mild':0.20,
    'extreme_high':0.95,'extreme_low':0.05,'missing_high':0.50,
}
QUESTION_TYPES = {
    'gender':('single',{}),'grade':('single',{}),
    'aware_gba':('likert',{'min':1,'max':5}),
    'participated':('yesno',{}),'suggestion':('open',{}),
}
```

### 统计计算（按维度）

```python
def compute_awareness_stats(d):
    s = {}; N = d['N']
    for key in ['aware_gba','aware_bqw','aware_incub','aware_policy']:
        clean = [int(x) for x in d[key] if x is not None]
        s[f'{key}_mean'] = mean(clean) if clean else 0
        s[f'{key}_high'] = sum(1 for x in clean if x>=4)/N if clean else 0
    for key in ['aware_gba','aware_bqw','aware_incub','aware_policy']:
        urban = [int(d[key][i]) for i in range(N) if d['hukou'][i]=='城镇' and d[key][i] is not None]
        rural = [int(d[key][i]) for i in range(N) if d['hukou'][i]=='农村' and d[key][i] is not None]
        s[f'{key}_urban'] = mean(urban) if urban else 0
        s[f'{key}_rural'] = mean(rural) if rural else 0
        s[f'{key}_diff'] = s[f'{key}_urban'] - s[f'{key}_rural']
    return s

def compute_participation_stats(d):
    N = d['N']
    return {'part_rate': sum(1 for x in d['participated'] if x=='是')/N,
            'will_rate': sum(1 for x in d['will_part'] if x=='是')/N}
```

### 渲染 + A 方案 + C 方案

```python
def render_report(stats, d, thresholds):
    doc = Document(); set_style_font(doc.styles['Normal'], '宋体', 12)
    for i,sz in [(1,22),(2,16),(3,14)]:
        set_style_font(doc.styles[f'Heading {i}'], '微软雅黑', sz, True)

    def p(text, bold=False, indent=False, size=12):
        pp = doc.add_paragraph(); pp.paragraph_format.line_spacing = 1.5
        if indent: pp.paragraph_format.first_line_indent = Cm(0.74)
        r = pp.add_run(text); set_run_font(r, '宋体', size, bold); return pp

    def st(text):
        pp = doc.add_paragraph(); pp.paragraph_format.line_spacing = 1.5
        pp.paragraph_format.space_before = Pt(6)
        r = pp.add_run(text); set_run_font(r, '宋体', 13, True); return pp

    p(f'本次回收有效问卷 {d["N"]} 份。')

    # A 方案：条件分支
    diff = stats['aware_gba_diff']; t = thresholds
    if diff >= t['gap_severe']:
        p(f'城乡认知差距高达 {diff:.2f} 分，处于严重分化区间。')
    elif diff >= t['gap_moderate']:
        p(f'城乡认知差距为 {diff:.2f} 分，差异明显。')
    elif diff >= t['gap_mild']:
        p(f'城乡认知差距仅 {diff:.2f} 分，差异可控。')
    else:
        p(f'城乡差距仅 {diff:.2f} 分或倒挂，意外模式。')

    # C 方案：异常标记
    alerts = detect_anomalies(stats, thresholds)
    if alerts:
        p('─'*40); st('⚠️ 以下段落需人工分析：')
        for i, a in enumerate(alerts, 1): p(f'  {i}. {a}')

    doc.save('output/报告.docx')
```

### C 方案规则

```python
def detect_anomalies(stats, thresholds):
    t = thresholds; alerts = []
    for key in ['aware_gba','aware_bqw','aware_incub','aware_policy']:
        d = stats.get(f'{key}_diff',0)
        if d > t['gap_severe']: alerts.append(f'{key} 城乡差距 {d:.2f}，严重分化')
        if d < 0: alerts.append(f'{key} 农村反超城镇 {abs(d):.2f} 分——倒挂')
    if stats.get('part_rate',0) > stats.get('will_rate',0):
        alerts.append('参与率 > 意愿率——逻辑异常')
    for key in ['aware_gba_high','aware_bqw_high','aware_incub_high','aware_policy_high']:
        v = stats.get(key,0)
        if v > t['extreme_high']: alerts.append(f'{key} = {v:.0%}，极端高值')
        if v < t['extreme_low']: alerts.append(f'{key} = {v:.0%}，极端低值')
    return alerts
```

---

## Part 4: 格式化速查

### `:.1f` vs `:.1%`（最易错）

| 写法 | 输入 | 输出 | 说明 |
|---|---|---|---|
| `:.1f` | 0.472 | 0.5 | 浮点 |
| `:.1%` | 0.472 | 47.2% | 自动×100+% |

百分比输入必须是小数（0.472），不是 47.2。

### 报告常用

```python
f"本次共回收有效问卷 {N} 份。"
f"均值 {m:.2f}，城镇 {urban:.2f}，农村 {rural:.2f}，差距 {diff:.2f} 分。"
f"了解率 {pct:.0%}，较上期提升 {delta:.1f} 个百分点。"
f"{cnt}/{total}（{cnt/total:.1%}）"
f"从大一 {fresh:.0%} 升至研究生 {grad:.0%}"
```

### 易错

| 错误 | 正确 |
|---|---|
| `f"{47.2:.1%}"` → 4720.0% | 确保是 0~1 小数 |
| `f"{None:.2f}"` TypeError | `v if v is not None else 0` |
| 百分点差用 `:.1%` → 0.5% | 用 `:.1f`，注明"个百分点" |

---

## Part 5: 一键入口 main.py

```python
import openpyxl

# 1. 清洗
wb = openpyxl.load_workbook('data/raw/问卷导出.xlsx')
ws = wb.active; header_row, data_start = detect_header_rows(ws)
columns = strip_meta_columns(ws, header_row)
raw_N = ws.max_row - data_start + 1
raw = {'N': raw_N, '_data_start': data_start}
for col_idx, name in columns.items():
    raw[name] = [ws.cell(row=r, column=col_idx).value for r in range(data_start, ws.max_row+1)]

keep, removed = remove_invalid(raw); raw = filter_by_index(raw, keep)
raw = normalize(raw, ENCODING_RULES)

all_flags = []; open_stats = []
for key, (qtype, params) in QUESTION_TYPES.items():
    if key not in raw: continue
    if qtype == 'yesno': normalize_yes_no(raw, key)
    elif qtype == 'fill': _, flags = clean_fill_by_type(raw, key); all_flags.extend(flags)
    elif qtype == 'open': open_stats.append(handle_open_ended(raw, key))

keep2, dropped = handle_missing(raw, 'mode'); raw = filter_by_index(raw, keep2)
rpt = cleaning_report(raw_N, keep, removed+dropped, all_flags, open_stats)
print(f"原始 {rpt['原始样本']} → 有效 {rpt['有效样本']}（剔除 {rpt['剔除率']}）")

# 2. 分析 + 图表
stats = compute_all_stats(raw, THRESHOLDS)
# ... 生成统计分析结果.xlsx + 分析图表.xlsx（调用 Part 2/2b 函数）

# 3. 报告
render_report(stats, raw, THRESHOLDS)
print("✅ 全流程完成，产出在 output/")
```

---

## 换数据操作

```
1. 问卷导出放 data/raw/
2. 改 config.py 的 COLUMN_MAP 和 QUESTION_TYPES
3. python scripts/main.py
4. 翻到报告末尾，看 ⚠️ 标记 → 贴数据给 AI → 替换分析段落
```

---

## 不适用

- 纯开放式访谈（没结构化数据）
- 问卷还在频繁改版
- 一篇过的短报告（不值得搭框架）
