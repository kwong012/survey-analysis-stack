# 问卷数据清洗

## 在工作流中的位置

```
问卷星/腾讯问卷导出（原始）→ data-clean → data/processed/（干净数据）
```

---

## 题型总览与清洗策略

| 题型 | 数据形态 | 清洗动作 | 后续分析 |
|---|---|---|---|
| 单选 | 单列 | 编码归一化 | `freq()` |
| 多选 | 多列0/1 或 逗号分隔 | 格式统一→0/1列 | `multi_select()` |
| 判断（是/否） | 两选项 | 编码归一化→1/0 | `freq()` |
| Likert量表 | 1-5数字 | 范围校验 | `describe()`, `likert_analysis()` |
| 填空（纯数值） | 全列数字 | float转换 + 离群检测 | `describe()` |
| 填空（混合） | 部分含文字 | 提取数值 + 标记确认 | 人工审核后走数值统计 |
| 填空（纯文本） | 短文本 | 去空格 + 词频 | 词频/手动分类 |
| 主观题 | 长文本 | **不处理，直接导出** | 人工编码 |

---

## 动作1：表头识别

```python
def detect_header_rows(ws):
    """跳过题目行，返回 (表头行号, 数据起始行号)"""
    for r in range(1, min(5, ws.max_row + 1)):
        first_cell = str(ws.cell(row=r, column=2).value or '')
        if first_cell in ('男','女','大一','大二','大三','大四','1','2','是','否') \
           or first_cell.isdigit():
            return (r - 1, r)
    return (1, 2)
```

---

## 动作2：元数据列剥离

```python
META_KEYWORDS = ['IP', 'ip', '开始时间', '结束时间', '答题时长', '所在地',
                 '城市', '序号', '编号', '提交时间', 'openid', 'OpenId']

def strip_meta_columns(ws, header_row):
    cols = {}
    for c in range(1, ws.max_column + 1):
        name = str(ws.cell(row=header_row, column=c).value or '')
        if any(kw in name for kw in META_KEYWORDS):
            continue
        cols[c] = name
    return cols
```

---

## 动作3：编码统一

```python
ENCODING_RULES = {
    'gender': {'男':'男', 'male':'男', 'Male':'男', '1':'男',
               '女':'女', 'female':'女', 'Female':'女', '2':'女'},
    'yn':     {'是':'是', 'yes':'是', 'Yes':'是', '1':'是',
               '否':'否', 'no':'否', 'No':'否', '2':'否'},
}

def normalize(raw_data, rules):
    for key, mapping in rules.items():
        if key in raw_data:
            raw_data[key] = [mapping.get(str(v), v) for v in raw_data[key]]
    return raw_data
```

---

## 动作4：判断/是非题归一

```python
def normalize_yes_no(data, key):
    yes_set = {'是','yes','Yes','YES','1','对','true','True','√'}
    no_set  = {'否','no','No','NO','2','错','false','False','×'}
    data[key] = [1 if str(v).strip() in yes_set else 0 if str(v).strip() in no_set else None
                 for v in data[key]]
    return data
```

---

## 动作5：无效卷剔除

```python
def remove_invalid(data):
    N = data['N']
    keep = list(range(N))
    removed = []
    for i in range(N):
        likert_keys = [k for k in data if k.startswith('aware_')]
        vals = [data[k][i] for k in likert_keys if data[k][i] is not None]
        if len(vals) >= 3 and len(set(vals)) == 1:
            removed.append({'row': i+2+data.get('_data_start',2), 'reason': '全选同一选项'})
            keep.remove(i)
            continue
    return keep, removed
```

---

## 动作6：缺失值处理（区分跳题缺失）

```python
SKIP_LOGIC = {
    'activity_type': {'dep': 'participated', 'trigger': ['是', '有参与', '参加']},
    # 键：被跳过的列名
    # dep：依赖列
    # trigger：依赖列取这些值时，本列应该非空
}

def handle_missing(data, strategy='mode'):
    N = data['N']
    to_drop = set()
    for key in data:
        if key in ('N', '_data_start') or key.startswith('_'):
            continue
        vals = data[key]
        none_idx = [i for i in range(N) if vals[i] is None or vals[i] == '']
        if not none_idx:
            continue
        # 跳题缺失
        if key in SKIP_LOGIC:
            rule = SKIP_LOGIC[key]
            dep_key = rule['dep']
            triggers = rule['trigger']
            real_missing = [i for i in none_idx if data[dep_key][i] in triggers]
            if not real_missing:
                continue
            none_idx = real_missing
        miss_rate = len(none_idx) / N
        if miss_rate > 0.30:
            to_drop.update(none_idx)
        elif strategy == 'mode':
            non_none = [v for v in vals if v not in (None, '')]
            fill = max(set(non_none), key=non_none.count) if non_none else 0
            for i in none_idx:
                vals[i] = fill
        elif strategy == 'flag':
            for i in none_idx:
                vals[i] = -1
    if to_drop:
        keep = [i for i in range(N) if i not in to_drop]
        return keep, list(to_drop)
    return list(range(N)), []
```

---

## 动作7：填空混合检测 + 自动路由

```python
import re

def classify_fill_column(data, key):
    """
    检测填空列属于哪种类型：
    - 'numeric': 全部可转数字 → float
    - 'mixed':   部分含文字但能提取数值 → 提取 + 标记
    - 'text':    以文本为主 → 词频
    """
    N = data['N']
    vals = [str(data[key][i] or '').strip() for i in range(N)]
    pure_num = 0
    has_num_in_text = 0
    pure_text = 0

    for v in vals:
        if not v:
            continue
        try:
            float(v)
            pure_num += 1
        except ValueError:
            nums = re.findall(r'\d+\.?\d*', v)
            if nums:
                has_num_in_text += 1
            else:
                pure_text += 1

    total = pure_num + has_num_in_text + pure_text
    if total == 0:
        return 'empty', {}
    if pure_num / total >= 0.90:
        return 'numeric', {}
    if (pure_num + has_num_in_text) / total >= 0.50:
        return 'mixed', {'extracted': has_num_in_text, 'total': total}
    return 'text', {}

def clean_fill_by_type(data, key):
    """根据分类结果执行对应清洗"""
    col_type, info = classify_fill_column(data, key)
    N = data['N']
    flags = []

    if col_type == 'numeric':
        data[key] = [float(str(v).strip()) if v is not None and str(v).strip() else None
                     for v in data[key]]

    elif col_type == 'mixed':
        new_vals = []
        for i in range(N):
            v = str(data[key][i] or '').strip()
            if not v:
                new_vals.append(None)
            else:
                try:
                    new_vals.append(float(v))
                except ValueError:
                    nums = re.findall(r'\d+\.?\d*', v)
                    if nums:
                        new_vals.append(float(nums[0]))
                        flags.append((i, f'{key}: "{v}" → 提取数值 {nums[0]}'))
                    else:
                        new_vals.append(None)  # 无法提取，置 None 避免下游炸掉
                        flags.append((i, f'{key}: "{v}" 无数字可提取 → 已置空'))
        data[key] = new_vals

    elif col_type == 'text':
        data[key] = [str(v).strip() if v is not None and str(v).strip() else None
                     for v in data[key]]

    return col_type, flags
```

---

## 动作8：主观题

```python
def handle_open_ended(data, key):
    """主观题：不修改数据，只统计并导出原文"""
    N = data['N']
    vals = [str(data[key][i] or '').strip() for i in range(N)]
    responded = [v for v in vals if v]
    lengths = [len(v) for v in responded]
    return {
        'key': key,
        '应答率': len(responded) / N,
        '平均字数': sum(lengths) / len(lengths) if lengths else 0,
        '响应文本': responded,
    }
```

---

## 动作9：多选题格式兼容

```python
def normalize_multi_select(data, col_key):
    """检测 0/1 多列 还是 逗号分隔单列，统一为 0/1 列"""
    N = data['N']
    vals = [str(data[col_key][i] or '') for i in range(N)]
    sample = [v for v in vals if v and v not in ('0','1','None','')]
    if sample and any(',' in v for v in sample):
        all_opts = set()
        for v in sample:
            all_opts.update(v.split(','))
        for opt in sorted(all_opts):
            new_key = f'{col_key}_{opt}'
            data[new_key] = [1 if opt in str(data[col_key][i] or '') else 0 for i in range(N)]
        return True
    return False
```

---

## 动作10：矩阵/量表列组映射

```python
COLUMN_GROUPS = {
    'aware_group': ['Q8_1', 'Q8_2', 'Q8_3', 'Q8_4'],
}

def load_matrix_columns(raw_data, groups):
    for group_name, col_names in groups.items():
        for i, cn in enumerate(col_names, 1):
            raw_data[f'{group_name}_{i}'] = raw_data[cn]
    return raw_data
```

---

## 辅助函数

```python
def filter_by_index(data, keep_idx):
    new = {'N': len(keep_idx)}
    for k, v in data.items():
        if k == 'N': continue
        new[k] = [v[i] for i in keep_idx]
    if '_data_start' in data:
        new['_data_start'] = data['_data_start']
    return new

def cleaning_report(raw_N, keep_idx, removed, fill_flags, open_stats):
    return {
        '原始样本': raw_N,
        '有效样本': len(keep_idx),
        '剔除数量': raw_N - len(keep_idx),
        '剔除率': f'{(raw_N-len(keep_idx))/raw_N:.1%}' if raw_N else '0%',
        '剔除明细': removed,
        '填空提取标记': fill_flags,
        '主观题摘要': open_stats,
    }
```

---

## 完整管线

```python
# 题型配置（唯一需改的）
QUESTION_TYPES = {
    'gender':   ('single',  {}),
    'grade':    ('single',  {}),
    'aware_1':  ('likert',  {'min':1,'max':5}),
    'participated': ('yesno', {}),
    'income':   ('fill',    {}),           # 自动检测数字/混合/文本
    'suggestion': ('open',  {}),           # 主观题：不处理，直接导出
}

# ── 执行 ──
wb = openpyxl.load_workbook('data/raw/问卷星导出.xlsx')
ws = wb.active
header_row, data_start = detect_header_rows(ws)
columns = strip_meta_columns(ws, header_row)

raw_N = ws.max_row - data_start + 1
raw = {'N': raw_N, '_data_start': data_start}
for col_idx, name in columns.items():
    raw[name] = [ws.cell(row=r, column=col_idx).value for r in range(data_start, ws.max_row+1)]

# 清洗
keep, removed = remove_invalid(raw)
raw = filter_by_index(raw, keep)
raw = normalize(raw, ENCODING_RULES)

all_flags = []
open_stats = []
for key, (qtype, params) in QUESTION_TYPES.items():
    if key not in raw:
        continue
    if qtype == 'yesno':
        normalize_yes_no(raw, key)
    elif qtype == 'fill':
        col_type, flags = clean_fill_by_type(raw, key)
        all_flags.extend(flags)
    elif qtype == 'open':
        open_stats.append(handle_open_ended(raw, key))

keep2, dropped = handle_missing(raw, strategy='mode')
raw = filter_by_index(raw, keep2)

# 报告
report = cleaning_report(raw_N, keep, removed+dropped, all_flags, open_stats)
print(f"原始 {report['原始样本']} → 有效 {report['有效样本']}（剔除 {report['剔除率']}）")
if all_flags:
    print(f"\n⚠️ {len(all_flags)} 处数值提取需确认：")
    for f in all_flags[:10]:
        print(f"  {f}")
if open_stats:
    for os_ in open_stats:
        print(f"\n主观题 [{os_['key']}]: 应答率 {os_['应答率']:.0%}, 均长 {os_['平均字数']:.0f} 字")

# 保存干净的 Excel
out_wb = openpyxl.Workbook()
out_ws = out_wb.active
for c, key in enumerate(raw.keys(), 1):
    out_ws.cell(row=1, column=c, value=key)
    for r in range(raw['N']):
        out_ws.cell(row=r+2, column=c, value=raw[key][r])
out_wb.save('data/processed/问卷数据_320份.xlsx')
```

---

## 策略分层

| 动作 | 策略 | 原因 |
|---|---|---|
| 全选同一选项 | 剔除 | 不可恢复 |
| 编码不统一 | 归一化 | 可逆 |
| 跳题缺失 | 不处理 | 合法空值 |
| 真缺失 <10% | mode填补 | 丢样本可惜 |
| 真缺失 >30% | 删行 | 填补引入偏差 |
| 纯数值填空 | float转换 | 直接统计 |
| 混合填空 | 提取数字 + 标记 | 保留数值，提醒审核 |
| 纯文本填空 | 词频 | 分类分析 |
| 主观题 | 原文导出 | 人工编码 |
