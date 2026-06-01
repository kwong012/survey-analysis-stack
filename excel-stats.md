# 问卷数据统计分析

## 工具选择策略

1. **Excel 原生功能** → 直接执行（需先确认分析工具库已启用）
2. **纯 Python stdlib** → 直接用
3. **需要 numpy/scipy** → 询问用户是否安装，等决策

---

## Excel 分析工具库：首次启用

文件 → 选项 → 加载项 → 管理"Excel 加载项" → 转到 → 勾选"分析工具库" → 确定

未启用时以下方法不可用：描述统计、t 检验、ANOVA、回归。

---

## 第一类：纯 Python stdlib（直接可用）

### 频率统计
```python
def freq(vals):
    d = {}
    for v in vals:
        d[v] = d.get(v, 0) + 1
    total = len(vals)
    return [(k, v, v/total) for k, v in sorted(d.items(), key=lambda x: -x[1])]
```

### 描述统计
```python
from statistics import mean, median, stdev, mode
def describe(col):
    clean = [x for x in col if x is not None]
    return {'n':len(clean),'mean':mean(clean),'median':median(clean),
            'stdev':stdev(clean),'min':min(clean),'max':max(clean),'mode':mode(clean)}
```

### Likert 量表
```python
def likert_analysis(col):
    n = len(col)
    return {'top2': sum(1 for x in col if x in (4,5))/n,
            'bottom2': sum(1 for x in col if x in (1,2))/n,
            'net': (sum(1 for x in col if x in (4,5))-sum(1 for x in col if x in (1,2)))/n}
```

### 多选题
```python
def multi_select(columns, labels):
    n = len(columns[0])
    return [(labels[i], sum(col), sum(col)/n) for i, col in enumerate(columns)]
```

### 交叉分组
```python
def group_rate(group_col, condition_fn):
    result = {}
    for g in set(group_col):
        idxs = [i for i, grp in enumerate(group_col) if grp == g]
        if idxs: result[g] = sum(1 for i in idxs if condition_fn(i))/len(idxs)
    return result
```

### 卡方检验（含 p 值近似）
```python
def chi_square(crosstab, row_labels, col_labels):
    row_tot = [sum(crosstab.get(r,{}).get(c,0) for c in col_labels) for r in row_labels]
    col_tot = [sum(crosstab.get(r,{}).get(c,0) for r in row_labels) for c in col_labels]
    total = sum(row_tot)
    chi2 = 0.0
    for i,r in enumerate(row_labels):
        for j,c in enumerate(col_labels):
            o = crosstab.get(r,{}).get(c,0)
            e = row_tot[i]*col_tot[j]/total if total>0 else 1
            if e>0: chi2 += (o-e)**2/e
    df = (len(row_labels)-1)*(len(col_labels)-1)
    return chi2, df

def chi_sq_pvalue(chi2, df, alpha=0.05):
    """近似 p 值。alpha=0.05 或 0.01
    临界值表：https://www.itl.nist.gov/div898/handbook/eda/section3/eda3674.htm"""
    table = {
        0.05: {1:3.84,2:5.99,3:7.81,4:9.49,5:11.07,6:12.59,7:14.07,8:15.51,9:16.92,10:18.31},
        0.01: {1:6.63,2:9.21,3:11.34,4:13.28,5:15.09,6:16.81,7:18.48,8:20.09,9:21.67,10:23.21},
    }
    t = table.get(alpha, {})
    if df in t:
        return f'<{alpha}' if chi2 > t[df] else f'>={alpha}'
    # df > 10：近似公式 sqrt(2*chi2) - sqrt(2*df-1) ≈ N(0,1)
    z = (2*chi2)**0.5 - (2*df-1)**0.5
    if abs(z) < 1.96:
        return f'>={alpha}（z={z:.2f}，近似正态）'
    return f'<{alpha}（z={z:.2f}，近似正态）'
```

数据格式要求：卡方检验需要**计数交叉表**（如 2×2 的城乡×参与），不能传原始清单。

### Pearson 相关系数
```python
def pearson_r(x, y):
    n = len(x); mx,my = mean(x),mean(y)
    sx,sy = stdev(x),stdev(y)
    cov = sum((x[i]-mx)*(y[i]-my) for i in range(n))/(n-1)
    return cov/(sx*sy) if sx and sy else 0
```
数据要求：两个等长连续变量。

### 独立样本 t 检验
```python
def ttest_ind(group_a, group_b):
    """两独立样本 t 检验。返回 (t 值, p 值近似)"""
    ma,mb = mean(group_a),mean(group_b)
    na,nb = len(group_a),len(group_b)
    if na < 2 or nb < 2:
        return None, '样本量不足'
    va,vb = stdev(group_a)**2,stdev(group_b)**2
    se = ((va/na)+(vb/nb))**0.5
    if se == 0:
        return None, '标准误为 0（数据异常）'
    t = (ma-mb)/se
    # df 近似（Welch-Satterthwaite）
    df_num = ((va/na)+(vb/nb))**2
    df_den = ((va/na)**2/(na-1)) + ((vb/nb)**2/(nb-1))
    df = df_num/df_den if df_den else 1
    # df >= 60 时 t ≈ N(0,1)，|t| > 1.96 ↔ p < 0.05
    if df >= 60:
        p_approx = '<0.05' if abs(t) > 1.96 else '>=0.05'
    else:
        p_approx = f'需查 t 表（df={df:.0f}）'
    return (t, p_approx)
```
数据要求：两组独立、数值型、样本量 ≥2。

### Cronbach's α
```python
def cronbach_alpha(items):
    k = len(items); n = len(items[0])
    var_total = stdev([sum(items[j][i] for j in range(k)) for i in range(n)])**2
    var_items = sum(stdev(col)**2 for col in items)
    return (k/(k-1))*(1-var_items/var_total) if var_total else 0
```
数据要求：k 个等长的 Likert 列。

---

## 第二类：Excel 原生功能（先确认工具库已启用）

| 方法 | Excel 操作 | 数据要求 |
|---|---|---|
| 描述统计 | 数据分析 → 描述统计 | 单列数值 |
| t 检验 | 数据分析 → t 检验：双样本等方差假设 | 两组数值分两列 |
| ANOVA 单因素 | 数据分析 → 方差分析：单因素 | 多组数值按列排布 |
| 回归 | 数据分析 → 回归 | Y 列 + X 列（可多列） |
| 相关系数 | `=CORREL()` | 两个等长区域 |
| 卡方 p 值 | `=CHISQ.TEST(实际, 期望)` | 两个等大区域 |

---

## 第三类：需要 numpy/scipy（先询问用户）

| 方法 | 所需库 | 安装命令 |
|---|---|---|
| ANOVA 多因素 | `statsmodels` | `pip install statsmodels` |
| 多元回归 | `statsmodels` | 同上 |
| 因子分析 | `sklearn` | `pip install scikit-learn` |
| Mann-Whitney U | `scipy` | `pip install scipy` |
| Kruskal-Wallis | `scipy` | 同上 |

**流程**：先问"需要安装 xxx，确认？"，同意后再装。若编译失败，加 `--only-binary=:all:`。

---

## 不适用
- 需要 SPSS 格式输出（`.sav`/`.spv`）— 导出 CSV 进 SPSS
- 数据 >10 万行 — 需 pandas 分块
