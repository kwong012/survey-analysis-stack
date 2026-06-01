# 项目目录规范

## 推荐结构

```
project/
├── data/
│   ├── raw/                ← 原始数据，只读不写（问卷星导出原样存放）
│   │   └── 2026-01-10-问卷星导出.xlsx
│   └── processed/          ← 清洗后数据（脚本从这里读取）
│       └── 问卷数据_320份.xlsx
│
├── scripts/
│   ├── config.py           ← COLUMN_MAP + THRESHOLDS（唯一需改的）
│   ├── load.py             ← load_data() + validate() + clean()
│   ├── stats.py            ← 按维度拆分的小函数
│   ├── render.py           ← 渲染 Word 报告
│   └── main.py             ← 一键跑全流程
│
├── output/                 ← 脚本产出（可删除，.gitignore 排除）
│   ├── 统计分析结果.xlsx
│   ├── 分析图表.xlsx
│   └── 报告.docx
│
├── docs/                   ← 文档
│   ├── 变量说明.md          ← 每个变量的含义、取值范围
│   ├── 方法记录.md          ← 为什么用卡方/t检验/这个阈值
│   └── 问卷编码表.xlsx      ← 题号→变量名→题型对照
│
├── archive/                ← 旧版报告归档
│   └── 2026-01-15-初版/
│
├── .gitignore
└── README.md
```

---

## .gitignore 模板

```gitignore
# 产出（可重跑恢复）
output/

# 数据（原始 + 处理过都可能涉隐私，全排除）
data/

# Python
__pycache__/
*.pyc

# 系统
.DS_Store
Thumbs.db
```

> ⚠️ `data/` 不应进版本控制，原始问卷可能含用户隐私。
> README 注明：找 4 号数据采集组要原始数据，放 `data/raw/`，再跑清洗脚本生成 `data/processed/`。

---

## 各目录规则

| 目录 | 规则 |
|---|---|
| `data/raw/` | 只读。问卷星导出原样存放，永不修改 |
| `data/processed/` | 脚本从这里读取。由清洗脚本从 raw/ 生成 |
| `scripts/` | 只放 .py。`main.py` 是唯一入口 |
| `output/` | 可删除。.gitignore 排除，不进版本控制 |
| `docs/` | 变量说明、方法记录、编码表 |
| `archive/` | 按日期归档旧版，不自动覆盖 |

---

## 命名规范

| 文件 | 命名 |
|---|---|
| 脚本 | 小写_下划线：`compute_stats.py` |
| 产出 | 中文描述：`统计分析结果.xlsx` |
| 数据 | `YYYY-MM-DD-描述.xlsx`：`2026-01-10-问卷星导出.xlsx` |
| 归档 | `YYYY-MM-DD-版本/`：`2026-01-15-定稿版/` |

---

## README 模板

```markdown
# 项目名

## 快速开始
1. 把清洗后数据放到 `data/processed/`
2. `python scripts/main.py`
3. 产出在 `output/`

## 依赖
openpyxl, xlsxwriter, python-docx, lxml
安装：`pip install openpyxl xlsxwriter python-docx lxml`

## 配置
问卷列号映射 → `scripts/config.py` 的 `COLUMN_MAP`
阈值调整 → `scripts/config.py` 的 `THRESHOLDS`
```

---

## 适用

- 问卷调研报告项目
- 数据迭代、多人协作
- 同课题反复调研
