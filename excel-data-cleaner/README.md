# excel-data-cleaner

> **通用 Excel / CSV 数据清洗 Skill** —— 把脏表洗成"字段统一、空值归一、样式统一、异常可追溯"的 `.xlsx`。

<p align="left">
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white" alt="Python 3.10+" /></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-green" alt="License: MIT" /></a>
  <a href="SKILL.md"><img src="https://img.shields.io/badge/Skill-TRAE%20%2F%20Claude%20Agent-purple" alt="Skill Format" /></a>
  <img src="https://img.shields.io/badge/version-v13-orange" alt="v13" />
</p>

`excel-data-cleaner` 是一个面向 **明细表**（订单 / 会员 / 商品 / 物流等含唯一性主键的表格）的数据清洗 Skill。它把脏数据按统一规则归一字段格式、空值、字符串编码、异常值，并应用统一样式（微软雅黑 9 号 + 黑色细边框 + 表头加粗浅蓝底），最终产出一个"开箱即用"的 `.xlsx` 文件。

> **看板 / 汇总报表**（无主键、含合计 / 汇总 / 总计等聚合指标）请改用 [`excel-style-cleaner`](#)。

---

## 目录

- [为什么需要这个 Skill](#为什么需要这个-skill)
- [特性一览](#特性一览)
- [快速上手](#快速上手)
- [使用场景](#使用场景)
- [配置规则](#配置规则)
- [Pipeline 概览](#pipeline-概览)
- [输出说明](#输出说明)
- [项目结构](#项目结构)
- [依赖](#依赖)
- [扩展与定制](#扩展与定制)
- [版本历史](#版本历史)
- [许可](#许可)

---

## 为什么需要这个 Skill

数据运营 / 业务运营日常会面对大量"看上去像 Excel 但内容是灾难"的明细表：

| 问题 | 真实场景 |
| --- | --- |
| **字段格式混乱** | 同一列"金额"里既有 `1234.50`，也有 `1,234.5`，还有 `¥1,234.50` |
| **时间格式不统一** | `2026/01/15`、`2026-01-15`、`15/01/2026`、`01:30:00 PM` 混在一起 |
| **空值字面量爆炸** | 空单元格 / `null` / `NULL` / `N/A` / `-` / `—` / `#N/A` 各自代表"没值" |
| **异常静默丢失** | 主键重复、ID 缺失、邮箱非法、手机号位数错，永远要人工对账才能发现 |
| **样式千差万别** | 标题字体混杂、列宽乱、没有边框，打开就头大 |
| **AI 生成水印** | 部分来源表带"含 AI 生成"水印层（WPS / Excel 残留），需要剥离 |

`excel-data-cleaner` 用 **一份 YAML 配置 + 一条命令** 把这些全部归一。

---

## 特性一览

| 类别 | 能力 |
| --- | --- |
| **多格式读入** | `.csv` / `.xls` / `.xlsx` 自动识别，统一读入 |
| **字段识别** | ID / 日期时间 / 月份 / 长文本 / 金额 / 百分比 / 数值 / 手机号 / 邮箱 / 布尔，按列名关键词启发式分类 |
| **空值归一** | `""` / `" "` / `null` / `N/A` / `-` / `—` / `#N/A` 等 15+ 字面量统一为空 |
| **字符串归一** | Unicode NFKC 全角→半角、tab/不间断空格替换、首尾 strip（文本字段保留段落内空白） |
| **时间统一** | `YYYY-MM-DD` 24 小时制；自动识别 12h/24h、`/`/`-` 分隔符 |
| **异常标识** | `DUP_ID` / `MISSING_ID` / `DUP_ROW` / `BAD_DATE` / `BAD_AMOUNT` / `BAD_EMAIL` / `BAD_PHONE` / `TEXT_TOO_LONG` / `OUT_OF_RANGE` / `STAT_OUTLIER` 10 种异常，写入 `_ANOMALIES` sheet |
| **统一样式** | 微软雅黑 9 号 + 黑色细边框 + 表头加粗浅蓝底 + 列宽 / 行高自适应 |
| **表头合并** | 多级表头自动识别 + 顶层横向合并 + 相邻空合并 + A 列纵向合并 |
| **AI 水印清理** | 自动剥离 WPS / Excel 残留的 `xl/drawings/*.xml` 装饰层（含"含 AI 生成"文本框） |
| **大数据量优化** | `engine="auto"` 自动切换 `xlsxwriter` 后端（≥50 万 cells 时提速 5-10 倍） |
| **可扩展** | 全部规则集中在 `config/cleaning_rules.yaml`，可按业务加字段 / 异常 / 样式 |

---

## 快速上手

### 1. 安装依赖

```bash
pip install pandas openpyxl xlrd pyyaml xlsxwriter
```

> Python ≥ 3.10

### 2. 命令行（最常用）

```bash
# 默认清洗 .csv / .xls / .xlsx → 输出 .xlsx
python scripts/clean_excel.py \
    --input examples/sample_dirty.csv \
    --output examples/sample_cleaned.xlsx \
    --rules config/cleaning_rules.yaml

# 大文件自动走 xlsxwriter（≥50 万 cells）
python scripts/clean_excel.py \
    --input huge_file.xlsx \
    --output huge_file_cleaned.xlsx

# 强制走 xlsxwriter
python scripts/clean_excel.py \
    --input huge_file.xlsx \
    --output huge_file_cleaned.xlsx \
    --engine xlsxwriter
```

### 3. Python API

```python
from scripts.clean_excel import clean_file

result = clean_file(
    input_path="path/to/dirty.xlsx",
    output_path="path/to/cleaned.xlsx",
    rules_path="config/cleaning_rules.yaml",   # 可选,缺省走内置默认规则
    # engine="auto",                          # auto / openpyxl / xlsxwriter
)

print(result["summary"])        # 命中规则数 / 空值数 / 异常数
print(result["input_sheets"])   # 输入 sheet 列表
```

### 4. 在 AI Agent（TRAE / Claude Agent 等）中调用

把本仓库放到 Agent 能识别的 `skills/` 目录下，然后用自然语言触发：

> "帮我清洗一下 `examples/sample_dirty.csv` 这个文件"

Agent 会自动加载 [`SKILL.md`](SKILL.md) 并按 pipeline 执行。

---

## 使用场景

### ✅ 适用

- 订单明细、会员明细、商品明细、物流明细、付款明细等 **明细表**
- 表格中存在 **唯一性字段主键**（订单号 / 会员号 / 商品 ID / 编号 / 工号 / 流水号）
- 表格中 **不存在**"合计 / 汇总 / 总计"等聚合字样
- 输入是 `.csv` / `.xls` / `.xlsx`
- 需要把多份格式不一的表归一为同一份规范的 `.xlsx`

### ❌ 不适用（请改用 `excel-style-cleaner`）

- 销售看板、月度业绩看板、经营分析报表等 **汇总类报表**
- 表格中含 **合计 / 汇总 / 总计** 等聚合指标
- 表格 **没有唯一性主键**，存在多级合并表头与分类汇总行
- 需要为达成率列加黄色 Data Bar、为环比 / 同比列加红绿箭头进度条

---

## 配置规则

默认规则已覆盖 80% 场景，所有可调阈值集中在 [`config/cleaning_rules.yaml`](config/cleaning_rules.yaml)。

### 字段识别配置

| 字段类型 | 识别关键词（部分） | 输出格式 | 显示格式 |
| --- | --- | --- | --- |
| `id`（主键/编号） | `id` / `编号` / `工号` / `订单号` / `...号` | 字符串 | `@` |
| `datetime` | `date` / `time` / `日期` / `时间` | `YYYY-MM-DD` | `yyyy-mm-dd` |
| `month` | `付款月份` / `交易月份` / `month` | `YYYY-MM` | `yyyy-mm` |
| `long_text` | `备注` / `说明` / `address` / `地址` | 字符串（保留段落内空白） | `@` |
| `percent` | `费率` / `占比` / `折扣` / `percentage` | 0~1 小数 | `0.00%` |
| `amount` | `amount` / `金额` / `gmv` / `手续费` / `余额` | 数值 | `0.00` |
| `number` | `count` / `qty` / `数量` / `年龄` | 数值 | `0.00` |
| `phone` | `phone` / `手机` | 11 位数字 | `@` |
| `email` | `email` / `邮箱` | 小写字符串 | `@` |
| `boolean` | `is_` / `flag` / `是否` | `True` / `False` | `@` |

### 异常类型与触发条件

| 异常 | 触发条件 |
| --- | --- |
| `MISSING_ID` | 主键列空 |
| `DUP_ID` | 主键出现 ≥ 2 次 |
| `DUP_ROW` | 整行所有列完全相同 |
| `BAD_DATE` | 日期 < 1970-01-01 或 > today+1 |
| `BAD_MONTH` | 月份无法解析为 `YYYY-MM` |
| `BAD_AMOUNT` | 金额含非数字字符 |
| `BAD_EMAIL` | 不匹配 RFC 5322 简化正则 |
| `BAD_PHONE` | 非 11 位数字（中国大陆） |
| `TEXT_TOO_LONG` | 长文本 > 500 字符 |
| `OUT_OF_RANGE` | 数值绝对值 > 1e12 |
| `STAT_OUTLIER` | 数值 Mean ± 2σ 之外（样本 ≥ 10） |

### 自定义示例

```yaml
# config/cleaning_rules.yaml
fields:
  id:
    column_keywords: [id, 编号, no, code, 工号, 订单号, 票号, 号]
    cell_format: "@"
    is_primary_key: true
  datetime:
    column_keywords: [datetime, date, time, 日期, 时间, timestamp]
    cell_format: "yyyy-mm-dd"
    force_24h: true
  amount:
    column_keywords: [amount, money, price, 金额, 价格, fee, gmv, 余额]
    cell_format: "0.00"
    allow_negative: true

anomalies:
  out_of_range_abs: 1e12
  stat_outlier_std_multiplier: 2
  stat_outlier_min_samples: 10

output_style:
  font_name: "微软雅黑"
  font_size: 9
  font_color: "FF000000"
  border_color: "FF000000"
```

---

## Pipeline 概览

整个清洗流程划分为 **四个阶段**，每个阶段职责单一、产出明确：

```
┌──────────────┐    ┌──────────────────┐    ┌──────────────────┐    ┌──────────────┐
│   阶段一     │    │      阶段二      │    │      阶段三      │    │   阶段四     │
│  Input ·     │ →  │  Format Clean ·  │ →  │ Visual Clean ·   │ →  │  Output ·    │
│  多格式读入  │    │    格式清洗      │    │    可视化清洗    │    │  统一输出    │
└──────────────┘    └──────────────────┘    └──────────────────┘    └──────────────┘
   • .csv/.xls/      ① 字段格式清洗         ① 统一样式               • 写出 .xlsx
     .xlsx 读入       ② 空值归一              · 有效范围识别
   • sheet 选择      ③ 时间统一              · 字体 / 字号
   • header=None     ④ 字符串归一            · 边框 / 颜色
   • 表头识别        ⑤ 异常标识              · 对齐 / 换行
                       · 主键校验          ② 行高 / 列宽自适应
                       · 类型校验          ③ 表头样式
                                              · 多级表头识别
                                              · 表头加粗
                                              · 表头背景填充
```

详见 [`SKILL.md`](SKILL.md) → "Pipeline Overview"。

---

## 输出说明

`<name>_cleaned.xlsx` 包含 **3 个 sheet**：

| Sheet | 内容 |
| --- | --- |
| **原始数据 sheet**（保留原 sheet 名） | 清洗后的明细数据 |
| **`_CLEAN_REPORT`** | 每列识别到的字段类型、命中规则数、空值数 |
| **`_ANOMALIES`** | 异常明细（`sheet_name` / `row` / `column` / `original_value` / `anomaly_type` / `reason`） |

控制台同时输出 `summary`：

```
{
  "input_sheets": ["Sheet1"],
  "matched_rules": {"id": 6, "datetime": 4, "amount": 31, ...},
  "null_normalized": 142,
  "anomaly_count": 18,
  "output_path": "examples/sample_cleaned.xlsx"
}
```

---

## 项目结构

```
excel-data-cleaner/
├── README.md                     # 本文件
├── SKILL.md                      # Skill 主文档（Agent 入口）
├── LICENSE                      # MIT License
├── config/
│   └── cleaning_rules.yaml       # 字段类型 / 异常阈值 / 样式配置
├── scripts/
│   └── clean_excel.py            # 主清洗脚本（CLI + Python API）
├── references/
│   └── field_patterns.md         # 字段识别正则与边界说明
├── tests/
│   ├── test_header_detection.py  # 表头识别单元测试
│   ├── test_data_rect.py         # 有效范围端到端测试
│   ├── test_effective_range.py   # 有效行 / 有效列 / 行高测试
│   ├── test_header_style_in_same_loop.py
│   ├── test_merge_excel_style_cleaner.py
│   └── test_top_level_merge.py
└── examples/
    ├── sample_dirty.csv          # 脏数据样例（输入）
    └── sample_cleaned.xlsx       # 清洗后样例（输出）
```

---

## 依赖

```bash
pip install pandas openpyxl xlrd pyyaml xlsxwriter
```

- **pandas** ≥ 1.3
- **openpyxl** ≥ 3.0（读 .xlsx + 默认写出引擎）
- **xlrd** ≥ 2.0（读 .xls）
- **xlsxwriter** ≥ 3.0（大文件后端，可选但推荐）
- **pyyaml** ≥ 6.0（规则文件）

---

## 扩展与定制

| 需求 | 改哪里 |
| --- | --- |
| **新增字段类型** | `config/cleaning_rules.yaml` 加 `fields.<type>` 块 + `scripts/clean_excel.py` 的 `_clean_cell` 加分支 |
| **新增异常类型** | `scripts/clean_excel.py` 的 `_collect_anomalies` 加扫描逻辑 + `SKILL.md` 异常表登记 |
| **自定义空值字面量** | `config/cleaning_rules.yaml` → `empty_values.literals` |
| **修改字体 / 字号 / 边框** | `config/cleaning_rules.yaml` → `output_style` |
| **强制走 xlsxwriter** | 调用时传 `engine="xlsxwriter"` |
| **保留 VBA 宏** | 当前 v13 **不支持** `.xlsm`，请改用其他工具或单独处理 |

---

## 版本历史

- **v13（2026-09-15）** — 删除旧水印逻辑 `_strip_ai_watermarks`；字段识别表全量同步 v10/v11/v12 关键词
- **v12（2026-09-14）** — `amount` 关键词追加"余额 / balance"；修 BJ 列 0.00 不生效
- **v11（2026-09-14）** — `amount` 关键词扩展（手续费 / 调整金额 / 成本 / 利润）；修 AV / BR 列；`_strip_ai_drawing_layers` 重构（不再静默吞异常）
- **v10（2026-09-13）** — `_write_xlsx_fast` 表头合并算法重写（抄 `excel-style-cleaner` v3.22+）；单元格 text_wrap / 居中
- **v6（2026-09-12）** — A1 合并范围扩展修复；`drawing` 清理 Permission 修复
- **v5（2026-09-12）** — 表头居中；`_strip_ai_drawing_layers` 新增（清理 AI 水印层）
- **v4（2026-09-11）** — 标题合并为空修复；V 列样式修复；列宽自适应 + 居中
- **v3（2026-09-11）** — 表头合并；百分比格式；数值还原；`xlsxwriter` 后端
- **v2（2026-09-10）** — 时间统一 24h 制；ID / 编号识别增强；统一样式
- **v1（2026-08-28）** — 初始版本

完整修订记录见 [`logs/log.md`](logs/log.md)。

---

## 许可

[MIT License](LICENSE) — 自由使用、修改、分发，请保留版权声明。

---

## 微信公众号

使用中遇到 bug 或有疑问，可扫描下方二维码在公众号后台留言，作者会定期回复：

![微信公众号二维码](assets/wechat-qr.jpg)