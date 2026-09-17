***

name: excel-data-cleaner
updated: 2026-09-15v13
description: >-
通用 Excel 数据清洗 Skill。使用场景:明细表(存在订单号 / 会员号 / ID 等唯一性字段主键、无合计 / 汇总 / 总计等聚合指标)。
看板 / 汇总报表场景请改用 excel-style-cleaner。处理 .csv/.xls/.xlsx,字段归一、空值归一、异常标识、样式归一(微软雅黑 9 号 + 黑色边框),输出 .xlsx。
Invoke when 用户要求清洗 Excel/CSV、字段标准化、空值归一、异常标识、统一样式。
中文触发: Excel 清洗、数据清洗、表格清洗、字段标准化、空值归一、异常标识、统一样式。
----------------------------------------------

# Excel Data Cleaner

通用 Excel 数据清洗 Skill。处理 `.csv` / `.xls` / `.xlsx` 三类表格,按统一的字段格式规则清洗后,输出带统一样式(微软雅黑 9 号字体 + 黑色细边框)的 `.xlsx` 文件,并附带清洗报告与异常标识 sheet。

## When to Invoke

> **使用场景定位**：本 skill 适用于 **明细表** —— 表格中存在一列能唯一标识每行记录的字段（订单号、会员号、ID、编号等），不存在"合计 / 汇总 / 总计"等聚合字样。
> 若表格为看板 / 汇总报表（无唯一性主键、含合计 / 汇总 / 总计 等聚合指标），请改用 `excel-style-cleaner` skill。

- 用户提到"清洗 Excel"、"清洗表格"、"清洗 CSV"、"规范字段格式"、"统一空值"

- 表格类型属于 **明细表**（订单明细 / 会员明细 / 商品明细 / 物流明细等）

- 表格中存在 **唯一性字段主键**（订单号 / 会员号 / 商品 ID / 编号 / 工号 / 流水号等）

- 表格中 **不存在**"合计 / 汇总 / 总计"等聚合字样

- 输入文件类型为 `.csv` / `.xls` / `.xlsx` 之一

- 需要把多种格式的表格归一为统一的 `.xlsx`

- 需要标记异常值并产出清洗报告

- 需要统一字体、字号、边框样式

> **本 Skill 不再支持** **`.xlsm`** **文件**(VBA 宏场景)。如需保留宏,请改用其他工具或单独处理。

### 不适用场景（请改用 `excel-style-cleaner`）

- 销售看板、月度业绩看板、经营分析报表等 **汇总类报表**

- 表格中含 **合计 / 汇总 / 总计** 等聚合指标

- 表格 **没有唯一性主键**，存在多级合并表头与分类汇总行

- 需要为达成率列加黄色 Data Bar、为环比 / 同比列加红绿箭头进度条

- 需要把金额列以万级格式（`0.0"万"`）展示

## Pipeline Overview(清洗流程架构)

整个清洗流程划分为 **四个阶段**,每个阶段职责单一、产出明确,可独立测试与复用:

```
┌──────────────┐    ┌──────────────────┐    ┌──────────────────┐    ┌──────────────┐
│   阶段一     │    │      阶段二      │    │      阶段三      │    │   阶段四     │
│  Input ·     │ →  │  Format Clean ·  │ →  │ Visual Clean ·   │ →  │  Output ·    │
│  多格式读入  │    │    格式清洗      │    │    可视化清洗    │    │  统一输出    │
└──────────────┘    └──────────────────┘    └──────────────────┘    └──────────────┘
   • .csv/.xls/      ① 字段格式清洗         ① 统一样式               • 写出 .xlsx
     .xlsx 读入       ② 空值归一              · 有效范围识别          • 附带 _CLEAN_REPORT
   • sheet 选择      ③ 时间统一              · 字体 / 字号             和 _ANOMALIES
   • header=None     ④ 字符串归一            · 边框 / 颜色              两个 sheet
   • 表头识别        ⑤ 异常标识              · 对齐 / 换行
                       · 主键校验            · 条件格式(若需)
                       · 类型校验          ② 行高 / 列宽自适应
                                          ③ 表头样式
                                              · 多级表头识别
                                              · 表头加粗
                                              · 表头背景填充
```

### 阶段对照表

| 阶段          | 关键产物                                        | 对应函数                                                                                                                                                                                            | <br /> |
| ----------- | ------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------ |
| 阶段一 · 多格式读入 | `{sheet: DataFrame}` + 整数列名 + `header_rows` | `read_input`、`detect_header_rows`、`_detect_header_rows_from_rows`                                                                                                                               | <br /> |
| 阶段二 · 格式清洗  | 清洗后的 `DataFrame` + `anomalies` 列表 + 字段类型映射  | `clean_dataframe`、`_clean_cell`、`_normalize_string`、`_parse_datetime`、`_parse_number` 等                                                                                                         | <br /> |
| 阶段三 · 可视化清洗 | 带样式的 xlsx 文件                                | `_apply_styles`、`_find_effective_range`、`_merge_header_cells`、`_merge_header_col_cells`、`_merge_header_row`、`_auto_fit_columns`、`_auto_fit_row_heights`、`_apply_cell_formats`、`_get_cell_value` | <br /> |
| 阶段四 · 统一输出  | 最终 `.xlsx`(仅数据 sheet)                       | `write_xlsx`、`_apply_cell_formats`                                                                                                                                                              | <br /> |

## Stage 1 · 多格式读入 (Input)

**职责**:把任意支持的格式统一读入内存,产出可被阶段二处理的 `DataFrame`。

### 支持格式

| 后缀      | 读取方式                                                       | 引擎                |
| ------- | ---------------------------------------------------------- | ----------------- |
| `.csv`  | `pd.read_csv(header=None, dtype=str)`                      | pandas            |
| `.xls`  | `pd.read_excel(header=None, dtype=str, engine='xlrd')`     | pandas + xlrd     |
| `.xlsx` | `pd.read_excel(header=None, dtype=str, engine='openpyxl')` | pandas + openpyxl |
| `.xlsm` | ❌ **不支持**(本 Skill 不再处理保留宏)                                 | —                 |

> 输入后缀不在上表 → 抛 `ValueError`,提示支持的后缀列表。

### 关键约定

- **列名统一为整数 0..n-1**:`header=None` 保证原始数据(包括首行表头)都保留为内容,不丢行

- **dtype 全部为 str**:避免 pandas 自动把 ID 转数字、长数字串转科学计数法

- **CSV 单独处理**:`keep_default_na=False` 防止空字符串被识别为 NaN

- **表头行自动识别**(可被阶段三复用):详见阶段三 · 表头识别

- **多 sheet 处理**:`.xls` / `.xlsx` 返回 `{sheet_name: DataFrame}` 字典

## Stage 2 · 格式清洗 (Format Clean)

**职责**:对阶段一的原始 DataFrame 做内容层清洗,产出"语义正确"的 DataFrame。

### ②-① 字段格式清洗

| 字段类别             | 检测启发式                                                                                            | 输出格式                       | 单元格 number\_format    |
| ---------------- | ------------------------------------------------------------------------------------------------- | -------------------------- | --------------------- |
| ID(主键/编号)        | 列名含 `id` / `编号` / `no` / `code` / `key` / `uuid` / `工号` / `订单号` / `流水号` / `单号` / `票号` / `**号**` | 字符串                        | `@`(文本)               |
| 日期时间(DateTime)   | 列名含 `datetime` / `timestamp` / `date` / `time` / `日期` / `时间` / `_dt` / `_ts`                     | `YYYY-MM-DD`(24h)          | `yyyy-mm-dd`          |
| 月份(Month)        | 列名含 `month` / `付款月份` / `交易月份`                                                                  | `YYYY-MM`                  | `yyyy-mm`             |
| 长文本(LongText)    | 列名含 `desc` / `remark` / `备注` / `说明` / `comment` / `address` / `地址` / `note` / `notes`          | 字符串(只 strip 首尾,保留段落内空行/空格) | `@`                   |
| 百分比(Percent)     | 列名含 `费率` / `比例` / `占比` / `折算比例` / `折扣` / `费率折扣` / `使用比` / `gmv比` / `percentage` / `rate`        | 数值(0~1,默认校验范围)            | `0.00%`               |
| 金额(Amount/Money) | 列名含 `amount` / `money` / `price` / `金额` / `价格` / `fee` / `cost` / `salary` / `revenue` / `income` / `gmv` / `比例费金额` / `手续费` / `调整金额` / `成本` / `利润` / `余额` / `balance` | 数值(2 位小数,负数允许)             | `0.00`                |
| 数值(Number)       | 列名含 `count` / `qty` / `quantity` / `数量` / `age` / `score` / `ratio` / `月份数` / `次数`               | 数值(底层值不变,仅显示格式)            | `0.00`                |
| 手机号(Phone)       | 列名含 `phone` / `mobile` / `手机` / `tel`                                                            | 11 位数字字符串(去非数字字符)          | `@`                   |
| 邮箱(Email)        | 列名含 `email` / `邮箱` / `mail`                                                                      | 小写字符串                      | `@`                   |
| 布尔(Boolean)      | 列名含 `is_` / `flag` / `enabled` / `是否` / `active`                                                | `True` / `False`           | `@`                   |

> **ID/编号兜底规则**:任何列名只要含"号"字(如 `合同号`、`卡号`、`发票号`、`编号` 已经在列表里),都会自动归一为 ID/编号类。配合单元格格式 `@`,避免长 ID 被 Excel 自动转科学计数法。

> 若列名无法识别,标记为 `untyped`,但仍享受字符串归一(NFKC + trim)。

### ②-② 空值归一

以下值统一视为空值(在内存中用 `None` 表示,在输出 xlsx 中写入空单元格):

- 空字符串 `""`

- 纯空白 `" "`、`"  "`、制表符、换行、不间断空格 `\xa0`(经过 NFKC + trim 归一后)

- 大小写不敏感的字符串:`null`、`NULL`、`None`、`N/A`、`n/a`、`NA`、`-`、`—`、`#N/A`、`\\N`

- pandas 读入后的 `NaN` / `NaT`

### ②-③ 时间统一(强制 24h 制)

`datetime` 字段无论输入是什么格式,统一归一为 `YYYY-MM-DD HH:MM:SS`(24 小时制):

- `2026-01-15` → `2026-01-15 00:00:00`

- `2026/01/15 09:30:00` → `2026-01-15 09:30:00`

- `15/01/2026 09:30` → `2026-01-15 09:30:00`(按 DD/MM/YYYY 解析)

- `01/15/2026 09:30` → `2026-01-15 09:30:00`(按 MM/DD/YYYY 解析)

- `01:30:00 PM` → `13:30:00`(12h 转 24h)

- `09:30:00 AM` → `09:30:00`(12h 转 24h)

不出现 `AM` / `PM` 标记。

### ②-④ 字符串归一(作用于所有字段)

| 步骤 | 行为                           | 配置项                                   |
| -- | ---------------------------- | ------------------------------------- |
| 1  | Unicode NFKC(全角→半角、不间断空格→空格) | `unicode_nfkc: true`                  |
| 2  | tab / `\xa0` → 普通空格          | 默认开启                                  |
| 3  | 首尾 strip                     | `trim: true`                          |
| 4  | 文本字段不清除段落内空白                 | `strip_internal_whitespace: false`    |
| 5  | 不合并段落内连续空白                   | `collapse_internal_whitespace: false` |
| 6  | 邮箱统一小写                       | `lowercase_email: true`               |

> 长文本字段(如"备注"、"地址")**只 strip 首尾空白**,不清除段落内部的空行与空格。

### ②-⑤ 异常标识

| 异常类型      | 触发条件                        | 标识                |
| --------- | --------------------------- | ----------------- |
| ID 重复     | 主键列出现 2 次及以上                | `[DUP_ID]`        |
| ID 缺失     | 主键列空                        | `[MISSING_ID]`    |
| 全行重复      | 整行的所有列值与之前某行完全相同            | `[DUP_ROW]`       |
| 日期/时间超出范围 | 日期 < 1970-01-01 或 > today+1 | `[BAD_DATE]`      |
| 月份非法      | 月份字段无法解析为 YYYY-MM           | `[BAD_MONTH]`     |
| 金额非法      | 金额字段含非数字字符                  | `[BAD_AMOUNT]`    |
| 长文本超长     | 字段长度 > 500 字符               | `[TEXT_TOO_LONG]` |
| 邮箱非法      | 字段不匹配 RFC 5322 简化正则         | `[BAD_EMAIL]`     |
| 手机号非法     | 字段非 11 位数字(中国大陆)            | `[BAD_PHONE]`     |
| 数值越界      | 数值字段绝对值 > 1e12              | `[OUT_OF_RANGE]`  |
| 统计学异常值    | 数值列 Mean ± 2σ 之外(样本 ≥ 10)   | `[STAT_OUTLIER]`  |

异常不会删除原始数据,而是写入 `_ANOMALIES` sheet:

```
sheet_name | row | column | original_value | anomaly_type | reason
```

## Stage 3 · 可视化清洗 (Visual Clean)

**职责**:对阶段二清洗后的 DataFrame,**在样式层做统一化**,不修改底层数据。

### ③-① 统一样式

#### 有效范围识别(关键,参考 `excel-style-cleaner`)

**术语对齐** **`excel-style-cleaner`** **的"整行 / 整列"**:

| 术语          | 定义                               |
| ----------- | -------------------------------- |
| 整行(如 `1:1`) | 含所有列(从第 1 列到最后一列)的一行             |
| 整列(如 `A:A`) | 含所有行(从第 1 行到最后一行)的一列             |
| **有效行**     | 不属于"整行为空"的所有行号(升序)               |
| **有效列**     | 不属于"整列为空"的所有列号(升序)               |
| **有效范围**    | `[首有效行, 末有效行] × [首有效列, 末有效列]` 矩形 |

**有效范围识别算法**(对齐 `excel-style-cleaner`):

1. `_find_empty_full_rows(ws)` → 返回"整行为空"的行号集合(用 `_is_empty` 判定每个单元格:`None` / 空字符串 / 纯空白)
2. `_find_effective_rows(ws)` = 全行号 - 整行为空的行号 → 升序有效行列表
3. `_find_empty_full_cols(ws)` → 返回"整列为空"的列号集合
4. `_find_effective_cols(ws)` = 全列号 - 整列为空的列号 → 升序有效列列表
5. `_find_effective_range(ws)` = `(effective_rows[0], effective_rows[-1], effective_cols[0], effective_cols[-1])`

**有效范围 = 样式、字体、合并、换行、列宽 / 行高自适应的唯一应用范围**。不在有效范围内的单元格保持原样(无样式、不动边框)。

#### 字体 / 字号 / 颜色 / 对齐

| 属性 | 取值                           | 适用范围         |
| -- | ---------------------------- | ------------ |
| 字体 | **微软雅黑**                     | 有效范围内全部非空单元格 |
| 字号 | **9**                        | 同上           |
| 颜色 | **黑色** `#000000`             | 同上           |
| 对齐 | 水平 `left` / 垂直 `center`(表头行) | 有效范围内        |

#### 边框规则

- **有效范围内的所有单元格**(包括范围内但值为空的单元格)**都**画黑色细边框(thin)

- 范围外的单元格不画边框(保留原样)

### ③-② 行高 / 列宽自适应

#### 列宽自适应

- 范围:有效范围的全部列

- **中文字符按 2 宽度**,ASCII 按 1 估算内容宽度

- **列宽** = `max(min_width=8, content_width + 2 padding)`,上限 `max_width=40`

- **超出 max\_width** → 列宽固定为 40,该列单元格 `wrap_text=True`,Excel 自动换行

- 不修改底层值,仅样式层面调整

#### 行高自适应(新增)

- 范围:有效范围的全部行

- 单列在有效列宽下所需的**行数** ≈ `ceil(content_width / col_width)`,中文字符 2 / ASCII 1

- **该行的行数** = `max(各列所需行数)`,下限 1

- **行高(磅)** = `max(min_height=15, base_line_height=15 × max_lines)`,下限 15 磅(Excel 默认)

- 不修改底层值,仅样式层面调整

> **与** **`excel-style-cleaner`** **的差异**:该 skill 仅做列宽自适应,本 Skill **同时做行高自适应**;实现基于列宽换行后的行数估算。

#### 条件格式

- 默认**不应用**条件格式(Data Bar / 颜色进度条等)

- 本 Skill 不适用场景中的"占比 / 环比 / 同比 / 达成率"等聚合指标已交由 `excel-style-cleaner` 处理

### ③-③ 表头样式

#### 多级表头识别(确认范围 · 表头行 / 内容行)

- **表头行数** 由 `detect_header_rows` 自动识别

- **表头行** = 有效范围内前 N 行(N = 自动识别的表头行数)

- **内容行** = 有效范围内非表头的行

**表头行自动识别算法**(基于"同列连续两行相近"):

- 扫描前 `MAX_HEADER_ROW_SCAN = 10` 行,跳过完全为空的行和说明/备注类文字行

- 按列从左到右扫描,每列从上到下

- **若找到首个同一列连续两行都非空且内容"相近"**(都是日期/都是数字/都是会员号/都是邮箱/都是 URL/都是手机号等),则这两行之上的所有行识别为表头行

- "相近"判定包括以下八种类型之一即可:

  1. 都是日期(两个都是日期格式)
  2. 都是数字(int/float,非 bool)
  3. 都是纯数字字符串(`-?[\d,]+(\.\d+)?%?`)
  4. 都是"XXX号"样式(含 `号$` 或 `号-?\d`)
  5. 都是邮箱
  6. 都是 URL(`https?://` 开头)
  7. 都是手机号(11 位)
  8. 弱相似:字符串长度差 ≤ 2 且都含字母数字

- 若扫描结束未找到,默认返回 1(常见单行表头)

#### 表头加粗 + 背景填充

- **加粗**:`Font(bold=True)`,仅作用于有效范围内表头行的非空单元格

- **背景填充**:`PatternFill("#D9E1F2")` 浅蓝色,作用范围同上

- **对齐**:`Alignment(horizontal="left", vertical="center")`

#### 表头单元格自动合并(从 `excel-style-cleaner` 复制)

##### 规则 A · 顶层合并(top-level · 横向)

- 若某一行在数据矩形 `[首有效列, 末有效列]` 范围内 **恰好只有 1 个非空 cell**,且该 cell 之后的 cell 全为空 → **将该 cell 横向合并到整个数据矩形**(即"总标题行"跨多列展示)

- 典型场景:第 1 行总标题(只填了 A1,后面列空),整行合并 A1:H1

##### 规则 B · 相邻空合并(横向 · `_merge_header_cells`)

- 找到该行每个非空 cell,若其**右侧 cell 为空**但**右侧列在其它行有内容**(整列非空)→ 向右合并,直到碰到非空 cell 或整列空

- 支持多格合并:如 A\_row:D\_row("name"+ 2 个空 cell + "value")

- 仅发生在表头行;数据区域不自动合并

##### 规则 C · 列纵向合并(纵向 · `_merge_header_col_cells`)

- 仅对\*\*第一列(A 列)\*\*做纵向合并

- 找到 A 列每个非空 cell,若其**下方 cell 为空**但**下方行在其它列有内容**(整行非空)→ 向下合并,直到碰到非空 cell 或整行空

- 支持多格合并:如 A1:A3(分类 1+2 行空 + 分类 2)

- **关键检查**:纵向合并前先检查下方 cell 是否已在任何合并范围内 → 若是,跳过本次纵向合并(避免重复)

- 仅发生在表头行 + 第一列;数据区域不自动合并

##### 调用顺序

```
边框 → 字体 → 表头加粗底色 → 横向合并(_merge_header_cells) → 纵向合并(_merge_header_col_cells) → 列宽自适应 → 行高自适应
```

> 顺序至关重要:边框先设(确保 Excel 打开整块边框完整),最后做合并;横向合并先做,纵向合并后做(纵向会跳过"已在合并范围内的 cell")。

## Stage 4 · 统一输出 (Output)

**职责**:把阶段三样式化的内容落盘,产出最终 `.xlsx` 文件。

### 输出规范

- **清洗结果文件**:`.xlsx`(唯一输出格式)

- **文件名规则**:`<原文件名>_cleaned.xlsx`

- **后缀强制校验**:`write_xlsx` 拒绝任何非 `.xlsx` 输出后缀

- **不生成额外 sheet**:只保留清洗后的数据 sheet,**不再生成** `_CLEAN_REPORT` 或 `_ANOMALIES` 等附加 sheet

- **列 number\_format(显示格式,不影响底层值)**:

  - 字段被识别为 `amount` / `number`(非文本数字):统一显示 `0.00`

  - 字段被识别为 `datetime`:显示 `yyyy-mm-dd hh:mm:ss`

  - 字段被识别为 `month`:显示 `yyyy-mm`

  - 字段被识别为 `id` / `long_text` / `phone` / `email` / `boolean`:显示 `@`(文本)

### 后处理

- `save()` 原子落盘,失败时保留临时文件便于排障

> **已废除**:之前版本中的 `X.X万` / `0.0,"万"` 数字格式应用逻辑已彻底删除,本 Skill 不再对任何字段应用此类格式化字符串。

## Usage

### 命令行(本地)

```bash
# 默认清洗(输出 .xlsx)
python scripts/clean_excel.py \
    --input path/to/file.csv \
    --output path/to/output_cleaned.xlsx \
    --rules config/cleaning_rules.yaml

# 只清洗某个 sheet
python scripts/clean_excel.py \
    --input path/to/file.xlsx \
    --output path/to/output_cleaned.xlsx \
    --sheet "Sheet1"
```

### Python API

```python
from scripts.clean_excel import clean_file

result = clean_file(
    input_path="path/to/file.xlsx",
    output_path="path/to/output_cleaned.xlsx",
    rules_path="config/cleaning_rules.yaml",
)
print(result["summary"])
print(result["input_sheets"])
```

### 在 Agent 中调用

当用户给出 Excel/CSV 文件并要求"清洗"、"整理字段"、"规范样式"时:

1. 调用 `scripts/clean_excel.py` 清洗
2. 把输出文件路径回传给用户
3. 控制台摘要(`summary` 字段)含命中规则数、空值数、异常数等
4. 不再支持 `.xlsm` 文件(会抛 `ValueError`)
5. **不生成** `_CLEAN_REPORT` / `_ANOMALIES` 等附加 sheet

## Inputs

- 输入文件:`.csv` / `.xls` / `.xlsx`

- 可选:自定义规则文件(YAML)

## Outputs

- `<name>_cleaned.xlsx`:清洗结果(仅数据 sheet,统一样式)

- 控制台摘要:命中规则数、空值归一数、异常标记数

## Directory Structure

```
excel-data-cleaner/
├── SKILL.md
├── README.md
├── .gitignore
├── config/
│   └── cleaning_rules.yaml      # 字段类型识别 + 异常阈值 + 样式配置
├── scripts/
│   └── clean_excel.py           # 主清洗脚本(单文件,所有阶段在此)
├── references/
│   └── field_patterns.md        # 字段识别正则与边界说明
├── tests/
│   ├── test_header_detection.py # 表头识别单元测试
│   ├── test_data_rect.py        # 有效范围端到端测试
│   ├── test_effective_range.py  # 有效行/有效列/行高测试
│   └── test_watermark_strip.py  # AI 生成水印防御测试
└── examples/
    ├── sample_dirty.csv         # 脏数据样例(输入)
    └── sample_cleaned.xlsx      # 清洗后样例(输出)
```

## Configuration

默认规则已覆盖 80% 场景;若需调整阈值或新增字段类别,编辑 `config/cleaning_rules.yaml`:

```yaml
fields:
  id:
    column_keywords: [id, 编号, no, code, 工号, 订单号, 票号, 号]
    cell_format: "@"
    is_primary_key: true
  datetime:
    column_keywords: [datetime, date, time, 日期, 时间, timestamp]
    output_format: "YYYY-MM-DD HH:MM:SS"
    cell_format: "yyyy-mm-dd hh:mm:ss"
    force_24h: true
string_normalize:
  unicode_nfkc: true
  trim: true
  strip_internal_whitespace: false
anomalies:
  future_date_tolerance_days: 1
  out_of_range_abs: 1e12
  stat_outlier_std_multiplier: 2
  stat_outlier_min_samples: 10
output_style:
  font_name: "微软雅黑"
  font_size: 9
  font_color: "FF000000"
  border_color: "FF000000"
header_detection:
  max_scan: 10
```

## Error Handling

- 输入文件不存在 → 抛出 `FileNotFoundError`,提示路径

- 文件格式不支持(含 `.xlsm`)→ 抛出 `ValueError`,列出支持的后缀

- 规则文件不存在 → 使用内置默认规则

- 写 xlsx 失败(权限/磁盘) → 抛出原始 IO 错误并保留临时文件

## Dependencies

- Python ≥ 3.10

- `pandas`, `openpyxl`, `xlrd`(读 `.xls`), `pyyaml`

安装:

```bash
pip install pandas openpyxl xlrd pyyaml
```

## 大数据量性能说明

> **场景**:单 sheet 超过 5 万行 × 50 列(约 250 万 cells)时,openpyxl 的 `wb.save()` 序列化成为瓶颈。

### 引擎选择(`clean_file(..., engine=...)`)

- `"auto"`(默认): 数据量 ≥ 50 万 cells 自动走 `xlsxwriter`,否则走 `openpyxl`
- `"openpyxl"`: 强制走 openpyxl(适合中小文件)
- `"xlsxwriter"`: 强制走 xlsxwriter(适合大文件)

### `xlsxwriter` 后端特性(`_write_xlsx_fast`)

- **表头合并**: 复用 skill 的规则 A/B/C,预计算合并范围后用 `worksheet.merge_range()` 一次性写出
- **空值处理**: None/NaN/空 一律用 `worksheet.write_blank(r, c, None, fmt)`,避免 `write()` 把 None 误转为数值 0
- **数值类型还原**: amount / number / percent 列的纯数字字符串用 `write_number()` 写出
- **boolean 列**: True/False 字符串用 `write_boolean()` 写出
- **color 6字符 RGB**: xlsxwriter 对 8 字符 aRGB 处理有 bug(再加 FF 变 10 字符,导致 openpyxl 解析失败);skill 内部统一剥到 6 字符 RGB

### 调用示例

```python
from clean_excel import clean_file

# 大文件自动走 xlsxwriter
result = clean_file(input_path, output_path)

# 强制走 xlsxwriter(适合已知大数据量场景)
result = clean_file(input_path, output_path, engine="xlsxwriter")

# 强制走 openpyxl(适合需要丰富样式的中等文件)
result = clean_file(input_path, output_path, engine="openpyxl")
```

### 字段类型: percent

新增 `percent` 字段类型,专门识别"费率/比例/占比/折算比例/折扣/费率折扣/使用比/gmv比/percentage/rate" 列:

```yaml
percent:
  cell_format: "0.00%"
  allow_negative: false
  min_value: 0
  max_value: 1  # 默认认为是 0~1 的小数比例
```

> **重要字段顺序**: `percent` 必须在 `amount` 之前,否则 `amount.gmv` 会抢先命中"结汇GMV占比"等列。

### `detect_field_type` 最长匹配

- 在所有 ftype 的关键词中,找到**最长匹配**的关键词(更具体的关键词胜出)
- 平局时按 fields 字典顺序优先
- 例:`"比例费金额(USD)"` 同时含 `比例`(2字符) 和 `比例费金额`(5字符) → 5 字符胜出 → amount
- 业务特定关键词(如 `"比例费金额"` / `"gmv占比"`)可精确胜出通用短关键词

### `xlsxwriter` 样式细节

- **表头**: 水平+垂直居中(vcenter),浅蓝底 (#D9E1F2),加粗
- **数据**: 水平+垂直居中(vcenter),`text_wrap=False`(不换行;靠列宽自适应保证可见)
- **列宽**: 按"表头 + 抽样数据行字符数"计算(中文 2 字符宽 / ASCII 1 字符宽),min=8 / max=40
- **合并标题**: `_write_xlsx_fast` 把原 cell 值传给 `merge_range()`(避免合并后标题为空)

### AI 水印清理

`_strip_ai_drawing_layers(xlsx_path)` 在写出完成后自动清理(本 Skill 的输出**不得**包含任何 AI 生成标识 / 水印 / 装饰图层 / 隐藏对象):

- 应用场景: 原文件含 WPS/Excel 插入的"AI 生成"水印(典型如 `xl/drawings/drawing_wm1.xml`,内含名为 `AILabel` 的文本框)
- 处理步骤:
  1. 一次性读取 zip 全部内容到内存,关闭后再处理(避免 Windows 双句柄冲突)
  2. 删除所有 `xl/drawings/drawing*.xml` 文件
  3. 修改 `xl/worksheets/_rels/sheet*.xml.rels` → 删除 drawing Relationship
  4. 修改 `xl/worksheets/sheet*.xml` → 删除 `<drawing r:id="..."/>` 标签
  5. 修改 `[Content_Types].xml` → 删除 drawing Override 声明
  6. 临时文件用 `tempfile.gettempdir()`(避免 sharing violation)
  7. 替换用 `os.replace()`;失败时不再静默吞异常,而是 fallback 留 `_cleaned.xlsx` 副产物
- 调用位置:
  - `_write_xlsx_fast` 完成 `workbook.close()` 后(大数据量场景)
  - `_apply_styles` 完成 `wb.save()` 后(中小文件场景)

> **v13 变更**: 删除了旧的 `_strip_ai_watermarks` 双重调用(扫 sharedStrings / VML 文本),新逻辑在 drawing 文件删除后已无残留文本,旧函数无存在意义。

## 修订记录

- **2026-09-15 v13**: 删除旧水印逻辑 `_strip_ai_watermarks`(扫 sharedStrings / VML 文本),`tests/test_watermark_strip.py` 同步删除;SKILL.md 字段识别表同步 v10/v11/v12 新增关键词(datetime/月/百分比/amount/number/boolean 全量)
- **2026-09-14 v12**: amount 关键词追加"余额/balance"修 BJ 列(以及所有"账户余额/可用余额/结余"类列)
- **2026-09-14 v11**: amount 关键词扩展(手续费/调整金额/成本/利润)修 AV 列;`clean_file` header_labels 多层 fallback(自下而上找非空) + 用 `pd.notna()` 过滤 numpy.nan 修 BR 列;`_strip_ai_drawing_layers` 重构(文件锁 fallback 留副产物 + 不再静默吞异常)彻底修 AI 水印
- **2026-09-13 v10**: `_write_xlsx_fast` 合并算法重写(抄 `excel-style-cleaner` v3.22+ 两遍扫描);datetime.cell_format 改为 `yyyy-mm-dd`;全表单元格 text_wrap/居中
- **2026-09-12 v6**: 修复 A1 合并范围(规则 B 不再被"整列空"中断,可一直扩展到下一个非空 cell);`_strip_ai_drawing_layers` 重构(一次性加载 zip + `os.replace` 替代 `shutil.move`,修复 Windows PermissionError)
- **2026-09-12 v5**: 表头改为水平居中(vcenter);新增 `_strip_ai_drawing_layers()` 清理 AI 水印 drawing 层(drawing_*.xml + sheet rels + Content_Types)
- **2026-09-11 v4**: 修复标题合并为空(`merge_range` 传原 cell 值);`detect_field_type` 改为最长匹配;amount 加 `gmv/比例费金额` 关键词,percent 加 `gmv占比` 关键词,month 改精确,number 加 `月份数/次数`;数据单元格居中(vcenter) + `text_wrap=False` + 列宽自适应
- **2026-09-11 v3**: 新增 `percent` 字段类型(费率/比例/占比识别为百分比);`number` 加 `gmv` 关键词(让 GMV 列还原为数值);新增 `_write_xlsx_fast`(xlsxwriter 后端);`write_xlsx` / `clean_file` 支持 `engine` 参数(`auto`/`openpyxl`/`xlsxwriter`)
- **2026-09-11 v2**:`_apply_styles` 增加 NamedStyle 共享对象 + `ws._cells` 快速路径(≥50 万 cells),优化大文件样式应用性能
- **2026-09-09**:R6 重复清理,删除 IDE 全局版,保留用户工作区版;发现 TRAE `disabledSkills` 机制
- **2026-08-29**:从 GLOBAL 迁移到 PROJECT(Data_Operation)scope
- **2026-08-28**:初始创建

