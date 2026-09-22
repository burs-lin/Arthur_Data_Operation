---
name: excel-style-cleaner
updated: 2026-09-14
description: >
  通用 Excel 样式清洗工具（v3.44）。**使用场景：看板 / 汇总报表（无唯一性主键、含合计 / 汇总 / 达成率 / 环比 / 同比等聚合指标）。明细表请改用 excel-data-cleaner。**
  四阶段清洗：① 多格式读入（.csv/.xls/.xlsx/.xlsm）→ ②
  **v3.44**：① **`is_date_like` 识别季度字符串**（`\d{4}\s*年\s*[Qq][1-4]` + `\d{2}\s*年\s*[Qq][1-4]`）——识别 "26年Q1" / "2026年Q1" / "26年q2" 等。② **`apply_number_formats` date fallback 修复**——sheet"0 季度核心指标达成情况"的 row 4 是季度（"26年Q1"），达成率列在 row 3 是"达成率"；旧版 `is_date_like("26年Q1")=False` → date fallback 失败 → last_header 误用 "26年Q1" → fmt 设为 `#,##0.00` 而非 `0.00%`。修复后达成率列正确显示 `0.00%` + 黄色 DataBar。③ **Phase 2 横向合不再受 `phase1_prev_cover` 约束**——Phase1 prev_cover 用于限制 Phase1 行表头"row N+1 不超 row N 合并列范围"，但 Phase2 处理的是数据行（列表头列），与 Phase1 行表头无关。改用 `max(header_cols)` 让 A=B=C 的相邻列能完整合并（如 sheet"2" row 10 A10:C10='<region_A1>' 全合）。
  **v3.42**：① **`is_date_like` 识别 2 位年份字符串**（`\d{2}\s*年\s*\d{1,2}\s*月`）——`format_header_dates` 把日期单元格式化为 "25年6月" 后，旧版正则只识别 4 位年份，导致 `apply_number_formats` 的 `last_header` date fallback 失效，整列走"普通数字"分支 → 错格式（如"参与考核人数"列被设成 `#,##0.00` 而非 `#,##0`）。② **条件格式双触发互斥**——`apply_rate_data_bar` 跳过已被 `apply_momyoy_data_bar` 标记的列（`ws._momyoy_cols`），并主动检查 `is_momyoy_in_header_cells` 兜底，避免环比/同比列同时被加上"黄色 Data Bar + 红绿 Data Bar + 3 个箭头公式"三重叠加。
  **v3.41**：① **统一数据单元格字段判断逻辑**——`apply_number_formats` 改为 5 优先级：percent → integer → GMV 万级 → 普通万级 → 普通数字；② **最后一层 header 判定 + date fallback**——避免上层"结汇率/达成率"等关键字污染整列；③ **GMV 关键字扫描严格化**——只匹配"明确的 GMV 列名"（`付款GMV/入账GMV/全量GMV/增量GMV/存量GMV/结汇.*GMV/GMV($)/GMV同比/GMV环比/GMV达成`），排除"GMV考核/GMV完成"等子模块名误判；④ **排除 GMV 列加进度条**——`apply_rate_data_bar` 主动排除 GMV 列（GMV 应是 X.X万格式而非达成率进度条）。
  **v3.40**：① **保留迷你图（Sparkline）**——openpyxl 加载时丢弃 `<x14:sparklineGroups>`，新增 `_preserve_sparklines_postsave` 从原始文件提取并合并到输出；含正确的 `<ext>` wrapper + `xmlns:x14` + `xmlns:mc` 命名空间声明。② **GMV 列强制万级格式**——`apply_number_formats` 新增 has_gmv 判定，与原均值万级并列。
  **v3.39**：① **B6:C6 不应合并修复**——`row1_horizontal_covers` 不再处理跨行合并（A1:C4），避免 `prev_row_col_cover` 错误继承；② **`_remove_ai_drawings_from_workbook`** ——加载后立即清空 `ws._drawing/charts/images`，防止 openpyxl save 时重新写回；③ **`_strip_ai_artifacts_postsave` 清理 `xl/drawings/` + `xl/media/` + `xl/charts/`** 等浮层对象。
  **v3.38**：① **修复最后一行误加粗**：移除 `bold_header_and_total_rows` 中"最后一行自动加粗"的逻辑（最后一行不一定是合计行，不应该自动加粗）。
  **v3.37**：① **关键字调整**：`PERCENT_NO_BAR_KEYWORDS` 增加"分润率/费率"（只有百分比格式，没有进度条）；`INTEGER_COUNT_KEYWORDS` 增加"链接数/代理商数/合同数/人数"（整数格式）；② **进度条上限固定为100%**：`apply_rate_data_bar` 的 DataBar 上限从动态计算改为固定 `end_value=1`（100%）；③ **表头识别修复**：`looks_like_header_row` 的说明关键字匹配从单字匹配改为短语匹配；④ **跨行合并覆盖继承**：`row1_horizontal_covers` 支持跨行合并。
  **v3.35**：① **合计行加粗增强**：B 列含"合计/总计/小计/汇总"时整行加粗；② **日期样式 `YY年M月`**；③ **表头合并扩展**：Phase1/Phase2 横向合并扩展允许空列接续；④ **比例样式统一**：`RATE_KEYWORDS` 增加"比例/比率"；⑤ **括号内容剥离**：新增 `_strip_parenthetical_content`。
  **v3.34**：① **`is_date_like` 识别 Excel 序列日期整数**；② **`format_date_cell` 支持序列日期数字**；③ **新增「占比」「比例」「客户数/会员数/账号数/订单数」列类型识别**；④ **`detect_layout` 表头行判定**：日期数字不再触发 `break`。
  **v3.31**：① **`_get_header_cells_in_col` 增强**：当某表头行单元格本身为空时，左右扩展寻找最近非空值；② **`detect_layout` header_cols 判定加强**；③ **`apply_number_formats` 复用 `_get_header_cells_in_col`**；④ **Phase1 prev_row_col_cover 横向合覆盖继承**。
  **v3.28**：① **`auto_fit_columns` 第一阶段 `horizontal_merges` 收集新增 `MAX_HEADER_TITLE_SPAN=6` 过滤**；② **条件格式触发条件改为"列的所有表头单元格中任意一个含关键字"**；③ **删除"严格金额列"触发条件**。
  **v3.27**：① **去除行列定义改为单元格定义**；② 新增「**表头单元格 = (r in header_rows) OR (c in header_cols) 的并集**」；③ `detect_layout` 重写；④ 下游 5 函数全部接受 `header_cells` 参数。
  **v3.22**：① 表头合并算法重构为「**逐行纵→横**」(`merge_header_by_rows`)；② 「**竖向优先**」规则；③ 「**下一行合并不超上一行**」规则。
  **v3.21**：① **新增 `merge_header_rectangles` 矩形合并二次扫描 pass2**；② **RATE_KEYWORDS 短关键字清理**；③ **AI 浮层对象清理范围扩展**。
  **v3.20**：① 万级底层值 /10000 + fmt `0.0万`；② 合计行不再被识别为表头；③ `is_numeric_for_header` 增强识别字符串数字；④ `format_header_dates` 跳过合并区域副格。
  **v3.19**：① 多层表头（header_rows 列表）智能合并 / 月份格式化 / 加粗全支持；② 合并范围 ≥2 列即可触发加粗；③ 关键字扩展；④ 阶段四新增 `_strip_ai_artifacts_postsave`。
  **v3.18**：阶段四清理 AI/工具生成的水印层（docProps/core.xml + docProps/app.xml）。
  **v3.17**：自检扩展覆盖千分位 / 短数字 / 负数红 / 自定义格式 6 类规则。
  **v3.16**：xwriter 模式沿用 input cell.number_format。
  **v3.15**：表头识别阈值暴露为 `--header-threshold` / `--header-max-scan`。
  **v3.13**：表头识别放宽到"连续 N 行含数字"。
  **v3.12**：打通 openpyxl → xlsxwriter；新增 xlsxwriter 写出引擎 `excel_writer.py`。
  **v3.9**：去除 AI 输出水印。
  **v3.7**：明细大表降级路径。其余 v3.4–v3.6 规则保留：先定位标题/内容区，统一微软雅黑 9
  号黑色字体与居中对齐，空行/空列不绘制边框，标题区空值保留原值用于合并、内容区空值填充 "-"；表头行 + 含"合计/总计/小计"的行加粗；横向合并单元格所在行 +
  是标题下首行或最后一行 → 整行加粗；达成率/通过率/完成率/环比/同比 一律 0.00%（底层值不变，根据表头行判断）；达成率/通过率/完成率 用黄色 Data
  Bar、环比/同比 用红色 Data Bar + 上中下箭头；金额列只居右不加 ¥；万级用 `0.0"万"` + 底层值
  /10000；表头行/表头列智能合并（先横后纵，纵合前检查已合并）；表头日期单元格统一为 YY年MM月 格式；列宽最小宽度自适应（内容不换行、标题≤2行）。
  Invoke when user needs to 清洗表格样式 / 规范化 Excel 报表 / 统一表格字体 / 给数据加进度条 / 表头智能合并 / 日期格式化 / 万级格式化 / 明细大表分块清洗 / 批量生成报表 / 去除 AI 水印 / 清除 AI 浮层对象 时使用。
---

# Excel Style Cleaner (v3.38) — 四阶段清洗架构

## 概述

本 skill 对任意 `.csv / .xls / .xlsx / .xlsm` 文件执行**四阶段规范化处理**，产出一份"视觉一致、可直接展示"的看板 / 汇总报表。

**底层实现**：
- `resources/excel_style_cleaner.py` — **openpyxl 引擎**（默认；保留原 sheet 结构、合并、批注）
- `resources/excel_writer.py` — **xlsxwriter 引擎**（`--xwriter` 模式；纯写新表；性能更好）

两套引擎共享同一套样式规则（关键字、数字格式、颜色、列宽）。CLI 主入口 `excel_style_cleaner.py` 已接通 `--xwriter` 等 13 个参数。

## When to Invoke（适用 / 不适用场景）

### 适用场景（适合本 skill）

- 销售看板 / 月度业绩看板 / 经营分析报表等**汇总类报表**
- 表格中含**合计 / 汇总 / 小计 / 达成率 / 完成率 / 环比 / 同比**等聚合指标
- **没有唯一性主键**，存在多级合并表头与分类汇总行
- 需要为达成率列加黄色 Data Bar、为环比 / 同比列加红绿箭头进度条
- 需要把金额列以万级格式（`0.0"万"`）展示
- 纯写新表场景（BI 导出 / 批量出表）→ 走 `--xwriter` 模式

### 不适用场景（请改用 `excel-data-cleaner`）

- 订单明细 / 会员明细 / 商品明细 / 物流明细等**明细表**
- 表格中存在**唯一性字段主键**（订单号 / 会员号 / ID / 编号 等）
- 需要按主键查重 / 字段类型推断 / 异常值标记等**数据层清洗**
- 仅做数据分析、统计、透视表 → 用 `xlsx` skill
- 需要创建新文档（不是清洗已有文档）→ 用 `xlsx` skill
- 需要读取 PDF/Word 中的表格数据 → 用 `pdf` / `docx` skill

## Pipeline Overview

整个清洗流程划分为 **四个阶段**，每个阶段职责单一、产出明确：

```
┌──────────────┐    ┌──────────────────┐    ┌──────────────────┐    ┌──────────────┐
│  Stage 1     │    │  Stage 2         │    │  Stage 3         │    │  Stage 4     │
│  Input ·     │ →  │  Format Clean ·  │ →  │  Visual Clean ·  │ →  │  Output ·    │
│  多格式读入  │    │  格式清洗        │    │  可视化清洗      │    │  统一输出    │
└──────────────┘    └──────────────────┘    └──────────────────┘    └──────────────┘
   • .csv/.xls/      ① 字段格式清洗         ① 统一样式               • 写出 .xlsx
     .xlsx/.xlsm        (智能数字格式)        · 应用范围识别          • 两种引擎可选:
     读入              ② 空值处理              (有效行/有效列)         openpyxl 读改存
   • sheet 选择      ③ 时间统一              · 字体/字号/颜色          xlsxwriter 写新表
   • header=None    ④ 字符串归一              · 边框/对齐/换行
   • 表头识别       ⑤ 异常标识              · 条件格式（Data Bar）
                                              ② 表头样式
                                              · 多级表头识别
                                              · 单元格合并
                                              · 加粗 + 背景填充
                                              ③ 行高列宽自适应
```

### 阶段对照表

| 阶段 | 关键产物 | 对应函数 / 引擎 |
|------|---------|----------------|
| Stage 1 · 多格式读入 | `{sheet: DataFrame}` + 表头行 + 单元格 format | `openpyxl.load_workbook` / `read_only=True, data_only=True` 快照 |
| Stage 2 · 格式清洗 | 清洗后 `DataFrame` + 字段类型映射 | `apply_smart_number_format` / `fill_empty_content_cells` / `format_header_dates` |
| Stage 3 · 可视化清洗 | 带样式的 xlsx | `smart_merge_cells` / `apply_base_style` / `bold_header_and_total_rows` / `apply_rate_data_bar` / `apply_momyoy_data_bar` / `apply_money_style` / `auto_fit_columns` |
| Stage 4 · 统一输出 | 最终 `.xlsx`（含 docProps 水印清理） | `clean_workbook` / `xwriter_export` / `write_styled_xlsx` / `write_multi_sheet_xlsx` |

### 两套实现引擎

| 引擎 | 触发 | 输入 | 输出 | 适用 |
|------|------|------|------|------|
| openpyxl 版（默认） | 不传 `--xwriter` | 任意 `.csv/.xls/.xlsx/.xlsm` | 修改原文件结构（保留 sheet 顺序、合并、批注等元数据） | 清洗旧表 |
| xlsxwriter 版 | 传 `--xwriter` | 任意 `.csv/.xls/.xlsx/.xlsm` | 新建 `.xlsx`（headers + data 重写） | 纯写新表 / BI 导出 / 批量出表 / 性能敏感 |

两套引擎**完全共享同一套样式规则**（数字格式 / 关键字 / 颜色 / 列宽），切换引擎不会改变样式效果。

---

## Stage 1 · 多格式读入 (Input)

**职责**：把任意支持的格式统一读入内存，产出可被 Stage 2 / 3 处理的内部表示。

### 支持格式

| 后缀 | 读取方式 | 引擎 |
|------|---------|------|
| `.csv` | `pd.read_csv(header=None, dtype=str)` | pandas |
| `.xls` | `pd.read_excel(header=None, dtype=str, engine='xlrd')` | pandas + xlrd |
| `.xlsx` | `pd.read_excel(header=None, dtype=str, engine='openpyxl')` | pandas + openpyxl |
| `.xlsm` | `load_workbook(...)` 保留宏 + 样式元数据 | openpyxl |

> 输入后缀不在上表 → 抛 `ValueError`，提示支持的后缀列表。

### 关键约定

- **xwriter 模式**：用 `read_only=True, data_only=True` 一次性读出 values + `cell.number_format`（v3.16 起，用于沿用 input 格式）
- **openpyxl 模式**：正常 load，可读写后保存
- **列名统一为整数 0..n-1**：避免丢表头行
- **CSV 单独处理**：`keep_default_na=False` 防止空字符串被识别为 NaN
- **多 sheet 处理**：`.xls` / `.xlsx` / `.xlsm` 返回 `{sheet_name: DataFrame}` 字典
- **明细大表自动降级**：行 × 列 ≥ `LARGE_SHEET_CELLS`（默认 200 万）→ 未传 `--force-clean` 时仅复制原文件并打印提示

---

## Stage 2 · 格式清洗 (Format Clean)

**职责**：对 Stage 1 读入的数据做内容层清洗，产出"语义正确"的内部表示。

### ②-① 字段格式清洗（智能数字格式）

根据**列标题关键字**与**列均值**自动选择数字格式：

| 触发条件 | 格式 | 示例 |
|---------|------|------|
| 列均值 ≥ 10,000 | `0"."0,"万";[Red]-0"."0,"万"` | `15580000` → `1558.0万` |
| 列标题含"达成率/完成率/达标率/完成比/完成度/通过率" | `0.00%;[Red]-0.00%` | `0.85` → `85.00%` |
| 列标题含"环比/同比/MoM/YoY" | `0.00%;[Red]-0.00%` | `0.12` → `12.00%` |
| 其它数字列 | `#,##0.00;[Red]-#,##0.00` | `1234.5` → `1,234.50` |

**v3.16 增强**：xwriter 模式沿用 input cell.number_format（`preserve_input_format=True`），优先级 **input 格式 > 表头关键字判定**。例如 input 是 `0.00%` → 输出仍是 `0.00%;[Red]-0.00%`，不重新判定。

### ②-② 空值处理

按"标题区 / 内容区"分区处理：

| 区域 | 行为 |
|------|------|
| 标题区（`header_rows` × 全部列） | 保留 None / 空字符串（不填 `-`） |
| 内容区（`data_rows` × `data_cols`） | 空值统一填 `-` |

空值定义（含 Stage 1 归一后）：

- 空字符串 `""`、纯空白（NFKC + trim 后仍为空）
- 大小写不敏感的字符串：`null` / `NULL` / `None` / `N/A` / `n/a` / `NA` / `-` / `—` / `#N/A` / `\\N`
- pandas 读入后的 `NaN` / `NaT`

### ②-③ 时间统一（表头日期格式）

仅作用于 **表头行** 的日期单元格（xlsxwriter 模式 + openpyxl 模式）：

- `2024-01` → `2024年01月`
- `2024-01-15` → `2024年01月`（表头场景仅取年月）
- `2024/01/15` → `2024年01月`

数据行日期**不**做归一（保持 input 原样）。

### ②-④ 字符串归一（作用于所有字段）

| 步骤 | 行为 |
|------|------|
| 1 | Unicode NFKC（全角→半角、不间断空格→空格） |
| 2 | tab / `\xa0` → 普通空格 |
| 3 | 首尾 strip |
| 4 | 文本字段不清除段落内空白 |
| 5 | 不合并段落内连续空白 |
| 6 | 邮箱统一小写 |

### ②-⑤ 异常标识

| 异常类型 | 触发条件 | 标识 |
|---------|---------|------|
| ID 重复 | 主键列出现 2 次及以上 | `[DUP_ID]` |
| ID 缺失 | 主键列空 | `[MISSING_ID]` |
| 全行重复 | 整行的所有列值与之前某行完全相同 | `[DUP_ROW]` |
| 日期/时间超出范围 | 日期 < 1970-01-01 或 > today+1 | `[BAD_DATE]` |
| 金额非法 | 金额字段含非数字字符 | `[BAD_AMOUNT]` |
| 长文本超长 | 字段长度 > 500 字符 | `[TEXT_TOO_LONG]` |
| 数值越界 | 数值字段绝对值 > 1e12 | `[OUT_OF_RANGE]` |
| 统计学异常值 | 数值列 Mean ± 2σ 之外（样本 ≥ 10） | `[STAT_OUTLIER]` |

> **看板 / 汇总报表场景**通常**不主动启用异常标识**（无唯一性主键），但本 skill 保留该能力以备需要。

---

## Stage 3 · 可视化清洗 (Visual Clean)

**职责**：对 Stage 2 清洗后的内容，**在样式层做统一化**，不修改底层数据。

### ③-① 统一样式

#### 应用范围识别（v3.27+ 单元格级定义）

| 术语 | 定义 |
|------|------|
| **表头行** | `header_rows` 列表中的行（自上而下扫描识别） |
| **表头列** | `header_cols` 列表中的列（在表头行范围 `[1, max(header_rows)]` 内逐列识别） |
| **表头单元格** | `(r, c) where r in header_rows OR c in header_cols` 的并集（v3.27+ 新增） |
| **数据行** | 剔除所有表头行后的行 |
| **数据列** | 剔除所有表头列后的列 |
| **数据单元格** | 有效行 × 有效列去除表头单元格的部分（v3.27+ 新增） |
| **整行(1:1)空** | 行 r 所有单元格都为空 → 该行不加边框 |
| **整列(A:A)空** | 列 c 所有单元格都为空 → 该列不加边框 |

**关键变更（v3.27+）**:
- 旧版：表头列范围 = 一刀切（整列视为表头/数据），导致数据行的表头列单元格（如 R3 A3）被错误地填 `-` 或应用数字格式
- 新版：表头列范围 = 线性识别（仅看表头行内的列值），但实际生效通过"表头单元格集合"——表头列范围内**所有行**的单元格都视为表头（保持整列线性判定）
- 数据单元格 = 去除表头单元格后的所有非空数据位置 —— 数字格式、Data Bar、居右、`-` 填充 都**只**作用在数据单元格

#### 字体 / 字号 / 颜色 / 对齐 / 边框

| 属性 | 取值 | 适用范围 |
|------|------|----------|
| 字体 | **微软雅黑** | 数据矩形内全部非空单元格 |
| 字号 | **9** | 同上 |
| 颜色 | **黑色** `#000000` | 同上 |
| 水平对齐 | `center` | 全部 |
| 垂直对齐 | `center` | 全部 |
| 边框 | 黑色细边框（thin） | 整行(1:1)或整列(A:A)为空时**不画** |

#### 条件格式（万 / 率 / 进度条 / 箭头）

| 触发条件 | 实现 | 视觉效果 |
|---------|------|---------|
| 列标题含"达成率/完成率/达标率/完成比/完成度/通过率" | 黄色 `FFFFD966` DataBar | 黄色进度条 |
| 列标题含"环比/同比/MoM/YoY" | 红色 `FFE06666` DataBar + 3 个 FormulaRule | 红 DataBar + 红↑/黄= /绿↓ |
| 列均值 ≥ 10,000 的数字列 | `0"."0,"万";[Red]-0"."0,"万"` | 数字按万级显示 |
| 严格金额列（含"金额/收入/支出/回款/定价"） | `0"."0,"万";[Red]-0"."0,"万"` + 右对齐 | 金额按万级显示，负数红 |

> 底层机制：openpyxl 把 DataBarRule / FormulaRule 写入 `ws.conditional_formatting`；xlsxwriter 写 `ws.conditional_format(...)`。Excel 打开时自动渲染。pandas 读取时**不会**保留条件格式（pandas 行为）。

#### 智能合并详细规则（v3.22+ — 逐行纵→横）

```
合并范围：仅标题行（第 1 行有值区间）和标题列（第 1 列有值区间）
内容行 / 内容列不做自动合并

  (a) v3.22+ 逐行合并（按 header_rows 顺序逐行处理）：
      对每个表头行 header_row：
        (1) 先竖向合并：本行每个源格（值格）向下延伸到第一个非空行
        (2) 再横向合并：本行剩余未合并的源格向右延伸

  (b) v3.22+ 竖向优先规则：
      若源格已竖向合并（A1:A2）→ 本行不再做横向合并
      避免 A1:A2 与 A1:B1 在同一 A1 处冲突

  (c) v3.22+ 下一行合并不超上一行：
      prev_row_col_cover[col] 记录「上一行该列源格的合并列范围」
      第 N+1 行横向合并：end_c ≤ prev_row_col_cover[c]
      首行初始化为 max_col（不限范围）

  (d) v3.21+ 矩形扩展 pass2（在逐行合并后）：
      已合并的横向区域 A1:B1 → 若下方 A2/B2 都空 → 扩展为 A1:B2
      已合并的竖向区域 A1:A2 → 若右方 B1/B2 都空 → 扩展为 A1:B2
      实现：merge_header_rectangles 二次扫描，unmerge 原区域 + re-merge 矩形
      关键：直接读 ws.cell(r, c).value，不走 get_cell_value（避免合并副格返回左上角值误判）
```

**注意**：
- 纯数字单元格不参与合并
- 内容行（第 2 行及以下）和内容列（第 2 列及以右）**不**做自动合并
- 已存在的合并区域不重复处理
- v3.21+ 矩形扩展**仅**当合并区在 [min_col, max_col] 范围内的下方行**整行**都为空才触发
- v3.22+ 规则 3 限制**仅**作用于**横向合并**;竖向合并不限制列范围（竖向只改行范围）

### ③-② 表头样式

#### 多级表头识别

- **表头行数** 由 `detect_header_rows_from_rows` 自动识别（v3.13 起：连续 N 行含数字才视为数据区起点；v3.15 起阈值暴露为 `--header-threshold` / `--header-max-scan` CLI 参数）
- **表头行** = 数据矩形内前 N 行（N = 自动识别的表头行数）
- **内容行** = 数据矩形内非表头的行

**v3.13 新算法伪代码**：
```
numeric_run = 0
for r in scan_rows:
    if not looks_like_header_row(r): continue
    if has_numeric:
        numeric_run += 1
    else:
        numeric_run = 0
    if numeric_run >= HEADER_NUMERIC_RUN_THRESHOLD:
        data_start = r - (threshold - 1)
        header_rows = filter(header_rows, < data_start)
        break
    header_rows.append(r)
```

**阈值选择指南**（v3.15+）：

| 阈值 | 适用场景 | 风险 |
|------|---------|------|
| 1 | 单层表头 + 数字数据 | 表头里有数字会被误判 |
| 2（默认）| 90% 场景（含"2024年"等含数字字符串的表头） | — |
| 3 | 5+ 行多层表头 / 表头有纯数字编号 | 数据行若夹 1 行文本会被误判 |

#### 表头加粗 + 背景填充 + 合并

- **加粗**：`Font(bold=True)`，仅作用于表头行的非空单元格
- **背景填充**：`PatternFill`（浅蓝），作用范围同上
- **单元格合并**：仅横向 + 标题行 + 内容行不自动合并
- **对齐**：`Alignment(horizontal="center", vertical="center")`

#### 标题 / 合计行加粗

- 表头行（`header_rows`）→ 加粗
- 含"合计/总计/小计/标题/汇总"的关键字行 → 加粗
- 横向合并单元格所在行 + 标题下首行 / 末行 → 加粗

### ③-③ 行高、列宽自适应

#### 列宽自适应

```
列宽 = max(
    内容区最大长度 + 2,         # 内容不换行
    ceil(标题区最大长度 / 2) + 2  # 标题最多 2 行
)
```

**字符宽度换算**：

| 字符类型 | 宽度 |
|---------|------|
| 中文字符 | 2.0 |
| ASCII 数字/字母 | 1.0 |
| 小数点 '.' / 千分位 ',' / 百分号 '%' | 0.5 |
| 加粗修正（+10%） | 中文 2.2 / ASCII 1.1 |

**Padding**：

- 默认 +2
- 万级列 +0.5（容纳 "万" 字符）
- 百分比列 +0.5（容纳 "%" 字符）
- 表头非空 +0.5（加粗修正）

**横向合并修正**：源格 → 内容字符宽度 / 合并列数 = 每列承担宽度（避免 B1 超长内容撑爆整列）。

#### 行高自适应

- 短内容（不需 wrap）：18 磅
- 1 行 wrap：30 磅
- 2 行 wrap：42 磅
- 横向合并的标题行：按 wrap 行数扩展

---

## Stage 4 · 统一输出 (Output)

**职责**：把 Stage 3 样式化的内容落盘，产出最终 `.xlsx` 文件，并清理 docProps 水印。

### 输出规范

- **清洗结果文件**：`.xlsx`（唯一输出格式）
- **文件名规则**：未指定输出路径时默认 `<input>.cleaned.xlsx`
- **xwriter 模式**：始终新建文件；openpyxl 模式：默认直接修改原文件（也可 `-o` 指定输出）
- **明细大表降级**：未传 `--force-clean` 时仅复制原文件

### 后处理 1：清理 AI/工具生成的水印（v3.18）

两套引擎都会在 `docProps/core.xml` 与 `docProps/app.xml` 中留下工具身份痕迹，**阶段四统一清理**：

| 引擎 | docProps/core.xml（用户元数据） | docProps/app.xml（应用元数据） |
|------|--------------------------------|--------------------------------|
| openpyxl | 写 `dc:creator` / `cp:lastModifiedBy` / `dc:title` 等 | 写 `<Application>Microsoft Excel Compatible / Openpyxl 3.x.x</Application>` ⚠️ |
| xlsxwriter | 写 `dc:creator="python-xlsxwriter"` 等 ⚠️ | 写 `<Application>Microsoft Excel</Application>`（中性，无工具痕迹）|

#### 清理策略

**docProps/core.xml**（用户元数据字段，**全部清空**）：

| 字段 | 引擎默认 | 清理后 |
|------|---------|--------|
| `dc:creator` | `""` (openpyxl) / `python-xlsxwriter` (xlsxwriter) | `""` |
| `cp:lastModifiedBy` | `""` / `python-xlsxwriter` | `""` |
| `dc:title` | `""` | `""` |
| `dc:subject` | `""` | `""` |
| `dc:description` | `""` | `""` |
| `cp:keywords` | `""` | `""` |
| `cp:category` | `""` | `""` |
| `dc:identifier` | `""` | `""` |
| `dc:language` | `""` | `""` |
| `cp:contentStatus` | `""` | `""` |
| `cp:revision` | `"1"` | `"1"`（保留，Office 必需）|
| `dcterms:created` / `dcterms:modified` | 时间戳 | 时间戳（保留，Office 必需）|

**docProps/app.xml**（应用元数据字段，**只清工具痕迹**）：

| 字段 | openpyxl 默认 | xlsxwriter 默认 | 清理后 |
|------|--------------|----------------|--------|
| `Application` | `Microsoft Excel Compatible / Openpyxl 3.x.x` ⚠️ | `Microsoft Excel` | `Microsoft Excel`（中性）|
| `Manager` / `Company` / `HyperlinkBase` | （不存在）| `""` | `""` |
| `AppVersion` / `DocSecurity` / `ScaleCrop` / `HeadingPairs` / `TitlesOfParts` | 标准字段 | 标准字段 | 保留（Office 必需）|

#### 实现：`_strip_watermarks(wb, xlsx_path)`

```python
from excel_style_cleaner import _strip_watermarks

# openpyxl 版（clean_workbook 内已自动调用）
_strip_watermarks(wb=wb)  # 清空 wb.properties
wb.save(output_path)
_strip_watermarks_postsave(str(output_path))  # 解 zip 改 app.xml

# xlsxwriter 版（write_styled_xlsx / write_multi_sheet_xlsx 内已自动调用）
# - 创建 wb 后立刻调 _strip_watermarks_xlsxwriter(wb)（即 set_properties({...})）
# - write 完成后调 _strip_watermarks_postsave(str(output_path))（双保险）
```

清理失败时静默兜底（不阻塞主流程）。

### 后处理 2：清理 AI/插件生成的浮层对象（v3.19）

AI 助手插件（如某些 WPS 插件 / "AI 帮我写" 插件）会在 xlsx 中插入**浮层对象 / 批注 / 嵌入图片** —— 这些通常显示为"AI 生成"灰色浮层，**不属于数据内容**，必须清理。

#### 删除的部件

| 部件 | 类型 | 触发场景 |
|------|------|----------|
| `xl/drawings/*.xml` + `_rels/` | 图形浮层（"AI 生成" 灰色浮层通常是 drawing） | AI 插件批量插入 |
| `xl/comments*.xml` | 批注（"AI 助手"批注） | AI 助手插件 |
| `xl/ctrlProps/*.xml` | ActiveX / 表单控件 | 表单残留 |
| `xl/activeX/*` | ActiveX 控件 | 老旧 Excel 模板 |
| `xl/embeddings/*` | 嵌入对象 | OLE 对象 |
| `xl/media/*` | 嵌入图片 / 截图 / 视频 | AI 生成的图标 |
| `xl/charts/*` | 图表 | 用户要求"禁止 AI 生成对象"含图表 |
| `xl/pivotTables/*` / `xl/pivotCache/*` | 数据透视表 | 用户要求含 |
| `xl/slicers/*` / `xl/tables/*` | 切片器 / 结构化表 | 用户要求含 |

#### 同时清理的引用

- `[Content_Types].xml`：删除对应的 `Override` 节点
- `xl/worksheets/_rels/sheet*.xml.rels`：删除对应的 `Relationship` 节点
- `xl/_rels/workbook.xml.rels`：删除对应的 `Relationship` 节点
- 每个 `xl/worksheets/sheet*.xml`：删除 `<drawing .../>` / `<legacyDrawing .../>` / `<picture .../>` / `<oleObjects .../>` 节点

#### 实现：`_strip_ai_artifacts_postsave(xlsx_path)`

```python
from excel_style_cleaner import _strip_ai_artifacts_postsave

# openpyxl 版（clean_workbook 内已自动调用）
wb.save(output_path)
_strip_watermarks_postsave(str(output_path))   # 后处理 1
_strip_ai_artifacts_postsave(str(output_path))  # 后处理 2

# xwriter 版（xwriter_export 内已自动调用）
write_multi_sheet_xlsx(...)
_strip_watermarks_postsave(str(output_path))   # 后处理 1
_strip_ai_artifacts_postsave(str(output_path))  # 后处理 2
```

清理失败时静默兜底（不阻塞主流程）。

#### 自检规则（v3.16 起，xwriter 模式）

xwriter 模式在写出后会自动检查输出文件的 `xl/styles.xml`，确保 input 格式被沿用。**v3.17 扩展为 6 类规则**：

| # | input 特征 | 输出期望 | 强制 |
|---|-----------|---------|------|
| 1 | 含 "%"（percent） | `0.00%;[Red]-0.00%` | ✅ |
| 2 | 含 "万" 或 "¥" | `0"."0,"万";[Red]-0"."0,"万"` | ✅ |
| 3 | 含 "#,##"（千级） | `#,##0.00;[Red]-#,##0.00` | ✅ |
| 4 | `0.0` / `0.00` / `0`（短数字） | 建议千级（记录但不强制） | ⚠️ |
| 5 | 含 "[Red]"（负数红） | 输出至少一个含 [Red] 的格式 | ✅ |
| 6 | 自定义格式 | 仅记入日志 | ℹ️ |

任一强制规则失败 → 抛 `ValueError`，提示用户传 `--no-preserve-input-format` 跳过。

### 颜色约定

| 用途 | HEX |
|------|-----|
| 黄色（达成率进度条） | `FFFFD966` |
| 红色（环比同比正向） | `FFE06666` |
| 绿色（环比同比负向） | `FF93C47D` |
| 黑色（边框 / 字体） | `FF000000` |

### v3.19：找回的 4 项规则

v3.18 重构时遗漏了4 项核心规则，v3.19 全部找回：

| # | 规则 | v3.18 状态 | v3.19 修复 |
|---|------|-----------|-----------|
| 1 | **表头合并**（多层表头） | ❌ `merge_header_cells` 只对 `header_row=1` 合并，第2/3 层未合并 | ✅ 支持 `header_rows` 列表，逐层合并 |
| 2 | **表头月份格式**（多层表头） | ❌ `format_header_dates` 只对 `header_row=1` 格式化 | ✅ 支持 `header_rows` 列表，逐层格式化 |
| 3 | **单元格合并加粗**（合并范围 ≥ 2 列） | ❌ 规则3 仅 `merge.max_col >= header_col_span` 才加粗 | ✅ 新判定 `merge_span >= 2` 即可触发 |
| 4 | **% 样式处理**（关键字缺失） | ❌ `RATE_KEYWORDS` 漏关键字，判定率低 | ✅ 扩展为 15 个关键字，含"完成度 / 达成 / 完成进度 / 完成情况" |

同时新增辅助函数 `_build_multi_layer_headers(ws, header_rows, max_col)`：按列拼接所有表头层的字符串用于关键字判定。多层表头时，关键字（如"完成度"）通常在中间层而非最后一层（最后一层往往是日期/编号）。

扩展后的关键字列表：

```python
RATE_KEYWORDS = [
    "达成率", "通过率", "完成率", "完成比", "达标率", "完成度",
    "完成率", "达成", "达标", "通过", "完成",
    "达成度", "完成进度", "完成情况", "达成情况",
]
MOM_KEYWORDS = ["环比", "MoM", "mom", "ring", "mom%", "月环比"]
YOY_KEYWORDS = ["同比", "YoY", "yoy", "year-over-year", "yoy%", "年同比"]
MOMYOY_KEYWORDS = MOM_KEYWORDS + YOY_KEYWORDS + ["同比环比", "环比同比", "环同比"]
```

---

## Usage

### 命令行（本地）

```powershell
# openpyxl 版（默认）—— 清洗旧表
python "<PROJECT_ROOT>/.agents/skills/excel-style-cleaner/resources/excel_style_cleaner.py" "<input_path>" [-o <output_path>]

# xwriter 版 —— 纯写新表（性能更好，BI 导出场景）
python "...excel_style_cleaner.py" "<input_path>" --xwriter
```

### 关键参数（13 个）

| 参数 | 说明 |
|------|------|
| `input` | 必需，输入文件路径（`.csv/.xls/.xlsx/.xlsm`） |
| `-o / --output` | 输出文件路径（默认 `<input>.cleaned.xlsx`） |
| `--header-rows 1-3` | 手动指定表头行（`1,2` 或 `1-3` 均可） |
| `--header-cols 1-2` | 手动指定表头列 |
| `--no-wan` | 禁用万级缩放（保留底层值 10000 不动） |
| `--max-rows N` | 明细大表仅清洗前 N 行（默认 2000；0 = 全量） |
| `--force-clean` | 对超大明细表（≥ 200 万 cells）也执行样式清洗（仅前 `--max-rows` 行） |
| `--debug` | 打印每个 sheet 写入的样式摘要（用于排查） |
| `--xwriter` | 走 xlsxwriter 纯写新表（v3.12+；性能更好；自动沿用 input 数字格式） |
| `--header-threshold N` | xwriter 模式：连续 N 行含数字视为数据区起点（v3.15+；默认 2） |
| `--header-max-scan N` | xwriter 模式：表头识别扫描行数（v3.15+；默认 10） |
| `--no-preserve-input-format` | xwriter 模式：关闭 input number_format 沿用（v3.16+） |
| `--keep-watermarks` | 保留 AI/工具生成的水印（默认会清理 docProps；v3.18+） |

### Python API

```python
from excel_style_cleaner import clean_workbook, xwriter_export
from pathlib import Path

# openpyxl 版（清洗旧表）
clean_workbook(Path("input.xlsx"), Path("output.xlsx"))

# xwriter 版（纯写新表，沿用 input 格式 + 自检）
xwriter_export(Path("input.xlsx"), Path("output.xlsx"))
xwriter_export(Path("input.xlsx"), Path("output.xlsx"), preserve_input_format=False)
xwriter_export(Path("input.xlsx"), Path("output.xlsx"), header_threshold=3)
```

### xlsxwriter 纯写新表 API

```python
from excel_writer import write_styled_xlsx, write_multi_sheet_xlsx

# 单 sheet
write_styled_xlsx(
    output_path="report.xlsx",
    headers=["区域", "销售额", "达成率"],
    data=[["<region>", 1500000, 0.92], ["<region>", 2300000, 0.85]],
    column_specs=[{"type": "auto"}, {"type": "money"}, {"type": "rate"}],
)

# 多 sheet
write_multi_sheet_xlsx(
    output_path="report.xlsx",
    sheets=[
        {"sheet_name": "汇总", "headers": [...], "data": [...], "column_specs": [...]},
        {"sheet_name": "明细", "headers": [...], "data": [...], "column_specs": [...]},
    ],
)
```

### 在 Agent 中调用

当用户给出 Excel 文件并要求"清洗样式 / 统一字体 / 加进度条 / 万级显示"时：

1. 判断场景：**看板 / 汇总报表** → 用本 skill；**明细表** → 改用 `excel-data-cleaner`
2. 调用 `excel_style_cleaner.py` 清洗（按需传 `--xwriter`）
3. 把输出文件路径回传给用户
4. 如启用 xwriter 自检失败，提示用户用 `--no-preserve-input-format` 跳过

---

## 决策记录（用户已确认）

| 决策点 | 用户选择 | 当前默认 |
|--------|---------|---------|
| 内容区空值是否填 `-` | 按 v3.1：内容区填 `-`，标题区不填 | 内容区填 `-`，标题区保留 None |
| 万级格式实现 | `0"."0,"万"` + 底层值不变 | 底层值不变 + 格式 `'0"."0,"万"'` |
| 百分比底层值 | 仍是小数 `0.85`（根据表头行判断） | 底层值 `0.85` 不变 |
| 表头日期格式 | 按 v3：`YY年MM月` | `2024年01月` |
| 列宽参数 | 最小宽度 + 内容不换行、标题≤2行 | `max(content_max+2, ceil(title_max/2)+2)` |
| 金额列货币符号 | 不需要 ¥ | 只居右，无货币符号 |
| 横向合并加粗规则 | 标题下首行 + 末行 → 加粗 | 表头行 + 横向合并行 + 末行 加粗 |
| 表头识别阈值 | 默认 2（v3.15） | 阈值暴露为 `--header-threshold` |
| xwriter 沿用 input 格式 | 是（v3.16） | `preserve_input_format=True` + 自检 |
| docProps 水印清理 | 是（v3.18） | 默认清理，`--keep-watermarks` 可关闭 |
| RATE_KEYWORDS 短关键字 | 删除「达成/达标/通过/完成」单字（v3.21） | 仅保留"成对"的关键字（达成率/通过率/完成率/完成比/达标率/完成度/达成度/完成进度/完成情况/达成情况），避免误匹配 |
| 矩形合并 pass2 | 是（v3.21） | 横向 A1:B1 + 下方整行空 → 扩展为 A1:B2；竖向 A1:A2 + 右方整列空 → 扩展为 A1:B2 |
| 表头合并算法 | 逐行纵→横（v3.22） | 按 header_rows 顺序逐行处理；每行先竖向再横向；竖向优先；下一行合并不超上一行 |
| 表头定义 | 单元格级并集（v3.27） | 表头单元格 = (r in header_rows) OR (c in header_cols)；数据单元格 = 有效行×有效列去除表头单元格 |
| 条件格式触发 | 表头单元格集合关键字（v3.28） | 列的所有表头单元格中任意一个含 RATE_KEYWORDS/MOMYOY_KEYWORDS → 触发 Data Bar |
| 金额列居右 | 已删除（v3.28） | 不再以 `is_money_column` 触发任何样式；`MONEY_KEYWORDS` 仅保留定义 |
| 表头行横向合右边界 | `max_col`（v3.45+） | Phase 1 横向合不受 `max(header_cols)` 限制——表头行源格可合到整片表头区（如 sheet"0" AS2 → AS2:AY3） |
| 表头行合并加粗范围 | 任何 ≥2 列横向合（v3.45+） | 删除 `is_first_after_header` 限制；任何 ≥2 列横向合并行都整行加粗 |
| `_get_header_cells_in_col` 扩展方向 | left-only（v3.45+） | 仅向左搜 max_search 列，不右扩展——防止跨业务组边界污染 date fallback 与 `apply_rate_data_bar` |
| `_extend_to_nearest_nonempty` 同上 | 保留原版 + left-only 变体（v3.45+） | `apply_number_formats` date fallback 与 `apply_rate_data_bar` 用 left-only；其他场景保留原版双向搜索 |
| 金额类列白名单 | `MONEY_LIKE_KEYWORDS` 强制万级（v3.45+） | 命中"金额/回款/营收/收入/销售额/销售金额/预缴/充值/退款/到账/付款/收款" → 强制万级，不依赖均值阈值 |
| Phase 2 横向合越界处理 | 三维判定（v3.47~v3.48） | (源格行位置: 表头行/数据行) × (next_val 语义: 同值/占位/空/不同值) × (a_col_has_label: True/False)；表头行源格合到 max_col；数据行源格仅严格同值允许跨 header_cols；该行 A 列有自己的标签时允许占位延伸 |
| `max(header_cols)` 限制用法 | v3.47 区分 v3.48 细化 | v3.47：仅当 `r not in header_rows_set` 时限制；v3.48：同时考虑 `a_col_has_label`，表头列 + 有自己 A 列标签的数据行可放宽 |
| 合并源格含 TITLE_KEYWORDS | 排除占位延伸（v3.48+） | 源格含"合计/总计/小计/汇总/标题" → 不允许占位延伸，防止 sheet"0" R9 B9='合计' 越界合到 BE9 |
| Phase 2.5 数据行非 header_col 列横向合 | 严格同值（v3.48+） | 遍历 c ∉ header_cols 的源格，仅严格同值合并，不允许占位延伸；解决 sheet"3.1" B12='<region_sub>' 横合到 C12 |
| 率列 vs GMV 列优先级 | 率列优先（v3.48+） | `apply_rate_data_bar` 中 `has_gmv and is_percent_column(last_header) → has_gmv=False`；即使含 GMV 关键字，本质是率列就按率列处理（修复"col_<biz_alias_2>..."缺 Data Bar） |
| `apply_rate_data_bar` 对 no-bar 列 | 跳过（v3.46+） | `is_percent_no_bar_column(last_header) → continue`；费率/汇率/分润率/分佣率/折算率 只走 percent 不加 Data Bar |
| `PERCENT_NO_BAR_KEYWORDS` 关键字 | 5 个（v3.46+） | 包含"分润率/费率/汇率/分佣率/折算率"；避免误判为普通列（无关键字命中 percent 判定） |
| `RATE_KEYWORDS` 含"留存率" | 加入（v3.46+） | "col_<biz_alias_2>..."含"留存率"被 `is_rate_column` 命中 → 走 percent 而非 GMV 万级 |
| `fill_empty_content_cells` 异常处理 | try-except 静默（v3.48+） | Phase 2 后跑时遇到合并副格写值会抛 AttributeError；加 try-except 静默跳过 |
| 竖向合并延伸判定 | 空/占位/与源格同值（v3.49+） | Phase 2 竖向合 next_val 允许 `is_value_for_merge` 或 `_can_extend_for_repeat_value`；支持重复值竖向合并（sheet"sheet_1_2" B2:B6 + B8:B16 + B17:B21） |
| Phase 1 表头行竖向合禁用重复值延伸 | 仅占位/空（v3.50+） | Phase 1 表头行源格（header_row ∈ header_rows）走竖向合时，禁止 `_can_extend_for_repeat_value` 延伸；防止 v3.49 在 sheet"2" R2 D2:D3 触发"重复值竖向合"→ D2 横合 D2:AS2 被 vertical_merged_cols 跳过 → merge_header_rectangles 扩展成 D2:Q3 越界 |
| Phase 2 列表头跳过数据行源格 | r ∈ header_rows 才处理（v3.50+），col=1 例外（v3.51+） | Phase 2 处理 header_col 时跳过 `r not in header_rows_set` 的源格；例外 col=1（A 列）不跳过——A 列是"行维度列"，整列所有行都是 header_cells 源格；防止 sheet"sheet_1_2" R7 B7='付款GMV'（数据行的"行维度标签"）被当作表头源格处理 → B7:C7 横合（C7 是数据）|
| Phase 2.5 跳过 A 列有标签的数据行 | A列真值 == cell_val 时跳过 c≠1 源格（v3.53+ 同 v3.52 语义） | Phase 2.5 处理数据行源格时，如果该行 A 列真值 == cell_val（同值），跳过 c≠1 的源格横向合；让 Phase 2 统一处理 A 列展开（A 列真值 != cell_val 时不跳过，如 R13 A13='<region_B>' B13='<region_sub>' 不同值，Phase 2.5 应当合 B13:C13；A 列是合并副格/占位时也不跳过）|
| Phase 1 行表头处理顺序 | 先横后竖（v3.53+） | v3.22+ 改为"先横后竖"——避免 R2 D2:D3 重复值竖向合冲突 D2:AS2 横合；行内扩展优先 → 先确定列范围再确定行范围 |
| Phase 2 列表头处理顺序 | 先横后竖（v3.53+，保留）| 列内扩展优先 → 按列号升序，每列内先横向合再竖向合 |
| Phase 2 占位跨 header_cols 边界 | col≥2 且 r∉header_rows 禁止占位跨边界（v3.53+） | 仅 `a_col_has_label=True 且 header_col=1` 时允许占位延伸；防止 sheet"sheet_1_2" R7 B7='付款GMV'（col=2）+ C7='-' → 不再合 B7:C7（C 列不在 header_cols）|
| merge_header_rectangles 独立合并干扰 | is_in_other_merge 排除嵌套合并（v3.53+） | 排除"以 exclude_merge 源格为左上角的合并"（如 A1:C1 + A1:A4 嵌套，不算独立合并），避免误阻断横向合向下扩展 |
| merge_header_rectangles 重叠 unmerge | unmerge 所有与新范围重叠的其他 merge（v3.53+） | 防止 openpyxl 在 merge A1:C4 时产生 A1:A4 + A1:C4 两个独立 merge 的 bug |
| apply_number_formats 行级覆盖 | B列指标标签命中 INTEGER_COUNT_KEYWORDS 覆盖列级格式（v3.54+） | 在 cell 循环里读 B 列（指标列 = max(header_cols)）的标签；若命中关键字该 cell 走整数 `#,##0`；覆盖列级"万级/普通"判定 |
| apply_number_formats 整列整数判定 | 均值<1万 + 整列都是整数 → `#,##0`（v3.54+）| 新增 `_column_all_integer` 辅助函数；优先级 5 普通数字列加分支 |

---

## 失败信号

| 信号 | 应对 |
|------|------|
| `❌ 输入文件不存在` | 检查路径；提示用户提供绝对路径 |
| `openpyxl` / `xlsxwriter` 缺失 | `pip install openpyxl xlsxwriter` |
| `.xls` 老格式保存失败 | 用 `xlrd`/`pyexcel` 预处理转 `.xlsx`；或用户手动另存 |
| WPS 不显示 Data Bar | WPS 部分版本对 DataBar 支持有限，可改用 Excel 打开 |
| 表头行识别错误 | 加 `--header-rows 1-3` 手动指定 |
| `⚠ 检测到明细大表（总_cells=…）` | 加 `--force-clean --max-rows N` 走快速路径 |
| 快速路径下未传 `--header-rows` 时只保留 1 行表头 | 加 `--header-rows 1-3` |
| xwriter 自检失败 | 提示用户用 `--no-preserve-input-format` 跳过（v3.16+） |
| `unrecognized arguments: --xwriter` | CLI 主入口未升级到 v3.19，需更新 `excel_style_cleaner.py` |

---

## 扩展定制

修改 `resources/excel_style_cleaner.py` 顶部常量：

```python
RATE_KEYWORDS = [...]      # 达成率/通过率/完成率类列标题关键字
MOMYOY_KEYWORDS = [...]    # 环比/同比类列标题关键字
MONEY_KEYWORDS = [...]     # 严格金额列（加 ¥ + 居右）
VOLUME_KEYWORDS = [...]    # 业务量列（按数字格式自动选）
TITLE_KEYWORDS = [...]     # 标题/合计行关键字
```

或修改 `config/keywords.yaml`（YAML 配置样例）。

修改后务必：
1. 同步更新本 SKILL.md 对应章节
2. 在 `logs/log.md` 追加一条更新记录
3. 运行 `references/_test_v317.py` 验证

---

## 外部资源

| 资源 | 路径 | 用途 | 何时加载 |
|------|------|------|---------|
| 主清洗脚本（openpyxl 版） | `resources/excel_style_cleaner.py` | 读取并清洗旧表 | 每次调用必读 |
| 新表写入（xlsxwriter 版） | `resources/excel_writer.py` | 纯写新表，BI / 批量出表 | `--xwriter` 时必读 |
| 端到端测试（v3.17 自检） | `references/_test_v317.py` | 验证 6 类格式自检 | 修改 v3.17 后必跑 |
| 端到端测试（基础） | `references/smoke_test.md` | 验证脚本可用性 | 修改后必跑 |
| 配置样例 | `config/keywords.yaml` | 关键字与颜色可配置化 | 二次定制时参考 |
| 更新日志 | `logs/log.md` | 全部版本变更记录 | 查阅历史时 |

---

## 版本变更

> **所有版本变更记录已迁移至 `logs/log.md`**。本节仅列出版本概览索引。

- **v3.54 (2026-09-20)**：用户反馈 sheet"sheet_1_2 col_<biz_alias_5>行使用 0.0万 格式"+"均值非≥1万的数字列若整列都是整数则用整数格式"。重设 apply_number_formats 的 cell 循环：① 新增 `_column_all_integer` 辅助函数——判定列是否所有非空数字 cell 都是整数；② 优先级 5 普通数字列判定：均值 < 1万 + 整列整数 → 走整数格式 `#,##0`，否则 `#,##0.00`；③ 新增**行级指标标签覆盖列级格式**：在 cell 循环里读 B 列（指标列 = max(header_cols)）的标签——若命中 INTEGER_COUNT_KEYWORDS（客户数/会员数/账号数/订单数/链接数/代理商数/合同数/人数），该 cell 走整数格式（覆盖列级"万级/普通"判定）。
  - 例：sheet"sheet_1_2" R17 B17='col_<biz_alias_5>'（label）→ C17~O17 走整数（而非整列"0.0万"）
  - 例：sheet"4" col 171 R2='col_<biz_alias_3>' → R4=656 走整数 `#,##0`
  - 例：sheet"4" col 185 R2='col_<biz_alias_4>' → R4=3096 走整数 `#,##0`
  - 列级 `is_integer_count_column(last_header)` 判定**不依赖 all_layers**——避免 R2='45809'(日期)+all_layers 含 'col_<biz_alias_5>' 时把整列误判为整数（导致 GMV 数据走整数格式）
- **v3.53 (2026-09-17)**：用户反馈"行表头处理应该先横后竖，列表头处理先竖后横"，并指出 Phase 2.5 不是冗余（处理"漏检的行维度列"）。重设计：① Phase 1 改为**先横后竖**——避免 R2 D2:D3 重复值竖向合冲突 D2:AS2 横合；② Phase 2 占位跨 header_cols 边界保护（仅 a_col_has_label=True 且 header_col=1 时允许占位延伸）——修复 sheet"sheet_1_2" R7 B7='付款GMV' + C7='-' 错误合并 B7:C7；③ merge_header_rectangles 加 `is_in_other_merge` 排除独立合并干扰，并 unmerge 重叠 merge（修复 A1:C1 + A1:A4 合并后 openpyxl 留两个独立 merge 的 bug）；④ 保留 v3.52 的 Phase 2.5 跳过逻辑（A 列真值 == cell_val 时跳过 c≠1 源格横向合）。
- **v3.52 (2026-09-17)**：① Phase 2.5 跳过条件从 v3.51 的"a_col_has_label=True 时跳过所有 c≠1 源格"精确为"A 列真值 == cell_val 时跳过 c≠1 源格"——防止 v3.51 误伤 sheet"2" R13（A13='<region_B>' vs B13='<region_sub>' 不同值）的 B13:C13 横合。
- **v3.51 (2026-09-17)**：① Phase 2 列表头处理 col=1（A 列）例外不跳过——A 列是"行维度列"，整列所有行都是 header_cells 源格；防止 v3.50 修复 Q1 时误伤 sheet"2" R5 A5='<biz_unit>' + B5='-' + C5='-' 应合 A5:C5 的场景；② Phase 2.5 处理数据行源格时，如果该行 A 列有自己的标签（a_col_has_label=True），跳过 c≠1 源格的横向合——让 Phase 2 统一处理 A 列展开（避免 sheet"2" R6 B6:C6 + A6:C6 两个独立 merge 共存）。
- **v3.50 (2026-09-17)**：① Phase 1 表头行竖向合禁用"重复值延伸"——防止 sheet"2" D2:D3 重复值竖向合导致 D2 横合 D2:AS2 被跳过、最终 merge_header_rectangles 扩展成 D2:Q3 越界；② Phase 2 列表头处理跳过 `r ∉ header_rows` 的源格——防止 sheet"sheet_1_2" R7 B7='付款GMV'（数据行的"行维度标签"）被当作表头源格参与 B7:C7 横合（C7 是数据）。
- **v3.49 (2026-09-17)**：① 竖向合并支持"重复值延伸"——Phase 2 竖向合 next_val 允许空/占位/与源格同值，修复 sheet"sheet_1_2" B 列 R2:R16 重复值 '入账GMV' 应当竖向合并为 B2:B6 + B8:B16 + B17:B21 的需求。
- **v3.48 (2026-09-17)**：① Phase 2 横向合引入 `a_col_has_label` 判定（该行 A 列有自己的标签 → 允许占位延伸）；② 排除"合计/总计/小计/汇总/标题"行源格的占位延伸（防止 sheet 0 R9 B9='合计' 越界合到 BE9）；③ `apply_rate_data_bar` 加"率列优先"判定——即使含 GMV 关键字，只要 `is_percent_column(last_header)=True` 就按率列处理（修复 sheet"4" 'col_<biz_alias_2>...' 缺 Data Bar）；④ `fill_empty_content_cells` 加 try-except 静默跳过合并副格（Phase 2 后跑时的 AttributeError）；⑤ 新增 Phase 2.5——遍历数据行内 c ∉ header_cols 的源格横向合（仅严格同值合并）。
- **v3.47 (2026-09-16)**：① Phase 2 横向合按"源格行位置 + next_val 语义"三维判定——表头行源格不受 `max(header_cols)` 限制（合到 max_col）；数据行源格仅当 next_val 与源格严格同值才允许跨 header_cols 边界；② 解决 0/45/46 误判（v3.46 删除限制后导致 B6=Q6 这种数据行横向合越界）。
- **v3.46 (2026-09-16)**：① Phase 2 横向合删除 `max(header_cols)` 限制——sheet"0" AS2='col_<biz_alias_7>）' 应合到 AS2:AY3；sheet"3.2" A7:C7='<region_B>' 应合（表头单元格重复值合并）；② `RATE_KEYWORDS` 加"留存率"——修复 sheet"4" 'col_<biz_alias_2>...' 误走万级；③ `PERCENT_NO_BAR_KEYWORDS` 扩充为含"费率/汇率/分润率/分佣率/折算率"——这 5 个关键字加 Data Bar 不加进度条；④ `apply_rate_data_bar` 加 `is_percent_no_bar_column` 拦截。
- **v3.45 (2026-09-16)**：① `merge_header_by_rows` Phase 1 横向合右边界用 `max_col` 而非 `max(header_cols)`——sheet"0" R2/R3 横向合不再受 v3.39 加的"next_c ≤ max(header_cols)"保护误伤；② `bold_header_and_total_rows` 删除 `is_first_after_header` 限制——任何 ≥2 列横向合并行整行加粗；③ 新增 `_extend_to_nearest_nonempty_left_only` 辅助函数——仅向左搜 max_search 列，不右扩展，防止跨业务组边界污染 `apply_number_formats` 的 date fallback；④ `apply_number_formats` 在 `apply_rate_data_bar` 之前调用 `_get_header_cells_in_col` 都改用 left-only 变体；⑤ 新增 `MONEY_LIKE_KEYWORDS` 白名单——命中后强制走万级，不依赖均值阈值（修复 sheet"2" AH 列（上层='col_<biz_alias_1>'，均值仅 8645.79）误判为普通数字）。
- **v3.44 (2026-09-14)**：① `is_date_like` 识别季度字符串（"26年Q1" 等），修复 sheet"0 达成率"列 fmt 错误；② Phase 2 横向合不再受 `phase1_prev_cover` 约束，修复 sheet"2" row 10 A10:C10 重复值合并只到 A:B 而非 A:C 的 bug。
- **v3.43 (2026-09-14)**：① **表头"重复值合并"**——新增 `_can_extend_for_repeat_value`，改造 `merge_header_by_rows` 的 4 处扩展判断；② 占位符 `'-` / 空 / 数字维持原逻辑。
- **v3.42 (2026-09-14)**：① `is_date_like` 识别 2 位年份字符串（`\d{2}\s*年\s*\d{1,2}\s*月`），修复 `format_header_dates` 格式化后 date fallback 失效 → 整数列错配普通数字格式的 bug；② `apply_rate_data_bar` / `apply_momyoy_data_bar` 互斥——前者跳过 `ws._momyoy_cols` 标记列，并主动检查 `is_momyoy_in_header_cells` 兜底，避免环比/同比列被双重条件格式叠加。
- **v3.41 (2026-09-14)**：① `apply_number_formats` 统一为 5 优先级判定；② 最后一层 header 判定 + date fallback；③ GMV 关键字扫描严格化（排除"GMV考核"等子模块名）；④ 排除 GMV 列加进度条。
- **v3.40 (2026-09-14)**：① `_preserve_sparklines_postsave` 保留迷你图（`<x14:sparklineGroups>` + `<ext>` wrapper + 命名空间）；② GMV 列强制万级格式。
- **v3.39 (2026-09-14)**：① B6:C6 不应合并修复；② `_remove_ai_drawings_from_workbook` 加载后立即清空 drawings；③ `_strip_ai_artifacts_postsave` 清理 drawings/media/charts 浮层对象。
- **v3.30 (2026-09-11)**：① `merge_header_by_rows` 入参改为 `header_cells` 集合 + 三段优先级分类（Phase 1 行表头先行再列，Phase 2 列表头先列再行）；② 关键修复——Phase 1 源格本身已被横向合（min_col==c）跳过竖向合；Phase 2 列表头竖向合时检查 next_r 同行横向合状态；③ 源格值改用 `ws.cell.value`（不被跨行合并覆盖判空影响）；④ 恢复 v3.22 prev_row_col_cover 约束；⑤ 关键场景验证：<biz_report> sheet1 → A5:A8/A9:A11/A12:A14 + A15:B15 + F3:G3 + AS2:AY3 (via AS1:BE1+AS1:AS3 合并路径)。
- **v3.29 (2026-09-11)**：① `merge_header_by_rows` 入参改为 `header_cells` 集合；② 修复 dangling externalLink 引用 Bug；③ 保留 `merge_header_by_legacy_rows` 旧接口。
- **v3.28 (2026-09-11)**：① `auto_fit_columns` 第一阶段 `horizontal_merges` 收集新增 `MAX_HEADER_TITLE_SPAN=6` 过滤——跳过跨度 >6 列的横向合并（如 `C1:AR1` 这类覆盖整片区域的"总标题"）；② 根因：原算法把总标题源格按 span 均分字符，每列 max_w ≈ 1，列宽塌回 `MIN_COL_WIDTH=8`；③ 修复后数据列按真实内容长度计算；表头宽度仍由第三步（标题行 wrap 扩展）按比例处理；④ **条件格式触发条件改为"列的所有表头单元格中任意一个含关键字"**——新增 `is_rate_in_header_cells` / `is_momyoy_in_header_cells` 辅助函数；`apply_rate_data_bar` / `apply_momyoy_data_bar` 改用新判定（不再依赖 `_build_multi_layer_headers` 拼接字符串）；⑤ **删除"严格金额列（含 金额/收入/支出/回款/定价）"触发条件**——`apply_money_alignment` 改空实现；`apply_number_formats` 内金额列居右删除。
- **v3.27 (2026-09-11)**：① 去除行列定义改为单元格级定义——`detect_layout` 返回 `(header_rows, header_cols, data_rows, data_cols, header_cells, data_cells)` 六元组；② 表头列扫描范围改为 `[1, max(header_rows)]`，逐列独立判定（非 break 终止）；③ 下游5 函数新增 `header_cells` 参数，仅对数据单元格应用样式（数字格式、Data Bar、居右、`-` 填充）
- **v3.22 (2026-09-10)**：① 表头合并算法重构为「逐行纵→横」(`merge_header_by_rows`)；② 新增「竖向优先」规则（已竖向合并的源格不再横向合并）；③ 新增「下一行合并不超上一行」规则（`prev_row_col_cover[c]` 列范围限制）
- **v3.21 (2026-09-07)**：① 新增 `merge_header_rectangles` 矩形合并 pass2（A1 有值 / A2/B1/B2 全空 → 合并为 A1:B2）；② `RATE_KEYWORDS` 删除「达成/达标/通过/完成」4 个单字（避免与「完成日期/目标达成次数」等列名误匹配）；③ `_AI_ARTIFACT_PARTS` 扩展 customXml/queryTable/connections/externalLinks/vbaProject
- **v3.20 (2026-09-07)**：① 万级底层值 /10000 + fmt `0.0万`（替代 v3.5-v3.6 的"显示 1500000.0万"多位小数问题）；② 合计行不再被识别为表头（含 TITLE_KEYWORDS 的行强制排除）；③ `is_numeric_for_header` 增强识别字符串数字；④ `format_header_dates` 跳过合并区域副格
- **v3.19 (2026-09-06)**：找回 4 项遗漏规则（多层表头合并 / 多层表头月份格式 / 合并范围 ≥2 列加粗 / RATE_KEYWORDS 扩展）；阶段四新增 `_strip_ai_artifacts_postsave` 清除 AI/插件生成的浮层对象（drawings/comments/ctrlProps/activeX/embeddings/media/charts/pivotTables/slicers/tables）
- **v3.18 (2026-09-06)**：阶段四清理 AI/工具生成的水印层（docProps/core.xml + docProps/app.xml）
- **v3.17 (2026-09-04)**：自检扩展覆盖千分位 / 短数字 / 负数红 / 自定义格式 6 类规则
- **v3.16 (2026-09-02)**：xwriter 模式沿用 input number_format + 自检规则 + 列宽增强
- **v3.15 (2026-09-01)**：表头识别阈值暴露为 `--header-threshold` / `--header-max-scan`
- **v3.14 (2026-09-01)**：xwriter 模式支持自动表头识别
- **v3.13 (2026-09-01)**：表头识别放宽到"连续 N 行含数字"
- **v3.12 (2026-08-31)**：打通 openpyxl → xlsxwriter
- **v3.11 (2026-08-30)**：百分比负数红 + 进度条确认
- **v3.10 (2026-08-30)**：列宽格式后缀追加
- **v3.9 (2026-08-30)**：去除 AI 输出水印
- **v3.8 (2026-08-30)**：Excel 双击自适配列宽 + 行高
- **v3.7 (2026-08-30)**：行高自适应 + debug 模式 + 明细大表降级
- **v3.6 (2026-08-29)**：关键 bug 修复（万级 / 百分比 / 表头加粗）
- **v3.5 (2026-08-29)**：用户反馈修订 4 项
- **v3.4 (2026-08-29)**：用户反馈修订 5 项
- **v3.3 (2026-08-29)**：修复 3 个关键 bug
- **v3.2 (2026-08-29)**：用户反馈细化 6 项
- **v3.1 (2026-08-28)**：标题区 / 内容区分区处理
- **v3 (2026-08-28)**：用户反馈修订
- **v2 (2026-08-28)**：重大增强（11 项新规则）
- **v1 (2026-08-27)**：初始创建（5 类样式规则）

详细修订记录见 [logs/log.md](computer://<PROJECT_ROOT>\.agents\skills\excel-style-cleaner\logs\log.md)。
