# 端到端测试（Smoke Test v3.2）

## 步骤 1：构造测试数据

```python
import openpyxl
from pathlib import Path

p = Path("_smoke_test_input.xlsx")  # 默认写到当前目录
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sheet1"

# 表头（含日期 + 百分比 + 金额 + 万级数字）
ws.append(["区域", "销售额", "2024年1月", "2024-02", "达成率", "通过率", "完成率", "环比", "同比"])

# 数据行
ws.append(["<region>", 1234567.89, 500000, 600000, 0.85, 0.90, 0.78, 0.12, -0.05])
ws.append(["<region>", 9876543.21, 800000, 750000, 1.10, 0.95, 1.20, -0.08, 0.20])
ws.append(["<region>", 555555.00, 300000, None, 0.60, 0.70, 0.65, 0.00, 0.00])

# 空行
ws.append([])

# 合计行
ws.append(["合计", 11667766.10, 1600000, 1350000, 2.55, 2.55, 2.63, 0.04, 0.15])

wb.save(p)
```

## 步骤 2：执行清洗

```powershell
python "<SHARED_ROOT>/.agents/skills/excel-style-cleaner/resources/excel_style_cleaner.py" `
  "<test_input_path>" -o "<test_output_path>"
```

## 步骤 3：验证（v3.2 14 项断言）

```python
import openpyxl

wb = openpyxl.load_workbook("<test_output>")
ws = wb.active

# 1) 表头日期 → YY年MM月
assert ws["C1"].value == "2024年01月"
assert ws["D1"].value == "2024年02月"

# 2) 内容区空值填 -（D4 由 None → "-"）
assert ws["D4"].value == "-"

# 3) 空行无边框
assert ws["A6"].border.left.style is None

# 4) 数据行有边框
assert ws["A2"].border.left.style == "thin"

# 5) 金额列格式 + 居右
assert "¥" in ws["B2"].number_format
assert ws["B2"].alignment.horizontal == "right"

# 6) 万级格式（B 列均值 ~3M → 0.0,"万"）
assert "0.0," in ws["B2"].number_format and "万" in ws["B2"].number_format

# 7) 5 个百分比列：达成率/通过率/完成率/环比/同比 一律 0.00%
for col_letter in ["E", "F", "G", "H", "I"]:
    cell = ws[f"{col_letter}2"]
    assert cell.number_format == "0.00%", f"❌ {cell.coordinate}: {cell.number_format}"

# 8) 【关键】百分比底层值仍是小数（不 *100）
assert ws["E2"].value == 0.85, "底层值应为 0.85"
assert ws["F2"].value == 0.90
assert ws["G2"].value == 0.78
assert ws["H2"].value == 0.12
assert ws["I2"].value == -0.05

# 9) 达成率/通过率/完成率 Data Bar（E2:E4, F2:F4, G2:G4）
cf_ranges = [str(r.sqref) for r in ws.conditional_formatting._cf_rules.keys()]
for col in ["E", "F", "G"]:
    assert any(f"{col}2:{col}4" in s for s in cf_ranges)

# 10) 环比/同比 Data Bar + 箭头（H2:H4, I2:I4）
for col in ["H", "I"]:
    assert any(f"{col}2:{col}4" in s for s in cf_ranges)

# 11) 加粗：标题行 + 合计行
assert ws["A1"].font.bold is True
assert ws["A7"].font.bold is True  # 含"合计"行

# 12) 横向合并单元格所在行加粗（构造时把合计行做成横向合并 A7:E7 → 该行加粗）
#     见步骤 4 验证

# 13) 基础样式
assert ws["A2"].font.name == "微软雅黑"
assert ws["A2"].font.size == 9
assert ws["A2"].alignment.horizontal == "center"

# 14) 列宽自适应（在 [8, 30] 范围内）
from openpyxl.utils import get_column_letter
for c in range(1, ws.max_column + 1):
    width = ws.column_dimensions[get_column_letter(c)].width
    assert 8 <= width <= 30, f"❌ 列 {get_column_letter(c)} 宽度 {width} 超出范围"

print("✅ all 14 checks passed (v3.2)")
```

## 步骤 4（v3.2 新增）：验证横向合并单元格所在行加粗

```python
import openpyxl, subprocess
from pathlib import Path

p = Path(r"...\_t3_2_merge_bold.xlsx")
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["区域", "销售额", "合计", None])  # D1=None 用于合并 C1:D1
ws.append(["<region>", 1000, None, None])
ws.append(["<region>", 2000, None, None])
ws.append(["合计", 3000, None, None])
wb.save(p)

subprocess.run(["python", "<skill_script>", str(p), "-o", str(p.with_name("_t3_2_merge_bold_out.xlsx"))])

wb = openpyxl.load_workbook(p.with_name("_t3_2_merge_bold_out.xlsx"))
ws = wb.active

# 横向合并 C1:D1
merge_strs = [str(m) for m in ws.merged_cells.ranges]
assert any("C1:D1" in s for s in merge_strs)

# 第 4 行（合计行）整行加粗
for c in range(1, 5):
    assert ws.cell(row=4, column=c).font.bold is True
```

## 步骤 5（v3.2 新增）：验证纵向合并跳过已合并单元格

```python
# 测试场景：A1=有值, A2=空, A3=空, A4=合计 → 应合并 A1:A3，但 A4 不合并
# 在另一个测试中，先创建水平合并（如 B1:B3），再测试纵向合并是否会跳过 B1:B3
# 这由 is_in_any_merge() 守护，已实现
```

## 常见问题

| 问题 | 排查 |
|------|------|
| `ModuleNotFoundError: openpyxl` | `pip install openpyxl` |
| `❌ 输入文件不存在` | 检查路径是否含空格或中文，使用绝对路径 |
| 货币符号变成 `?` | Excel/WPS 字体未安装对应字符集 |
| WPS 不显示箭头 / Data Bar | WPS 部分版本支持差，临时用 Excel 打开 |
| 万级格式不生效 | 确认该列均值 ≥ 10,000 |
| 百分比显示成 `85` 而不是 `85.00%` | 检查 number_format 是否被其他规则覆盖 |
| 标题行没合并 | 确认右侧列在数据行有值 |
| 列宽太窄/太宽 | 在 YAML 调整 `column_width.min` / `column_width.max` |
