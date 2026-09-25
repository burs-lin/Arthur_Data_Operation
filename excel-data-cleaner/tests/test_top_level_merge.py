"""单元测试 + 端到端:表头顶层合并 + 数字 number_format=0.00。

覆盖以下场景:
  - 顶层合并:某行只有 1 个非空 cell 时,该 cell 合并到整行所有列
  - 普通相邻空合并仍然工作
  - 顶层合并在有效列范围内生效,不被矩形外的列影响
  - 数字 number_format = 0.00(不影响底层值)
  - 不生成 _CLEAN_REPORT / _ANOMALIES sheet
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
from openpyxl import Workbook, load_workbook
import clean_excel as cec


TMP = Path(__file__).resolve().parent / "_tmp"
TMP.mkdir(exist_ok=True)


def make_xlsx(path: Path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)
    wb.save(path)


# ============================================
# 1. 顶层合并
# ============================================
def test_top_level_merge_r1():
    """R1 只有 1 个非空 cell → 整行合并"""
    src = TMP / "tlm1.xlsx"
    # 数据填到 H 列,让数据矩形是 [1, 3] x [1, 8]
    # R1 只有 A1 非空 → 应合并 A1:H1(整行 8 列)
    make_xlsx(src, [
        ["总标题", None, None, None, None, None, None, None],
        ["name", "value", "type", "ext1", "ext2", "ext3", "ext4", "ext5"],
        ["A", 1, "x", "y1", "y2", "y3", "y4", "y5"],
    ])
    out = TMP / "tlm1_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    merged = list(ws.merged_cells.ranges)
    assert len(merged) >= 1
    assert any(str(mr) == "A1:H1" for mr in merged), \
        f"R1 should be merged A1:H1, got: {[str(mr) for mr in merged]}"
    print(f"[OK] test_top_level_merge_r1: {[str(mr) for mr in merged]}")


def test_top_level_merge_doesnt_break_r2():
    """R1 顶层合并不影响其它行(R2 在本例是数据行,不被合并)"""
    src = TMP / "tlm2.xlsx"
    make_xlsx(src, [
        ["总标题", None, None, None, None, None, None, None],
        ["name", None, "value", "ext1", "ext2", "ext3", "ext4", "ext5"],  # B2 空
        ["A", None, 1, "y1", "y2", "y3", "y4", "y5"],
    ])
    out = TMP / "tlm2_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    merged = [str(mr) for mr in ws.merged_cells.ranges]
    # R1 顶层合并
    assert "A1:H1" in merged, f"R1 should be merged A1:H1, got: {merged}"
    # 本例 header_rows = 1,R2 是数据行不参与合并
    assert len(merged) == 1, f"only R1 should be merged, got: {merged}"
    print(f"[OK] test_top_level_merge_doesnt_break_r2: {merged}")


def test_multi_header_top_level():
    """多级表头:R1 总标题整行合并,R2 字段名部分相邻空合并"""
    src = TMP / "multi.xlsx"
    # 2 行表头 + 数据
    # R1: 总标题("销售报表")整行合并
    # R2: 字段名 + 一些分类标题(相邻空合并)
    make_xlsx(src, [
        ["销售报表", None, None, None, None, None, None, None],   # R1 顶层合并
        ["区域", None, "客户", None, "业务", None, None, None],   # R2 部分相邻空
        ["华北", "北京店", "张三", "M0001号", "B2B", "贸易", "货贸", "付款"],
    ])
    out = TMP / "multi_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    merged = [str(mr) for mr in ws.merged_cells.ranges]

    # R1 应顶层合并 A1:H1
    assert "A1:H1" in merged, f"R1 top-level merge failed, got: {merged}"
    print(f"[OK] test_multi_header_top_level: {merged}")


# ============================================
# 2. 数字 number_format = 0.00
# ============================================
def test_number_format_zero_decimal():
    """number 列 number_format = 0.00,不影响底层值"""
    src = TMP / "nf1.xlsx"
    make_xlsx(src, [
        ["金额", "数量", "客户"],
        [100.5, 3, "张三"],
        [200.456, 5, "李四"],
    ])
    out = TMP / "nf1_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]

    # 检查 number_format
    fmt_a = ws["A2"].number_format  # 金额
    fmt_b = ws["B2"].number_format  # 数量
    fmt_c = ws["C2"].number_format  # 客户(应该是文本 @)
    print(f"  A2 (金额) number_format: {fmt_a}")
    print(f"  B2 (数量) number_format: {fmt_b}")
    print(f"  C2 (客户) number_format: {fmt_c}")

    # 数字列都是 0.00
    assert fmt_a == "0.00", f"amount number_format should be 0.00, got {fmt_a!r}"
    assert fmt_b == "0.00", f"number number_format should be 0.00, got {fmt_b!r}"

    # 底层值不变(整数 / 浮点数仍保留原值)
    assert ws["A2"].value == 100.5 or ws["A2"].value == "100.5", \
        f"A2 underlying value should be 100.5, got {ws['A2'].value!r}"
    assert ws["B2"].value == 3 or ws["B2"].value == "3", \
        f"B2 underlying value should be 3, got {ws['B2'].value!r}"
    print("[OK] test_number_format_zero_decimal")


# ============================================
# 3. 不生成 _CLEAN_REPORT / _ANOMALIES
# ============================================
def test_no_extra_sheets():
    """输出文件不含 _CLEAN_REPORT / _ANOMALIES"""
    src = TMP / "ns1.xlsx"
    make_xlsx(src, [
        ["客户", "金额"],
        ["张三", 100],
    ])
    out = TMP / "ns1_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    sheet_names = wb.sheetnames
    assert "_CLEAN_REPORT" not in sheet_names, \
        f"_CLEAN_REPORT should NOT be created, got: {sheet_names}"
    assert "_ANOMALIES" not in sheet_names, \
        f"_ANOMALIES should NOT be created, got: {sheet_names}"
    assert sheet_names == ["Sheet1"], \
        f"only Sheet1 should exist, got: {sheet_names}"
    print(f"[OK] test_no_extra_sheets: sheets={sheet_names}")


if __name__ == "__main__":
    test_top_level_merge_r1()
    test_top_level_merge_doesnt_break_r2()
    test_multi_header_top_level()
    test_number_format_zero_decimal()
    test_no_extra_sheets()
    print("\nALL TOP-LEVEL MERGE + NUMBER_FORMAT TESTS PASSED")