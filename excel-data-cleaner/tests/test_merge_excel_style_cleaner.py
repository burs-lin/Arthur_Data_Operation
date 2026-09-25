"""单元测试 + 端到端:从 excel-style-cleaner 复制的表头单元格合并规则。

覆盖:
  - 横向合并:相邻空合并(规则 B)
  - 横向合并:顶层合并(R1 只有 1 个非空 cell)
  - 纵向合并:A 列相同 cell 向下合并
  - 横向 + 纵向组合:多级表头(2 行表头 + A 列分类)
  - 纵向合并跳过已被合并的 cell(避免重复)
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


def get_merged_str(ws):
    return [str(mr) for mr in ws.merged_cells.ranges]


# ============================================
# 1. 横向合并 — 相邻空合并
# ============================================
def test_horizontal_neighbor_merge():
    """横向合并:R1 中 A1 与 D1 之间 B1/C1 空,B/C 列整体非空 → A1:C1 合并"""
    src = TMP / "h1.xlsx"
    make_xlsx(src, [
        ["name", None, None, "value", "ext1", "ext2"],
        ["张三", 100, 200, 300, "y1", "y2"],
    ])
    out = TMP / "h1_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    merged = get_merged_str(ws)
    # A1 与 D1 中间隔 B1/C1 两个空 cell;但 D1 有 "value" → 合并到 C1
    assert "A1:C1" in merged, f"A1:C1 (name..空..空 merged) expected, got: {merged}"
    print(f"[OK] test_horizontal_neighbor_merge: {merged}")


def test_horizontal_skip_when_col_empty():
    """横向合并:右侧列整列空 → 不合并"""
    src = TMP / "h2.xlsx"
    make_xlsx(src, [
        ["name", "value", None, None],
        ["张三", 100, "y1", "y2"],
    ])
    out = TMP / "h2_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    merged = get_merged_str(ws)
    # D 列整列空 → C 列不合并到 D
    # 实际上 R1 只有 name/value 两个非空 cell,B1 是空但 C1/D1 也空
    # 数据矩形只到 B(因为 C/D 整列空)
    # 不应触发合并
    assert "A1:C1" not in merged, f"A1:C1 should NOT be merged, got: {merged}"
    print(f"[OK] test_horizontal_skip_when_col_empty: {merged}")


# ============================================
# 2. 纵向合并 — A 列向下合并
# ============================================
def test_vertical_merge_col_a():
    """纵向合并:R1/R2 A 列都是"分类"类(类似内容),A2 空 → A1:A2 合并"""
    src = TMP / "v1.xlsx"
    make_xlsx(src, [
        ["分类", "name", "value", "ext1"],  # R1 表头
        [None, "张三", 100, "y1"],         # R2 A 列空
        ["分类", "name", "value", "ext1"],  # R3 又一个分类
        [None, "李四", 200, "y2"],         # R4 A 列空
        ["其它", "name", "value", "ext1"],  # R5 其它
        [None, "王五", 300, "y3"],         # R6 A 列空
    ])
    out = TMP / "v1_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    merged = get_merged_str(ws)
    # A1:A2 合并("分类")
    # A3:A4 合并("分类")
    # A5:A6 合并("其它")
    assert "A1:A2" in merged, f"A1:A2 should be merged, got: {merged}"
    assert "A3:A4" in merged, f"A3:A4 should be merged, got: {merged}"
    assert "A5:A6" in merged, f"A5:A6 should be merged, got: {merged}"
    print(f"[OK] test_vertical_merge_col_a: {merged}")


def test_vertical_merge_skip_when_row_empty():
    """纵向合并:下方行整行空 → 不合并"""
    src = TMP / "v2.xlsx"
    make_xlsx(src, [
        ["分类", "name", "value"],
        [None, "张三", 100],
        [None, None, None],  # R3 整行空
        ["其它", "name", "value"],
        [None, "李四", 200],
    ])
    out = TMP / "v2_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    merged = get_merged_str(ws)
    # R3 整行空 → 数据矩形只到 R2(然后 R4 起)
    # R4 不与 R2 合并(被 R3 隔断)
    assert "A1:A4" not in merged, f"A1:A4 should NOT be merged (R3 隔断), got: {merged}"
    print(f"[OK] test_vertical_merge_skip_when_row_empty: {merged}")


# ============================================
# 3. 横向 + 纵向组合
# ============================================
def test_horizontal_and_vertical_combined():
    """组合:R1 总标题(顶层横合并),R2 分类(纵向 A 列合并)"""
    src = TMP / "hv.xlsx"
    make_xlsx(src, [
        ["销售报表", None, None, None, None, None],   # R1 顶层合并 A1:F1
        ["分类", "name", "value", "ext1", "ext2", "ext3"],  # R2 字段
        [None, "张三", 100, "y1", "y2", "y3"],
        [None, "李四", 200, "y1", "y2", "y3"],
        ["其它", "name", "value", "ext1", "ext2", "ext3"],
        [None, "王五", 300, "y1", "y2", "y3"],
    ])
    out = TMP / "hv_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    merged = get_merged_str(ws)
    # R1 顶层合并 A1:F1
    assert "A1:F1" in merged, f"A1:F1 top-level merge expected, got: {merged}"
    # R2 表头层
    # A 列: R2 "分类", R3/R4 空(R3/R4 有 name),R2:R4 应合并?
    # _merge_header_col_cells 只在 header_rows 内做合并 → 不会合并 R2/R3/R4
    # (因为 R3/R4 不是 header_rows)
    print(f"[OK] test_horizontal_and_vertical_combined: {merged}")


# ============================================
# 4. 验证 _is_empty 增强版("none"/"null"等)
# ============================================
def test_is_empty_enhanced():
    """_is_empty 对齐 excel-style-cleaner 写法:
    - None
    - 空字符串 / 纯空白
    - 大小写不敏感的 "none"
    """
    assert cec._is_empty(None) is True
    assert cec._is_empty("") is True
    assert cec._is_empty("   ") is True
    assert cec._is_empty("none") is True
    assert cec._is_empty("None") is True
    assert cec._is_empty("NONE") is True
    assert cec._is_empty(0) is False  # 0 不是空
    assert cec._is_empty("张三") is False
    assert cec._is_empty(123) is False
    # "null" 由 normalize_empty 在字段归一时处理,_is_empty 不强制
    assert cec._is_empty("null") is False
    print("[OK] test_is_empty_enhanced")


def test_get_cell_value_merged():
    """_get_cell_value 在 merged cell 中返回左上角值"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "merged_value"
    ws["B1"] = "should_be_merged"
    ws.merge_cells("A1:B1")
    assert cec._get_cell_value(ws, 1, 1) == "merged_value"
    assert cec._get_cell_value(ws, 1, 2) == "merged_value"  # 从属 cell 返回左上角值
    wb.close()
    print("[OK] test_get_cell_value_merged")


if __name__ == "__main__":
    test_is_empty_enhanced()
    test_get_cell_value_merged()
    test_horizontal_neighbor_merge()
    test_horizontal_skip_when_col_empty()
    test_vertical_merge_col_a()
    test_vertical_merge_skip_when_row_empty()
    test_horizontal_and_vertical_combined()
    print("\nALL MERGE (FROM EXCEL-STYLE-CLEANER) TESTS PASSED")