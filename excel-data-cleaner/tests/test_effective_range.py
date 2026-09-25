"""单元测试:有效行 / 有效列 / 有效范围识别 + 行高自适应。

覆盖以下场景:
  - 有效行/有效列基础识别
  - 整行为空 / 整列为空被排除
  - 有效范围(首/末)
  - 行高自适应(列宽换行后的行数估算)
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
from openpyxl import Workbook, load_workbook
import clean_excel as cec


# ============================================
# 1. 有效行/有效列/有效范围识别
# ============================================
def test_effective_basic():
    """基础:全填满的 sheet"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "a"
    ws["B1"] = "b"
    ws["A2"] = "c"
    ws["B2"] = "d"
    assert cec._find_empty_full_rows(ws) == set()
    assert cec._find_empty_full_cols(ws) == set()
    assert cec._find_effective_rows(ws) == [1, 2]
    assert cec._find_effective_cols(ws) == [1, 2]
    assert cec._find_effective_range(ws) == (1, 2, 1, 2)
    print("[OK] test_effective_basic")


def test_effective_with_empty_full_row():
    """整行为空(1:1)的行被排除"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "a"
    ws["B1"] = "b"
    # 第 2 行全空
    ws["A3"] = "c"
    ws["B3"] = "d"
    assert cec._find_empty_full_rows(ws) == {2}
    assert cec._find_effective_rows(ws) == [1, 3]
    assert cec._find_effective_range(ws) == (1, 3, 1, 2)
    print("[OK] test_effective_with_empty_full_row")


def test_effective_with_empty_full_col():
    """整列为空(A:A)的列被排除"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "a"
    ws["C1"] = "c"
    ws["A2"] = "x"
    ws["C2"] = "y"
    assert cec._find_empty_full_cols(ws) == {2}
    assert cec._find_effective_cols(ws) == [1, 3]
    assert cec._find_effective_range(ws) == (1, 2, 1, 3)
    print("[OK] test_effective_with_empty_full_col")


def test_effective_whitespace_only():
    """纯空白单元格(如" ")视为空"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "a"
    ws["B1"] = " "  # 纯空白
    ws["C1"] = "b"
    assert 1 not in cec._find_empty_full_rows(ws), "R1 还有 A1='a' 和 C1='b',不应被认作空"
    print("[OK] test_effective_whitespace_only")


# ============================================
# 2. 行高自适应
# ============================================
def test_row_height_basic():
    """行高自适应:短内容 → 行高 = base_line_height"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "短"
    ws["B1"] = "也短"
    cec._auto_fit_columns(ws, 1, 2, header_rows=0)
    cec._auto_fit_row_heights(ws, 1, 2, header_rows=0)
    h = ws.row_dimensions[1].height
    assert h is not None and h >= 15.0, f"row height too small: {h}"
    print(f"[OK] test_row_height_basic (h={h})")


def test_row_height_long_content():
    """行高自适应:长内容需要换行 → 行高增大"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "x" * 200  # 远超过 max_width=40
    cec._auto_fit_columns(ws, 1, 1, header_rows=0)
    cec._auto_fit_row_heights(ws, 1, 1, header_rows=0)
    h = ws.row_dimensions[1].height
    assert h > 15.0, f"long content should expand row height, got h={h}"
    print(f"[OK] test_row_height_long_content (h={h})")


def test_row_height_via_e2e():
    """端到端:clean_file 后行高被设置"""
    src = Path(__file__).resolve().parent / "_tmp" / 't_rh.xlsx'
    src.parent.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "姓名"
    ws["B1"] = "订单内容"
    ws["A2"] = "张三"
    ws["B2"] = "超长的订单描述 " * 20
    wb.save(src)
    out = src.with_name(src.stem + "_out.xlsx")
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    h1 = ws.row_dimensions[2].height
    assert h1 is not None and h1 > 15.0, f"row 2 should be expanded, got {h1}"
    print(f"[OK] test_row_height_via_e2e (R2 height={h1})")


if __name__ == "__main__":
    test_effective_basic()
    test_effective_with_empty_full_row()
    test_effective_with_empty_full_col()
    test_effective_whitespace_only()
    test_row_height_basic()
    test_row_height_long_content()
    test_row_height_via_e2e()
    print("\nALL EFFECTIVE RANGE + ROW HEIGHT TESTS PASSED")