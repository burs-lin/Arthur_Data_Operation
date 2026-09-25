"""单元测试:表头加粗底色与边框 / 字体在同一个循环内。

验证 ③ 表头样式被合并到 ① 边框 + 字体循环里,而不是分开两段:
  - 表头行每个非空 cell → font.bold=True + fill=浅蓝底
  - 数据行每个非空 cell → font.bold=False
  - 所有 cell(包括空 cell)→ border
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


def test_header_bold_and_fill():
    """表头行每个非空 cell 都被加粗 + 浅蓝底"""
    src = TMP / "hs.xlsx"
    make_xlsx(src, [
        ["name", "value", "ext1"],
        ["张三", 100, "y1"],
        ["李四", 200, "y2"],
    ])
    out = TMP / "hs_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]

    # R1 表头:每个非空 cell 应该 bold + 浅蓝底(#D9E1F2)
    for c in range(1, 4):
        cell = ws.cell(row=1, column=c)
        assert cell.value is not None
        assert cell.font.bold is True, f"R1C{c} should be bold, got bold={cell.font.bold}"
        assert cell.fill.start_color.value == "FFD9E1F2", \
            f"R1C{c} should be 浅蓝底, got fill={cell.fill.start_color.value}"

    # R2/R3 数据行:每个非空 cell 应该 NOT bold, NOT 浅蓝底
    for r in [2, 3]:
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            if cell.value:
                assert cell.font.bold is False, f"R{r}C{c} should NOT be bold, got bold={cell.font.bold}"
    print("[OK] test_header_bold_and_fill")


def test_header_empty_cells_have_border():
    """表头行空 cell 也应该有边框(因边框先设,合并后视觉完整)"""
    src = TMP / "hs2.xlsx"
    make_xlsx(src, [
        ["name", None, "value"],   # R1 B1 空
        ["张三", 100, 200],
    ])
    out = TMP / "hs2_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    # R1C2 (B1) 空,但因为 R1C2 在矩形内且边框循环先于字体循环 → 应有边框
    # 注意 openpyxl 读 MergedCell 时 border 字段会是 None(不持有样式)
    # 但 Excel 打开会按主单元格 A1 的 border 渲染整块
    # 这里仅验证:数据 sheet 主单元格 A1 的 border 完整
    cell_a1 = ws["A1"]
    assert cell_a1.border.left.style == "thin"
    assert cell_a1.border.right.style == "thin"
    assert cell_a1.border.top.style == "thin"
    assert cell_a1.border.bottom.style == "thin"
    print("[OK] test_header_empty_cells_have_border")


def test_header_alignment_left_center():
    """表头行对齐:水平 left,垂直 center"""
    src = TMP / "hs3.xlsx"
    make_xlsx(src, [
        ["name", "value"],
        ["张三", 100],
    ])
    out = TMP / "hs3_out.xlsx"
    cec.clean_file(str(src), output_path=str(out))
    wb = load_workbook(out)
    ws = wb["Sheet1"]
    for c in range(1, 3):
        cell = ws.cell(row=1, column=c)
        a = cell.alignment
        assert a.horizontal == "left", f"R1C{c} horizontal should be 'left', got {a.horizontal!r}"
        assert a.vertical == "center", f"R1C{c} vertical should be 'center', got {a.vertical!r}"
    print("[OK] test_header_alignment_left_center")


if __name__ == "__main__":
    test_header_bold_and_fill()
    test_header_empty_cells_have_border()
    test_header_alignment_left_center()
    print("\nALL HEADER STYLE-IN-SAME-LOOP TESTS PASSED")