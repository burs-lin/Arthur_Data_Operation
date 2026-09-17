"""单元测试:表头识别(`_detect_header_rows_from_rows` + `_values_are_similar`)。

覆盖以下场景:
  - 基础:数字列相近
  - 日期列相近
  - 会员号列相近
  - 邮箱列相近
  - URL 列相近
  - 手机号列相近
  - 3 行表头
  - 数据中间夹空行
  - 全空
  - 只有一行
  - 全是文字(无相近连续)
  - 首列说明文字被跳过
  - max_scan 截断
  - 弱相似
"""
import sys
sys.path.insert(0, r"D:/Trae_Work/Data_Operation/.agents/skills/excel-data-cleaner/scripts")
from typing import Any, List
import clean_excel as cec


def show(name: str, expected: int, rows: List[List[Any]], **kw):
    got = cec._detect_header_rows_from_rows(rows, **kw)
    status = "OK" if got == expected else "FAIL"
    print(f"[{status}] {name}: expected={expected} got={got}")
    assert got == expected, f"{name}: expected {expected} got {got}"


def test_basic_number():
    show("basic_number", 1, [
        ["订单号", "客户", "金额"],
        ["A001", "张三", 100],
        ["A002", "李四", 200],
    ])


def test_date_column():
    show("date_column", 1, [
        ["月份", "门店", "GMV"],
        ["2024-01", "北京店", 5000],
        ["2024-02", "深圳店", 6000],
    ])


def test_member_id():
    show("member_id", 1, [
        ["会员号", "姓名", "积分"],
        ["VIP001号", "张三", 100],
        ["VIP002号", "李四", 200],
    ])


def test_email_col():
    show("email_col", 1, [
        ["name", "email"],
        ["a", "a@x.com"],
        ["b", "b@x.com"],
    ])


def test_3row_header():
    show("3row_header", 3, [
        ["销售报表", None, None],
        ["区域", "门店", "GMV"],
        ["华北", "北京店", "合计"],
        ["华南", "深圳店", 500],
        ["华东", "上海店", 800],
    ])


def test_gap_in_data():
    show("gap_in_data", 1, [
        ["订单号", "金额"],
        ["A1", 100],
        [None, None],
        ["A2", 200],
    ])


def test_empty():
    show("empty", 1, [])


def test_single_row():
    show("single_row", 1, [
        ["a", "b"],
    ])


def test_all_text():
    show("all_text", 1, [
        ["a", "b"],
        ["c", "d"],
        ["e", "f"],
        ["g", "h"],
    ])


def test_note_skipped():
    show("note_skipped", 1, [
        ["请查看报表说明,数据从第 2 行开始", "", ""],
        ["月份", "门店", "GMV"],
        ["2024-01", "北京店", 5000],
        ["2024-02", "深圳店", 6000],
    ])


def test_max_scan():
    show("max_scan", 1, [
        ["a", "b"],
        ["c", "d"],
        ["e", "f"],
        [1, 2],
        [3, 4],
    ], max_scan=3)


def test_url_col():
    show("url_col", 1, [
        ["name", "url"],
        ["a", "https://a.com"],
        ["b", "https://b.com"],
    ])


def test_phone_col():
    show("phone_col", 1, [
        ["姓名", "手机号"],
        ["张三", "13800138000"],
        ["李四", "13900139000"],
    ])


def test_weak_similar():
    show("weak_similar", 1, [
        ["code", "value"],
        ["AB12", 100],
        ["AB34", 200],
    ])


if __name__ == "__main__":
    test_basic_number()
    test_date_column()
    test_member_id()
    test_email_col()
    test_3row_header()
    test_gap_in_data()
    test_empty()
    test_single_row()
    test_all_text()
    test_note_skipped()
    test_max_scan()
    test_url_col()
    test_phone_col()
    test_weak_similar()
    print("\nALL 14 HEADER DETECTION TESTS PASSED")