"""v3.17 自检扩展测试：覆盖 6 种格式

修复：每个测试用例至少 3 行数据，避免「连续 2 行含数字」阈值未触发导致全表被当表头。
"""
import sys, subprocess, openpyxl, os
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent  # 当前 skill 根目录（同级 scripts/）
RES = SKILL / "scripts"
sys.path.insert(0, str(RES))

def make_input(fmt_cases, n_data_rows=5):
    """造测试表：1 行表头 + n_data_rows 行数据，第 2 列起按 fmt_cases 设置格式"""
    wb = openpyxl.Workbook()
    ws = wb.active
    # 表头
    ws["A1"] = "区域"
    for c, item in enumerate(fmt_cases, start=2):
        h, _, _ = item
        ws.cell(row=1, column=c, value=h)
    # 数据行（至少 5 行，确保 threshold=2 能触发）
    regions = ["<region>", "<region>", "<region>", "<region>", "<region>"]
    for r in range(n_data_rows):
        ws.cell(row=r + 2, column=1, value=regions[r % len(regions)])
        for c, item in enumerate(fmt_cases, start=2):
            _, sample_v, fmt = item
            # 每行值略有差异，避免都是同一个数
            v = sample_v * (1 + r * 0.05) if isinstance(sample_v, (int, float)) else sample_v
            if isinstance(sample_v, int):
                v = int(v)
            cell = ws.cell(row=r + 2, column=c, value=v)
            cell.number_format = fmt
    return wb

def run_case(name, fmt_cases, expect_pass=True, extra_args=None):
    INPUT = SKILL / "references" / ("_v317_" + name + ".xlsx")
    OUTPUT = SKILL / "references" / ("_v317_" + name + "_out.xlsx")
    wb = make_input(fmt_cases)
    wb.save(INPUT)

    cmd = [
        sys.executable, str(RES / "excel_style_cleaner.py"),
        str(INPUT), "-o", str(OUTPUT),
        "--xwriter",
    ]
    if extra_args:
        cmd.extend(extra_args)
    res = subprocess.run(cmd, capture_output=True, text=True)
    pass_ = (res.returncode == 0)
    label = "通过" if expect_pass else "失败"
    print("\n=== " + name + " (期望 " + label + ") ===")
    print("   returncode: " + str(res.returncode))
    last_lines = res.stdout.strip().split("\n")[-5:]
    for ln in last_lines:
        if ln.strip():
            print("   | " + ln)
    if res.returncode != 0 and res.stderr:
        stderr_lines = res.stderr.strip().split("\n")[-3:]
        for ln in stderr_lines:
            if ln.strip():
                print("   ERR | " + ln)
    status = "[OK]" if (pass_ == expect_pass) else "[FAIL]"
    print("   " + status + " " + name)
    return pass_ == expect_pass

results = []

# Case 1: percent
results.append(run_case("percent", [("达成率", 0.92, "0.00%")], True))

# Case 2: 万级
results.append(run_case("wan", [("销售额", 1500000, '0"."0,"万"')], True))

# Case 3: 千分位
results.append(run_case("thousand", [("数量", 12345, '#,##0')], True))

# Case 4: 短数字
results.append(run_case("short_num", [("小计", 100, '0.0')], True))

# Case 5: [Red] 负数红（百分比格式）
results.append(run_case("red", [("环比", -0.05, '0.00%;[Red]-0.00%')], True))

# Case 6: 自定义（日期，不强制检查）
results.append(run_case("custom", [("日期", 45123, 'yyyy-mm-dd')], True))

# Case 7: 跳过自检（传 --no-preserve-input-format）
results.append(run_case(
    "percent_no_check",
    [("达成率", 0.92, "0.00%")],
    True,
    extra_args=["--no-preserve-input-format"],
))

# 清理临时文件
for f in os.listdir(SKILL / "references"):
    if f.startswith("_v317_"):
        try:
            (SKILL / "references" / f).unlink()
        except OSError:
            pass

print("\n\n=== 总览 ===")
print("通过: " + str(sum(results)) + "/" + str(len(results)))
if all(results):
    print("[OK] 全部 7 个测试用例通过")
else:
    print("[FAIL] 部分用例失败")
