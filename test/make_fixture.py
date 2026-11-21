# -*- coding: utf-8 -*-
"""生成合成明细账夹具 + ground truth。

要点：
  * 全部为合成数据，不含任何真实客户、企业名、金额。
  * 故意做脏：标题行、空行、合计行、文本金额、千分位、括号负数、
    三位小数、文本日期、日期缺失、期间外、零金额、大额（进全查层）。
  * ground truth 用 Decimal 独立算一遍写到 ground_truth.json，
    与 JS 侧解析结果比对——两边是两条独立路径，不是自己验自己。
"""
import json
import os
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "fixture")
os.makedirs(OUT_DIR, exist_ok=True)

CLIENT = "合成测试商贸有限公司"

rows = []
base = date(2026, 1, 10)
for i in range(40):
    d = base + timedelta(days=i * 8)
    rows.append((d, "记-%03d" % (i + 1), "销售货款%02d" % (i + 1),
                 "1122 应收账款", Decimal("137.25") * (i + 1), None))

EDGE = [
    (None, "记-901", "日期缺失测试", "1122 应收账款", Decimal("1000.00"), None),
    (date(2027, 2, 1), "记-902", "期后发生", "1122 应收账款", Decimal("2000.00"), None),
    (date(2026, 6, 15), "记-903", "红冲至零", "1122 应收账款", Decimal("0"), None),
    (date(2026, 6, 16), "记-904", "文本金额", "1122 应收账款", "1,234.56", None),
    (date(2026, 6, 17), "记-905", "三位小数", "1122 应收账款", "1,000.005", None),
    (date(2026, 6, 18), "记-906", "大额销售", "1122 应收账款", Decimal("500000.00"), None),
    ("2026/7/5", "记-907", "收款", "1122 应收账款", None, Decimal("3000.00")),
]
rows.extend(EDGE)


def q(v):
    s = str(v).replace(",", "").replace("元", "").strip()
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    if s.endswith("-"):
        s = "-" + s[:-1]
    return Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def to_cents(v):
    return int(q(v) * 100)


included = []
excluded = {"date_missing": 0, "out_of_period": 0, "zero": 0}
for d, v, s, a, dr, cr in rows:
    dr_c = to_cents(dr) if dr is not None else None
    cr_c = to_cents(cr) if cr is not None else None
    signed = (dr_c or 0) - (cr_c or 0)
    if d is None:
        excluded["date_missing"] += 1
        continue
    if isinstance(d, date) and d.year != 2026:
        excluded["out_of_period"] += 1
        continue
    if signed == 0:
        excluded["zero"] += 1
        continue
    included.append(signed)

THRESH = 10000000  # 100,000.00 元 → 分
full = [x for x in included if abs(x) >= THRESH]
rest = [x for x in included if abs(x) < THRESH]

gt = {
    "client": CLIENT,
    "row_count_total": len(rows),
    "included": len(included),
    "excluded": excluded,
    "total_abs_cents": sum(abs(x) for x in included),
    "net_cents": sum(included),
    "debit_cents": sum(x for x in included if x > 0),
    "credit_cents": sum(-x for x in included if x < 0),
    "threshold_cents": THRESH,
    "full_count": len(full),
    "rest_count": len(rest),
    "full_abs_cents": sum(abs(x) for x in full),
    "rest_abs_cents": sum(abs(x) for x in rest),
    "specific": {
        "text_amount_1234_56": to_cents("1,234.56"),
        "three_decimal": to_cents("1,000.005"),
        "paren_negative": to_cents("(800.00)"),
        "trailing_minus": to_cents("500.00-"),
        "yuan_suffix": to_cents("1,200元"),
        "large": to_cents("500000.00"),
    },
}

# ---------- 主表：借贷双列，脏格式；表头固定在第 3 行 ----------
wb = Workbook()
ws = wb.active
ws.title = "应收账款明细账"
ws["A1"] = "%s  2026年度  应收账款  明细账" % CLIENT
ws["A2"] = "科目编码：1122    单位：人民币元"

header = ["记账日期", "凭证号", "摘要", "科目名称", "借方金额", "贷方金额", "余额", "制单人"]
for i, h in enumerate(header, start=1):
    c = ws.cell(row=3, column=i, value=h)
    c.font = Font(bold=True)

r = 4
for item in rows:
    d, v, s, a, dr, cr = item
    ws.cell(row=r, column=1, value=d)
    ws.cell(row=r, column=2, value=v)
    ws.cell(row=r, column=3, value=s)
    ws.cell(row=r, column=4, value=a)
    if dr is not None:
        ws.cell(row=r, column=5, value=dr)
    if cr is not None:
        ws.cell(row=r, column=6, value=cr)
    ws.cell(row=r, column=8, value="张三")
    r += 1
    if v == "记-010":
        r += 1  # 空行

# 小计 / 合计行（非明细，应被跳过并报告）
for label in ("小计", "合计"):
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=3, value="本期发生额")
    ws.cell(row=r, column=5, value="1,000.00")
    r += 1

# ---------- 单列金额 + 方向 ----------
ws2 = wb.create_sheet("单列金额")
h2 = ["业务日期", "单号", "摘要", "金额", "方向"]
for i, h in enumerate(h2, start=1):
    ws2.cell(row=1, column=i, value=h)
data2 = [
    (date(2026, 3, 1), "P-001", "收款A", Decimal("800.00"), "贷"),
    (date(2026, 3, 2), "P-002", "收款B", "(1,200.00)", None),
    (date(2026, 3, 3), "P-003", "付款C", Decimal("450.50"), "借"),
    (date(2026, 3, 4), "P-004", "收款D", "2,000.00", "贷"),
]
for i, (d, v, s, amt, dirn) in enumerate(data2, start=2):
    ws2.cell(row=i, column=1, value=d)
    ws2.cell(row=i, column=2, value=v)
    ws2.cell(row=i, column=3, value=s)
    ws2.cell(row=i, column=4, value=amt)
    if dirn:
        ws2.cell(row=i, column=5, value=dirn)

ws3 = wb.create_sheet("使用说明")
ws3["A1"] = "本表由测试夹具生成，仅供程序验证使用。"

path = os.path.join(OUT_DIR, "sample_ledger.xlsx")
wb.save(path)

# ---------- 真实世界格式（结构照抄，数据全合成） ----------
# 四张表各代表一类真实导出的毛病：标题行占两行、凭证号拆列、透视表扁平导出、
# 科目余额表混在同一个工作簿里。全部为合成内容。
ws4 = wb.create_sheet("序时账")
ws4["A1"] = "记账凭证"
ws4["A2"] = "2026年满足条件查询的全部凭证"
HDR4 = ["日期", "字", "号", "摘要", "科目代码", "科目名称", "借方金额", "贷方金额",
        "客户档案代码", "客户档案名称", "现金流量项目代码", "现金流量项目名称"]
for i, h in enumerate(HDR4, start=1):
    ws4.cell(row=3, column=i, value=h)

r = 4
xs_total = Decimal("0")
for n in range(1, 7):
    amt = Decimal("1000.00") + n * 100
    xs_total += amt * 2
    dr_line = [date(2026, n, 10), "记", n, "合成业务%02d" % n, "112201", "应收账款-合成客户", amt, None]
    cr_line = [date(2026, n, 10), "记", n, "合成业务%02d" % n, "600101", "主营业务收入-合成", None, amt]
    for line in (dr_line, cr_line):
        for i, v in enumerate(line, start=1):
            if v is not None:
                ws4.cell(row=r, column=i, value=v)
        ws4.cell(row=r, column=9, value="KH001")
        ws4.cell(row=r, column=10, value="合成客户")
        r += 1

# SAP 式扁平导出：单行表头，科目/借方/贷方散在很后面
ws5 = wb.create_sheet("明细表")
HDR5 = ["公司代码", "公司代码描述", "凭证编号", "会计年度", "期间", "货币", "公司货币",
        "归档凭证编号", "凭证日期", "过账日期", "分配", "凭证行号",
        "科目", "原币金额", "借方", "贷方"]
for i, h in enumerate(HDR5, start=1):
    ws5.cell(row=1, column=i, value=h)

sap_total = Decimal("0")
for i in range(5):
    amt = Decimal("500.00")
    sap_total += amt
    ws5.cell(row=i + 2, column=1, value="T01")
    ws5.cell(row=i + 2, column=2, value="合成测试公司")
    ws5.cell(row=i + 2, column=3, value="20000000%02d" % (i + 1))
    ws5.cell(row=i + 2, column=9, value="2025-1%d-02" % (i + 1))
    ws5.cell(row=i + 2, column=13, value="100201 银行存款")
    ws5.cell(row=i + 2, column=14, value=amt)
    if i % 2 == 0:
        ws5.cell(row=i + 2, column=15, value=amt)
    else:
        ws5.cell(row=i + 2, column=16, value=amt)

# 制单凭证：日期列叫「制单日期」——真实数据里就是它把工具打回了
ws6 = wb.create_sheet("填制凭证")
HDR6 = ["制单日期", "凭证字号", "凭证本币总金额", "审核人", "制单人", "摘要",
        "科目编码", "科目名称", "辅助项", "借方(本币)", "贷方(本币)"]
for i, h in enumerate(HDR6, start=1):
    ws6.cell(row=1, column=i, value=h)

tz_total = Decimal("0")
r = 2
for n in (1, 2):
    amt = Decimal("300.00") * n
    tz_total += amt * 2
    ws6.cell(row=r, column=1, value="2025-02-1%d" % n)
    ws6.cell(row=r, column=2, value="记-%04d" % n)
    ws6.cell(row=r, column=6, value="合成费用")
    ws6.cell(row=r, column=7, value="66020313")
    ws6.cell(row=r, column=8, value="管理费用-合成")
    ws6.cell(row=r, column=10, value=amt)
    r += 1
    ws6.cell(row=r, column=1, value="2025-02-1%d" % n)
    ws6.cell(row=r, column=2, value="记-%04d" % n)
    ws6.cell(row=r, column=6, value="合成费用")
    ws6.cell(row=r, column=7, value="10026101")
    ws6.cell(row=r, column=8, value="银行存款-合成")
    ws6.cell(row=r, column=11, value=amt)
    r += 1

# 科目余额表：两行标题 + 两级表头，没有日期。混在同一个工作簿里，绝不能被选中
ws7 = wb.create_sheet("科目余额表")
ws7["A1"] = "账户汇总表"
ws7["A2"] = "2026年1月至6月"
ws7["A3"] = "账户代码1"
ws7["C3"] = "账户名称1"
ws7["E3"] = "期初余额"
ws7["I3"] = "期间发生额"
for i, h in enumerate(["账户代码1", "账户代码2", "账户名称1", "账户名称2", "方向", "金额",
                       "借方金额", "贷方金额", "借方金额", "贷方金额"], start=1):
    ws7.cell(row=4, column=i, value=h)
for i in range(3):
    ws7.cell(row=5 + i, column=1, value="1002%02d" % (i + 1))
    ws7.cell(row=5 + i, column=3, value="合成科目%02d" % (i + 1))
    ws7.cell(row=5 + i, column=5, value="借")
    ws7.cell(row=5 + i, column=6, value=Decimal("1000.00"))

# 两级表头：上一行是「本期发生额」这种分类（合并两列），下一行才是「借方金额」
ws8 = wb.create_sheet("两级表头")
for i, h in enumerate(["记账日期", "凭证字号", "摘要", "科目名称", "本期发生额", None, "余额"], start=1):
    if h:
        ws8.cell(row=1, column=i, value=h)
ws8.merge_cells("E1:F1")
for i, h in enumerate([None, None, None, None, "借方金额", "贷方金额", "金额"], start=1):
    if h:
        ws8.cell(row=2, column=i, value=h)

lvl_total = Decimal("0")
for n in range(1, 6):
    amt = Decimal("200.00") * n
    lvl_total += amt
    ws8.cell(row=n + 2, column=1, value=date(2026, 3, n))
    ws8.cell(row=n + 2, column=2, value="记-%02d" % n)
    ws8.cell(row=n + 2, column=3, value="合成两级业务%02d" % n)
    ws8.cell(row=n + 2, column=4, value="应收账款-两级")
    ws8.cell(row=n + 2, column=5, value=amt)
    ws8.cell(row=n + 2, column=7, value=Decimal("0"))

gt["twolevel"] = {"records": 5, "total_abs_cents": int(lvl_total * 100)}

# 竖排：一笔业务占一列（字段名竖着排），转置后才是「一行一条明细」
ws9 = wb.create_sheet("竖排凭证")
rows9 = [
    ["日期", date(2026, 1, 5), date(2026, 1, 6), date(2026, 1, 7)],
    ["凭证号", "记-01", "记-02", "记-03"],
    ["摘要", "合成竖排A", "合成竖排B", "合成竖排C"],
    ["科目", "1122 应收账款", "1122 应收账款", "1122 应收账款"],
    ["借方", Decimal("100.00"), Decimal("200.00"), Decimal("300.00")],
]
for r_i, row in enumerate(rows9, start=1):
    for c_i, v in enumerate(row, start=1):
        ws9.cell(row=r_i, column=c_i, value=v)

gt["vertical"] = {"records": 3, "total_abs_cents": 60000}

gt["realworld"] = {
    "xushizhang": {"records": 12, "total_abs_cents": int(xs_total * 100),
                   "first_voucher": "记-1", "account": "应收账款-合成客户", "account_code": "112201"},
    "sap": {"records": 5, "total_abs_cents": int(sap_total * 100)},
    "tianzhi": {"records": 4, "total_abs_cents": int(tz_total * 100), "first_voucher": "记-0001"},
    "balance_sheet_usable": False,
}

# 上面几张表是在第一次 save 之后才建的，必须重新保存一次
wb.save(path)

gt["amount_mode"] = {
    "count": len(data2),
    "total_abs_cents": sum(abs(to_cents(x[3])) for x in data2),
    "net_cents": sum(to_cents(x[3]) * (-1 if x[4] == "贷" else 1) for x in data2),
}

with open(os.path.join(OUT_DIR, "ground_truth.json"), "w", encoding="utf-8") as f:
    json.dump(gt, f, ensure_ascii=False, indent=2)

print("fixture:", path)
print(json.dumps(gt, ensure_ascii=False, indent=2))
