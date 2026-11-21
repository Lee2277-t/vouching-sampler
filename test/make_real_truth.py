# -*- coding: utf-8 -*-
"""为两份公开的真实序时账算期望值。

数据来源（公开仓库，非客户数据）：
  aging      https://github.com/Cathywang1988/aging-generator
  valuation  https://github.com/zer0-lyz/valuation-detail-table

这里用 openpyxl 独立扫一遍，和 JS 侧（SheetJS + core.js）比。
序时账的借贷必须相等 —— 这条恒等式是最硬的独立校验，不依赖任何一方的实现。
"""
import json
import os
import shutil
from decimal import Decimal

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(os.path.dirname(HERE), "samples")
FIXTURE = os.path.join(HERE, "fixture")

SPEC = {
    # 文件名: (数据起始行, 借方列, 贷方列, 科目列, 1-based)
    "aging_xushizhang.xlsx": (2, 5, 6, 8),
    "valuation_xushizhang.xlsx": (2, 8, 9, 14),
}


def num(v):
    if v is None or v == "":
        return None
    return Decimal(str(v))


def scan(path, first_row, dcol, ccol, acol):
    ws = load_workbook(path, data_only=True)["Sheet1"]
    txn = 0
    debit = Decimal("0")
    credit = Decimal("0")
    abs_total = Decimal("0")
    skipped = 0
    receivable = 0
    for r in range(first_row, ws.max_row + 1):
        d = num(ws.cell(row=r, column=dcol).value)
        c = num(ws.cell(row=r, column=ccol).value)
        if d is None and c is None:
            skipped += 1
            continue
        txn += 1
        debit += d or 0
        credit += c or 0
        abs_total += abs((d or 0) - (c or 0))
        acct = str(ws.cell(row=r, column=acol).value or "")
        if "应收" in acct:
            receivable += 1
    return {
        "transactions": txn,
        "skipped_no_amount": skipped,
        "debit_cents": int(debit * 100),
        "credit_cents": int(credit * 100),
        "total_abs_cents": int(abs_total * 100),
        "account_filter_receivable": receivable,
    }


out = {}
for name, (first_row, dcol, ccol, acol) in SPEC.items():
    src = os.path.join(SAMPLES, name)
    if not os.path.exists(src):
        raise SystemExit("缺样例文件: " + src)
    out[name] = scan(src, first_row, dcol, ccol, acol)
    # 拷一份进 fixture，测试不依赖 samples/ 是否还在
    shutil.copyfile(src, os.path.join(FIXTURE, name))

with open(os.path.join(FIXTURE, "real_truth.json"), "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print(json.dumps(out, ensure_ascii=False, indent=2))
