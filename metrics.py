#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
metrics.py —— 由原始科目计算净息差、不良率、拨备覆盖率，并与年报披露值对账。

设计原则：
1. 输入是原始科目（利息收入、利息支出、平均生息资产、不良余额、贷款总额、损失准备），
   不是人工算好的指标——这样"分析工具"才名实相符；
2. 披露值只用于对账，不参与计算（计算与核对分离，才能发现算错）；
3. 拨备覆盖率同时输出"含 FVOCI / 仅摊余"两种口径：FVOCI 准备波动大（招行 2022 年
   单年 65.63 亿元，占准备 2.51%），漏掉会让覆盖率假跳 10 个百分点以上。

用法：
    python metrics.py            # 打印对账表
    python metrics.py --check    # 对帐通过退出 0，否则退出 1（可挂 CI）
"""

import argparse
import csv
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
RAW_FILE = BASE_DIR / "data_raw.csv"

TOL_BP = 1.0    # 净息差、不良率的容差（bp）
TOL_PP = 0.1    # 拨备覆盖率的容差（百分点）

# 年报披露值（仅用于对账）。None 表示该年未取到披露值，不做比对。
REPORTED = {
    ("招商银行", 2021): {"nim": 2.48, "npl": 0.91, "coverage": 483.87},
    ("招商银行", 2022): {"nim": 2.40, "npl": 0.96, "coverage": 450.79},
    ("招商银行", 2023): {"nim": 2.15, "npl": 0.95, "coverage": 437.70},
    ("招商银行", 2024): {"nim": 1.98, "npl": 0.95, "coverage": 411.98},
    ("工商银行", 2021): {"nim": 2.11, "npl": 1.42, "coverage": None},
    ("工商银行", 2022): {"nim": 1.92, "npl": 1.38, "coverage": None},
    ("工商银行", 2023): {"nim": 1.61, "npl": 1.36, "coverage": None},
    ("工商银行", 2024): {"nim": 1.42, "npl": 1.34, "coverage": 214.91},
}

FLOAT_FIELDS = [
    "利息收入", "利息支出", "平均生息资产", "不良贷款余额",
    "贷款和垫款总额", "贷款损失准备", "贷款损失准备_FVOCI",
]


def load_raw(path=RAW_FILE):
    """读取原始科目表，返回 [{'银行','年份', 数值字段...}]。"""
    path = Path(path)
    if not path.exists():
        raise SystemExit(f"[错误] 未找到原始科目表：{path}")
    with open(path, encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    out = []
    for row in rows:
        item = {"银行": row["银行"].strip(), "年份": int(row["年份"])}
        for field in FLOAT_FIELDS:
            item[field] = float(row[field])
        out.append(item)
    return out


def nim(row):
    """净息差（%）= 净利息收入 / 平均生息资产。"""
    return (row["利息收入"] - row["利息支出"]) / row["平均生息资产"] * 100


def npl(row):
    """不良率（%）= 不良贷款余额 / 贷款和垫款总额。"""
    return row["不良贷款余额"] / row["贷款和垫款总额"] * 100


def coverage(row, include_fvoci=True):
    """拨备覆盖率（%）= 贷款损失准备 / 不良贷款余额；可选择是否含 FVOCI 准备。"""
    provision = row["贷款损失准备"]
    if not include_fvoci:
        provision -= row["贷款损失准备_FVOCI"]
    return provision / row["不良贷款余额"] * 100


def reconcile(row):
    """比对一行，返回各项偏差（bp / pp）与是否通过。"""
    ref = REPORTED.get((row["银行"], row["年份"]))
    if ref is None:
        return None
    out = {
        "nim": (nim(row) - ref["nim"]) * 100,
        "npl": (npl(row) - ref["npl"]) * 100,
        "coverage": None,
    }
    if ref["coverage"] is not None:
        out["coverage"] = coverage(row) - ref["coverage"]
    out["ok"] = (
        abs(out["nim"]) <= TOL_BP
        and abs(out["npl"]) <= TOL_BP
        and (out["coverage"] is None or abs(out["coverage"]) <= TOL_PP)
    )
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(description="原始科目 → 指标计算与年报对账")
    parser.add_argument("--data", type=Path, default=RAW_FILE, help="原始科目 CSV")
    parser.add_argument("--check", action="store_true", help="仅校验，全部通过退出 0")
    args = parser.parse_args(argv)

    rows = load_raw(args.data)
    failed = 0

    print(f"[数据] {args.data.name}：{len(rows)} 行，"
          f"{len({r['银行'] for r in rows})} 家银行")
    header = (f"{'银行':<6}{'年份':<6}{'NIM':>8}{'偏差bp':>8}"
              f"{'不良率':>8}{'偏差bp':>8}{'覆盖率':>10}{'偏差pp':>8}"
              f"{'仅摊余':>9}{'漏FVOCI':>9}")
    print(header)
    print("-" * len(header.encode("gbk", "ignore")))

    for row in rows:
        ref = REPORTED.get((row["银行"], row["年份"]))
        res = reconcile(row)
        dev_nim = f"{res['nim']:+.2f}" if res else "—"
        dev_npl = f"{res['npl']:+.2f}" if res else "—"
        dev_cov = f"{res['coverage']:+.2f}" if res and res["coverage"] is not None else "—"
        if res and res["coverage"] is not None:
            gap = coverage(row, include_fvoci=False) - coverage(row)
        else:
            gap = coverage(row, include_fvoci=False) - coverage(row)
        print(f"{row['银行']:<6}{row['年份']:<6}{nim(row):>8.2f}{dev_nim:>8}"
              f"{npl(row):>8.2f}{dev_npl:>8}{coverage(row):>10.2f}{dev_cov:>8}"
              f"{coverage(row, include_fvoci=False):>9.2f}{gap:>+9.2f}")
        if res and not res["ok"]:
            failed += 1

    checked = sum(1 for r in rows if (r["银行"], r["年份"]) in REPORTED)
    print("-" * len(header.encode("gbk", "ignore")))
    print(f"[对账] 可核对 {checked} 行，超差 {failed} 行"
          f"（容差：NIM/NPL ≤ {TOL_BP} bp，覆盖率 ≤ {TOL_PP} pp）")
    print(f"[口径] 拨备覆盖率含 FVOCI；'仅摊余'列展示漏计 FVOCI 时的偏差（pp）")

    if args.check and failed:
        print("[结果] 未通过", file=sys.stderr)
        return 1
    print("[结果] 通过" if failed == 0 else "[结果] 存在超差行")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
