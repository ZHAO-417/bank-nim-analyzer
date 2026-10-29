#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
main.py —— 商业银行净息差（NIM）与不良贷款率（NPL）可视化工具

版本：v0.3.0

功能
----
1. 读取 data.csv（标准 UTF-8-BOM CSV；同时兼容旧版“改名为 .csv 的 Excel 工作簿”，
   后者经 BytesIO 读取，可规避 openpyxl 3.1 对扩展名的校验）；
2. 用 pandas 完成字段清洗、透视与汇总统计；
3. 用 matplotlib 绘制两张图并保存为 PNG：
      chart_nim.png  折线图：各银行净息差走势 + 行业均值 + 息差溢价标注
      chart_npl.png  柱状图：各银行不良率对比
4. 自动选择系统中文字体（Microsoft YaHei / SimHei），解决中文乱码与负号乱码。

用法
----
    python main.py
    python main.py --data data.csv --outdir outputs

依赖
----
    pandas, matplotlib, openpyxl（仅当数据文件实为 .xlsx / 旧版 xlsx-csv 时需要）
"""

import argparse
import sys
from io import BytesIO
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import font_manager

import metrics as mx

__version__ = "0.3.0"

# --------------------------------------------------------------------------- #
# 0. 全局配置
# --------------------------------------------------------------------------- #
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DATA = BASE_DIR / "data.csv"
DEFAULT_RAW = BASE_DIR / "data_raw.csv"

CHART_NIM_NAME = "chart_nim.png"
CHART_NPL_NAME = "chart_npl.png"

# 数据表的标准列名，顺序即输出顺序
REQUIRED_COLS = ["银行", "年份", "净息差%", "不良率%"]
NUM_COLS = ["净息差%", "不良率%"]

# 兼容不同的表头写法（含空格、全角括号、英文缩写）
COLUMN_ALIASES = {
    "银行": ["银行", "银行名称", "机构", "机构名称", "bank"],
    "年份": ["年份", "年度", "year"],
    "净息差%": ["净息差%", "净息差(%)", "净息差", "nim", "净利息收益率"],
    "不良率%": ["不良率%", "不良率(%)", "不良率", "npl", "不良贷款率"],
}

# 行业基准：商业银行整体净息差（%）。
# 来源：2020 年为国家金融监督管理总局原表；2021—2024 年为银保监会/金监总局季度监管指标
# （2021 年为三源交叉整理值）。2025Q4 为 1.42%，不在样本期内故不绘入。
INDUSTRY_NIM = {2020: 2.10, 2021: 2.08, 2022: 1.91, 2023: 1.69, 2024: 1.52}
INDUSTRY_NAME = "商业银行行业均值"
INDUSTRY_SOURCE = "银保监会/国家金融监督管理总局 2020—2024 年商业银行净息差"

# 品牌感配色；未收录的银行按顺序取 FALLBACK_COLORS
BANK_COLORS = {
    "招商银行": "#C8102E",   # 招行红
    "工商银行": "#1F4E79",   # 工行深蓝
    "建设银行": "#2E86C1",
    "农业银行": "#1E8449",
    "中国银行": "#B7950B",
    "交通银行": "#6C3483",
    "邮储银行": "#CA6F1E",
}
FALLBACK_COLORS = ["#C8102E", "#1F4E79", "#2E86C1", "#1E8449", "#6C3483", "#CA6F1E"]

# 中文字体候选：Windows 优先微软雅黑，其次黑体
CJK_FONTS = ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC",
             "Source Han Sans SC", "PingFang SC", "WenQuanYi Micro Hei"]


# --------------------------------------------------------------------------- #
# 1. 中文字体
# --------------------------------------------------------------------------- #
def setup_chinese_font():
    """把 matplotlib 的默认字体切换为系统已安装的中文字体。

    返回实际选用的字体名；若一个都没找到则打印告警（此时中文会显示为方框）。
    """
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in CJK_FONTS:
        if name in installed:
            plt.rcParams["font.sans-serif"] = [name] + [
                f for f in plt.rcParams["font.sans-serif"] if f != name
            ]
            plt.rcParams["font.family"] = "sans-serif"
            plt.rcParams["axes.unicode_minus"] = False  # 负号正常显示
            return name

    print("[警告] 未找到中文字体，图表中的汉字可能显示为方框。", file=sys.stderr)
    print(f"      已尝试：{', '.join(CJK_FONTS)}", file=sys.stderr)
    return None


# --------------------------------------------------------------------------- #
# 2. 读取与清洗
# --------------------------------------------------------------------------- #
def _normalize_header(name):
    """统一表头写法：去空格、全角括号转半角、转小写。"""
    text = str(name).replace("\ufeff", "").strip()
    text = text.replace("（", "(").replace("）", ")").replace(" ", "")
    return text.lower()


_LOOKUP = {
    _normalize_header(alias): canon
    for canon, aliases in COLUMN_ALIASES.items()
    for alias in aliases
}


def load_data(path):
    """读取数据文件，自动识别真实格式。

    v0.1 的 data.csv 实际是 Excel 工作簿（文件头 PK\\x03\\x04），仅因 openpyxl 3.0
    对扩展名宽容才能读取；openpyxl 3.1 起会对 .csv 后缀直接抛 InvalidFileException。
    v0.2 起改为读取字节流、经 BytesIO 交给 read_excel，从而绕开扩展名校验，
    旧文件仍可打开；仓库内的 data.csv 已转为标准 UTF-8-BOM CSV。
    """
    path = Path(path)
    if not path.exists():
        raise SystemExit(f"[错误] 未找到数据文件：{path}")

    raw = path.read_bytes()
    if raw[:4] == b"PK\x03\x04":                 # xlsx / xlsm（zip 容器）
        df = pd.read_excel(BytesIO(raw), sheet_name=0)
        print(f"[读取] {path.name} 是 Excel 工作簿，已用 read_excel 读取字节流载入。")
        return df

    last_error = None
    for encoding in ("utf-8-sig", "gbk"):
        try:
            df = pd.read_csv(BytesIO(raw), encoding=encoding)
            print(f"[读取] {path.name} 以 {encoding} 编码载入。")
            return df
        except UnicodeDecodeError as exc:
            last_error = exc

    raise SystemExit(f"[错误] 无法解码数据文件 {path}：{last_error}")


def clean_data(df):
    """规范列名、剔除空行、统一数值类型，返回干净的长表。"""
    rename = {col: _LOOKUP[_normalize_header(col)]
              for col in df.columns if _normalize_header(col) in _LOOKUP}
    df = df.rename(columns=rename)

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"[错误] 缺少必需列：{missing}\n"
            f"       文件实际列名：{list(df.columns)}\n"
            f"       支持的表头写法见 COLUMN_ALIASES。"
        )

    df = df[REQUIRED_COLS].copy()
    df = df.dropna(how="all")

    df["银行"] = df["银行"].astype(str).str.strip()
    df["年份"] = pd.to_numeric(df["年份"], errors="coerce")
    for col in NUM_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.replace({"银行": {"": np.nan, "nan": np.nan, "None": np.nan}})
    df = df.dropna(subset=["银行", "年份"] + NUM_COLS)
    df["年份"] = df["年份"].astype(int)
    df = df.drop_duplicates(subset=["银行", "年份"], keep="last")
    df = df.sort_values(["年份"], kind="stable").reset_index(drop=True)

    if df.empty:
        raise SystemExit("[错误] 清洗后没有任何有效数据，请检查 data.csv 的内容。")

    return df


def to_wide(df, value_col):
    """长表 -> 宽表：行为年份、列为银行（保持数据中银行首次出现的顺序）。"""
    wide = df.pivot_table(index="年份", columns="银行",
                          values=value_col, aggfunc="last")
    order = [bank for bank in df["银行"].unique() if bank in wide.columns]
    return wide.reindex(columns=order).sort_index()


# --------------------------------------------------------------------------- #
# 3. 绘图
# --------------------------------------------------------------------------- #
def _color_of(bank, index):
    return BANK_COLORS.get(bank, FALLBACK_COLORS[index % len(FALLBACK_COLORS)])


def _style_axes(ax):
    """去掉上/右边框、统一网格风格。"""
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color("#B0B7C3")
    ax.spines["bottom"].set_color("#B0B7C3")
    ax.grid(axis="y", linestyle="--", linewidth=0.8, color="#D5D8DC", alpha=0.9)
    ax.set_axisbelow(True)
    ax.tick_params(colors="#34495E", labelsize=10.5)


def plot_nim(nim_wide, out_path):
    """折线图：净息差（NIM）走势，附行业均值与息差溢价标注。"""
    years = [int(y) for y in nim_wide.index]
    banks = list(nim_wide.columns)

    fig, ax = plt.subplots(figsize=(9.4, 5.4), dpi=110)

    for i, bank in enumerate(banks):
        series = nim_wide[bank].astype(float)
        color = _color_of(bank, i)
        ax.plot(years, series.values, color=color, linewidth=2.4,
                marker="o", markersize=7.5, markerfacecolor="white",
                markeredgewidth=2.2, markeredgecolor=color,
                label=bank, zorder=3)

        # 数值标签统一置于数据点上方 7 pt，并加白色描边，避免与行业虚线交叠
        for x, y in zip(years, series.values):
            ax.annotate(f"{y:.2f}", xy=(x, y), xytext=(0, 7),
                        textcoords="offset points", ha="center", va="bottom",
                        fontsize=9, fontweight="bold", color=color, zorder=6,
                        bbox=dict(boxstyle="round,pad=0.12", fc="white",
                                  ec="none", alpha=0.75))

    # 行业均值：只画样本年份内存在的 2022—2024 年，避免用一条直线冒充全期基准
    ind_years = sorted(y for y in INDUSTRY_NIM if years[0] <= y <= years[-1])
    if ind_years:
        ax.plot(ind_years, [INDUSTRY_NIM[y] for y in ind_years],
                linestyle="--", linewidth=1.6, color="#95A5A6",
                marker="s", markersize=5, label=f"{INDUSTRY_NAME}（NIM）",
                zorder=2)

    # 息差溢价：在空白区给出说明框，避免箭头与数据标签（尤其 1.42）相互遮挡
    if len(banks) == 2:
        top, bottom = banks[0], banks[1]
        if nim_wide[top].mean() < nim_wide[bottom].mean():
            top, bottom = bottom, top
        first, last = years[0], years[-1]
        gap_first = (float(nim_wide[top].loc[first]) - float(nim_wide[bottom].loc[first])) * 100
        gap_last = (float(nim_wide[top].loc[last]) - float(nim_wide[bottom].loc[last])) * 100
        if np.isfinite(gap_first) and np.isfinite(gap_last):
            ax.text(0.02, 0.05,
                    f"两线间距＝息差溢价：{gap_first:.0f} bp（{first}）"
                    f"→ {gap_last:.0f} bp（{last}）",
                    transform=ax.transAxes, fontsize=9, color="#5D6D7E",
                    va="bottom", ha="left",
                    bbox=dict(boxstyle="round,pad=0.38", fc="#F8F9FA",
                              ec="#D5D8DC", lw=0.8))

    y_min, y_max = float(nim_wide.min().min()), float(nim_wide.max().max())
    pad = max((y_max - y_min) * 0.35, 0.15)
    floor = max(0.0, round(y_min - pad, 1))
    ax.set_ylim(floor, round(y_max + pad * 0.8, 1))

    ax.set_xticks(years)
    ax.set_xlim(years[0] - 0.35, years[-1] + 0.35)
    ax.set_xlabel("年份", fontsize=11.5, color="#2C3E50", labelpad=8)
    ax.set_ylabel("净息差（%）", fontsize=11.5, color="#2C3E50", labelpad=8)
    ax.set_title(f"{' vs '.join(banks)}：净息差（NIM）走势（{years[0]}—{years[-1]}）",
                 fontsize=15, fontweight="bold", color="#1B2631", pad=16)

    _style_axes(ax)
    ax.legend(loc="upper right", frameon=False, fontsize=10,
              handlelength=2.2, borderaxespad=0.6)

    note = (f"数据来源：data.csv（各银行年度报告）｜行业均值为{INDUSTRY_SOURCE}｜"
            f"纵轴自 {floor:.1f}% 起，以突出差异")
    fig.text(0.005, -0.045, note, fontsize=8.2, color="#7F8C8D")

    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"[输出] {out_path}")


def plot_npl(npl_wide, out_path):
    """分组柱状图：不良贷款率（NPL）对比。"""
    years = [int(y) for y in npl_wide.index]
    banks = list(npl_wide.columns)
    n = len(banks)

    x = np.arange(len(years))
    group_width = 0.76
    bar_width = group_width / n

    fig, ax = plt.subplots(figsize=(9.4, 5.4), dpi=110)

    for i, bank in enumerate(banks):
        values = npl_wide[bank].astype(float).values
        offset = (i - (n - 1) / 2) * bar_width
        bars = ax.bar(x + offset, values, width=bar_width * 0.9,
                      color=_color_of(bank, i), label=bank,
                      edgecolor="white", linewidth=1.0, zorder=3)
        for bar in bars:
            height = bar.get_height()
            ax.annotate(f"{height:.2f}",
                        xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3), textcoords="offset points",
                        ha="center", va="bottom", fontsize=9,
                        color="#34495E", fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels([str(y) for y in years])
    ax.set_xlim(-0.55, len(years) - 0.45)
    # 顶部留白并取 0.1 的整数倍，使刻度落成整齐数值（如 2.00）
    top = float(npl_wide.max().max())
    ax.set_ylim(0, np.ceil(top * 1.25 * 10) / 10)

    ax.set_xlabel("年份", fontsize=11.5, color="#2C3E50", labelpad=8)
    ax.set_ylabel("不良贷款率（%）", fontsize=11.5, color="#2C3E50", labelpad=8)
    ax.set_title(f"{' vs '.join(banks)}：不良贷款率（NPL）对比（{years[0]}—{years[-1]}）",
                 fontsize=15, fontweight="bold", color="#1B2631", pad=16)

    _style_axes(ax)
    ax.legend(loc="upper right", frameon=False, fontsize=10.5,
              handlelength=1.6, borderaxespad=0.6)

    fig.text(0.005, -0.045,
             "数据来源：data.csv（各银行年度报告）｜柱顶数值为不良贷款率（%）；"
             "行业不良率 1.50% 与拨备覆盖率 205.21% 为 2025Q4 数据，时点不同，未绘入图内",
             fontsize=8.2, color="#7F8C8D")

    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"[输出] {out_path}")


# --------------------------------------------------------------------------- #
# 4. 控制台汇总
# --------------------------------------------------------------------------- #
def summarize(df):
    """打印数据透视表与关键变动指标（单位：bp = 0.01 个百分点）。"""
    print("\n" + "=" * 64)
    print("一、净息差（%）年度数据")
    print("=" * 64)
    nim_wide = to_wide(df, "净息差%")
    print(nim_wide.round(2).to_string())

    print("\n" + "=" * 64)
    print("二、不良率（%）年度数据")
    print("=" * 64)
    npl_wide = to_wide(df, "不良率%")
    print(npl_wide.round(2).to_string())

    rows = []
    for bank, group in df.groupby("银行", sort=False):
        group = group.sort_values("年份")
        nim = group["净息差%"].astype(float)
        npl = group["不良率%"].astype(float)
        rows.append({
            "银行": bank,
            "区间": f"{int(group['年份'].iloc[0])}-{int(group['年份'].iloc[-1])}",
            "净息差均值%": round(nim.mean(), 2),
            "净息差变动(bp)": int(round((nim.iloc[-1] - nim.iloc[0]) * 100)),
            "不良率均值%": round(npl.mean(), 2),
            "不良率变动(bp)": int(round((npl.iloc[-1] - npl.iloc[0]) * 100)),
        })

    stats = pd.DataFrame(rows)
    print("\n" + "=" * 64)
    print("三、关键变动汇总（bp = 0.01 个百分点）")
    print("=" * 64)
    print(stats.to_string(index=False))

    print("\n" + "=" * 64)
    print("四、快速观察")
    print("=" * 64)
    for row in rows:
        nim_dir = "收窄" if row["净息差变动(bp)"] < 0 else "走阔"
        npl_dir = "下降" if row["不良率变动(bp)"] < 0 else "上升"
        print(f"· {row['银行']}：净息差由 {nim_wide[row['银行']].iloc[0]:.2f}% "
              f"{nim_dir}至 {nim_wide[row['银行']].iloc[-1]:.2f}%"
              f"（{row['净息差变动(bp)']:+d} bp）；不良率{npl_dir} "
              f"{abs(row['不良率变动(bp)'])} bp。")

    if len(nim_wide.columns) == 2:
        hi, lo = nim_wide.columns[0], nim_wide.columns[1]
        if nim_wide[hi].mean() < nim_wide[lo].mean():
            hi, lo = lo, hi
        first, last = int(nim_wide.index[0]), int(nim_wide.index[-1])
        gap_first = (nim_wide[hi].loc[first] - nim_wide[lo].loc[first]) * 100
        gap_last = (nim_wide[hi].loc[last] - nim_wide[lo].loc[last]) * 100
        print(f"· 息差溢价：{hi} 相对 {lo} 由 {gap_first:.0f} bp 变化至 {gap_last:.0f} bp"
              f"（{gap_last - gap_first:+.0f} bp）。")
    print()


# --------------------------------------------------------------------------- #
# 5. 入口
# --------------------------------------------------------------------------- #
def build_from_raw(path):
    """由原始科目计算 NIM 与 NPL，并与年报披露值对账（v0.3 指标下沉）。"""
    rows = mx.load_raw(path)
    records, checked = [], 0
    max_nim = max_npl = max_cov = 0.0
    for row in rows:
        records.append({
            "银行": row["银行"],
            "年份": row["年份"],
            "净息差%": mx.nim(row),
            "不良率%": mx.npl(row),
        })
        res = mx.reconcile(row)
        if res is None:
            continue
        checked += 1
        max_nim = max(max_nim, abs(res["nim"]))
        max_npl = max(max_npl, abs(res["npl"]))
        if res["coverage"] is not None:
            max_cov = max(max_cov, abs(res["coverage"]))

    df = pd.DataFrame(records)
    print(f"[指标] 由 {Path(path).name} 的原始科目计算：{len(rows)} 行（拨备覆盖率含 FVOCI）")
    if checked:
        print(f"[对账] 可核对 {checked} 行；最大偏差 NIM {max_nim:.2f} bp、"
              f"NPL {max_npl:.2f} bp、覆盖率 {max_cov:.2f} pp")
    return df


def parse_args():
    parser = argparse.ArgumentParser(
        description="由原始科目计算净息差/不良率并绘图，或直接读披露值 data.csv"
    )
    parser.add_argument("--source", choices=("disclosed", "raw"), default="disclosed",
                        help="disclosed=读 data.csv 披露值（默认，覆盖 2020—2024）；"
                             "raw=由原始科目计算（覆盖 2021—2024）")
    parser.add_argument("--raw-data", type=Path, default=DEFAULT_RAW,
                        help=f"原始科目表路径（默认：{DEFAULT_RAW.name}）")
    parser.add_argument("--data", type=Path, default=DEFAULT_DATA,
                        help=f"披露值表路径（--source disclosed 时使用，默认：{DEFAULT_DATA.name}）")
    parser.add_argument("--outdir", type=Path, default=BASE_DIR,
                        help="图片输出目录（默认：脚本所在目录）")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser.parse_args()


def main():
    args = parse_args()

    print(f"[版本] v{__version__}")
    font = setup_chinese_font()
    print(f"[字体] 使用中文字体：{font}" if font else "[字体] 未找到中文字体")

    if args.source == "raw":
        df = build_from_raw(args.raw_data)
    else:
        df = clean_data(load_data(args.data))

    print(f"[数据] 共 {len(df)} 行，{df['银行'].nunique()} 家银行，"
          f"年份 {df['年份'].min()}—{df['年份'].max()}")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    nim_wide = to_wide(df, "净息差%")
    npl_wide = to_wide(df, "不良率%")

    plot_nim(nim_wide, outdir / CHART_NIM_NAME)
    plot_npl(npl_wide, outdir / CHART_NPL_NAME)

    summarize(df)
    print("完成：两张图已生成。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
