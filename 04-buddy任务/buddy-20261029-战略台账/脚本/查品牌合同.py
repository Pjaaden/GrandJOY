#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
品牌合同查询（本地版）：按品牌查当前有效合同台账，输出 Excel。

用法：
    # 查一个或多个品牌的当前有效合同
    python 查品牌合同.py --品牌 NIKE,李宁
    python 查品牌合同.py --品牌 安踏          # 集团名自动展开
    python 查品牌合同.py --品牌 NIKE --含未起租   # 含尚未起租的合同
    python 查品牌合同.py --品牌 NIKE --到期窗口 90

输出：
    产出/品牌合同查询_<品牌>_<日期>.xlsx
    同时在终端打印表格预览。

口径（与飞书查询页一致）：
    当前有效 = 期次起租日 ≤ 今天 ≤ 期次到期日
    「合同最终到期日」= 同一合同编号下所有期次到期日的最大值（续约谈判看这个）
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 台账核心 import build_all, load_source  # noqa: E402

# 品牌集团归属（与 查台账.py 保持一致）
GROUPS = {
    "安踏": {"ANTA", "ANTA GUANJUN", "ANTA SUPERSTORE", "FILA", "FILA FUSION",
             "FILA KIDS", "MAIA ACTIVE"},
    "耐克": {"NIKE", "Jordan", "耐克儿童", "Nike Kicks Lounge"},
    "阿迪": {"adidas", "ADIDAS ORIGINALS", "Adidas Mega"},
    "李宁": {"李宁", "李宁YOUNG", "中国李宁", "李宁户外"},
    "波司登": {"波司登"},
    "新百伦": {"New Balance", "New Balance Kids"},
    "骆驼": {"CAMEL"},
    "斐乐": {"FILA", "FILA FUSION", "FILA KIDS"},
    "MLB": {"MLB", "MLB KIDS"},
    "安德玛": {"Under Armour", "UNDER ARMOUR KIDS", "UNDER ARMOUR YOUTH"},
}

OUT_COLS = [
    ("品牌/行业", None),
    ("项目名称", None),
    ("楼层", None),
    ("资源位编号", None),
    ("商户名称", None),
    ("合同编号", None),
    ("合同状态", None),
    ("签约面积(㎡)", "#,##0.00"),
    ("单位日租金(元/㎡/日)", "#,##0.0000"),
    ("抽成扣率", "0.00%"),
    ("物业费(元/㎡/月)", "#,##0.0000"),
    ("推广费(元/㎡/月)", "#,##0.0000"),
    ("期次起租日", "yyyy-mm-dd"),
    ("期次到期日", "yyyy-mm-dd"),
    ("合同最终到期日", "yyyy-mm-dd"),
    ("距最终到期天数", "0"),
    ("所属区域", None),
    ("三级业态", None),
    ("是否战略品牌", None),
    ("合同份数", "0"),
]


def D(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and v.strip():
        try:
            return dt.date.fromisoformat(v.strip()[:10])
        except ValueError:
            return None
    return None


def expand(names):
    """把集团名展开成品牌集合；未识别的按字面品牌名处理。"""
    out = set()
    for n in names:
        n = n.strip()
        if not n:
            continue
        out |= GROUPS.get(n, {n})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--品牌", required=True, help="品牌名，逗号分隔；支持集团名")
    ap.add_argument("--源文件", default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "源数据", "运动男装合同台账1103.xlsx"))
    ap.add_argument("--含未起租", action="store_true",
                    help="同时包含尚未起租但未到期的合同（默认只看正在履约的）")
    ap.add_argument("--到期窗口", type=int, default=0,
                    help="只看 N 天内到期的（0=不限）")
    ap.add_argument("--输出目录", default=None)
    args = ap.parse_args()

    names = [x for x in args.品牌.replace("，", ",").split(",") if x.strip()]
    # 品牌名可能含空格（如 "FILA KIDS"）→ 支持用 | 分隔多个品牌
    expanded = []
    for chunk in names:
        expanded += [c for c in chunk.split("|") if c.strip()]
    names = expanded
    brands = expand(names)
    today = dt.date.today()

    data = build_all(args.源文件)
    _, meta = load_source(args.源文件)
    detail = data["合同台账明细"]

    # 合同最终到期日 = 同合同编号下期次到期日最大值
    contract_end = {}
    for d in detail:
        cid = d["合同编号"]
        e = D(d["到期日"])
        if cid and e and (cid not in contract_end or e > contract_end[cid]):
            contract_end[cid] = e
    # 每份合同的份数
    period_cnt = {}
    for d in detail:
        cid = d["合同编号"]
        period_cnt[cid] = period_cnt.get(cid, 0) + 1

    rows = []
    for d in detail:
        if d["品牌/行业"] not in brands:
            continue
        s, e = D(d["起租日"]), D(d["到期日"])
        if not s or not e:
            continue
        if e < today:
            continue                       # 已到期
        if not args.含未起租 and s > today:
            continue                       # 尚未起租
        cend = contract_end.get(d["合同编号"])
        if args.到期窗口 and cend and (cend - today).days > args.到期窗口:
            continue
        r = dict(d)
        r["合同最终到期日"] = cend
        r["距最终到期天数"] = (cend - today).days if cend else None
        r["合同份数"] = period_cnt.get(d["合同编号"], 1)
        rows.append(r)

    rows.sort(key=lambda x: (x["品牌/行业"], x["合同最终到期日"] or dt.date.max))

    # 终端预览
    print(f"品牌：{args.品牌}（展开 {len(brands)} 个：{', '.join(sorted(brands))}）")
    print(f"口径：{'当前履约 + 未起租' if args.含未起租 else '当前正在履约（起租日 ≤ 今天 ≤ 到期日）'}"
          f"{f'，{args.到期窗口}天内到期' if args.到期窗口 else ''}")
    print(f"数据版本 {meta.get('stat_date')}　|　命中 {len(rows)} 份期次合同\n")
    hdr = f"{'品牌':<14}{'项目':<16}{'铺位':<14}{'面积':>9} {'日租':>7} {'最终到期':>12} {'剩余':>6}"
    print(hdr)
    print("-" * len(hdr))
    for r in rows[:40]:
        print(f"{str(r['品牌/行业']):<14}{str(r['项目名称']):<16}{str(r['资源位编号']):<14}"
              f"{r['签约面积'] or 0:>9.2f} {r['单位日租金(元/平方米/日)'] or 0:>7.2f} "
              f"{str(r['合同最终到期日']):>12} {r['距最终到期天数']:>6}")
    if len(rows) > 40:
        print(f"... 共 {len(rows)} 行（完整结果见 Excel）")

    # 汇总
    area = sum(r["签约面积"] or 0 for r in rows)
    wsum = sum((r["签约面积"] or 0) * (r["单位日租金(元/平方米/日)"] or 0) for r in rows)
    print(f"\n合计：面积 {area:,.2f} ㎡　|　加权日租 {(wsum/area if area else 0):.2f} 元/㎡/日"
          f"　|　合同份数 {len({r['合同编号'] for r in rows})}")
    if rows:
        nearest = min(r["合同最终到期日"] for r in rows if r["合同最终到期日"])
        print(f"最早到期：{nearest}（剩 {(nearest - today).days} 天）")

    # 输出 Excel
    out_dir = args.输出目录 or os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "产出")
    os.makedirs(out_dir, exist_ok=True)
    tag = "_".join(n.replace(" ", "") for n in names) or "全部"
    path = os.path.join(out_dir, f"品牌合同查询_{tag}_{today:%Y%m%d}.xlsx")
    write_xlsx(path, rows, args.品牌, today, meta.get("stat_date"))
    print(f"\n已生成：{path}")


def write_xlsx(path, rows, brand_arg, today, version):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "品牌合同查询"

    ws["A1"] = f"品牌合同查询 · {brand_arg}"
    ws["A1"].font = Font(size=14, bold=True, color="1F3864")
    ws["A2"] = (f"数据版本 {version}　|　查询日期 {today:%Y-%m-%d}　|　"
                f"命中 {len(rows)} 份期次合同　|　"
                "口径：起租日 ≤ 今天 ≤ 到期日（当前正在履约）")
    ws["A2"].font = Font(size=10, color="666666")

    hr = 4
    for i, (h, _) in enumerate(OUT_COLS, start=1):
        c = ws.cell(row=hr, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(horizontal="center", vertical="center")

    for ri, r in enumerate(rows, start=hr + 1):
        for ci, (h, fmt) in enumerate(OUT_COLS, start=1):
            v = r.get(h)
            if v is None:
                continue
            cell = ws.cell(row=ri, column=ci, value=v)
            if fmt:
                cell.number_format = fmt
            if isinstance(v, dt.date):
                cell.number_format = "yyyy-mm-dd"

    widths = [16, 18, 14, 15, 26, 20, 12, 12, 18, 12, 18, 18, 13, 13, 15, 14, 16, 14, 12, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = f"A{hr+1}"
    ws.auto_filter.ref = f"A{hr}:{openpyxl.utils.get_column_letter(len(OUT_COLS))}{hr + len(rows)}"

    # 汇总页
    area = sum(r["签约面积"] or 0 for r in rows)
    wnum = sum((r["签约面积"] or 0) * (r["单位日租金(元/平方米/日)"] or 0) for r in rows)
    sm = wb.create_sheet("汇总")
    sm["A1"] = f"品牌合同查询汇总 · {brand_arg}"
    sm["A1"].font = Font(size=14, bold=True, color="1F3864")
    sm["A2"] = f"查询日期 {today:%Y-%m-%d}　|　数据版本 {version}"
    sm["A2"].font = Font(size=10, color="666666")
    kpis = [
        ("命中期次合同", len(rows), "#,##0"),
        ("涉及合同份数", len({r["合同编号"] for r in rows}), "#,##0"),
        ("涉及品牌数", len({r["品牌/行业"] for r in rows}), "#,##0"),
        ("签约面积合计(㎡)", round(area, 2), "#,##0.00"),
        ("加权日租(元/㎡/日)", round(wnum / area, 4) if area else 0, "#,##0.0000"),
        ("战略品牌占比",
         f"{sum(1 for r in rows if r.get('是否战略品牌') == '是')}/{len(rows)}"
         if rows else "-", None),
    ]
    for i, (k, v, fmt) in enumerate(kpis, start=4):
        sm.cell(row=i, column=1, value=k).font = Font(size=10, color="666666")
        c = sm.cell(row=i, column=2, value=v)
        c.font = Font(size=13, bold=True, color="1F3864")
        if fmt:
            c.number_format = fmt
    sm.column_dimensions["A"].width = 22
    sm.column_dimensions["B"].width = 18

    wb.save(path)


if __name__ == "__main__":
    main()
