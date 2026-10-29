#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
战略台账查询工具 —— 让常见招商问题一句话就能查。

用法：
    python 查台账.py 到期 --品牌 ANTA --月数 12
    python 查台账.py 到期 --商户 斐乐 --月数 6
    python 查台账.py 品牌
    python 查台账.py 区域
    python 查台账.py 项目
    python 查台账.py 校验

示例：
    # 安踏系 12 个月内到期的合同
    python 查台账.py 到期 --品牌 ANTA,FILA --月数 12

    # 某商户全部在租合同
    python 查台账.py 合同 --商户 斐乐服饰
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 台账核心 import build_all  # noqa: E402

# 品牌集团归属（用于 --品牌 支持集团维度查询）
GROUPS = {
    "安踏": {"ANTA", "ANTA GUANJUN", "ANTA SUPERSTORE", "FILA", "FILA FUSION", "MAIA ACTIVE",
             "salomon", "saucony", "wilson", "Peak Performance"},
    "耐克": {"NIKE", "Jordan", "耐克儿童", "Nike Kicks Lounge", "ADIDAS ORIGINALS", "adidas",
             "Adidas Mega", "ADIDAS FOS", "Onitsuka Tiger"},
    "阿迪": {"adidas", "ADIDAS ORIGINALS", "Adidas Mega", "ADIDAS FOS", "Onitsuka Tiger"},
    "李宁": {"李宁", "李宁YOUNG", "中国李宁", "李宁户外"},
    "波司登": {"波司登"},
    "新百伦": {"New Balance", "New Balance Kids"},
    "骆驼": {"CAMEL"},
    "斐乐": {"FILA", "FILA FUSION"},
}


def _date(s: str | None) -> dt.date | None:
    return dt.date.fromisoformat(s) if s else None


def _expand_brands(names: list[str]) -> set[str]:
    """把集团名展开成品牌集合；未识别的按字面品牌名处理。"""
    out = set()
    for n in names:
        out |= GROUPS.get(n, {n})
    return out


def cmd_到期(args, data):
    c = data["合同级汇总"]
    today = dt.date.today()
    months = args.月数
    limit_days = months * 30 if months else 365

    sel = c
    if args.品牌:
        brands = _expand_brands(args.品牌.split(","))
        sel = [x for x in sel if x["品牌/行业"] in brands]
    if args.商户:
        sel = [x for x in sel if args.商户 in (x["商户名称"] or "")]
    if args.项目:
        sel = [x for x in sel if args.项目 in (x["项目名称"] or "")]
    if args.区域:
        sel = [x for x in sel if args.区域 in (x["所属区域"] or "")]

    soon = []
    for x in sel:
        e = _date(x["合同到期日"])
        if not e:
            continue
        days = (e - today).days
        if 0 <= days <= limit_days:
            soon.append((days, x))
    soon.sort(key=lambda t: t[0])

    label = "、".join(args.品牌.split(",")) if args.品牌 else "全部品牌"
    print(f"\n【{label}】未来 {months} 个月内到期合同：{len(soon)} 份\n")
    if not soon:
        print("  无")
        return
    total_area = sum(x["签约面积(平方米)"] or 0 for _, x in soon)
    total_cost = sum(x["期间成本合计(元)"] or 0 for _, x in soon)
    for i, (days, x) in enumerate(soon, 1):
        print(f"{i}. 【{x['品牌/行业']}】{x['项目名称']} {x.get('楼层','')} {x['资源位编号']}")
        print(f"   商户 {x['商户名称']} | {x['合同起租日']} → {x['合同到期日']}（剩 {days} 天）")
        print(f"   面积 {x['签约面积(平方米)']}㎡ | {x['所属区域']} | {x['三级业态']} | 成本 {x['期间成本合计(元)']:,.0f} 元")
    print(f"\n  合计 {len(soon)} 份 | 面积 {total_area:,.1f}㎡ | 期间成本 {total_cost:,.0f} 元")
    print("\n  按品牌:", dict(Counter(x["品牌/行业"] for _, x in soon)))
    print("  按区域:", dict(Counter(x["所属区域"] for _, x in soon)))
    print("  按项目:", dict(Counter(x["项目名称"] for _, x in soon)))


def cmd_合同(args, data):
    c = data["合同级汇总"]
    sel = c
    if args.商户:
        sel = [x for x in sel if args.商户 in (x["商户名称"] or "")]
    if args.品牌:
        sel = [x for x in sel if x["品牌/行业"] in _expand_brands(args.品牌.split(","))]
    if args.项目:
        sel = [x for x in sel if args.项目 in (x["项目名称"] or "")]
    print(f"\n共 {len(sel)} 份合同\n")
    for x in sel:
        print(f"  {x['合同编号']:22} {x['品牌/行业']:16} {x['项目名称']:14} {x.get('楼层','')} "
              f"{x['合同起租日']}→{x['合同到期日']} {x['签约面积(平方米)']}㎡ {x['期间成本合计(元)']:>14,.0f}")
    print("\n  按品牌:", dict(Counter(x["品牌/行业"] for x in sel)))
    print("  按商户:", dict(Counter(x["商户名称"] for x in sel)))
    print("  按区域:", dict(Counter(x["所属区域"] for x in sel)))


def cmd_品牌(args, data):
    for b in data["品牌汇总"]:
        print(f"  {b['品牌/行业']:20} 合同{b['合同数']:3}  项目{b['项目数']:2}  区域{b['所属区域数']}  "
              f"面积{b['签约面积合计(平方米)']:>9,.1f}㎡  成本{b['期间成本合计(元)']:>14,.0f}  "
              f"日租{b['平均单位日租金(元/平方米/日)'] or 0:>8.2f}  扣率{b['平均抽成扣率'] or 0:>7.2%}")


def cmd_区域(args, data):
    for r in data["区域业态汇总"]:
        print(f"  {r['所属区域']:12} {r['三级业态']:8} 合同{r['合同数']:4}  面积{r['签约面积合计(平方米)']:>9,.1f}㎡  "
              f"租金{r['期间固定租金总额(元)']:>14,.0f}  成本{r['期间成本合计(元)']:>14,.0f}  "
              f"日租{r['平均单位日租金(元/平方米/日)'] or 0:>8.2f}  扣率{r['平均抽成扣率'] or 0:>7.2%}")


def cmd_项目(args, data):
    agg = defaultdict(lambda: {"n": 0, "area": 0.0, "cost": 0.0})
    for x in data["合同级汇总"]:
        b = agg[x["项目名称"]]
        b["n"] += 1
        b["area"] += x["签约面积(平方米)"] or 0
        b["cost"] += x["期间成本合计(元)"] or 0
    for k, v in sorted(agg.items(), key=lambda x: -x[1]["cost"]):
        print(f"  {k:20} 合同{v['n']:3}  面积{v['area']:>10,.1f}㎡  成本{v['cost']:>14,.0f}")


def cmd_校验(args, data):
    d = data["合同台账明细"]
    cnt = Counter(x["数据校验"] for x in d)
    print(f"\n明细 {len(d)} 行校验分布：")
    for k, v in cnt.most_common():
        print(f"  {k:24} {v:4} 行 ({v/len(d)*100:.1f}%)")
    bad = [x for x in d if x["数据校验"] not in ("一致",)]
    if bad:
        print("\n非「一致」行的差异明细（前 20）：")
        for x in sorted(bad, key=lambda y: abs(y["固定租金校验差异(元)"] or 0), reverse=True)[:20]:
            print(f"  {x['项目名称']:14} {x['品牌/行业']:14} 期次{x['合同期间']}  "
                  f"差异{x['固定租金校验差异(元)'] or 0:>14,.2f}  {x['数据校验']}")


def main():
    ap = argparse.ArgumentParser(description="战略台账查询工具")
    ap.add_argument("命令", choices=["到期", "合同", "品牌", "区域", "项目", "校验"])
    ap.add_argument("--品牌", default="")
    ap.add_argument("--商户", default="")
    ap.add_argument("--项目", default="")
    ap.add_argument("--区域", default="")
    ap.add_argument("--月数", type=int, default=12)
    ap.add_argument("--源文件", default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "源数据", "运动男装合同台账1029.xlsx"))
    args = ap.parse_args()

    data = build_all(args.源文件)
    {"到期": cmd_到期, "合同": cmd_合同, "品牌": cmd_品牌,
     "区域": cmd_区域, "项目": cmd_项目, "校验": cmd_校验}[args.命令](args, data)


if __name__ == "__main__":
    main()
