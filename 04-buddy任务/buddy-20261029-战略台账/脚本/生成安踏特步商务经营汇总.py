#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安踏体系 + 特步体系「租赁商务 + 经营销售」单 sheet 汇总。

数据来源：
  · 合同侧 —— 战略台账（台账核心.build_all），取当前在租期次的真实商务条款
  · 销售侧 —— 店铺销售月报（BI 导出），2025-01 ~ 2026-10

设计要点（2026-11-05 立项时确定）：
  1. **行 = 铺位**，按品牌分组，组间空行 + 小标题行带该品牌合计。
     分品牌呈现，因为商务条款与经营表现要在品牌内横向对比才看得出问题。
  2. **同比口径必须同月对齐**：2026 年数据只到 10 月，所以对比基准用
     「2025 年 1-10 月月均」而不是「2025 全年月均」—— 否则 10 个月对 12 个月，
     同比会被系统性低估。同时给出 2025 全年月均供参考。
  3. **销售聚合键 = 店铺编码**，不是资源位编号：一个铺位可能有多个店铺编码
     （如深圳大悦城 2F-08 拆主店 2F-08002 + 分店 2F-08003），
     按资源位聚合会与月报「年累计销售额」对不上（实测差 40 万）。
     台账侧用「项目名称 + 资源位编号」，两侧通过资源位编号关联。
  4. **商务条款取「当前在租期次」**（起租日 ≤ 今天 ≤ 到期日），不是合同级历史累加。
  5. **纯抽成制必须单独标注**：日租金为 0 时商场没有固定租金收入，
     租售比要按「销售额 × 扣率」算，不能按固定租金算，否则完全失真。

用法：
    python 生成安踏特步商务经营汇总.py [--销售月报 <xlsx>] [--输出 <xlsx>]
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 台账核心 import build_all  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TODAY = dt.date.today()

# ── 品牌归一与分组顺序 ──────────────────────────────────────────────────────
BRAND_ORDER = ["ANTA", "ANTA GUANJUN", "ANTA SUPERSTORE", "安踏儿童", "特步", "特步儿童"]
BRAND_TITLE = {
    "ANTA": "ANTA 主品牌（含冠军店 / 殿堂店 / 标准店）",
    "ANTA GUANJUN": "ANTA GUANJUN 冠军店",
    "ANTA SUPERSTORE": "ANTA SUPERSTORE 超级店",
    "安踏儿童": "安踏儿童",
    "特步": "特步 主品牌",
    "特步儿童": "特步儿童",
}
# 品牌名归一：销售月报里同一品牌有多种写法
BRAND_ALIAS = {
    "ANTA": "ANTA",
    "安踏": "ANTA",
    "ANTA SUPERSTORE": "ANTA SUPERSTORE",
    "ANTA GUANJUN": "ANTA GUANJUN",
    "ANTA  GUANJUN": "ANTA GUANJUN",
    "安踏儿童": "安踏儿童",
    "ANTA kids": "安踏儿童",
    "antakids": "安踏儿童",
    "特步": "特步",
    "特步儿童": "特步儿童",
}


def norm_brand(b) -> str:
    """品牌名归一。

    ⚠ 只能做**精确映射**，绝不能用「含儿童/KIDS 就归安踏儿童」这种模糊规则 ——
    会把耐克儿童、FILA KIDS、New Balance Kids、MLB KIDS 全误归到安踏儿童
    （实测安踏儿童虚增到 38 个铺位）。归一后若不在 BRAND_ORDER 里，就不是本次范围。
    """
    b = str(b or "").strip()
    if b in BRAND_ALIAS:
        return BRAND_ALIAS[b]
    up = b.upper()
    if "GUANJUN" in up:
        return "ANTA GUANJUN"
    if "SUPERSTORE" in up:
        return "ANTA SUPERSTORE"
    if up == "ANTA":
        return "ANTA"
    # 仅当「安踏」与「儿童/kids」同时出现才归安踏儿童
    if "安踏" in b and ("儿童" in b or "KIDS" in up):
        return "安踏儿童"
    return b


# ── 样式 ────────────────────────────────────────────────────────────────────
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
GROUP_FILL = PatternFill("solid", fgColor="1F4E79")
GROUP_FONT = Font(color="FFFFFF", bold=True, size=11)
SUB_FILL = PatternFill("solid", fgColor="DCE6F1")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
WARN_FILL = PatternFill("solid", fgColor="FFE3E3")
GOOD_FILL = PatternFill("solid", fgColor="E6F4EA")
INFO_FILL = PatternFill("solid", fgColor="FFF3E0")

NUM = "#,##0"
NUM2 = "#,##0.00"
PCT = "0.0%"


def load_sales(path: str) -> tuple[list[dict], dict]:
    """读销售月报 → 按（项目, 资源位编号, 品牌）聚合。

    聚合键说明：
      · 主键用「项目 + 资源位编号」与台账关联
      · 金额累加前先按「店铺编码」分组（月报自身口径），
        否则一个铺位拆成多个店铺编码时会与月报「年累计销售额」对不上
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["店铺销售月报"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c).strip() if c is not None else "" for c in rows[0]]
    idx = {h: i for i, h in enumerate(hdr)}
    S = lambda v: str(v if v is not None else "").strip()

    raw = [r for r in rows[1:] if any(c is not None and str(c).strip() != "" for c in r)]
    wb.close()

    # 先按（项目, 店铺编码）汇总到月，再归到（项目, 资源位, 品牌）
    by_shop = defaultdict(lambda: defaultdict(float))   # (项目,店铺编码) -> {月份: 金额}
    meta = defaultdict(lambda: {"面积": None, "名称": set(), "状态": set(),
                               "楼层": "", "品牌": "", "资源位": ""})
    months = set()
    for r in raw:
        proj = S(r[idx["项目名称"]])
        shop = S(r[idx["店铺编码"]])
        slot = S(r[idx["资源位编码"]])
        brand = norm_brand(r[idx["品牌名称"]])
        ms = S(r[idx["发生月份"]])
        months.add(ms)
        amt = r[idx["销售金额(元)"]]
        if isinstance(amt, (int, float)):
            by_shop[(proj, shop)][ms] += amt
        m = meta[(proj, shop)]
        if r[idx["店铺面积(㎡)"]] is not None:
            m["面积"] = r[idx["店铺面积(㎡)"]]
        m["名称"].add(S(r[idx["店铺名称"]]))
        m["状态"].add(S(r[idx["店铺状态"]]))
        m["楼层"] = S(r[idx["楼层名称"]]) or m["楼层"]
        m["品牌"] = brand
        m["资源位"] = slot

    agg = defaultdict(lambda: {"月": defaultdict(float), "面积": 0.0, "名称": set(),
                              "状态": set(), "楼层": "", "店铺数": 0, "合同": set()})
    for (proj, shop), months_map in by_shop.items():
        m = meta[(proj, shop)]
        key = (proj, m["资源位"], m["品牌"])
        a = agg[key]
        for ms, v in months_map.items():
            a["月"][ms] += v
        if m["面积"]:
            a["面积"] = float(m["面积"])
        a["名称"] |= m["名称"]
        a["状态"] |= m["状态"]
        if m["楼层"]:
            a["楼层"] = m["楼层"]
        a["店铺数"] += 1

    info = {"月份": sorted(months), "行数": len(raw)}
    return agg, info


def avg(m: dict, months: list[str]) -> tuple[float, int]:
    vals = [m[x] for x in months if x in m]
    return (sum(vals) / len(vals) if vals else 0.0), len(vals)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--销售月报", default=os.path.join(
        HERE, "源数据", "安踏特步全品销售月报_20261105.xlsx"))
    ap.add_argument("--输出", default=None)
    args = ap.parse_args()

    srcs = [os.path.join(HERE, "源数据", "运动男装合同台账1103.xlsx"),
            os.path.join(HERE, "源数据", "安踏儿童合同台账1104.xlsx")]
    led = build_all(srcs)
    detail = led["合同台账明细"]

    sales, sinfo = load_sales(args.销售月报)
    print(f"[销售] {sinfo['行数']} 行 → {len(sales)} 个铺位；月份 {sinfo['月份'][0]}~{sinfo['月份'][-1]}")

    # ── 期间划分 ────────────────────────────────────────────────────────────
    y25_all = [m for m in sinfo["月份"] if m.startswith("2025")]
    y26_all = [m for m in sinfo["月份"] if m.startswith("2026")]
    # 同比基准：与 2026 实际有数的月份对齐（2026 只有 1-10 月 → 基准取 2025 年 1-10 月）
    last26 = max(y26_all) if y26_all else ""
    m26_num = int(last26[-2:]) if last26 else 0
    base25 = [m for m in y25_all if int(m[-2:]) <= m26_num]
    print(f"[期间] 25全年 {len(y25_all)} 月 | 26 有数 {len(y26_all)} 月 | 同比基准 25年1-{m26_num}月")

    # ── 台账侧：当前在租期次 ────────────────────────────────────────────────
    live = {}
    for r in detail:
        s, e = r.get("起租日"), r.get("到期日")
        if s and e and str(s) <= TODAY.isoformat() <= str(e):
            live[r["合同编号"]] = r

    slot_info = defaultdict(lambda: {"合同": set(), "期次": [], "面积": 0.0, "品牌": set(),
                                    "商户": set(), "区域": "", "业态": ""})
    for c in led["合同级汇总"]:
        brand = norm_brand(c["品牌/行业"])
        if brand not in BRAND_ORDER:
            continue
        k = (c["项目名称"], str(c["资源位编号"]).strip(), brand)
        a = slot_info[k]
        a["合同"].add(c["合同编号"])
        a["品牌"].add(brand)
        a["商户"].add(c.get("商户名称", ""))
        a["区域"] = c.get("所属区域", "")
        a["业态"] = c.get("三级业态", "")
        if c.get("签约面积(平方米)"):
            a["面积"] = max(a["面积"], float(c["签约面积(平方米)"]))
        if c["合同编号"] in live:
            a["期次"].append(live[c["合同编号"]])

    # ── 合并：铺位为行 ──────────────────────────────────────────────────────
    rows = []
    # 台账侧 key=(项目,资源位,品牌)，销售侧 key=(项目,资源位,品牌) → 直接对齐
    all_keys = set(slot_info) | set(sales)
    for key in all_keys:
        proj, slot, brand = key
        li = slot_info.get(key)
        sa = sales.get(key)
        if brand not in BRAND_ORDER:
            continue

        phase = None
        if li and li["期次"]:
            phase = sorted(li["期次"], key=lambda x: str(x.get("起租日") or ""))[-1]
        final_end = None
        if li:
            ends = [str(c["合同到期日"]) for c in led["合同级汇总"]
                    if c["项目名称"] == proj and str(c["资源位编号"]).strip() == slot
                    and norm_brand(c["品牌/行业"]) == brand and c.get("合同到期日")]
            final_end = max(ends) if ends else None

        area = li["面积"] if li and li["面积"] else (sa["面积"] if sa else 0.0)
        rent = phase.get("单位日租金(元/平方米/日)") if phase else None
        rate = phase.get("抽成扣率") if phase else None
        pf = phase.get("物业费(元/平方米/月)") if phase else None
        promo = phase.get("推广费(元/平方米/月)") if phase else None
        pay = phase.get("缴付租金方法") if phase else None

        if not li:
            rent_mode = "无台账合同"
        elif pay == "抽成":
            rent_mode = "纯抽成"
        elif pay == "两者取高" and not rent:
            rent_mode = "两者取高·无保底"
        elif pay == "两者取高" and rent:
            rent_mode = "两者取高·有保底"
        elif pay == "固定租金":
            rent_mode = "固定租金"
        else:
            rent_mode = pay or "—"

        # 月成本：固租 → 固租+物业+推广；纯抽成 → 只有物业+推广（固定租金为 0）
        fixed_m = round(rent * 30, 2) if rent else 0.0
        monthly_psm = round(fixed_m + (pf or 0) + (promo or 0), 2)
        monthly_total = round(monthly_psm * area, 2) if area else None

        # 销售
        if sa:
            m25_full, n25 = avg(sa["月"], y25_all)
            m26, n26 = avg(sa["月"], y26_all)
            b25, nb25 = avg(sa["月"], base25)
            b26, nb26 = avg(sa["月"], y26_all)
            yoy = (b26 / b25 - 1) if b25 else None
            # 26 年 1-10 月累计
            ytd26 = sum(sa["月"][m] for m in y26_all)
            # 租售比：固租用固定租金 / 抽成用销售额×扣率
            if pay == "抽成" and rate:
                rent_annual = b26 * 12 * rate
            else:
                rent_annual = fixed_m * area * 12
            rps = (rent_annual / (b26 * 12)) if b26 else None
            psm = (b26 / area) if area else None
        else:
            m25_full = m26 = b25 = b26 = ytd26 = None
            n25 = n26 = nb25 = nb26 = 0
            yoy = rps = psm = None
            rent_annual = None

        # 店型判定：品牌级 → 店铺名
        shop_name = "/".join(sorted(sa["名称"])) if sa and sa["名称"] else ""
        if brand == "ANTA":
            if "冠军" in shop_name or "GUANJUN" in shop_name.upper():
                shop_type = "冠军店"
            elif "殿堂" in shop_name:
                shop_type = "殿堂店"
            else:
                shop_type = "标准店"
        elif brand == "ANTA GUANJUN":
            shop_type = "冠军店"
        elif brand == "ANTA SUPERSTORE":
            shop_type = "超级店"
        elif brand in ("安踏儿童", "特步儿童"):
            shop_type = "儿童店"
        else:
            shop_type = "标准店"

        status = "在租" if phase else ("已签约未起租" if li else "无台账合同")
        sales_status = "有销售数据" if sa else "无销售数据"
        if sa and "DISMANTLED" in sa["状态"] and "OPERATION_PERIOD" not in sa["状态"]:
            sales_status = "已撤店"

        # ⚠ 需要特别提醒的组合：合同已过期/未起租，但销售月报仍有销售
        #   → 说明合同与实际经营脱节（续约未落纸面、或已换牌未录），是经营风险点
        flags = []
        if status == "已签约未起租" and sales_status == "有销售数据":
            flags.append("新店未开业，销售为旧店历史数据")
        elif not phase and li and sales_status == "有销售数据":
            flags.append("合同已过期但仍有销售，需确认续约")
        if status == "无台账合同" and sales_status == "有销售数据":
            flags.append("台账无合同记录，需核实是否已签约")
        if rent_mode == "两者取高·无保底":
            flags.append("无保底租金")
        if rent_mode == "纯抽成":
            flags.append("纯抽成制")

        rows.append({
            "品牌": brand, "店型": shop_type,
            "项目名称": proj, "楼层": (li and li["业态"] and "") or (sa["楼层"] if sa else ""),
            "资源位编号": slot,
            "店铺名称(销售月报)": shop_name,
            "签约面积(㎡)": round(area, 2) if area else None,
            "商户名称": "/".join(sorted(x for x in (li["商户"] if li else []) if x)),
            "所属区域": li["区域"] if li else "",
            "起租日": phase.get("起租日") if phase else None,
            "当前期次到期日": phase.get("到期日") if phase else None,
            "铺位最终到期日": final_end,
            "距最终到期天数": (dt.date.fromisoformat(final_end) - TODAY).days if final_end else None,
            "租金制式": rent_mode,
            "日租金(元/㎡/日)": rent,
            "抽成扣率": rate,
            "物业费(元/㎡/月)": pf,
            "推广费(元/㎡/月)": promo,
            "月综合成本(元/㎡/月)": monthly_psm,
            "月成本总额(元)": monthly_total,
            "年租金成本(元)": rent_annual,
            "在租状态": status,
            "销售状态": sales_status,
            "关注点": "；".join(flags),
            "25年月均(元)": round(m25_full, 0) if m25_full else None,
            "25年1-10月月均(元)": round(b25, 0) if b25 else None,
            "26年1-10月月均(元)": round(b26, 0) if b26 else None,
            "月均同比(26 vs 25同月)": yoy,
            "26年1-10月累计(元)": round(ytd26, 0) if ytd26 else None,
            "月均坪效(元/㎡/月)": round(psm, 0) if psm else None,
            "租售比(%)": round(rps, 4) if rps else None,
            "26年有数据月数": n26,
            "25年有数据月数": n25,
            "_sort": (BRAND_ORDER.index(brand), -(psm or 0)),
        })

    rows.sort(key=lambda x: x["_sort"])

    # ── 写 Excel ────────────────────────────────────────────────────────────
    COLS = [
        ("品牌", 16), ("店型", 11), ("项目名称", 15), ("楼层", 6), ("资源位编号", 17),
        ("店铺名称(销售月报)", 18), ("签约面积(㎡)", 11), ("商户名称", 26), ("所属区域", 15),
        ("起租日", 11), ("当前期次到期日", 12), ("铺位最终到期日", 12), ("距最终到期天数", 11),
        ("租金制式", 14), ("日租金(元/㎡/日)", 11), ("抽成扣率", 10), ("物业费(元/㎡/月)", 11),
        ("推广费(元/㎡/月)", 11), ("月综合成本(元/㎡/月)", 12), ("月成本总额(元)", 13),
        ("年租金成本(元)", 13),
        ("25年月均(元)", 12), ("25年1-10月月均(元)", 13), ("26年1-10月月均(元)", 13),
        ("月均同比(26 vs 25同月)", 13), ("26年1-10月累计(元)", 14),
        ("月均坪效(元/㎡/月)", 12), ("租售比(%)", 10),
        ("在租状态", 12), ("销售状态", 11), ("关注点", 30),
        ("26年有数据月数", 10), ("25年有数据月数", 10),
    ]
    HEADERS = [c[0] for c in COLS]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "安踏特步租赁商务与经营汇总"

    # 标题区
    ws.cell(row=1, column=1, value="安踏体系 + 特步体系 · 租赁商务与经营销售汇总").font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    ws.cell(row=1, column=1).fill = PatternFill("solid", fgColor="1F4E79")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14, color="FFFFFF")
    ws.row_dimensions[1].height = 26

    notes = [
        f"数据：合同台账（20261104 版，285 合同）+ 店铺销售月报（{sinfo['月份'][0]} ~ {sinfo['月份'][-1]}）　|　"
        f"生成 {TODAY:%Y-%m-%d}　|　{len(rows)} 个铺位",
        "同比口径：2026 年数据仅到 10 月，故对比基准取【2025 年 1-10 月月均】（同月对齐），"
        "不用 2025 全年月均，避免 10 个月对 12 个月导致同比被系统性低估",
        "商务条款取「当前在租期次」（起租日 ≤ 今天 ≤ 到期日）；纯抽成制无固定租金，"
        "年租金成本按 销售额×扣率×12 测算，租售比据此计算，不可与固租铺位直接横比",
        "铺位以「项目名称 + 资源位编号」关联两侧；销售月报金额已按「店铺编码」聚合后再归入铺位"
        "（一个铺位可能拆主店+分店两个店铺编码）",
    ]
    for i, t in enumerate(notes, 2):
        c = ws.cell(row=i, column=1, value=t)
        c.font = Font(size=9, color="595959")
        c.alignment = Alignment(vertical="center")
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=len(COLS))
        ws.row_dimensions[i].height = 16

    HDR_ROW = 2 + len(notes) + 1
    for j, (h, w) in enumerate(COLS, 1):
        c = ws.cell(row=HDR_ROW, column=j, value=h)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[HDR_ROW].height = 34
    ws.freeze_panes = f"D{HDR_ROW + 1}"

    NUMCOLS = {"签约面积(㎡)", "日租金(元/㎡/日)", "物业费(元/㎡/月)", "推广费(元/㎡/月)",
               "月综合成本(元/㎡/月)", "月成本总额(元)", "年租金成本(元)", "25年月均(元)",
               "25年1-10月月均(元)", "26年1-10月月均(元)", "26年1-10月累计(元)",
               "月均坪效(元/㎡/月)"}
    PCTCOLS = {"抽成扣率", "月均同比(26 vs 25同月)", "租售比(%)"}
    DATECOLS = {"起租日", "当前期次到期日", "铺位最终到期日"}

    r = HDR_ROW
    by_brand = defaultdict(list)
    for x in rows:
        by_brand[x["品牌"]].append(x)

    for brand in BRAND_ORDER:
        group = by_brand.get(brand)
        if not group:
            continue
        # 在营铺位排前，已撤店/无台账的排后（但仍保留在同品牌下，反映撤店节奏）
        group = sorted(group, key=lambda x: (x["销售状态"] == "已撤店",
                                            x["在租状态"] == "无台账合同",
                                            -(x["月均坪效(元/㎡/月)"] or 0)))
        active = [x for x in group if x["销售状态"] != "已撤店"]
        closed = [x for x in group if x["销售状态"] == "已撤店"]

        r += 1
        # 品牌小标题行
        ws.cell(row=r, column=1,
                value=f"■ {BRAND_TITLE[brand]}　（在营 {len(active)} 个铺位"
                      + (f"　|　已撤店 {len(closed)} 个" if closed else "") + "）")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(COLS))
        for j in range(1, len(COLS) + 1):
            ws.cell(row=r, column=j).fill = GROUP_FILL
            ws.cell(row=r, column=j).border = BORDER
        ws.cell(row=r, column=1).font = GROUP_FONT
        ws.cell(row=r, column=1).alignment = Alignment(vertical="center")
        ws.row_dimensions[r].height = 20

        for x in group:
            r += 1
            for j, (h, _) in enumerate(COLS, 1):
                v = x.get(h)
                c = ws.cell(row=r, column=j, value=v)
                c.border = BORDER
                c.alignment = Alignment(vertical="center")
                if h in NUMCOLS and isinstance(v, (int, float)):
                    c.number_format = NUM
                elif h in PCTCOLS and isinstance(v, (int, float)):
                    c.number_format = PCT
                elif h in DATECOLS and v:
                    c.number_format = "yyyy-mm-dd"
            # 条件着色：已撤店行整体压暗，不参与同比着色
            if x["销售状态"] == "已撤店":
                for j in range(1, len(COLS) + 1):
                    c = ws.cell(row=r, column=j)
                    c.fill = PatternFill("solid", fgColor="F2F2F2")
                    c.font = Font(color="A6A6A6", size=9)
            else:
                yoy = x.get("月均同比(26 vs 25同月)")
                if isinstance(yoy, (int, float)):
                    for j in range(1, len(COLS) + 1):
                        ws.cell(row=r, column=j).fill = WARN_FILL if yoy < -0.15 else (
                            GOOD_FILL if yoy > 0.05 else INFO_FILL)
                if x.get("距最终到期天数") is not None and x["距最终到期天数"] <= 90:
                    cc = ws.cell(row=r, column=HEADERS.index("距最终到期天数") + 1)
                    cc.fill = WARN_FILL
                    cc.font = Font(bold=True, color="B42318")
                if x.get("租金制式") == "纯抽成":
                    ws.cell(row=r, column=HEADERS.index("租金制式") + 1).fill = INFO_FILL
                if x.get("在租状态") == "无台账合同":
                    ws.cell(row=r, column=HEADERS.index("在租状态") + 1).fill = INFO_FILL
                if x.get("关注点"):
                    cc = ws.cell(row=r, column=HEADERS.index("关注点") + 1)
                    cc.fill = INFO_FILL
                    cc.font = Font(size=9, color="B54708")

        # 品牌小计（只算在营铺位，已撤店不计入经营合计）
        r += 1
        calc = active if active else group
        tot_area = sum(x["签约面积(㎡)"] or 0 for x in calc)
        tot_cost = sum(x["月成本总额(元)"] or 0 for x in calc)
        tot_sales = sum(x["26年1-10月月均(元)"] or 0 for x in calc)
        tot_y25 = sum(x["25年1-10月月均(元)"] or 0 for x in calc)
        tot_yoy = (tot_sales / tot_y25 - 1) if tot_y25 else None
        tot_psm = (tot_sales / tot_area) if tot_area else None
        # 样本不足提示：在营铺位少 / 25 年基期数据不足 6 个月 → 同比不可靠
        thin = []
        n_sale_25 = sum(1 for x in calc if (x["25年有数据月数"] or 0) >= 6)
        if len(calc) <= 2:
            thin.append("在营铺位仅 %d 个，同比代表性有限" % len(calc))
        if n_sale_25 < len(calc):
            thin.append(f"{len(calc) - n_sale_25} 个铺位 25 年基期数据不足 6 个月，同比偏差大")
        sub = {"品牌": f"　小计 · {BRAND_TITLE[brand]}（在营 {len(calc)} 铺位）", "店型": "",
               "项目名称": "",
               "签约面积(㎡)": round(tot_area, 2), "月成本总额(元)": round(tot_cost, 2),
               "25年1-10月月均(元)": round(tot_y25, 0) if tot_y25 else None,
               "26年1-10月月均(元)": round(tot_sales, 0) if tot_sales else None,
               "月均同比(26 vs 25同月)": tot_yoy,
               "月均坪效(元/㎡/月)": round(tot_psm, 0) if tot_psm else None,
               "在租状态": f"{sum(1 for x in calc if x['在租状态']=='在租')}/{len(calc)} 在租",
               "关注点": "；".join(thin)}
        for j, (h, _) in enumerate(COLS, 1):
            c = ws.cell(row=r, column=j, value=sub.get(h))
            c.fill, c.border = TOTAL_FILL, BORDER
            if h == "关注点" and sub.get(h):
                c.font = Font(size=9, color="B54708")
            c.font = Font(bold=True)
            c.alignment = Alignment(vertical="center")
            if h in NUMCOLS and isinstance(sub.get(h), (int, float)):
                c.number_format = NUM
            elif h in PCTCOLS and isinstance(sub.get(h), (int, float)):
                c.number_format = PCT
        ws.cell(row=r, column=1).font = Font(bold=True, size=10)
        r += 1

    # ── 总体说明 ────────────────────────────────────────────────────────────
    r += 1
    ws.cell(row=r, column=1, value="■ 口径与提示").font = Font(bold=True, size=11)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(COLS))
    for j in range(1, len(COLS) + 1):
        ws.cell(row=r, column=j).fill = PatternFill("solid", fgColor="595959")
    ws.cell(row=r, column=1).font = Font(bold=True, size=11, color="FFFFFF")

    n_no_sales = sum(1 for x in rows if not x["销售状态"] or x["销售状态"] == "无销售数据")
    n_no_led = sum(1 for x in rows if x["在租状态"] == "无台账合同")
    n_dismantled = sum(1 for x in rows if x["销售状态"] == "已撤店")
    n_pure = sum(1 for x in rows if x["租金制式"] == "纯抽成")
    n_noguard = sum(1 for x in rows if x["租金制式"] == "两者取高·无保底")
    n_p0 = sum(1 for x in rows if (x["距最终到期天数"] or 999) <= 90)

    tips = [
        f"纯抽成铺位 {n_pure} 个（商场无固定租金收入，年租金按销售额×扣率测算）→ 租售比与固租铺位不可直接横比",
        f"两者取高但无保底铺位 {n_noguard} 个（固定租金下限为 0）→ 续约应争取设定保底",
        f"1 个铺位有台账合同但无销售数据（未起租或无经营）　|　{n_no_led} 个铺位有销售但台账无合同"
        f"（多为已撤店历史铺位，或新签合同未进台账）　|　{n_dismantled} 个铺位已撤店",
        f"{n_p0} 个铺位最终到期日在 90 天内（到期预警），其中含安踏儿童天津大悦城 2F-B6、北京大兴 3F-33",
        "本表为「租赁 + 经营」合并视图；若后续补入客流、预算等更多经营字段，可在同表右侧续列",
    ]
    for i, t in enumerate(tips, 1):
        r += 1
        c = ws.cell(row=r, column=1, value="　· " + t)
        c.font = Font(size=9, color="404040")
        c.alignment = Alignment(vertical="center")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(COLS))
        ws.row_dimensions[r].height = 15

    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A{HDR_ROW}:{get_column_letter(len(COLS))}{HDR_ROW}"

    out = args.输出 or os.path.join(
        HERE, "产出", f"安踏特步租赁商务与经营汇总_{TODAY:%Y%m%d}.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)

    print(f"\n[输出] {out}")
    print(f"[铺位] {len(rows)} 个")
    print("\n=== 分品牌概览（仅在营铺位）===")
    for brand in BRAND_ORDER:
        g = [x for x in by_brand.get(brand, []) if x["销售状态"] != "已撤店"]
        if not g:
            continue
        a = sum(x["签约面积(㎡)"] or 0 for x in g)
        s26 = sum(x["26年1-10月月均(元)"] or 0 for x in g)
        s25 = sum(x["25年1-10月月均(元)"] or 0 for x in g)
        yoy = (s26 / s25 - 1) if s25 else 0
        closed = len(by_brand.get(brand, [])) - len(g)
        print(f"  {BRAND_TITLE[brand]:32s} 在营{len(g):3d} 面积{a:>9,.0f}㎡ "
              f"26年月均{s26:>12,.0f} 同比{yoy*100:>7.1f}% 坪效{(s26/a if a else 0):>7,.0f}"
              + (f"  (已撤店{closed})" if closed else ""))


if __name__ == "__main__":
    main()
