#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
拉取安踏体系 + 特步体系的租约台账，为后续接经营月报（销售数据）打底。

范围：
  · 安踏 ANTA（含 ANTA SUPERSTORE、ANTA GUANJUN 两种店型）
  · ANTA KIDS（安踏儿童）—— 当前台账中**无此品牌**，脚本会如实标注缺口
  · 特步、特步儿童

输出（Excel）：
  1) 铺位租约台账  —— 粒度=铺位，含当前在租期次的真实租金 / 扣率 / 到期日 / 续约优先级
  2) 合同期次明细  —— 粒度=合同期次，完整签约历史（含已过期期次）
  3) 品牌汇总      —— 按体系/品牌/店型汇总
  4) 经营数据待补  —— 预置销售数据列骨架（金额、坪效、客流等），等月报填入

用法：
    python 拉取安踏特步租约.py [--源文件 <BI导出xlsx>] [--输出 <输出xlsx路径>]
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 台账核心 import (  # noqa: E402
    build_contract, build_detail, load_sources, split_source_args, _date,
)

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

# ── 品牌范围定义 ────────────────────────────────────────────────────────────
# 安踏体系：台账里是 4 个独立品牌名
#   （SUPERSTORE / GUANJUN 是独立店型，各自有独立铺位；安踏儿童 2026-11-05 补充导出后并入）
ANTA_BRANDS = {"ANTA", "ANTA SUPERSTORE", "ANTA  GUANJUN", "安踏儿童"}
XTEP_BRANDS = {"特步", "特步儿童"}
# 注：2026-11-05 起「安踏儿童」已由 BI 单独导出并入总表（源数据/安踏儿童合同台账1104.xlsx），
#     此前台账缺该品牌 —— 不再登记为缺失。
MISSING_BRANDS: dict[str, str] = {}

# 品牌 → 体系 / 店型
def brand_meta(b: str, shop_name: str = "") -> tuple[str, str]:
    """体系与店型判定。

    ⚠ 口径提醒：BI 里安踏的店型标注**不统一** ——
      · 品牌级标了 GUANJUN（杭州大悦城 L1-02）、SUPERSTORE（厦门大悦城 L117、L118）
      · 店铺名称级标了「ANTA 冠军」（北京朝阳 4F-50）、「ANTA 殿堂」（北京朝阳 4F-84），
        但这两家的**品牌字段仍是 ANTA**
    所以店型判定顺序：品牌级 GUANJUN/SUPERSTORE → 店铺名含「冠军」→ 店铺名含「殿堂」→ 标准店。
    两种口径并列保留，便于核对 BI 标注是否需要修正。
    """
    b = (b or "").strip()
    sn = (shop_name or "").strip()
    bu = b.upper()
    if bu == "ANTA SUPERSTORE":
        return "安踏体系", "SUPERSTORE 超级店"
    if "GUANJUN" in bu:
        return "安踏体系", "GUANJUN 冠军店"
    if "SUPERSTORE" in bu:
        return "安踏体系", "SUPERSTORE 超级店"
    if b == "安踏儿童":
        # 店铺名称有「安踏儿童」「ANTA kids」「antakids」「3F-33 ANTA kids」多种写法
        return "安踏体系", "儿童店"
    if b == "特步":
        return "特步体系", "标准店"
    if b == "特步儿童":
        return "特步体系", "儿童店"
    # ANTA 主品牌：用店铺名称细分店型
    if b.upper() == "ANTA":
        if "冠军" in sn or "GUANJUN" in sn.upper():
            return "安踏体系", "冠军店（店铺名标注）"
        if "殿堂" in sn:
            return "安踏体系", "殿堂店（店铺名标注）"
        return "安踏体系", "标准店"
    return "其他", "—"


def in_scope(b: str) -> bool:
    b = (b or "").strip()
    if b in ANTA_BRANDS or b in XTEP_BRANDS:
        return True
    # 兜底：安踏品牌名带空格变体（如 "ANTA  GUANJUN"）
    return b.upper().startswith("ANTA") or b.upper().startswith("特步")


# ── 样式 ────────────────────────────────────────────────────────────────────
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
GRP_FILL = PatternFill("solid", fgColor="E8EEF7")
WARN_FILL = PatternFill("solid", fgColor="FFE3E3")
OK_FILL = PatternFill("solid", fgColor="E6F4EA")
INFO_FILL = PatternFill("solid", fgColor="FFF3E0")


def style_sheet(ws, headers, widths=None, freeze="A2"):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def finish(ws, headers, rows, num_cols=()):
    for i, r in enumerate(rows, 2):
        for j, h in enumerate(headers, 1):
            v = r.get(h)
            c = ws.cell(row=i, column=j, value=v)
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=False)
            if h in num_cols and isinstance(v, (int, float)):
                c.number_format = "#,##0.00" if abs(v) < 1000 else "#,##0"
            elif isinstance(v, (dt.date, dt.datetime)):
                c.number_format = "yyyy-mm-dd"
    ws.sheet_view.showGridLines = False


def main():
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ap = argparse.ArgumentParser()
    ap.add_argument("--源文件", default=os.path.join(
        here, "源数据", "运动男装合同台账1103.xlsx") +
        "," + os.path.join(here, "源数据", "安踏儿童合同台账1104.xlsx"),
        action="append",
        help="BI 导出 xlsx；可多次传参或用逗号分隔（多源合并）")
    ap.add_argument("--输出", default=None)
    ap.add_argument("--截至", default=None, help="判断在租的基准日，默认今天")
    args = ap.parse_args()

    today = dt.date.fromisoformat(args.截至) if args.截至 else dt.date.today()
    srcs = split_source_args(args.源文件)
    src, meta = load_sources(srcs)
    detail_all = build_detail(src)
    contract_all = build_contract(detail_all, src)

    # ── 过滤范围 ────────────────────────────────────────────────────────────
    detail = [r for r in detail_all if in_scope(r.get("品牌/行业"))]
    contracts = [r for r in contract_all if in_scope(r.get("品牌/行业"))]
    found_brands = {r["品牌/行业"].strip() for r in detail}
    print(f"[范围] 明细 {len(detail)} 行 / 合同 {len(contracts)} 份 / 品牌 {sorted(found_brands)}")

    # ⚠ 台账的明细层/合同层**没有保留「店铺名称」字段**，但安踏的店型判定要用它
    #   （「ANTA 冠军」「ANTA 殿堂」只在店铺名称里，品牌字段仍是 ANTA）。
    # 这里直接从 BI 源表按 合同编号 + 期次 建映射补回来。
    shop_by_cid = {}
    for r in src:
        cid = str(r.get("合同编号", "")).strip()
        if cid and cid not in shop_by_cid:
            shop_by_cid[cid] = str(r.get("店铺名称") or "").strip()

    def shop_of(contract_no: str) -> str:
        return shop_by_cid.get(str(contract_no or "").strip(), "")

    # ── 1. 铺位级租约台账 ────────────────────────────────────────────────────
    by_cid = defaultdict(list)
    for r in detail:
        by_cid[r["合同编号"]].append(r)

    live_phase = {}
    future_phase = {}
    for r in detail:
        s, e = _date(r.get("起租日")), _date(r.get("到期日"))
        if not (s and e):
            continue
        if s <= today <= e:
            live_phase[r["合同编号"]] = r
        elif s > today:
            # 已签约但尚未起租（即将开业）—— 经营上有意义，不能当成空铺
            if r["合同编号"] not in future_phase or s < _date(future_phase[r["合同编号"]]["起租日"]):
                future_phase[r["合同编号"]] = r

    slots = defaultdict(list)
    for x in contracts:
        slots[(x["项目名称"], x["资源位编号"])].append(x)

    slot_rows = []
    for (proj, slot_no), cons in slots.items():
        ends = [_date(x["合同到期日"]) for x in cons if x.get("合同到期日")]
        if not ends:
            continue
        final = max(ends)

        live_c = None
        for x in cons:
            if x["合同编号"] in live_phase:
                live_c = x
                break
        if live_c is None:
            live_c = sorted(cons, key=lambda y: y.get("合同到期日") or "")[-1]
        phase = live_phase.get(live_c["合同编号"])

        # ⚠ 三态区分（不能都算「无在租期次」）：
        #   在租    = 有期次满足 起租日 ≤ 今天 ≤ 到期日
        #   未起租  = 已签约，起租日在今天之后（即将开业的新店）
        #   空置    = 全部期次都已过期
        if phase is not None:
            status = "在租"
        else:
            fut = [future_phase[x["合同编号"]] for x in cons
                   if x["合同编号"] in future_phase]
            if fut:
                phase = min(fut, key=lambda y: y["起租日"])
                status = f"未起租（{_date(phase['起租日']).strftime('%Y-%m-%d')} 开业）"
            else:
                status = "空置·合同已过期"

        brand = live_c["品牌/行业"]
        shop_name = shop_of(live_c["合同编号"])
        system, shop_type = brand_meta(brand, shop_name)
        days = (final - today).days
        if days < 0:
            pri = "已过期·待核查"
        elif days <= 90:
            pri = "P0 立即启动"
        elif days <= 180:
            pri = "P1 半年内启动"
        elif days <= 365:
            pri = "P2 一年内启动"
        else:
            pri = "P3 一年内以上"

        rent = phase.get("单位日租金(元/平方米/日)") if phase else None
        rate = phase.get("抽成扣率") if phase else None
        pf = phase.get("物业费(元/平方米/月)") if phase else None
        promo = phase.get("推广费(元/平方米/月)") if phase else None
        pay = phase.get("缴付租金方法") if phase else None

        # ⚠ 成本口径必须区分租金制式，否则横向比较会严重误导：
        #   固租 / 两者取高且有日租 → 月成本 = 固租 + 物业 + 推广
        #   纯抽成（日租=0）        → 商场**没有**固定租金收入，实际按销售额扣率分成
        #   两者取高但日租=0        → 同样无保底（合同下限为 0），固定租金收入可能为 0
        #   三者「月成本」含义完全不同，必须靠「租金制式」列区分
        fixed_rent_m = round(rent * 30, 2) if rent else 0.0
        monthly = round(fixed_rent_m + (pf or 0) + (promo or 0), 2)

        if pay == "抽成":
            rent_mode = "纯抽成（无固定租金）"
        elif pay == "两者取高" and not rent:
            rent_mode = "两者取高但无保底"
        elif pay == "两者取高" and rent:
            rent_mode = "两者取高（有保底）"
        elif pay == "固定租金":
            rent_mode = "固定租金"
        else:
            rent_mode = pay or "—"

        area = live_c.get("签约面积(平方米)")
        # 月成本总额 = 月综合成本 × 面积（纯抽成时此值偏小，须配合租金制式解读）
        cost_total = round(monthly * area, 2) if (monthly is not None and area) else None

        seg = " | ".join(
            f"{x['品牌/行业']}({x['合同起租日']}~{x['合同到期日']})"
            for x in sorted(cons, key=lambda y: y.get("合同起租日") or "")
        )

        notes = []
        if status.startswith("未起租"):
            notes.append("已签约未起租，即将开业")
        elif status.startswith("空置"):
            notes.append("全部期次均已过期，需确认铺位现状")
        if pay == "抽成":
            notes.append("纯抽成制，无固定租金，实收需用销售额×扣率测算")
        elif pay == "两者取高" and not rent:
            notes.append("两者取高但固定租金下限为0（无保底），续约应争取设保底")
        if len(cons) > 1:
            notes.append(f"含{len(cons)}份分段合同")
        if len({x["品牌/行业"] for x in cons}) > 1:
            notes.append("铺位已换品牌")
        if any(str(x.get("合同状态", "")) == "审核中" for x in cons):
            notes.append("存在审核中的合同期次，需核实")

        slot_rows.append({
            "体系": system,
            "品牌": brand,
            "店型": shop_type,
            "店铺名称(源表)": shop_name,
            "项目名称": proj,
            "楼层": live_c.get("楼层", ""),
            "资源位编号": slot_no,
            "商户名称": live_c.get("商户名称", ""),
            "签约面积(㎡)": area,
            "所属区域": live_c.get("所属区域", ""),
            "三级业态": live_c.get("三级业态", ""),
            "合同份数": len(cons),
            "分段合同历史": seg,
            "当前在租期次": phase.get("合同期间") if phase else None,
            "当前期次起租日": phase.get("起租日") if phase else None,
            "当前期次到期日": phase.get("到期日") if phase else None,
            "最终到期日": final.isoformat(),
            "距最终到期天数": days,
            "续约优先级": pri,
            "在租状态": status,
            "租金制式": rent_mode,
            "日租金(元/㎡/日)": rent,
            "抽成扣率": rate,
            "物业费(元/㎡/月)": pf,
            "推广费(元/㎡/月)": promo,
            "月固租(元/㎡/月)": fixed_rent_m,
            "月综合成本(元/㎡/月)": monthly,
            "月成本总额(元)": cost_total,
            "日租金计算方式": phase.get("日租金计算方式") if phase else None,
            "缴付租金方法": pay,
            "免租期天数(天)": phase.get("免租期天数(天)") if phase else None,
            "备注": "；".join(notes),
        })

    order = {"P0 立即启动": 0, "P1 半年内启动": 1, "P2 一年内启动": 2, "P3 一年内以上": 3}
    slot_rows.sort(key=lambda x: (x["体系"], x["品牌"], order.get(x["续约优先级"], 9), x["距最终到期天数"]))

    # ── 2. 合同期次明细 ─────────────────────────────────────────────────────
    ph_rows = []
    for r in sorted(detail, key=lambda x: (x["品牌/行业"], x["项目名称"], x["资源位编号"], x.get("合同期间") or 0)):
        s, e = _date(r.get("起租日")), _date(r.get("到期日"))
        state = "在租" if (s and e and s <= today <= e) else ("未起租" if (s and s > today) else "已过期")
        system, shop_type = brand_meta(r["品牌/行业"], shop_of(r["合同编号"]))
        ph_rows.append({
            "体系": system,
            "品牌": r["品牌/行业"],
            "店型": shop_type,
            "项目名称": r["项目名称"],
            "楼层": r.get("楼层", ""),
            "资源位编号": r["资源位编号"],
            "合同编号": r["合同编号"],
            "合同名称": r.get("合同名称", ""),
            "合同期间": r.get("合同期间"),
            "起租日": r.get("起租日"),
            "到期日": r.get("到期日"),
            "签约面积(㎡)": r.get("签约面积(平方米)"),
            "日租金(元/㎡/日)": r.get("单位日租金(元/平方米/日)"),
            "抽成扣率": r.get("抽成扣率"),
            "物业费(元/㎡/月)": r.get("物业费(元/平方米/月)"),
            "推广费(元/㎡/月)": r.get("推广费(元/平方米/月)"),
            "期间固定租金总额(元)": r.get("期间固定租金总额(元)"),
            "期间物业费总额(元)": r.get("期间物业费总额(元)"),
            "期间推广费总额(元)": r.get("期间推广费总额(元)"),
            "缴付租金方法": r.get("缴付租金方法"),
            "日租金计算方式": r.get("日租金计算方式"),
            "免租期天数(天)": r.get("免租期天数(天)"),
            "期间天数(天)": r.get("期间天数(天)"),
            "在租状态": state,
            "合同状态": r.get("合同状态", ""),
            "商户名称": r.get("商户名称", ""),
        })

    # ── 3. 品牌汇总 ─────────────────────────────────────────────────────────
    agg = defaultdict(lambda: {"铺位": 0, "面积": 0.0, "月成本": 0.0, "合同": set(),
                               "在租": 0, "空置": 0, "未起租": 0, "到期min": None, "到期max": None,
                               "纯抽成": 0, "无保底": 0, "有保底": 0, "固租": 0})
    for s in slot_rows:
        k = (s["体系"], s["品牌"], s["店型"])
        a = agg[k]
        a["铺位"] += 1
        a["面积"] += s["签约面积(㎡)"] or 0
        a["月成本"] += s["月成本总额(元)"] or 0
        a["合同"].add(s["项目名称"] + "|" + s["资源位编号"])
        st = s["在租状态"]
        if st == "在租":
            a["在租"] += 1
        elif st.startswith("未起租"):
            a["未起租"] = a.get("未起租", 0) + 1
        else:
            a["空置"] += 1
        rm = s["租金制式"]
        if "纯抽成" in rm:
            a["纯抽成"] += 1
        elif "无保底" in rm:
            a["无保底"] += 1
        elif "有保底" in rm:
            a["有保底"] += 1
        else:
            a["固租"] += 1
        d = s["最终到期日"]
        a["到期min"] = d if a["到期min"] is None or d < a["到期min"] else a["到期min"]
        a["到期max"] = d if a["到期max"] is None or d > a["到期max"] else a["到期max"]

    sum_rows = []
    for (system, brand, shop_type), a in sorted(agg.items()):
        area = round(a["面积"], 2)
        sum_rows.append({
            "体系": system,
            "品牌": brand,
            "店型": shop_type,
            "项目数": len({x["项目名称"] for x in slot_rows
                        if x["品牌"] == brand and x["体系"] == system}),
            "铺位数": a["铺位"],
            "在租铺位": a["在租"],
            "未起租铺位": a.get("未起租", 0),
            "无在租期次": a["空置"],
            "签约面积(㎡)": area,
            "月成本总额(元)": round(a["月成本"], 2),
            "其中纯抽成铺位": a["纯抽成"],
            "其中无保底铺位": a["无保底"],
            "最早最终到期日": a["到期min"],
            "最晚最终到期日": a["到期max"],
            "状态": "正常" if a["空置"] == 0 else f"{a['空置']} 个铺位合同已过期",
        })

    # 体系合计
    sys_agg = defaultdict(lambda: {"铺位": 0, "面积": 0.0, "月成本": 0.0,
                                   "在租": 0, "空置": 0, "未起租": 0})
    for s in slot_rows:
        a = sys_agg[s["体系"]]
        a["铺位"] += 1
        a["面积"] += s["签约面积(㎡)"] or 0
        a["月成本"] += s["月成本总额(元)"] or 0
        st = s["在租状态"]
        if st == "在租":
            a["在租"] += 1
        elif st.startswith("未起租"):
            a["未起租"] += 1
        else:
            a["空置"] += 1
    total_rows = []
    for system, a in sys_agg.items():
        sub = [x for x in slot_rows if x["体系"] == system]
        total_rows.append({
            "体系": system, "品牌": "【小计】", "店型": "—",
            "项目数": len({x["项目名称"] for x in sub}),
            "铺位数": a["铺位"], "在租铺位": a["在租"],
            "未起租铺位": a["未起租"], "无在租期次": a["空置"],
            "签约面积(㎡)": round(a["面积"], 2),
            "月成本总额(元)": round(a["月成本"], 2),
            "其中纯抽成铺位": sum(1 for x in sub if "纯抽成" in x["租金制式"]),
            "其中无保底铺位": sum(1 for x in sub if "无保底" in x["租金制式"]),
            "最早最终到期日": None, "最晚最终到期日": None,
            "状态": "正常" if a["空置"] == 0 else f"{a['空置']} 个铺位合同已过期",
        })

    # ── 4. 经营数据待补（销售数据骨架）────────────────────────────────────
    biz_rows = []
    for s in slot_rows:
        biz_rows.append({
            "体系": s["体系"], "品牌": s["品牌"], "店型": s["店型"],
            "项目名称": s["项目名称"], "楼层": s["楼层"], "资源位编号": s["资源位编号"],
            "签约面积(㎡)": s["签约面积(㎡)"],
            "在租状态": s["在租状态"],
            "起始月": None, "结束月": None,
            "月销售额(元)": None, "月客流(人次)": None, "月坪效(元/㎡)": None,
            "月租售比(%)": None, "累计销售(元)": None,
            "同比(%)": None, "环比(%)": None,
            "销售来源": "经营月报", "备注": None,
        })

    # ── 写 Excel ────────────────────────────────────────────────────────────
    out_path = args.输出 or os.path.join(
        here, "产出", f"安踏特步体系租约台账_截至{today:%Y%m%d}.xlsx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb = openpyxl.Workbook()

    # 说明页
    ws0 = wb.active
    ws0.title = "说明"
    info = [
        ["安踏体系 + 特步体系 · 租约台账（台账口径）", ""],
        ["", ""],
        ["数据源文件", meta.get("file", "")],
        ["BI 统计日期", meta.get("stat_date", "")],
        ["在租判定基准日", today.isoformat()],
        ["台账明细总行数", meta.get("rows", "")],
        ["本次范围明细行数", len(detail)],
        ["本次范围合同份数", len(contracts)],
        ["本次范围铺位数", len(slot_rows)],
        ["", ""],
        ["【口径说明】", ""],
        ["粒度", "铺位级 = 项目名称 + 资源位编号（同一铺位可能有多份分段合同/换品牌）"],
        ["最终到期日", "该铺位所有合同到期日的最晚值（这才是需要总部介入续约的时点）"],
        ["当前租金", "当前正在履约期次的真实单价，不是合同级历史加总"],
        ["在租判定", "起租日 ≤ 基准日 ≤ 到期日"],
        ["铺位状态三态", "在租 / 未起租（已签约，尚未开业）/ 空置·合同已过期 —— 三者不可混为一谈"],
        ["月综合成本", "日租金×30 + 物业费 + 推广费"],
        ["月成本总额", "月综合成本 × 签约面积"],
        ["", ""],
        ["【⚠ 租金制式必须区分】", ""],
        ["两者取高（有保底）", "有固定租金下限，月成本可直接横向比较"],
        ["两者取高但无保底", "日租金记 0 = 合同下限为 0，商场无保底租金；"
                          "月成本只含物业+推广，续约应争取设定保底"],
        ["纯抽成（无固定租金）", "商场没有固定租金收入，实际按销售额×扣率分成；"
                            "月成本只含物业+推广，**不能**与固租铺位直接比大小"],
        ["", ""],
        ["【重要提示】", ""],
    ]
    for b, note in MISSING_BRANDS.items():
        info.append([f"⚠ {b}（{note}）", "当前 BI 导出中无此品牌，台账未收录。"
                                        "需确认：是否在「运动+男装」筛选范围外，或尚未签署合同。"])
    info += [
        ["⚠ 店型标注不统一", "安踏的 GUANJUN / SUPERSTORE 在品牌字段标注，"
                          "而「ANTA 冠军」「ANTA 殿堂」只出现在店铺名称字段（品牌字段仍是 ANTA）。"
                          "本表「店型」列已按 品牌级→店铺名 顺序统一判定，并保留「店铺名称(源表)」列备查。"],
        ["【后续】", "「经营数据待补」页已预置销售数据列骨架，待经营月报到位后填入"],
        ["", ""],
        ["子表", "铺位租约台账 / 合同期次明细 / 品牌汇总 / 经营数据待补"],
    ]
    for i, (a, b) in enumerate(info, 1):
        ws0.cell(row=i, column=1, value=a).font = Font(bold=(i in (1,) or a.startswith("【")), size=11)
        ws0.cell(row=i, column=2, value=b).alignment = Alignment(wrap_text=True, vertical="top")
    ws0.column_dimensions["A"].width = 26
    ws0.column_dimensions["B"].width = 86
    ws0.sheet_view.showGridLines = False

    # 铺位租约台账
    H1 = list(slot_rows[0].keys())
    ws1 = wb.create_sheet("铺位租约台账")
    style_sheet(ws1, H1, [10, 16, 22, 18, 16, 6, 20, 30, 11, 14, 12, 9, 46, 11, 13, 13, 13, 11, 11,
                           20, 13, 11, 11, 11, 13, 13, 14, 14, 16, 40])
    finish(ws1, H1, slot_rows, num_cols=set(H1) - {"最终到期日", "当前期次起租日", "当前期次到期日"})
    for i, r in enumerate(slot_rows, 2):
        if r["续约优先级"].startswith(("P0", "P1")):
            for j in range(1, len(H1) + 1):
                ws1.cell(row=i, column=j).fill = WARN_FILL
        elif r["在租状态"] != "在租" or r["续约优先级"].startswith("已过期"):
            for j in range(1, len(H1) + 1):
                ws1.cell(row=i, column=j).fill = INFO_FILL

    # 合同期次明细
    H2 = list(ph_rows[0].keys())
    ws2 = wb.create_sheet("合同期次明细")
    style_sheet(ws2, H2, [10, 16, 20, 16, 6, 20, 20, 26, 8, 12, 12, 11, 13, 11, 11, 11,
                           16, 16, 16, 13, 13, 10, 9, 10, 12, 30])
    finish(ws2, H2, ph_rows, num_cols=set(H2) - {"起租日", "到期日"})
    for i, r in enumerate(ph_rows, 2):
        if r["在租状态"] == "在租":
            ws2.cell(row=i, column=26).fill = OK_FILL
        elif r["在租状态"] == "未起租":
            ws2.cell(row=i, column=26).fill = INFO_FILL

    # 品牌汇总
    H3 = list(sum_rows[0].keys())
    ws3 = wb.create_sheet("品牌汇总")
    style_sheet(ws3, H3, [12, 22, 22, 9, 9, 10, 12, 13, 15, 13, 13, 15, 15, 26])
    finish(ws3, H3, sum_rows, num_cols={"签约面积(㎡)", "月成本总额(元)"})
    for i, r in enumerate(sum_rows, 2):
        if r["无在租期次"]:
            ws3.cell(row=i, column=7).fill = INFO_FILL
    r0 = len(sum_rows) + 2
    ws3.cell(row=r0, column=1, value="【体系小计】").font = Font(bold=True)
    for j, h in enumerate(H3, 1):
        c = ws3.cell(row=r0, column=j, value=None)
        c.fill, c.border = GRP_FILL, BORDER
    for k, t in enumerate(total_rows):
        for j, h in enumerate(H3, 1):
            v = t.get(h)
            c = ws3.cell(row=r0 + 1 + k, column=j, value=v)
            c.fill, c.border = GRP_FILL, BORDER
            if h in ("签约面积(㎡)", "月成本总额(元)") and isinstance(v, (int, float)):
                c.number_format = "#,##0.00"

    # 经营数据待补
    H4 = list(biz_rows[0].keys())
    ws4 = wb.create_sheet("经营数据待补")
    style_sheet(ws4, H4, [10, 16, 20, 16, 6, 20, 11, 11, 10, 10, 14, 13, 14, 12, 14, 10, 10, 12, 18])
    finish(ws4, H4, biz_rows, num_cols={"签约面积(㎡)", "月销售额(元)", "月坪效(元/㎡)",
                                    "累计销售(元)", "月租售比(%)", "同比(%)", "环比(%)"})
    ws4.insert_rows(1)
    ws4.cell(row=1, column=1, value="⚠ 销售数据待补：本页销售列已留空，待经营月报到位后填入。"
                                   "每行对应「铺位租约台账」中的一个铺位。")
    ws4.cell(row=1, column=1).font = Font(bold=True, color="B54708")
    ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(H4))
    ws4.cell(row=1, column=1).fill = INFO_FILL
    ws4.freeze_panes = "A3"

    wb.save(out_path)

    # ── 控制台摘要 ──────────────────────────────────────────────────────────
    print(f"\n[输出] {out_path}")
    n_live = sum(1 for x in slot_rows if x["在租状态"] == "在租")
    n_fut = sum(1 for x in slot_rows if x["在租状态"].startswith("未起租"))
    n_vac = sum(1 for x in slot_rows if x["在租状态"].startswith("空置"))
    print(f"[铺位] {len(slot_rows)} 个 | 在租 {n_live} | 未起租 {n_fut} | 空置(合同已过期) {n_vac}")
    print("\n=== 品牌汇总 ===")
    for r in sum_rows:
        print(f"  {r['体系']:8s} {r['品牌']:20s} {r['店型']:22s} "
              f"铺位{r['铺位数']:3d} 在租{r['在租铺位']:3d} 未起租{r.get('未起租铺位',0):2d} "
              f"面积{r['签约面积(㎡)']:9,.2f}㎡ 月成本{r['月成本总额(元)']:14,.0f}元 "
              f"到期 {r['最早最终到期日']}~{r['最晚最终到期日']}")
    print("\n=== 体系小计 ===")
    for r in total_rows:
        print(f"  {r['体系']:8s} 铺位{r['铺位数']:3d} 在租{r['在租铺位']:3d} 未起租{r.get('未起租铺位',0):2d} "
              f"面积{r['签约面积(㎡)']:9,.2f}㎡ 月成本{r['月成本总额(元)']:14,.0f}元")
    missing = [b for b in MISSING_BRANDS if b not in found_brands]
    if missing:
        print(f"\n⚠ 台账中缺失的品牌: {missing} → 已在「说明」页标注缺口")
    # 续约重点
    p01 = [x for x in slot_rows if x["续约优先级"].startswith(("P0", "P1"))]
    if p01:
        print(f"\n=== 一年内需启动续约（P0/P1）{len(p01)} 个铺位 ===")
        for x in sorted(p01, key=lambda y: y["距最终到期天数"]):
            print(f"  {x['距最终到期天数']:5d}天 {x['续约优先级']:8s} {x['品牌']:18s} "
                  f"{x['项目名称']:14s} {x['资源位编号']:16s} 到期 {x['最终到期日']}")


if __name__ == "__main__":
    main()
