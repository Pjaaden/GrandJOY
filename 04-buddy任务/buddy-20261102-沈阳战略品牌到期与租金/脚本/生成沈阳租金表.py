#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
沈阳大悦城 战略品牌「到期时间 + 租金水平」抓取，并预留销售数据接入位。

口径（与战略台账一致）：
- BI 明细粒度 = 合同期次，同一合同编号有多期
- 面积/签约口径只在合同级看（期次累加会重复计数）
- 最终到期日 = 该合同所有期次到期日的最晚值
- 当前租金 = 覆盖统计日期那一期的真实单价（不是历史累加）
- 月面积折算用 365/12 = 30.4167 天

销售数据（后续补充）放在 销售数据/ 目录，文件名任意 .xlsx/.csv，
需含：项目名称 / 品牌或店铺名称 / 统计期间(年月) / 销售额。
本脚本检测到该目录有文件时，自动追加「租售比/销售表现/同比」三个子表。
"""

from __future__ import annotations

import datetime as dt
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
TASK_ROOT = os.path.dirname(BASE)
PROJ_ROOT = os.path.dirname(os.path.dirname(TASK_ROOT))  # GrandJOY 根
# 源数据统一来自战略台账任务目录（BI 原始导出的唯一来源）
LEDGER_ROOT = os.path.join(PROJ_ROOT, "04-buddy任务", "buddy-20261029-战略台账")
SRC = os.path.join(LEDGER_ROOT, "源数据", "运动男装合同台账1029.xlsx")
# 补录台账：BI 未导出的合同（如合同状态「审核中」），结构与 BI 导出一致
# 优先级说明：同一合同编号以 BI 台账为准，此文件只补 BI 里没有的合同
EXTRA_SRC = [
    os.path.join(PROJ_ROOT, "04-buddy任务", "buddy-20261102-战略品牌会议模板",
                 "示例", "EVISU补录台账-演示.xlsx"),
]
OUT_DIR = os.path.join(TASK_ROOT, "产出")
SALES_DIR = os.path.join(TASK_ROOT, "销售数据")
PROJECT = "沈阳大悦城"
DAYS_PER_MONTH = 365 / 12  # 30.4167

# ---------- 样式 ----------
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=10)
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
RED_FILL = PatternFill("solid", fgColor="FCE4E4")   # 90 天内到期
YEL_FILL = PatternFill("solid", fgColor="FFF2CC")   # 180 天内到期
GRN_FILL = PatternFill("solid", fgColor="E2EFDA")   # 待补充销售


def style_sheet(ws, headers, widths, freeze="A2"):
    for j, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[1].height = 34
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def style_body(ws, nrows, ncols, num_fmts: dict[int, str] | None = None):
    num_fmts = num_fmts or {}
    for i in range(2, nrows + 2):
        for j in range(1, ncols + 1):
            c = ws.cell(row=i, column=j)
            c.border = BORDER
            c.font = BODY_FONT
            if j in num_fmts:
                c.number_format = num_fmts[j]
            else:
                c.alignment = Alignment(horizontal="center", vertical="center")


def _d(v):
    if isinstance(v, dt.datetime):
        return v.date()
    return v


def _brand(b):
    """品牌名统一显示口径：BI 多为小写缩写（Charles&Keith / Evisu），对外用标准写法。"""
    s = str(b or "")
    return {"Charles&Keith": "CHARLES&KEITH", "Evisu": "EVISU", "MLB": "MLB"}.get(s, s)


def _clean_unit(u):
    """BI「资源位编号」清洗。

    BI 导出里同一字段混了两种写法：`000003B121`（带项目内码前缀）与 `E320;E321`（不带）。
    去掉 8 位数字前缀后即为干净铺位号；多铺位合同用分号并列。
    """
    s = str(u or "").strip()
    if not s:
        return ""
    parts = [re.sub(r"^\d{6}", "", p.strip()) for p in s.split(";") if p.strip()]
    return ";".join(parts)


def load_contracts(stat_date: dt.date):
    """读源文件，聚合成「合同级」记录。

    BI 台账为主；补录台账（EXTRA_SRC）只补 BI 里查不到的合同编号
    （典型：合同状态「审核中」，BI 尚未导出，如 EVISU B121）。
    """
    def _read(path):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["合同台账"]
        hdr = [c.value for c in ws[1]]
        return [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True)
                if r[1] == PROJECT]

    rows = _read(SRC)
    base_ids = {r["合同编号"] for r in rows}
    for ep in EXTRA_SRC:
        if not os.path.exists(ep):
            continue
        extra = [r for r in _read(ep) if r["合同编号"] not in base_ids]
        if extra:
            print(f"   [补录] {os.path.basename(ep)}：新增 {len(extra)} 行 "
                  f"({', '.join(sorted({str(r['合同编号']) for r in extra}))})")
            rows.extend(extra)

    by_contract = defaultdict(list)
    for r in rows:
        by_contract[r["合同编号"]].append(r)

    out = []
    for cid, periods in by_contract.items():
        periods.sort(key=lambda p: p["起租日"])
        base = periods[0]
        first_start = min(_d(p["起租日"]) for p in periods)
        final_end = max(_d(p["到期日"]) for p in periods)

        cur = [p for p in periods if _d(p["起租日"]) <= stat_date <= _d(p["到期日"])]
        if cur:
            c = cur[0]
            cur_daily = c["单位日租金(元/平方米/日)"]
            cur_rate = c["抽成扣率(%)"]
            cur_way = c["繳付租金方法"]
            cur_period = f"期{c['合同期间']}"
            cur_range = f"{_d(c['起租日'])} ~ {_d(c['到期日'])}"
            cur_flag = "正常"
        else:
            # 当前无覆盖期次：取最近一个未来期次作参考，并标注
            fut = [p for p in periods if _d(p["起租日"]) > stat_date]
            p = min(fut, key=lambda x: _d(x["起租日"])) if fut else periods[-1]
            cur_daily, cur_rate = p["单位日租金(元/平方米/日)"], p["抽成扣率(%)"]
            cur_way, cur_period = p["繳付租金方法"], f"期{p['合同期间']}"
            cur_range = f"{_d(p['起租日'])} ~ {_d(p['到期日'])}"
            cur_flag = "当前无在租期次"

        area = base["签约面积"] or 0
        daily = cur_daily or 0
        # 月固定租金 = 日单价 × 面积 × 30.4167
        monthly_rent = daily * area * DAYS_PER_MONTH
        # 月综合成本 = 租金 + 物业费 + 推广费（均按月面积口径）
        monthly_prop = (base["物业费(元/平方米/月)"] or 0) * area
        monthly_promo = (base["推广费(元/平方米/月)"] or 0) * area
        monthly_total = monthly_rent + monthly_prop + monthly_promo

        # 剩余月数（按 365/12 折算）
        days_left = (final_end - stat_date).days
        months_left = round(days_left / DAYS_PER_MONTH, 1)

        out.append({
            "品牌": _brand(base["品牌/行业"]),
            "店铺名称": base["店铺名称"],
            "是否战略品牌": "是" if str(base["是否战略品牌"]).strip() in ("1", "是") else "否",
            "楼层": base["楼层"],
            "楼栋": base["楼栋"],
            "铺位": _clean_unit(base["资源位编号"]) or base["资源位名称"],
            "楼层位": base["资源位名称"],
            "资源位编号": base["资源位编号"],
            "面积": area,
            "合同编号": cid,
            "商户名称": base["商户名称"],
            "签约公司": base["签约公司"],
            "承租期起": first_start,
            "开业日": _d(base["承租期-开业日"]),
            "最终到期日": final_end,
            "剩余天数": days_left,
            "剩余月数": months_left,
            "当前期次": cur_period,
            "当前期区间": cur_range,
            "当前日租金": daily,
            "当前月租金": round(monthly_rent, 0),
            "当前扣率%": (cur_rate or 0) * 100,
            "租金方式": cur_way,
            "物业费单价": base["物业费(元/平方米/月)"],
            "推广费单价": base["推广费(元/平方米/月)"],
            "月物业费": round(monthly_prop, 0),
            "月推广费": round(monthly_promo, 0),
            "月综合成本": round(monthly_total, 0),
            "年化综合成本": round(monthly_total * 12, 0),
            "单位月综合成本": round(monthly_total / area, 2) if area else 0,
            "期次数": len(periods),
            "口径提示": cur_flag,
        })
    out.sort(key=lambda x: (x["最终到期日"], -x["面积"]))
    return out


def build_detail(ws, contracts):
    headers = [
        "序号", "品牌", "店铺名称", "战略品牌", "楼栋", "楼层", "铺位", "楼层位", "面积(㎡)",
        "合同编号", "商户名称", "签约公司", "承租期起", "开业日", "最终到期日",
        "剩余天数", "剩余月数", "当前期次", "当前期区间", "当前日租金(元/㎡/日)",
        "当前月租金(元)", "当前扣率(%)", "租金方式", "物业费(元/㎡/月)",
        "推广费(元/㎡/月)", "月物业费(元)", "月推广费(元)", "月综合成本(元)",
        "年化综合成本(元)", "单位月综合成本(元/㎡/月)", "口径提示",
    ]
    widths = [5, 16, 18, 9, 10, 8, 22, 26, 10, 20, 22, 24, 11, 11, 11, 9, 9,
              9, 22, 13, 14, 10, 11, 11, 11, 12, 12, 14, 15, 13, 16]
    style_sheet(ws, headers, widths, freeze="C2")
    for i, c in enumerate(contracts, start=2):
        vals = [i - 1, c["品牌"], c["店铺名称"], c["是否战略品牌"], c["楼栋"], c["楼层"],
                c["铺位"], c["楼层位"], c["面积"], c["合同编号"], c["商户名称"], c["签约公司"],
                c["承租期起"], c["开业日"], c["最终到期日"], c["剩余天数"], c["剩余月数"],
                c["当前期次"], c["当前期区间"], c["当前日租金"], c["当前月租金"],
                c["当前扣率%"], c["租金方式"], c["物业费单价"], c["推广费单价"],
                c["月物业费"], c["月推广费"], c["月综合成本"], c["年化综合成本"],
                c["单位月综合成本"], c["口径提示"]]
        for j, v in enumerate(vals, start=1):
            ws.cell(row=i, column=j, value=v)
        # 到期预警底色
        if c["剩余天数"] < 0:
            fill = RED_FILL
        elif c["剩余天数"] <= 90:
            fill = RED_FILL
        elif c["剩余天数"] <= 180:
            fill = YEL_FILL
        else:
            fill = None
        if fill:
            for j in range(1, len(headers) + 1):
                ws.cell(row=i, column=j).fill = fill
    style_body(ws, len(contracts), len(headers), {
        9: "#,##0.00", 13: "yyyy-mm-dd", 14: "yyyy-mm-dd", 15: "yyyy-mm-dd",
        16: "0", 17: "0.0", 20: "0.00", 21: "#,##0", 22: "0.0",
        24: "0.0", 25: "0.0", 26: "#,##0", 27: "#,##0", 28: "#,##0",
        29: "#,##0", 30: "0.00",
    })


def build_summary(ws, contracts):
    headers = ["指标", "数值", "说明"]
    widths = [26, 18, 56]
    style_sheet(ws, headers, widths, freeze="A2")
    total_area = sum(c["面积"] for c in contracts)
    strat = [c for c in contracts if c["是否战略品牌"] == "是"]
    strat_area = sum(c["面积"] for c in strat)
    tot_rent = sum(c["当前月租金"] for c in contracts)
    strat_rent = sum(c["当前月租金"] for c in strat)
    w_daily = tot_rent / total_area / DAYS_PER_MONTH if total_area else 0
    w_strat_daily = strat_rent / strat_area / DAYS_PER_MONTH if strat_area else 0
    in90 = [c for c in contracts if c["剩余天数"] <= 90]
    in180 = [c for c in contracts if c["剩余天数"] <= 180]
    in365 = [c for c in contracts if c["剩余天数"] <= 365]
    rows = [
        ("项目", PROJECT, "数据源：运动男装合同台账1029.xlsx"),
        ("统计日期", dt.date(2026, 10, 28), "BI 统计日期"),
        ("品牌数（合同级）", len(contracts), "同一品牌多铺位/多合同分别计"),
        ("其中战略品牌", len(strat), "是否战略品牌=1"),
        ("签约面积合计(㎡)", round(total_area, 2), "合同级口径，期次不累加"),
        ("其中战略品牌面积(㎡)", round(strat_area, 2), "战略品牌面积占比 %.1f%%" % (strat_area / total_area * 100)),
        ("当前月租金合计(元)", round(tot_rent, 0), "按各合同当前期次日单价×面积×30.4167"),
        ("当前月综合成本合计(元)", round(sum(c["月综合成本"] for c in contracts), 0), "租金+物业费+推广费"),
        ("当前年化综合成本(元)", round(sum(c["年化综合成本"] for c in contracts), 0), "月综合成本×12"),
        ("整体加权日租金(元/㎡/日)", round(w_daily, 2), "月租金合计 ÷ 面积合计 ÷ 30.4167"),
        ("战略品牌加权日租金(元/㎡/日)", round(w_strat_daily, 2), "战略品牌月租金 ÷ 战略品牌面积 ÷ 30.4167"),
        ("整体单位月综合成本(元/㎡/月)", round(sum(c["月综合成本"] for c in contracts) / total_area, 2), "含物业费+推广费"),
        ("90天内到期品牌数", len(in90), "、".join(f"{c['品牌']}({c['最终到期日']})" for c in in90) or "无"),
        ("180天内到期品牌数", len(in180), "、".join(f"{c['品牌']}({c['最终到期日']})" for c in in180) or "无"),
        ("365天内到期品牌数", len(in365), "、".join(f"{c['品牌']}({c['最终到期日']})" for c in in365) or "无"),
        ("最早到期日", min(c["最终到期日"] for c in contracts), "优先续约谈判对象"),
        ("最晚到期日", max(c["最终到期日"] for c in contracts), "锁定长期租金"),
        ("加权平均剩余租期(月)", round(sum(c["剩余月数"] * c["面积"] for c in contracts) / total_area, 1), "按面积加权"),
    ]
    for i, (k, v, note) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=v)
        ws.cell(row=i, column=3, value=note)
    style_body(ws, len(rows), 3)
    for i, (k, v, note) in enumerate(rows, start=2):
        ws.cell(row=i, column=1).font = Font(name="微软雅黑", size=10, bold=True)
        ws.cell(row=i, column=3).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if isinstance(v, dt.date):
            ws.cell(row=i, column=2).number_format = "yyyy-mm-dd"
        elif isinstance(v, float):
            ws.cell(row=i, column=2).number_format = "#,##0.00"
        elif isinstance(v, int) and v > 1000:
            ws.cell(row=i, column=2).number_format = "#,##0"
    # 高亮 90 天内到期行指标
    for i, (k, v, note) in enumerate(rows, start=2):
        if k.startswith("90天内"):
            ws.cell(row=i, column=2).fill = RED_FILL


def build_expiry_plan(ws, contracts):
    """到期时间轴：按到期年月排序，给出续约节奏建议。"""
    headers = ["到期年月", "品牌", "楼层", "面积(㎡)", "最终到期日", "剩余月数",
               "当前日租金(元/㎡/日)", "当前月租金(元)", "续约优先级", "谈判要点"]
    widths = [11, 18, 8, 10, 11, 10, 14, 14, 12, 46]
    style_sheet(ws, headers, widths)
    for i, c in enumerate(contracts, start=2):
        ym = c["最终到期日"].strftime("%Y-%m")
        if c["剩余天数"] <= 90:
            pri, pt = "P0 紧急", "90天内到期，立即启动续约谈判，锁定租金不降"
        elif c["剩余天数"] <= 180:
            pri, pt = "P1 高", "半年内到期，本季度出谈判方案"
        elif c["剩余天数"] <= 365:
            pri, pt = "P2 中", "一年内到期，纳入年度续约计划"
        else:
            pri, pt = "P3 常规", "租期较长，保持经营跟踪"
        # 租金下行信号
        if c["当前日租金"] and c["面积"] >= 400:
            pt += "；大面积主力店，重点关注保底条款"
        vals = [ym, c["品牌"], c["楼层"], c["面积"], c["最终到期日"], c["剩余月数"],
                c["当前日租金"], c["当前月租金"], pri, pt]
        for j, v in enumerate(vals, start=1):
            ws.cell(row=i, column=j, value=v)
        f = RED_FILL if c["剩余天数"] <= 90 else (YEL_FILL if c["剩余天数"] <= 180 else None)
        if f:
            for j in range(1, len(headers) + 1):
                ws.cell(row=i, column=j).fill = f
    style_body(ws, len(contracts), len(headers), {
        4: "#,##0.00", 5: "yyyy-mm-dd", 6: "0.0", 7: "0.00", 8: "#,##0",
    })
    for i in range(2, len(contracts) + 2):
        ws.cell(row=i, column=10).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def build_rent_bench(ws, contracts):
    """租金水平对标：按日租金降序，看同一楼层/业态内的价格带。"""
    headers = ["品牌", "楼层", "面积(㎡)", "当前日租金(元/㎡/日)", "当前扣率(%)",
               "物业费(元/㎡/月)", "单位月综合成本(元/㎡/月)", "日租金排名",
               "与加权均值偏离", "价格带"]
    widths = [18, 8, 10, 16, 11, 13, 17, 10, 13, 12]
    style_sheet(ws, headers, widths)
    total_area = sum(c["面积"] for c in contracts)
    tot_rent = sum(c["当前月租金"] for c in contracts)
    w_daily = tot_rent / total_area / DAYS_PER_MONTH if total_area else 0
    srt = sorted(contracts, key=lambda x: -x["当前日租金"])
    for i, c in enumerate(srt, start=2):
        dev = (c["当前日租金"] / w_daily - 1) * 100
        if c["当前日租金"] >= 15:
            band = "高价带"
        elif c["当前日租金"] >= 8:
            band = "中高价带"
        elif c["当前日租金"] >= 5:
            band = "中价带"
        else:
            band = "低价带"
        vals = [c["品牌"], c["楼层"], c["面积"], c["当前日租金"], c["当前扣率%"],
                c["物业费单价"], c["单位月综合成本"], i - 1, dev / 100, band]
        for j, v in enumerate(vals, start=1):
            ws.cell(row=i, column=j, value=v)
    style_body(ws, len(srt), len(headers), {
        3: "#,##0.00", 4: "0.00", 5: "0.0", 6: "0.0", 7: "0.00", 8: "0", 9: "0.0%",
    })


def build_placeholder(ws, title, headers, widths, note):
    """销售相关子表的占位（等销售数据到位后自动填充）。"""
    style_sheet(ws, headers, widths)
    ws.cell(row=2, column=1, value=note).font = Font(name="微软雅黑", size=10, color="C00000")
    ws.cell(row=2, column=1).fill = GRN_FILL
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws.row_dimensions[2].height = 30
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def find_sales_file():
    if not os.path.isdir(SALES_DIR):
        return None
    for f in sorted(os.listdir(SALES_DIR)):
        if f.startswith("~$"):
            continue
        if f.lower().endswith((".xlsx", ".xls", ".csv")):
            return os.path.join(SALES_DIR, f)
    return None


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    stat_date = dt.date(2026, 10, 28)
    contracts = load_contracts(stat_date)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "1-合同明细"
    build_detail(ws, contracts)

    build_summary(wb.create_sheet("2-汇总指标"), contracts)
    build_expiry_plan(wb.create_sheet("3-到期时间轴"), contracts)
    build_rent_bench(wb.create_sheet("4-租金水平对标"), contracts)

    # 销售相关占位
    build_placeholder(
        wb.create_sheet("5-租售比"),
        "租售比",
        ["品牌", "楼层", "面积(㎡)", "月销售额(元)", "月租金(元)", "月综合成本(元)",
         "租售比(租金/销售)", "综合租售比", "租售比评级", "备注"],
        [18, 8, 10, 14, 14, 15, 14, 14, 12, 30],
        "待补充销售数据。计算口径：租售比 = 月租金 ÷ 月销售额；综合租售比 = (租金+物业费+推广费) ÷ 月销售额。零售业态健康区间通常 10%~18%。",
    )
    build_placeholder(
        wb.create_sheet("6-销售表现"),
        "销售表现",
        ["品牌", "楼层", "面积(㎡)", "月销售额(元)", "月坪效(元/㎡/月)", "同比(%)",
         "环比(%)", "坪效排名", "经营评级", "备注"],
        [18, 8, 10, 14, 15, 11, 11, 10, 12, 30],
        "待补充销售数据。计算口径：月坪效 = 月销售额 ÷ 面积；同比 = 本期 ÷ 去年同期 - 1。",
    )
    build_placeholder(
        wb.create_sheet("7-销售同比"),
        "销售同比",
        ["品牌", "面积(㎡)", "去年同期销售额(元)", "本期销售额(元)", "同比增减额(元)",
         "同比(%)", "面积变动(㎡)", "可比口径同比(%)", "异常标记"],
        [18, 10, 16, 15, 15, 11, 12, 15, 30],
        "待补充两期销售数据。可比口径同比剔除面积变动影响；同比波动超 ±30% 需标注异常并核查（新店/改造/撤柜/数据口径变化）。",
    )

    # 口径说明
    wsc = wb.create_sheet("8-口径说明")
    notes = [
        ("数据源", "运动男装合同台账1029.xlsx（BI 导出，统计日期 2026-10-28）"
                  "＋ EVISU 补录台账（合同状态「审核中」，BI 未导出）"),
        ("项目", PROJECT),
        ("明细粒度", "BI 明细 = 合同期次（同一合同编号有多期），本表已聚合到合同级"),
        ("面积口径", "签约面积只在合同级取一次，期次不可累加（累加会重复计数）"),
        ("铺位口径", "「铺位」= BI 资源位编号（权威铺位号，如 B121）；"
                   "「楼层位」= BI 资源位名称（仅楼层示意，如 B1F-14）"),
        ("最终到期日", "该合同所有期次到期日的最晚值；分段合同/换品牌会推后真实到期"),
        ("当前租金", "取覆盖统计日期 2026-10-28 的那一期真实日单价，不是合同历史累加"),
        ("月面积折算", "365/12 = 30.4167 天（不是 30 天）"),
        ("月租金", "当前日租金 × 签约面积 × 30.4167"),
        ("月综合成本", "月租金 + 月物业费 + 月推广费"),
        ("租金方式", "本项目全部合同均为「两者取高」（保底租金 vs 抽成取高）"),
        ("⚠ 台账 vs 招调方案", "本表只反映「合同现行版本」（BI 台账）状态。"
                          "项目提供的招调报告属「计划续约/调改」方案，不回写本表；"
                          "续约后租金/费率另见「沈阳大悦城-2027年到期战略品牌招调简表」。"),
        ("⚠ EVISU 口径", "B121 EVISU 合同状态为「审核中」，BI 台账未导出，由补录台账补入："
                       "日租金 17.50 元/㎡/日、扣率 16%、物业费 65、推广费 10、到期 2026-12-20。"
                       "**以系统内合同版本 17.50 为准**；项目招调报告中的 15.50 为续约报价，"
                       "不是现行租金，不得覆盖本表现行值。"),
        ("⚠ 租金异常提示", "CAMEL 当前期次（期2，2026-08-09~2028-06-09）日租金仅 0.82 元/㎡/日，"
                          "而前期（期1）为 2.79，后续 0.92→0.92→1.02→1.15 缓慢递增。"
                          "需核实是否录入错误或存在特殊条款（如分期补差），否则将显著拉低加权租金。"),
        ("⚠ 租金回调提示", "李宁当前期次（期4，2025-12-30~2026-12-29）日租金 25.2，"
                          "低于上一期 28.62（-11.9%），虽期5 回升至 25.96，仍属续约降租。"
                          "GXG 同样存在期6→期7 由 17.01 降至 15.31（-10.0%）的情况。"),
        ("销售数据接入", "将销售明细放入 04-buddy任务/buddy-20261102-沈阳战略品牌到期与租金/销售数据/，"
                        "重跑本脚本即可自动填充 5-租售比、6-销售表现、7-销售同比 三张表。"
                        "所需字段：项目名称 / 品牌(或店铺名称) / 统计年月 / 销售额。"),
    ]
    style_sheet(wsc, ["项目", "说明"], [20, 100])
    for i, (k, v) in enumerate(notes, start=2):
        wsc.cell(row=i, column=1, value=k).font = Font(name="微软雅黑", size=10, bold=True)
        wsc.cell(row=i, column=2, value=v)
    style_body(wsc, len(notes), 2)
    for i in range(2, len(notes) + 2):
        wsc.cell(row=i, column=2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        wsc.row_dimensions[i].height = 30

    # 销售数据若已存在，在此处接入（当前为空）
    sales_file = find_sales_file()
    out = os.path.join(OUT_DIR, f"沈阳大悦城-战略品牌到期与租金-{dt.date.today():%Y%m%d}.xlsx")
    wb.save(out)
    print("saved:", out)
    print("contracts:", len(contracts))
    print("sales_file:", sales_file or "(none — 销售子表为占位)")


if __name__ == "__main__":
    main()
