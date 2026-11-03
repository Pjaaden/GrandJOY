#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""沈阳 2027 年到期战略品牌 —— 14 列精简版（战略级次全 A）

数据来源：同目录 `会前同步.py` 的主表产出（29 列），本脚本只做「取列 + 单位换算 + 排版」，
不重算任何口径，避免两版数字打架。若主表更新，重跑本脚本即可同步。

列口径说明：
- 固定租金(元/㎡/天) = 主表「现行日租」= 合同当前版本单价（不是招调报告的「现租金」）
- 抽成比例          = 主表「现行扣率」/ 100
- 25年月均销售(万元) = 主表「25年月均销售」/ 10000
- 2026销售(万元)     = 主表「26年月均销售」/ 10000（月均口径，非全年累计）
- 同比              = 主表「月均销售同比」（两年同月份交集口径）
- 租售比            = 主表「现行租售比」= max(固租, 销售×扣率) ÷ 26年月均销售
- 27年招调计划      = 「续约 x 年」，换牌铺位追加「调整为 新品牌」
- 品牌              = 只写原品牌；换牌铺位的新品牌只出现在招调计划列
"""

from __future__ import annotations

import os
import re
import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
TASK = os.path.dirname(BASE)
SRC = os.path.join(TASK, "产出", "沈阳大悦城-2027年到期战略品牌招调简表.xlsx")
OUT = os.path.join(TASK, "产出", "沈阳大悦城-2027年到期战略品牌招调简表-14列版.xlsx")

TITLE = "沈阳大悦城 · 2027 年到期战略品牌清单"

HDRS = [
    "序号", "品牌", "战略级次", "铺位号", "合同面积\n(㎡)", "起租日", "到期日",
    "固定租金\n(元/㎡/天)", "抽成比例", "25年月均销售\n(万元)",
    "26年月均销售\n(万元)", "同比", "租售比", "27年招调计划",
]
WIDTHS = [5, 14, 8, 12, 10, 11, 11, 12, 9, 13, 13, 9, 9, 46]

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
BODY = Font(name="微软雅黑", size=10)
TOT_FILL = PatternFill("solid", fgColor="D9E2F3")
RED_F = PatternFill("solid", fgColor="FCE4E4")
YEL_F = PatternFill("solid", fgColor="FFF2CC")


def read_main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["招调简表"]
    hdr = [str(c.value).replace("\n", "") for c in ws[4]]
    ix = {h: i + 1 for i, h in enumerate(hdr)}
    rows = []
    for r in range(5, ws.max_row + 1):
        u = ws.cell(row=r, column=ix["铺位号"]).value
        if not u or str(u).startswith("合计"):
            continue
        g = lambda k: ws.cell(row=r, column=ix[k]).value
        # 主表品牌可能是「波司登 → RE'VAN+ariose years」这种换牌写法：
        # 本表品牌列只留原品牌，新品牌归到「27年招调计划」列。
        raw = str(g("品牌") or "")
        brand, newb = raw, None
        for sep in ("→", "->", "—>"):
            if sep in raw:
                brand, newb = (s.strip() for s in raw.split(sep, 1))
                break
        rows.append({
            "品牌": brand.strip(),
            "计划新品牌": (newb or "").strip() or None,
            "铺位号": u,
            "面积": g("面积(㎡)"),
            "到期": g("原合同到期"),
            "日租": g("现行日租(元/㎡/日)"),
            "扣率": g("现行扣率(%)"),
            "25均": g("25年月均销售(元)"),
            "26均": g("26年月均销售(元)"),
            "同比": g("月均销售同比(%)"),
            "租售比": g("现行租售比(%)"),
            "计划日租": g("计划日租(元/㎡/日)"),
            "计划扣率": g("计划扣率(%)"),
            "计划年限": g("计划年限(年)"),
        })
    return rows


def start_date(exp):
    """起租日：合同现行期次里，起始不晚于到期日的那一段。

    本表取的是「当前在租合同」的承租期起租日，与到期日同属一个合同周期；
    数据源主表未单列起租日，这里按「到期日往前推整租期」不可靠，
    故直接从台账明细（若存在）补齐——见 build()。
    """
    return None


def pick_plan_text(d):
    """27年招调计划：一句话讲清「换不换牌 + 租金怎么变 + 费率 + 续几年」。

    格式：续约降至 15.50 元（-11.4%），扣率 16%，续约 1 年
    换牌铺位把「调整为 新品牌」提到最前面，让人第一眼看到要换品牌。
    """
    pr = d["计划年限"] or 1
    pr = int(pr) if float(pr) == int(pr) else pr

    parts = []
    newb = d.get("计划新品牌")
    if newb:
        parts.append(f"调整为 {newb}")

    cur, plan = d["日租"], d["计划日租"]
    if plan is None:
        parts.append("续约租金待定")
    else:
        if cur:
            up = plan / cur - 1
            if abs(up) < 5e-4:          # |涨幅| < 0.05%，视为维持
                parts.append(f"续约维持 {plan:.2f} 元")
            else:
                arrow = "降" if up < 0 else "涨"
                parts.append(f"续约{arrow}至 {plan:.2f} 元（{up * 100:+.1f}%）")
        else:
            parts.append(f"续约至 {plan:.2f} 元")

    if d["计划扣率"] is not None:
        parts.append(f"扣率 {d['计划扣率']:.0f}%")
    parts.append(f"续约 {pr} 年")
    return "，".join(parts)


def build():
    rows = read_main()
    # 起租日 / 最终到期日 从 8-sheet 表补（那里有完整合同周期）
    # 铺位号两边写法不同：主表 `E320-321` / 8sheet `E320;E321`，需归一化后再比
    rent_wb = os.path.join(TASK, "..", "buddy-20261102-沈阳战略品牌到期与租金",
                           "产出", "沈阳大悦城-战略品牌到期与租金-20261102.xlsx")

    def parts(u):
        """拆成铺位号集合。

        `-` 在 `E320-321` 里是**区间连接符**（同楼层连续铺位），要展开成 E320、E321；
        在 `A1F-01` 这类里又是楼栋分隔，故仅在同前缀+纯数字结尾时才当区间处理。
        """
        s = str(u or "").strip().replace(" ", "")
        for a in (";", "；", "/", "、"):
            s = s.replace(a, "|")
        out = set()
        for p in s.split("|"):
            if not p:
                continue
            m = re.fullmatch(r"([A-Za-z]+)(\d+)-(\d+)", p)
            if m and len(m.group(3)) == len(m.group(2)):
                pre, a, b = m.group(1), int(m.group(2)), int(m.group(3))
                if b - a < 20:      # 合理区间，展开
                    w = len(m.group(2))
                    out |= {f"{pre}{i:0{w}d}" for i in range(a, b + 1)}
                    continue
            out.add(p)
        return out

    recs = []          # [(parts_set, 起租日, 到期日)]
    if os.path.exists(rent_wb):
        w2 = openpyxl.load_workbook(rent_wb, data_only=True)["1-合同明细"]
        h2 = [str(c.value).replace("\n", "") for c in w2[1]]
        j2 = {h: i + 1 for i, h in enumerate(h2)}
        for r in range(2, w2.max_row + 1):
            ps = parts(w2.cell(row=r, column=j2["铺位"]).value)
            recs.append((ps,
                         w2.cell(row=r, column=j2["承租期起"]).value,
                         w2.cell(row=r, column=j2["最终到期日"]).value))

    def lookup(unit):
        """铺位号两侧写法不同，按优先级匹配：
        1) 分段集合完全相同                    （B114 = B114）
        2) 主表分段是全量集合的子集的唯一一条    （A101/A102 ⊂ A101;A102;A102A）
        3) 主表分段在候选里命中率最高且唯一      （E320-321 vs E320;E321，靠品牌兜底）
        匹配不唯一则返回 None，宁可留空也不猜。
        """
        ps = parts(unit)
        exact = [r for r in recs if r[0] == ps]
        if len(exact) == 1:
            return exact[0][1], exact[0][2]
        sub = [r for r in recs if ps <= r[0]]
        if len(sub) == 1:
            return sub[0][1], sub[0][2]
        cand = [r for r in recs if ps & r[0]]
        if len(cand) == 1:
            return cand[0][1], cand[0][2]
        return None

    # 起租日 = **该品牌在本铺位的最初合同起始日**，不是当前这一期合同的起租日。
    # 来源：8-sheet 表的「承租期起」，那是合同级字段，同一合同多期次值相同，
    # 即该合同链起租日；已逐条与 BI 台账首期「起租日」核对一致
    # （B305 2017-11-19 / C119 2019-09-06 / E114 2022-12-30 / A101 2024-05-20 /
    #  B114 2025-06-01 / E320-321 2025-12-30）。
    #
    # 例外：EVISU B121 台账无更早合同链（BI 主台账零命中，补录台账仅 2026-01-01 一期），
    # 2026-01-01 是年度续约合同写法、不是首进日。首进日按项目提供口径硬编码 2023-06-01。
    START_OVERRIDE = {"B121": datetime.date(2023, 6, 1)}

    def _dt(v):
        """到期日可能是 datetime 也可能是 'YYYY-MM-DD' 字符串，统一成 date，
        否则该列会因字符串单元格绕过 number_format 而显示成文本。"""
        if isinstance(v, datetime.datetime):
            return v.date()
        if isinstance(v, datetime.date):
            return v
        if isinstance(v, str):
            for f in ("%Y-%m-%d", "%Y/%m/%d"):
                try:
                    return datetime.datetime.strptime(v.strip(), f).date()
                except ValueError:
                    pass
        return v

    for d in rows:
        d["到期"] = _dt(d["到期"])
        rec = lookup(d["铺位号"])
        if rec:
            d["起租日"], d["到期日"] = _dt(rec[0]), _dt(rec[1])
        else:
            d["起租日"], d["到期日"] = None, d["到期"]
            print(f"  [警告] {d['铺位号']} 未唯一匹配到合同周期")
        ov = START_OVERRIDE.get(d["铺位号"])
        if ov:
            d["起租日"] = ov
            print(f"  [覆盖] {d['铺位号']} 起租日按最初进场口径 {ov}（台账无更早合同链）")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "战略品牌清单"

    n = len(HDRS)
    # 标题（表后不加任何说明行）
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    c = ws.cell(row=1, column=1, value=TITLE)
    c.font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 表头
    HR = 2
    for j, (h, w) in enumerate(zip(HDRS, WIDTHS), start=1):
        c = ws.cell(row=HR, column=j, value=h)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[HR].height = 34

    # 数据
    for i, d in enumerate(rows, start=1):
        r = i + HR
        vals = [
            i, d["品牌"], "A", d["铺位号"], d["面积"], d["起租日"], d["到期日"],
            d["日租"], d["扣率"] / 100.0,
            round((d["25均"] or 0) / 10000, 1),
            round((d["26均"] or 0) / 10000, 1),
            d["同比"], d["租售比"], pick_plan_text(d),
        ]
        for j, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = BORDER
            cell.font = BODY
            if j in (14,):
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
        fmts = {5: "#,##0.00", 6: "yyyy-mm-dd", 7: "yyyy-mm-dd", 8: "0.00",
                9: "0.0%", 10: "#,##0.0", 11: "#,##0.0", 12: "+0.0%;-0.0%;0.0%",
                13: "0.0%"}
        for j, f in fmts.items():
            ws.cell(row=r, column=j).number_format = f
        # 租售比预警底色
        ratio = d["租售比"] or 0
        if ratio >= 0.40:
            ws.cell(row=r, column=13).fill = RED_F
        elif ratio >= 0.30:
            ws.cell(row=r, column=13).fill = YEL_F
        ws.row_dimensions[r].height = 22

    # 合计行
    tr = len(rows) + HR + 1
    area = sum(d["面积"] for d in rows)
    s25 = sum((d["25均"] or 0) for d in rows)
    s26 = sum((d["26均"] or 0) for d in rows)
    wday = sum(d["日租"] * d["面积"] for d in rows) / area
    # 加权租售比：按 26 年销售加权实际收租
    actual = sum((d["26均"] or 0) * (d["租售比"] or 0) for d in rows)
    wratio = actual / s26
    ws.cell(row=tr, column=1, value="合计 / 加权")
    ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=2)
    ws.cell(row=tr, column=5, value=round(area, 2)).number_format = "#,##0.00"
    ws.cell(row=tr, column=8, value=round(wday, 2)).number_format = "0.00"
    ws.cell(row=tr, column=10, value=round(s25 / 10000, 1)).number_format = "#,##0.0"
    ws.cell(row=tr, column=11, value=round(s26 / 10000, 1)).number_format = "#,##0.0"
    ws.cell(row=tr, column=12, value=(s26 / s25 - 1)).number_format = "+0.0%;-0.0%;0.0%"
    ws.cell(row=tr, column=13, value=round(wratio, 4)).number_format = "0.0%"
    for j in range(1, n + 1):
        c = ws.cell(row=tr, column=j)
        c.fill = TOT_FILL
        c.font = Font(name="微软雅黑", size=10, bold=True)
        c.border = BORDER
        if c.alignment is None or j not in (1,):
            c.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=tr, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[tr].height = 22

    ws.freeze_panes = f"C{HR + 1}"
    ws.auto_filter.ref = (f"A{HR}:{get_column_letter(n)}{len(rows) + HR}")

    wb.save(OUT)
    print("saved:", OUT)
    print("rows:", len(rows))
    for d in rows:
        print(f"  {d['铺位号']:<12} {str(d['品牌'])[:22]:<24} "
              f"起 {d['起租日']} 到 {d['到期日']} 租 {d['日租']}")


if __name__ == "__main__":
    build()
