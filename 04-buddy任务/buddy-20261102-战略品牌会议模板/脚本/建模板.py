# -*- coding: utf-8 -*-
"""
生成「战略品牌到期与招调」空白模板（可直接向已有文件追加 sheet，也可新建）。

用法：
  # 1) 新建一个空模板文件
  python 建模板.py --out /path/新模板.xlsx

  # 2) 向已有的数据文件追加两个 sheet（主表已存在时报错，不覆盖）
  python 建模板.py --append /path/已有数据.xlsx

版式来源：2026-11-02 沈阳项目人工调整终版（用户手动加列 + 加预览页 + 删注脚后的形态）
"""
import argparse
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 版式常量（与人工终版一致）
FONT_CN = "微软雅黑"
C_TITLE = "1F4E79"      # 深蓝标题 / 表头底
C_SUB = "808080"        # 灰色副标题
C_RED = "C00000"        # 涨（中式：涨红跌绿）
C_GREEN = "008000"
C_TOTAL = "D9E2F3"      # 合计行浅蓝

GRADE_FILL = {
    "极高危": ("C00000", "FFFFFF"),
    "过高":   ("FFC7CE", None),
    "偏高":   ("FFEB9C", None),
    "健康":   ("C6EFCE", None),
    "偏低":   ("DDEBF7", None),
}
GRADE_EMPTY = "EDEDED"

# 主表 29 列：(表头, 列宽, 数字格式, 汇总方式)
#   汇总方式：sum=直接相加｜wavg=按面积加权｜ratio=分子分母各自相加再相除｜yoy=另算｜note=不汇总
COLUMNS = [
    ("铺位号", 10, "General", "note"),
    ("品牌", 14, "General", "note"),
    ("楼层", 7, "General", "note"),
    ("面积(㎡)", 10, "0.00", "sum"),
    ("战略品牌", 9, "General", "note"),
    ("原合同到期", 12, "General", "note"),
    ("计划签约", 10, "General", "note"),
    ("计划开业", 9, "General", "note"),
    ("现行日租\n(元/㎡/日)", 11, "0.00", "wavg"),
    ("现行扣率(%)", 10, "0.0", "note"),
    ("实际收租\n(元/月)", 12, "#,##0", "sum"),
    ("现行物业费\n(元/㎡/月)", 11, "0.0", "note"),
    ("现行月综合成本\n(元)", 14, "#,##0", "sum"),
    ("计划日租\n(元/㎡/日)", 11, "0.00", "wavg"),
    ("计划扣率(%)", 10, "0.0", "note"),
    ("租金同比(%)", 10, "+0.0%;-0.0%;0.0%", "note"),
    ("计划年限(年)", 10, "0", "note"),
    ("计划月租金\n(元)", 13, "#,##0", "sum"),
    ("计划月综合成本\n(元)", 14, "#,##0", "sum"),
    ("25年月均销售\n(元)", 14, "#,##0", "sum"),
    ("26年月均销售\n(元)", 14, "#,##0", "sum"),
    ("月均销售同比(%)", 12, "+0.0%;-0.0%;0.0%", "note"),
    ("26年最新月销售\n(元)", 15, "#,##0", "note"),
    ("现行租售比(%)", 11, "0.0%", "ratio"),
    ("计划后租售比(%)", 12, "0.0%", "ratio"),
    ("租售比变动(pp)", 11, "+0.0;-0.0;0.0", "note"),
    ("租售比评级", 16, "General", "note"),
    ("抽成/固租", 10, "0.0%", "note"),
    ("手动备注", 34, "General", "note"),   # ← 人工终版新增列，纯手工填写
]
NCOL = len(COLUMNS)

# 预览页 7 列（B–H，A 列留白作左边距）
PREVIEW = [
    ("铺位", 28.0, "General"),
    ("品牌", 28.7, "General"),
    ("日租 现→计", 17.6, "General"),
    ("26年月均", 15.1, "#,##0"),
    ("同比", 15.5, "0.00%"),
    ("计划租售比", 16.7, "0.00%"),
    ("评级", 13.0, "General"),
]

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor=C_TITLE)
HDR_FONT = Font(name=FONT_CN, size=9, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_CN, size=9)
CENTER = Alignment(horizontal="center", vertical="center")


def build_main(ws, project="", title_suffix="战略品牌到期与招调简表", subtitle=""):
    """在主 sheet 上绘制 29 列表头（数据区留空，供脚本或手工填入）。"""
    ws.title = ws.title if ws.title not in ("Sheet", "Sheet1") else "招调简表"

    MERGE_TO = NCOL - 1        # AC「手动备注」是人工手填列，不并入标题合并区
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=MERGE_TO)
    c = ws.cell(row=1, column=1, value=f"{project} · {title_suffix}".strip(" ·"))
    c.font = Font(name=FONT_CN, size=14, bold=True, color=C_TITLE)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=MERGE_TO)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = Font(name=FONT_CN, size=8.5, italic=True, color=C_SUB)
    c.alignment = Alignment(horizontal="left", vertical="center")

    HR = 4
    for j, (name, w, fmt, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=HR, column=j, value=name)
        if j == NCOL:      # 「手动备注」列：人工终版为纯手填，不套表头样式
            cell.font = Font(name="宋体", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.font = HDR_FONT
            cell.fill = HDR_FILL
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[HR].height = 42

    ws.freeze_panes = "C5"
    ws.sheet_view.zoomScale = 107
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def build_preview(ws):
    """预览页：7 列极简视图，供会议现场看。A 列留白做左边距。

    版式严格照搬人工终版：
      · 无标题行，表头直接在第 2 行（第 1 行留白）
      · 表头主题色 3 + tint 0.8、微软雅黑 10、无加粗
      · 数据行 / 合计行纯白填充 FFFFFFFF
      · 右边框 medium #EBEBEB，左边框 thin；H 列右边框 thin
    """
    ws.title = "简洁版预览"
    HR = 2
    HDR_F = PatternFill("solid", fgColor="D9E2F3")      # 主题色 3 / tint 0.8 的等效色
    BODY_F = PatternFill("solid", fgColor="FFFFFFFF")
    R_MED = Side(style="medium", color="EBEBEB")
    R_THIN = Side(style="thin")
    T_MED = Side(style="medium", color="EBEBEB")

    for j, (name, w, _fmt) in enumerate(PREVIEW, start=2):     # 从 B 列开始
        cell = ws.cell(row=HR, column=j, value=name)
        cell.font = Font(name=FONT_CN, size=10)
        cell.fill = HDR_F
        cell.alignment = CENTER
        cell.border = Border(left=R_THIN, right=R_MED, top=R_THIN, bottom=None)
        ws.column_dimensions[get_column_letter(j)].width = w
    # 表头最右列（H）右边框改细
    ws.cell(row=HR, column=1 + len(PREVIEW)).border = Border(
        left=R_THIN, right=R_THIN, top=R_THIN, bottom=None)

    ws.column_dimensions["A"].width = 13.0      # 人工终版 A 列为 13
    return ws


# 预览页数据写入（版式与人工终版逐格对齐）
def _pv_num(v, nd=2):
    return "—" if v is None else f"{v:,.{nd}f}"


def _pv_pct(v, dash="—"):
    """负数写成「−18.5%」文本（人工终版的习惯），非负返回原值走百分比格式。"""
    if v is None:
        return dash
    return f"−{abs(v)*100:.1f}%" if v < 0 else v


def fill_preview(ws, rows):
    """把 rows（主表数据 dict 列表）写进预览页，含合计行。"""
    R_MED = Side(style="medium", color="EBEBEB")
    R_THIN = Side(style="thin")
    T_MED = Side(style="medium", color="EBEBEB")
    BODY_F = PatternFill("solid", fgColor="FFFFFFFF")
    HR = 2
    LAST = 1 + len(PREVIEW)          # H 列

    TOT_F = PatternFill("solid", fgColor="D9E2F3")     # 合计行：主题色 3 / tint 0.8

    def put(r, c, v, fmt="General", bold=False, total=False):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name=FONT_CN, size=10, bold=bold)
        cell.fill = TOT_F if total else BODY_F
        cell.alignment = CENTER
        cell.number_format = fmt
        right = R_THIN if c == LAST else R_MED
        cell.border = Border(left=R_THIN, right=right, top=T_MED, bottom=None)
        return cell

    r = HR + 1
    for d in rows:
        put(r, 2, d.get("铺位号"))
        put(r, 3, str(d.get("品牌") or "").replace(" → ", "→"))
        put(r, 4, f"{_pv_num(d.get('现行日租'))} → {_pv_num(d.get('计划日租'))}")
        put(r, 5, round(d.get("26年月均销售") or 0), "#,##0")
        yv = round(d["月均销售同比"], 3) if d.get("月均销售同比") is not None else None
        put(r, 6, _pv_pct(yv), "General" if isinstance(_pv_pct(yv), str) else "0.00%")
        put(r, 7, round(d["计划后租售比"], 3) if d.get("计划后租售比") is not None else None, "0.00%")
        # 预览页评级去掉「（须降租）」这类行动提示后缀，保留「（关注）/（预警）」
        put(r, 8, str(d.get("租售比评级") or "").replace("（须降租）", ""))
        ws.row_dimensions[r].height = 17.55
        r += 1

    # 合计行
    tot_area = sum((d.get("面积(㎡)") or 0) for d in rows) or 1
    dd = [d for d in rows if d.get("26年月均销售")]
    sale = sum(d.get("26年月均销售") or 0 for d in dd)
    r_s = sum(d.get("实际收租") or 0 for d in dd)
    r_p = sum(max(d.get("计划月租金") or 0,
                  (d.get("26年月均销售") or 0) * (d.get("计划扣率") or 0) / 100) for d in dd)
    a25 = sum(d.get("25年月均销售") or 0 for d in dd)
    a26 = sale
    cur_d = sum((d.get("现行日租") or 0) * (d.get("面积(㎡)") or 0) for d in rows) / tot_area
    plan_d = sum((d.get("计划日租") or 0) * (d.get("面积(㎡)") or 0) for d in rows) / tot_area

    put(r, 2, "合计", total=True)
    put(r, 3, None, total=True)
    put(r, 4, f"{cur_d:.2f} → {plan_d:.2f}", total=True)
    put(r, 5, round(a26), "#,##0", total=True)
    tv = round(a26 / a25 - 1, 3) if a25 else None
    put(r, 6, _pv_pct(tv), "General" if isinstance(_pv_pct(tv), str) else "0.00%", total=True)
    put(r, 7, f"{round(r_s/sale*100,1):.1f}% → {round(r_p/sale*100,1):.1f}%" if sale else "—", total=True)
    put(r, 8, None, total=True)
    r += 1
    return r


def fill_row(ws, r, d):
    """向主表第 r 行写入一条铺位记录并套用版式。d 为 dict，键与 COLUMNS 顺序对应。"""
    vals = [
        d.get("铺位号"), d.get("品牌"), d.get("楼层"), d.get("面积(㎡)"), d.get("战略品牌"),
        d.get("原合同到期"), d.get("计划签约"), d.get("计划开业"),
        d.get("现行日租"), d.get("现行扣率"), d.get("实际收租"), d.get("现行物业费"), d.get("现行月综合成本"),
        d.get("计划日租"), d.get("计划扣率"), d.get("租金同比"), d.get("计划年限"),
        d.get("计划月租金"), d.get("计划月综合成本"),
        d.get("25年月均销售"), d.get("26年月均销售"), d.get("月均销售同比"), d.get("26年最新月销售"),
        d.get("现行租售比"), d.get("计划后租售比"), d.get("租售比变动"),
        d.get("租售比评级"), d.get("抽成/固租"), d.get("手动备注"),
    ]
    for j, (v, (_, _, fmt, _)) in enumerate(zip(vals, COLUMNS), start=1):
        cell = ws.cell(row=r, column=j, value=v)
        if j == NCOL:      # 手动备注列留给人工，不套自动版式
            cell.font = Font(name=FONT_CN, size=9)
            cell.alignment = CENTER
            continue
        cell.font = BODY_FONT
        cell.border = BORDER
        cell.alignment = CENTER
        cell.number_format = fmt

    # 评级底色
    grade = str(d.get("租售比评级") or "")
    gc = ws.cell(row=r, column=27)
    for key, (bg, fg) in GRADE_FILL.items():
        if key in grade:
            gc.fill = PatternFill("solid", fgColor=bg)
            if fg:
                gc.font = Font(name=FONT_CN, size=9, bold=True, color=fg)
            break
    else:
        gc.fill = PatternFill("solid", fgColor=GRADE_EMPTY)

    # 涨红跌绿（中式惯例）：租金同比 / 销售同比 / 租售比变动
    for j in (16, 22, 26):
        c = ws.cell(row=r, column=j)
        if isinstance(c.value, (int, float)):
            c.font = Font(name=FONT_CN, size=9,
                          color=C_GREEN if c.value < 0 else C_RED)
    return r


def fill_total(ws, r, rows):
    """合计/加权行。rows 为主表数据 dict 列表。"""
    tot_area = sum(d.get("面积(㎡)") or 0 for d in rows) or 1
    ws.cell(row=r, column=1, value="合计/加权")
    ws.cell(row=r, column=4, value=sum(d.get("面积(㎡)") or 0 for d in rows))
    ws.cell(row=r, column=9,
            value=sum((d.get("现行日租") or 0) * (d.get("面积(㎡)") or 0) for d in rows) / tot_area)
    ws.cell(row=r, column=11, value=sum(d.get("实际收租") or 0 for d in rows))
    ws.cell(row=r, column=13, value=sum(d.get("现行月综合成本") or 0 for d in rows))
    ws.cell(row=r, column=14,
            value=sum((d.get("计划日租") or 0) * (d.get("面积(㎡)") or 0) for d in rows) / tot_area)
    ws.cell(row=r, column=18, value=sum(d.get("计划月租金") or 0 for d in rows))
    ws.cell(row=r, column=19, value=sum(d.get("计划月综合成本") or 0 for d in rows))
    ws.cell(row=r, column=21, value=sum(d.get("26年月均销售") or 0 for d in rows))

    dd = [d for d in rows if d.get("26年月均销售")]
    if dd:
        rent_s = sum(d.get("实际收租") or 0 for d in dd)
        rent_p = sum(max(d.get("计划月租金") or 0,
                         (d.get("26年月均销售") or 0) * (d.get("计划扣率") or 0) / 100) for d in dd)
        sale = sum(d.get("26年月均销售") or 0 for d in dd)
        ws.cell(row=r, column=24, value=rent_s / sale)
        ws.cell(row=r, column=25, value=rent_p / sale)
        ws.cell(row=r, column=26, value=(rent_p - rent_s) / sale * 100)
        a25 = sum(d.get("25年月均销售") or 0 for d in dd)
        if a25:
            a26 = sum(d.get("26年月均销售") or 0 for d in dd)
            ws.cell(row=r, column=22, value=a26 / a25 - 1)

    for j in range(1, NCOL + 1):
        cell = ws.cell(row=r, column=j)
        cell.font = Font(name=FONT_CN, size=9, bold=True)
        cell.border = BORDER
        cell.alignment = CENTER
        cell.fill = PatternFill("solid", fgColor=C_TOTAL)
        cell.number_format = COLUMNS[j - 1][2]
    # 合计行的销售同比按人工终版用 0.0%（明细行才是带正负号的 +0.0%;-0.0%）
    ws.cell(row=r, column=22).number_format = "0.0%"
    ws.row_dimensions[r].height = 20
    return r


def build_new(path, project=""):
    wb = openpyxl.Workbook()
    build_main(wb.active, project=project,
               subtitle="数据口径：合同台账 ＋ 店铺销售月报　｜　月面积折算 30.4167 天　｜　金额单位：元")
    build_preview(wb.create_sheet())
    wb.save(path)
    print(f"已生成模板：{path}")


def append_to(path, project=""):
    wb = openpyxl.load_workbook(path)
    if "简洁版预览" in wb.sheetnames:
        print("已存在「简洁版预览」，跳过追加，避免覆盖。", file=sys.stderr)
        return
    build_preview(wb.create_sheet())
    wb.save(path)
    print(f"已追加预览页：{path}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", help="新建模板文件路径")
    ap.add_argument("--append", help="向已有文件追加预览页")
    ap.add_argument("--project", default="", help="项目名，如「沈阳大悦城」")
    a = ap.parse_args()
    if a.out:
        build_new(a.out, a.project)
    elif a.append:
        append_to(a.append, a.project)
    else:
        ap.print_help()
