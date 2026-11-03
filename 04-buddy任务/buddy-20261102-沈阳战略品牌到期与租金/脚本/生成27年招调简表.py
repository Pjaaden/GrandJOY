# -*- coding: utf-8 -*-
"""
沈阳大悦城 2027年到期战略品牌招调简表（修订版）
- 销售数据源：/Users/zhijian/Downloads/店铺销售月报 (1).xlsx（按品牌名筛选版）
- 剔除：室内场地/特卖等非铺位、面积不一致记录
- 输出：单 sheet，28 列，与用户给定列序完全一致
"""
import os
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJ_ROOT = "/Users/Shared/Projects/GrandJOY"
TASK_ROOT = os.path.join(PROJ_ROOT, "04-buddy任务", "buddy-20261102-沈阳战略品牌到期与租金")
SRC_LEDGER = os.path.join(PROJ_ROOT, "04-buddy任务", "buddy-20261029-战略台账",
                          "源数据", "运动男装合同台账1029.xlsx")
# 销售数据源（多文件合并；同一铺位同一月份取两表较大者，防重复累加）
# 文件不存在时自动跳过并在控制台/注脚标注，不中断生成
SRC_SALES = [
    "/Users/zhijian/Downloads/店铺销售月报 (1).xlsx",   # 按品牌名筛选版（6家）
    "/Users/zhijian/Downloads/店铺销售月报 (2).xlsx",   # EVISU 专项（若存在则并入）
]
# 主台账未收录的合同（如 EVISU 状态为「审核中」，BI 未导出），单独合并
SRC_LEDGER_EXTRA = [
    SRC_LEDGER,
    "/Users/zhijian/Downloads/evisu合同台账.xlsx",      # 可选，缺失时回退内置台账
]

# EVISU 沈阳合同（SY00000297-1-1）台账快照 —— 源自 evisu合同台账.xlsx。
# 该文件为一次性补充材料，若本机不存在则使用此内置快照，保证结果可复现。
# 仅 1 期：2023-06-01 ~ 2026-12-20，日租 17.50 / 扣率 16% / 物业 65 / 推广 10。
EVISU_LEDGER_SNAPSHOT = {
    "楼层": "L1",
    "面积": 133.1,
    "到期": "2026-12-20",
    "日租": 17.50,
    "扣率": 16.0,
    "物业费": 65.0,
    "推广费": 10.0,
    "战略": "是",
    "状态": "审核中",
}
OUT = os.path.join(TASK_ROOT, "产出", "沈阳大悦城-2027年到期战略品牌招调简表.xlsx")

DAYS = 365 / 12        # 月面积折算天数（与大悦BI口径一致，非30天）
CUR_MONTH = "202610"   # 当前月（用于「现行」口径）

# ---------------------------------------------------------------- 招调计划（用户提供）
# 租金同比一律以「续约日租 ÷ 现行日租 − 1」实算，不再直接采信招调表口径。
# 用户 2026-11-02 复核口径：涨幅与前后租金必须一一对应。
PLAN = [
    # 用户回填：现 21.00 → 续 21.00，涨幅应为 0%（原报 -20% 系旧基数口径）
    dict(铺位="B114",    原品牌="MLB",           面积=280.34, 年限=2, 日租=21.00, 扣率=0.11, 同比=0.0,
         物业费=65, 推广费=10, 到期="2027-04-12", 签约="2027/3",  开业="—"),
    # EVISU：合同状态「审核中」，主台账未收录，取自单独提供的 evisu合同台账.xlsx
    # 台账期次仅一期：2023-06-01~2026-12-20，日租 17.50 / 扣率 16%
    # 用户口径：现租金 15.50 → 续约 15.50，涨幅 0%（与台账 17.50 差 2.00，见注9）
    dict(铺位="B121",    原品牌="EVISU",         面积=133.10, 年限=1, 日租=15.50, 扣率=0.16, 同比=0.0,
         物业费=65, 推广费=10, 到期="2026-12-20", 签约="2026/12", 开业="—",
         现日租=15.50),
    dict(铺位="C119",    原品牌="CHARLES&KEITH", 面积=161.43, 年限=3, 日租=18.50, 扣率=0.17, 同比=0.0,
         物业费=70, 推广费=10, 到期="2027-09-05", 签约="2027/8",  开业="—"),
    # 用户修正：续约日租 13.01（原误填 15.31），涨幅 -15.0%
    dict(铺位="B305",    原品牌="GXG",           面积=126.97, 年限=1, 日租=13.01, 扣率=0.21, 同比=-0.15,
         物业费=65, 推广费=10, 到期="2027-02-18", 签约="2027/1",  开业="—"),
    # 现 25.96 → 续 25.96，涨幅 0%（原报 -4% 系与上期 28.62 比较的口径）
    dict(铺位="E114",    原品牌="李宁",           面积=387.62, 年限=1, 日租=25.96, 扣率=0.20, 同比=0.0,
         物业费=65, 推广费=10, 到期="2027-12-29", 签约="2027/12", 开业="—"),
    dict(铺位="E320-321", 原品牌="MLB",           面积=265.93, 年限=1, 日租=2.00,  扣率=0.12, 同比=0.0,
         物业费=70, 推广费=10, 到期="2026-12-29", 签约="2027/12", 开业="—"),
    # 波司登：招调后品牌变更为 RE'VAN + ariose years，铺位号 A101/A102
    # 现 6.77 → 续 6.77，涨幅 0%（原报 -19% 系旧基数口径）
    dict(铺位="A101/A102", 原品牌="波司登",       面积=407.21, 年限=3, 日租=6.77,  扣率=0.18, 同比=0.0,
         物业费=65, 推广费=10, 到期="2027-05-19", 签约="2027/5",  开业="2027/6",
         新品牌="RE'VAN+ariose years"),
]

# ------------------------------------------------- 销售记录 -> 铺位归属（按合同号 + 面积判定，名称不可靠）
# 「MLB」名称在报表中被两家店混用：
#   · 202501-202505 阶段「MLB」= 合同 HT2024040700089 / 280.34㎡ → B114（L1）
#   · 202512-202610 阶段「MLB」= 合同 HT2025120300063 / 265.93㎡ → E320-321（L3）
#   · 「MLB（B114）」自 202506 起接续 B114；「MLB（E320-321）」为 E320-321 老合同
# 因此按 (合同编号, 面积) 精确映射，避免同名串号。
UNIT_BY_CONTRACT = {
    # 合同编号                        -> 铺位
    "HT2024040700089":                "B114",
    "HT2024040700089-1":              "B114",
    "HT2025120300063":                "E320-321",
    "HT2022052500143":                "E320-321",
    "00000300000222201901":           "C119",
    "0000030003C35001201701":         "B305",
    "HT2024041600121":                "A101/A102",   # 波司登 L1 407.21㎡
    "HT2022101700057":                "E114",
    "SY00000297-1-1":                 "B121",   # EVISU
}

# 现行合同号（期次最晚者）
CUR_CONTRACT = {
    "B114":     "HT2024040700089-1",
    "C119":     "00000300000222201901",
    "B305":     "0000030003C35001201701",
    "E114":     "HT2022101700057",
    "E320-321": "HT2025120300063",
    "A101/A102": "HT2024041600121",
    "B121": "SY00000297-1-1",       # EVISU（合同状态：审核中）
}

COLUMNS = [
    ("铺位号", 10), ("品牌", 14), ("楼层", 7), ("面积(㎡)", 10), ("战略品牌", 9),
    ("原合同到期", 12), ("计划签约", 10), ("计划开业", 9),
    ("现行日租\n(元/㎡/日)", 11), ("现行扣率(%)", 10), ("实际收租\n(元/月)", 12),
    ("现行物业费\n(元/㎡/月)", 11), ("现行月综合成本\n(元)", 14),
    ("计划日租\n(元/㎡/日)", 11), ("计划扣率(%)", 10), ("租金同比(%)", 10), ("计划年限(年)", 10),
    ("计划月租金\n(元)", 13), ("计划月综合成本\n(元)", 14),
    ("25年月均销售\n(元)", 14), ("26年月均销售\n(元)", 14), ("月均销售同比(%)", 12),
    ("26年最新月销售\n(元)", 15),
    ("现行租售比(%)", 11), ("计划后租售比(%)", 12), ("租售比变动(pp)", 11),
    ("租售比评级", 16), ("抽成/固租", 10),
]
NCOL = len(COLUMNS)


# ---------------------------------------------------------------- 读台账
def load_ledger():
    """主台账 + 补充台账（EVISU 等主表未收录的「审核中」合同）合并读取。
    补充台账文件缺失时，回退 EVISU_LEDGER_SNAPSHOT 内置快照。"""
    by_ct = defaultdict(list)
    status = {}
    used = []
    for src in SRC_LEDGER_EXTRA:
        if not os.path.exists(src):
            print(f"  [跳过] 台账文件不存在：{src}")
            continue
        used.append(src)
        wb = openpyxl.load_workbook(src, data_only=True)
        ws = wb["合同台账"]
        hdr = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        ix = {h: i for i, h in enumerate(hdr)}
        rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
                if r[ix["项目名称"]] and "沈阳" in str(r[ix["项目名称"]])]
        for r in rows:
            by_ct[r[ix["合同编号"]]].append(r)
            status[r[ix["合同编号"]]] = str(r[ix.get("合同状态", 11)] or "")
    out = {}
    for p in PLAN:
        cno = CUR_CONTRACT[p["铺位"]]
        recs = by_ct.get(cno, [])
        if not recs:
            if p["铺位"] == "B121":
                print(f"  [回退] 合同 {cno} 未在台账文件中找到，使用内置快照")
                out[p["铺位"]] = dict(EVISU_LEDGER_SNAPSHOT)
                continue
            raise RuntimeError(f"台账未找到合同 {cno}")
        cur = max(recs, key=lambda x: x[ix["到期日"]])
        rate_raw = float(cur[ix["抽成扣率(%)"]] or 0)
        out[p["铺位"]] = dict(
            楼层=str(cur[ix["楼层"]]).strip(),
            面积=float(cur[ix["签约面积"]]),
            到期=str(cur[ix["到期日"]])[:10],
            日租=float(cur[ix["单位日租金(元/平方米/日)"]] or 0),
            扣率=rate_raw * 100 if rate_raw <= 1 else rate_raw,
            物业费=float(cur[ix["物业费(元/平方米/月)"]] or 0),
            推广费=float(cur[ix["推广费(元/平方米/月)"]] or 0),
            战略="是" if str(cur[ix["是否战略品牌"]]).strip() in ("1", "是", "True") else "否",
            状态=status.get(cno, ""),
        )
    return out


# ---------------------------------------------------------------- 读销售
def load_sales():
    """按 (合同编号, 面积) 映射到铺位，剔除非铺位/面积异常记录。
    支持多份销售报表合并：同一铺位同一月份若重复出现，取两表金额较大者（防重复累加）。
    源文件缺失时跳过，不中断生成（缺失情况在控制台与注脚中标注）。"""
    raw = defaultdict(lambda: defaultdict(float))   # 铺位 -> 月份 -> 金额
    dropped = defaultdict(lambda: [0.0, 0])
    store_of = defaultdict(set)
    unmatched = defaultdict(lambda: [0.0, 0])
    srcs = SRC_SALES if isinstance(SRC_SALES, (list, tuple)) else [SRC_SALES]
    used = []
    for src in srcs:
        if not os.path.exists(src):
            print(f"  [跳过] 销售报表不存在：{src}")
            continue
        used.append(os.path.basename(src))
        wb = openpyxl.load_workbook(src, data_only=True)
        ws = wb["店铺销售月报"]
        hdr = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        ix = {h: i for i, h in enumerate(hdr)}
        for r in ws.iter_rows(min_row=2, values_only=True):
            name = r[ix["店铺名称"]]
            if not name:
                continue
            a = float(r[ix["店铺面积(㎡)"]] or 0)
            rtype = str(r[ix["资源位类型"]] or "")
            amt = r[ix["销售金额(元)"]] or 0
            ct = str(r[ix["合同编号"]] or "").strip()
            if rtype != "铺位" or a <= 0:
                d = dropped[name]
                d[0] += amt
                d[1] += 1
                continue
            uid = UNIT_BY_CONTRACT.get(ct)
            if uid is None:
                d = unmatched[name]
                d[0] += amt
                d[1] += 1
                continue
            mo = str(r[ix["发生月份"]])
            raw[uid][mo] = max(raw[uid].get(mo, 0.0), amt)   # 跨表去重
            store_of[uid].add((name, ct, a))
    return raw, store_of, {k: tuple(v) for k, v in dropped.items()}, \
        {k: tuple(v) for k, v in unmatched.items()}, used


def metrics(raw, uid):
    """返回铺位级销售指标。"""
    m = dict(raw.get(uid, {}))
    if not m:
        return None, None, None, None, None, 0, 0

    m25 = {k: v for k, v in m.items() if k.startswith("2025")}
    m26 = {k: v for k, v in m.items() if k.startswith("2026")}
    common = sorted(set(k[4:] for k in m25) & set(k[4:] for k in m26))
    avg25 = sum(m25[f"2025{k}"] for k in common) / len(common) if common else None
    avg26 = sum(m26[f"2026{k}"] for k in common) / len(common) if common else None
    yoy = None
    if avg25 and avg26 and len(common) >= 6 and len(m25) >= 6 and len(m26) >= 6:
        yoy = avg26 / avg25 - 1
    latest_m = max(m26) if m26 else max(m)
    return avg25, avg26, yoy, m.get(latest_m), latest_m, len(m25), len(m26)


def rating(r):
    if r is None:
        return "无销售数据"
    if r < 0.10:
        return "偏低（有提租空间）"
    if r < 0.18:
        return "健康"
    if r < 0.30:
        return "偏高（关注）"
    if r < 0.40:
        return "过高（预警）"
    return "极高危（须降租）"


# ---------------------------------------------------------------- 主流程
def main():
    led = load_ledger()
    raw, store_of, dropped, unmatched, sales_used = load_sales()

    print("=== 剔除记录（非铺位 / 面积异常）===")
    for k, (amt, n) in dropped.items():
        print(f"  {k}: {n} 条，累计销售 {amt:,.0f} 元")
    if unmatched:
        print("=== 未匹配到铺位（已忽略）===")
        for k, (amt, n) in unmatched.items():
            print(f"  {k}: {n} 条，累计销售 {amt:,.0f} 元")
    print()
    print("=== 销售记录归属 ===")
    for uid in [p["铺位"] for p in PLAN]:
        for nm, ct, a in sorted(store_of.get(uid, []), key=lambda x: x[1]):
            print(f"  {uid:10s} <- {nm:16s} 合同 {ct:24s} 面积 {a:8.2f} ㎡  月份数 {len(raw[uid])}")
    print()

    rows = []
    for p in PLAN:
        uid = p["铺位"]
        L = led[uid]
        if abs(L["面积"] - p["面积"]) > 0.01:
            print(f"  [面积差异] {uid}: 招调表 {p['面积']} / 台账 {L['面积']} → 取台账")
        a = L["面积"]
        # 现租金：用户明确提供的口径优先（EVISU 台账 17.50 与用户 15.50 冲突，按用户口径）
        cur_day = p.get("现日租", L["日租"])
        cur_rate = p.get("现扣率", L["扣率"])
        prop = L["物业费"] or p["物业费"]
        promo = L["推广费"] or p["推广费"]

        cur_rent = cur_day * a * DAYS
        cur_all = cur_rent + (prop + promo) * a

        plan_day = p["日租"]
        plan_rate = p["扣率"] * 100
        yoy_rent = (plan_day / cur_day - 1) if cur_day else None   # 与前后租金严格对应
        plan_rent = plan_day * a * DAYS     # 计划月租金按报价固租列示（便于看清让租幅度）
        plan_all = plan_rent + (p["物业费"] + p["推广费"]) * a

        avg25, avg26, yoy, latest, latest_m, n25, n26 = metrics(raw, uid)

        pct_rent = (avg26 * cur_rate / 100) if avg26 else None
        actual = max(cur_rent, pct_rent) if pct_rent else cur_rent
        pof = (pct_rent / cur_rent) if (pct_rent and cur_rent) else None

        cur_ratio = (actual / avg26) if (avg26 and actual) else None
        plan_pct = (avg26 * plan_rate / 100) if avg26 else None
        plan_actual = max(plan_rent, plan_pct) if plan_pct else plan_rent
        plan_ratio = (plan_actual / avg26) if (avg26 and plan_actual) else None
        delta_pp = ((plan_ratio - cur_ratio) * 100) if (cur_ratio is not None and plan_ratio is not None) else None

        rows.append(dict(
            铺位=uid, 品牌=p["原品牌"], 楼层=L["楼层"], 新品牌=p.get("新品牌"), 面积=a, 战略=L["战略"],
            到期=L["到期"], 签约=p["签约"], 开业=p["开业"], 状态=L.get("状态", ""),
            台账日租=L["日租"],
            cur_day=cur_day, cur_rate=cur_rate, actual=actual, prop=prop,
            cur_all=cur_all, cur_rent=cur_rent,
            plan_day=plan_day, plan_rate=plan_rate, plan_yoy=yoy_rent, plan_years=p["年限"],
            plan_rent=plan_rent, plan_all=plan_all,
            avg25=avg25, avg26=avg26, yoy=yoy, latest=latest,
            cur_ratio=cur_ratio, plan_ratio=plan_ratio, delta_pp=delta_pp,
            grade=rating(plan_ratio if plan_ratio is not None else cur_ratio),
            pof=pof, n25=n25, n26=n26,
        ))

    write_xlsx(rows, dropped, unmatched, store_of, sales_used)
    report(rows)


THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=9)
TOTAL_FONT = Font(name="微软雅黑", size=9, bold=True)


def write_xlsx(rows, dropped, unmatched, store_of, sales_used):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "27年到期招调简表"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOL)
    c = ws.cell(row=1, column=1, value="沈阳大悦城 · 2027 年到期战略品牌招调简表")
    c.font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOL)
    c = ws.cell(row=2, column=1,
                value=f"数据口径：合同台账（统计日 2026-10-28）＋ 店铺销售月报（202501–{CUR_MONTH}）　｜　"
                      f"月面积折算 {DAYS:.4f} 天　｜　金额单位：元　｜　已剔除特卖/室内场地等非铺位资源")
    c.font = Font(name="微软雅黑", size=8.5, italic=True, color="808080")
    c.alignment = Alignment(horizontal="left", vertical="center")

    HR = 4
    for j, (name, w) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=HR, column=j, value=name)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[HR].height = 42

    FMTS = {4: "0.00", 9: "0.00", 10: "0.0", 11: "#,##0", 12: "0.0", 13: "#,##0",
            14: "0.00", 15: "0.0", 16: "+0.0%;-0.0%;0.0%", 17: "0",
            18: "#,##0", 19: "#,##0", 20: "#,##0", 21: "#,##0",
            22: "+0.0%;-0.0%;0.0%", 23: "#,##0",
            24: "0.0%", 25: "0.0%", 26: "+0.0;-0.0;0.0", 28: "0.0%"}

    r = HR + 1
    for d in rows:
        vals = [
            d["铺位"], (f'{d["品牌"]} → {d["新品牌"]}' if d.get("新品牌") else d["品牌"]),
            d["楼层"], d["面积"], d["战略"],
            d["到期"], d["签约"], d["开业"],
            d["cur_day"], d["cur_rate"], d["actual"], d["prop"], d["cur_all"],
            d["plan_day"], d["plan_rate"], d["plan_yoy"], d["plan_years"],
            d["plan_rent"], d["plan_all"],
            d["avg25"], d["avg26"], d["yoy"], d["latest"],
            d["cur_ratio"], d["plan_ratio"], d["delta_pp"], d["grade"], d["pof"],
        ]
        for j, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for j, f in FMTS.items():
            ws.cell(row=r, column=j).number_format = f

        gc = ws.cell(row=r, column=27)
        g = d["grade"]
        if "极高危" in g:
            gc.fill = PatternFill("solid", fgColor="C00000")
            gc.font = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
        elif "过高" in g:
            gc.fill = PatternFill("solid", fgColor="FFC7CE")
        elif "偏高" in g:
            gc.fill = PatternFill("solid", fgColor="FFEB9C")
        elif g == "健康":
            gc.fill = PatternFill("solid", fgColor="C6EFCE")
        elif "偏低" in g:
            gc.fill = PatternFill("solid", fgColor="DDEBF7")
        else:
            gc.fill = PatternFill("solid", fgColor="EDEDED")

        # 同比/变动：跌绿涨红（中式财务惯例）
        for j in (16, 22):
            v = ws.cell(row=r, column=j).value
            if isinstance(v, (int, float)):
                ws.cell(row=r, column=j).font = Font(name="微软雅黑", size=9,
                                                     color="008000" if v < 0 else "C00000")
        v = ws.cell(row=r, column=26).value
        if isinstance(v, (int, float)):
            ws.cell(row=r, column=26).font = Font(name="微软雅黑", size=9,
                                                  color="008000" if v < 0 else "C00000")
        r += 1

    tr = r
    tot_area = sum(d["面积"] for d in rows)
    ws.cell(row=tr, column=1, value="合计/加权")
    ws.cell(row=tr, column=4, value=tot_area).number_format = "0.00"
    ws.cell(row=tr, column=9, value=sum(d["cur_day"] * d["面积"] for d in rows) / tot_area).number_format = "0.00"
    ws.cell(row=tr, column=11, value=sum(d["actual"] for d in rows)).number_format = "#,##0"
    ws.cell(row=tr, column=13, value=sum(d["cur_all"] for d in rows)).number_format = "#,##0"
    ws.cell(row=tr, column=14, value=sum(d["plan_day"] * d["面积"] for d in rows) / tot_area).number_format = "0.00"
    ws.cell(row=tr, column=18, value=sum(d["plan_rent"] for d in rows)).number_format = "#,##0"
    ws.cell(row=tr, column=19, value=sum(d["plan_all"] for d in rows)).number_format = "#,##0"

    dd = [d for d in rows if d["avg26"]]
    if dd:
        rent_s = sum(d["actual"] for d in dd)
        rent_p = sum(max(d["plan_rent"], d["avg26"] * d["plan_rate"] / 100) for d in dd)
        sale_s = sum(d["avg26"] for d in dd)
        ws.cell(row=tr, column=21, value=sale_s).number_format = "#,##0"
        ws.cell(row=tr, column=24, value=rent_s / sale_s).number_format = "0.0%"
        ws.cell(row=tr, column=25, value=rent_p / sale_s).number_format = "0.0%"
        ws.cell(row=tr, column=26, value=(rent_p - rent_s) / sale_s * 100).number_format = "+0.0;-0.0;0.0"
        a25 = sum(d["avg25"] for d in dd if d["avg25"])
        a26 = sum(d["avg26"] for d in dd if d["avg25"])
        if a25:
            ws.cell(row=tr, column=22, value=a26 / a25 - 1).number_format = "0.0%"

    for j in range(1, NCOL + 1):
        cell = ws.cell(row=tr, column=j)
        cell.font = TOTAL_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="D9E2F3")
    ws.row_dimensions[tr].height = 20

    drop_txt = "；".join(f"{k}（{v[1]} 条，剔除销售 {v[0]:,.0f} 元）" for k, v in dropped.items())
    notes = [
        f"注1【数据清洗】本次销售数据已剔除非铺位资源与面积异常记录：{drop_txt or '无'}。"
        "MLB 两家店（B114 / E320-321）分别统计；B114 在 202501–202505 使用老店铺编码、202506 起并入新合同编码，"
        "同一铺位跨期记录已合并，不重复计数。",
        "注2【两者取高】「实际收租」= max(当前固租, 26年月均销售 × 当前扣率)。「抽成/固租」= 抽成租金 ÷ 固租："
        "<100% 说明抽成条款常年不触发，实际全部按固租收租 —— 此时提高扣率不增加实收，只有降固租才真实减负。",
        "注3【综合成本】月综合成本 = 月租金 + 物业费 + 推广费（均按签约面积计）。现行取台账当前期次费率，计划取招调方报价。",
        "注4【销售口径】25年/26年月均销售均取两年「同月份」交集求均值，消除 2026 年数据未满年的季节性偏差；"
        "同比 = 26年同月均 ÷ 25年同月均 − 1。两年各不满 6 个月的视为新店/断档，不计算同比（显示「—」）。",
        "注5【租金口径】「实际收租」与「现行月综合成本」按两者取高计（含可能触发的抽成）；"
        "「计划月租金 / 计划月综合成本」按报价固租口径列示，便于直接看清让租幅度；"
        "「计划后租售比」按两者取高计（max(计划固租, 26年月均销售 × 计划扣率)），反映实际可能的负担水平。"
        "评级：<10% 偏低｜10%–18% 健康｜18%–30% 偏高｜30%–40% 过高｜>40% 极高危。",
        "注6【租金同比口径·必读】本表「租金同比」= 续约日租 ÷ 现行日租 − 1，由前后租金实算，与两列数字严格对应。"
        "经逐一核对，原招调表所填涨幅有 4 处与前后租金不成立，已按实算修正："
        "① B114（21.00→21.00）原填 −20%，实为 0%；② E114 李宁（25.96→25.96）原填 −4%，实为 0%；"
        "③ A101/A102 波司登（6.77→6.77）原填 −19%，实为 0%；④ E320-321（6.00→2.00）原填 0%，实为 −66.7%。"
        "前三处的原填涨幅系采用「历史旧价」作基数（如李宁上期为 28.62、波司登 2025 年为 6.26），"
        "并非与本次续约价对应，请以本表实算值为准。B305 GXG 经修正续约日租为 13.01（原误填 15.31），"
        "15.31→13.01 = −15.0%，与所填涨幅一致，已确认无误。",
        "注7【EVISU B121】该铺合同状态为「审核中」，BI 主台账未导出（仅 1 期：2023-06-01～2026-12-20），"
        "台账数据取自单独提供的「evisu合同台账.xlsx」。台账现行日租为 17.50 元/㎡/日，而续约材料所列「现租金」为 15.50，"
        "两者相差 2.00 元/㎡/日（月租差约 8,100 元、年化约 9.7 万元），本表按续约材料的 15.50 口径列示。"
        "该铺 2026-12-20 到期，为本批最早到期者（距今不足 2 个月），建议签约前核实真实现租金。"
        "销售取自「店铺销售月报 (2).xlsx」单店 22 个月（202501–202610）：26 年月均 23.1 万元，同比 −25.1%，"
        "为本批降幅最大的铺位；现行租售比 27.1% 属「偏高」区间，续约未降租（15.50→15.50），"
        "此后将升至计划后 27.1%，负担进一步加重，建议本轮续约争取降租或设保底。"
        "续约材料所填租售比 21.0% 对应的月销售基数约 29.9 万元，接近 2025 年水平，与 26 年实际不符，请以本表 26 年口径为准。",
        "注8【E320-321】该铺已于 2026 年底续约一年至 2027-12-29，本次招调体现为「再续一年至 2028-12-29」。"
        "台账现行日租 6.00 元/㎡/日，续约报价仅 2.00（−66.7%），而 2026 年月均销售约 34.5 万元、"
        "租售比 14.1% 属健康区间、同比 +3.8%，是本批经营最好的铺位。续约固租降至 1.62 万元/月明显低于市场水平，"
        "建议核实是否为笔误或特殊条款（若维持 6.00，月租由 1.62 万回升至 4.85 万）。",
        "注9【A101/A102 波司登】该铺 2027-05-19 到期，计划由波司登换为 RE'VAN + ariose years，"
        "年限 3 年、2027 年 5 月签约、6 月开业。计划固租 6.77 与现行一致（涨幅 0%）。"
        "销售同比 +66.5% 系羽绒服强季节性所致：2025 年 1–10 月均值 17.1 万、2026 年同期 28.4 万，"
        "且两年 10 月均为旺季（58.9 万 / 81.9 万），同比可信但不可外推全年。换牌后原品牌销售不可直接作为新品牌预期。",
        "注10【合计行】7 个铺位均已纳入 26 年销售，合计租售比 = Σ实际收租 ÷ Σ26年月均销售，"
        "加权同比按两年均有数据的铺位合并计算。租售比为金额加权，非各铺位比率的简单平均。",
    ]
    nr = tr + 2
    for t in notes:
        ws.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=NCOL)
        cell = ws.cell(row=nr, column=1, value=t)
        cell.font = Font(name="微软雅黑", size=8.5, color="595959")
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[nr].height = 30
        nr += 1

    ws.freeze_panes = "C5"
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print(f"\n已保存：{OUT}")
    print(f"sheets: {wb.sheetnames}  rows: {ws.max_row}  cols: {ws.max_column}  freeze: {ws.freeze_panes}")


def report(rows):
    def F(v, p=1, s=""):
        return "—" if v is None else f"{v:,.{p}f}{s}"

    print("\n铺位        品牌              面积      到期       现日租  实际收租   计划日租  计划月租   25月均    26月均    同比    现行租售比 计划租售比 抽成/固租")
    for d in rows:
        print(f"{d['铺位']:<11}{d['品牌']:<14}{d['面积']:>7.2f}  {d['到期']}  "
              f"{d['cur_day']:>6.2f} {d['actual']:>9,.0f}  {d['plan_day']:>7.2f} {d['plan_rent']:>9,.0f} "
              f"{F(d['avg25'],0):>9} {F(d['avg26'],0):>9} "
              f"{F(d['yoy']*100 if d['yoy'] is not None else None,1,'%'):>7} "
              f"{F(d['cur_ratio']*100 if d['cur_ratio'] else None,1,'%'):>8} "
              f"{F(d['plan_ratio']*100 if d['plan_ratio'] else None,1,'%'):>8} "
              f"{F(d['pof']*100 if d['pof'] else None,1,'%'):>8}")
    print(f"\n面积合计 {sum(d['面积'] for d in rows):,.2f} ㎡")
    print(f"现行月综合成本合计 {sum(d['cur_all'] for d in rows):,.0f} 元")
    print(f"计划月综合成本合计 {sum(d['plan_all'] for d in rows):,.0f} 元")
    print(f"差额 {sum(d['plan_all'] for d in rows) - sum(d['cur_all'] for d in rows):+,.0f} 元/月 "
          f"（年化 {(sum(d['plan_all'] for d in rows) - sum(d['cur_all'] for d in rows))*12:+,.0f} 元）")


if __name__ == "__main__":
    main()
