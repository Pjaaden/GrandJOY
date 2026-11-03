# -*- coding: utf-8 -*-
"""
会前同步：从 BI 合同台账 + 店铺销售月报，一键生成「战略品牌到期与招调简表」。

═══════════════════════════════════════════════════════════════════
用法（三种场景）
═══════════════════════════════════════════════════════════════════
1) 只同步、不调整租金（会前摸底，大多数情况）
   python 会前同步.py --project 沈阳 --year 2027 --out 产出/沈阳-2027到期.xlsx

2) 带招调计划（已有报价，需要算计划后租售比与让租幅度）
   python 会前同步.py --project 沈阳 --year 2027 --plan 招调计划.xlsx --out ...

3) 指定销售报表 / 台账（默认自动找最新）
   python 会前同步.py --project 沈阳 --year 2027 \
       --ledger 台账.xlsx --sales 月报1.xlsx --sales 月报2.xlsx --out ...

═══════════════════════════════════════════════════════════════════
口径（重要，不要改）
═══════════════════════════════════════════════════════════════════
· 台账明细粒度 = 合同期次，同一合同号有多期 → 取「到期日最晚」的一期作现行口径
· 月面积折算 = 365/12 = 30.4167 天（大悦 BI 口径，不是 30 天）
· 「两者取高」：实际收租 = max(固租, 销售 × 扣率)
· 销售归属必须按 (合同编号, 面积) 映射 —— 报表「店铺名称」不可靠，同名品牌会串号
· 同比用「两年同月份交集」求均值再比，两年各不满 6 个月视为新店/断档，不计算
· 租金涨幅一律实算：计划日租 ÷ 现行日租 − 1，绝不采信招调表手填值
· 租售比评级：<10% 偏低 ｜ 10~18% 健康 ｜ 18~30% 偏高 ｜ 30~40% 过高 ｜ >40% 极高危
"""
import argparse
import glob
import os
import re
import sys
from collections import defaultdict

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 建模板 import (COLUMNS, PREVIEW, NCOL, FONT_CN, CENTER,  # noqa: E402
                     build_main, build_preview, fill_row, fill_total,
                     fill_preview)
from 铺位号 import clean_unit                                  # noqa: E402
from 口径覆盖 import read_override, apply_override             # noqa: E402

DAYS = 365 / 12

# 自动找源文件的候选目录
SEARCH_DIRS = [
    "/Users/zhijian/Downloads",
    "/Users/Shared/Projects/GrandJOY/04-buddy任务/buddy-20261029-战略台账/源数据",
    "/Users/Shared/Projects/GrandJOY/04-buddy任务/buddy-20261029-战略台账/产出",
]


# ---------------------------------------------------------------- 源文件发现
def find_files(patterns):
    """按 glob 模式在各候选目录里找，返回按修改时间倒序的路径列表。"""
    hits = []
    for d in SEARCH_DIRS:
        for p in patterns:
            hits += glob.glob(os.path.join(d, p))
    hits = sorted(set(hits), key=os.path.getmtime, reverse=True)
    return hits


def _is_detail_ledger(path):
    """校验是否为 BI 明细台账（有「合同台账」sheet 且含关键列）。
    产出汇总表 / 单表版虽然也叫台账，列结构不同，必须排除。
    注意：BI 导出文件不带 dimension，read_only 模式只能读到第 1 列，必须用普通模式。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        if "合同台账" not in wb.sheetnames:
            return False
        ws = wb["合同台账"]
        hdr = set()
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            hdr = {str(c) for c in row if c}
            break
        return {"项目名称", "合同编号", "到期日", "签约面积"}.issubset(hdr)
    except Exception:
        return False


def pick_ledger(explicit=None):
    if explicit:
        if not _is_detail_ledger(explicit):
            raise SystemExit(f"{explicit} 不是 BI 明细台账（需含「合同台账」sheet 及 项目名称/合同编号/到期日/签约面积 列）")
        return explicit
    # 优先源数据目录（BI 原始导出），其次 Downloads
    cands = find_files(["*合同台账*.xlsx", "*有效合同台账*.xlsx"])
    cands = [c for c in cands if not os.path.basename(c).startswith("~$")]
    good = [c for c in cands if _is_detail_ledger(c)]
    if not good:
        raise SystemExit("找不到可用的 BI 明细台账（需含「合同台账」sheet），请用 --ledger 指定")
    # 源数据目录优先
    for c in good:
        if "/源数据/" in c:
            return c
    return good[0]


def pick_sales(explicit=None):
    if explicit:
        return explicit
    hits = find_files(["店铺销售月报*.xlsx", "*销售月报*.xlsx"])
    hits = [h for h in hits if not os.path.basename(h).startswith("~$")]
    if not hits:
        raise SystemExit("找不到销售月报，请用 --sales 指定")
    return hits


# ---------------------------------------------------------------- 读台账
def read_ledger(path, project_kw):
    """返回 (by_contract, meta)。by_contract: 合同号 -> 期次行列表。"""
    by_ct = defaultdict(list)
    meta = {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["合同台账"] if "合同台账" in wb.sheetnames else wb[wb.sheetnames[0]]
    hdr = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    ix = {h: i for i, h in enumerate(hdr) if h}

    need = ["项目名称", "合同编号", "到期日", "签约面积", "单位日租金(元/平方米/日)", "抽成扣率(%)"]
    miss = [k for k in need if k not in ix]
    if miss:
        raise SystemExit(f"台账缺少必需列：{miss}")

    for r in ws.iter_rows(min_row=2, values_only=True):
        proj = r[ix["项目名称"]]
        if not proj or project_kw not in str(proj):
            continue
        cno = r[ix["合同编号"]]
        if not cno:
            continue
        by_ct[str(cno).strip()].append(r)
    meta["ix"] = ix
    meta["hdr"] = hdr
    # 台账「统计日期」形如 20261028，用于副标题说明数据截止日
    try:
        first = next(ws.iter_rows(min_row=2, max_row=2, values_only=True))
        si = ix.get("统计日期")
        meta["统计日期"] = str(first[si] or "") if si is not None else ""
    except StopIteration:
        meta["统计日期"] = ""
    return by_ct, meta


def norm_rate(v):
    v = float(v or 0)
    return v * 100 if v <= 1 else v


def current_period(rows, ix):
    """取到期日最晚的一期作为现行口径。"""
    def key(r):
        v = r[ix["到期日"]]
        return v if hasattr(v, "year") else str(v or "")
    return max(rows, key=key)


def contract_info(rows, ix):
    cur = current_period(rows, ix)

    def g(*names, default=None):
        for n in names:
            i = ix.get(n)
            if i is not None and i < len(cur) and cur[i] not in (None, ""):
                return cur[i]
        return default

    # 铺位号取值优先级：
    #   BI「资源位编号」是权威铺位号（000003B114 / E320;E321 / 000003A101;...）；
    #   「资源位名称」只是楼层位描述（B1F-21 / E3F-20;E3F-21），不是铺位号。
    #   因此优先取「资源位编号」；仅当它缺失或呈楼层位形态时才退回「资源位名称」。
    import re as _re
    FLOORISH = _re.compile(r"[A-Z]{1,2}\d{0,2}F-\d+")

    def pick_unit(code, name):
        code = str(code or "").strip()
        name = str(name or "").strip()
        if code and not _re.fullmatch(FLOORISH.pattern, code):
            return code
        return name or code

    raw_unit = pick_unit(g("资源位编号", default=""), g("资源位名称", default=""))

    return dict(
        楼层=str(g("楼层", "店铺所属楼层", default="") or "").strip(),
        面积=float(g("签约面积", default=0) or 0),
        到期=str(g("到期日", default=""))[:10],
        日租=float(g("单位日租金(元/平方米/日)", default=0) or 0),
        扣率=norm_rate(g("抽成扣率(%)", default=0)),
        物业费=float(g("物业费(元/平方米/月)", default=0) or 0),
        推广费=float(g("推广费(元/平方米/月)", default=0) or 0),
        战略="是" if str(g("是否战略品牌", default="")).strip() in ("1", "是", "True") else "否",
        品牌=str(g("品牌/行业", "店铺名称", default="") or ""),
        状态=str(g("合同状态", "合同状态名称历史", default="") or ""),
        铺位=raw_unit,
    )


# ---------------------------------------------------------------- 读销售
def read_sales(paths, resolve):
    """读取销售月报并归属到铺位。

    关键坑（必须遵守）：
      1. 报表「店铺名称」不可靠 —— 同名品牌（如 MLB）会混用多家店
      2. 同一铺位在生命周期内可能有多个合同编号（续约/改签），报表会**并行记录**：
         MLB B114 = HT2024040700089（202501-202505）＋ HT2024040700089-1（202506-202610）
         MLB E320-321 = HT2022052500143 ＋ HT2025120300063
         两者在重叠月份各有记录 → 必须按**铺位**聚合、同月取大者，绝不能跨合同号累加
      3. resolve(合同号, 面积, 品牌) -> 铺位，由调用方提供三级归属

    返回 (raw, dropped)
      raw: 铺位号 -> {月份: 金额}
    """
    raw = defaultdict(lambda: defaultdict(float))
    dropped = defaultdict(lambda: [0.0, 0])

    for p in paths:
        wb = openpyxl.load_workbook(p, data_only=True)
        if "店铺销售月报" not in wb.sheetnames:
            print(f"  [跳过] {os.path.basename(p)} 无「店铺销售月报」sheet", file=sys.stderr)
            continue
        ws = wb["店铺销售月报"]
        hdr = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        ix = {h: i for i, h in enumerate(hdr) if h}
        for k in ("销售金额(元)", "发生月份", "合同编号"):
            if k not in ix:
                raise SystemExit(f"{p} 缺少列：{k}")
        for r in ws.iter_rows(min_row=2, values_only=True):
            name = r[ix.get("店铺名称", 0)]
            if not name:
                continue
            a = float(r[ix.get("店铺面积(㎡)", 0)] or 0)
            rtype = str(r[ix.get("资源位类型", 0)] or "")
            amt = r[ix["销售金额(元)"]] or 0
            cno = str(r[ix["合同编号"]] or "").strip()
            brand = str(r[ix.get("品牌名称", 0)] or name)
            # 剔除非铺位（特卖/室内场地）与面积异常
            if (rtype and rtype != "铺位") or a <= 0:
                d = dropped[str(name)]
                d[0] += amt
                d[1] += 1
                continue
            uid = resolve(cno, a, brand)
            if uid is None:
                d = dropped[str(name)]
                d[0] += amt
                d[1] += 1
                continue
            mo = str(r[ix["发生月份"]])
            # 跨合同号 / 跨表：同铺位同月取较大者（并行记录，非累加）
            raw[uid][mo] = max(raw[uid].get(mo, 0.0), amt)
    return raw, {k: tuple(v) for k, v in dropped.items()}


def build_unit_map(by_ct, ix):
    """构建 合同号 → 铺位号 映射。

    同一铺位会有多个合同号（续约/改签），用「清洗后的铺位号 + 签约面积」做归并键，
    把面积视为同等（±0.5㎡ 容差）的合同归到同一铺位。

    另外：报表里会残留**已终止的老合同号**（例如 MLB 的 HT2024040700089、
    HT2022052500143），它们在当前台账里已查不到。对这类合同号，另建
    `area_brand` 兜底表：按 (面积, 品牌名归一化) 反查铺位。
    """
    unit_of = {}
    groups = []
    area_brand = {}          # (round(area,1), 品牌归一) -> 铺位
    for cno, rows in by_ct.items():
        info = contract_info(rows, ix)
        u = clean_unit(info["铺位"])
        area = info["面积"]
        matched = None
        for g in groups:
            if g["unit"] == u and abs(g["area"] - area) <= 0.5:
                matched = g
                break
        if matched is None:
            matched = {"unit": u, "area": area, "contracts": [], "brand": info["品牌"]}
            groups.append(matched)
        matched["contracts"].append(cno)
        unit_of[cno] = u
        key = (round(area, 1), _norm_brand(info["品牌"]))
        area_brand.setdefault(key, u)
    return unit_of, groups, area_brand


def _norm_brand(b):
    """品牌名归一化：去括号备注（MLB（B114）→ mlb）、去空格、转小写。"""
    s = re.sub(r"[（(].*?[）)]", "", str(b or ""))
    return s.replace(" ", "").lower()


def resolve_unit(cno, area, brand, unit_of, area_brand):
    """三级归属：合同号 → (面积,品牌) → 面积唯一匹配。都不中返回 None。"""
    if cno in unit_of:
        return unit_of[cno]
    k = (round(area, 1), _norm_brand(brand))
    if k in area_brand:
        return area_brand[k]
    cands = {u for (a, _), u in area_brand.items() if abs(a - round(area, 1)) <= 0.5}
    return cands.pop() if len(cands) == 1 else None


def min_yoy(m):
    """可比同比：两年同月份交集；两年各不足 6 个月不计算。"""
    m25 = {k: v for k, v in m.items() if k.startswith("2025")}
    m26 = {k: v for k, v in m.items() if k.startswith("2026")}
    common = sorted(set(k[4:] for k in m25) & set(k[4:] for k in m26))
    if not common:
        return None, None, None, 0, 0
    a25 = sum(m25[f"2025{k}"] for k in common) / len(common)
    a26 = sum(m26[f"2026{k}"] for k in common) / len(common)
    yoy = a26 / a25 - 1 if (a25 and a26 and len(m25) >= 6 and len(m26) >= 6) else None
    return a25, a26, yoy, len(m25), len(m26)


def grade(r):
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


# ---------------------------------------------------------------- 读招调计划（可选）
def read_plan(path):
    """招调计划表：铺位号/品牌/计划日租/计划扣率/计划年限/计划签约/计划开业/新品牌。
    列名模糊匹配，缺列则该项留空。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hdr = [str(c or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]

    def find(*kws):
        for i, h in enumerate(hdr):
            if any(k in h for k in kws):
                return i
        return None

    cols = dict(
        铺位=find("铺位", "资源位"),
        计划日租=find("续约租金", "计划日租"),
        计划扣率=find("续约扣率", "计划扣率"),
        计划年限=find("续约租期", "年限"),
        计划签约=find("签约"),
        计划开业=find("开业"),
        新品牌=find("新品牌"),
    )
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if cols["铺位"] is None:
            break
        key = str(r[cols["铺位"]] or "").strip()
        if not key:
            continue
        d = {}
        for k, i in cols.items():
            if i is None or i >= len(r):
                continue
            v = r[i]
            if k in ("计划日租",):
                d[k] = float(v) if isinstance(v, (int, float)) else None
            elif k == "计划扣率":
                d[k] = norm_rate(v) if isinstance(v, (int, float)) else None
            elif k == "计划年限":
                d[k] = int(v) if isinstance(v, (int, float)) else None
            else:
                d[k] = str(v).strip() if v is not None else None
        out[key] = d
    return out


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--project", required=True, help="项目名关键词，如「沈阳」")
    ap.add_argument("--year", type=int, default=2027, help="招调年份，默认 2027")
    ap.add_argument("--year-到期", dest="year_到期", action="store_true",
                    help="严格只取到期日在该年份内的合同（默认取 ≤ 该年底的全部待处理合同，"
                         "以覆盖上年底已到期但尚未续签的铺位，如 EVISU）")
    ap.add_argument("--ledger", help="合同台账路径（默认自动找最新）")
    ap.add_argument("--extra-ledger", action="append",
                    help="补录台账路径，可多次传。用于状态「审核中」等 BI 未导出、但会前必须纳入的合同")
    ap.add_argument("--sales", action="append", help="销售月报路径，可多次传（默认自动找全部）")
    ap.add_argument("--plan", help="招调计划表路径（可选）")
    ap.add_argument("--override",
                    help="人工口径覆盖表（铺位号/现行日租/现行扣率/品牌/现行物业费/推广费/备注）。"
                         "用于台账费率非最终执行价、品牌对外名称不一致等情况")
    ap.add_argument("--unit-order",
                    help="铺位显示顺序，逗号分隔，如 B114,B121,C119,B305。不传则按到期日排序")
    ap.add_argument("--out", required=True, help="输出 xlsx 路径")
    a = ap.parse_args()

    ledger = pick_ledger(a.ledger)
    sales = a.sales or pick_sales()
    print(f"台账：{ledger}")
    print(f"销售：{len(sales)} 份")
    for s in sales:
        print(f"   - {os.path.basename(s)}")

    by_ct, meta = read_ledger(ledger, a.project)
    ix = meta["ix"]

    # 补录台账外的合同（如状态「审核中」BI 未导出、单独提供的台账）
    if a.extra_ledger:
        for ep in a.extra_ledger:
            extra, meta2 = read_ledger(ep, a.project)
            n = 0
            for cno, rows in extra.items():
                if cno not in by_ct:
                    by_ct[cno] = rows
                    n += 1
            print(f"   [补录] {os.path.basename(ep)}：新增 {n} 个合同")

    unit_of, groups, area_brand = build_unit_map(by_ct, ix)
    raw, dropped = read_sales(
        sales, lambda cno, a, b: resolve_unit(cno, a, b, unit_of, area_brand))
    plan = read_plan(a.plan) if a.plan else {}
    ov = read_override(a.override) if a.override else {}
    if ov:
        print(f"口径覆盖：{len(ov)} 个铺位（{os.path.basename(a.override)}）")

    # 挑出目标年份到期 + 战略品牌
    # --year 是「计划签约年份」：默认取 到期日 <= 该年底 且 尚未签约的合同，
    # 也就是「2027 年招调会上要处理的铺位」——包含 2026 年底已到期、待续签的（如 EVISU）。
    picked = []
    seen_units = set()
    year_end = f"{a.year}-12-31"
    for cno, rows in by_ct.items():
        info = contract_info(rows, ix)
        exp = info["到期"]
        if a.year_到期:
            hit = exp.startswith(str(a.year))
        else:
            hit = exp <= year_end
        if not hit:
            continue
        if info["战略"] != "是":
            continue
        u = clean_unit(info["铺位"])
        if u in seen_units:          # 同铺位多合同只取一个
            continue
        seen_units.add(u)
        picked.append((cno, info, u))
    picked.sort(key=lambda x: x[1]["到期"])

    # 可选：按人工习惯的铺位顺序排列（会议逐铺过一遍，楼层/动线顺序更好讲）
    if a.unit_order:
        order = {u.strip(): i for i, u in enumerate(a.unit_order.split(","))}
        picked.sort(key=lambda x: (order.get(x[2], 999), x[1]["到期"]))

    if not picked:
        raise SystemExit(f"{a.project} 项目没有 {a.year} 年前到期（或 {a.year} 年到期）的战略品牌合同")

    label = f"{a.year} 年到期" if a.year_到期 else f"{a.year} 年招调窗口内到期（≤{year_end}）"
    print(f"\n{label}战略品牌：{len(picked)} 个铺位")
    if dropped:
        for k, (amt, n) in dropped.items():
            print(f"  [剔除] {k}：{n} 条，销售 {amt:,.0f} 元（非铺位/面积异常/无法归属）")

    rows = []
    for cno, info, uid in picked:
        # 人工口径覆盖：优先于台账值（台账费率可能与实际执行价不同）
        info, ov_used = apply_override(info, ov.get(uid))
        if ov_used:
            print(f"  [覆盖] {uid}：{'、'.join(ov_used)}")
        m = raw.get(uid, {})
        a25, a26, yoy, n25, n26 = min_yoy(m)
        m26 = [k for k in m if k.startswith("2026")]
        latest = m[max(m26)] if m26 else None
        p = plan.get(uid, {})

        area = info["面积"]
        cur_day, cur_rate = info["日租"], info["扣率"]
        cur_rent = cur_day * area * DAYS
        pct = a26 * cur_rate / 100 if a26 else None
        actual = max(cur_rent, pct) if pct else cur_rent
        pof = pct / cur_rent if (pct and cur_rent) else None

        plan_day = p.get("计划日租") or cur_day      # 无计划则默认维持现价
        plan_rate = p.get("计划扣率") or cur_rate
        plan_rent = plan_day * area * DAYS
        plan_pct = a26 * plan_rate / 100 if a26 else None
        plan_actual = max(plan_rent, plan_pct) if plan_pct else plan_rent

        cur_ratio = actual / a26 if a26 else None
        plan_ratio = plan_actual / a26 if a26 else None
        delta = (plan_ratio - cur_ratio) * 100 if (cur_ratio is not None and plan_ratio is not None) else None

        brand = info["品牌"]
        newb = p.get("新品牌")

        # 计划段物业/推广费：允许单独覆盖（续约后费率可能调整，与现行列不同）
        ovd = ov.get(uid, {})
        plan_fee = (float(ovd["计划物业费"]) if "计划物业费" in ovd else info["物业费"])
        plan_promo = (float(ovd["计划推广费"]) if "计划推广费" in ovd else info["推广费"])

        rows.append(dict(
            **{"铺位号": uid, "品牌": f"{brand} → {newb}" if newb else brand,
               "楼层": info["楼层"], "面积(㎡)": area, "战略品牌": info["战略"],
               "原合同到期": info["到期"],
               "计划签约": p.get("计划签约") or "—", "计划开业": p.get("计划开业") or "—",
               "现行日租": cur_day, "现行扣率": cur_rate, "实际收租": actual,
               "现行物业费": info["物业费"], "现行月综合成本": cur_rent + (info["物业费"] + info["推广费"]) * area,
               "计划日租": plan_day, "计划扣率": plan_rate,
               "租金同比": (plan_day / cur_day - 1) if cur_day else None,
               "计划年限": p.get("计划年限") or "—",
               "计划月租金": plan_rent,
               "计划月综合成本": plan_rent + (plan_fee + plan_promo) * area,
               "25年月均销售": a25, "26年月均销售": a26, "月均销售同比": yoy,
               "26年最新月销售": latest,
               "现行租售比": cur_ratio, "计划后租售比": plan_ratio, "租售比变动": delta,
               "租售比评级": grade(plan_ratio if plan_ratio is not None else cur_ratio),
               "抽成/固租": pof,
               # AC「手动备注」列刻意留空：这是人工手填列，任何自动内容都会在会前被覆盖掉
               "手动备注": None,
               }
        ))

    # 写文件
    # 副标题按人工终版口径写「数据是什么」而不是「文件名是什么」
    stat = meta.get("统计日期", "")
    stat_s = f"{stat[:4]}-{stat[4:6]}-{stat[6:8]}" if len(stat) == 8 else (stat or "—")
    mons = sorted({mo for m in raw.values() for mo in m})
    span = f"（{mons[0]}–{mons[-1]}）" if mons else ""
    wb = openpyxl.Workbook()
    ws = wb.active
    build_main(ws, project=a.project, title_suffix=f"{a.year} 年到期战略品牌招调简表",
               subtitle=f"数据口径：合同台账（统计日 {stat_s}）＋ 店铺销售月报{span}"
                        f"　｜　月面积折算 {DAYS:.4f} 天　｜　金额单位：元"
                        f"　｜　已剔除特卖/室内场地等非铺位资源")
    r = 5
    for d in rows:
        fill_row(ws, r, d)
        r += 1
    fill_total(ws, r, rows)
    pv = wb.create_sheet()
    build_preview(pv)
    fill_preview(pv, rows)

    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    wb.save(a.out)

    # 控制台速览
    print(f"\n{'铺位':<10}{'品牌':<16}{'面积':>8}  {'到期':<11}{'日租':>7}{'实收':>10}{'26月均':>10}{'同比':>8}{'租售比':>8}  评级")
    for d in rows:
        def F(v, p=1, s=""):
            return "—" if v is None else f"{v:,.{p}f}{s}"
        print(f"{d['铺位号']:<10}{d['品牌'][:14]:<16}{d['面积(㎡)']:>8.2f}  {d['原合同到期']:<11}"
              f"{d['现行日租']:>7.2f}{d['实际收租']:>10,.0f}{F(d['26年月均销售'],0):>10}"
              f"{F(d['月均销售同比']*100 if d['月均销售同比'] is not None else None,1,'%'):>8}"
              f"{F(d['计划后租售比']*100 if d['计划后租售比'] else None,1,'%'):>8}  {d['租售比评级']}")
    print(f"\n已保存：{a.out}")


if __name__ == "__main__":
    main()
