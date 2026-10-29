#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全国招商总部 · 运动/男装板块 战略品牌合同台账 —— 核心处理模块

用途
----
从 BI 系统导出的「运动男装合同台账」原始 xlsx 中，抽取并标准化出招商总部战略台账。
BI 导出的每一行是一个「合同期间（期次）」，本模块保留该粒度作为明细台账，
同时向上汇总出合同级、区域级、品牌级三个维度，为后续动态分析与区域对比打底。

口径要点（已用 1029 版全量 925 行反推校验）
------------------------------------------
1. 明细粒度 = 合同期次。源表「起租日/到期日」是期次级边界，「合同期间」是期次序号；
   「承租期-起租日/到期日」才是整份合同的总边界，仅用于合同级汇总。
2. 期间固定租金总额：BI 值 = 单位日租金 × 签约面积 × (期间天数 − 免租期天数)。
   差异常见「两者取高」类合同（BI 侧另有取高逻辑）与免租期超出期间长度的负值情形。
3. 期间物业费总额 / 期间推广费总额：BI 值 ≈ 单价 × 签约面积 × 期间天数 ÷ (365/12)。
   差额主要来自 BI 按自然月对齐与按天折算的取整差异。
4. 面积类指标一律在「合同级」汇总，避免期次重复计数。
"""

from __future__ import annotations

import datetime as dt
import math
import re
from collections import defaultdict
from typing import Any

import openpyxl

# ── 常量 ────────────────────────────────────────────────────────────────────
DAYS_PER_MONTH = 365 / 12  # BI 月面积折算口径

# 明细台账 17 个核心表头（顺序即用户指定顺序，不可改动）
CORE_HEADERS = [
    "项目名称",
    "楼层",
    "资源位编号",
    "品牌/行业",
    "合同状态",
    "签约面积",
    "商户名称",
    "起租日",
    "到期日",
    "单位日租金(元/平方米/日)",
    "抽成扣率",
    "物业费(元/平方米/月)",
    "推广费(元/平方米/月)",
    "合同期间",
    "期间固定租金总额(元)",
    "期间物业费总额(元)",
    "期间推广费总额(元)",
]

# 扩展列（服务于后续动态分析 / 区域对比，排在 17 列之后）
EXT_HEADERS = [
    "所属区域",
    "三级业态",
    "合同编号",
    "合同名称",
    "是否战略品牌",
    "缴付租金方法",
    "日租金计算方式",
    "免租期天数(天)",
    "期间天数(天)",
    "期间月数",
    "固定租金校验差异(元)",
    "数据校验",
]

# 合同级汇总表头
CONTRACT_HEADERS = [
    "合同编号",
    "合同名称",
    "项目名称",
    "所属区域",
    "资源位编号",
    "品牌/行业",
    "商户名称",
    "合同状态",
    "签约面积(平方米)",
    "三级业态",
    "期次数",
    "合同起租日",
    "合同到期日",
    "合同总天数",
    "期间固定租金总额(元)",
    "期间物业费总额(元)",
    "期间推广费总额(元)",
    "期间成本合计(元)",
]

# 区域 × 业态 汇总表头
REGION_HEADERS = [
    "所属区域",
    "三级业态",
    "合同数",
    "铺位行数",
    "签约面积合计(平方米)",
    "期间固定租金总额(元)",
    "期间物业费总额(元)",
    "期间推广费总额(元)",
    "期间成本合计(元)",
    "平均单位日租金(元/平方米/日)",
    "平均抽成扣率",
]

# 品牌汇总表头
BRAND_HEADERS = [
    "品牌/行业",
    "所属区域数",
    "项目数",
    "合同数",
    "签约面积合计(平方米)",
    "期间固定租金总额(元)",
    "期间成本合计(元)",
    "平均单位日租金(元/平方米/日)",
    "平均抽成扣率",
    "涉及三级业态",
]

# 字段说明表头
SPEC_HEADERS = ["序号", "字段名称", "源表来源列", "口径与计算规则", "备注"]

# 版本记录表头
VERSION_HEADERS = [
    "数据版本",
    "统计日期",
    "源文件名",
    "明细行数",
    "合同数",
    "项目数",
    "区域数",
    "生成时间",
    "更新人",
    "备注",
]


# ── 工具函数 ────────────────────────────────────────────────────────────────
def _num(v: Any) -> float | None:
    if v is None or v == "":
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(f) else f


def _txt(v: Any) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def _date(v: Any) -> dt.date | None:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and v.strip():
        s = v.strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
            try:
                return dt.datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


def _dstr(v: Any) -> str | None:
    d = _date(v)
    return d.isoformat() if d else None


def _round(v: float | None, nd: int = 2) -> float | None:
    return None if v is None else round(v, nd)


# ── 读取源表 ────────────────────────────────────────────────────────────────
def _locate_sheet(wb: openpyxl.Workbook) -> str:
    """找到含「合同编号」表头的工作表（源表名可能随版本变化）。"""
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            if row and any(str(c).strip() == "合同编号" for c in row if c is not None):
                return ws.title
    return wb.sheetnames[0]


def load_source(path: str) -> tuple[list[dict], dict]:
    """读取 BI 导出的合同台账，返回 (原始行列表, 源元信息)。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = _locate_sheet(wb)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("源文件为空")
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = [dict(zip(header, r)) for r in rows[1:] if any(c is not None and str(c).strip() != "" for c in r)]
    wb.close()

    stat_dates = {_txt(r.get("统计日期")) for r in data}
    stat_dates.discard(None)
    meta = {
        "sheet": sheet,
        "rows": len(data),
        "stat_date": sorted(stat_dates)[-1] if stat_dates else "",
        "file": path.split("/")[-1],
    }
    return data, meta


# ── 明细台账构建 ────────────────────────────────────────────────────────────
def build_detail(src: list[dict]) -> list[dict]:
    """把原始行标准化为明细台账行（含 17 核心列 + 扩展列 + 校验列）。"""
    out: list[dict] = []

    for r in src:
        s_date = _date(r.get("起租日"))
        e_date = _date(r.get("到期日"))
        days = (e_date - s_date).days + 1 if s_date and e_date else None

        rent = _num(r.get("单位日租金(元/平方米/日)"))
        area = _num(r.get("签约面积"))
        pf = _num(r.get("物业费(元/平方米/月)"))
        promo = _num(r.get("推广费(元/平方米/月)"))
        rate = _num(r.get("抽成扣率(%)"))
        free = _num(r.get("装修期/免租期(天)")) or 0.0
        pay = _txt(r.get("繳付租金方法"))

        rent_total = _num(r.get("期间固定租金总额(元)"))
        pf_total = _num(r.get("期间物业费总额(元)"))
        promo_total = _num(r.get("期间推广费总额(元)"))

        # 固定租金校验：BI 值 vs 日租金 × 面积 × (天数 − 免租期)
        diff = None
        if rent is not None and area is not None and days is not None and rent_total is not None:
            expect = rent * area * max(days - free, 0)
            diff = round(rent_total - expect, 2)

        # 校验结论分类
        if rent is None and rent_total is None:
            check = "抽成制·BI无固定租金"
        elif diff is None:
            check = "无法校验"
        elif abs(diff) <= max(1.0, abs(rent_total or 0) * 0.001):
            check = "一致"
        elif free > (days or 0):
            check = "免租期超期·负值"
        elif pay == "两者取高":
            check = "口径差异·两者取高"
        else:
            check = "口径差异"

        period = r.get("合同期间")
        try:
            period_i = int(str(period).strip())
        except (TypeError, ValueError):
            period_i = None

        months = round(days / DAYS_PER_MONTH, 3) if days else None

        out.append(
            {
                # 17 核心列
                "项目名称": _txt(r.get("项目名称")),
                "楼层": _txt(r.get("楼层")),
                "资源位编号": _txt(r.get("资源位编号")),
                "品牌/行业": _txt(r.get("品牌/行业")),
                "合同状态": _txt(r.get("合同状态")),
                "签约面积": _round(area),
                "商户名称": _txt(r.get("商户名称")),
                "起租日": _dstr(s_date),
                "到期日": _dstr(e_date),
                "单位日租金(元/平方米/日)": _round(rent, 4),
                "抽成扣率": rate,
                "物业费(元/平方米/月)": _round(pf, 4),
                "推广费(元/平方米/月)": _round(promo, 4),
                "合同期间": period_i,
                "期间固定租金总额(元)": _round(rent_total),
                "期间物业费总额(元)": _round(pf_total),
                "期间推广费总额(元)": _round(promo_total),
                # 扩展列
                "所属区域": _txt(r.get("所属区域")),
                "三级业态": _txt(r.get("三级业态")),
                "合同编号": _txt(r.get("合同编号")),
                "合同名称": _txt(r.get("合同名称")),
                "是否战略品牌": "是" if str(r.get("是否战略品牌")).strip() == "1" else "否",
                "缴付租金方法": pay,
                "日租金计算方式": _txt(r.get("日租金计算方式")),
                "免租期天数(天)": int(free) if free else 0,
                "期间天数(天)": days,
                "期间月数": months,
                "固定租金校验差异(元)": diff,
                "数据校验": check,
            }
        )
    return out


# ── 合同级汇总 ──────────────────────────────────────────────────────────────
def build_contract(detail: list[dict], src: list[dict]) -> list[dict]:
    """按合同编号汇总：面积取合同内最大值，成本类求和，日期取最早/最晚。"""
    # 用源表补合同级总边界（承租期-起租日/到期日）
    contract_span: dict[str, tuple[dt.date | None, dt.date | None]] = {}
    for r in src:
        cid = _txt(r.get("合同编号"))
        if not cid:
            continue
        cs, ce = _date(r.get("承租期-起租日")), _date(r.get("承租期-到期日"))
        prev = contract_span.get(cid)
        if prev is None:
            contract_span[cid] = (cs, ce)
        else:
            contract_span[cid] = (
                min([x for x in (prev[0], cs) if x], default=None),
                max([x for x in (prev[1], ce) if x], default=None),
            )

    groups: dict[str, list[dict]] = defaultdict(list)
    for row in detail:
        groups[row["合同编号"] or "(缺失合同编号)"].append(row)

    out = []
    for cid, rows in groups.items():
        first = rows[0]
        areas = [r["签约面积"] for r in rows if r["签约面积"] is not None]
        span = contract_span.get(cid, (None, None))
        if span[0] is None or span[1] is None:
            dates = [(_date(r["起租日"]), _date(r["到期日"])) for r in rows]
            span = (
                min([d[0] for d in dates if d[0]], default=None),
                max([d[1] for d in dates if d[1]], default=None),
            )
        total_days = (span[1] - span[0]).days + 1 if span[0] and span[1] else None

        def ssum(key: str) -> float:
            return round(sum(r[key] or 0 for r in rows), 2)

        rent_sum = ssum("期间固定租金总额(元)")
        pf_sum = ssum("期间物业费总额(元)")
        promo_sum = ssum("期间推广费总额(元)")

        out.append(
            {
                "合同编号": cid,
                "合同名称": first["合同名称"],
                "项目名称": first["项目名称"],
                "所属区域": first["所属区域"],
                "资源位编号": first["资源位编号"],
                "品牌/行业": first["品牌/行业"],
                "商户名称": first["商户名称"],
                "合同状态": first["合同状态"],
                "签约面积(平方米)": _round(max(areas) if areas else None),
                "三级业态": first["三级业态"],
                "期次数": len(rows),
                "合同起租日": span[0].isoformat() if span[0] else None,
                "合同到期日": span[1].isoformat() if span[1] else None,
                "合同总天数": total_days,
                "期间固定租金总额(元)": rent_sum,
                "期间物业费总额(元)": pf_sum,
                "期间推广费总额(元)": promo_sum,
                "期间成本合计(元)": round(rent_sum + pf_sum + promo_sum, 2),
            }
        )

    out.sort(key=lambda x: (x["所属区域"] or "", x["项目名称"] or "", x["品牌/行业"] or "", x["合同编号"] or ""))
    return out


# ── 区域 × 业态 汇总 ────────────────────────────────────────────────────────
def build_region(contract: list[dict], detail: list[dict]) -> list[dict]:
    """以合同级为基数统计区域×业态，避免期次重复计数面积。"""
    detail_by_cid: dict[str, list[dict]] = defaultdict(list)
    for r in detail:
        detail_by_cid[r["合同编号"] or "(缺失合同编号)"].append(r)

    cells: dict[tuple[str, str], dict] = defaultdict(
        lambda: {"合同数": 0, "铺位行数": 0, "面积": 0.0, "租金": 0.0, "物业": 0.0, "推广": 0.0,
                 "租金样本": [], "扣率样本": []}
    )

    for c in contract:
        key = (c["所属区域"] or "(未标注)", c["三级业态"] or "(未标注)")
        b = cells[key]
        b["合同数"] += 1
        b["铺位行数"] += c["期次数"]
        b["面积"] += c["签约面积(平方米)"] or 0
        b["租金"] += c["期间固定租金总额(元)"] or 0
        b["物业"] += c["期间物业费总额(元)"] or 0
        b["推广"] += c["期间推广费总额(元)"] or 0
        for d in detail_by_cid.get(c["合同编号"], []):
            if d["单位日租金(元/平方米/日)"] is not None:
                b["租金样本"].append(d["单位日租金(元/平方米/日)"])
            if d["抽成扣率"] is not None:
                b["扣率样本"].append(d["抽成扣率"])

    out = []
    for (region, biz), b in sorted(cells.items()):
        cost = b["租金"] + b["物业"] + b["推广"]
        out.append(
            {
                "所属区域": region,
                "三级业态": biz,
                "合同数": b["合同数"],
                "铺位行数": b["铺位行数"],
                "签约面积合计(平方米)": _round(b["面积"]),
                "期间固定租金总额(元)": _round(b["租金"]),
                "期间物业费总额(元)": _round(b["物业"]),
                "期间推广费总额(元)": _round(b["推广"]),
                "期间成本合计(元)": _round(cost),
                "平均单位日租金(元/平方米/日)": _round(sum(b["租金样本"]) / len(b["租金样本"]), 4) if b["租金样本"] else None,
                "平均抽成扣率": _round(sum(b["扣率样本"]) / len(b["扣率样本"]), 4) if b["扣率样本"] else None,
            }
        )
    return out


# ── 品牌汇总 ────────────────────────────────────────────────────────────────
def build_brand(contract: list[dict], detail: list[dict]) -> list[dict]:
    detail_by_cid: dict[str, list[dict]] = defaultdict(list)
    for r in detail:
        detail_by_cid[r["合同编号"] or "(缺失合同编号)"].append(r)

    groups: dict[str, list[dict]] = defaultdict(list)
    for c in contract:
        groups[c["品牌/行业"] or "(未标注)"].append(c)

    out = []
    for brand, cs in groups.items():
        rent = sum(c["期间固定租金总额(元)"] or 0 for c in cs)
        cost = sum(c["期间成本合计(元)"] or 0 for c in cs)
        area = sum(c["签约面积(平方米)"] or 0 for c in cs)
        rents, rates = [], []
        bizs = set()
        for c in cs:
            if c["三级业态"]:
                bizs.add(c["三级业态"])
            for d in detail_by_cid.get(c["合同编号"], []):
                if d["单位日租金(元/平方米/日)"] is not None:
                    rents.append(d["单位日租金(元/平方米/日)"])
                if d["抽成扣率"] is not None:
                    rates.append(d["抽成扣率"])
        out.append(
            {
                "品牌/行业": brand,
                "所属区域数": len({c["所属区域"] for c in cs if c["所属区域"]}),
                "项目数": len({c["项目名称"] for c in cs if c["项目名称"]}),
                "合同数": len(cs),
                "签约面积合计(平方米)": _round(area),
                "期间固定租金总额(元)": _round(rent),
                "期间成本合计(元)": _round(cost),
                "平均单位日租金(元/平方米/日)": _round(sum(rents) / len(rents), 4) if rents else None,
                "平均抽成扣率": _round(sum(rates) / len(rates), 4) if rates else None,
                "涉及三级业态": "、".join(sorted(bizs)),
            }
        )
    out.sort(key=lambda x: (-(x["期间成本合计(元)"] or 0), x["品牌/行业"] or ""))
    return out


# ── 字段说明 ────────────────────────────────────────────────────────────────
def build_spec() -> list[dict]:
    s = [
        ("项目名称", "项目名称", "原样取自 BI 导出的项目名称", "全国 29 个项目"),
        ("楼层", "楼层", "原样取自 BI 导出的楼层", "如 1F、2F、B1"),
        ("资源位编号", "资源位编号", "原样取自 BI 导出的资源位编号", "一铺多号时以分号分隔"),
        ("品牌/行业", "品牌/行业", "原样取自 BI 导出的品牌字段", "空值统一标注「(未标注)」"),
        ("合同状态", "合同状态", "原样取自 BI 导出的合同状态", "双签 / 已批准"),
        ("签约面积", "签约面积", "原样取自 BI 导出的签约面积（平方米）", "合同级汇总取合同内最大值，避免期次重复计数"),
        ("商户名称", "商户名称", "原样取自 BI 导出的签约商户", ""),
        ("起租日", "起租日", "期次级起租日（BI「起租日」列）", "不是整份合同的总起租日"),
        ("到期日", "到期日", "期次级到期日（BI「到期日」列）", "不是整份合同的总到期日"),
        ("单位日租金(元/平方米/日)", "单位日租金(元/平方米/日)", "原样取自 BI", "抽成制合同为空"),
        ("抽成扣率", "抽成扣率(%)", "原样取自 BI，以小数存储（0.16 = 16%）", "显示为百分比格式，可直接参与计算"),
        ("物业费(元/平方米/月)", "物业费(元/平方米/月)", "原样取自 BI", ""),
        ("推广费(元/平方米/月)", "推广费(元/平方米/月)", "原样取自 BI", ""),
        ("合同期间", "合同期间", "期次序号，转为整数", "1 = 第一期，2 = 第二期……"),
        ("期间固定租金总额(元)", "期间固定租金总额(元)", "原样取自 BI（该期次合计）", "口径：日租金 × 面积 × (期间天数 − 免租期)"),
        ("期间物业费总额(元)", "期间物业费总额(元)", "原样取自 BI（该期次合计）", "口径：物业费 × 面积 × 期间天数 ÷ (365/12)"),
        ("期间推广费总额(元)", "期间推广费总额(元)", "原样取自 BI（该期次合计）", "口径：推广费 × 面积 × 期间天数 ÷ (365/12)"),
        ("所属区域", "所属区域", "原样取自 BI", "北方 / 华东 / 西南 / 华南区域商业管理部"),
        ("三级业态", "三级业态", "原样取自 BI", "运动装、鞋类、男装、户外用品、服饰集合店等"),
        ("合同编号", "合同编号", "原样取自 BI", "合同级汇总主键"),
        ("合同名称", "合同名称", "原样取自 BI", ""),
        ("是否战略品牌", "是否战略品牌", "BI「是否战略品牌」1→是，0→否", ""),
        ("缴付租金方法", "繳付租金方法", "原样取自 BI", "固租 / 抽成 / 两者取高"),
        ("日租金计算方式", "日租金计算方式", "原样取自 BI", "日面积 / 月面积"),
        ("免租期天数(天)", "装修期/免租期(天)", "原样取自 BI", "装修期/免租期天数"),
        ("期间天数(天)", "派生", "到期日 − 起租日 + 1", ""),
        ("期间月数", "派生", "期间天数 ÷ (365/12)", "保留 3 位小数"),
        ("固定租金校验差异(元)", "派生", "BI 期间固定租金总额 − (日租金 × 面积 × (期间天数 − 免租期))", "接近 0 表示口径一致"),
        ("数据校验", "派生", "按校验差异与合同类型自动分类", "一致 / 口径差异·两者取高 / 免租期超期·负值 / 抽成制·BI无固定租金 / 无法校验"),
    ]
    return [
        {"序号": i, "字段名称": n, "源表来源列": src, "口径与计算规则": rule, "备注": note}
        for i, (n, src, rule, note) in enumerate(s, 1)
    ]


# ── 版本记录 ────────────────────────────────────────────────────────────────
def build_version(meta: dict, detail: list[dict], contract: list[dict], note: str = "") -> list[dict]:
    return [
        {
            "数据版本": meta.get("stat_date") or "未标注",
            "统计日期": meta.get("stat_date") or "",
            "源文件名": meta.get("file", ""),
            "明细行数": len(detail),
            "合同数": len(contract),
            "项目数": len({d["项目名称"] for d in detail if d["项目名称"]}),
            "区域数": len({d["所属区域"] for d in detail if d["所属区域"]}),
            "生成时间": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "更新人": meta.get("owner", "闫志坚"),
            "备注": note or "首次建账",
        }
    ]


# ── 主入口 ──────────────────────────────────────────────────────────────────
def build_all(path: str, note: str = "") -> dict[str, list[dict]]:
    """一次性产出全部子表数据。"""
    src, meta = load_source(path)
    detail = build_detail(src)
    contract = build_contract(detail, src)
    region = build_region(contract, detail)
    brand = build_brand(contract, detail)
    spec = build_spec()
    version = build_version(meta, detail, contract, note)

    return {
        "合同台账明细": detail,
        "合同级汇总": contract,
        "区域业态汇总": region,
        "品牌汇总": brand,
        "字段说明": spec,
        "版本记录": version,
    }


# 各子表的列顺序与数据类型（供 xlsx 生成与飞书写入共用）
SHEET_SPECS: dict[str, dict] = {
    "合同台账明细": {
        "headers": CORE_HEADERS + EXT_HEADERS,
        "dtypes": {
            "签约面积": "float64", "单位日租金(元/平方米/日)": "float64", "抽成扣率": "float64",
            "物业费(元/平方米/月)": "float64", "推广费(元/平方米/月)": "float64",
            "期间固定租金总额(元)": "float64", "期间物业费总额(元)": "float64",
            "期间推广费总额(元)": "float64", "合同期间": "int64", "免租期天数(天)": "int64",
            "期间天数(天)": "int64", "期间月数": "float64", "固定租金校验差异(元)": "float64",
        },
        "formats": {
            "签约面积": "#,##0.00", "单位日租金(元/平方米/日)": "#,##0.0000", "抽成扣率": "0.00%",
            "物业费(元/平方米/月)": "#,##0.0000", "推广费(元/平方米/月)": "#,##0.0000",
            "期间固定租金总额(元)": "#,##0.00", "期间物业费总额(元)": "#,##0.00",
            "期间推广费总额(元)": "#,##0.00", "期间月数": "#,##0.000", "固定租金校验差异(元)": "#,##0.00",
        },
    },
    "合同级汇总": {
        "headers": CONTRACT_HEADERS,
        "dtypes": {
            "签约面积(平方米)": "float64", "期次数": "int64", "合同总天数": "int64",
            "期间固定租金总额(元)": "float64", "期间物业费总额(元)": "float64",
            "期间推广费总额(元)": "float64", "期间成本合计(元)": "float64",
        },
        "formats": {
            "签约面积(平方米)": "#,##0.00", "期间固定租金总额(元)": "#,##0.00",
            "期间物业费总额(元)": "#,##0.00", "期间推广费总额(元)": "#,##0.00", "期间成本合计(元)": "#,##0.00",
        },
    },
    "区域业态汇总": {
        "headers": REGION_HEADERS,
        "dtypes": {
            "合同数": "int64", "铺位行数": "int64", "签约面积合计(平方米)": "float64",
            "期间固定租金总额(元)": "float64", "期间物业费总额(元)": "float64",
            "期间推广费总额(元)": "float64", "期间成本合计(元)": "float64",
            "平均单位日租金(元/平方米/日)": "float64", "平均抽成扣率": "float64",
        },
        "formats": {
            "签约面积合计(平方米)": "#,##0.00", "期间固定租金总额(元)": "#,##0.00",
            "期间物业费总额(元)": "#,##0.00", "期间推广费总额(元)": "#,##0.00",
            "期间成本合计(元)": "#,##0.00", "平均单位日租金(元/平方米/日)": "#,##0.0000",
            "平均抽成扣率": "0.00%",
        },
    },
    "品牌汇总": {
        "headers": BRAND_HEADERS,
        "dtypes": {
            "所属区域数": "int64", "项目数": "int64", "合同数": "int64",
            "签约面积合计(平方米)": "float64", "期间固定租金总额(元)": "float64",
            "期间成本合计(元)": "float64", "平均单位日租金(元/平方米/日)": "float64", "平均抽成扣率": "float64",
        },
        "formats": {
            "签约面积合计(平方米)": "#,##0.00", "期间固定租金总额(元)": "#,##0.00",
            "期间成本合计(元)": "#,##0.00", "平均单位日租金(元/平方米/日)": "#,##0.0000",
            "平均抽成扣率": "0.00%",
        },
    },
    "字段说明": {
        "headers": SPEC_HEADERS,
        "dtypes": {"序号": "int64"},
        "formats": {},
    },
    "版本记录": {
        "headers": VERSION_HEADERS,
        "dtypes": {"明细行数": "int64", "合同数": "int64", "项目数": "int64", "区域数": "int64"},
        "formats": {},
    },
}
