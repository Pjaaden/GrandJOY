# -*- coding: utf-8 -*-
"""
人工口径覆盖表 —— 战略品牌会议模板专用。

用途：台账数据与谈判/业务实际口径不一致时，在会前用一张小表统一覆盖，
不必改脚本。典型场景：
  · 合同状态「审核中」，BI 台账的费率不是最终执行价（如 EVISU 台账 17.50 / 实际 15.50）
  · 品牌名要按对外口径显示（Evisu → EVISU）
  · 台账费率缺失，需人工补

格式（.xlsx，第一个 sheet，第 1 行表头）：
  铺位号 | 现行日租 | 现行扣率 | 品牌 | 备注
  B121  | 15.50   | 16      | EVISU | 台账17.50为审核中费率，实际执行15.50

只填需要覆盖的列，留空的列用台账值。铺位号须与主表一致（清洗后的铺位号）。
"""
import openpyxl

FIELDS = {
    "铺位号": "_unit",
    "现行日租": "日租",
    "日租": "日租",
    "现行扣率": "扣率",
    "扣率": "扣率",
    "品牌": "品牌",
    "现行物业费": "物业费",
    "物业费": "物业费",
    "推广费": "推广费",
    "计划物业费": "计划物业费",
    "备注": "_note",
}

# 计划段专用覆盖（只影响「计划月综合成本」，不改现行列）
PLAN_FIELDS = {"计划物业费", "计划推广费"}


def read_override(path):
    """读取覆盖表，返回 铺位号 -> {字段: 值}。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hdr = [str(c or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    ix = {}
    for i, h in enumerate(hdr):
        key = FIELDS.get(h)
        if key:
            ix[key] = i
    unit_i = ix.pop("_unit", None)
    if unit_i is None:
        unit_i = next((i for i, h in enumerate(hdr) if "铺位" in h), None)
    if unit_i is None:
        raise SystemExit(f"覆盖表缺少「铺位号」列：{hdr}")

    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        u = str(r[unit_i] or "").strip()
        if not u:
            continue
        d = {}
        for k, i in ix.items():
            if i < len(r) and r[i] not in (None, ""):
                d[k] = r[i]
        if d:
            out[u] = d
    return out


def apply_override(info, ov):
    """把覆盖值套到合同信息上。返回 (info, 应用的字段列表)。"""
    if not ov:
        return info, []
    info = dict(info)
    used = []
    for k, v in ov.items():
        if k == "_note":
            continue
        if k in ("日租", "扣率", "物业费", "推广费", "计划物业费"):
            info[k] = float(v)
        else:
            info[k] = v
        used.append(k)
    return info, used
