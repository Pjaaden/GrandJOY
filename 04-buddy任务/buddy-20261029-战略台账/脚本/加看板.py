#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
在飞书台账表首部插入「台账看板」子表（关键指标 + 区域分布 + 品牌 Top15）。

用法：
    python 加看板.py --表格token <token>
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 台账核心 import build_all  # noqa: E402

LARK_BIN = os.path.expanduser(
    "~/.grandjoybuddy/binaries/node/workspace/node_modules/.bin/lark-cli"
)


def run(cmd, stdin_data=None):
    env = dict(os.environ)
    env["PATH"] = os.path.dirname(LARK_BIN) + os.pathsep + env.get("PATH", "")
    p = subprocess.run(cmd, env=env, input=stdin_data, capture_output=True, text=True)
    try:
        return json.loads(p.stdout.strip() or "{}")
    except json.JSONDecodeError:
        return {"raw": p.stdout, "stderr": p.stderr[-1500:]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--表格token", required=True)
    ap.add_argument("--源文件", default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "源数据", "运动男装合同台账1029.xlsx"))
    args = ap.parse_args()

    data = build_all(args.源文件)
    contract = data["合同级汇总"]
    brand = data["品牌汇总"]
    detail = data["合同台账明细"]

    total_area = sum(c["签约面积(平方米)"] or 0 for c in contract)
    total_rent = sum(c["期间固定租金总额(元)"] or 0 for c in contract)
    total_cost = sum(c["期间成本合计(元)"] or 0 for c in contract)

    # 1) 复用已存在的「台账看板」子表；没有才新建到最前
    info = run([LARK_BIN, "sheets", "+workbook-info", "--as", "user",
                "--url", f"https://my.feishu.cn/sheets/{args.表格token}"])
    sid = None
    for sh in (info.get("data", {}) or {}).get("sheets", []) or []:
        if sh.get("sheet_name") == "台账看板":
            sid = sh.get("sheet_id")
            break

    if sid is None:
        res = run([LARK_BIN, "sheets", "+sheet-create", "--as", "user",
                   "--url", f"https://my.feishu.cn/sheets/{args.表格token}",
                   "--title", "台账看板", "--index", "0",
                   "--row-count", "80", "--col-count", "8"])
        if not res.get("ok"):
            print("[ERR] 建看板失败:", json.dumps(res, ensure_ascii=False)[:1200])
            return
        d = res.get("data", {}) or {}
        sid = (d.get("sheet") or {}).get("sheet_id") or d.get("sheet_id")
        print("[OK] 看板子表已建, sheet_id =", sid)
    else:
        print("[OK] 复用已有看板子表, sheet_id =", sid)

    # 1.5) 写入前先清空看板区域，保证多次更新幂等（不留上一版残留）
    clr = run([LARK_BIN, "sheets", "+cells-batch-clear", "--as", "user",
               "--url", f"https://my.feishu.cn/sheets/{args.表格token}",
               "--ranges", '["台账看板!A1:H200"]', "--yes"])
    print("[OK] 看板区已清空:", clr.get("ok"))

    # 2) 组装单元格
    cells: list[list[dict]] = []

    def row(*items):
        cells.append([i if isinstance(i, dict) else {"value": i} for i in items])

    def blank(n):
        # 注意：cells 数组里不接受 {"value": null}（schema 校验会拒绝），
        # 空位必须用空对象 {} 占位，语义是「该格不做任何修改、保留原值」。
        return [{} for _ in range(n)]

    row({"value": "全国招商总部 · 运动/男装板块 战略品牌合同台账",
         "cell_styles": {"font_size": 16, "font_weight": "bold", "font_color": "#1F3864",
                         "horizontal_alignment": "left", "vertical_alignment": "middle"}})
    cells.append(blank(6))
    row({"value": f"数据版本 {data['版本记录'][0]['数据版本']}　|　明细 {len(detail)} 行 / 合同 {len(contract)} 份　|　生成 {dt.datetime.now():%Y-%m-%d %H:%M}",
         "cell_styles": {"font_size": 10, "font_color": "#666666", "horizontal_alignment": "left"}})
    cells.append(blank(6))

    kpis = [
        ("合同数（份）", len(contract), "#,##0"),
        ("签约面积合计（㎡）", round(total_area, 2), "#,##0.00"),
        ("期间固定租金合计（元）", round(total_rent, 2), "#,##0.00"),
        ("期间成本合计（元）", round(total_cost, 2), "#,##0.00"),
        ("覆盖项目数", len({d["项目名称"] for d in detail if d["项目名称"]}), "#,##0"),
        ("覆盖品牌数", len({d["品牌/行业"] for d in detail if d["品牌/行业"]}), "#,##0"),
    ]
    for i, (k, v, fmt) in enumerate(kpis):
        col = 1 + (i % 3) * 2
        r = 5 + (i // 3) * 3
        while len(cells) < r - 1:
            cells.append(blank(6))
        while len(cells) == r - 1:
            cells.append([{"value": k, "cell_styles": {"font_size": 10, "font_color": "#666666",
                                                       "horizontal_alignment": "left"}},
                          {}])
        while len(cells) == r:
            cells.append([{"value": v, "cell_styles": {"font_size": 15, "font_weight": "bold",
                                                       "font_color": "#1F3864", "number_format": fmt,
                                                       "horizontal_alignment": "left"}},
                          {}])
        cells.append(blank(6))

    while len(cells) < 11:
        cells.append(blank(6))

    # 区域分布
    row({"value": "一、区域分布（合同级口径，面积不重复计数）",
         "cell_styles": {"font_size": 12, "font_weight": "bold", "font_color": "#1F3864",
                         "horizontal_alignment": "left"}})
    row(*[{"value": h, "cell_styles": {"font_weight": "bold", "background_color": "#E8EEF7",
                                       "horizontal_alignment": "center",
                                       "border_styles": {"bottom": {"style": "solid", "weight": "thin", "color": "#D6DCE4"}}}}
             for h in ["所属区域", "合同数", "签约面积（㎡）", "期间固定租金（元）", "期间成本合计（元）", "成本占比"]])
    agg = {}
    for c in contract:
        b = agg.setdefault(c["所属区域"] or "(未标注)", {"n": 0, "area": 0.0, "rent": 0.0, "cost": 0.0})
        b["n"] += 1
        b["area"] += c["签约面积(平方米)"] or 0
        b["rent"] += c["期间固定租金总额(元)"] or 0
        b["cost"] += c["期间成本合计(元)"] or 0
    for k, v in sorted(agg.items(), key=lambda x: -x[1]["cost"]):
        row({"value": k, "cell_styles": {"horizontal_alignment": "left"}},
            {"value": v["n"], "cell_styles": {"number_format": "#,##0", "horizontal_alignment": "right"}},
            {"value": round(v["area"], 2), "cell_styles": {"number_format": "#,##0.00", "horizontal_alignment": "right"}},
            {"value": round(v["rent"], 2), "cell_styles": {"number_format": "#,##0.00", "horizontal_alignment": "right"}},
            {"value": round(v["cost"], 2), "cell_styles": {"number_format": "#,##0.00", "horizontal_alignment": "right"}},
            {"value": (v["cost"] / total_cost) if total_cost else 0,
             "cell_styles": {"number_format": "0.0%", "horizontal_alignment": "right"}})

    cells.append(blank(6))
    row({"value": "二、品牌 Top15（按期间成本合计）",
         "cell_styles": {"font_size": 12, "font_weight": "bold", "font_color": "#1F3864",
                         "horizontal_alignment": "left"}})
    row(*[{"value": h, "cell_styles": {"font_weight": "bold", "background_color": "#E8EEF7",
                                       "horizontal_alignment": "center",
                                       "border_styles": {"bottom": {"style": "solid", "weight": "thin", "color": "#D6DCE4"}}}}
             for h in ["品牌/行业", "合同数", "项目数", "签约面积（㎡）", "期间成本合计（元）", "平均抽成扣率"]])
    for b in brand[:15]:
        rate = b["平均抽成扣率"]
        row({"value": b["品牌/行业"], "cell_styles": {"horizontal_alignment": "left"}},
            {"value": b["合同数"], "cell_styles": {"number_format": "#,##0", "horizontal_alignment": "right"}},
            {"value": b["项目数"], "cell_styles": {"number_format": "#,##0", "horizontal_alignment": "right"}},
            {"value": b["签约面积合计(平方米)"], "cell_styles": {"number_format": "#,##0.00", "horizontal_alignment": "right"}},
            {"value": b["期间成本合计(元)"], "cell_styles": {"number_format": "#,##0.00", "horizontal_alignment": "right"}},
            {"value": rate, "cell_styles": {"number_format": "0.00%", "horizontal_alignment": "right"} if rate is not None else {}})

    cells.append(blank(6))
    row({"value": "三、使用与更新说明",
         "cell_styles": {"font_size": 12, "font_weight": "bold", "font_color": "#1F3864",
                         "horizontal_alignment": "left"}})
    notes = [
        "1. 明细粒度 = 合同期次（合同期间：1=第一期，2=第二期……）。起租日/到期日为该期次边界，非整份合同边界。",
        "2. 面积类指标看「合同级汇总」「区域业态汇总」，避免期次重复计数；「合同台账明细」为期次级。",
        "3. 「数据校验」列：一致 / 口径差异·两者取高 / 免租期超期·负值 / 抽成制·BI无固定租金 / 口径差异。",
        "4. 抽成扣率以小数存储（0.16 = 16%）可直接计算；抽成制合同无固定租金，日租金与固定租金总额为空属正常。",
        "5. 更新方式：新版 BI 导出放入 源数据/ 目录 → 运行 脚本/更新台账.sh → 本地 xlsx 重生成并覆盖回写飞书本表。",
        "6. 飞书位置：我的空间 / 大悦商业 / 战略台账 / 本表。",
    ]
    for n in notes:
        row({"value": n, "cell_styles": {"font_size": 10, "font_color": "#333333", "horizontal_alignment": "left"}})

    payload = [{
        "sheet_id": sid,
        "range": f"A1:F{len(cells)}",
        "cells": cells,
    }]
    res = run([LARK_BIN, "sheets", "+cells-set", "--as", "user",
               "--url", f"https://my.feishu.cn/sheets/{args.表格token}",
               "--writes", "-"], stdin_data=json.dumps(payload, ensure_ascii=False))
    print("[看板写入]", json.dumps(res, ensure_ascii=False)[:800])

    # 3) 版式收尾：列宽 + 隐藏网格线（清空会重置样式，每次更新后重新设）
    url = f"https://my.feishu.cn/sheets/{args.表格token}"
    run([LARK_BIN, "sheets", "+cols-resize", "--as", "user", "--url", url,
         "--sheet-name", "台账看板", "--range", "A:F", "--width", "190"])
    run([LARK_BIN, "sheets", "+cols-resize", "--as", "user", "--url", url,
         "--sheet-name", "台账看板", "--range", "A:A", "--width", "250"])
    run([LARK_BIN, "sheets", "+sheet-hide-gridline", "--as", "user", "--url", url,
         "--sheet-name", "台账看板"])
    print("[OK] 看板版式已设置（列宽 / 隐藏网格线）")


if __name__ == "__main__":
    main()
