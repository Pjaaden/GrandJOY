#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
为飞书台账配置到期预警条件格式（红/橙/黄三级），并保证子表顺序。

用法：
    python 加条件格式.py --表格token <token>
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess

LARK_BIN = os.path.expanduser(
    "~/.grandjoybuddy/binaries/node/workspace/node_modules/.bin/lark-cli"
)

# (预警文字, 背景色, 字体色)
RULES = [
    ("🔴 90天内到期", "#FFE3E3", "#B42318"),
    ("🟠 半年内到期", "#FFF3E0", "#B54708"),
    ("🟡 一年内到期", "#FFFCE6", "#A16207"),
]
# 续约优先级（P0 最急 → P3）
PRIORITY_RULES = [
    ("P0 立即启动", "#FFE3E3", "#B42318"),
    ("P1 半年内启动", "#FFF3E0", "#B54708"),
    ("P2 一年内启动", "#FFFCE6", "#A16207"),
    ("P3 一年内以上", "#E8EEF7", "#1F3864"),
]
ALERT_SHEET = "到期预警"
SLOT_SHEET = "铺位续约清单"
MAX_ROWS = 500


def run(cmd, stdin_data=None):
    env = dict(os.environ)
    env["PATH"] = os.path.dirname(LARK_BIN) + os.pathsep + env.get("PATH", "")
    p = subprocess.run(cmd, env=env, input=stdin_data, capture_output=True, text=True)
    try:
        return json.loads(p.stdout.strip() or "{}")
    except json.JSONDecodeError:
        return {"raw": p.stdout, "stderr": p.stderr[-1200:]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--表格token", required=True)
    args = ap.parse_args()
    url = f"https://my.feishu.cn/sheets/{args.表格token}"

    # 已有规则先清掉（条件格式规则对象不能靠清单元格清掉，需显式删）
    for sheet in (ALERT_SHEET, SLOT_SHEET):
        lst = run([LARK_BIN, "sheets", "+cond-format-list", "--as", "user",
                   "--url", url, "--sheet-name", sheet])
        rules = (lst.get("data") or {}).get("rules") or (lst.get("data") or {}).get("conditional_formats") or []
        for r in rules:
            rid = r.get("rule_id") or r.get("id")
            if rid:
                run([LARK_BIN, "sheets", "+cond-format-delete", "--as", "user", "--url", url,
                     "--sheet-name", sheet, "--rule-id", rid, "--yes"])

    rng = json.dumps([f"A2:P{MAX_ROWS}"], ensure_ascii=False)
    for text, back, fore in RULES:
        props = json.dumps({
            "style": {"back_color": back, "fore_color": fore},
            "attrs": [{"formula": [f'=$A2="{text}"']}],
        }, ensure_ascii=False)
        res = run([LARK_BIN, "sheets", "+cond-format-create", "--as", "user", "--url", url,
                   "--sheet-name", ALERT_SHEET, "--rule-type", "expression",
                   "--ranges", rng, "--properties", props])
        print(f"[条件格式] {ALERT_SHEET} {text}: {'OK' if res.get('ok') else res}")

    # 铺位续约清单：按优先级上色（O 列 = 续约优先级）
    slot_rng = json.dumps([f"A2:AA{MAX_ROWS}"], ensure_ascii=False)
    for text, back, fore in PRIORITY_RULES:
        props = json.dumps({
            "style": {"back_color": back, "fore_color": fore},
            "attrs": [{"formula": [f'=$O2="{text}"']}],
        }, ensure_ascii=False)
        res = run([LARK_BIN, "sheets", "+cond-format-create", "--as", "user", "--url", url,
                   "--sheet-name", SLOT_SHEET, "--rule-type", "expression",
                   "--ranges", slot_rng, "--properties", props])
        print(f"[条件格式] {SLOT_SHEET} {text}: {'OK' if res.get('ok') else res}")

    # 子表顺序：看板 → 明细 → 合同级 → 到期预警 → 铺位续约 → 区域业态 → 品牌 → 字段说明 → 版本记录
    order = ["台账看板", "合同台账明细", "合同级汇总", ALERT_SHEET, SLOT_SHEET,
             "区域业态汇总", "品牌汇总", "字段说明", "版本记录"]
    info = run([LARK_BIN, "sheets", "+workbook-info", "--as", "user", "--url", url])
    names = [s.get("sheet_name") for s in (info.get("data") or {}).get("sheets", [])]
    for idx, want in enumerate(order):
        if want not in names:
            continue
        cur = names.index(want)
        if cur != idx:
            run([LARK_BIN, "sheets", "+sheet-move", "--as", "user", "--url", url,
                 "--sheet-name", want, "--index", str(idx)])
            names.remove(want)
            names.insert(idx, want)
    print("[顺序] 已对齐:", " → ".join(names))


if __name__ == "__main__":
    main()
