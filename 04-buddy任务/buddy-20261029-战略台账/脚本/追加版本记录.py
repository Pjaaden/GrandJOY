#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
向飞书台账的「版本记录」子表追加一条本次更新的版本记录（增量，不覆盖历史）。

用法：
    python 追加版本记录.py --表格token <token> --源文件 <BI导出xlsx>
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
VERSION_HEADERS = [
    "数据版本", "统计日期", "源文件名", "明细行数", "合同数", "项目数", "区域数",
    "生成时间", "更新人", "备注",
]


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
    ap.add_argument("--源文件", required=True)
    ap.add_argument("--备注", default="定期更新")
    args = ap.parse_args()

    data = build_all(args.源文件)
    v = dict(data["版本记录"][0])
    v["备注"] = args.备注
    v["生成时间"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    url = f"https://my.feishu.cn/sheets/{args.表格token}"

    # 读现有版本记录，算出末行
    res = run([LARK_BIN, "sheets", "+csv-get", "--as", "user", "--url", url,
               "--sheet-name", "版本记录", "--range", "A1:J300"])
    csv_text = (res.get("data", {}) or {}).get("annotated_csv", "")
    last_row = 0
    for line in csv_text.splitlines():
        if line.startswith("[row="):
            try:
                last_row = max(last_row, int(line[5:line.index("]")]))
            except ValueError:
                pass

    # 若已有同数据版本的记录，则跳过（幂等：同一版 BI 重复跑不重复记）
    if any(f"数据版本 {v['数据版本']}" in ln or f",{v['数据版本']}," in ln for ln in csv_text.splitlines()):
        print(f"[跳过] 数据版本 {v['数据版本']} 已有记录，不重复追加")
        return

    target = last_row + 1
    cells = [[{"value": v.get(h)} for h in VERSION_HEADERS]]
    payload = [{"sheet_name": "版本记录", "range": f"A{target}:J{target}", "cells": cells}]
    out = run([LARK_BIN, "sheets", "+cells-set", "--as", "user", "--url", url, "--writes", "-"],
              stdin_data=json.dumps(payload, ensure_ascii=False))
    print(f"[OK] 版本记录已追加至第 {target} 行:", json.dumps(out, ensure_ascii=False)[:200])


if __name__ == "__main__":
    main()
