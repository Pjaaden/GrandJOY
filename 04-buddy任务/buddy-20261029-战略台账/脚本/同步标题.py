#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把飞书表的标题同步为「名称_数据版本版_更新YYYYMMDD」。

为什么需要：
  表格标题里带了数据版本和更新日期，人在飞书文件夹里一眼就能分清哪张是当前版、哪张是旧版。
  但版本号会随每次 BI 导出变化，写死在标题里就会过期 —— 所以每次更新完自动同步一次。

用法：
    python 同步标题.py --表格token <token> --源文件 <BI导出xlsx> [--名称 "前缀名"]
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
    # ⚠ 飞书 CLI 出错时把 JSON 错误体写到 **stderr**、stdout 为空。
    #    只解析 stdout 会拿到 {}，看不出真实错误码 —— 两边都要试。
    for blob in (p.stdout, p.stderr):
        blob = (blob or "").strip()
        if blob.startswith("{"):
            try:
                return json.loads(blob)
            except json.JSONDecodeError:
                continue
    return {"ok": False, "raw": (p.stdout or "")[:200], "stderr": (p.stderr or "")[-400:]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--表格token", required=True)
    ap.add_argument("--源文件", required=True, action="append",
                    help="BI 导出的合同台账 xlsx；可多次传参或用逗号分隔（多源合并）")
    ap.add_argument("--名称", default="全国运动男装战略品牌合同台账", help="标题前缀")
    ap.add_argument("--类型", default="sheet", help="sheet|docx|base")
    args = ap.parse_args()

    from 台账核心 import split_source_args
    data = build_all(split_source_args(args.源文件))
    version = str(data["版本记录"][0]["数据版本"])
    today = dt.date.today().strftime("%Y%m%d")
    title = f"{args.名称}_{version}版_更新{today}"

    res = run([LARK_BIN, "drive", "+update-title", "--as", "user",
               "--type", args.类型, "--token", args.表格token, "--title", title])
    if res.get("ok"):
        print(f"[标题] 已同步 → {title}")
        return
    # 飞书在「标题与当前完全相同」时返回 981001 internal error —— 属于无变化的正常情况
    code = (res.get("error") or {}).get("code")
    if code == 981001:
        print(f"[标题] 无需变更（标题已是）→ {title}")
        return
    print(f"[标题] 同步失败（不影响数据）: {str(res)[:200]}")


if __name__ == "__main__":
    main()
