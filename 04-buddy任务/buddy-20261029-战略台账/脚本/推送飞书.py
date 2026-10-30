#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把台账推送到飞书电子表格（我的空间 / 大悦商业 / 战略台账）。

用法：
    python 推送飞书.py --源文件 <BI导出的xlsx> [--表格token <已有表token>] [--仅导出载荷]

行为：
  - 未给 --表格token：新建一张电子表格「全国运动男装战略品牌合同台账」
  - 给了 --表格token：覆盖式刷新该表的全部子表（mode=overwrite），用于定期更新
  - 载荷用 typed 协议（+table-put --sheets），日期落真日期、金额/扣率保数值型
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 台账核心 import SHEET_SPECS, build_all, load_source  # noqa: E402

LARK_BIN = os.path.expanduser(
    "~/.grandjoybuddy/binaries/node/workspace/node_modules/.bin/lark-cli"
)
FOLDER_TOKEN = "ZRbRfWrV7lMEKgdYWPocyboBngf"  # 大悦商业 / 战略台账
TABLE_TITLE = "全国运动男装战略品牌合同台账"

# 日期列：写真日期而非文本
DATE_COLS = {"起租日", "到期日", "合同起租日", "合同到期日"}
# 版本记录单独维护（见 更新版本记录.py），不参与覆盖刷新，避免历史被清掉
TYPED_SHEETS = [s for s in SHEET_SPECS if s != "版本记录"]
# 各子表清理范围（覆盖刷新前先清，防止行数变少时残留旧行）
CLEAR_RANGES = {
    "合同台账明细": "A1:AZ3000",
    "合同级汇总": "A1:AZ1000",
    "到期预警": "A1:AZ500",
    "铺位续约清单": "A1:AZ500",
    "区域业态汇总": "A1:AZ500",
    "品牌汇总": "A1:AZ500",
    "字段说明": "A1:AZ200",
}


def to_typed(name: str, rows: list[dict]) -> dict:
    spec = SHEET_SPECS[name]
    headers = spec["headers"]
    dtypes = dict(spec.get("dtypes", {}))
    formats = dict(spec.get("formats", {}))
    for h in headers:
        if h in DATE_COLS:
            dtypes[h] = "datetime64[ns]"
    data = []
    for r in rows:
        line = []
        for h in headers:
            v = r.get(h)
            if v is None:
                line.append(None)
            elif h in DATE_COLS:
                line.append(str(v))
            elif isinstance(v, bool):
                line.append(v)
            elif isinstance(v, (int, float)):
                line.append(v)
            else:
                line.append(str(v))
        data.append(line)
    item = {
        "name": name,
        "columns": headers,
        "data": data,
        "header": True,
        "mode": "overwrite",
    }
    if dtypes:
        item["dtypes"] = dtypes
    if formats:
        item["formats"] = formats
    return item


def build_styles(name: str, spec: dict, n_rows: int) -> dict:
    """为每个子表生成配套样式：表头、冻结、列宽、每列数字格式。

    注意：必须为每个数据列显式声明 number_format。
    原因：清空只清内容不清格式，旧列残留的日期格式会让新写入的整数
    显示成 1900-01-01（列数变化时尤其明显）。
    """
    headers = spec["headers"]
    dtypes = spec.get("dtypes", {})
    formats = spec.get("formats", {})

    def col_letter(n: int) -> str:
        s = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            s = chr(65 + rem) + s
        return s

    end = col_letter(len(headers))
    last_row = max(n_rows + 1, 2)
    styles = {
        "name": name,
        "cell_styles": [
            {
                "range": f"A1:{end}1",
                "font_weight": "bold",
                "background_color": "#1F3864",
                "font_color": "#FFFFFF",
                "horizontal_alignment": "center",
                "vertical_alignment": "middle",
                "border_styles": {
                    "bottom": {"style": "solid", "weight": "thin", "color": "#D6DCE4"}
                },
            }
        ],
        "row_sizes": [{"range": "1:1", "type": "pixel", "size": 30}],
        "col_sizes": [],
        "freeze": {"rows": 1, "cols": 0},
    }

    # 每列显式数字格式 + 对齐
    for i, h in enumerate(headers, start=1):
        cl = col_letter(i)
        dt_name = dtypes.get(h, "object")
        cell_style = {}
        if h in DATE_COLS or dt_name.startswith("datetime"):
            cell_style["number_format"] = formats.get(h, "yyyy-mm-dd")
            cell_style["horizontal_alignment"] = "center"
        elif dt_name.startswith("int"):
            cell_style["number_format"] = formats.get(h, "0")
            cell_style["horizontal_alignment"] = "right"
        elif dt_name.startswith("float"):
            cell_style["number_format"] = formats.get(h, "#,##0.00")
            cell_style["horizontal_alignment"] = "right"
        else:
            cell_style["number_format"] = formats.get(h, "@")
            cell_style["horizontal_alignment"] = "left"
        styles["cell_styles"].append({"range": f"{cl}2:{cl}{last_row}", **cell_style})

    # 列宽
    for i, h in enumerate(headers, start=1):
        width = 120
        if h in ("合同名称", "商户名称", "备注", "口径与计算规则", "源表来源列", "字段名称", "项目名称"):
            width = 240
        elif h in ("数据校验", "涉及三级业态", "所属区域", "品牌/行业", "到期预警", "预警分级"):
            width = 170
        elif "(元)" in h:
            width = 160
        styles["col_sizes"].append({"range": f"{col_letter(i)}:{col_letter(i)}", "type": "pixel", "size": width})
    return styles


def run(cmd: list[str], cwd: str | None = None, stdin_data: str | None = None) -> dict:
    env = dict(os.environ)
    env["PATH"] = os.path.dirname(LARK_BIN) + os.pathsep + env.get("PATH", "")
    proc = subprocess.run(
        cmd,
        cwd=cwd,
        env=env,
        input=stdin_data,
        capture_output=True,
        text=True,
    )
    out = proc.stdout.strip()
    try:
        parsed = json.loads(out) if out else {}
    except json.JSONDecodeError:
        parsed = {"raw": out}
    if not parsed.get("ok", False):
        parsed["_stderr"] = proc.stderr[-2000:]
        parsed["_returncode"] = proc.returncode
    return parsed


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--源文件", required=True)
    ap.add_argument("--表格token", default=None, help="已有飞书表 token（刷新用）")
    ap.add_argument("--仅导出载荷", action="store_true")
    ap.add_argument("--载荷目录", default=None)
    args = ap.parse_args()

    data = build_all(args.源文件)
    _, meta = load_source(args.源文件)

    sheets_payload = {"sheets": [to_typed(n, data[n]) for n in TYPED_SHEETS]}
    styles_payload = {"styles": [build_styles(n, SHEET_SPECS[n], len(data[n])) for n in TYPED_SHEETS]}

    out_dir = args.载荷目录 or os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "产出", "_载荷")
    os.makedirs(out_dir, exist_ok=True)
    sp = os.path.join(out_dir, "sheets.json")
    st = os.path.join(out_dir, "styles.json")
    with open(sp, "w", encoding="utf-8") as f:
        json.dump(sheets_payload, f, ensure_ascii=False)
    with open(st, "w", encoding="utf-8") as f:
        json.dump(styles_payload, f, ensure_ascii=False)
    print(f"[载荷] {sp}  ({os.path.getsize(sp)/1024:.0f} KB)")
    print(f"[载荷] {st}  ({os.path.getsize(st)/1024:.0f} KB)")

    if args.仅导出载荷:
        return

    if not args.表格token:
        # 新建工作簿：--sheets 走 stdin，--styles 走 @相对路径
        cmd = [
            LARK_BIN, "sheets", "+workbook-create", "--as", "user",
            "--title", TABLE_TITLE,
            "--folder-token", FOLDER_TOKEN,
            "--sheets", "-",
            "--styles", "@styles.json",
        ]
        res = run(cmd, cwd=out_dir, stdin_data=open(sp, encoding="utf-8").read())
        print(json.dumps(res, ensure_ascii=False, indent=2)[:3000])
    else:
        # 覆盖刷新：先清空各子表旧数据，防行数变少时残留
        # 注意：+cells-batch-clear 遇到不存在的子表会整体失败，
        # 所以先查工作簿结构，只清真实存在的子表。
        info = run([LARK_BIN, "sheets", "+workbook-info", "--as", "user",
                    "--url", f"https://my.feishu.cn/sheets/{args.表格token}"])
        exists = {s.get("sheet_name") for s in (info.get("data") or {}).get("sheets", [])}
        ranges = [f"{n}!{r}" for n, r in CLEAR_RANGES.items() if n in exists]
        if ranges:
            clr = run([LARK_BIN, "sheets", "+cells-batch-clear", "--as", "user",
                       "--url", f"https://my.feishu.cn/sheets/{args.表格token}",
                       "--ranges", json.dumps(ranges, ensure_ascii=False), "--yes"])
            print("[清空旧数据]", "OK" if clr.get("ok") else json.dumps(clr, ensure_ascii=False)[:200])

        cmd = [
            LARK_BIN, "sheets", "+table-put", "--as", "user",
            "--spreadsheet-token", args.表格token,
            "--sheets", "-",
            "--styles", "@styles.json",
        ]
        res = run(cmd, cwd=out_dir, stdin_data=open(sp, encoding="utf-8").read())
        print(json.dumps(res, ensure_ascii=False, indent=2)[:3000])


if __name__ == "__main__":
    main()
