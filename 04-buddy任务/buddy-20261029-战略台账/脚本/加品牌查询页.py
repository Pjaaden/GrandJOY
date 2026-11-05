#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
构建飞书台账的「品牌合同查询」子表：选定一个或多个品牌 → 表格端实时得到当前有效合同台账。

用法：
    python 加品牌查询页.py --表格token <token> [--源文件 <BI导出的xlsx>]

═══ 设计说明（2026-11-03 沙箱实测定稿）═══

【为什么做成两张表】
第一版把 937 行明细留在「合同台账明细」子表、查询页全部跨表引用，一次重算要触碰约
**1057 万**单元格（MAXIFS 562,200 + INDEX 5,622,000 + SUMPRODUCT 3,511,876 + 跨表 SEARCH 877,969），
直接把整张台账拖垮 —— workbook-info / csv-get 全部 40400 record not found，连删子表都失败。

本版改为：
  · 「查询数据」子表 —— 用 typed 协议写入的**静态快照**（日期是真日期、数字是真数字），零重算成本
  · 「品牌合同查询」子表 —— 用户界面（品牌多选下拉 + 汇总条 + 结果区），公式只做轻量引用
  重算量级约 **6 万**，是上一版的 1/170。

【踩过的坑（都已在沙箱复现并修正）】
1. **+cells-set 会把数字和日期写成文本** → 文本日期和 TODAY() 比较是按字符比的（结果全错），
   文本数字也无法参与加权求和。快照必须用 **+table-put 的 typed 协议**
   （dtype=datetime64[ns] / float64）写，落成真日期真数字。
2. **行号错位**：body 从 A1 写入，body[i] 对应表格第 i+1 行。写死 FIRST_ROW 常量会让
   结果区公式整体偏移一行（表现为「品牌对、其他列全是 0」）。
   现在所有行号都由 body 实际长度推导。
3. **子表列数不够时写入静默失败**：+cells-set 写超界范围不报错、也不写。
   写快照前先 +dim-insert 扩列。
4. **+dim-insert 的 --count 必须是字符串**（传 int 会 TypeError）。
5. **跨表只能引用单格，不能用区域数组**：
   - INDEX(跨表区域, 标量行号)   ✅ 可用
   - SMALL(跨表区域, k)          ❌ #REF! —— SMALL 必须在同表
   - AGGREGATE(跨表区域, ...)    ❌ #REF!
   所以「命中序号」辅助列放在查询页同表，只跨表引用快照的**单个单元格**做判定。
6. **多选下拉**取值是逗号分隔字符串（如 "NIKE,李宁"），
   匹配必须逗号包裹 SEARCH(",李宁,", ","&选区&",")，否则会误命中「李宁YOUNG」。

【飞书公式能力实测结论】
✓ INDEX、SMALL、COUNTIF、COUNTA、SUMIF、SUMPRODUCT、MAXIFS、MINIFS、
  TEXTJOIN、SEARCH、ISNUMBER、IFERROR、AND、OR、TODAY
✗ FILTER 动态数组溢出 → #REF!
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 台账核心 import build_all, load_source  # noqa: E402

LARK_BIN = os.path.expanduser(
    "~/.grandjoybuddy/binaries/node/workspace/node_modules/.bin/lark-cli"
)

QUERY_SHEET = "品牌合同查询"     # 用户界面页
DATA_SHEET = "查询数据"          # 静态快照页
N_RESULT = 60                    # 结果区行数（单品牌在租最多 22 行）

# 快照列（顺序即 typed 写入的列序）
SNAP_H = [
    "品牌/行业", "期次起租日", "期次到期日", "合同最终到期日", "合同编号",
    "项目名称", "楼层", "资源位编号", "商户名称", "合同状态",
    "签约面积(㎡)", "单位日租金(元/㎡/日)", "抽成扣率",
    "物业费(元/㎡/月)", "推广费(元/㎡/月)", "所属区域", "三级业态",
    "是否战略品牌", "数据校验",
]
SNAP_COL = {h: chr(ord("A") + i) for i, h in enumerate(SNAP_H)}   # 查询数据页列字母
SNAP_DTYPE = {
    "期次起租日": "datetime64[ns]", "期次到期日": "datetime64[ns]",
    "合同最终到期日": "datetime64[ns]", "签约面积(㎡)": "float64",
    "单位日租金(元/㎡/日)": "float64", "抽成扣率": "float64",
    "物业费(元/㎡/月)": "float64", "推广费(元/㎡/月)": "float64",
}
SNAP_FORMAT = {
    "期次起租日": "yyyy-mm-dd", "期次到期日": "yyyy-mm-dd",
    "合同最终到期日": "yyyy-mm-dd", "签约面积(㎡)": "#,##0.00",
    "单位日租金(元/㎡/日)": "#,##0.0000", "抽成扣率": "0.00%",
    "物业费(元/㎡/月)": "#,##0.0000", "推广费(元/㎡/月)": "#,##0.0000",
}

# 结果区列：(表头, 快照列名, 数字格式)  —— "__DAYS__" 为派生列
RESULT_COLS = [
    ("品牌/行业", "品牌/行业", None),
    ("项目名称", "项目名称", None),
    ("楼层", "楼层", None),
    ("资源位编号", "资源位编号", None),
    ("商户名称", "商户名称", None),
    ("合同编号", "合同编号", None),
    ("合同状态", "合同状态", None),
    ("签约面积(㎡)", "签约面积(㎡)", "#,##0.00"),
    ("单位日租金(元/㎡/日)", "单位日租金(元/㎡/日)", "#,##0.0000"),
    ("抽成扣率", "抽成扣率", "0.00%"),
    ("物业费(元/㎡/月)", "物业费(元/㎡/月)", "#,##0.0000"),
    ("推广费(元/㎡/月)", "推广费(元/㎡/月)", "#,##0.0000"),
    ("期次起租日", "期次起租日", "yyyy-mm-dd"),
    ("期次到期日", "期次到期日", "yyyy-mm-dd"),
    ("合同最终到期日", "合同最终到期日", "yyyy-mm-dd"),
    ("距最终到期天数", "__DAYS__", "0"),
    ("所属区域", "所属区域", None),
    ("三级业态", "三级业态", None),
    ("是否战略品牌", "是否战略品牌", None),
]

HELPER_COL = "T"    # 查询页：命中序号辅助列
PTR_COL = "V"      # 查询页：指针列
BRAND_COL = "W"    # 查询页：品牌清单（下拉数据源）


def run(cmd, stdin_data=None, retries=4, cwd=None):
    env = dict(os.environ)
    env["PATH"] = os.path.dirname(LARK_BIN) + os.pathsep + env.get("PATH", "")
    last = {}
    for attempt in range(retries):
        p = subprocess.run(cmd, cwd=cwd, env=env, input=stdin_data,
                           capture_output=True, text=True)
        # ⚠ 飞书 CLI 出错时把 JSON 错误体写到 **stderr**、stdout 为空。
        #    只解析 stdout 会拿到 {}，看不出真实错误码（40400 重试逻辑会因此失效）—— 两边都要试。
        last = {}
        for blob in (p.stdout, p.stderr):
            blob = (blob or "").strip()
            if blob.startswith("{"):
                try:
                    last = json.loads(blob)
                    break
                except json.JSONDecodeError:
                    continue
        if not last:
            last = {"ok": False, "raw": (p.stdout or "")[:300], "stderr": (p.stderr or "")[-800:]}
        if last.get("ok"):
            return last
        msg = json.dumps(last, ensure_ascii=False)
        if "40400" in msg or "record not found" in msg:
            time.sleep(3 + attempt * 3)
            continue
        return last
    return last


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and v.strip():
        try:
            return date.fromisoformat(v.strip()[:10])
        except ValueError:
            return None
    return None


def ensure_sheet(url, name, index=None, rows=30, cols=24):
    """确保子表存在，返回 sheet_id。"""
    info = run([LARK_BIN, "sheets", "+workbook-info", "--as", "user", "--url", url])
    for s in (info.get("data") or {}).get("sheets", []) or []:
        if s.get("sheet_name") == name:
            return s.get("sheet_id"), s
    cmd = [LARK_BIN, "sheets", "+sheet-create", "--as", "user", "--url", url,
           "--title", name, "--row-count", str(rows), "--col-count", str(cols)]
    if index is not None:
        cmd += ["--index", str(index)]
    res = run(cmd)
    if not res.get("ok"):
        print(f"[ERR] 建子表失败 {name}:", json.dumps(res, ensure_ascii=False)[:400])
        return None, None
    time.sleep(3)
    sid = (res.get("data") or {}).get("sheet_id")
    print(f"[OK] 子表已建 {name}, sheet_id = {sid}")
    return sid, {"sheet_name": name, "sheet_id": sid, "column_count": cols,
                 "row_count": rows}


def find_or_create_query_workbook(folder_token):
    """品牌查询页**单独建一张表**，与主台账物理隔离。

    为什么必须独立：公式页一旦出问题是会波及整张表的 ——
    2026-11-03 就因为把查询页放进主台账，重算风暴把主台账锁死 40+ 分钟，
    连删子表都失败。独立表即使公式再重，也只影响自己。
    """
    name = "品牌合同查询工具"
    # 找同文件夹下已有的同名表
    lst = run([LARK_BIN, "drive", "+list", "--as", "user",
                "--folder-token", folder_token])
    files = ((lst.get("data") or {}).get("files")
             or (lst.get("data") or {}).get("items")
             or (lst.get("data") or {}).get("list")
             or [])
    for f in files:
        if f.get("name") == name and f.get("type") == "sheet":
            tok = f.get("token")
            if tok:
                print(f"[OK] 复用已有查询工具表, token = {tok}")
                return tok
    # 没有就新建一张空表
    res = run([LARK_BIN, "sheets", "+workbook-create", "--as", "user",
               "--title", name, "--folder-token", folder_token])
    if not res.get("ok"):
        print("[ERR] 建查询工具表失败:", json.dumps(res, ensure_ascii=False)[:400])
        return None
    sp = (res.get("data") or {}).get("spreadsheet") or {}
    tok = sp.get("spreadsheet_token")
    print(f"[OK] 查询工具表已建, token = {tok}")
    print(f"     {sp.get('url','')}")
    return tok


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--表格token", required=True,
                    help="主台账 token（查询页默认另建独立表，此参数仅用于定位源数据）")
    ap.add_argument("--查询表token", default=None,
                    help="查询工具表 token；不给则新建一张独立表")
    ap.add_argument("--文件夹token", default="ZRbRfWrV7lMEKgdYWPocyboBngf",
                    help="查询工具表存放文件夹（战略台账）")
    ap.add_argument("--源文件", default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "源数据", "运动男装合同台账1103.xlsx"))
    args = ap.parse_args()

    # ── 定位查询工具表（独立于主台账，公式页出问题也不会波及台账）──
    qtok = args.查询表token or find_or_create_query_workbook(args.文件夹token)
    if not qtok:
        print("[ERR] 没有可用的查询工具表，终止")
        return
    url = f"https://my.feishu.cn/sheets/{qtok}"
    print(f"[目标] 查询工具表 {qtok}")
    print(f"       {url}")

    from 台账核心 import load_sources, split_source_args
    srcs = split_source_args(args.源文件)
    data = build_all(srcs)
    _, meta = load_sources(srcs)
    detail = data["合同台账明细"]

    # 合同最终到期日 = 同合同编号下期次到期日最大值（Python 预算 → 静态值）
    contract_end = {}
    for d in detail:
        cid = d["合同编号"]
        e = to_date(d["到期日"])
        if cid and e and (cid not in contract_end or e > contract_end[cid]):
            contract_end[cid] = e

    snap = []
    for d in detail:
        snap.append({
            "品牌/行业": d["品牌/行业"],
            "期次起租日": to_date(d["起租日"]),
            "期次到期日": to_date(d["到期日"]),
            "合同最终到期日": contract_end.get(d["合同编号"]),
            "合同编号": d["合同编号"],
            "项目名称": d["项目名称"],
            "楼层": d["楼层"],
            "资源位编号": d["资源位编号"],
            "商户名称": d["商户名称"],
            "合同状态": d["合同状态"],
            "签约面积(㎡)": d["签约面积"],
            "单位日租金(元/㎡/日)": d["单位日租金(元/平方米/日)"],
            "抽成扣率": d["抽成扣率"],
            "物业费(元/㎡/月)": d["物业费(元/平方米/月)"],
            "推广费(元/㎡/月)": d["推广费(元/平方米/月)"],
            "所属区域": d["所属区域"],
            "三级业态": d["三级业态"],
            "是否战略品牌": d["是否战略品牌"],
            "数据校验": d["数据校验"],
        })
    n_snap = len(snap)
    last = n_snap + 1                       # 快照最后一行（含表头）
    brands = sorted({s["品牌/行业"] for s in snap if s["品牌/行业"]})

    payload_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "产出", "_载荷")
    os.makedirs(payload_dir, exist_ok=True)

    # ── 1) 「查询数据」页：typed 静态快照 ──────────────────────────────
    #    新建的表自带一个空白 Sheet1，删掉避免困惑
    for s in ((run([LARK_BIN, "sheets", "+workbook-info", "--as", "user", "--url", url])
               .get("data") or {}).get("sheets") or []):
        if s.get("sheet_name") in ("Sheet1", "Sheet 1"):
            run([LARK_BIN, "sheets", "+sheet-delete", "--as", "user", "--url", url,
                 "--sheet-name", s["sheet_name"], "--yes"])
            time.sleep(2)

    _, dinfo = ensure_sheet(url, DATA_SHEET, index=1, rows=20, cols=len(SNAP_H))
    if dinfo is None:
        return
    need_cols = len(SNAP_H)
    if dinfo.get("column_count", 0) < need_cols:
        res = run([LARK_BIN, "sheets", "+dim-insert", "--as", "user", "--url", url,
                   "--sheet-name", DATA_SHEET,
                   "--position", col_letter(dinfo.get("column_count", 20) + 1),
                   "--count", str(need_cols - dinfo.get("column_count", 20))])
        print("[扩列]", "OK" if res.get("ok") else json.dumps(res, ensure_ascii=False)[:150])
        time.sleep(2)

    snap_data = []
    for s in snap:
        row = []
        for h in SNAP_H:
            v = s[h]
            if v is None:
                row.append(None)
            elif isinstance(v, date):
                row.append(v.isoformat())
            else:
                row.append(v)
        snap_data.append(row)
    snap_payload = {"sheets": [{
        "name": DATA_SHEET, "columns": SNAP_H, "data": snap_data,
        "header": True, "mode": "overwrite",
        "dtypes": SNAP_DTYPE, "formats": SNAP_FORMAT,
    }]}
    sp = os.path.join(payload_dir, "query_snapshot.json")
    with open(sp, "w", encoding="utf-8") as f:
        json.dump(snap_payload, f, ensure_ascii=False)
    t = time.time()
    res = run([LARK_BIN, "sheets", "+table-put", "--as", "user", "--url", url,
               "--sheets", f"@{sp}"], cwd=payload_dir)
    print(f"[写入快照 {n_snap} 行]",
          "OK" if res.get("ok") else json.dumps(res, ensure_ascii=False)[:250],
          f"({time.time() - t:.1f}s)")

    # ── 2) 「品牌合同查询」页：用户界面 ────────────────────────────────
    _, qinfo = ensure_sheet(url, QUERY_SHEET, index=1, rows=30, cols=22)
    if qinfo is None:
        return

    # 清空（先删筛选器：残留筛选会让被筛掉的行清不掉）
    run([LARK_BIN, "sheets", "+filter-delete", "--as", "user", "--url", url,
         "--sheet-name", QUERY_SHEET, "--yes"])
    clr = run([LARK_BIN, "sheets", "+cells-batch-clear", "--as", "user", "--url", url,
               "--ranges", json.dumps([f"{QUERY_SHEET}!A1:BZ2000"], ensure_ascii=False),
               "--yes"])
    print("[OK] 查询页已清空:", clr.get("ok"))

    body: list[list[dict]] = []

    def title(text, size=16, color="#1F3864"):
        return {"value": text, "cell_styles": {
            "font_size": size, "font_weight": "bold", "font_color": color,
            "horizontal_alignment": "left", "vertical_alignment": "middle"}}

    def label(text):
        return {"value": text, "cell_styles": {
            "font_size": 10, "font_color": "#666666",
            "horizontal_alignment": "left", "vertical_alignment": "middle"}}

    body.append([title("品牌合同查询 · 当前有效合同台账")])
    body.append([label(f"数据版本 {meta.get('stat_date')}　|　明细 {n_snap} 行　|　"
                       "在租判定：起租日 ≤ 今天 ≤ 到期日（随日期自动滚动）")])
    body.append([label("① 选择品牌（可多选）"), {"value": "", "cell_styles": {
        "background_color": "#FFF7E6", "font_weight": "bold",
        "horizontal_alignment": "left"}}, {}, {}, {}, {}])
    body.append([label("② 汇总"), title("合同数", 11),
                 {"value": 0, "cell_styles": {"font_size": 13, "font_weight": "bold",
                                              "font_color": "#1F3864",
                                              "number_format": "#,##0",
                                              "horizontal_alignment": "left"}},
                 title("签约面积(㎡)", 11),
                 {"value": 0, "cell_styles": {"font_size": 13, "font_weight": "bold",
                                              "font_color": "#1F3864",
                                              "number_format": "#,##0.00",
                                              "horizontal_alignment": "left"}},
                 title("加权日租(元/㎡/日)", 11),
                 {"value": 0, "cell_styles": {"font_size": 13, "font_weight": "bold",
                                              "font_color": "#1F3864",
                                              "number_format": "#,##0.00",
                                              "horizontal_alignment": "left"}},
                 title("最早到期", 11),
                 {"value": "", "cell_styles": {"font_size": 13, "font_weight": "bold",
                                               "font_color": "#B42318",
                                               "number_format": "yyyy-mm-dd",
                                               "horizontal_alignment": "left"}},
                 {}, {}, {}, {}, {}])
    body.append([{"value": h, "cell_styles": {
        "font_weight": "bold", "background_color": "#1F3864", "font_color": "#FFFFFF",
        "horizontal_alignment": "center", "vertical_alignment": "middle",
        "border_styles": {"bottom": {"style": "solid", "weight": "thin", "color": "#D6DCE4"}}}}
        for h, _, _ in RESULT_COLS])

    # ⚠ 行号一律由 body 长度推导：body 从 A1 写入，body[i] → 表格第 i+1 行
    HDR_ROW = len(body)
    FIRST_ROW = HDR_ROW + 1
    SEL_ROW = len(body) - 3
    STAT_ROW = len(body) - 2
    SEL_CELL = f"$B${SEL_ROW}"
    print(f"[行号] 下拉 B{SEL_ROW} / 统计条 第{STAT_ROW}行 / 表头 第{HDR_ROW}行 / "
          f"结果区 第{FIRST_ROW}~{FIRST_ROW + N_RESULT - 1}行")

    days_col = col_letter(
        [h for h, _, _ in RESULT_COLS].index("距最终到期天数") + 1)
    end_col = col_letter(
        [h for h, _, _ in RESULT_COLS].index("合同最终到期日") + 1)
    for k in range(1, N_RESULT + 1):
        r = FIRST_ROW + k - 1
        vals = []
        for _, src, fmt in RESULT_COLS:
            if src == "__DAYS__":
                f = f'=IFERROR(IF(${PTR_COL}{r}="","",${end_col}{r}-TODAY()),"")'
            else:
                cl = SNAP_COL[src]
                # ⚠ 指针为空时必须显式返回空串：INDEX(区域,"") 在飞书里会返回首行内容，
                #   导致结果区出现一串重复行（表现为命中数 22 但显示 65 行）
                f = (f'=IFERROR(IF(${PTR_COL}{r}="","",'
                     f'INDEX(\'{DATA_SHEET}\'!${cl}$2:${cl}${last},${PTR_COL}{r})),"")')
            st = {"horizontal_alignment": "left"}
            if fmt:
                st["number_format"] = fmt
                st["horizontal_alignment"] = "right" if fmt != "yyyy-mm-dd" else "center"
            vals.append({"formula": f, "cell_styles": st})
        body.append(vals)

    bp = os.path.join(payload_dir, "query_body.json")
    with open(bp, "w", encoding="utf-8") as f:
        json.dump(body, f, ensure_ascii=False)
    res = run([LARK_BIN, "sheets", "+cells-set", "--as", "user", "--url", url,
               "--sheet-name", QUERY_SHEET, "--range", "A1", "--cells", f"@{bp}"],
              cwd=payload_dir)
    print(f"[写入主体] {'OK' if res.get('ok') else json.dumps(res, ensure_ascii=False)[:250]}")

    # ── 3) 辅助列：品牌命中 + 在租 → 返回相对行号（同表，跨表只引单格）──
    helper = [[{"formula": (
        f'=IF(AND(ISNUMBER(SEARCH(","&\'{DATA_SHEET}\'!$A{i}&",",","&{SEL_CELL}&",")),'
        f"'{DATA_SHEET}'!$B{i}<=TODAY(),'{DATA_SHEET}'!$C{i}>=TODAY()),{i - 1},\"\")"
    )}] for i in range(2, last + 1)]
    hp = os.path.join(payload_dir, "query_helper.json")
    with open(hp, "w", encoding="utf-8") as f:
        json.dump(helper, f, ensure_ascii=False)
    res = run([LARK_BIN, "sheets", "+cells-set", "--as", "user", "--url", url,
               "--sheet-name", QUERY_SHEET,
               "--range", f"{HELPER_COL}2:{HELPER_COL}{last}",
               "--cells", f"@{hp}"], cwd=payload_dir)
    print("[写入辅助列]", "OK" if res.get("ok") else json.dumps(res, ensure_ascii=False)[:250])

    # ── 4) 指针列：SMALL 取第 k 个命中（必须同表）────────────────────
    ptr = [[{"formula": (
        f'=IFERROR(SMALL(${HELPER_COL}$2:${HELPER_COL}${last},{k}),"")')}]
        for k in range(1, N_RESULT + 1)]
    pp = os.path.join(payload_dir, "query_pointer.json")
    with open(pp, "w", encoding="utf-8") as f:
        json.dump(ptr, f, ensure_ascii=False)
    res = run([LARK_BIN, "sheets", "+cells-set", "--as", "user", "--url", url,
               "--sheet-name", QUERY_SHEET,
               "--range", f"{PTR_COL}{FIRST_ROW}:{PTR_COL}{FIRST_ROW + N_RESULT - 1}",
               "--cells", f"@{pp}"], cwd=payload_dir)
    print("[写入指针列]", "OK" if res.get("ok") else json.dumps(res, ensure_ascii=False)[:250])

    # ── 5) 品牌清单 + 多选下拉 ────────────────────────────────────────
    brand_rows = [[{"value": b}] for b in brands]
    bpath = os.path.join(payload_dir, "query_brands.json")
    with open(bpath, "w", encoding="utf-8") as f:
        json.dump(brand_rows, f, ensure_ascii=False)
    res = run([LARK_BIN, "sheets", "+cells-set", "--as", "user", "--url", url,
               "--sheet-name", QUERY_SHEET,
               "--range", f"{BRAND_COL}2:{BRAND_COL}{len(brands) + 1}",
               "--cells", f"@{bpath}"], cwd=payload_dir)
    print(f"[写入品牌清单 {len(brands)} 个]",
          "OK" if res.get("ok") else json.dumps(res, ensure_ascii=False)[:250])

    res = run([LARK_BIN, "sheets", "+dropdown-set", "--as", "user", "--url", url,
               "--sheet-name", QUERY_SHEET, "--range", SEL_CELL.replace("$", ""),
               "--source-range", f"'{QUERY_SHEET}'!{BRAND_COL}2:{BRAND_COL}{len(brands) + 1}",
               "--multiple", "--highlight=false"])
    print("[品牌下拉]", "OK" if res.get("ok") else json.dumps(res, ensure_ascii=False)[:250])

    # ── 6) 汇总条 ────────────────────────────────────────────────────
    a_c = SNAP_COL["签约面积(㎡)"]
    r_c = SNAP_COL["单位日租金(元/㎡/日)"]
    e_c = SNAP_COL["合同最终到期日"]
    stats = {
        f"C{STAT_ROW}": f'=IF({SEL_CELL}="",0,COUNT(${HELPER_COL}$2:${HELPER_COL}${last}))',
        f"E{STAT_ROW}": (f'=IF({SEL_CELL}="",0,SUMIF(${HELPER_COL}$2:${HELPER_COL}${last},">0",'
                        f"'{DATA_SHEET}'!${a_c}$2:${a_c}${last}))"),
        f"G{STAT_ROW}": (f'=IFERROR(SUMPRODUCT((ISNUMBER(${HELPER_COL}$2:${HELPER_COL}${last}))*'
                        f"'{DATA_SHEET}'!${a_c}$2:${a_c}${last}*"
                        f"'{DATA_SHEET}'!${r_c}$2:${r_c}${last})/E{STAT_ROW},0)"),
        f"I{STAT_ROW}": (f'=IFERROR(MINIFS(\'{DATA_SHEET}\'!${e_c}$2:${e_c}${last},'
                        f'${HELPER_COL}$2:${HELPER_COL}${last},">0"),"")'),    }
    for addr, f in stats.items():
        fmt = ("#,##0" if addr.startswith("C") else
               "#,##0.00" if addr.startswith(("E", "G")) else "yyyy-mm-dd")
        one = [[{"formula": f, "cell_styles": {
            "font_size": 13, "font_weight": "bold", "font_color": "#1F3864",
            "number_format": fmt, "horizontal_alignment": "left"}}]]
        r = run([LARK_BIN, "sheets", "+cells-set", "--as", "user", "--url", url,
                 "--sheet-name", QUERY_SHEET, "--range", addr, "--cells",
                 json.dumps(one, ensure_ascii=False)])
        if not r.get("ok"):
            print(f"[统计 {addr}] 失败:", json.dumps(r, ensure_ascii=False)[:200])
    print(f"[统计条] 已写入 第{STAT_ROW}行")

    # ── 7) 版式 + 隐藏辅助区 ──────────────────────────────────────────
    widths = []
    for i, (h, _, _) in enumerate(RESULT_COLS, start=1):
        w = 120
        if h in ("品牌/行业", "项目名称", "商户名称", "三级业态", "数据校验"):
            w = 170
        elif h == "合同编号":
            w = 190
        elif h in ("所属区域", "合同状态", "是否战略品牌"):
            w = 140
        elif "㎡" in h or "元" in h:
            w = 150
        widths.append({"range": f"{col_letter(i)}:{col_letter(i)}",
                       "type": "pixel", "size": w})
    # ⚠ +styles-put 的目标表名要写在 --styles 的 `name` 字段里，
    #   它**不接受** --sheet-name（会报 unknown flag）。
    styles = {
        "styles": [{
            "name": QUERY_SHEET,
            "cell_styles": [
                {"range": f"A{HDR_ROW}:{col_letter(len(RESULT_COLS))}{HDR_ROW}",
                 "font_weight": "bold", "background_color": "#1F3864",
                 "font_color": "#FFFFFF", "horizontal_alignment": "center",
                 "vertical_alignment": "middle"},
            ],
            "row_sizes": [{"range": f"{HDR_ROW}:{HDR_ROW}", "type": "pixel", "size": 28}],
            "col_sizes": widths,
            "freeze": {"rows": HDR_ROW, "cols": 0},
        }],
    }
    res = run([LARK_BIN, "sheets", "+styles-put", "--as", "user", "--url", url,
               "--styles", json.dumps(styles, ensure_ascii=False)])
    print("[版式]", "OK" if res.get("ok") else json.dumps(res, ensure_ascii=False)[:200])

    res = run([LARK_BIN, "sheets", "+dim-hide", "--as", "user", "--url", url,
               "--sheet-name", QUERY_SHEET, "--range", f"{HELPER_COL}:{BRAND_COL}"])
    print("[隐藏辅助列]", "OK" if res.get("ok") else json.dumps(res, ensure_ascii=False)[:200])

    # ── 8) 条件格式：距最终到期天数分级 ──────────────────────────────
    for thr, back, fore in [(90, "#FFE3E3", "#B42318"),
                             (180, "#FFF3E0", "#B54708"),
                             (365, "#FFFCE6", "#A16207")]:
        props = json.dumps({
            "style": {"back_color": back, "fore_color": fore},
            "attrs": [{"formula": [f'=AND(${days_col}{FIRST_ROW}<>"",'
                                   f"${days_col}{FIRST_ROW}<={thr})"]}],
        }, ensure_ascii=False)
        run([LARK_BIN, "sheets", "+cond-format-create", "--as", "user", "--url", url,
             "--sheet-name", QUERY_SHEET, "--rule-type", "expression",
             "--ranges", json.dumps([f"{days_col}{FIRST_ROW}:"
                                     f"{days_col}{FIRST_ROW + N_RESULT - 1}"]),
             "--properties", props])

    print(f"\n[OK] 「{QUERY_SHEET}」已生成：品牌 {len(brands)} 个 / 快照 {n_snap} 行 / "
          f"结果区 {N_RESULT} 行")
    print(f"     下拉 {SEL_CELL}（可多选）；统计条 第{STAT_ROW}行；表头 第{HDR_ROW}行")
    print(f"     位置：{url}")
    print(f"     ⚠ 查询页是**独立表**，公式页出重算问题也只影响这张表，不波及主台账")


if __name__ == "__main__":
    main()
