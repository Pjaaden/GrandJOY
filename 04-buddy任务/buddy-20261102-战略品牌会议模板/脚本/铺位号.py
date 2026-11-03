# -*- coding: utf-8 -*-
"""
铺位号清洗规则 —— 战略品牌会议模板专用。

BI 台账「资源位编号」是系统内部编码，不是人看的铺位号。三种形态：
  "000003B114"                      → B114
  "000003B102;000003B103;..."       → B102-B105（多铺位合并，需缩写）
  "E107;E108"                       → E107-108（同层连续，合并）
  "000003A101;000003A102;000003A102A" → A101/A102（含子铺位）

规则：去掉 "000003" 等前缀 → 若全为同字母前缀+数字，且数字连续或为同一区段，折叠为
「首号-末号」；若含字母后缀（如 A102A）或跨层，保留主号用 "/" 连接。
"""
import re

PREFIX_RE = re.compile(r"^0{4,}\d{0,4}")


def _strip(code):
    c = str(code).strip()
    c = PREFIX_RE.sub("", c)
    return c


def _split(code):
    return [x for x in re.split(r"[;,、\s]+", str(code)) if x]


def clean_unit(code):
    """把 BI 资源位编号转成人看的铺位号。
    注意：BI「资源位编号」有时存的是**资源位名称**（如 EVISU 的 B1F-21），
    而真实铺位号在「资源位名称 / 店铺编号」里，调用方可传 candidates 依次尝试。"""
    if code is None:
        return ""
    code = str(code).strip()
    # B1F-21 这类是楼层+序号，不是标准铺位号；若形如 <楼层>-<数字> 原样返回
    if re.fullmatch(r"[A-Z]{1,2}\d{0,2}F?-\d+", code):
        return code
    parts = [_strip(p) for p in _split(code)]
    parts = [p for p in parts if p]
    if not parts:
        return _strip(code)
    if len(parts) == 1:
        return parts[0]

    # 尝试合并：同字母前缀 + 纯数字 + 数字连续
    m = [re.fullmatch(r"([A-Za-z]+)(\d+)", p) for p in parts]
    if all(m):
        letters = {x.group(1) for x in m}
        if len(letters) == 1:
            L = letters.pop()
            nums = sorted(int(x.group(2)) for x in m)
            if len(nums) == 1:
                return f"{L}{nums[0]}"
            # 连续且两位数以上 → "B114-115"；两铺位 → "E320-321"
            if nums == list(range(nums[0], nums[0] + len(nums))):
                return f"{L}{nums[0]}-{nums[-1]}"
            return f"{L}{nums[0]}/{L}{nums[-1]}"

    # 混合：先去除纯子铺位（A102A 是 A102 的附属）
    mains = []
    for p in parts:
        base = re.sub(r"([A-Za-z]+\d+)[A-Za-z]+$", r"\1", p)
        if base not in mains:
            mains.append(base)
    fm = [re.fullmatch(r"([A-Za-z]+)(\d+)", x) for x in mains]
    if all(fm):
        letters = {x.group(1) for x in fm}
        nums = sorted(int(x.group(2)) for x in fm)
        if len(letters) == 1:
            L = letters.pop()
            if len(nums) == 1:
                return f"{L}{nums[0]}"
            # 混合形态（含附属铺位）一律用 "/" 连接，如 A101/A102
            if nums == list(range(nums[0], nums[0] + len(nums))):
                return "/".join(f"{L}{n}" for n in nums)
    return "/".join(mains)


if __name__ == "__main__":
    tests = [
        "000003B114", "000003B102;000003B103;000003B104;000003B105;000003B203;000003B204;000003B205",
        "E107;E108", "E320;E321", "000003A101;000003A102;000003A102A",
        "E109;E113;E207;E208;E209", "000003BX108;000003BX202;000003BX206",
        "000003C119", "E114", "000003B305",
    ]
    for t in tests:
        print(f"{t[:52]:<54} -> {clean_unit(t)}")
