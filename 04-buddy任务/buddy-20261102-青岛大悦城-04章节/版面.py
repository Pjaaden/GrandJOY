#!/usr/bin/env python3
"""第四章（第九-十四页）的统一版面定义 —— HTML 与 PPTX 共用同一套坐标。

坐标系：EMU（1 pt = 12700 EMU），与原 PPT 模板完全一致。
原模板画布 9359900 x 1979930 EMU（10.24" x 2.17"，比例 4.727:1）。

排版参照原模板：
  · 左侧章节栏「04 / 听，先锋潮涌」—— 沿用第 9-14 页原坐标
  · 红色大标题（微软雅黑 加粗 FF0000）—— 沿用第 9 页 14pt
  · 正文（9.5pt）与关键词标签 —— 沿用第 10 页矩形样式（accent1 填充）
  · 图片卡 + 卡内说明条 —— 沿用第 8 页「图片 + 浅蓝说明条 + 黄底斜体副标题」
"""
from PIL import Image, ImageDraw, ImageFont

# ---------- 画布 ----------
EMU_W, EMU_H = 9359900, 1979930

# ---------- 颜色（取自原模板 theme2 / 页面实测） ----------
C_BG = 'FFFCF8'          # 暖白底板
C_ACCENT = '4472C4'      # 主题色 accent1
C_ACCENT_DARK = '2F5597' # accent1 shade
C_ACCENT_TINT = 'DCE3F2' # accent1 lumMod20/lumOff80
C_RED = 'FF0000'         # 大标题红
C_INK = '0F1115'         # 正文墨色
C_YELLOW = 'FFFF00'      # 高亮

# ---------- 字号（pt）----------
SZ_RAIL_NO = 40       # 「04」
SZ_RAIL_NAME = 22     # 「听，先锋潮涌」
SZ_TITLE = 14         # 红色大标题（与原第 9 页一致）
SZ_LEAD = 9.5         # 正文（与原第 9 页一致）
SZ_CHIP = 9.5         # 关键词标签
SZ_CAP_T = 11         # 图注标题
SZ_CAP_S = 9          # 图注副标题
SZ_FOOT = 8           # 页脚小字
# 定义式标语分两段字号：主定位语（第一站）大一号，支撑语（首店/发售平台）小一号。
# 实测：16pt + 8.5pt，「黄海首站 | 区域首店 · 潮流发售平台」恰好在 200 万 EMU 内排下。
SZ_SLOGAN_A = 16
SZ_SLOGAN_B = 8.5

FONT_CN = '微软雅黑'
FONT_NUM = 'Noto Sans SC'

# ---------- 版面坐标（EMU）----------
G = {
    # 左侧章节栏（沿用原第 9-14 页）
    'rail_x': 116550,
    'rail_no_y': 540000,
    'rail_name_y': 1042035,
    'rail_name_w': 1803400,
    # 文案栏
    'text_x': 2000000,
    'text_w': 2000000,
    'title_y': 165000,
    'title_h': 270000,
    'lead_lh': 1.7,
    'lead_gap': 50000,        # 正文与标签之间的间距
    'chip_h': 280000,
    'chip_gap_y': 25000,
    'chip_gap_x': 60000,
    'chip_pad_x': 91440,
    # 图片卡
    'card_right': 9210000,
    'card_top': 150000,
    'card_w': 1620000,
    'card_h': 1675000,
    'card_gap': 110000,
    # 卡内说明条
    'cap_h': 500000,
    # 页脚
    'foot_y': 1700000,
}

# ---------- 文字量宽 ----------
_MF = '/System/Library/Fonts/Hiragino Sans GB.ttc'
_mf_cache = {}


def _font():
    if 'f' not in _mf_cache:
        try:
            _mf_cache['f'] = ImageFont.truetype(_MF, 200, index=1)
        except Exception:
            _mf_cache['f'] = ImageFont.load_default()
    return _mf_cache['f']


def text_w_emu(text, pt):
    """测量文本宽度，换算成 EMU。"""
    d = ImageDraw.Draw(Image.new('L', (8, 8)))
    return int(d.textlength(text, font=_font()) / 200 * pt * 12700)


def wrap_lines(text, pt, width_emu):
    """按可用宽度估算换行后的行数。

    浏览器/PPT 实际字体（微软雅黑、PingFang）与测量用字体（Hiragino）字宽略有差异，
    因此预留 8% 余量，避免实际渲染时多折一行导致溢出。
    """
    d = ImageDraw.Draw(Image.new('L', (8, 8)))
    f = _font()
    limit = (width_emu / 12700.0 * 200 / pt) / 1.08   # 折算到测量字号（含安全余量）
    lines, cur = 1, 0.0
    for ch in text:
        if ch == '\n':
            lines += 1
            cur = 0
            continue
        w = d.textlength(ch, font=f)
        if cur + w > limit:
            lines += 1
            cur = w
        else:
            cur += w
    return lines


def chip_width(text):
    return text_w_emu(text, SZ_CHIP) + G['chip_pad_x'] * 2


def layout_text(lead):
    """计算文案栏内各元素的纵向位置，返回 (lead_y, lead_h, chips_y)。"""
    lead_y = G['title_y'] + G['title_h'] + 40000
    n = wrap_lines(lead, SZ_LEAD, G['text_w'])
    line_h = int(SZ_LEAD * G['lead_lh'] * 12700)
    lead_h = n * line_h
    chips_y = lead_y + lead_h + G['lead_gap']
    return lead_y, lead_h, chips_y


# ---------- 定义式标语（可选文案块，如「黄海先锋时尚第一站 | 区域首店 | 潮流发售平台」）----------
# 结构：第一段 = 定位语（SZ_SLOGAN_A，强调蓝加粗）；其余段 = 支撑语（SZ_SLOGAN_B，墨色加粗）；
# 段间用浅蓝「 | 」。实测宽度：16pt+10pt 恰好在文案栏 200 万 EMU 内排下。
def slogan_parts(slogan):
    """拆出 (定位语, [支撑语, ...])。"""
    rows = [r.strip() for r in slogan.replace('\n', ' ').split('|') if r.strip()]
    return rows[0], rows[1:]


def layout_slogan(slogan, y):
    """返回 (y, h)：标语块占用的纵向范围。"""
    line_h = int(SZ_SLOGAN_A * 1.55 * 12700)
    return y, line_h


def slogan_width(slogan):
    head, rest = slogan_parts(slogan)
    w = text_w_emu(head, SZ_SLOGAN_A)
    for s in rest:
        w += text_w_emu(' | ', SZ_SLOGAN_B) + text_w_emu(s, SZ_SLOGAN_B)
    return w


def check_slogan(page):
    """标语块校验：宽度不溢出文案栏、底部不越界。"""
    problems = []
    if not page.get('slogan'):
        return problems
    w = slogan_width(page['slogan'])
    if w > G['text_w']:
        problems.append(f"标语行超宽：{w} > {G['text_w']}（{page['slogan']}）")
    if page.get('lead'):
        _, _, tag_y = layout_text(page['lead'])
        sy, sh = layout_slogan(page['slogan'], tag_y)
        if sy + sh > EMU_H:
            problems.append(f"标语底部溢出画布：{sy+sh} > {EMU_H}")
    return problems


def layout_chips(keywords, chips_y):
    """把关键词排成两列，返回 [(x, y, w, text), ...]。"""
    half = (len(keywords) + 1) // 2
    col_a, col_b = keywords[:half], keywords[half:]
    wa = max((chip_width(k) for k in col_a), default=0)
    wb = max((chip_width(k) for k in col_b), default=0)
    gap = G['chip_gap_x']
    if wa + wb + gap > G['text_w']:                 # 兜底：整体等比压缩
        s = (G['text_w'] - gap) / (wa + wb)
        wa, wb = int(wa * s), int(wb * s)
    out = []
    xa = G['text_x']
    xb = xa + wa + gap
    for i, k in enumerate(col_a):
        out.append((xa, chips_y + i * (G['chip_h'] + G['chip_gap_y']), wa, k))
    for i, k in enumerate(col_b):
        out.append((xb, chips_y + i * (G['chip_h'] + G['chip_gap_y']), wb, k))
    return out


def layout_cards(n):
    """返回第 i 张卡片的 (x, y, w, h)。"""
    total = n * G['card_w'] + (n - 1) * G['card_gap']
    x0 = G['card_right'] - total
    return [(x0 + i * (G['card_w'] + G['card_gap']), G['card_top'],
             G['card_w'], G['card_h']) for i in range(n)]


def src_rect(img_w, img_h, box_w, box_h):
    """按目标框比例计算 a:srcRect（居中裁切），避免图片拉伸变形。"""
    box_ratio = box_w / box_h
    img_ratio = img_w / img_h
    if img_ratio > box_ratio:          # 图更宽 -> 裁左右
        keep = box_ratio / img_ratio
        c = (1 - keep) / 2
        return int(c * 100000), 0, int(c * 100000), 0
    keep = img_ratio / box_ratio        # 图更高 -> 裁上下
    c = (1 - keep) / 2
    return 0, int(c * 100000), 0, int(c * 100000)


def check_page(page):
    """自动校验一页的排版：溢出、越界、元素重叠。返回问题列表。"""
    problems = []
    lead_y, lead_h, tag_y = layout_text(page['lead'])
    has_slogan = bool(page.get('slogan'))
    # 有标语时，标语占据 chip 行的位置（二选一，不并存）
    boxes = [] if has_slogan else layout_chips(page['keywords'], tag_y)
    cards = layout_cards(len(page['images']))

    # 1. 文案区不能越出画布下边界
    if has_slogan:
        _, sh = layout_slogan(page['slogan'], tag_y)
        tag_bottom = tag_y + sh
        if tag_bottom > EMU_H:
            problems.append(f"标语块底部溢出画布：{tag_bottom} > {EMU_H}")
    else:
        tag_bottom = max((y + G['chip_h']) for _, y, _, _ in boxes) if boxes else tag_y
        if tag_bottom > EMU_H:
            problems.append(f"关键词标签底部溢出画布：{tag_bottom} > {EMU_H}")

    # 2. 正文底不能压到标签/标语
    if lead_y + lead_h > tag_y:
        problems.append(f"正文与标签行重叠：lead底={lead_y+lead_h} > 标签顶={tag_y}")

    # 3. 文案栏右边界不能压到第一张图卡
    text_right = G['text_x'] + G['text_w']
    if text_right > cards[0][0]:
        problems.append(f"文案栏压到图卡：text右={text_right} > 卡左={cards[0][0]}")

    # 4. 图卡整体不能越出画布
    for i, (x, y, w, h) in enumerate(cards):
        if x < 0 or y < 0 or x + w > EMU_W or y + h > EMU_H:
            problems.append(f"第{i+1}张图卡越界：({x},{y},{w},{h})")

    # 5. 图卡之间不能重叠
    for i in range(len(cards) - 1):
        a, b = cards[i], cards[i + 1]
        if a[0] + a[2] > b[0]:
            problems.append(f"第{i+1}、{i+2}张图卡重叠")

    # 6. 左侧章节栏不能压到文案栏
    if G['rail_x'] + G['rail_name_w'] > G['text_x']:
        problems.append("章节栏压到文案栏")

    # 7. 关键词两列不能超出文案栏
    if boxes:
        chip_right = max(x + w for x, _, w, _ in boxes)
        if chip_right > G['text_x'] + G['text_w']:
            problems.append(f"关键词标签越出文案栏：{chip_right} > {G['text_x']+G['text_w']}")

    # 8. 标语块宽度
    problems.extend(check_slogan(page))

    return problems
