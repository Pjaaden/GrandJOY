#!/usr/bin/env python3
"""生成第四章（第 9-14 页）的单文件 HTML 演示稿。

版面坐标与 PPTX 版完全共用 版面.py，保证两种产物视觉一致。
设计沿用原 PPT 模板：暖白底、accent1 蓝、红色大标题、左侧章节栏、
图片卡 + 卡内浅蓝说明条（黄底斜体副标题）。

用法：python3 生成HTML.py
"""
import base64
import io
import json
import os

from PIL import Image

import 版面
from 版面 import G, EMU_W, EMU_H

BASE = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE, '图片')
OUT_DIR = os.path.join(BASE, '输出')
OUT_FILE = os.path.join(OUT_DIR, '青岛大悦城推介_第四章_先锋潮涌.html')

FONT_CN = ('"Microsoft YaHei","微软雅黑","PingFang SC","Hiragino Sans GB",'
           '"Noto Sans SC","Source Han Sans SC",sans-serif')
FONT_NUM = ('"Noto Sans SC","PingFang SC","Hiragino Sans GB",'
            '"Microsoft YaHei",sans-serif')


def emu2cqw(e):
    """EMU -> cqw（容器宽度的百分比），cqw 保证缩放自适应。"""
    return e / EMU_W * 100


def emu2cqh(e):
    return e / EMU_H * 100


def pt2cqw(pt):
    """pt -> cqw（相对画布宽度）。"""
    return pt * 12700 / EMU_W * 100


def img_to_datauri(path, max_w=1400, quality=86):
    im = Image.open(path).convert('RGB')
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, 'JPEG', quality=quality, optimize=True, progressive=True)
    return 'data:image/jpeg;base64,' + base64.b64encode(buf.getvalue()).decode('ascii')


def build_page(page, uri_map):
    boxes = 版面.layout_cards(len(page['images']))
    cards = []
    for im, (x, y, w, h) in zip(page['images'], boxes):
        cards.append(f'''
      <figure class="card" style="left:{emu2cqw(x):.4f}cqw; top:{emu2cqh(y):.4f}cqh;
              width:{emu2cqw(w):.4f}cqw; height:{emu2cqh(h):.4f}cqh;">
        <img src="{uri_map[im['src']]}" alt="{im['cap_t']}">
        <figcaption class="cap" style="height:{emu2cqh(G['cap_h']):.4f}cqh;">
          <span class="cap-t">{im['cap_t']}</span>
          <span class="cap-s">{im['cap_s']}</span>
        </figcaption>
      </figure>''')

    chips = []
    lead_y, lead_h, chips_y = 版面.layout_text(page['lead'])
    if page.get('slogan'):
        _, sh = 版面.layout_slogan(page['slogan'], chips_y)
        head, rest = 版面.slogan_parts(page['slogan'])
        inner = f'<span class="sl-head">{head}</span>'
        for s in rest:
            inner += f'<span class="sl-sep"> | </span><span class="sl-sub">{s}</span>'
        chips.append(f'<div class="slogan" style="left:{emu2cqw(G["text_x"]):.4f}cqw; '
                     f'top:{emu2cqh(chips_y):.4f}cqh; width:{emu2cqw(G["text_w"]):.4f}cqw; '
                     f'height:{emu2cqh(sh):.4f}cqh;">{inner}</div>')
    for (x, y, w, kw) in ([] if page.get('slogan')
                          else 版面.layout_chips(page['keywords'], chips_y)):
        chips.append(f'<span class="chip" style="left:{emu2cqw(x):.4f}cqw; '
                     f'top:{emu2cqh(y):.4f}cqh; width:{emu2cqw(w):.4f}cqw; '
                     f'height:{emu2cqh(G["chip_h"]):.4f}cqh;">{kw}</span>')

    return f'''
  <section class="slide" id="{page['id']}">
    <div class="rail">
      <div class="rail-no" style="left:{emu2cqw(G['rail_x']):.4f}cqw; top:{emu2cqh(G['rail_no_y']):.4f}cqh;">{page['no']}</div>
      <div class="rail-name" style="left:{emu2cqw(G['rail_x']):.4f}cqw; top:{emu2cqh(G['rail_name_y']):.4f}cqh;">{page['chapter']}</div>
      <div class="rail-foot" style="left:{emu2cqw(G['rail_x']):.4f}cqw; top:{emu2cqh(G['foot_y']):.4f}cqh;">青岛大悦城 · 品牌推介发布会</div>
    </div>

    <div class="text">
      <h2 class="title" style="left:{emu2cqw(G['text_x']):.4f}cqw; top:{emu2cqh(G['title_y']):.4f}cqh;
          width:{emu2cqw(G['text_w']):.4f}cqw;">{page['title']}</h2>
      <p class="lead" style="left:{emu2cqw(G['text_x']):.4f}cqw; top:{emu2cqh(lead_y):.4f}cqh;
         width:{emu2cqw(G['text_w']):.4f}cqw; height:{emu2cqh(lead_h):.4f}cqh;">{page['lead']}</p>
    </div>

    <div class="chips">{''.join(chips)}
    </div>
    <div class="cards">{''.join(cards)}
    </div>
  </section>'''


def main():
    with open(os.path.join(BASE, '内容.json'), encoding='utf-8') as f:
        data = json.load(f)
    meta, pages = data['meta'], data['pages']
    for p in pages:
        p['no'] = meta['chapter_no']
        p['chapter'] = meta['chapter_name']

    os.makedirs(OUT_DIR, exist_ok=True)

    # 版面自检：溢出 / 重叠 / 越界
    bad = 0
    for p in pages:
        for prob in 版面.check_page(p):
            print(f'  [版面警告] {p["id"]}: {prob}')
            bad += 1
    if bad:
        print(f'共 {bad} 处版面问题，请调整 内容.json / 版面.py 后重试')
        raise SystemExit(1)
    print('版面自检通过')

    uri_map = {}
    for p in pages:
        for im in p['images']:
            if im['src'] in uri_map:
                continue
            uri_map[im['src']] = img_to_datauri(os.path.join(IMG_DIR, im['src']))
    print(f'已编码 {len(uri_map)} 张图片')

    body = '\n'.join(build_page(p, uri_map) for p in pages)
    html = build_html(body)
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'已生成 {OUT_FILE}  ({os.path.getsize(OUT_FILE)/1024/1024:.2f} MB, {len(pages)} 页)')


def build_html(body):
    """把若干页拼成单文件 HTML（版面样式与演示稿完全一致）。"""
    return f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>青岛大悦城 · 品牌推介发布会 — 第四章 先锋潮涌</title>
<style>
  :root {{
    --bg: #{版面.C_BG};
    --accent: #{版面.C_ACCENT};
    --accent-dark: #{版面.C_ACCENT_DARK};
    --accent-tint: #{版面.C_ACCENT_TINT};
    --red: #{版面.C_RED};
    --ink: #{版面.C_INK};
    --yellow: #{版面.C_YELLOW};
    --font-cn: {FONT_CN};
    --font-num: {FONT_NUM};
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  html, body {{ background:#E9E5DE; }}
  body {{ font-family:var(--font-cn); color:var(--ink);
          -webkit-font-smoothing:antialiased; padding:2.6cqw 0 4cqw; }}
  .deck {{ container-type:inline-size; width:100%; }}

  .slide {{
    position:relative; width:100%;
    aspect-ratio:{EMU_W} / {EMU_H};
    background:var(--bg); overflow:hidden;
    margin:0 auto 2.4cqw; max-width:2400px;
  }}
  .slide > * {{ position:absolute; }}

  /* 左侧章节栏 */
  .rail-no {{
    position:absolute; font-family:var(--font-num); font-weight:400;
    font-size:{pt2cqw(版面.SZ_RAIL_NO):.4f}cqw; line-height:1; color:var(--ink);
  }}
  .rail-name {{
    position:absolute; font-family:var(--font-num); font-weight:400;
    font-size:{pt2cqw(版面.SZ_RAIL_NAME):.4f}cqw; line-height:1.2; color:var(--ink);
    white-space:nowrap;
  }}
  .rail-foot {{
    position:absolute; font-size:{pt2cqw(版面.SZ_FOOT):.4f}cqw;
    color:#9AA0A6; letter-spacing:0.05em; white-space:nowrap;
  }}

  /* 文案 */
  .title {{
    position:absolute; font-weight:700; color:var(--red);
    font-size:{pt2cqw(版面.SZ_TITLE):.4f}cqw; line-height:1.25;
  }}
  .lead {{
    position:absolute; font-size:{pt2cqw(版面.SZ_LEAD):.4f}cqw;
    line-height:{G['lead_lh']}; color:#2B2F36; text-align:justify;
  }}

  /* 关键词标签 */
  .chip {{
    position:absolute; display:flex; align-items:center; justify-content:center;
    background:var(--accent); color:#fff;
    font-size:{pt2cqw(版面.SZ_CHIP):.4f}cqw; font-weight:700; letter-spacing:0.02em;
    border:{emu2cqw(2540):.4f}cqw solid var(--accent-dark);
    white-space:nowrap; overflow:hidden;
  }}

  /* 定义式标语（替代标签行）：定位语 16pt 蓝，支撑语 10pt 墨色 */
  .slogan {{
    position:absolute; display:flex; align-items:baseline;
    white-space:nowrap; font-weight:700; line-height:1.55;
  }}
  .slogan .sl-head {{
    font-size:{pt2cqw(版面.SZ_SLOGAN_A):.4f}cqw; color:var(--accent);
    letter-spacing:0.01em;
  }}
  .slogan .sl-sub {{
    font-size:{pt2cqw(版面.SZ_SLOGAN_B):.4f}cqw; color:var(--ink);
  }}
  .slogan .sl-sep {{
    font-size:{pt2cqw(版面.SZ_SLOGAN_B):.4f}cqw; color:#7D9AD1; font-weight:700;
    padding:0 {emu2cqw(12700):.4f}cqw;
  }}

  /* 图片卡 */
  .card {{ position:absolute; overflow:hidden; background:#EFEAE3; }}
  .card img {{ width:100%; height:100%; object-fit:cover; display:block; }}
  .cap {{
    position:absolute; left:0; top:0; right:0;
    background:rgba(220,227,242,0.94);
    padding:{emu2cqh(18288):.4f}cqh {emu2cqw(45720):.4f}cqw;
    display:flex; flex-direction:column; gap:{emu2cqh(9144):.4f}cqh;
  }}
  .cap-t {{
    font-size:{pt2cqw(版面.SZ_CAP_T):.4f}cqw; font-weight:700; color:var(--accent);
    line-height:1.25; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
  }}
  .cap-s {{
    font-size:{pt2cqw(版面.SZ_CAP_S):.4f}cqw; font-style:italic; color:var(--ink);
    line-height:1.35; background:var(--yellow); padding:0 {emu2cqw(2540):.4f}cqw;
    justify-self:start; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    max-width:100%;
  }}

  @page {{ size:2400px 508px; margin:0; }}
  @media print {{
    html, body {{ background:#fff; padding:0; }}
    .slide {{ margin:0; page-break-after:always; break-after:page; }}
    .slide:last-child {{ page-break-after:auto; break-after:auto; }}
  }}
</style>
</head>
<body>
  <div class="deck">
{body}
  </div>
</body>
</html>
'''


if __name__ == '__main__':
    main()
