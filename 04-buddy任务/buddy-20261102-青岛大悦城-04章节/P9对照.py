#!/usr/bin/env python3
"""P9 文案对照图：定义式标语（「XX第一站 | 区域首店 | 潮流发售平台」）候选比较。

用 playwright + 本机 Edge 按像素精确截图，竖向拼成一张图。
产物：输出/_对比/文案对比.png
"""
import asyncio
import json
import os

from PIL import Image

import 生成HTML as H
import 版面

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE, '输出', '_对比')
PAGE_W = 2200
TAG_H = 48

VARIANTS = [
    ('A｜黄海首站 | 潮流发售平台    （短促有力，「第一站+发售平台」两层定义都在）', {
        'id': 'vA', 'title': '潮起黄海・新品共展',
        'lead': '凭黄海航线与免签便利，接住首尔当季潮流。招的不是品牌，是首发权。',
        'slogan': '黄海首站 | 潮流发售平台',
    }),
    ('B｜黄海首站 | 区域首店·潮流发售平台    （三段全保留，压缩分隔符）', {
        'id': 'vB', 'title': '潮起黄海・新品共展',
        'lead': '凭黄海航线与免签便利，接住首尔当季潮流。招的不是品牌，是首发权。',
        'slogan': '黄海首站 | 区域首店·潮流发售平台',
    }),
    ('C｜国际首站 | 潮流发售平台    （只换地理词，最接近原始口径）', {
        'id': 'vC', 'title': '国际潮头・新品共展',
        'lead': '一小时航程接住首尔、东京当季潮流。招的不是品牌，是首发权。',
        'slogan': '国际首站 | 潮流发售平台',
    }),
]

CSS = f'''<style>
  body{{padding-top:{TAG_H}px!important;background:#fff!important}}
  .tag{{position:absolute;left:14px;top:8px;font:600 15px/1.6 "PingFang SC";
       color:#2F5597;letter-spacing:.04em}}
  .slide{{max-width:100%!important;margin:0!important;outline:1px solid #D9D3C9}}
</style>'''


async def main():
    from playwright.async_api import async_playwright

    with open(os.path.join(BASE, '内容.json'), encoding='utf-8') as f:
        data = json.load(f)
    meta = data['meta']
    p9 = next(p for p in data['pages'] if p['id'] == 'p9')
    p9 = dict(p9)
    p9.pop('keywords', None)

    os.makedirs(OUT_DIR, exist_ok=True)
    uri_map = {im['src']: H.img_to_datauri(os.path.join(BASE, '图片', im['src']))
               for im in p9['images']}

    page_h = PAGE_W * 1979930 // 9359900
    shots = []
    async with async_playwright() as pw:
        br = await pw.chromium.launch(channel='msedge')
        pg = await br.new_page(viewport={'width': PAGE_W, 'height': page_h + 10},
                               device_scale_factor=2)
        for i, (label, v) in enumerate(VARIANTS):
            page = dict(p9); page.update(v)
            page['keywords'] = []
            page['no'] = meta['chapter_no']
            page['chapter'] = meta['chapter_name']
            probs = 版面.check_page(page)
            print(f'{label[:10]}: {"版面 OK" if not probs else probs}')

            html = H.build_html(H.build_page(page, uri_map)).replace('</head>', CSS + '</head>')
            html = html.replace('<body>', f'<body><div class="tag">{label}</div>', 1)
            tmp = os.path.join(OUT_DIR, f'_v{i}.html')
            with open(tmp, 'w', encoding='utf-8') as f:
                f.write(html)
            await pg.goto('file://' + tmp, wait_until='load')
            await pg.wait_for_timeout(700)
            shot = os.path.join(OUT_DIR, f'_v{i}.png')
            await pg.screenshot(path=shot, clip={'x': 0, 'y': 0,
                                                 'width': PAGE_W, 'height': page_h + TAG_H})
            shots.append(shot)
        await br.close()

    ims = [Image.open(s) for s in shots]
    gap = 6
    canvas = Image.new('RGB', (ims[0].width, sum(im.height for im in ims) + gap * (len(ims) - 1)), 'white')
    y = 0
    for im in ims:
        canvas.paste(im, (0, y))
        y += im.height + gap
    out = os.path.join(OUT_DIR, '文案对比.png')
    canvas.save(out)
    canvas.resize((canvas.width // 2, canvas.height // 2), Image.LANCZOS).save(
        os.path.join(OUT_DIR, '文案对比_小.png'))
    print('已生成', out, f'{os.path.getsize(out)//1024} KB', canvas.size)


asyncio.run(main())
