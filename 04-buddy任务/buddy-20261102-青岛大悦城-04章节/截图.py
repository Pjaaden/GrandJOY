#!/usr/bin/env python3
"""用 playwright + 本机 Edge 把 HTML 演示稿逐页按像素截图。

Edge --headless --screenshot 会把 cqw 按视口宽而非容器宽计算，整页缩成一团，
因此改用 playwright 驱动本机 Edge 渲染后按 px 裁剪。

用法：python3 截图.py [html路径] [输出目录]
"""
import asyncio
import os
import sys

from PIL import Image

EDGE_CHANNEL = 'msedge'
BASE = os.path.dirname(os.path.abspath(__file__))
HTML = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    BASE, '输出', '青岛大悦城推介_第四章_先锋潮涌.html')
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE, '输出', 'shots')
PAGES = ['p9', 'p10', 'p11', 'p12', 'p13', 'p14']
W = 2400
SCALE = 2


async def main():
    from playwright.async_api import async_playwright

    os.makedirs(OUT, exist_ok=True)
    h = round(W * 1979930 / 9359900)
    async with async_playwright() as pw:
        br = await pw.chromium.launch(channel=EDGE_CHANNEL)
        pg = await br.new_page(viewport={'width': W, 'height': h + 20},
                               device_scale_factor=SCALE)
        await pg.goto('file://' + HTML, wait_until='load')
        await pg.wait_for_timeout(1200)
        for i, pid in enumerate(PAGES):
            el = await pg.query_selector(f'.slide#{pid}')
            if not el:
                print(f'{pid} 未找到')
                continue
            shot = os.path.join(OUT, f'{i+1}_{pid}.png')
            await el.screenshot(path=shot)
            im = Image.open(shot)
            print(f'{pid} -> {shot} ({os.path.getsize(shot)//1024} KB, {im.size[0]}x{im.size[1]})')
        await br.close()


asyncio.run(main())
