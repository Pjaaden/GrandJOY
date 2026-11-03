#!/usr/bin/env python3
"""图片素材统一后处理：去水印 -> 裁边 -> 重命名归档。

生成器把 "AI生成 / WORKBUDDY" 半透明水印固定叠在右下角，同一尺寸下所有图片
的水印像素坐标完全一致。

处理逻辑：
1. 逐图在右下角紧窗口内检测水印字形像素（暗底呈中性灰 / 亮底呈近白）。
2. 取"高置信度"检测结果的并集作为该尺寸的标准字形模板
   （背景干净时检测到的笔画最全，其余图只是漏检部分笔画）。
3. 把模板按同一像素坐标套用到该尺寸所有图片，cv2.inpaint 修复。
4. 裁掉底部残余水印带，统一归档到 图片/ 目录。
"""
import glob
import os
import shutil

import cv2
import numpy as np

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, '图片素材')
DST = os.path.join(BASE, '图片')

# 源文件 -> 主题名
NAME_MAP = {
    'Modern_minimalist_fashion_bout_2026-11-02T01-26-49.png': 'p9_boutique_night.png',
    'Seoul_select_shop_street_at_du_2026-11-02T01-31-49.png': 'p9_seoul_street.png',
    'Minimalist_fashion_boutique_in_2026-11-02T01-33-49.png': 'p9_boutique.png',
    'Limited_fashion_drop_launch_in_2026-11-02T01-59-06.png': 'p9_launch.png',
    'Fresh_seafood_market_stall__li_2026-11-02T01-32-15.png': 'p10_seafood.png',
    'Modern_Korean_barbecue_restaur_2026-11-02T01-32-41.png': 'p10_bbq.png',
    'Seaside_coffee_terrace_overloo_2026-11-02T01-33-19.png': 'p10_cafe.png',
    'Climbers_on_indoor_bouldering__2026-11-02T01-39-08.png': 'p11_climbing.png',
    'Pickleball_court_golden_hour___2026-11-02T01-39-54.png': 'p11_pickleball.png',
    'Outdoor_hiking_gear_laid_out_o_2026-11-02T02-05-29.png': 'p11_outdoor.png',
    'Bookstore_nestled_in_green_hil_2026-11-02T01-59-34.png': 'p12_bookstore.png',
    'Minimal_gallery_room_with_one__2026-11-02T02-00-03.png': 'p12_onebook.png',
    'Young_people_reading_books_at__2026-11-02T02-00-31.png': 'p12_reading.png',
    'Curated_vintage_shop_interior__2026-11-02T02-02-03.png': 'p13_vintage.png',
    'Artisan_hands_crafting_leather_2026-11-02T02-02-55.png': 'p13_craft.png',
    'Outdoor_stepped_pedestrian_all_2026-11-02T02-03-23.png': 'p13_steps.png',
    'Live_music_venue_interior__ban_2026-11-02T02-03-51.png': 'p14_livehouse.png',
    'Rooftop_beer_garden_at_night___2026-11-02T02-04-34.png': 'p14_rooftop.png',
    'Small_black_box_theatre__intim_2026-11-02T02-05-02.png': 'p14_theatre.png',
}

WIN_W, WIN_H = 118, 66      # 右下角检测窗口
CROP_BOTTOM = 58            # 裁掉的底部高度


def detect(img):
    """返回 (角落实例内的水印掩膜, 命中像素数)。"""
    h, w = img.shape[:2]
    rx0, rx1 = max(0, w - WIN_W), w
    ry0, ry1 = max(0, h - WIN_H), h
    corner = img[ry0:ry1, rx0:rx1]
    b, g, r = corner[:, :, 0].astype(np.int16), corner[:, :, 1].astype(np.int16), corner[:, :, 2].astype(np.int16)
    mx = np.maximum(np.maximum(b, g), r)
    mn = np.minimum(np.minimum(b, g), r)
    val = (b + g + r) / 3.0
    chroma = mx - mn
    gray = (chroma <= 20) & (val >= 104) & (val <= 154)   # 暗底上的中性灰字
    white = (chroma <= 26) & (val >= 234)                  # 亮底上的近白字
    sig = (gray | white).astype(np.uint8)
    return sig, int(sig.sum())


def build_template(paths):
    """取高置信度检测结果的并集，得到标准字形模板与全图坐标。"""
    votes = None
    ref_size = None
    best_n = 0
    for p in paths:
        img = cv2.imread(p, cv2.IMREAD_COLOR)
        if img is None:
            continue
        h, w = img.shape[:2]
        sig, n = detect(img)
        if n == 0:
            continue
        if n > best_n:
            best_n = n
            ref_size = (w, h)
        if votes is None or sig.shape != votes.shape:
            if votes is None:
                votes = sig
                ref_size = (w, h)
            continue
        votes = np.maximum(votes, sig)   # 并集：任一图检出即纳入
    if votes is None or votes.sum() == 0:
        return None, None, None
    # 取最大连通簇
    n, labels, stats, _ = cv2.connectedComponentsWithStats(votes, connectivity=8)
    if n <= 1:
        return None, None, None
    idx = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))
    x, y, cw, ch, area = stats[idx]
    tmpl = np.zeros_like(votes)
    tmpl[y:y + ch, x:x + cw] = votes[y:y + ch, x:x + cw]
    # 膨胀覆盖抗锯齿与半透明描边
    tmpl = cv2.dilate(tmpl, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7)), iterations=1)
    w, h = ref_size
    gx = w - WIN_W + x
    gy = h - WIN_H + y
    return tmpl, gx, gy


def main():
    if os.path.isdir(DST):
        shutil.rmtree(DST)
    os.makedirs(DST, exist_ok=True)

    missing = [s for s in NAME_MAP if not os.path.exists(os.path.join(SRC, s))]
    if missing:
        print('缺失源文件:')
        for m in missing:
            print('  ', m)

    # 按尺寸分组
    groups = {}
    for src_name, dst_name in NAME_MAP.items():
        src = os.path.join(SRC, src_name)
        if not os.path.exists(src):
            continue
        img = cv2.imread(src, cv2.IMREAD_COLOR)
        if img is None:
            continue
        h, w = img.shape[:2]
        groups.setdefault((w, h), []).append((src, dst_name))

    for size, items in sorted(groups.items()):
        print(f'尺寸 {size[0]}x{size[1]}: {len(items)} 张')
        tmpl, gx, gy = build_template([s for s, _ in items])
        if tmpl is None:
            print('  未检出水印，跳过')
        else:
            gh, gw = tmpl.shape[:2]
            print(f'  水印模板 {gw}x{gh} @ ({gx},{gy})')
        for src, dst_name in items:
            img = cv2.imread(src, cv2.IMREAD_COLOR)
            h, w = img.shape[:2]
            if tmpl is not None:
                mask = np.zeros((h, w), np.uint8)
                gh, gw = tmpl.shape[:2]
                x1, y1 = min(gx + gw, w), min(gy + gh, h)
                mask[gy:y1, gx:x1] = tmpl[:y1 - gy, :x1 - gx]
                if mask.sum():
                    img = cv2.inpaint(img, mask, 7, cv2.INPAINT_TELEA)
            out_path = os.path.join(DST, dst_name)
            cv2.imwrite(out_path, img)
            print(f'  -> {dst_name:24s} {w}x{h}')

    print(f'完成，输出 {len(os.listdir(DST))} 张到 {DST}')


if __name__ == '__main__':
    main()
