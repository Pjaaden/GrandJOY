#!/usr/bin/env python3
"""去除 AI 生成图片右下角的 "AI生成 / WORKBUDDY" 水印。

水印是生成器叠加的固定素材：在同一尺寸下，所有图片的水印像素坐标完全一致。
做法：
1. 选一张右下角背景最暗的图作参照 —— 半透明白色叠加在近黑背景上呈中性灰 ~128，
   据此提取水印字形的精确像素集合。
2. 轻微膨胀吃掉抗锯齿边缘。
3. 该掩膜按相同绝对坐标套用到同尺寸的所有图片，cv2.inpaint 修复。
"""
import glob
import os
import sys

import cv2
import numpy as np

WIN_W, WIN_H = 150, 75   # 右下角检测窗口


def detect_glyph(img):
    """在右下角窗口内检测水印字形，返回 (mask, x0, y0)；未检出返回 (None,None,None)。"""
    h, w = img.shape[:2]
    rx0, rx1 = max(0, w - WIN_W), w
    ry0, ry1 = max(0, h - WIN_H), h
    corner = img[ry0:ry1, rx0:rx1]
    b, g, r = corner[:, :, 0].astype(np.int16), corner[:, :, 1].astype(np.int16), corner[:, :, 2].astype(np.int16)
    mx = np.maximum(np.maximum(b, g), r)
    mn = np.minimum(np.minimum(b, g), r)
    val = (b + g + r) / 3.0
    # 暗底 + 半透明白字 -> 中性灰
    sig = ((mx - mn) <= 20) & (val >= 108) & (val <= 150)
    if sig.sum() < 400:
        return None, None, None
    sig = sig.astype(np.uint8)
    n, labels, stats, _ = cv2.connectedComponentsWithStats(sig, connectivity=8)
    if n <= 1:
        return None, None, None
    # 取面积最大的簇
    idx = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))
    x, y, cw, ch, area = stats[idx]
    if area < 400:
        return None, None, None
    out = np.zeros_like(sig)
    out[y:y + ch, x:x + cw] = sig[y:y + ch, x:x + cw]
    return out, rx0 + x, ry0 + y


def main(folder):
    files = sorted(glob.glob(os.path.join(folder, '*.png')))
    groups = {}
    for f in files:
        img = cv2.imread(f, cv2.IMREAD_COLOR)
        if img is None:
            continue
        h, w = img.shape[:2]
        groups.setdefault((w, h), []).append(f)

    total = 0
    for size, group in sorted(groups.items()):
        print(f'尺寸 {size[0]}x{size[1]}: {len(group)} 张')
        # 参照图：右下角最暗
        def corner_dark(f):
            img = cv2.imread(f, cv2.IMREAD_COLOR)
            h, w = img.shape[:2]
            return float(img[h - WIN_H:h, w - WIN_W:w].mean())
        ref = min(group, key=corner_dark)
        ref_img = cv2.imread(ref, cv2.IMREAD_COLOR)
        glyph, gx, gy = detect_glyph(ref_img)
        if glyph is None:
            print(f'  参照图 {os.path.basename(ref)[:44]} 未检出水印，该尺寸跳过')
            continue
        gh, gw = glyph.shape[:2]
        # 膨胀 2px 覆盖抗锯齿边缘与半透明描边
        k = np.ones((5, 5), np.uint8)
        glyph = cv2.dilate(glyph, k, iterations=1)
        gh, gw = glyph.shape[:2]
        print(f'  参照 {os.path.basename(ref)[:44]} -> 水印 {gw}x{gh} @ ({gx},{gy})')
        for f in group:
            img = cv2.imread(f, cv2.IMREAD_COLOR)
            h, w = img.shape[:2]
            mask = np.zeros((h, w), np.uint8)
            x1, y1 = min(gx + gw, w), min(gy + gh, h)
            mask[gy:y1, gx:x1] = glyph[:y1 - gy, :x1 - gx]
            if mask.sum() == 0:
                continue
            out = cv2.inpaint(img, mask, 6, cv2.INPAINT_TELEA)
            cv2.imwrite(f, out)
            total += 1
            print(f'    已去除: {os.path.basename(f)[:52]}')
    print(f'完成，共去除 {total} 张水印')


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else '.')
