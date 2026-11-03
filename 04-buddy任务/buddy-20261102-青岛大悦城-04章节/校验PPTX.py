#!/usr/bin/env python3
"""PPTX 渲染校验器：直接解析生成出来的 slide XML，按 shape 实际坐标绘制 PNG。

用途：本机没有可用的 PPT->PDF 渲染器（LibreOffice/WPS CLI 均不可用），
因此写一个最小渲染器，按 OOXML 中的真实 xfrm 坐标与图片关系还原画面，
用来检查是否有溢出、重叠、错位等问题。
"""
import os
import re
import sys
import zipfile

from PIL import Image, ImageDraw, ImageFont

# 该 PPT 画布是 10.24 英寸的超宽横幅，按 pt 计算的文字磅值极小（
# 换算后 px 只有几像素），PIL 在小字号下计算遮罩会触发误报的 DecompressionBombError。
Image.MAX_IMAGE_PIXELS = None

SCALE = 2400 / 9359900  # 目标宽度 2400px

FONT_DIR = '/System/Library/Fonts'
FONT_CN = '/System/Library/Fonts/Hiragino Sans GB.ttc'
FONT_NUM = '/System/Library/Fonts/Hiragino Sans GB.ttc'
FONT_EN = '/System/Library/Fonts/Supplemental/Arial.ttf'


def load_font(sz_pt, bold=False, italic=False, cjk=True):
    px = max(8, int(sz_pt / 72 * (2400 / (9359900 / 914400))))
    if cjk:
        idx = 1 if bold else 0
        try:
            return ImageFont.truetype(FONT_CN, px, index=idx)
        except Exception:
            pass
    name = 'Arial Bold Italic.ttf' if (bold and italic) else (
        'Arial Bold.ttf' if bold else ('Arial Italic.ttf' if italic else 'Arial.ttf'))
    try:
        return ImageFont.truetype(os.path.join(FONT_DIR, 'Supplemental', name), px)
    except Exception:
        return ImageFont.load_default()


def solid_of(shape_xml, sppr_only=False):
    """取填充色。sppr_only=True 时只在 <p:spPr> 内查找，避免误取文字颜色。"""
    scope = shape_xml
    if sppr_only:
        m = re.search(r'<p:spPr>.*?</p:spPr>', shape_xml, re.S)
        if not m:
            return None
        scope = m.group(0)
    m = re.search(r'<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', scope)
    if m:
        return '#' + m.group(1)
    m = re.search(r'<a:solidFill>\s*<a:schemeClr val="([a-z0-9]+)"', scope)
    if m:
        return {'accent1': '#4472C4', 'tx1': '#000000', 'lt1': '#FFFFFF',
                'dk1': '#000000'}.get(m.group(1), '#CCCCCC')
    return None


def alpha_of(shape_xml):
    m = re.search(r'<a:srgbClr val="[0-9A-Fa-f]{6}">\s*<a:alpha val="(\d+)"', shape_xml)
    if m:
        return int(m.group(1)) / 100000
    return 1.0


def render_slide(xml, rels, zf, out_png):
    W, H = 2400, int(1979930 * SCALE)
    img = Image.new('RGB', (W, H), '#FFFCF8')
    d = ImageDraw.Draw(img, 'RGBA')

    # 背景渐变
    grad = re.search(r'<p:bg>.*?</p:bg>', xml, re.S)
    if grad:
        g = grad.group(0)
        stops = re.findall(r'<a:gs pos="(\d+)">\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', g)
        if len(stops) >= 2:
            cols = [(int(p) / 100000, '#' + c) for p, c in stops]
            grad_img = Image.new('RGB', (W, H))
            gd = ImageDraw.Draw(grad_img)
            for y in range(H):
                for i in range(len(cols) - 1):
                    p0, c0 = cols[i]
                    p1, c1 = cols[i + 1]
                    if p0 <= y / H <= p1:
                        t = (y / H - p0) / max(1e-6, (p1 - p0))
                        c0r = tuple(int(c0[j:j + 2], 16) for j in (1, 3, 5))
                        c1r = tuple(int(c1[j:j + 2], 16) for j in (1, 3, 5))
                        col = tuple(int(c0r[k] + (c1r[k] - c0r[k]) * t) for k in range(3))
                        gd.line([(0, y), (W, y)], fill=col)
                        break
            img = grad_img
            d = ImageDraw.Draw(img, 'RGBA')

    shapes = [m.group(0) for m in re.finditer(r'<p:(?:sp|pic)>.*?</p:(?:sp|pic)>', xml, re.S)]
    for sh in shapes:
        xf = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"/>', sh)
        if not xf:
            continue
        x, y, cx, cy = (int(v) for v in xf.groups())
        px, py = x * SCALE, y * SCALE
        pw, ph = cx * SCALE, cy * SCALE

        if sh.startswith('<p:pic'):
            emb = re.search(r'r:embed="([^"]+)"', sh)
            if emb:
                tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % emb.group(1), rels)
                if tgt:
                    path = tgt.group(1).replace('../', 'ppt/')
                    try:
                        with zf.open(path) as fh:
                            import io
                            pim = Image.open(io.BytesIO(fh.read())).convert('RGB')
                        pim = pim.resize((max(1, int(pw)), max(1, int(ph))), Image.LANCZOS)
                        img.paste(pim, (int(px), int(py)))
                    except Exception as e:
                        d.rectangle([px, py, px + pw, py + ph], outline='#FF00FF', width=3)
            continue

        # 文本框 / 形状
        fill = solid_of(sh, sppr_only=True)
        al = alpha_of(sh)
        if fill:
            rgba = tuple(int(fill[j:j + 2], 16) for j in (1, 3, 5)) + (int(al * 255),)
            d.rectangle([px, py, px + pw, py + ph], fill=rgba)

        # 文字（同段落内各 run 字号可能不同，逐 run 排布）
        for pi, para in enumerate(re.findall(r'<a:p>.*?</a:p>', sh, re.S)):
            runs = re.findall(r'<a:r>(.*?)</a:r>', para, re.S)
            if not runs:
                continue
            algn = re.search(r'algn="(\w+)"', para)
            algn = algn.group(1) if algn else 'l'
            total_w = 0
            for r in runs:
                t = ''.join(re.findall(r'<a:t>(.*?)</a:t>', r, re.S))
                t = (t.replace('&amp;', '&').replace('&lt;', '<')
                      .replace('&gt;', '>').replace('&quot;', '"'))
                m = re.search(r'sz="(\d+)"', r)
                rsz = int(m.group(1)) / 100 if m else 1800
                rf = load_font(rsz, 'b="1"' in r, 'i="1"' in r)
                total_w += d.textlength(t, font=rf)
            if algn == 'ctr':
                tx = px + (pw - total_w) / 2
            elif algn == 'r':
                tx = px + pw - total_w
            else:
                tx = px + 6
            ty = py + 6 + pi * (18 / 72 * (2400 / (9359900 / 914400)) * 1.45)
            for r in runs:
                t = ''.join(re.findall(r'<a:t>(.*?)</a:t>', r, re.S))
                t = (t.replace('&amp;', '&').replace('&lt;', '<')
                      .replace('&gt;', '>').replace('&quot;', '"'))
                m = re.search(r'sz="(\d+)"', r)
                rsz = int(m.group(1)) / 100 if m else 1800
                rf = load_font(rsz, 'b="1"' in r, 'i="1"' in r)
                col = solid_of(r) or '#000000'
                hl = re.search(r'<a:highlight>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', r)
                tw = d.textlength(t, font=rf)
                if hl:
                    hh = int(rsz / 72 * (2400 / (9359900 / 914400)) * 1.45)
                    base = ty + (18 - rsz) / 72 * (2400 / (9359900 / 914400))
                    d.rectangle([tx - 3, base, tx + tw + 3, base + hh], fill='#' + hl.group(1))
                d.text((tx, ty), t, font=rf, fill=col)
                tx += tw

    img.save(out_png)
    return out_png


def main():
    pptx = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else '/tmp/ppt_probe/pptx_render'
    os.makedirs(outdir, exist_ok=True)
    with zipfile.ZipFile(pptx) as zf:
        names = zf.namelist()
        pres = zf.read('ppt/presentation.xml').decode('utf-8')
        prels = zf.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
        order = re.findall(r'<p:sldId id="\d+" r:id="(rId\d+)"/>', pres)
        rm = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', prels))
        for i, rid in enumerate(order, 1):
            target = rm[rid]
            spath = 'ppt/' + target.replace('../', '')
            relpath = 'ppt/slides/_rels/' + os.path.basename(target) + '.rels'
            try:
                rels = zf.read(relpath).decode('utf-8')
            except KeyError:
                rels = ''
            xml = zf.read(spath).decode('utf-8')
            out = os.path.join(outdir, f'pptx_{i}.png')
            render_slide(xml, rels, zf, out)
            print(f'slide{i} -> {out}')


if __name__ == '__main__':
    main()
