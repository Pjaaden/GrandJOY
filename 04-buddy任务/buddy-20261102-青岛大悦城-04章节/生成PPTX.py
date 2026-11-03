#!/usr/bin/env python3
"""基于原 PPT 模板生成「第四章 先锋潮涌」6 页 PPTX。

策略：直接改写原 pptx 的 zip 结构 —— 保留 master / layout / theme / 背景渐变，
只重建第 9-14 页（第四章），使产物与原文件完全同源同版式。

版面坐标全部来自 版面.py，与 HTML 版共用同一套定义。

产出：输出/青岛大悦城推介_第四章_先锋潮涌.pptx
"""
import io
import json
import os
import re
import shutil
import zipfile

from PIL import Image

import 版面
from 版面 import G, EMU_W, EMU_H

BASE = os.path.dirname(os.path.abspath(__file__))
SRC_PPTX = ('/Users/zhijian/Library/Mobile Documents/com~apple~CloudDocs/Desktop/'
            '大悦城GrandJOY/_总部招商/筹开:升级项目/青岛大悦城/'
            '青岛大悦城品牌推介发布会 · 演讲大纲_v4.pptx')
IMG_DIR = os.path.join(BASE, '图片')
OUT_DIR = os.path.join(BASE, '输出')
OUT_PPTX = os.path.join(OUT_DIR, '青岛大悦城推介_第四章_先锋潮涌.pptx')


def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
             .replace('"', '&quot;'))


def run(text, sz, color, bold=False, italic=False, font=版面.FONT_CN, highlight=None):
    attrs = f'lang="zh-CN" altLang="en-US" sz="{int(round(sz))}" b="{1 if bold else 0}" i="{1 if italic else 0}"'
    fill = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
    hl = (f'<a:highlight><a:srgbClr val="{highlight}"/></a:highlight>' if highlight else '')
    fnt = (f'<a:latin typeface="{font}" panose="020B0503020204020204" charset="-122"/>'
           f'<a:ea typeface="{font}" panose="020B0503020204020204" charset="-122"/>'
           f'<a:sym typeface="+mn-ea"/>')
    return (f'<a:r><a:rPr {attrs}>{fill}{hl}{fnt}</a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r>')


def txbox(sid, name, x, y, w, h, runs, align='l', anchor='t'):
    return f'''<p:sp>
<p:nvSpPr><p:cNvPr id="{sid}" name="{name}" txBox="1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="{anchor}"/>
<a:lstStyle/><a:p><a:pPr marL="0" indent="0" algn="{align}"><a:lnSpc><a:spcPct val="{int(G['lead_lh']*100000)}"/></a:lnSpc></a:pPr>{''.join(runs)}</a:p></p:txBody>
</p:sp>'''


def chip(sid, name, x, y, w, h, text):
    return f'''<p:sp>
<p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{版面.C_ACCENT}"/></a:solidFill>
<a:ln w="12700"><a:solidFill><a:srgbClr val="{版面.C_ACCENT_DARK}"/></a:solidFill></a:ln></p:spPr>
<p:txBody><a:bodyPr wrap="square" lIns="{G['chip_pad_x']}" tIns="0" rIns="{G['chip_pad_x']}" bIns="0" rtlCol="0" anchor="ctr"/>
<a:lstStyle/><a:p><a:pPr marL="0" indent="0" algn="ctr"/>
{run(text, 版面.SZ_CHIP*100, 'FFFFFF', bold=True)}</a:p></p:txBody>
</p:sp>'''


def caption(sid, name, x, y, w, h, title, sub):
    return f'''<p:sp>
<p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{版面.C_ACCENT_TINT}"><a:alpha val="94000"/></a:srgbClr></a:solidFill>
<a:ln><a:noFill/></a:ln></p:spPr>
<p:txBody><a:bodyPr wrap="square" lIns="45720" tIns="18288" rIns="45720" bIns="0" rtlCol="0" anchor="t"/>
<a:lstStyle/>
<a:p><a:pPr marL="0" indent="0" algn="l"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>{run(title, 版面.SZ_CAP_T*100, 版面.C_ACCENT, bold=True)}</a:p>
<a:p><a:pPr marL="0" indent="0" algn="l"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>{run(sub, 版面.SZ_CAP_S*100, 版面.C_INK, italic=True, highlight=版面.C_YELLOW)}</a:p>
</p:txBody>
</p:sp>'''


def picture(sid, name, x, y, w, h, rid, l, t, r, b):
    return f'''<p:pic>
<p:nvPicPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
<p:blipFill><a:blip r:embed="{rid}"/><a:srcRect l="{l}" t="{t}" r="{r}" b="{b}"/>
<a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>'''


def slogan_box(sid, name, x, y, w, h, slogan):
    """定义式标语：第一段 16pt 强调蓝加粗（定位），其余 10pt 墨色加粗（支撑），
    段间浅蓝「 | 」。单段文本，行内变字号靠各自的 <a:r> 控制。"""
    head, rest = 版面.slogan_parts(slogan)
    runs = [run(head, 版面.SZ_SLOGAN_A * 100, 版面.C_ACCENT, bold=True)]
    for s in rest:
        runs.append(run(' | ', 版面.SZ_SLOGAN_B * 100, 'A9BEE0'))
        runs.append(run(s, 版面.SZ_SLOGAN_B * 100, 版面.C_INK, bold=True))
    return f'''<p:sp>
<p:nvSpPr><p:cNvPr id="{sid}" name="{name}" txBox="1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"/>
<a:lstStyle/><a:p><a:pPr marL="0" indent="0" algn="l">
<a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>{''.join(runs)}</a:p></p:txBody>
</p:sp>'''


def build_slide(page, rid_map, crop_map, uid):
    sid = uid
    shapes = []

    # 暖白底板（原第 9 页 AutoShape 2）
    shapes.append(f'''<p:sp>
<p:nvSpPr><p:cNvPr id="{sid}" name="AutoShape 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9360000" cy="1980000"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 0"/></a:avLst></a:prstGeom>
<a:solidFill><a:srgbClr val="{版面.C_BG}"/></a:solidFill>
<a:ln w="25400"><a:noFill/></a:ln></p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr sz="405"/></a:p></p:txBody>
</p:sp>''')
    sid += 1

    # 左侧章节栏
    shapes.append(txbox(sid, 'RailNo', G['rail_x'], G['rail_no_y'], 900000, 520000,
                        [run(page['no'], 版面.SZ_RAIL_NO * 100, '000000', font=版面.FONT_NUM)]))
    sid += 1
    shapes.append(txbox(sid, 'RailName', G['rail_x'], G['rail_name_y'], G['rail_name_w'], 400000,
                        [run(page['chapter'], 版面.SZ_RAIL_NAME * 100, '000000', font=版面.FONT_NUM)]))
    sid += 1
    shapes.append(txbox(sid, 'RailFoot', G['rail_x'], G['foot_y'], G['rail_name_w'], 240000,
                        [run('青岛大悦城 · 品牌推介发布会', 版面.SZ_FOOT * 100, '9AA0A6')]))
    sid += 1

    # 红色大标题
    shapes.append(txbox(sid, 'Title', G['text_x'], G['title_y'], G['text_w'], G['title_h'],
                        [run(page['title'], 版面.SZ_TITLE * 100, 版面.C_RED, bold=True)]))
    sid += 1

    # 导语（纵向位置按实测行数推算，避免与下方标签重叠）
    lead_y, lead_h, chips_y = 版面.layout_text(page['lead'])
    shapes.append(txbox(sid, 'Lead', G['text_x'], lead_y, G['text_w'], lead_h,
                        [run(page['lead'], 版面.SZ_LEAD * 100, '2B2F36')]))
    sid += 1

    # 关键词标签 / 定义式标语（二选一）
    if page.get('slogan'):
        _, sh = 版面.layout_slogan(page['slogan'], chips_y)
        shapes.append(slogan_box(sid, 'Slogan', G['text_x'], chips_y,
                                 G['text_w'], sh, page['slogan']))
        sid += 1
    else:
        for (x, y, w, kw) in 版面.layout_chips(page['keywords'], chips_y):
            shapes.append(chip(sid, f'Chip{sid}', x, y, w, G['chip_h'], kw))
            sid += 1

    # 图片卡 + 说明条
    boxes = 版面.layout_cards(len(page['images']))
    for i, im in enumerate(page['images']):
        x, y, w, h = boxes[i]
        l, t, r, b = crop_map[im['src']]
        shapes.append(picture(sid, f'Pic{i}', x, y, w, h, rid_map[im['src']], l, t, r, b))
        sid += 1
        shapes.append(caption(sid, f'Cap{i}', x, y, w, G['cap_h'], im['cap_t'], im['cap_s']))
        sid += 1

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst>
<a:gs pos="0"><a:srgbClr val="CDDAF0"><a:alpha val="100000"/></a:srgbClr></a:gs>
<a:gs pos="50000"><a:srgbClr val="E4D6EE"><a:alpha val="100000"/></a:srgbClr></a:gs>
<a:gs pos="100000"><a:srgbClr val="FFDCD7"><a:alpha val="100000"/></a:srgbClr></a:gs>
</a:gsLst><a:lin ang="2700000"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{''.join(shapes)}
</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''


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

    tmp = os.path.join(OUT_DIR, '_build')
    if os.path.isdir(tmp):
        shutil.rmtree(tmp)
    os.makedirs(tmp)
    with zipfile.ZipFile(SRC_PPTX) as z:
        z.extractall(tmp)

    # ---- 图片：压缩 + 写入 media ----
    media_dir = os.path.join(tmp, 'ppt', 'media')
    img_names, crop_map = {}, {}
    next_media = 20
    for p in pages:
        for im in p['images']:
            src = os.path.join(IMG_DIR, im['src'])
            raw = Image.open(src).convert('RGB')
            fname = f'image{next_media}.jpeg'
            next_media += 1
            raw.save(os.path.join(media_dir, fname), 'JPEG', quality=88, optimize=True)
            img_names[im['src']] = fname
            crop_map[im['src']] = 版面.src_rect(raw.width, raw.height,
                                                G['card_w'], G['card_h'])
    print(f'已写入 {len(img_names)} 张图片')

    # ---- 重建 slide9-14 ----
    for idx, page in enumerate(pages):
        n = 9 + idx
        slide_path = os.path.join(tmp, 'ppt', 'slides', f'slide{n}.xml')
        rels_path = os.path.join(tmp, 'ppt', 'slides', '_rels', f'slide{n}.xml.rels')

        rels = open(rels_path, encoding='utf-8').read()
        rels = re.sub(r'<Relationship[^>]*Type="[^"]*/image"[^>]*/>', '', rels)
        add, rid_map = [], {}
        for i, im in enumerate(page['images']):
            rid = f'rIdImg{i}'
            rid_map[im['src']] = rid
            add.append(f'<Relationship Id="{rid}" '
                       f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                       f'Target="../media/{img_names[im["src"]]}"/>')
        rels = rels.replace('</Relationships>', ''.join(add) + '</Relationships>')
        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels)

        with open(slide_path, 'w', encoding='utf-8') as f:
            f.write(build_slide(page, rid_map, crop_map, 10))

    # ---- presentation.xml：只保留第 9-14 页 ----
    pres_path = os.path.join(tmp, 'ppt', 'presentation.xml')
    pres = open(pres_path, encoding='utf-8').read()
    ids = re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', pres)
    keep = ids[8:14]
    pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>',
                  '<p:sldIdLst>' + ''.join(f'<p:sldId id="{a}" r:id="{b}"/>' for a, b in keep)
                  + '</p:sldIdLst>', pres, flags=re.S)
    with open(pres_path, 'w', encoding='utf-8') as f:
        f.write(pres)

    # ---- presentation.xml.rels ----
    prels_path = os.path.join(tmp, 'ppt', '_rels', 'presentation.xml.rels')
    prels = open(prels_path, encoding='utf-8').read()
    keep_rids = {b for _, b in keep}
    for _, rid in ids:
        if rid not in keep_rids:
            prels = re.sub(r'<Relationship Id="%s"[^>]*/>' % rid, '', prels)
    with open(prels_path, 'w', encoding='utf-8') as f:
        f.write(prels)

    # ---- 删除多余页 ----
    removed = list(range(1, 9)) + list(range(15, 18))
    for i in removed:
        for fp in (os.path.join(tmp, 'ppt', 'slides', f'slide{i}.xml'),
                   os.path.join(tmp, 'ppt', 'slides', '_rels', f'slide{i}.xml.rels'),
                   os.path.join(tmp, 'ppt', 'notesSlides', f'notesSlide{i}.xml'),
                   os.path.join(tmp, 'ppt', 'notesSlides', '_rels', f'notesSlide{i}.xml.rels')):
            if os.path.exists(fp):
                os.remove(fp)

    nm_path = os.path.join(tmp, 'ppt', 'notesMasters', '_rels', 'notesMaster1.xml.rels')
    if os.path.exists(nm_path):
        nm = open(nm_path, encoding='utf-8').read()
        for i in removed:
            nm = re.sub(r'<Relationship Id="[^"]*"[^>]*Target="\.\./notesSlides/notesSlide%d\.xml"[^>]*/>' % i,
                        '', nm)
        with open(nm_path, 'w', encoding='utf-8') as f:
            f.write(nm)

    # ---- [Content_Types].xml ----
    ct_path = os.path.join(tmp, '[Content_Types].xml')
    ct = open(ct_path, encoding='utf-8').read()
    for i in removed:
        ct = ct.replace(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>', '')
        ct = ct.replace(f'<Override PartName="/ppt/notesSlides/notesSlide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>', '')
    if 'Extension="jpeg"' not in ct:
        ct = ct.replace('<Default Extension="png"',
                        '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png"')
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct)

    # ---- 打包 ----
    if os.path.exists(OUT_PPTX):
        os.remove(OUT_PPTX)
    with zipfile.ZipFile(OUT_PPTX, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(tmp):
            for fn in files:
                full = os.path.join(root, fn)
                z.write(full, os.path.relpath(full, tmp))
    print(f'已生成 {OUT_PPTX}  ({os.path.getsize(OUT_PPTX)/1024/1024:.2f} MB)')


if __name__ == '__main__':
    main()
