#!/usr/bin/env python3
"""P9 文案候选清单 —— 用真实版面实测字数/行数/宽高，只保留「能直接上页面」的方案。

判定口径（与 版面.py 一致）：
  · 标题正文二选一：单行标题若能在正文栏 2000000 EMU 内排下（≤ 14 字），
    正文可用空间从 lead_y 到画布底还有约 166 万 EMU → 可排 4 行导语 + 1 行关键词
  · 若沿用「XX・XX」八字体例，正文只有 3 行导语 + 2 行关键词
  · 关键词：4 个，chip = 文字实测宽 + 2×91440，两列合计不得超 2000000（否则被压缩字号）

用法：python3 P9候选.py
"""
import 版面
from 版面 import G, SZ_LEAD, SZ_CHIP, EMU_H, wrap_lines, text_w_emu

LEAD_TOP = G['title_y'] + G['title_h'] + 40000


def space_rows():
    """正文栏从 lead_y 到底边，按 9.5pt/1.7 行高能排几行标签（含 4 行导语的情况）。"""
    line_h = int(SZ_LEAD * G['lead_lh'] * 12700)
    rows = (EMU_H - LEAD_TOP) // line_h
    return line_h, rows


def report(title, lead, kws):
    ok, notes = True, []
    tw = text_w_emu(title, 版面.SZ_TITLE)
    t_lines = 1 if tw <= G['text_w'] else 2
    if t_lines > 1:
        ok = False
        notes.append(f'标题超宽 {tw}>{G["text_w"]}')

    line_h, _ = space_rows()
    n = wrap_lines(lead, SZ_LEAD, G['text_w'])
    lead_h = n * line_h
    chips_y = LEAD_TOP + lead_h + G['lead_gap']
    row_h = G['chip_h'] + G['chip_gap_y']
    need = chips_y + row_h * (len(kws) + 1) // 2 - G['chip_gap_y']
    if need > EMU_H:
        ok = False
        notes.append(f'关键词底 {need} 超过画布 {EMU_H}')
    if n > 4:
        ok = False
        notes.append(f'导语 {n} 行（上限 4）')

    ws = [版面.chip_width(k) for k in kws]
    half = (len(kws) + 1) // 2
    pair = max(ws[:half]) + max(ws[half:]) + G['chip_gap_x']
    if pair > G['text_w']:
        ok = False
        notes.append(f'两列 chip 合计 {pair} > {G["text_w"]}，会被压缩字号')

    print(f'[{"OK " if ok else "NG "}] {title}  ({len(title)}字, 宽{tw})')
    print(f'      {lead}   ->  {n} 行')
    print(f'      关键词：{" | ".join(f"{k}({ws[i]})" for i, k in enumerate(kws))}  合计{pair}')
    if notes:
        print('      ⚠ ' + '；'.join(notes))
    print()
    return ok


if __name__ == '__main__':
    print(f'正文栏：x={G["text_x"]} 宽={G["text_w"]}；标题上限 14 字(单行)\n')
    print('=' * 72)
    print('A 组｜主推：单行标题 + 4 行导语（信息量最大）')
    print('=' * 72)
    report('国际潮头・新品共展',
           '一小时航程接住首尔、东京的当季潮流。我们招的不是品牌，是首发权，是早一季的时间差。',
           ['首发权', '早一季红利', '区域首店', '限量发售'])

    print('=' * 72)
    print('B 组｜沿用八字体例（标题对仗最稳，导语 3 行）')
    print('=' * 72)
    report('全球潮头・新品共展',
           '当季新品在这里比同线城市早落地一季。不招品牌，只招「首发权」与时间差。',
           ['首发权', '时间差红利', '区域首店', '限量发售'])

    print('=' * 72)
    print('C 组｜招商意图最直接')
    print('=' * 72)
    report('首发即主场',
           '当季新品在这里比同线城市早落地一季。不招品牌，只招「首发权」与时间差。',
           ['全球 / 区域首店', '当季首发', '限量发售', '品牌发声'])
