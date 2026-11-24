# -*- coding: utf-8 -*-
"""圓餅圖資料標籤自動排版：避免小扇形的標籤互相重疊。

做法：
1. 每張圓餅圖明確指定繪圖區（圓餅的位置與大小），讓圓心、半徑可以精確計算。
2. 依數值算出每個扇形的角度，估算每個標籤的寬高。
3. 由小扇形到大扇形依序替每個標籤找位置：優先放在自己的扇形裡；
   放不下或會撞到其他標籤時，往內／往外、左右微調；真的放不下才放到圓外。
   不可超出圖表範圍，也不可壓到標題、圖例。
4. 位置以「資料標籤置中（ctr）」為基準，寫成 manualLayout 位移。
"""
import math, json, copy

EMU_PT = 12700


def _find_all(x, n, out):
    if isinstance(x, dict):
        if x.get('n') == n:
            out.append(x)
        for v in x.values():
            _find_all(v, n, out)
    elif isinstance(x, list):
        for v in x:
            _find_all(v, n, out)


def _sizes(x):
    out = []

    def g(y):
        if isinstance(y, dict):
            a = y.get('a')
            if isinstance(a, dict) and 'sz' in a:
                out.append(int(a['sz']))
            for v in y.values():
                g(v)
        elif isinstance(y, list):
            for v in y:
                g(v)
    g(x)
    return out


def _title_text(ch):
    t = []
    _find_all(ch['top'], 'title', t)
    if not t:
        return '', None
    rich = []
    _find_all(t[0], 'rich', rich)
    if rich:
        txt = ''.join(r.get('t', '') for p in rich[0]['tb']['paras'] for r in p['runs'])
    else:
        txt = ch['series'][0].get('tx', {}).get('pts', {}).get('0', '')
    sz = _sizes(t[0])
    return txt, (sz[0] / 100.0 if sz else 18.0)


def _manual(node_parent_name, ch):
    """取得 legend 的 manualLayout（比例）。"""
    out = []
    _find_all(ch['top'], node_parent_name, out)
    if not out:
        return None
    lay = [k for k in out[0]['k'] if k.get('n') == 'layout']
    if not lay:
        return None
    ml = [k for k in lay[0]['k'] if k.get('n') == 'manualLayout']
    if not ml:
        return None
    d = {k['n']: k['a'].get('val') for k in ml[0]['k']}
    return {k: float(d[k]) for k in ('x', 'y', 'w', 'h') if k in d}


def _set_plot_layout(ch, x, y, w, h):
    pa = []
    _find_all(ch['top'], 'plotArea', pa)
    lay = [k for k in pa[0]['k'] if k.get('n') == 'layout']
    node = {'n': 'manualLayout', 'pre': 'c', 'a': {}, 'k': [
        {'n': 'layoutTarget', 'pre': 'c', 'a': {'val': 'inner'}, 'k': []},
        {'n': 'xMode', 'pre': 'c', 'a': {'val': 'edge'}, 'k': []},
        {'n': 'yMode', 'pre': 'c', 'a': {'val': 'edge'}, 'k': []},
        {'n': 'x', 'pre': 'c', 'a': {'val': '%.6f' % x}, 'k': []},
        {'n': 'y', 'pre': 'c', 'a': {'val': '%.6f' % y}, 'k': []},
        {'n': 'w', 'pre': 'c', 'a': {'val': '%.6f' % w}, 'k': []},
        {'n': 'h', 'pre': 'c', 'a': {'val': '%.6f' % h}, 'k': []}]}
    if lay:
        lay[0]['k'] = [node]
    else:
        pa[0]['k'].insert(0, {'n': 'layout', 'pre': 'c', 'a': {}, 'k': [node]})


def _char_w(c):
    if c.isdigit():
        return 0.56
    return {';': 0.30, ' ': 0.28, '%': 0.88, '.': 0.28}.get(c, 1.0)


def _label_box(text, pt):
    w = sum(_char_w(c) for c in text) * pt * EMU_PT + 2 * 38100 + 0.25 * pt * EMU_PT
    h = 1.32 * pt * EMU_PT + 2 * 19050
    return w, h


def _round_half_up(x):
    return int(math.floor(x + 0.5))


def _overlap(a, b, pad=0):
    return not (a[2] + pad <= b[0] or b[2] + pad <= a[0] or a[3] + pad <= b[1] or b[3] + pad <= a[1])


def layout_chart(ch, fw, fh):
    """回傳 (結果說明, 是否全部放得下)。fw/fh：圖表外框大小（EMU）。"""
    se = ch['series'][0]
    n = se['val']['n']
    vals = [float(se['val']['pts'].get(str(i), 0)) for i in range(n)]
    total = sum(vals)
    dls = [k for k in se['k'] if k['n'] == 'dLbls']
    if not dls or total <= 0:
        return 'no-labels', True
    dl = dls[0]
    flags = {x['n']: x['a'].get('val') for x in dl['k'] if x['n'].startswith('show')}
    show_pct = flags.get('showPercent') == '1'
    tx = [x for x in dl['k'] if x['n'] == 'txPr']
    szs = _sizes(tx[0]) if tx else []
    pt = (szs[0] / 100.0) if szs else 18.0

    # --- 標題、圖例、繪圖區 ---
    ttxt, tpt = _title_text(ch)
    tw = max(len(ttxt), 1) * tpt * EMU_PT * 1.0 + 0.6 * tpt * EMU_PT
    th = 1.45 * tpt * EMU_PT
    ty0 = 0.015 * fh
    title_box = ((fw - tw) / 2, ty0, (fw + tw) / 2, ty0 + th)
    leg = _manual('legend', ch)
    pl = _manual('plotArea', ch)
    if pl and 'w' in pl:          # 原本就有指定繪圖區（例如行業別）：沿用
        D = min(pl['w'] * fw, pl['h'] * fh)
        cx, cy = (pl['x'] + pl['w'] / 2) * fw, (pl['y'] + pl['h'] / 2) * fh
    else:
        band_top = title_box[3] + 0.02 * fh
        band_bot = (leg['y'] * fh - 0.02 * fh) if leg else 0.97 * fh
        D = min(0.92 * fw, band_bot - band_top) * 0.97
        cx, cy = fw / 2.0, (band_top + band_bot) / 2.0
        _set_plot_layout(ch, (cx - D / 2) / fw, (cy - D / 2) / fh, D / fw, D / fh)
    R = D / 2.0
    if leg:
        legend_box = (leg['x'] * fw, leg['y'] * fh, (leg['x'] + leg['w']) * fw, (leg['y'] + leg['h']) * fh)
    else:                          # 自動圖例在下方：圓餅下緣以下都當作圖例區
        legend_box = (0, cy + R + 0.015 * fh, fw, fh)
    obstacles = [title_box, legend_box]
    pad = 0.010 * fh
    def attempt(pt):
        ang, a0 = [], 0.0
        for v in vals:
            span = 2 * math.pi * v / total
            ang.append((a0, a0 + span))
            a0 += span
        labels = []
        for i, v in enumerate(vals):
            if v <= 0:
                continue
            txt = ('%d; %d%%' % (v, _round_half_up(v / total * 100))) if show_pct else '%d' % v
            w, h = _label_box(txt, pt)
            labels.append((i, v, w, h))
        placed, boxes = {}, []
        for i, v, w, h in sorted(labels, key=lambda t: t[1]):      # 小扇形先排
            s0, s1 = ang[i]
            mid, span = (s0 + s1) / 2, s1 - s0
            half = math.hypot(w, h) / 2
            cands = []
            for rf in (0.66, 0.56, 0.76, 0.46, 0.84, 0.36, 0.26):
                for df in (0, -0.22, 0.22, -0.4, 0.4):
                    cands.append((abs(rf - 0.66) + abs(df) * 0.9, mid + df * span, rf * R))
            for k, dth in enumerate((0, -0.12, 0.12, -0.25, 0.25, -0.4, 0.4, -0.6, 0.6)):
                cands.append((4 + k * 0.2, mid + dth, R + half * 0.85 + 0.04 * R))
                cands.append((5 + k * 0.2, mid + dth, R + half * 0.85 + 0.22 * R))
            def crosses(b):
                # 標籤框是否跨過扇形的分隔線（只看在圓內的部分）
                if span >= 2 * math.pi - 1e-9:
                    return False
                for px in (b[0], (b[0] + b[2]) / 2, b[2]):
                    for py in (b[1], (b[1] + b[3]) / 2, b[3]):
                        dx, dy = px - cx, cy - py
                        if math.hypot(dx, dy) > R:
                            continue
                        a_ = math.atan2(dx, dy) % (2 * math.pi)
                        rel = (a_ - s0) % (2 * math.pi)
                        if rel > span:
                            return True
                return False
            scored = []
            for score, th_, r in cands:
                x, y = cx + r * math.sin(th_), cy - r * math.cos(th_)
                b = (x - w / 2, y - h / 2, x + w / 2, y + h / 2)
                if r < R and crosses(b):
                    score += 2.5
                scored.append((score, th_, r))
            best = None
            for score, th_, r in sorted(scored, key=lambda c: c[0]):
                x, y = cx + r * math.sin(th_), cy - r * math.cos(th_)
                b = (x - w / 2, y - h / 2, x + w / 2, y + h / 2)
                if b[0] < 0 or b[1] < 0 or b[2] > fw or b[3] > fh:
                    continue
                if any(_overlap(b, o, pad * 0.5) for o in obstacles):
                    continue
                if any(_overlap(b, o, pad) for o in boxes):
                    continue
                best = ('m', th_, r, b)
                break
            if best is None:
                return None
            placed[i] = best
            boxes.append(best[3])
        return placed

    used_pt = pt
    placed = attempt(pt)
    while placed is None and used_pt > pt * 0.75:
        used_pt = round(used_pt * 0.9, 1)
        placed = attempt(used_pt)
    if placed is None:
        return 'FAILED', False
    if used_pt != pt:
        for t in tx:
            def g(y):
                if isinstance(y, dict):
                    a = y.get('a')
                    if isinstance(a, dict) and 'sz' in a:
                        a['sz'] = str(int(used_pt * 100))
                    for vv in y.values():
                        g(vv)
                elif isinstance(y, list):
                    for vv in y:
                        g(vv)
            g(t)

    # --- 寫入：逐點 dLbl（以「置中」為基準的位移；位移單位＝圖表寬高的比例）---
    flag_nodes = [x for x in dl['k'] if x['n'] in ('showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize')]
    newl = []
    acc = 0.0
    for i, v in enumerate(vals):
        idx = {'n': 'idx', 'pre': 'c', 'a': {'val': str(i)}, 'k': []}
        mid = (acc + v / 2) / total * 2 * math.pi
        acc += v
        if v <= 0:
            newl.append({'n': 'dLbl', 'pre': 'c', 'a': {}, 'k': [idx, {'n': 'delete', 'pre': 'c', 'a': {'val': '1'}, 'k': []}]})
            continue
        _, th_, r, b = placed[i]
        bx, by = cx + 0.5 * R * math.sin(mid), cy - 0.5 * R * math.cos(mid)
        x, y = cx + r * math.sin(th_), cy - r * math.cos(th_)
        ox, oy = (x - bx) / fw, (y - by) / fh
        k = [idx]
        if abs(ox) > 1e-5 or abs(oy) > 1e-5:
            k.append({'n': 'layout', 'pre': 'c', 'a': {}, 'k': [{'n': 'manualLayout', 'pre': 'c', 'a': {}, 'k': [
                {'n': 'x', 'pre': 'c', 'a': {'val': '%.6f' % ox}, 'k': []},
                {'n': 'y', 'pre': 'c', 'a': {'val': '%.6f' % oy}, 'k': []}]}]})
        k.append({'n': 'dLblPos', 'pre': 'c', 'a': {'val': 'ctr'}, 'k': []})
        k += copy.deepcopy(flag_nodes)
        newl.append({'n': 'dLbl', 'pre': 'c', 'a': {}, 'k': k})
    dl['k'] = newl + [x for x in dl['k'] if x['n'] != 'dLbl']
    for x in dl['k']:
        if x['n'] == 'dLblPos':
            x['a']['val'] = 'ctr'
        if x['n'] == 'showLeaderLines':
            x['a']['val'] = '1'
    outside = sum(1 for i in placed if placed[i][2] > R)
    return {'pt': used_pt, 'orig_pt': pt, 'outside': outside, 'R': R, 'c': (cx, cy), 'boxes': {i: placed[i][3] for i in placed}}, True


def fix_all(S, flat_shapes):
    frames = {}
    for s in S['slides']:
        for sh in flat_shapes(s['cSld']['shapes']):
            if 'chart' in sh:
                frames.setdefault(sh['chart'], set()).add(tuple(sh['xfrm']['ext']))
    report = {}
    for ci, ext in frames.items():
        ch = S['charts'][ci]
        pie = []
        _find_all(ch['top'], 'pieChart', pie)
        if not pie:
            continue
        assert len(ext) == 1, (ci, ext)
        fw, fh = list(ext)[0]
        report[ci] = layout_chart(ch, fw, fh)
    return report
