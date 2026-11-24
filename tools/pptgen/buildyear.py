# -*- coding: utf-8 -*-
"""以 v1.5 簡報規格為底，套入 115 年統計（來自獨立統計程式），輸出 spec_115.json。
用法：python3 build115.py <indep115.json> <companies_sorted.json> <out_spec.json>
"""
import json, sys, copy
import os, gzip
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from lib115 import T, flat_shapes, related

D = json.load(open(sys.argv[1], encoding='utf-8'))
COMPANIES = json.load(open(sys.argv[2], encoding='utf-8'))
OUT = sys.argv[3]
YEAR = sys.argv[4]          # 民國年（封面用）
_BASE = gzip.open(os.path.join(HERE, 'spec_base.json.gz'), 'rt', encoding='utf-8').read()
S = json.loads(_BASE)
REF = json.loads(_BASE)
REF_CHART = {'場次分布': 144, '違規公司分布': 145, '假日違規': 146, '非假日違規': 147, '行業別': 148, '局限空間地點': 149,
             '局限空間類型': 149, '是否偏遠': 149}
SL = lambda i: S['slides'][i - 1]          # 舊編號
LOG = []

# ---------- 共用 ----------
IDENT = [max(sh['cnv'].get('ident') or 0 for s in S['slides'] for sh in flat_shapes(s['cSld']['shapes'])) + 1000]


def new_ident():
    IDENT[0] += 1
    return IDENT[0]


def next_id(slide):
    return str(max(int(sh['cnv']['id']) for sh in flat_shapes(slide['cSld']['shapes'])) + 1)


def charts_on(slide):
    return [sh for sh in flat_shapes(slide['cSld']['shapes']) if 'chart' in sh]


def set_runs_text(sh, text):
    """把文字框改成單一段落文字（保留第一個 run 的格式）。"""
    p = sh['tx']['paras'][0]
    r0 = p['runs'][0]
    r0['t'] = text
    p['runs'] = [r0]
    sh['tx']['paras'] = [p]


def set_total(sh, n):
    runs = sh['tx']['paras'][0]['runs']
    assert [r['t'] for r in runs][0] == '總場次：' and runs[-1]['t'] == '場', [r['t'] for r in runs]
    assert len(runs) == 3
    runs[1]['t'] = '%d ' % n


# ---------- 亮起顏色（橘＝有資料，灰白＝無資料） ----------
def sc(val, *mods):
    return {'t': 'schemeClr', 'a': {'val': val}, 'mods': [list(m) for m in mods]}


PAL = {
    'path': {'orange': [sc('accent4'), sc('accent2')],
             'grey': [sc('bg1', ('lumMod', '95000')), sc('bg1', ('lumMod', '50000'))]},
    'btn': {'orange': [sc('accent4'), sc('accent2')],
            'grey': [sc('bg1', ('lumMod', '85000')), sc('bg1', ('lumMod', '65000'))]},
}


def recolor(sh, kind, to):
    f = sh['sppr'].get('fill')
    if not f or f.get('t') != 'grad':
        return False                     # 未亮起（例如畫面外的按鈕）不處理
    frm = 'grey' if to == 'orange' else 'orange'
    cur = [c for _, c in f['stops']]
    if cur == PAL[kind][to]:
        return True
    assert cur == PAL[kind][frm], (kind, json.dumps(cur, ensure_ascii=False))
    for st, c in zip(f['stops'], PAL[kind][to]):
        st[1] = copy.deepcopy(c)
    return True


def restyle(slide, X, to):
    r = related(slide, X)
    done = []
    for kind in ('path', 'btn'):
        for sh in r[kind]:
            if recolor(sh, kind, to):
                done.append(kind)
    return done


# ---------- 圖表數據 ----------
def chart_role(ci):
    se = S['charts'][ci]['series']
    if len(se) == 17:
        return 'items'
    s = se[0]
    name = s.get('tx', {}).get('pts', {}).get('0', '')
    cats = list(s['cat']['pts'].values()) if 'cat' in s else []
    if '儲槽' in cats:
        return '局限空間地點'
    if '製造業' in cats:
        return '行業別'
    if '五項以上' in cats:
        return '非假日違規' if name.startswith('非假日') else '假日違規'
    if cats[:2] == ['假日', '非假日']:
        return '違規公司分布' if '違規公司' in name else '場次分布'
    if '缺氧作業' in cats:
        return '局限空間類型'
    if '偏遠' in cats:
        return '是否偏遠'
    raise ValueError((ci, name, cats))


def find_kid(node, n):
    return [k for k in node['k'] if k.get('n') == n]


def set_series_data(ci, mapping):
    """mapping: OrderedDict 類別 -> 數值；依類別重寫 cat/val、dPt、資料標籤。"""
    ch = S['charts'][ci]
    s = ch['series'][0]
    cats = list(mapping.keys())
    vals = [mapping[c] for c in cats]
    n = len(cats)
    old_n = s['val']['n']
    s['cat']['n'] = n
    s['cat']['pts'] = {str(i): c for i, c in enumerate(cats)}
    s['val']['n'] = n
    s['val']['pts'] = {str(i): str(v) for i, v in enumerate(vals)}
    # dPt：數量要與類別數一致
    dpts = [k for k in s['k'] if k['n'] == 'dPt']
    if n > old_n:
        assert n == 3 and old_n == 2, (ci, n, old_n)
        extra = copy.deepcopy(dpts[1])
        [x for x in extra['k'] if x['n'] == 'idx'][0]['a']['val'] = '2'
        [x for x in extra['k'] if x['n'] == 'spPr'][0]['sp']['fill'] = {'t': 'solid', 'c': sc('accent3')}
        for x in extra['k']:
            if x['n'] == 'extLst':
                x['exts'] = [{'uri': e['uri'], 'uid': True} for e in x['exts']]
        pos = s['k'].index(dpts[-1]) + 1
        s['k'].insert(pos, extra)
    elif n < old_n:
        for d in dpts:
            if int([x for x in d['k'] if x['n'] == 'idx'][0]['a']['val']) >= n:
                s['k'].remove(d)
    # 資料標籤：拿掉逐點手動位置，數值為 0 的點不顯示標籤
    dl = [k for k in s['k'] if k['n'] == 'dLbls']
    if dl and not any(k['n'].startswith('show') for k in dl[0]['k']):
        # 原圖整個數列的標籤被關掉（只靠逐點標籤顯示）→ 換成同類圖表的標準標籤設定
        ref = [k for k in REF['charts'][REF_CHART[chart_role(ci)]]['series'][0]['k'] if k['n'] == 'dLbls'][0]
        LOG.append('圖表 %d（%s）原本整列標籤關閉，改用標準標籤設定' % (ci, chart_role(ci)))
        s['k'][s['k'].index(dl[0])] = dl[0] = copy.deepcopy(ref)
    if dl:
        dl = dl[0]
        dl['k'] = [k for k in dl['k'] if k['n'] != 'dLbl']
        zeros = [i for i, v in enumerate(vals) if v == 0]
        dl['k'][0:0] = [{'n': 'dLbl', 'pre': 'c', 'a': {}, 'k': [
            {'n': 'idx', 'pre': 'c', 'a': {'val': str(i)}, 'k': []},
            {'n': 'delete', 'pre': 'c', 'a': {'val': '1'}, 'k': []}]} for i in zeros]


def chart_values(ci):
    s = S['charts'][ci]['series']
    if len(s) == 17:
        return {x['tx']['pts']['0']: int(x['val']['pts']['0']) for x in s}
    return {s[0]['cat']['pts'][str(i)]: int(s[0]['val']['pts'][str(i)]) for i in range(s[0]['val']['n'])}


DECK_CATS = {'局限空間地點': ['儲槽', '機械設備內部', '人孔', '管道與地下空間', '密閉空間', '其他']}
RENAME = {'局限空間地點': {'密閉空間': '蓄水池或化糞池'}}
FIX3 = ['假日', '非假日', '無違規']


def fill_chart(ci, charts_data):
    role = chart_role(ci)
    want = charts_data[role]
    s = S['charts'][ci]['series'][0]
    deck = [s['cat']['pts'][str(i)] for i in range(s['cat']['n'])]
    deck = [RENAME.get(role, {}).get(c, c) if c not in want else c for c in deck]
    if role == '違規公司分布':
        deck = FIX3
    assert set(want) <= set(deck), (ci, role, deck, list(want))
    set_series_data(ci, {c: want.get(c, 0) for c in deck})
    return role


def clone_chart(ci):
    S['charts'].append(copy.deepcopy(S['charts'][ci]))
    return len(S['charts']) - 1


def set_chart_title(ci, text):
    """把圖表標題（rich text）與數列名稱改掉。"""
    hit = [0]

    def walk(x, in_title=False):
        if isinstance(x, dict):
            t = in_title or x.get('n') == 'title'
            if t and 'tb' in x and x.get('n') == 'rich':
                paras = x['tb']['paras']
                runs = [r for p in paras for r in p['runs'] if 't' in r]
                runs[0]['t'] = text
                for r in runs[1:]:
                    r['t'] = ''
                hit[0] += 1
            for v in x.values():
                walk(v, t)
        elif isinstance(x, list):
            for v in x:
                walk(v, in_title)
    walk(S['charts'][ci]['top'])
    assert hit[0] == 1, hit
    S['charts'][ci]['series'][0]['tx']['pts'] = {'0': text}


# ---------- 「無資料」小字（圓餅圖沒有資料時） ----------
NODATA_P1 = [sh for sh in flat_shapes(SL(66)['cSld']['shapes']) if T(sh) == '無資料'][0]   # 第一頁、假日違規位置
NODATA_P2 = [sh for sh in flat_shapes(SL(67)['cSld']['shapes']) if T(sh) == '無資料'][0]   # 第二頁（畫面外上方）
BIG_NODATA = [sh for sh in SL(4)['cSld']['shapes'] if T(sh) == '無資料'][0]
EMU = 100000


def is_small_nodata(sh):
    return sh['k'] == 'sp' and T(sh) == '無資料' and sh['sppr']['xfrm']['ext'][0] < 3000000


def add_nodata_boxes(slide, page, data, idents):
    slide['cSld']['shapes'] = [sh for sh in slide['cSld']['shapes'] if not is_small_nodata(sh)]
    for role, dx in (('假日違規', 0), ('非假日違規', 3600000)):
        if sum(data['charts'][role].values()) == 0:
            box = copy.deepcopy(NODATA_P1 if page == 1 else NODATA_P2)
            box['sppr']['xfrm']['off'][0] += dx
            box['cnv']['id'] = next_id(slide)
            box['cnv']['name'] = '無資料 ' + role
            box['cnv']['ident'] = idents[role]
            slide['cSld']['shapes'].append(box)


# ---------- 偏遠地區標籤 ----------
def remote_chip(slide, ident):
    tot = [sh for sh in slide['cSld']['shapes'] if sh['cnv']['name'] == '總場次'][0]
    chip = copy.deepcopy(tot)
    chip['cnv'] = {'id': next_id(slide), 'name': '偏遠地區', 'a': {}, 'links': [], 'ident': ident}
    chip['cnvsp'] = {'a': {}, 'locks': None}
    sp = chip['sppr']
    sp['xfrm'] = {'a': {}, 'off': [1650000, 200000], 'ext': [1300000, 430000]}
    sp['geom'] = {'t': 'prst', 'prst': 'roundRect', 'av': [['adj', 'val 50000']]}
    sp['fill'] = {'t': 'solid', 'c': {'t': 'srgbClr', 'a': {'val': '4E9A3A'}, 'mods': []}}
    chip['tx']['body'] = {'a': {'wrap': 'none', 'lIns': '72000', 'tIns': '0', 'rIns': '72000', 'bIns': '0',
                                'rtlCol': '0', 'anchor': 'ctr'}, 'parts': []}
    p = chip['tx']['paras'][0]
    p['pPr'] = {'tag': 'pPr', 'a': {'algn': 'ctr'}, 'parts': []}
    r0 = p['runs'][0]
    r0['rPr']['a']['sz'] = '1600'
    r0['t'] = '偏遠地區'
    p['runs'] = [r0]
    chip['tx']['paras'] = [p]
    slide['cSld']['shapes'].append(chip)


# ---------- 局限空間類型（每區第二頁）、全區新增圖 ----------
def scale_fonts(x, f):
    if isinstance(x, dict):
        for k, v in x.items():
            if k == 'sz' and isinstance(v, str) and v.isdigit():
                x[k] = str(int(round(int(v) * f / 100.0)) * 100)
            else:
                scale_fonts(v, f)
    elif isinstance(x, list):
        for v in x:
            scale_fonts(v, f)


def add_pie(slide, src_frame, role, title, off, ext, ident, data, font=1.0, src_chart=None):
    ci = clone_chart(src_chart if src_chart is not None else src_frame['chart'])
    set_chart_title(ci, title)
    def strip_layout(x):
        if isinstance(x, dict):
            if x.get('n') == 'layout':
                x['k'] = []
            for v in x.values():
                strip_layout(v)
        elif isinstance(x, list):
            for v in x:
                strip_layout(v)
    if src_chart is not None:                      # 全區頁的小圓餅：重新指定繪圖區與圖例位置
        def ml(x, y, w, h, inner=False):
            k = ([{'n': 'layoutTarget', 'pre': 'c', 'a': {'val': 'inner'}, 'k': []}] if inner else []) + [
                {'n': 'xMode', 'pre': 'c', 'a': {'val': 'edge'}, 'k': []}, {'n': 'yMode', 'pre': 'c', 'a': {'val': 'edge'}, 'k': []},
                {'n': 'x', 'pre': 'c', 'a': {'val': str(x)}, 'k': []}, {'n': 'y', 'pre': 'c', 'a': {'val': str(y)}, 'k': []},
                {'n': 'w', 'pre': 'c', 'a': {'val': str(w)}, 'k': []}, {'n': 'h', 'pre': 'c', 'a': {'val': str(h)}, 'k': []}]
            return [{'n': 'manualLayout', 'pre': 'c', 'a': {}, 'k': k}]

        def setl(x, parent):
            if isinstance(x, dict):
                if x.get('n') == 'layout' and parent == 'plotArea':
                    x['k'] = ml(0.25, 0.2, 0.5, 0.62, True)
                elif x.get('n') == 'layout' and parent == 'legend':
                    x['k'] = ml(0.0, 0.86, 1.0, 0.14)
                for v in x.get('k', []) if isinstance(x.get('k'), list) else []:
                    setl(v, x.get('n'))
                for kk, v in x.items():
                    if kk != 'k':
                        setl(v, parent)
            elif isinstance(x, list):
                for v in x:
                    setl(v, parent)
        setl(S['charts'][ci]['top'], None)
    if font != 1.0:
        scale_fonts(S['charts'][ci]['top'], font)
        scale_fonts(S['charts'][ci]['series'], font)
    set_series_data(ci, data)
    fr = copy.deepcopy(src_frame)
    fr['chart'] = ci
    fr['cnv'] = {'id': next_id(slide), 'name': title, 'a': {}, 'links': [], 'ident': ident}
    fr['xfrm']['off'] = list(off)
    fr['xfrm']['ext'] = list(ext)
    slide['cSld']['shapes'].append(fr)
    assert chart_role(ci) == role
    return ci


def frame_by_role(slide, role):
    hits = [sh for sh in charts_on(slide) if chart_role(sh['chart']) == role]
    assert len(hits) == 1, (role, len(hits))
    return hits[0]


# ---------- 各區頁面 ----------
def update_district_pages(X, p1, p2, data, idents):
    for page, sl in ((1, p1), (2, p2)):
        roles = []
        for fr in charts_on(sl):
            roles.append(fill_chart(fr['chart'], data['charts']))
        assert sorted(roles) == sorted(['場次分布', '違規公司分布', '假日違規', '非假日違規', '局限空間地點', '行業別']), (X, page, roles)
        tots = [sh for sh in sl['cSld']['shapes'] if sh['cnv']['name'] == '總場次']
        assert len(tots) == 1
        set_total(tots[0], data['total'])
        add_nodata_boxes(sl, page, data, idents)
        if data['remote'] is True:
            remote_chip(sl, idents['chip'])
        # 局限空間類型：第二頁在畫面上，第一頁放在畫面外下方（與其他第二頁圖表一致）
        src = frame_by_role(sl, '局限空間地點')
        y = 2450000 if page == 2 else 2450000 + 6100000
        add_pie(sl, src, '局限空間類型', '局限空間類型', (9450000, y), (2700000, 3300000), idents['ctype'],
                data['charts']['局限空間類型'])


def title_text_set(slide, X):
    for sh in slide['cSld']['shapes']:
        if 'tbl' in sh and len(sh['tbl']['rows']) == 1:
            cell = sh['tbl']['rows'][0]['cells'][0]
            runs = [r for p in cell['tx']['paras'] for r in p['runs'] if 't' in r]
            runs[0]['t'] = X
            for r in runs[1:]:
                r['t'] = ''
            return
    raise ValueError


def title_text(slide):
    for sh in slide['cSld']['shapes']:
        if 'tbl' in sh and len(sh['tbl']['rows']) == 1:
            return ''.join(x.get('t', '') for p in sh['tbl']['rows'][0]['cells'][0]['tx']['paras'] for x in p['runs'])


# ---------- 群組 ----------
grp = {}
for sh in flat_shapes(SL(2)['cSld']['shapes']):
    if sh['cnv']['name'].startswith('botton'):
        h = [l['tgt'][1] for l in sh['cnv']['links'] if l['tag'] == 'hlinkHover' and l['tgt']]
        grp[T(sh)] = h[0]
hs = sorted(grp.values())
order = sorted(grp, key=lambda x: grp[x])
TPL_P1, TPL_P2, TPL_HOVER = copy.deepcopy(SL(60)), copy.deepcopy(SL(61)), copy.deepcopy(SL(59))     # 有資料的範本：小港區

new_slides = [SL(1), SL(2)]
old_of_new = [1, 2]          # 新投影片對應的舊編號（新增的頁面記 None）

for X in order:
    h = grp[X]
    nxt = min([x for x in hs if x > h] + [107])
    pages = list(range(h + 1, nxt))
    d = D['districts'][X]
    has115 = d['total'] > 0
    had = len(pages) == 2
    hover = SL(h)
    idents = {'假日違規': new_ident(), '非假日違規': new_ident(), 'chip': new_ident(), 'ctype': new_ident()}
    if has115:
        restyle(hover, X, 'orange')
        if had:
            p1, p2 = SL(pages[0]), SL(pages[1])
            if title_text(p2) != X:
                LOG.append('%s 第二頁標題原本是「%s」，已更正' % (X, title_text(p2)))
                title_text_set(p2, X)
            new = [(hover, h), (p1, pages[0]), (p2, pages[1])]
        else:
            # 無資料 → 有資料：用該區自己的「無資料」頁當第一頁底（地圖、按鈕、圖示位置都是該區的）
            nd = SL(pages[0])
            p1 = copy.deepcopy(nd)
            p1['cSld']['shapes'] = [sh for sh in p1['cSld']['shapes'] if T(sh) != '無資料']
            for fr in charts_on(TPL_P1):
                f2 = copy.deepcopy(fr)
                f2['chart'] = clone_chart(fr['chart'])
                f2['cnv']['id'] = next_id(p1)
                p1['cSld']['shapes'].append(f2)
            tot = [sh for sh in p1['cSld']['shapes'] if sh['cnv']['name'] == '總場次'][0]
            tpl_tot = [sh for sh in TPL_P1['cSld']['shapes'] if sh['cnv']['name'] == '總場次'][0]
            tot['cnv']['links'] = copy.deepcopy(tpl_tot['cnv']['links'])
            restyle(p1, X, 'orange')
            # 第二頁：小港區第二頁為底，換成本區的圖示、標題
            p2 = copy.deepcopy(TPL_P2)
            icon = copy.deepcopy(related(p1, X)['path'][0])
            p2['cSld']['shapes'] = [icon if (sh['k'] == 'sp' and sh['cnv']['name'] == '小港區') else sh
                                    for sh in p2['cSld']['shapes']]
            assert related(p2, X)['path'] and not related(p2, '小港區')['path']
            for fr in charts_on(p2):
                fr['chart'] = clone_chart(fr['chart'])
            title_text_set(p2, X)
            # 亮起按鈕加上「點擊→下一頁」
            btn = related(hover, X)['btn'][0]
            tpl_btn = related(TPL_HOVER, '小港區')['btn'][0]
            btn['cnv']['links'] = [l for l in btn['cnv']['links'] if l['tag'] != 'hlinkClick'] + \
                [copy.deepcopy(l) for l in tpl_btn['cnv']['links'] if l['tag'] == 'hlinkClick']
            new = [(hover, h), (p1, None), (p2, None)]
            LOG.append('%s：114 年無資料、115 年有資料，新增兩頁圖表' % X)
        update_district_pages(X, new[1][0], new[2][0], d, idents)
    else:
        restyle(hover, X, 'grey')
        if had:
            # 有資料 → 無資料：用該區第一頁為底，拿掉圖表，換成「無資料」
            nd = copy.deepcopy(SL(pages[0]))
            nd['cSld']['shapes'] = [sh for sh in nd['cSld']['shapes'] if 'chart' not in sh and not is_small_nodata(sh)]
            big = copy.deepcopy(BIG_NODATA)
            big['cnv']['id'] = next_id(nd)
            nd['cSld']['shapes'].append(big)
            tot = [sh for sh in nd['cSld']['shapes'] if sh['cnv']['name'] == '總場次'][0]
            tot['cnv']['links'] = []
            set_total(tot, 0)
            restyle(nd, X, 'grey')
            new = [(hover, h), (nd, None)]
            LOG.append('%s：114 年有資料、115 年無資料，改成「無資料」頁' % X)
        else:
            nd = SL(pages[0])
            tot = [sh for sh in nd['cSld']['shapes'] if sh['cnv']['name'] == '總場次'][0]
            assert ''.join(r['t'] for r in tot['tx']['paras'][0]['runs']).replace(' ', '') == '總場次：0場'
            new = [(hover, h), (nd, pages[0])]
    for sl, o in new:
        new_slides.append(sl)
        old_of_new.append(o)

# ---------- 全區總覽 ----------
SUM = D['summary']
for i in (107, 108):
    sl = SL(i)
    for fr in charts_on(sl):
        fill_chart(fr['chart'], SUM['charts'])
    set_total([sh for sh in sl['cSld']['shapes'] if sh['cnv']['name'] == '總場次'][0], SUM['total'])
sl108 = SL(108)
src = frame_by_role(sl108, '局限空間地點')
dist_src = frame_by_role(TPL_P2, '局限空間地點')['chart']     # 用各區頁面的圖表樣式（版面較適合小尺寸）
if '是否偏遠' in SUM['charts']:
    add_pie(sl108, src, '局限空間類型', '局限空間類型', (8850000, 1700000), (2900000, 2250000), new_ident(), SUM['charts']['局限空間類型'], 0.75, dist_src)
    add_pie(sl108, src, '是否偏遠', '是否偏遠', (8850000, 3950000), (2900000, 2250000), new_ident(), SUM['charts']['是否偏遠'], 0.75, dist_src)
else:   # 115 年以前不統計是否偏遠：只放局限空間類型
    add_pie(sl108, src, '局限空間類型', '局限空間類型', (8850000, 2300000), (2900000, 2700000), new_ident(), SUM['charts']['局限空間類型'], 0.8, dist_src)

# 缺失項目統計（第 111 頁）
items = {it['name']: it['v'] for it in SUM['items']}
sl111 = SL(111)
bar = frame_by_role(sl111, 'items')
for se in S['charts'][bar['chart']]['series']:
    name = se['tx']['pts']['0']
    key = '無' if name == '無/空白' else name
    se['tx']['pts']['0'] = key
    se['val']['pts']['0'] = str(items[key])
tbls = [sh for sh in sl111['cSld']['shapes'] if 'tbl' in sh and len(sh['tbl']['rows']) > 1]
t_items = [t for t in tbls if len(t['tbl']['rows']) == 18][0]['tbl']
t_mm = [t for t in tbls if len(t['tbl']['rows']) == 2][0]['tbl']


def cell_set(cell, text):
    runs = [r for p in cell['tx']['paras'] for r in p['runs'] if 't' in r]
    runs[0]['t'] = text
    for r in runs[1:]:
        r['t'] = ''
    cell['tx']['paras'] = [p for p in cell['tx']['paras'] if any('t' in r for r in p['runs'])][:1]


def cell_text(cell):
    return ''.join(x.get('t', '') for p in cell['tx']['paras'] for x in p['runs'])


for row in t_items['rows'][1:]:
    lab = cell_text(row['cells'][0]).strip()
    key = '無' if lab == '無/空白' else '第%s項' % lab
    if lab == '無/空白':
        cell_set(row['cells'][0], '無')
    cell_set(row['cells'][1], str(items[key]))
num = {k: v for k, v in items.items() if k != '無'}
mx, mn = max(num.values()), min(num.values())
cell_set(t_mm['rows'][0]['cells'][1], '、'.join(k[1:-1] for k, v in num.items() if v == mx))
cell_set(t_mm['rows'][1]['cells'][1], '、'.join(k[1:-1] for k, v in num.items() if v == mn))

# 缺失五項以上公司（第 109、110 頁）：加「日期」欄
PER = 15
chunks = [COMPANIES[:PER], COMPANIES[PER:2 * PER]]
assert len(COMPANIES) <= 2 * PER
for i, chunk in zip((109, 110), chunks):
    sl = SL(i)
    t = [sh for sh in sl['cSld']['shapes'] if 'tbl' in sh and len(sh['tbl']['rows']) > 1][0]['tbl']
    W = sum(t['cols'])
    t['cols'] = [1300000, 3450000, 1550000, W - 1300000 - 3450000 - 1550000]
    head, body_tpl = t['rows'][0], t['rows'][1]
    newhead = copy.deepcopy(head)
    newhead['cells'] = [copy.deepcopy(head['cells'][0])] + newhead['cells']
    cell_set(newhead['cells'][0], '日期')
    cell_set(newhead['cells'][3], '缺失項目（項次）')
    rows = [newhead]
    for comp, n, its, date in chunk:
        r = copy.deepcopy(body_tpl)
        r['cells'] = [copy.deepcopy(body_tpl['cells'][0])] + r['cells']
        for k, (c, v) in enumerate(zip(r['cells'], (date, comp, str(n), its))):
            cell_set(c, v)
            for p in c['tx']['paras']:
                for run in p['runs']:
                    if run.get('rPr'):
                        run['rPr']['a']['sz'] = ('1200' if len(v) > 30 else '1400') if k == 3 else '1600'
        rows.append(r)
    t['rows'] = rows

# 封面年份
for sh in flat_shapes(SL(1)['cSld']['shapes']):
    for p in (sh.get('tx') or {'paras': []})['paras']:
        for r in p['runs']:
            if r.get('t') == '114':
                r['t'] = YEAR
                LOG.append('封面年份 → ' + YEAR)

# ---------- 重新編號、更新連結 ----------
new_of_old = {o: i + 1 for i, o in enumerate(old_of_new) if o}
new_slides += [SL(i) for i in range(107, 113)]
base = len(old_of_new)
for k, i in enumerate(range(107, 113)):
    new_of_old[i] = base + k + 1
bad = []
for si, sl in enumerate(new_slides, 1):
    for sh in flat_shapes(sl['cSld']['shapes']):
        for l in sh['cnv']['links']:
            if l['tgt'] and l['tgt'][0] == 'slide':
                o = l['tgt'][1]
                if o not in new_of_old:
                    bad.append((si, sh['cnv']['name'], o))
                else:
                    l['tgt'] = ['slide', new_of_old[o]]
        for p in (sh.get('tx') or {'paras': []})['paras']:
            for r in p['runs']:
                for l in ((r.get('rPr') or {}).get('links') or []):
                    if l.get('tgt') and l['tgt'][0] == 'slide':
                        l['tgt'] = ['slide', new_of_old[l['tgt'][1]]]
assert not bad, bad
# 新增投影片需要獨立的投影片物件（深拷貝避免共用）
S['slides'] = [copy.deepcopy(s) for s in new_slides]
import labelfix
LREP = labelfix.fix_all(S, flat_shapes)
fails = [ci for ci, (r, ok) in LREP.items() if not ok]
shrunk = [(ci, r['orig_pt'], r['pt']) for ci, (r, ok) in LREP.items() if ok and isinstance(r, dict) and r['pt'] != r['orig_pt']]
outs = sum(r['outside'] for r, ok in LREP.values() if ok and isinstance(r, dict))
print('圓餅圖標籤排版：%d 張，放不下 %d，縮字 %d 張，放到圓外的標籤 %d 個' % (len(LREP), len(fails), len(shrunk), outs))
json.dump({str(k): (v[0] if not isinstance(v[0], dict) else {kk: vv for kk, vv in v[0].items()}) for k, v in LREP.items()}, open(OUT + '.labels.json', 'w'), ensure_ascii=False, default=list)
json.dump(S, open(OUT, 'w'), ensure_ascii=False)
print('slides', len(S['slides']), 'charts', len(S['charts']))
print('\n'.join(LOG))
