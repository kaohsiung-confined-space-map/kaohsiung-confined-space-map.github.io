"""佈景主題、母片、版面配置、表格樣式、簡報層級設定。"""
from dml import *
from slide import *

P = '{%s}' % NS['p']


# ---------------- 佈景主題 ----------------
def parse_theme(root):
    te = root.find(A + 'themeElements')
    cs = te.find(A + 'clrScheme')
    fs = te.find(A + 'fontScheme')
    fm = te.find(A + 'fmtScheme')

    def fonts(el):
        d = {'base': [], 'scripts': []}
        for c in kids(el):
            if ln_(c) == 'font':
                d['scripts'].append([c.get('script'), c.get('typeface')])
            else:
                d['base'].append([ln_(c), attrs_of(c)])
        return d

    def effect_style(es):
        d = {'fx': None, 'scene3d': None, 'sp3d': None}
        for c in kids(es):
            n = ln_(c)
            if n == 'effectLst':
                d['fx'] = parse_effects(c)
            elif n == 'scene3d':
                d['scene3d'] = parse_scene3d(c)
            elif n == 'sp3d':
                d['sp3d'] = parse_sp3d(c)
            else:
                unknown(c, 'effectStyle')
        return d

    fam = root.find('.//' + q('thm15:themeFamily'))
    return {
        'name': root.get('name'),
        'colors': {'name': cs.get('name'), 'list': [[ln_(c), find_color(c)] for c in kids(cs)]},
        'fonts': {'name': fs.get('name'), 'major': fonts(fs.find(A + 'majorFont')), 'minor': fonts(fs.find(A + 'minorFont'))},
        'fmt': {'name': fm.get('name'),
                'fill': [parse_fill(c) for c in kids(fm.find(A + 'fillStyleLst'))],
                'ln': [parse_line(c) for c in kids(fm.find(A + 'lnStyleLst'))],
                'effect': [effect_style(c) for c in kids(fm.find(A + 'effectStyleLst'))],
                'bg': [parse_fill(c) for c in kids(fm.find(A + 'bgFillStyleLst'))]},
        'family': attrs_of(fam) if fam is not None else None,
    }


def emit_theme(d):
    root = etree.Element(q('a:theme'), nsmap={'a': NS['a'], 'r': NS['r']})
    root.set('name', d['name'])
    te = sub(root, 'a:themeElements')
    cs = sub(te, 'a:clrScheme', {'name': d['colors']['name']})
    for n, c in d['colors']['list']:
        sub(cs, 'a:' + n).append(emit_color(c))
    fs = sub(te, 'a:fontScheme', {'name': d['fonts']['name']})
    for k in ('major', 'minor'):
        f = sub(fs, 'a:%sFont' % k)
        for n, a in d['fonts'][k]['base']:
            sub(f, 'a:' + n, a)
        for s, t in d['fonts'][k]['scripts']:
            sub(f, 'a:font', {'script': s, 'typeface': t})
    fm = sub(te, 'a:fmtScheme', {'name': d['fmt']['name']})
    l = sub(fm, 'a:fillStyleLst')
    for x in d['fmt']['fill']:
        l.append(emit_fill(x))
    l = sub(fm, 'a:lnStyleLst')
    for x in d['fmt']['ln']:
        l.append(emit_line(x))
    l = sub(fm, 'a:effectStyleLst')
    for x in d['fmt']['effect']:
        es = sub(l, 'a:effectStyle')
        es.append(emit_effects(x['fx'] or []))
        if x['scene3d']:
            es.append(emit_scene3d(x['scene3d']))
        if x['sp3d']:
            es.append(emit_sp3d(x['sp3d']))
    l = sub(fm, 'a:bgFillStyleLst')
    for x in d['fmt']['bg']:
        l.append(emit_fill(x))
    sub(root, 'a:objectDefaults')
    sub(root, 'a:extraClrSchemeLst')
    if d['family']:
        ext = sub(sub(root, 'a:extLst'), 'a:ext', {'uri': '{05A4C25C-085E-4340-85A3-A5531E510DB2}'})
        sub(ext, 'thm15:themeFamily', d['family'])
    return root


# ---------------- 形狀樹（母片／版面配置／投影片共用） ----------------
def parse_cSld(cs, ctx, charts):
    d = {'name': cs.get('name'), 'bg': parse_bg(cs.find(P + 'bg')), 'shapes': []}
    tree = cs.find(P + 'spTree')
    for el in kids(tree):
        if ln_(el) in ('nvGrpSpPr', 'grpSpPr'):
            continue
        d['shapes'].append(parse_shape(el, ctx, charts))
    for c in kids(cs):
        if ln_(c) not in ('bg', 'spTree', 'extLst'):
            unknown(c, 'cSld')
    return d


def emit_cSld(d, ctx, key, build_chart, add_media):
    cs = E('p:cSld', {'name': d['name']})
    if d['bg']:
        cs.append(emit_bg(d['bg']))
    tree = sub(cs, 'p:spTree')
    nv = sub(tree, 'p:nvGrpSpPr')
    sub(nv, 'p:cNvPr', {'id': 1, 'name': ''})
    sub(nv, 'p:cNvGrpSpPr')
    sub(nv, 'p:nvPr')
    gp = sub(tree, 'p:grpSpPr')
    gp.append(emit_xfrm({'a': {}, 'off': [0, 0], 'ext': [0, 0], 'chOff': [0, 0], 'chExt': [0, 0]}))
    for i, s in enumerate(d['shapes']):
        tree.append(emit_shape(s, ctx, (key, s['cnv']['name'], i), build_chart, add_media))
    ext = sub(sub(cs, 'p:extLst'), 'p:ext', {'uri': '{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}'})
    sub(ext, 'p14:creationId', {'val': stable_int('sld', key)})
    return cs


def parse_clrmap_ovr(el):
    if el is None:
        return None
    c = kids(el)[0]
    return ['override', attrs_of(c)] if ln_(c) == 'overrideClrMapping' else ['master', None]


def emit_clrmap_ovr(d):
    el = E('p:clrMapOvr')
    if d[0] == 'master':
        sub(el, 'a:masterClrMapping')
    else:
        sub(el, 'a:overrideClrMapping', d[1])
    return el


# ---------------- 母片 ----------------
def parse_master(root, ctx):
    tx = root.find(P + 'txStyles')
    return {'cSld': parse_cSld(root.find(P + 'cSld'), ctx, None),
            'clrMap': attrs_of(root.find(P + 'clrMap')),
            'styles': [[ln_(s), parse_lststyle(s, ctx)] for s in kids(tx)],
            'a': attrs_of(root)}


def emit_master(d, ctx, n_layouts, add_media):
    root = etree.Element(q('p:sldMaster'), nsmap={'a': NS['a'], 'r': NS['r'], 'p': NS['p']})
    for k, v in d['a'].items():
        root.set(k, v)
    root.append(emit_cSld(d['cSld'], ctx, 'master', None, add_media))
    sub(root, 'p:clrMap', d['clrMap'])
    lst = sub(root, 'p:sldLayoutIdLst')
    for i in range(n_layouts):
        rid = ctx.rel(RT + 'slideLayout', '../slideLayouts/slideLayout%d.xml' % (i + 1))
        sub(lst, 'p:sldLayoutId', {'id': 2147483697 + i, 'r:id': rid})
    ctx.rel(RT + 'theme', '../theme/theme1.xml')
    tx = sub(root, 'p:txStyles')
    for n, lst_ in d['styles']:
        tx.append(emit_lststyle(lst_, ctx, 'p:' + n))
    return root


def parse_layout(root, ctx):
    return {'a': attrs_of(root), 'cSld': parse_cSld(root.find(P + 'cSld'), ctx, None),
            'clr': parse_clrmap_ovr(root.find(P + 'clrMapOvr'))}


def emit_layout(d, ctx, i, add_media):
    root = etree.Element(q('p:sldLayout'), nsmap={'a': NS['a'], 'r': NS['r'], 'p': NS['p']})
    for k, v in d['a'].items():
        root.set(k, v)
    root.append(emit_cSld(d['cSld'], ctx, ('layout', i), None, add_media))
    if d['clr']:
        root.append(emit_clrmap_ovr(d['clr']))
    ctx.rel(RT + 'slideMaster', '../slideMasters/slideMaster1.xml')
    return root


# ---------------- 表格樣式（通用 DrawingML 樹，顏色/線條/填色用語意模型） ----------------
def parse_tree(el):
    n = ln_(el)
    if n in COLOR_TAGS:
        return {'t': 'color', 'c': parse_color(el)}
    if n in FILL_TAGS:
        return {'t': 'fill', 'f': parse_fill(el)}
    if n == 'ln':
        return {'t': 'ln', 'l': parse_line(el)}
    return {'t': 'node', 'n': n, 'a': attrs_of(el), 'k': [parse_tree(c) for c in kids(el)]}


def emit_tree(d):
    if d['t'] == 'color':
        return emit_color(d['c'])
    if d['t'] == 'fill':
        return emit_fill(d['f'])
    if d['t'] == 'ln':
        return emit_line(d['l'])
    el = E('a:' + d['n'], d['a'])
    for k in d['k']:
        el.append(emit_tree(k))
    return el
