"""DrawingML 語意模型：parse_*（XML → dict）與 emit_*（dict → XML）。
未知元素一律丟例外，確保每種用到的格式都有被理解並重建。"""
from lxml import etree

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
    'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'p15': 'http://schemas.microsoft.com/office/powerpoint/2012/main',
    'p159': 'http://schemas.microsoft.com/office/powerpoint/2015/09/main',
    'c14': 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
    'c15': 'http://schemas.microsoft.com/office/drawing/2012/chart',
    'c16': 'http://schemas.microsoft.com/office/drawing/2014/chart',
    'c16r2': 'http://schemas.microsoft.com/office/drawing/2015/06/chart',
    'c16r3': 'http://schemas.microsoft.com/office/drawing/2017/03/chart',
    'thm15': 'http://schemas.microsoft.com/office/thememl/2012/main',
}
A = '{%s}' % NS['a']
R_ID = '{%s}id' % NS['r']


def q(tag):
    pre, local = tag.split(':')
    return '{%s}%s' % (NS[pre], local)


def ln_(el):
    return etree.QName(el).localname


def pfx(el):
    uri = etree.QName(el).namespace
    for k, v in NS.items():
        if v == uri:
            return k
    raise ValueError(uri)


def kids(el):
    return [c for c in el if isinstance(c.tag, str)]


def E(tag, attrs=None, *children):
    nsm = {tag.split(':')[0]: NS[tag.split(':')[0]]}
    for k in (attrs or {}):
        if ':' in k:
            nsm[k.split(':')[0]] = NS[k.split(':')[0]]
    el = etree.Element(q(tag), nsmap=nsm)
    for k, v in (attrs or {}).items():
        if v is None:
            continue
        el.set(q(k) if ':' in k else k, str(v))
    for c in children:
        if c is not None:
            el.append(c)
    return el


def sub(parent, tag, attrs=None):
    el = E(tag, attrs)
    parent.append(el)
    return el


def attrs_of(el):
    out = {}
    for k, v in el.attrib.items():
        qn = etree.QName(k)
        if qn.namespace:
            pre = [p for p, u in NS.items() if u == qn.namespace][0]
            out['%s:%s' % (pre, qn.localname)] = v
        else:
            out[k] = v
    return out


class Unknown(Exception):
    pass


def unknown(el, where):
    raise Unknown('%s: 未處理的元素 %s' % (where, etree.QName(el).text))


# ---------------- 顏色 ----------------
COLOR_TAGS = ('srgbClr', 'schemeClr', 'prstClr', 'sysClr', 'scrgbClr', 'hslClr')


def parse_color(el):
    d = {'t': ln_(el), 'a': attrs_of(el), 'mods': []}
    for m in kids(el):
        d['mods'].append([ln_(m), m.get('val')])
    return d


def emit_color(d):
    el = E('a:' + d['t'], d['a'])
    for name, val in d['mods']:
        sub(el, 'a:' + name, {'val': val} if val is not None else None)
    return el


def find_color(el):
    for c in kids(el):
        if ln_(c) in COLOR_TAGS:
            return parse_color(c)
    return None


# ---------------- 填色 ----------------
FILL_TAGS = ('noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill')


def parse_fill(el):
    t = ln_(el)
    if t == 'noFill':
        return {'t': 'none'}
    if t == 'grpFill':
        return {'t': 'grp'}
    if t == 'solidFill':
        return {'t': 'solid', 'c': find_color(el)}
    if t == 'gradFill':
        d = {'t': 'grad', 'a': attrs_of(el), 'stops': [], 'shade': None, 'tile': None}
        for c in kids(el):
            n = ln_(c)
            if n == 'gsLst':
                for gs in kids(c):
                    d['stops'].append([int(gs.get('pos')), find_color(gs)])
            elif n == 'lin':
                d['shade'] = {'t': 'lin', 'a': attrs_of(c)}
            elif n == 'path':
                f = c.find(A + 'fillToRect')
                d['shade'] = {'t': 'path', 'a': attrs_of(c), 'ftr': attrs_of(f) if f is not None else None}
            elif n == 'tileRect':
                d['tile'] = attrs_of(c)
            else:
                unknown(c, 'gradFill')
        return d
    if t == 'pattFill':
        fg = el.find(A + 'fgClr'); bg = el.find(A + 'bgClr')
        return {'t': 'patt', 'prst': el.get('prst'), 'fg': find_color(fg), 'bg': find_color(bg)}
    unknown(el, 'fill')


def emit_fill(d):
    t = d['t']
    if t == 'none':
        return E('a:noFill')
    if t == 'grp':
        return E('a:grpFill')
    if t == 'solid':
        return E('a:solidFill', None, emit_color(d['c']))
    if t == 'grad':
        el = E('a:gradFill', d['a'])
        gl = sub(el, 'a:gsLst')
        for pos, c in d['stops']:
            gl.append(E('a:gs', {'pos': pos}, emit_color(c)))
        s = d['shade']
        if s:
            if s['t'] == 'lin':
                sub(el, 'a:lin', s['a'])
            else:
                p = sub(el, 'a:path', s['a'])
                if s['ftr'] is not None:
                    sub(p, 'a:fillToRect', s['ftr'])
        if d['tile'] is not None:
            sub(el, 'a:tileRect', d['tile'])
        return el
    if t == 'patt':
        return E('a:pattFill', {'prst': d['prst']},
                 E('a:fgClr', None, emit_color(d['fg'])), E('a:bgClr', None, emit_color(d['bg'])))
    raise ValueError(t)


def take_fill(el):
    for c in kids(el):
        if ln_(c) in FILL_TAGS:
            return parse_fill(c)
    return None


# ---------------- 線條 ----------------
def parse_line(el):
    d = {'a': attrs_of(el), 'fill': None, 'dash': None, 'join': None, 'head': None, 'tail': None}
    for c in kids(el):
        n = ln_(c)
        if n in FILL_TAGS:
            d['fill'] = parse_fill(c)
        elif n == 'prstDash':
            d['dash'] = c.get('val')
        elif n in ('round', 'bevel', 'miter'):
            d['join'] = [n, attrs_of(c)]
        elif n == 'headEnd':
            d['head'] = attrs_of(c)
        elif n == 'tailEnd':
            d['tail'] = attrs_of(c)
        else:
            unknown(c, 'ln')
    return d


def emit_line(d, tag='a:ln'):
    el = E(tag, d['a'])
    if d['fill']:
        el.append(emit_fill(d['fill']))
    if d['dash']:
        sub(el, 'a:prstDash', {'val': d['dash']})
    if d['join']:
        sub(el, 'a:' + d['join'][0], d['join'][1])
    if d['head'] is not None:
        sub(el, 'a:headEnd', d['head'])
    if d['tail'] is not None:
        sub(el, 'a:tailEnd', d['tail'])
    return el


# ---------------- 特效 ----------------
EFFECTS = ('blur', 'fillOverlay', 'glow', 'innerShdw', 'outerShdw', 'prstShdw', 'reflection', 'softEdge')


def parse_effects(el):
    out = []
    for c in kids(el):
        n = ln_(c)
        if n not in EFFECTS:
            unknown(c, 'effectLst')
        out.append({'t': n, 'a': attrs_of(c), 'c': find_color(c)})
    return out


def emit_effects(lst):
    el = E('a:effectLst')
    for e in lst:
        x = sub(el, 'a:' + e['t'], e['a'])
        if e['c']:
            x.append(emit_color(e['c']))
    return el


def parse_scene3d(el):
    d = {}
    for c in kids(el):
        n = ln_(c)
        if n in ('camera', 'lightRig'):
            rot = c.find(A + 'rot')
            d[n] = {'a': attrs_of(c), 'rot': attrs_of(rot) if rot is not None else None}
        else:
            unknown(c, 'scene3d')
    return d


def emit_scene3d(d):
    el = E('a:scene3d')
    for n in ('camera', 'lightRig'):
        x = sub(el, 'a:' + n, d[n]['a'])
        if d[n]['rot']:
            sub(x, 'a:rot', d[n]['rot'])
    return el


def parse_sp3d(el):
    d = {'a': attrs_of(el), 'parts': []}
    for c in kids(el):
        n = ln_(c)
        if n in ('bevelT', 'bevelB'):
            d['parts'].append([n, attrs_of(c), None])
        elif n in ('extrusionClr', 'contourClr'):
            d['parts'].append([n, None, find_color(c)])
        else:
            unknown(c, 'sp3d')
    return d


def emit_sp3d(d):
    el = E('a:sp3d', d['a'])
    for n, a, c in d['parts']:
        x = sub(el, 'a:' + n, a)
        if c:
            x.append(emit_color(c))
    return el


# ---------------- 幾何 ----------------
def parse_geom(el):
    if ln_(el) == 'prstGeom':
        av = el.find(A + 'avLst')
        gds = [[g.get('name'), g.get('fmla')] for g in kids(av)] if av is not None else []
        return {'t': 'prst', 'prst': el.get('prst'), 'av': gds}
    # 手繪多邊形：只保留真正決定外形的路徑；連接點公式(gdLst/cxnLst)不影響外觀
    d = {'t': 'cust', 'rect': None, 'paths': []}
    rect = el.find(A + 'rect')
    if rect is not None:
        d['rect'] = attrs_of(rect)
    for p in el.find(A + 'pathLst'):
        cmds = []
        for c in p:
            n = ln_(c)
            pts = [int(pt.get('x')) for pt in c] , [int(pt.get('y')) for pt in c]
            flat = [v for xy in zip(*pts) for v in xy]
            if n == 'moveTo':
                cmds.append(['M'] + flat)
            elif n == 'lnTo':
                cmds.append(['L'] + flat)
            elif n == 'cubicBezTo':
                cmds.append(['C'] + flat)
            elif n == 'close':
                cmds.append(['Z'])
            else:
                unknown(c, 'path')
        d['paths'].append({'a': attrs_of(p), 'cmds': cmds})
    return d


def emit_geom(d):
    if d['t'] == 'prst':
        el = E('a:prstGeom', {'prst': d['prst']})
        av = sub(el, 'a:avLst')
        for name, fmla in d['av']:
            sub(av, 'a:gd', {'name': name, 'fmla': fmla})
        return el
    el = E('a:custGeom')
    for t in ('avLst', 'gdLst', 'ahLst', 'cxnLst'):
        sub(el, 'a:' + t)
    if d['rect']:
        sub(el, 'a:rect', d['rect'])
    else:
        sub(el, 'a:rect', {'l': 'l', 't': 't', 'r': 'r', 'b': 'b'})
    pl = sub(el, 'a:pathLst')
    tags = {'M': 'a:moveTo', 'L': 'a:lnTo', 'C': 'a:cubicBezTo', 'Z': 'a:close'}
    for p in d['paths']:
        pe = sub(pl, 'a:path', p['a'])
        for cmd in p['cmds']:
            c = sub(pe, tags[cmd[0]])
            v = cmd[1:]
            for i in range(0, len(v), 2):
                sub(c, 'a:pt', {'x': v[i], 'y': v[i + 1]})
    return el


# ---------------- xfrm ----------------
def parse_xfrm(el):
    if el is None:
        return None
    d = {'a': attrs_of(el)}
    for c in kids(el):
        d[ln_(c)] = [int(v) for v in c.attrib.values()]
    return d


def emit_xfrm(d, tag='a:xfrm'):
    el = E(tag, d['a'])
    names = {'off': ('x', 'y'), 'ext': ('cx', 'cy'), 'chOff': ('x', 'y'), 'chExt': ('cx', 'cy')}
    for k in ('off', 'ext', 'chOff', 'chExt'):
        if k in d:
            sub(el, 'a:' + k, dict(zip(names[k], d[k])))
    return el


# ---------------- 文字 ----------------
def parse_font(el):
    return attrs_of(el)


def parse_rpr(el, ctx):
    if el is None:
        return None
    d = {'tag': ln_(el), 'a': attrs_of(el), 'ln': None, 'fill': None, 'fx': None, 'hl': None,
         'fonts': [], 'links': []}
    for c in kids(el):
        n = ln_(c)
        if n == 'ln':
            d['ln'] = parse_line(c)
        elif n in FILL_TAGS:
            d['fill'] = parse_fill(c)
        elif n == 'effectLst':
            d['fx'] = parse_effects(c)
        elif n == 'highlight':
            d['hl'] = find_color(c)
        elif n in ('latin', 'ea', 'cs', 'sym'):
            d['fonts'].append([n, attrs_of(c)])
        elif n in ('hlinkClick', 'hlinkMouseOver'):
            d['links'].append(ctx.parse_link(c))
        else:
            unknown(c, 'rPr')
    return d


def emit_rpr(d, ctx):
    el = E('a:' + d['tag'], d['a'])
    if d['ln']:
        el.append(emit_line(d['ln']))
    if d['fill']:
        el.append(emit_fill(d['fill']))
    if d['fx'] is not None:
        el.append(emit_effects(d['fx']))
    if d['hl']:
        el.append(E('a:highlight', None, emit_color(d['hl'])))
    for n, a in d['fonts']:
        sub(el, 'a:' + n, a)
    for l in d['links']:
        el.append(ctx.emit_link(l))
    return el


BU = ('buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx', 'buFont',
      'buNone', 'buAutoNum', 'buChar', 'buBlip')


def parse_ppr(el, ctx):
    if el is None:
        return None
    d = {'tag': ln_(el), 'a': attrs_of(el), 'parts': []}
    for c in kids(el):
        n = ln_(c)
        if n in ('lnSpc', 'spcBef', 'spcAft'):
            s = kids(c)[0]
            d['parts'].append([n, [ln_(s), s.get('val')]])
        elif n in ('buClr',):
            d['parts'].append([n, find_color(c)])
        elif n in BU:
            d['parts'].append([n, attrs_of(c)])
        elif n == 'tabLst':
            d['parts'].append([n, [attrs_of(t) for t in kids(c)]])
        elif n == 'defRPr':
            d['parts'].append([n, parse_rpr(c, ctx)])
        else:
            unknown(c, 'pPr')
    return d


def emit_ppr(d, ctx):
    el = E('a:' + d['tag'], d['a'])
    for n, v in d['parts']:
        if n in ('lnSpc', 'spcBef', 'spcAft'):
            sub(sub(el, 'a:' + n), 'a:' + v[0], {'val': v[1]})
        elif n == 'buClr':
            sub(el, 'a:buClr').append(emit_color(v))
        elif n == 'tabLst':
            t = sub(el, 'a:tabLst')
            for a in v:
                sub(t, 'a:tab', a)
        elif n == 'defRPr':
            el.append(emit_rpr(v, ctx))
        else:
            sub(el, 'a:' + n, v)
    return el


def parse_lststyle(el, ctx):
    if el is None:
        return None
    return [parse_ppr(c, ctx) for c in kids(el)]


def emit_lststyle(lst, ctx, tag='a:lstStyle'):
    el = E(tag)
    for p in lst or []:
        el.append(emit_ppr(p, ctx))
    return el


def parse_bodypr(el):
    d = {'a': attrs_of(el), 'parts': []}
    for c in kids(el):
        n = ln_(c)
        if n in ('spAutoFit', 'noAutofit', 'normAutofit', 'flatTx'):
            d['parts'].append([n, attrs_of(c)])
        elif n == 'prstTxWarp':
            d['parts'].append([n, {'prst': c.get('prst'), 'av': [[g.get('name'), g.get('fmla')] for g in kids(c.find(A + 'avLst'))]}])
        elif n == 'scene3d':
            d['parts'].append([n, parse_scene3d(c)])
        else:
            unknown(c, 'bodyPr')
    return d


def emit_bodypr(d):
    el = E('a:bodyPr', d['a'])
    for n, v in d['parts']:
        if n == 'prstTxWarp':
            w = sub(el, 'a:prstTxWarp', {'prst': v['prst']})
            av = sub(w, 'a:avLst')
            for name, f in v['av']:
                sub(av, 'a:gd', {'name': name, 'fmla': f})
        elif n == 'scene3d':
            el.append(emit_scene3d(v))
        else:
            sub(el, 'a:' + n, v)
    return el


def parse_txbody(el, ctx):
    """el = p:txBody / c:rich / c:txPr / a:txBody"""
    if el is None:
        return None
    d = {'body': None, 'lst': None, 'paras': []}
    for c in kids(el):
        n = ln_(c)
        if n == 'bodyPr':
            d['body'] = parse_bodypr(c)
        elif n == 'lstStyle':
            d['lst'] = parse_lststyle(c, ctx)
        elif n == 'p':
            para = {'pPr': None, 'runs': [], 'end': None}
            for r in kids(c):
                rn = ln_(r)
                if rn == 'pPr':
                    para['pPr'] = parse_ppr(r, ctx)
                elif rn == 'r':
                    para['runs'].append({'k': 'r', 'rPr': parse_rpr(r.find(A + 'rPr'), ctx), 't': r.find(A + 't').text or ''})
                elif rn == 'br':
                    para['runs'].append({'k': 'br', 'rPr': parse_rpr(r.find(A + 'rPr'), ctx)})
                elif rn == 'fld':
                    para['runs'].append({'k': 'fld', 'a': attrs_of(r), 'rPr': parse_rpr(r.find(A + 'rPr'), ctx),
                                         'pPr': parse_ppr(r.find(A + 'pPr'), ctx), 't': (r.find(A + 't').text if r.find(A + 't') is not None else None)})
                elif rn == 'endParaRPr':
                    para['end'] = parse_rpr(r, ctx)
                else:
                    unknown(r, 'p')
            d['paras'].append(para)
        else:
            unknown(c, 'txBody')
    return d


def _t(text):
    el = E('a:t')
    el.text = text
    if text != text.strip():
        el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return el


def emit_txbody(d, ctx, tag='p:txBody'):
    el = E(tag)
    if d['body'] is not None:
        el.append(emit_bodypr(d['body']))
    if d['lst'] is not None:
        el.append(emit_lststyle(d['lst'], ctx))
    for para in d['paras']:
        p = sub(el, 'a:p')
        if para['pPr']:
            p.append(emit_ppr(para['pPr'], ctx))
        for r in para['runs']:
            if r['k'] == 'r':
                re_ = sub(p, 'a:r')
                if r['rPr']:
                    re_.append(emit_rpr(r['rPr'], ctx))
                re_.append(_t(r['t']))
            elif r['k'] == 'br':
                b = sub(p, 'a:br')
                if r['rPr']:
                    b.append(emit_rpr(r['rPr'], ctx))
            else:
                f = sub(p, 'a:fld', r['a'])
                if r['rPr']:
                    f.append(emit_rpr(r['rPr'], ctx))
                if r['pPr']:
                    f.append(emit_ppr(r['pPr'], ctx))
                if r['t'] is not None:
                    f.append(_t(r['t']))
        if para['end']:
            p.append(emit_rpr(para['end'], ctx))
    return el


# ---------------- 形狀屬性 spPr ----------------
def parse_sppr(el, ctx):
    if el is None:
        return None
    d = {'a': attrs_of(el), 'xfrm': None, 'geom': None, 'fill': None, 'ln': None, 'fx': None,
         'scene3d': None, 'sp3d': None, 'hidden': []}
    for c in kids(el):
        n = ln_(c)
        if n == 'xfrm':
            d['xfrm'] = parse_xfrm(c)
        elif n in ('prstGeom', 'custGeom'):
            d['geom'] = parse_geom(c)
        elif n in FILL_TAGS:
            d['fill'] = parse_fill(c)
        elif n == 'ln':
            d['ln'] = parse_line(c)
        elif n == 'effectLst':
            d['fx'] = parse_effects(c)
        elif n == 'scene3d':
            d['scene3d'] = parse_scene3d(c)
        elif n == 'sp3d':
            d['sp3d'] = parse_sp3d(c)
        elif n == 'extLst':
            # a14:hiddenFill / hiddenLine：舊版相容用的「隱藏時填色」
            for ext in kids(c):
                for h in kids(ext):
                    hn = ln_(h)
                    if hn == 'hiddenFill':
                        d['hidden'].append(['hiddenFill', take_fill(h)])
                    elif hn == 'hiddenLine':
                        d['hidden'].append(['hiddenLine', parse_line(h)])
                    else:
                        unknown(h, 'spPr/extLst')
        else:
            unknown(c, 'spPr')
    return d


HIDDEN_URI = {'hiddenFill': '{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}', 'hiddenLine': '{91240B29-F687-4F45-9708-019B960494DF}'}


def emit_sppr(d, ctx, tag='p:spPr'):
    el = E(tag, d['a'])
    if d['xfrm']:
        el.append(emit_xfrm(d['xfrm']))
    if d['geom']:
        el.append(emit_geom(d['geom']))
    if d['fill']:
        el.append(emit_fill(d['fill']))
    if d['ln']:
        el.append(emit_line(d['ln']))
    if d['fx'] is not None:
        el.append(emit_effects(d['fx']))
    if d['scene3d']:
        el.append(emit_scene3d(d['scene3d']))
    if d['sp3d']:
        el.append(emit_sp3d(d['sp3d']))
    if d['hidden']:
        ext_lst = sub(el, 'a:extLst')
        for kind, v in d['hidden']:
            ext = sub(ext_lst, 'a:ext', {'uri': HIDDEN_URI[kind]})
            if kind == 'hiddenFill':
                ext.append(E('a14:hiddenFill', None, emit_fill(v)))
            else:
                ext.append(emit_line(v, 'a14:hiddenLine'))
    return el


def parse_style(el):
    if el is None:
        return None
    return [[ln_(c), c.get('idx'), find_color(c)] for c in kids(el)]


def emit_style(lst):
    el = E('p:style')
    for n, idx, c in lst:
        x = sub(el, 'a:' + n, {'idx': idx})
        if c:
            x.append(emit_color(c))
    return el
