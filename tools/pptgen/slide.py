"""投影片層級：形狀、圖片、表格、圖表框、超連結、背景、轉場。"""
import uuid, hashlib
from lxml import etree
from dml import *

P = '{%s}' % NS['p']


class Ctx:
    """解析時：把 r:id 轉成語意目標；產生時：替目標建立新的關聯。"""
    def __init__(self, rels=None, slide_index=None):
        self.rels = rels or {}            # 解析用 rId -> (type, target, mode)
        self.slide_index = slide_index or {}
        self.out_rels = []                # 產生用 [(rId, type, target, mode)]
        self.n = 0

    # ---- 解析 ----
    def parse_link(self, el):
        a = attrs_of(el)
        rid = a.pop('r:id', None)
        tgt = None
        if rid:
            typ, target, mode = self.rels[rid]
            if typ.endswith('/slide'):
                tgt = ['slide', self.slide_index[target.split('/')[-1]]]
            elif mode == 'External':
                tgt = ['url', target]
            else:
                raise ValueError(typ)
        if kids(el):
            unknown(kids(el)[0], 'hlink')
        return {'tag': ln_(el), 'a': a, 'tgt': tgt}

    # ---- 產生 ----
    def rel(self, typ, target, mode=None):
        for rid, t, tg, m in self.out_rels:
            if t == typ and tg == target and m == mode:
                return rid
        self.n += 1
        rid = 'rId%d' % self.n
        self.out_rels.append((rid, typ, target, mode))
        return rid

    def emit_link(self, d):
        a = dict(d['a'])
        if d['tgt']:
            k, v = d['tgt']
            if k == 'slide':
                a = {'r:id': self.rel(RT + 'slide', '../slides/slide%d.xml' % v), **a}
            else:
                a = {'r:id': self.rel(RT + 'hyperlink', v, 'External'), **a}
        else:
            a = {'r:id': '', **a}
        return E('a:' + d['tag'], a)


RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'


def stable_guid(*parts):
    h = hashlib.md5('|'.join(map(str, parts)).encode()).hexdigest()
    return '{%s}' % str(uuid.UUID(h)).upper()


def stable_int(*parts):
    return int(hashlib.md5('|'.join(map(str, parts)).encode()).hexdigest()[:8], 16)


# ---------------- 非視覺屬性 ----------------
IDENT = {}


def parse_cnvpr(el, ctx):
    a = attrs_of(el)
    d = {'id': a.pop('id'), 'name': a.pop('name'), 'a': a, 'links': [], 'ident': None}
    for c in kids(el):
        n = ln_(c)
        if n in ('hlinkClick', 'hlinkHover'):
            d['links'].append(ctx.parse_link(c))
        elif n == 'extLst':
            for ext in kids(c):
                for x in kids(ext):
                    if ln_(x) != 'creationId':
                        unknown(x, 'cNvPr/ext')
                    # 只記錄「哪些物件是同一個身分」(編號)，不保留原始識別碼
                    d['ident'] = IDENT.setdefault(x.get('id'), len(IDENT))
        else:
            unknown(c, 'cNvPr')
    return d


def emit_cnvpr(d, ctx, key):
    el = E('p:cNvPr', {'id': d['id'], 'name': d['name'], **d['a']})
    for l in d['links']:
        el.append(ctx.emit_link(l))
    # 每個身分群組產生全新的識別碼；原本跨頁共用身分的物件，重建後仍共用，「變形」配對關係不變
    if d.get('ident') is not None:
        ext = sub(sub(el, 'a:extLst'), 'a:ext', {'uri': '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}'})
        sub(ext, 'a16:creationId', {'id': stable_guid('rebuild-identity', d['ident'])})
    return el


def parse_nvpr(el):
    d = {'a': attrs_of(el), 'ph': None}
    for c in kids(el):
        n = ln_(c)
        if n == 'ph':
            d['ph'] = attrs_of(c)
        elif n == 'extLst':
            pass  # p14:modId 為版本戳記，重建時重新產生
        else:
            unknown(c, 'nvPr')
    return d


def emit_nvpr(d, modid=None):
    el = E('p:nvPr', d['a'])
    if d['ph'] is not None:
        sub(el, 'p:ph', d['ph'])
    if modid is not None:
        ext = sub(sub(el, 'p:extLst'), 'p:ext', {'uri': '{D42A27DB-BD31-4B8C-83A1-F6EECF244321}'})
        sub(ext, 'p14:modId', {'val': modid})
    return el


def parse_locks(el):
    """cNvSpPr / cNvPicPr / cNvGraphicFramePr"""
    d = {'a': attrs_of(el), 'locks': None}
    for c in kids(el):
        if ln_(c) in ('spLocks', 'picLocks', 'graphicFrameLocks', 'grpSpLocks'):
            d['locks'] = [ln_(c), attrs_of(c)]
        else:
            unknown(c, 'cNv*Pr')
    return d


def emit_locks(tag, d):
    el = E(tag, d['a'])
    if d['locks']:
        sub(el, 'a:' + d['locks'][0], d['locks'][1])
    return el


# ---------------- 形狀 ----------------
def parse_shape(el, ctx, charts):
    t = ln_(el)
    if t == 'sp':
        nv = el.find(P + 'nvSpPr')
        return {'k': 'sp', 'cnv': parse_cnvpr(nv.find(P + 'cNvPr'), ctx),
                'cnvsp': parse_locks(nv.find(P + 'cNvSpPr')), 'nvpr': parse_nvpr(nv.find(P + 'nvPr')),
                'sppr': parse_sppr(el.find(P + 'spPr'), ctx), 'style': parse_style(el.find(P + 'style')),
                'tx': parse_txbody(el.find(P + 'txBody'), ctx), 'a': attrs_of(el)}
    if t == 'pic':
        nv = el.find(P + 'nvPicPr')
        bf = el.find(P + 'blipFill')
        blip = bf.find(A + 'blip')
        bfx = []
        for c in kids(blip):
            n = ln_(c)
            if n == 'extLst':
                continue  # a14:useLocalDpi
            if n == 'duotone':
                bfx.append([n, None, [parse_color(x) for x in kids(c)]])
            else:
                bfx.append([n, attrs_of(c), None])
        src = bf.find(A + 'srcRect')
        mode = [ln_(c) for c in kids(bf) if ln_(c) in ('stretch', 'tile')]
        typ, target, _ = ctx.rels[blip.get(q('r:embed'))]
        return {'k': 'pic', 'cnv': parse_cnvpr(nv.find(P + 'cNvPr'), ctx),
                'cnvpic': parse_locks(nv.find(P + 'cNvPicPr')), 'nvpr': parse_nvpr(nv.find(P + 'nvPr')),
                'img': target.split('/')[-1], 'bfa': attrs_of(bf), 'bfx': bfx,
                'src': attrs_of(src) if src is not None else None, 'mode': mode[0],
                'sppr': parse_sppr(el.find(P + 'spPr'), ctx), 'style': parse_style(el.find(P + 'style'))}
    if t == 'graphicFrame':
        nv = el.find(P + 'nvGraphicFramePr')
        gd = el.find(A + 'graphic/' + A + 'graphicData')
        d = {'k': 'gf', 'cnv': parse_cnvpr(nv.find(P + 'cNvPr'), ctx),
             'cnvgf': parse_locks(nv.find(P + 'cNvGraphicFramePr')), 'nvpr': parse_nvpr(nv.find(P + 'nvPr')),
             'xfrm': parse_xfrm(el.find(P + 'xfrm')), 'a': attrs_of(el)}
        uri = gd.get('uri')
        if uri.endswith('/table'):
            d['tbl'] = parse_table(gd.find(A + 'tbl'), ctx)
        elif uri.endswith('/chart'):
            cid = gd.find(q('c:chart')).get(R_ID)
            d['chart'] = charts(ctx.rels[cid][1])
        else:
            raise Unknown(uri)
        return d
    unknown(el, 'spTree')


def emit_shape(d, ctx, key, build_chart, add_media):
    k = d['k']
    if k == 'sp':
        el = E('p:sp', d['a'])
        nv = sub(el, 'p:nvSpPr')
        nv.append(emit_cnvpr(d['cnv'], ctx, key))
        nv.append(emit_locks('p:cNvSpPr', d['cnvsp']))
        nv.append(emit_nvpr(d['nvpr']))
        el.append(emit_sppr(d['sppr'], ctx))
        if d['style']:
            el.append(emit_style(d['style']))
        if d['tx']:
            el.append(emit_txbody(d['tx'], ctx))
        return el
    if k == 'pic':
        el = E('p:pic')
        nv = sub(el, 'p:nvPicPr')
        nv.append(emit_cnvpr(d['cnv'], ctx, key))
        nv.append(emit_locks('p:cNvPicPr', d['cnvpic']))
        nv.append(emit_nvpr(d['nvpr']))
        bf = sub(el, 'p:blipFill', d['bfa'])
        rid = ctx.rel(RT + 'image', '../media/' + add_media(d['img']))
        blip = sub(bf, 'a:blip', {'r:embed': rid})
        for n, a, cols in d['bfx']:
            x = sub(blip, 'a:' + n, a)
            for c in cols or []:
                x.append(emit_color(c))
        ext = sub(sub(blip, 'a:extLst'), 'a:ext', {'uri': '{28A0092B-C50C-407E-A947-70E740481C1C}'})
        sub(ext, 'a14:useLocalDpi', {'val': '0'})
        if d['src'] is not None:
            sub(bf, 'a:srcRect', d['src'])
        sub(sub(bf, 'a:' + d['mode']), 'a:fillRect') if d['mode'] == 'stretch' else sub(bf, 'a:tile')
        el.append(emit_sppr(d['sppr'], ctx))
        if d['style']:
            el.append(emit_style(d['style']))
        return el
    if k == 'gf':
        el = E('p:graphicFrame', d['a'])
        nv = sub(el, 'p:nvGraphicFramePr')
        nv.append(emit_cnvpr(d['cnv'], ctx, key))
        nv.append(emit_locks('p:cNvGraphicFramePr', d['cnvgf']))
        nv.append(emit_nvpr(d['nvpr'], stable_int('mod', key, d['cnv']['id'])))
        el.append(emit_xfrm(d['xfrm'], 'p:xfrm'))
        g = sub(el, 'a:graphic')
        if 'tbl' in d:
            gd = sub(g, 'a:graphicData', {'uri': 'http://schemas.openxmlformats.org/drawingml/2006/table'})
            gd.append(emit_table(d['tbl'], ctx, key))
        else:
            gd = sub(g, 'a:graphicData', {'uri': 'http://schemas.openxmlformats.org/drawingml/2006/chart'})
            target = build_chart(d['chart'])
            rid = ctx.rel(RT + 'chart', '../charts/' + target)
            sub(gd, 'c:chart', {'r:id': rid})
        return el
    raise ValueError(k)


# ---------------- 表格 ----------------
CELL_LN = ('lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr')


def parse_table(el, ctx):
    tp = el.find(A + 'tblPr')
    sid = tp.find(A + 'tableStyleId')
    d = {'a': attrs_of(tp), 'style': sid.text if sid is not None else None,
         'tpfill': take_fill(tp), 'cols': [], 'rows': []}
    for c in kids(tp):
        if ln_(c) not in ('tableStyleId',) + FILL_TAGS:
            unknown(c, 'tblPr')
    for gc in el.find(A + 'tblGrid'):
        d['cols'].append(int(gc.get('w')))
    for tr in el.findall(A + 'tr'):
        row = {'h': int(tr.get('h')), 'cells': []}
        for tc in tr.findall(A + 'tc'):
            pr = tc.find(A + 'tcPr')
            cell = {'a': attrs_of(tc), 'tx': parse_txbody(tc.find(A + 'txBody'), ctx),
                    'pa': attrs_of(pr) if pr is not None else None, 'lines': [], 'fill': None}
            if pr is not None:
                for c in kids(pr):
                    n = ln_(c)
                    if n in CELL_LN:
                        cell['lines'].append([n, parse_line(c)])
                    elif n in FILL_TAGS:
                        cell['fill'] = parse_fill(c)
                    else:
                        unknown(c, 'tcPr')
            row['cells'].append(cell)
        d['rows'].append(row)
    return d


def emit_table(d, ctx, key):
    el = E('a:tbl')
    tp = sub(el, 'a:tblPr', d['a'])
    if d['tpfill']:
        tp.append(emit_fill(d['tpfill']))
    if d['style']:
        sub(tp, 'a:tableStyleId').text = d['style']
    grid = sub(el, 'a:tblGrid')
    for i, w in enumerate(d['cols']):
        gc = sub(grid, 'a:gridCol', {'w': w})
        ext = sub(sub(gc, 'a:extLst'), 'a:ext', {'uri': '{9D8B030D-6E8A-4147-A177-3AD203B41FA5}'})
        sub(ext, 'a16:colId', {'val': stable_int('col', key, i)})
    for j, row in enumerate(d['rows']):
        tr = sub(el, 'a:tr', {'h': row['h']})
        for cell in row['cells']:
            tc = sub(tr, 'a:tc', cell['a'])
            tc.append(emit_txbody(cell['tx'], ctx, 'a:txBody'))
            if cell['pa'] is not None:
                pr = sub(tc, 'a:tcPr', cell['pa'])
                for n, l in cell['lines']:
                    pr.append(emit_line(l, 'a:' + n))
                if cell['fill']:
                    pr.append(emit_fill(cell['fill']))
        ext = sub(sub(tr, 'a:extLst'), 'a:ext', {'uri': '{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}'})
        sub(ext, 'a16:rowId', {'val': stable_int('row', key, j)})
    return el


# ---------------- 背景與轉場 ----------------
def parse_bg(el):
    if el is None:
        return None
    c = kids(el)[0]
    if ln_(c) == 'bgPr':
        fill = take_fill(c)
        fx = c.find(A + 'effectLst')
        return {'t': 'pr', 'a': attrs_of(c), 'fill': fill, 'fx': parse_effects(fx) if fx is not None else None}
    return {'t': 'ref', 'idx': c.get('idx'), 'c': find_color(c)}


def emit_bg(d):
    el = E('p:bg')
    if d['t'] == 'pr':
        pr = sub(el, 'p:bgPr', d['a'])
        pr.append(emit_fill(d['fill']))
        if d['fx'] is not None:
            pr.append(emit_effects(d['fx']))
    else:
        sub(el, 'p:bgRef', {'idx': d['idx']}).append(emit_color(d['c']))
    return el


def parse_transition(sld):
    ac = sld.find(q('mc:AlternateContent'))
    if ac is None:
        tr = sld.find(P + 'transition')
        if tr is None:
            return None
        return {'choice': None, 'fallback': _parse_tr(tr)}
    ch = ac.find(q('mc:Choice'))
    fb = ac.find(q('mc:Fallback'))
    return {'requires': ch.get('Requires'), 'choice': _parse_tr(ch.find(P + 'transition')),
            'fallback': _parse_tr(fb.find(P + 'transition'))}


def _parse_tr(tr):
    a = attrs_of(tr)
    eff = [[pfx(c), ln_(c), attrs_of(c)] for c in kids(tr)]
    for c in kids(tr):
        if kids(c):
            unknown(kids(c)[0], 'transition')
    return {'a': a, 'eff': eff}


def _emit_tr(d):
    el = E('p:transition', d['a'])
    for pre, n, a in d['eff']:
        sub(el, '%s:%s' % (pre, n), a)
    return el


def emit_transition(d):
    if d['choice'] is None:
        return _emit_tr(d['fallback'])
    ac = E('mc:AlternateContent')
    ch = sub(ac, 'mc:Choice', {'Requires': d['requires']})
    ch.append(_emit_tr(d['choice']))
    sub(ac, 'mc:Fallback').append(_emit_tr(d['fallback']))
    return ac
