"""圖表：解析成「數據 + 設定 + 格式」模型，再重新產生圖表 XML 與內嵌 Excel 數據表。"""
import io
from lxml import etree
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from dml import *
from slide import stable_guid

C = '{%s}' % NS['c']
TEXT_EL = {'f', 'formatCode', 'v', 'separator'}
SHEET = '工作表1'


class NoLinks:
    def parse_link(self, el):
        unknown(el, 'chart link')

    def emit_link(self, d):
        raise ValueError


NL = NoLinks()


def parse_node(el):
    """通用設定節點；格式（spPr/txPr/rich）與數據（cat/val/tx）交給專門的模型。"""
    n = ln_(el)
    pre = pfx(el)
    if n == 'spPr' and pre == 'c':
        return {'n': n, 'sp': parse_sppr(el, NL)}
    if n == 'txPr':
        return {'n': n, 'tb': parse_txbody(el, NL)}
    if n == 'rich':
        return {'n': n, 'tb': parse_txbody(el, NL)}
    if n in TEXT_EL:
        return {'n': n, 'pre': pre, 'text': el.text or ''}
    if n == 'extLst':
        exts = []
        for ext in kids(el):
            uri = ext.get('uri')
            inner = kids(ext)
            if inner and ln_(inner[0]) == 'uniqueId':
                exts.append({'uri': uri, 'uid': True})       # 重建時重新編號
            elif inner and ln_(inner[0]) == 'style':
                continue
            else:
                exts.append({'uri': uri, 'kids': [parse_node(x) for x in inner]})
        return {'n': n, 'exts': exts}
    return {'n': n, 'pre': pre, 'a': attrs_of(el), 'k': [parse_node(c) for c in kids(el)]}


def emit_node(d, uid):
    n = d['n']
    if n == 'spPr':
        return emit_sppr(d['sp'], NL, 'c:spPr')
    if n in ('txPr', 'rich'):
        return emit_txbody(d['tb'], NL, 'c:' + n)
    if 'text' in d:
        el = E('%s:%s' % (d['pre'], n))
        el.text = d['text']
        return el
    if n == 'extLst':
        el = E('c:extLst')
        for x in d['exts']:
            ext = sub(el, 'c:ext', {'uri': x['uri']})
            if x.get('uid'):
                sub(ext, 'c16:uniqueId', {'val': uid()})
            else:
                for k in x['kids']:
                    ext.append(emit_node(k, uid))
        return el
    el = E('%s:%s' % (d['pre'], n), d['a'])
    for k in d['k']:
        el.append(emit_node(k, uid))
    return el


# ---------- 數據 ----------
def parse_data(el):
    """c:tx / c:cat / c:val 內的 strRef/numRef → 純數據"""
    ref = kids(el)[0]
    kind = ln_(ref)
    if kind not in ('strRef', 'numRef'):
        unknown(ref, 'data')
    cache = ref.find(C + ('strCache' if kind == 'strRef' else 'numCache'))
    fc = cache.find(C + 'formatCode')
    n = int(cache.find(C + 'ptCount').get('val'))
    pts = {int(p.get('idx')): p.find(C + 'v').text for p in cache.findall(C + 'pt')}
    return {'kind': kind, 'fmt': fc.text if fc is not None else None, 'n': n, 'pts': pts}


def emit_data(tag, d, ref):
    el = E(tag)
    r = sub(el, 'c:' + d['kind'])
    sub(r, 'c:f').text = ref
    cache = sub(r, 'c:strCache' if d['kind'] == 'strRef' else 'c:numCache')
    if d['fmt'] is not None:
        sub(cache, 'c:formatCode').text = d['fmt']
    sub(cache, 'c:ptCount', {'val': d['n']})
    for i in sorted(d['pts']):
        sub(sub(cache, 'c:pt', {'idx': i}), 'c:v').text = d['pts'][i]
    return el


def parse_chart(xml_bytes):
    root = etree.fromstring(xml_bytes)
    model = {'top': [], 'series': []}
    for c in kids(root):
        n = ln_(c)
        if n == 'AlternateContent':     # 圖表樣式編號 c14:style / c:style
            model['top'].append({'n': 'style', 'c14': c.find('.//' + q('c14:style')).get('val'),
                                 'c': c.find('.//' + C + 'style').get('val')})
        elif n == 'externalData':
            model['top'].append({'n': 'externalData', 'auto': c.find(C + 'autoUpdate').get('val')})
        else:
            model['top'].append(parse_node_with_series(c, model))
    return model


def parse_node_with_series(el, model):
    if ln_(el) == 'ser':
        s = {'n': 'ser', 'k': []}
        for c in kids(el):
            cn = ln_(c)
            if cn in ('tx', 'cat', 'val'):
                s[cn] = parse_data(c)
                s['k'].append({'n': '@' + cn})
            else:
                s['k'].append(parse_node(c))
        model['series'].append(s)
        return {'n': '@ser', 'i': len(model['series']) - 1}
    if ln_(el) in TEXT_EL or ln_(el) in ('spPr', 'txPr', 'extLst', 'rich'):
        return parse_node(el)
    d = {'n': ln_(el), 'pre': pfx(el), 'a': attrs_of(el), 'k': []}
    for c in kids(el):
        d['k'].append(parse_node_with_series(c, model))
    return d


def build_chart(model, chart_no):
    """回傳 (圖表 XML bytes, 內嵌活頁簿 bytes)"""
    counter = [0]

    def uid():
        counter[0] += 1
        return '{%08X-%s}' % (counter[0], stable_guid('chart', chart_no)[10:-1])

    series = model['series']
    for s_ in series:   # JSON 會把整數索引變字串，這裡還原
        for k in ('tx', 'cat', 'val'):
            if k in s_:
                s_[k]['pts'] = {int(i): v for i, v in s_[k]['pts'].items()}
    nmax = max([s['cat']['n'] if 'cat' in s else 0 for s in series] + [s['val']['n'] for s in series])

    def emit_any(d):
        if d['n'] == '@ser':
            s = series[d['i']]
            col = get_column_letter(d['i'] + 2)
            el = E('c:ser')
            for k in s['k']:
                if k['n'] == '@tx':
                    el.append(emit_data('c:tx', s['tx'], "%s!$%s$1" % (SHEET, col)))
                elif k['n'] == '@cat':
                    el.append(emit_data('c:cat', s['cat'], "%s!$A$2:$A$%d" % (SHEET, s['cat']['n'] + 1) if s['cat']['n'] > 1 else "%s!$A$2" % SHEET))
                elif k['n'] == '@val':
                    n = s['val']['n']
                    el.append(emit_data('c:val', s['val'], "%s!$%s$2:$%s$%d" % (SHEET, col, col, n + 1) if n > 1 else "%s!$%s$2" % (SHEET, col)))
                else:
                    el.append(emit_node(k, uid))
            return el
        if d['n'] in ('spPr', 'txPr', 'rich', 'extLst') or 'text' in d:
            return emit_node(d, uid)
        el = E('%s:%s' % (d['pre'], d['n']), d['a'])
        for k in d['k']:
            el.append(emit_any(k))
        return el

    root = etree.Element(q('c:chartSpace'), nsmap={'c': NS['c'], 'a': NS['a'], 'r': NS['r'],
                                                    'c16r2': NS['c16r2']})
    for d in model['top']:
        if d['n'] == 'style':
            ac = sub(root, 'mc:AlternateContent')
            ch = sub(ac, 'mc:Choice', {'Requires': 'c14'})
            sub(ch, 'c14:style', {'val': d['c14']})
            sub(sub(ac, 'mc:Fallback'), 'c:style', {'val': d['c']})
        elif d['n'] == 'externalData':
            ex = sub(root, 'c:externalData', {'r:id': 'rId1'})
            sub(ex, 'c:autoUpdate', {'val': d['auto']})
        else:
            root.append(emit_any(d))
    # 內嵌 Excel：依數據重新產生
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    if series and 'cat' in series[0]:
        for i, v in series[0]['cat']['pts'].items():
            ws.cell(row=i + 2, column=1, value=v)
    for j, s in enumerate(series):
        if 'tx' in s:
            ws.cell(row=1, column=j + 2, value=s['tx']['pts'].get(0))
        for i, v in s['val']['pts'].items():
            try:
                fv = float(v)
                v = int(fv) if fv.is_integer() else fv
            except (TypeError, ValueError):
                pass
            ws.cell(row=i + 2, column=j + 2, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    return root, buf.getvalue()
