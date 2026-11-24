"""用法：
  python rebuild.py extract 原檔.pptx spec.json     # 拆成物件規格
  python rebuild.py build   spec.json 新檔.pptx     # 依規格從零產生
"""
import sys, json, re, zipfile, base64, datetime
from lxml import etree
from dml import *
from slide import *
from parts import *
import chart as CH

P = '{%s}' % NS['p']


def read_rels(z, part):
    d = part.rsplit('/', 1)
    path = '%s/_rels/%s.rels' % (d[0], d[1])
    out = {}
    if path not in z.namelist():
        return out
    for r in etree.fromstring(z.read(path)):
        out[r.get('Id')] = (r.get('Type'), r.get('Target'), r.get('TargetMode'))
    return out


def extract(src, dst):
    z = zipfile.ZipFile(src)
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    prels = read_rels(z, 'ppt/presentation.xml')
    order = [prels[s.get(R_ID)][1].split('/')[-1] for s in pres.find(P + 'sldIdLst')]
    slide_index = {f: i + 1 for i, f in enumerate(order)}

    spec = {'charts': [], 'media': {}}

    def charts_for(ctx_rels_dir):
        def load(target):
            path = 'ppt/charts/' + target.split('/')[-1]
            spec['charts'].append(CH.parse_chart(z.read(path)))
            return len(spec['charts']) - 1
        return load

    # 佈景主題、母片、版面配置
    spec['theme'] = parse_theme(etree.fromstring(z.read('ppt/theme/theme1.xml')))
    mctx = Ctx(read_rels(z, 'ppt/slideMasters/slideMaster1.xml'))
    spec['master'] = parse_master(etree.fromstring(z.read('ppt/slideMasters/slideMaster1.xml')), mctx)
    mroot = etree.fromstring(z.read('ppt/slideMasters/slideMaster1.xml'))
    layout_files = [mctx.rels[x.get(R_ID)][1].split('/')[-1] for x in mroot.find(P + 'sldLayoutIdLst')]
    spec['layouts'] = []
    for lf in layout_files:
        part = 'ppt/slideLayouts/' + lf
        spec['layouts'].append(parse_layout(etree.fromstring(z.read(part)), Ctx(read_rels(z, part))))
    layout_no = {lf: i + 1 for i, lf in enumerate(layout_files)}

    # 投影片
    spec['slides'] = []
    for f in order:
        part = 'ppt/slides/' + f
        root = etree.fromstring(z.read(part))
        rels = read_rels(z, part)
        ctx = Ctx(rels, slide_index)
        lay = [v[1] for v in rels.values() if v[0].endswith('/slideLayout')][0].split('/')[-1]
        for c in kids(root):
            if ln_(c) not in ('cSld', 'clrMapOvr', 'AlternateContent', 'transition', 'extLst'):
                unknown(c, 'sld')
        s = {'a': attrs_of(root), 'layout': layout_no[lay],
             'cSld': parse_cSld(root.find(P + 'cSld'), ctx, charts_for(None)),
             'clr': parse_clrmap_ovr(root.find(P + 'clrMapOvr')),
             'transition': parse_transition(root)}
        spec['slides'].append(s)
        for sh in _walk(s['cSld']['shapes']):
            if sh['k'] == 'pic' and sh['img'] not in spec['media']:
                spec['media'][sh['img']] = base64.b64encode(z.read('ppt/media/' + sh['img'])).decode()
    for lay in spec['layouts'] + [spec['master']]:
        for sh in lay['cSld']['shapes']:
            if sh['k'] == 'pic':
                raise ValueError('layout picture')

    # 簡報層級
    pctx = Ctx()
    spec['pres'] = {'a': attrs_of(pres), 'sldSz': attrs_of(pres.find(P + 'sldSz')),
                    'notesSz': attrs_of(pres.find(P + 'notesSz')),
                    'defaultTextStyle': parse_lststyle(pres.find(P + 'defaultTextStyle'), pctx),
                    'first_slide_id': int(pres.find(P + 'sldIdLst')[0].get('id'))}
    pp = etree.fromstring(z.read('ppt/presProps.xml'))
    sp = pp.find(P + 'showPr')
    spec['presProps'] = {'show': attrs_of(sp),
                         'pen': find_color(sp.find(P + 'penClr')),
                         'laser': find_color(sp.find('.//' + q('p14:laserClr'))),
                         'mru': [parse_color(c) for c in kids(pp.find(P + 'clrMru'))] if pp.find(P + 'clrMru') is not None else []}
    ts = etree.fromstring(z.read('ppt/tableStyles.xml'))
    spec['tableStyles'] = {'def': ts.get('def'), 'styles': [parse_tree(c) for c in kids(ts)]}
    core = etree.fromstring(z.read('docProps/core.xml'))
    spec['core'] = {etree.QName(c).localname: c.text for c in core}
    json.dump(spec, open(dst, 'w'), ensure_ascii=False)
    print('slides', len(spec['slides']), 'charts', len(spec['charts']), 'media', len(spec['media']))


def _walk(shapes):
    for s in shapes:
        yield s


# ======================== 產生 ========================
CT = {
    'pres': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'master': 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    'layout': 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    'slide': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    'theme': 'application/vnd.openxmlformats-officedocument.theme+xml',
    'chart': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    'presProps': 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
    'viewProps': 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
    'tableStyles': 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
    'core': 'application/vnd.openxmlformats-package.core-properties+xml',
    'app': 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
}
ALL_NS = {k: v for k, v in NS.items()}


def ser(root, keep=()):
    """統一整理命名空間後輸出。"""
    used = set()
    for el in root.iter():
        if isinstance(el.tag, str):
            used.add(etree.QName(el).namespace)
            for k in el.attrib:
                if etree.QName(k).namespace:
                    used.add(etree.QName(k).namespace)
    top = {k: v for k, v in NS.items() if v in used or k in keep}
    etree.cleanup_namespaces(root, top_nsmap=top, keep_ns_prefixes=list(keep))
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def rels_xml(rels):
    root = etree.Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationships',
                         nsmap={None: 'http://schemas.openxmlformats.org/package/2006/relationships'})
    for rid, typ, target, mode in rels:
        r = etree.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
        r.set('Id', rid); r.set('Type', typ); r.set('Target', target)
        if mode:
            r.set('TargetMode', mode)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def build(src, dst):
    spec = json.load(open(src))
    files = {}
    overrides = []

    def put(path, data, ct=None):
        files[path] = data
        if ct:
            overrides.append(('/' + path, ct))

    used_media = set()

    def add_media(name):
        used_media.add(name)
        return name

    chart_no = [0]

    def build_chart(idx):
        chart_no[0] += 1
        n = chart_no[0]
        root, xlsx = CH.build_chart(spec['charts'][idx], n)
        name = 'chart%d.xml' % n
        put('ppt/charts/' + name, ser(root, ['c14', 'c16r2']), CT['chart'])
        put('ppt/embeddings/Microsoft_Excel_Worksheet%d.xlsx' % n, xlsx)
        files['ppt/charts/_rels/%s.rels' % name] = rels_xml(
            [('rId1', RT + 'package', '../embeddings/Microsoft_Excel_Worksheet%d.xlsx' % n, None)])
        return name

    # 佈景主題
    put('ppt/theme/theme1.xml', ser(emit_theme(spec['theme'])), CT['theme'])
    # 版面配置與母片
    n_lay = len(spec['layouts'])
    for i, lay in enumerate(spec['layouts']):
        ctx = Ctx()
        root = emit_layout(lay, ctx, i + 1, add_media)
        put('ppt/slideLayouts/slideLayout%d.xml' % (i + 1), ser(root), CT['layout'])
        files['ppt/slideLayouts/_rels/slideLayout%d.xml.rels' % (i + 1)] = rels_xml(ctx.out_rels)
    mctx = Ctx()
    put('ppt/slideMasters/slideMaster1.xml', ser(emit_master(spec['master'], mctx, n_lay, add_media)), CT['master'])
    files['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = rels_xml(mctx.out_rels)

    # 投影片
    for i, s in enumerate(spec['slides']):
        ctx = Ctx()
        ctx.rel(RT + 'slideLayout', '../slideLayouts/slideLayout%d.xml' % s['layout'])
        root = etree.Element(q('p:sld'), nsmap={'a': NS['a'], 'r': NS['r'], 'p': NS['p']})
        for k, v in s['a'].items():
            root.set(k, v)
        root.append(emit_cSld(s['cSld'], ctx, ('slide', i + 1), build_chart, add_media))
        root.append(emit_clrmap_ovr(s['clr']))
        keep = []
        if s['transition']:
            root.append(emit_transition(s['transition']))
            if s['transition'].get('requires'):
                keep.append(s['transition']['requires'])
                # 讓 mc:Choice 的 Requires 前綴在 Choice 上就宣告
                ch = root.find('.//' + q('mc:Choice'))
                pre = s['transition']['requires']
                new = etree.Element(ch.tag, nsmap={pre: NS[pre], 'p14': NS['p14']})
                for k, v in ch.attrib.items():
                    new.set(k, v)
                for c in list(ch):
                    new.append(c)
                ch.getparent().replace(ch, new)
        put('ppt/slides/slide%d.xml' % (i + 1), ser(root, keep), CT['slide'])
        files['ppt/slides/_rels/slide%d.xml.rels' % (i + 1)] = rels_xml(ctx.out_rels)

    for name in used_media:
        files['ppt/media/' + name] = base64.b64decode(spec['media'][name])

    # presentation.xml
    pr = spec['pres']
    pctx = Ctx()
    root = etree.Element(q('p:presentation'), nsmap={'a': NS['a'], 'r': NS['r'], 'p': NS['p']})
    for k, v in pr['a'].items():
        root.set(k, v)
    ml = sub(root, 'p:sldMasterIdLst')
    sub(ml, 'p:sldMasterId', {'id': 2147483696, 'r:id': pctx.rel(RT + 'slideMaster', 'slideMasters/slideMaster1.xml')})
    sl = sub(root, 'p:sldIdLst')
    for i in range(len(spec['slides'])):
        sub(sl, 'p:sldId', {'id': 256 + i, 'r:id': pctx.rel(RT + 'slide', 'slides/slide%d.xml' % (i + 1))})
    sub(root, 'p:sldSz', pr['sldSz'])
    sub(root, 'p:notesSz', pr['notesSz'])
    root.append(emit_lststyle(pr['defaultTextStyle'], pctx, 'p:defaultTextStyle'))
    ext = sub(sub(root, 'p:extLst'), 'p:ext', {'uri': '{EFAFB233-063F-42B5-8137-9DF3F51BA10A}'})
    sub(ext, 'p15:sldGuideLst')
    pctx.rel(RT + 'presProps', 'presProps.xml')
    pctx.rel(RT + 'viewProps', 'viewProps.xml')
    pctx.rel(RT + 'theme', 'theme/theme1.xml')
    pctx.rel(RT + 'tableStyles', 'tableStyles.xml')
    put('ppt/presentation.xml', ser(root), CT['pres'])
    files['ppt/_rels/presentation.xml.rels'] = rels_xml(pctx.out_rels)

    # 放映設定
    pp = spec['presProps']
    root = etree.Element(q('p:presentationPr'), nsmap={'a': NS['a'], 'r': NS['r'], 'p': NS['p']})
    sh = sub(root, 'p:showPr', pp['show'])
    sub(sh, 'p:present'); sub(sh, 'p:sldAll')
    sub(sh, 'p:penClr').append(emit_color(pp['pen']))
    el = sub(sh, 'p:extLst')
    sub(sub(el, 'p:ext', {'uri': '{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}'}), 'p14:laserClr').append(emit_color(pp['laser']))
    sub(sub(el, 'p:ext', {'uri': '{2FDB2607-1784-4EEB-B798-7EB5836EED8A}'}), 'p14:showMediaCtrls', {'val': 1})
    if pp['mru']:
        m = sub(root, 'p:clrMru')
        for c in pp['mru']:
            m.append(emit_color(c))
    el = sub(root, 'p:extLst')
    sub(sub(el, 'p:ext', {'uri': '{E76CE94A-603C-4142-B9EB-6D1370010A27}'}), 'p14:discardImageEditData', {'val': 0})
    sub(sub(el, 'p:ext', {'uri': '{D31A062A-798A-4329-ABDD-BBA856620510}'}), 'p14:defaultImageDpi', {'val': 32767})
    sub(sub(el, 'p:ext', {'uri': '{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}'}), 'p15:chartTrackingRefBased', {'val': 1})
    put('ppt/presProps.xml', ser(root), CT['presProps'])

    root = etree.Element(q('p:viewPr'), nsmap={'a': NS['a'], 'r': NS['r'], 'p': NS['p']})
    nv = sub(root, 'p:normalViewPr', {'horzBarState': 'maximized'})
    sub(nv, 'p:restoredLeft', {'sz': 13514, 'autoAdjust': 0}); sub(nv, 'p:restoredTop', {'sz': 94660})
    cv = sub(sub(sub(root, 'p:slideViewPr'), 'p:cSldViewPr', {'snapToGrid': 0}), 'p:cViewPr', {'varScale': 1})
    sc = sub(cv, 'p:scale'); sub(sc, 'a:sx', {'n': 75, 'd': 100}); sub(sc, 'a:sy', {'n': 75, 'd': 100})
    sub(cv, 'p:origin', {'x': 0, 'y': 0})
    sub(cv.getparent(), 'p:guideLst')
    sub(root, 'p:gridSpacing', {'cx': 72008, 'cy': 72008})
    put('ppt/viewProps.xml', ser(root), CT['viewProps'])

    root = etree.Element(q('a:tblStyleLst'), nsmap={'a': NS['a']})
    root.set('def', spec['tableStyles']['def'])
    for t in spec['tableStyles']['styles']:
        root.append(emit_tree(t))
    put('ppt/tableStyles.xml', ser(root), CT['tableStyles'])

    # 文件摘要
    now = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
    core = spec['core']
    put('docProps/core.xml', ((
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>%s</dc:title><dc:creator>%s</dc:creator><cp:lastModifiedBy>%s</cp:lastModifiedBy><cp:revision>1</cp:revision>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
        '</cp:coreProperties>') % (core.get('title', ''), core.get('creator', ''), core.get('creator', ''), now, now))
        .encode(), CT['core'])
    put('docProps/app.xml', (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Microsoft Office PowerPoint</Application><PresentationFormat>寬螢幕</PresentationFormat>'
        '<Slides>%d</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><AppVersion>16.0000</AppVersion></Properties>'
        % len(spec['slides'])).encode(), CT['app'])

    # 套件結構
    files['_rels/.rels'] = rels_xml([
        ('rId1', RT + 'officeDocument', 'ppt/presentation.xml', None),
        ('rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml', None),
        ('rId3', RT + 'extended-properties', 'docProps/app.xml', None)])
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="png" ContentType="image/png"/>'
          '<Default Extension="jpeg" ContentType="image/jpeg"/>'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
          '<Default Extension="xml" ContentType="application/xml"/>']
    for part, c in overrides:
        ct.append('<Override PartName="%s" ContentType="%s"/>' % (part, c))
    ct.append('</Types>')
    out = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    out.writestr('[Content_Types].xml', ''.join(ct))
    for path in sorted(files):
        out.writestr(path, files[path])
    out.close()
    print('wrote', dst, 'charts', chart_no[0])


if __name__ == '__main__':
    {'extract': extract, 'build': build}[sys.argv[1]](sys.argv[2], sys.argv[3])
