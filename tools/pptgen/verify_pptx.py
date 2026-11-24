# -*- coding: utf-8 -*-
"""直接讀 115 年 PPTX 的 XML（不經過產生程式的規格檔），抽出所有數字，與網站發布的統計逐項比對。
用法：python3 verify_pptx.py <pptx> <網站 115.json> <獨立統計 indep115.json>
"""
import sys, json, zipfile, io, re, posixpath
from lxml import etree
import openpyxl

pptx, web_path, indep_path, YEAR = sys.argv[1:5]
WEB = json.load(open(web_path, encoding='utf-8'))
IND = json.load(open(indep_path, encoding='utf-8'))
z = zipfile.ZipFile(pptx)
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
X = lambda el, path: el.xpath(path, namespaces=NS)
ERR = []
CHECKS = [0]


def ok(cond, msg):
    CHECKS[0] += 1
    if not cond:
        ERR.append(msg)


def rels(part):
    d, f = posixpath.split(part)
    p = posixpath.join(d, '_rels', f + '.rels')
    if p not in z.namelist():
        return {}
    out = {}
    for r in etree.fromstring(z.read(p)):
        out[r.get('Id')] = (r.get('Type').rsplit('/', 1)[-1], posixpath.normpath(posixpath.join(d, r.get('Target'))) if r.get('TargetMode') != 'External' else r.get('Target'))
    return out


pres = etree.fromstring(z.read('ppt/presentation.xml'))
prel = rels('ppt/presentation.xml')
slides = [prel[s.get('{%s}id' % NS['r'])][1] for s in X(pres, '//p:sldIdLst/p:sldId')]
idx_of = {s: i + 1 for i, s in enumerate(slides)}
SW, SH = 12192000, 6858000


def text_of(el):
    return ''.join(X(el, './/a:t/text()'))


def off_ext(sp):
    o = X(sp, './p:spPr/a:xfrm/a:off|./p:xfrm/a:off')
    e = X(sp, './p:spPr/a:xfrm/a:ext|./p:xfrm/a:ext')
    return (int(o[0].get('x')), int(o[0].get('y')), int(e[0].get('cx')), int(e[0].get('cy'))) if o else None


def onslide(b):
    # 露出畫面不到 1 mm（36000 EMU）的視為在畫面外（原簡報有圖表只差一點點沒完全移出畫面）
    m = 36000
    return b and b[0] + b[2] > m and b[0] < SW - m and b[1] + b[3] > m and b[1] < SH - m


def read_chart(part):
    root = etree.fromstring(z.read(part))
    sers = X(root, '//c:ser')
    out = []
    for s in sers:
        name = X(s, './c:tx//c:v/text()')
        cats = X(s, './c:cat//c:pt/c:v/text()')
        vals = [float(v) for v in X(s, './c:val//c:pt/c:v/text()')]
        nval = int(X(s, './c:val//c:ptCount/@val')[0])
        out.append({'name': name[0] if name else None, 'cats': cats, 'vals': vals, 'n': nval})
    # 內嵌 Excel 的數值
    r = rels(part)
    emb = [t for k, (ty, t) in r.items() if ty == 'package']
    wbv = None
    if emb:
        wb = openpyxl.load_workbook(io.BytesIO(z.read(emb[0])))
        ws = wb.active
        wbv = [[c.value for c in row] for row in ws.iter_rows()]
    return out, wbv


def chart_role(sers):
    if len(sers) == 17:
        return 'items'
    s = sers[0]
    cats = s['cats']
    if '儲槽' in cats: return '局限空間地點'
    if '製造業' in cats: return '行業別'
    if '五項以上' in cats: return '非假日違規' if s['name'].startswith('非假日') else '假日違規'
    if cats[:2] == ['假日', '非假日']: return '違規公司分布' if '違規公司' in s['name'] else '場次分布'
    if '缺氧作業' in cats: return '局限空間類型'
    if '偏遠' in cats: return '是否偏遠'
    raise ValueError(s)


def slide_info(part):
    root = etree.fromstring(z.read(part))
    r = rels(part)
    info = {'part': part, 'charts': [], 'texts': [], 'tables': [], 'shapes': [], 'links': []}
    for gf in X(root, '//p:graphicFrame'):
        b = off_ext(gf)
        ch = X(gf, './/c:chart/@r:id')
        if ch:
            sers, wbv = read_chart(r[ch[0]][1])
            info['charts'].append({'box': b, 'on': onslide(b), 'sers': sers, 'wb': wbv, 'role': chart_role(sers)})
        tb = X(gf, './/a:tbl')
        if tb:
            rows = [[text_of(tc) for tc in X(tr, './a:tc')] for tr in X(tb[0], './a:tr')]
            info['tables'].append({'box': b, 'rows': rows})
    for sp in X(root, '//p:sp'):
        name = X(sp, './p:nvSpPr/p:cNvPr/@name')[0]
        t = text_of(sp)
        fill = X(sp, './p:spPr/a:gradFill//a:gs')
        gs = [(g.get('pos'), X(g, './a:schemeClr/@val'), X(g, './a:schemeClr/a:lumMod/@val')) for g in fill]
        lk = []
        for h in X(sp, './p:nvSpPr/p:cNvPr/a:hlinkClick|./p:nvSpPr/p:cNvPr/a:hlinkHover'):
            act = h.get('action') or ''
            rid = h.get('{%s}id' % NS['r'])
            tgt = idx_of.get(r[rid][1]) if rid and rid in r and r[rid][0] == 'slide' else None
            lk.append((etree.QName(h).localname, act.split('jump=')[-1] if 'jump=' in act else act, tgt))
        info['shapes'].append({'name': name, 'text': t, 'box': off_ext(sp), 'grad': gs, 'links': lk,
                               'hidden': X(sp, './p:nvSpPr/p:cNvPr/@hidden') == ['1']})
    return info


INFO = [slide_info(p) for p in slides]
print('投影片數', len(INFO))

ORANGE = [('0', ['accent4'], []), ('100000', ['accent2'], [])]
GREY = [('0', ['bg1'], ['95000']), ('100000', ['bg1'], ['50000'])]
BTN_O = [('50000', ['accent4'], []), ('100000', ['accent2'], [])]
BTN_G = [('50000', ['bg1'], ['85000']), ('100000', ['bg1'], ['65000'])]
DIST = list(WEB['districts'])


def title(si):
    for t in si['tables']:
        if len(t['rows']) == 1 and len(t['rows'][0]) == 1:
            return t['rows'][0][0]


def total_text(si):
    ts = [s['text'] for s in si['shapes'] if s['name'] == '總場次']
    return ts


def num(v):
    return int(v) if float(v).is_integer() else v


def check_chart(where, ch, want):
    s = ch['sers'][0]
    got = dict(zip(s['cats'], [num(v) for v in s['vals']]))
    # PPT 圖表保留原簡報的類別；資料裡沒有出現的類別必須是 0
    extra = {k: v for k, v in got.items() if k not in want}
    ok(all(v == 0 for v in extra.values()), '%s %s 多出的類別不是 0：%s' % (where, ch['role'], extra))
    want = {k: want.get(k, 0) for k in got} if set(want) <= set(got) else want
    ok(s['n'] == len(s['cats']) == len(s['vals']), '%s %s 點數不一致' % (where, ch['role']))
    ok(got == want, '%s %s 圖表快取 %s ≠ 網站 %s' % (where, ch['role'], got, want))
    # 內嵌 Excel
    wb = ch['wb']
    wcats = [r[0] for r in wb[1:] if r[0] is not None]
    wvals = [r[1] for r in wb[1:] if r[0] is not None]
    ok(dict(zip(wcats, wvals)) == want, '%s %s 內嵌Excel %s ≠ 網站 %s' % (where, ch['role'], dict(zip(wcats, wvals)), want))


# ---------- 1. 地圖頁按鈕 → 亮起頁 → 資料頁 ----------
MAP = INFO[1]
btns = {s['text']: s for s in MAP['shapes'] if s['name'].startswith('botton')}
ok(sorted(btns) == sorted(DIST), '地圖頁按鈕與 38 區不符')
summary_pages = []
for X_ in DIST:
    d = WEB['districts'][X_]
    hov = [l[2] for l in btns[X_]['links'] if l[0] == 'hlinkHover']
    ok(len(hov) == 1, '%s 按鈕沒有滑過連結' % X_)
    h = hov[0]
    hs = INFO[h - 1]
    # 亮起的區塊與按鈕
    lit = [s for s in hs['shapes'] if s['grad'] and s['name'] == X_]
    litb = [s for s in hs['shapes'] if s['grad'] and s['name'].startswith('botton') and s['text'] == X_]
    others = [s['name'] for s in hs['shapes'] if s['grad'] and not s['name'].startswith('botton') and s['name'] in DIST and s['name'] != X_]
    ok(len(lit) == 1 and len(litb) == 1 and not others, '%s 亮起頁的亮起物件不對 %d %d %s' % (X_, len(lit), len(litb), others))
    want_c = ORANGE if d['total'] > 0 else GREY
    want_b = BTN_O if d['total'] > 0 else BTN_G
    ok(lit and lit[0]['grad'] == want_c, '%s 亮起顏色錯（總場次 %d）%s' % (X_, d['total'], lit and lit[0]['grad']))
    ok(litb and litb[0]['grad'] == want_b, '%s 按鈕亮起顏色錯 %s' % (X_, litb and litb[0]['grad']))
    # 資料頁
    pages = []
    j = h + 1
    while j <= len(INFO) and title(INFO[j - 1]) == X_:
        pages.append(j)
        j += 1
    if d['total'] > 0:
        ok(len(pages) == 2, '%s 應有 2 頁資料頁，實際 %d' % (X_, len(pages)))
        clk = [l for l in litb[0]['links'] if l[0] == 'hlinkClick'] if litb else []
        ok(clk and clk[0][1] == 'nextslide', '%s 亮起按鈕沒有「點擊到下一頁」' % X_)
        for k, pg in enumerate(pages):
            si = INFO[pg - 1]
            ok(total_text(si) == ['總場次：%d 場' % d['total']], '%s 第%d頁總場次文字 %s' % (X_, k + 1, total_text(si)))
            roles = sorted(c['role'] for c in si['charts'])
            ok(roles == sorted(['場次分布', '違規公司分布', '假日違規', '非假日違規', '局限空間地點', '行業別', '局限空間類型']),
               '%s 第%d頁圖表組成 %s' % (X_, k + 1, roles))
            for ch in si['charts']:
                check_chart('%s 第%d頁' % (X_, k + 1), ch, d['charts'][ch['role']])
            onroles = sorted(c['role'] for c in si['charts'] if c['on'])
            exp_on = ['假日違規', '場次分布', '違規公司分布', '非假日違規'] if k == 0 else ['局限空間地點', '局限空間類型', '行業別']
            ok(onroles == sorted(exp_on), '%s 第%d頁畫面上的圖表 %s' % (X_, k + 1, onroles))
            # 「無資料」小字
            nod = [s for s in si['shapes'] if s['text'] == '無資料']
            exp_n = sum(1 for r in ('假日違規', '非假日違規') if sum(d['charts'][r].values()) == 0)
            ok(len(nod) == exp_n, '%s 第%d頁「無資料」數量 %d，應為 %d' % (X_, k + 1, len(nod), exp_n))
            for s in nod:   # 位置要在對應圓餅圖中央
                role = '假日違規' if s['box'][0] < 6000000 else '非假日違規'
                ok(sum(d['charts'][role].values()) == 0, '%s 第%d頁「無資料」放在有資料的 %s 上' % (X_, k + 1, role))
            chip = [s for s in si['shapes'] if s['text'] == '偏遠地區']
            ok(len(chip) == (1 if d['remote'] is True else 0), '%s 第%d頁偏遠標籤 %d（remote=%s）' % (X_, k + 1, len(chip), d['remote']))
            back = [l for s in si['shapes'] if s['name'] == 'back' and onslide(s['box']) for l in s['links']]
            ok(back == ([('hlinkClick', 'ppaction://hlinksldjump', 2)] if k == 0 else [('hlinkClick', 'previousslide', None)]),
               '%s 第%d頁返回連結 %s' % (X_, k + 1, back))
            if k == 0:
                tl = [l for s in si['shapes'] if s['name'] == '總場次' for l in s['links']]
                ok(tl == [('hlinkClick', 'nextslide', None)], '%s 第1頁總場次連結 %s' % (X_, tl))
            icon = [s for s in si['shapes'] if s['name'] == X_ and onslide(s['box'])]
            ok(len(icon) == 1 and icon[0]['grad'] == ORANGE, '%s 第%d頁區域圖示顏色' % (X_, k + 1))
    else:
        ok(len(pages) == 1, '%s 應有 1 頁「無資料」頁，實際 %d' % (X_, len(pages)))
        si = INFO[pages[0] - 1]
        ok(not si['charts'], '%s 無資料頁不應有圖表' % X_)
        ok(any(s['text'] == '無資料' for s in si['shapes']), '%s 無資料頁沒有「無資料」' % X_)
        ok(total_text(si) == ['總場次：0 場'], '%s 無資料頁總場次 %s' % (X_, total_text(si)))
        icon = [s for s in si['shapes'] if s['name'] == X_ and onslide(s['box'])]
        ok(len(icon) == 1 and icon[0]['grad'] == GREY, '%s 無資料頁圖示顏色' % X_)

# ---------- 2. 全區總覽 ----------
S = WEB['summary']
sums = [i for i, si in enumerate(INFO, 1) if title(si) == '高雄市全區總覽']
p1 = [i for i in sums if any(c['role'] == '場次分布' and c['on'] for c in INFO[i - 1]['charts'])][0]
p2 = [i for i in sums if any(c['role'] == '局限空間地點' and c['on'] for c in INFO[i - 1]['charts'])][0]
p4 = [i for i in sums if any(c['role'] == 'items' for c in INFO[i - 1]['charts'])][0]
for pg in (p1, p2):
    si = INFO[pg - 1]
    ok(total_text(si) == ['總場次：%d 場' % S['total']], '全區 總場次文字 %s' % total_text(si))
    for ch in si['charts']:
        check_chart('全區第%d頁' % pg, ch, S['charts'][ch['role']])
ok(sorted(c['role'] for c in INFO[p2 - 1]['charts'] if c['on']) == sorted(['局限空間地點', '行業別', '局限空間類型'] + (['是否偏遠'] if '是否偏遠' in S['charts'] else [])),
   '全區地點頁圖表組成')
# 缺失項目統計
si = INFO[p4 - 1]
bar = [c for c in si['charts'] if c['role'] == 'items'][0]
got = {s['name']: num(s['vals'][0]) for s in bar['sers']}
want = {it['name']: it['v'] for it in S['items']}
ok(got == want, '缺失項目長條圖 %s ≠ %s' % (got, want))
wb = bar['wb']
ok(dict(zip(wb[0][1:], wb[1][1:])) == want, '缺失項目長條圖內嵌Excel ≠ 網站')
t18 = [t for t in si['tables'] if len(t['rows']) == 18][0]['rows']
tgot = {('無' if r[0] == '無' else '第%s項' % r[0]): int(r[1]) for r in t18[1:]}
ok(tgot == want, '缺失項目表格 %s ≠ %s' % (tgot, want))
t2 = [t for t in si['tables'] if len(t['rows']) == 2][0]['rows']
nums = {k: v for k, v in want.items() if k != '無'}
mx = '、'.join(k[1:-1] for k, v in nums.items() if v == max(nums.values()))
mn = '、'.join(k[1:-1] for k, v in nums.items() if v == min(nums.values()))
ok(t2 == [['缺失最多項目', mx], ['缺失最少項目', mn]], '最多/最少 %s，應為 %s/%s' % (t2, mx, mn))
# 缺失五項以上公司
rows = []
for si in INFO:
    if title(si) == '單次工程缺失五項以上公司':
        big = [t for t in si['tables'] if len(t['rows']) > 1][0]['rows']
        ok(big[0] == ['日期', '承攬廠商', '缺失項次總數', '缺失項目（項次）'], '公司表標題列 %s' % big[0])
        rows += big[1:]
wantc = [[c[3], c[0], str(c[1]), c[2]] for c in S['companies']]
ok(rows == wantc, '缺失五項以上公司表與網站不同（PPT %d 列、網站 %d 列）' % (len(rows), len(wantc)))
# 封面
cover = ''.join(s['text'] for s in INFO[0]['shapes'])
ok('中華民國%s年' % YEAR in cover and len(re.findall(r'1[01]\d', cover)) == 1, '封面年份 %s' % cover)

# ---------- 3. 獨立統計 vs 網站（再確認一次） ----------
for k in ('districts',):
    ok(IND[k] == WEB[k], '獨立統計各區 ≠ 網站')
s1 = {k: v for k, v in IND['summary'].items() if k != 'companies'}
s2 = {k: v for k, v in WEB['summary'].items() if k != 'companies'}
ok(s1 == s2, '獨立統計全區 ≠ 網站')
ok(sorted(map(tuple, IND['summary']['companies'])) == sorted(map(tuple, WEB['summary']['companies'])), '獨立統計公司名單 ≠ 網站')

print('檢查項目', CHECKS[0], '不符', len(ERR))
for e in ERR:
    print(' ✗', e)
n_ch = sum(len(si['charts']) for si in INFO)
print('PPT 內圖表數（含畫面外的轉場用複本）', n_ch)
sys.exit(1 if ERR else 0)
