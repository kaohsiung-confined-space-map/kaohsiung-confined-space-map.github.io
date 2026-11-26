# -*- coding: utf-8 -*-
"""依網站上的年度資料（data/<年>.json）自動產生地圖簡報 PPTX，並逐項比對 PPT 內的數字與網站資料。
比對全部一致才輸出；任何一年失敗都不影響其他年度與網站本身。
用法：python3 tools/pptgen/make_pptx.py <網站根目錄> <輸出資料夾>
年度代號可為 115 或同一年的另一份地圖 115_1、115_2…（封面與比對用資料內的民國年）
輸出：<輸出資料夾>/<代號>.pptx 與 index.json（{"years": [成功的年度], "failed": {年度: 原因}}）"""
import sys, os, json, subprocess, tempfile, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
site, out = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
os.makedirs(out, exist_ok=True)
years = json.load(open(os.path.join(site, 'data', 'years.json'), encoding='utf-8'))['years']
ok, failed = [], {}
import re
def idkey(v):
    m = re.fullmatch(r'(\d+)(?:_(\d+))?', str(v))
    return (int(m.group(1)), int(m.group(2) or 0)) if m else (10**9, 0)
for y in sorted(years, key=idkey):
    data_path = os.path.join(site, 'data', '%s.json' % y)
    tmp = tempfile.mkdtemp()
    try:
        data = json.load(open(data_path, encoding='utf-8'))
        roc = str(data.get('year') or idkey(y)[0])
        comp = os.path.join(tmp, 'companies.json')
        json.dump(data['summary']['companies'], open(comp, 'w', encoding='utf-8'), ensure_ascii=False)
        spec = os.path.join(tmp, 'spec.json')
        pptx = os.path.join(tmp, '%s.pptx' % y)
        run = lambda *a: subprocess.run([sys.executable, *a], cwd=HERE, capture_output=True, text=True)
        r = run('buildyear.py', data_path, comp, spec, roc)
        print(r.stdout); print(r.stderr, file=sys.stderr)
        if r.returncode:
            raise RuntimeError('產生簡報規格失敗（例如資料出現原簡報沒有的新分類）：' + (r.stderr.strip().splitlines() or ['?'])[-1])
        r = run('rebuild.py', 'build', spec, pptx)
        print(r.stdout); print(r.stderr, file=sys.stderr)
        if r.returncode:
            raise RuntimeError('組成 PPTX 失敗：' + (r.stderr.strip().splitlines() or ['?'])[-1])
        r = run('verify_pptx.py', pptx, data_path, data_path, roc)
        print(r.stdout); print(r.stderr, file=sys.stderr)
        if r.returncode:
            raise RuntimeError('PPT 數字與網站資料比對不符，不提供下載')
        shutil.copy(pptx, os.path.join(out, '%s.pptx' % y))
        ok.append(y)
        print('✓ %s 年 PPT 產生並驗證完成' % y)
    except Exception as e:
        failed[str(y)] = str(e)
        print('✗ %s 年：%s' % (y, e), file=sys.stderr)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
import time
json.dump({'years': ok, 'failed': failed, 'built': int(time.time())}, open(os.path.join(out, 'index.json'), 'w', encoding='utf-8'), ensure_ascii=False)
print('完成：', ok, '失敗：', failed)
sys.exit(1 if failed else 0)  # 有年度失敗時讓 Actions 顯示紅字提醒（網站仍會照常部署）
