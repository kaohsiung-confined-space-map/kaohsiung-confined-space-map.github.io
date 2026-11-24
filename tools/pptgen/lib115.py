import json
def T(sh): return ''.join(r.get('t','') for p in (sh.get('tx') or {'paras':[]})['paras'] for r in p['runs'])
def flat_shapes(shs):
    out=[]
    def walk(s):
        for sh in s:
            out.append(sh)
            if sh['k']=='grp': walk(sh['shapes'])
    walk(shs); return out
def related(slide, X):
    """與區 X 相關的三種形狀：地圖區塊/圖示、按鈕、標籤"""
    out={'path':[], 'btn':[], 'label':[]}
    for sh in flat_shapes(slide['cSld']['shapes']):
        n=sh['cnv']['name']
        if sh['k']!='sp': continue
        if n==X: out['path'].append(sh)
        elif n.startswith('botton') and T(sh)==X: out['btn'].append(sh)
        elif n.startswith('文字方塊') and T(sh)==X: out['label'].append(sh)
    return out
