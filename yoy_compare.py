"""
yoy_compare.py
兩季各區同期比較頁面產生器（25/26 vs 26/27）
─────────────────────────────────────────────
用法：python yoy_compare.py          → 輸出 yoy_compare.html（本機開啟，不含個資）
口徑：
  純課   團號含 ZN(野澤) / ZB(斑尾) / ZY(湯澤)
  團客   團號開頭 DBC-N / S4-J-N → 野澤；DBC-B / S4-J-B → 斑尾；（龍平不比較）
  其他團號（25/26 的 YL/YP/PK…）一律不計
  天數   以各季開賣日 5/1 起算（25/26 = 2025/5/1，26/27 = 2026/5/1）
  25/26  5/1 前有先售（day<0），頁面可切換含/不含
"""
import os, re, sys, json, glob, warnings
from datetime import datetime
import pandas as pd

warnings.filterwarnings('ignore')
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
from agents.sales import load_xlsx, find_xlsx_files   # noqa: E402

CATS = [
    ('野澤團客', '野澤', '團客'), ('野澤純課', '野澤', '純課'),
    ('斑尾團客', '斑尾', '團客'), ('斑尾純課', '斑尾', '純課'),
    ('湯澤純課', '湯澤', '純課'),
]


def categorize(g):
    s = str(g)
    if 'ZN' in s: return '野澤純課'
    if 'ZB' in s: return '斑尾純課'
    if 'ZY' in s: return '湯澤純課'
    if re.match(r'^(DBC-|S4-J-)N', s): return '野澤團客'
    if re.match(r'^(DBC-|S4-J-)B', s): return '斑尾團客'
    return None


def prep(path, start):
    df = load_xlsx(path)
    df['cat'] = df['團號'].apply(categorize)
    df = df[df['cat'].notna()].copy()
    df['fee'] = pd.to_numeric(df['團費'].astype(str).str.replace(',', '').str.replace(' ', ''),
                              errors='coerce').fillna(0)
    df['su'] = pd.to_datetime(df['報名日期'], errors='coerce')
    df['dep'] = pd.to_datetime(df['出發日期'], errors='coerce', format='mixed')
    df['day'] = (df['su'] - pd.Timestamp(start)).dt.days
    df = df[df['day'].notna()]
    return df


def pack(df):
    out = {}
    for key, _, _ in CATS:
        s = df[df['cat'] == key]
        days = {}
        for d, g in s.groupby('day'):
            days[int(d)] = [int(len(g)), int(g['fee'].sum())]
        dep = {}
        for m, g in s.groupby(s['dep'].dt.month):
            if pd.notna(m): dep[int(m)] = [int(len(g)), int(g['fee'].sum())]
        out[key] = {'days': days, 'dep': dep}
    return out


def main():
    cur, hist = find_xlsx_files()
    prev = dict(hist).get('25/26')
    if not prev or not cur:
        raise SystemExit('找不到 25-26 或當季原始大總表')
    m = re.search(r'原始大總表(\d{2})(\d{2})\.xlsx$', os.path.basename(cur))
    asof = datetime(2026, int(m.group(1)), int(m.group(2))) if m else datetime.now()
    cut = (asof - datetime(2026, 5, 1)).days
    a = prep(prev, '2025-05-01')
    b = prep(cur, '2026-05-01')
    last_signup = b['su'].max().strftime('%m/%d')
    data = {
        'asof': asof.strftime('%Y/%m/%d'), 'cut': cut, 'lastSignup': last_signup,
        'srcPrev': os.path.basename(prev), 'srcCur': os.path.basename(cur),
        'cats': [{'key': k, 'area': ar, 'type': t} for k, ar, t in CATS],
        'y25': pack(a), 'y26': pack(b),
    }
    html = open(os.path.join(BASE, 'yoy_template.html'), encoding='utf-8').read()
    html = html.replace('__DATA__', json.dumps(data, ensure_ascii=False))
    out = os.path.join(BASE, 'yoy_compare.html')
    open(out, 'w', encoding='utf-8').write(html)
    print('OK', out, '| 資料日', data['asof'], '| 開賣後第', cut, '天')


if __name__ == '__main__':
    main()
