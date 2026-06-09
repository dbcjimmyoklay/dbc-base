"""
agents/sales.py
SALES 銷量分析員
─────────────────────────────────────────────
職責：
  1. 讀取當季最新原始大總表 + 歷年大總表
  2. 計算各種統計（總覽、每日新增、出發分布、客群、加購）
  3. 識別回頭客（中文姓名+生日 OR 護照號碼）
  4. 與去年同期/同日比較（YoY）
  5. 輸出 portal/sales_data.json（純統計，不含個資）

檔案辨識：
  原始大總表XXXX.xlsx  → 當季最新（XXXX = 任意，取最新 mtime）
  原始大總表YY-YY.xlsx → 歷年資料（如 25-26）

觸發時機：
  納入 main.py --all 每日流程，CLEANER 後執行
  單獨：python main.py --sales
─────────────────────────────────────────────
"""

import os
import sys
import re
import json
import glob
from datetime import datetime
from collections import defaultdict

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_JSON = os.path.join(BASE_DIR, 'portal', 'sales_data.json')
AREAS       = ['野澤', '斑尾', '湯澤', '龍平']


# ═══════════════════════════════════════════
# 檔案辨識
# ═══════════════════════════════════════════

def find_xlsx_files():
    """掃描 BASE_DIR 找出當季 + 歷年大總表
    當季 = 不符 YY-YY 命名 的 原始大總表*.xlsx 中最新者
    歷年 = 符合 原始大總表YY-YY.xlsx
    """
    all_files = glob.glob(os.path.join(BASE_DIR, '原始大總表*.xlsx'))
    current_candidates = []
    historical = []

    for f in all_files:
        name = os.path.basename(f)
        m = re.search(r'原始大總表(\d\d)-(\d\d)\.xlsx$', name)
        if m:
            season = f"{m.group(1)}/{m.group(2)}"
            historical.append((season, f))
        else:
            current_candidates.append(f)

    current = max(current_candidates, key=os.path.getmtime) if current_candidates else None
    historical.sort(key=lambda x: x[0])
    return current, historical


def load_xlsx(path):
    """同 cleaner 邏輯：第2列為標題、第3列起為資料、重複欄位加 _N
    並對「團號」做 forward-fill（科威格式：同訂單只有第一人有團號）
    """
    df_raw = pd.read_excel(path, header=None)
    headers = df_raw.iloc[1].tolist()
    seen = {}
    new_headers = []
    for h in headers:
        s = str(h)
        if s in seen:
            seen[s] += 1
            new_headers.append(f"{s}_{seen[s]}")
        else:
            seen[s] = 1
            new_headers.append(s)
    df = df_raw.iloc[2:].copy()
    df.columns = new_headers
    df = df.reset_index(drop=True)
    if '團號' in df.columns:
        df = df[df['團號'] != '數量小計'].copy()
    # 同訂單第 2 人開始 團號 空白 → 用訂單編號補齊
    df = forward_fill_group_by_order(df)
    return df


# ═══════════════════════════════════════════
# 區域 / Level / 加購 判斷
# ═══════════════════════════════════════════

def get_area(g):
    """團號 → 雪場（與 config.py ITEM_RULES 一致）
    ZN=野澤純課, ZB=斑尾純課, ZY=湯澤純課,
    YS=龍平,  NT=野澤,  BT=斑尾,
    N=野澤,  B=斑尾
    """
    if pd.isna(g): return ''
    s = str(g)
    if 'ZN' in s: return '野澤'
    if 'ZB' in s: return '斑尾'
    if 'ZY' in s: return '湯澤'
    if 'YS' in s: return '龍平'
    if 'NT' in s: return '野澤'
    if 'BT' in s: return '斑尾'
    if 'N'  in s: return '野澤'
    if 'B'  in s: return '斑尾'
    return ''


def is_pure(g):
    if pd.isna(g): return False
    s = str(g)
    return 'ZN' in s or 'ZB' in s or 'ZY' in s


def forward_fill_group_by_order(df):
    """科威格式：團號只在「該團首列」出現，同團後續列團號空白
    （包含同訂單其他人 + 同團不同訂單的旅客）
    用「行順序 forward-fill」沿用上一列團號，直到遇到新團號。
    """
    if '團號' not in df.columns:
        return df
    df = df.copy()
    # 將空字串視為缺值
    df['團號'] = df['團號'].replace(r'^\s*$', pd.NA, regex=True)
    df['團號'] = df['團號'].ffill()
    return df


def norm_level(lv):
    if pd.isna(lv) or str(lv).strip() == '': return ''
    m = re.search(r'(\d+)', str(lv))
    return f'lv{m.group(1)}' if m else ''


def find_addon_cols(headers):
    H = [str(h) for h in headers]
    return {
        '續滑含教練': [h for h in H if ('續滑' in h or '加滑' in h) and '含' in h and '不含' not in h and '保留' not in h],
        '續滑不含':   [h for h in H if ('續滑' in h or '加滑' in h) and '不含' in h and '保留' not in h],
        '專屬教練':   [h for h in H if '專屬' in h],
        '租借衣褲':   [h for h in H if '租借' in h and '雪衣' in h],
        '租借護膝':   [h for h in H if '租借' in h and '護膝' in h],
        '租借護臀':   [h for h in H if '租借' in h and '護臀' in h],
        '租借手套':   [h for h in H if '租借' in h and '手套' in h],
    }


def row_has_value(row, cols):
    for c in cols:
        if c in row.index:
            v = row[c]
            if pd.notna(v) and v != 0 and str(v).strip() not in ('', '0'):
                return True
    return False


# ═══════════════════════════════════════════
# 回頭客 key
# ═══════════════════════════════════════════

def customer_keys(row):
    """產生可能的旅客唯一 key（任一相同 = 同一人，個資留在 Python 端不上 JSON）
    優先順位：身分證、護照、中文姓名+生日
    """
    keys = []
    id_no = str(row.get('身分證', '') or '').strip()
    if id_no and id_no.lower() != 'nan' and len(id_no) >= 5:
        keys.append(f"I:{id_no.upper()}")
    passport = str(row.get('護照號碼', '') or '').strip()
    if passport and passport.lower() != 'nan' and len(passport) >= 4:
        keys.append(f"P:{passport.upper()}")
    name = str(row.get('中文姓名', '') or '').strip()
    bday = str(row.get('生日', '') or '').strip()
    if name and bday and bday.lower() != 'nan' and name.lower() != 'nan':
        keys.append(f"NB:{name}|{bday}")
    return keys


# ═══════════════════════════════════════════
# 統計
# ═══════════════════════════════════════════

def compute_overview(df):
    df = df.copy()
    df['_area'] = df['團號'].apply(get_area)
    df['_pure'] = df['團號'].apply(is_pure)

    # 統計未分團的訂單（團號空白 → 科威尚未派團）
    unassigned = int((df['_area'] == '').sum())
    df['_fee']  = pd.to_numeric(
        df['團費'].astype(str).str.replace(',', '').str.replace(' ', ''),
        errors='coerce',
    )

    total_pax     = len(df)
    total_orders  = int(df['訂單編號'].nunique()) if '訂單編號' in df.columns else 0
    total_revenue = int(df['_fee'].sum())
    avg_fee       = int(df['_fee'].mean()) if total_pax > 0 else 0

    areas = {}
    for a in AREAS:
        sub = df[df['_area'] == a]
        if len(sub) == 0: continue
        pure = int(sub['_pure'].sum())
        group = len(sub) - pure
        areas[a] = {
            'total': len(sub),
            'group': group,
            'pure':  pure,
            'pct':   round(len(sub) / total_pax * 100, 1) if total_pax else 0,
        }

    return {
        'total_pax':     total_pax,
        'total_orders':  total_orders,
        'total_revenue': total_revenue,
        'avg_fee':       avg_fee,
        'areas':         areas,
        'unassigned':    unassigned,
    }


def compute_daily_signup(df):
    df = df.copy()
    df['_area']   = df['團號'].apply(get_area)
    df['_signup'] = pd.to_datetime(df['報名日期'], errors='coerce')

    grouped = df.groupby([df['_signup'].dt.date, '_area']).size().unstack(fill_value=0)
    result = []
    for d in sorted(grouped.index):
        if not pd.notna(d): continue
        row = {'d': str(d), 'total': 0}
        for a in AREAS:
            v = int(grouped.loc[d, a]) if a in grouped.columns else 0
            row[a] = v
            row['total'] += v
        result.append(row)
    return result


def compute_cumulative(daily_list):
    cum = 0
    out = []
    for r in daily_list:
        cum += r['total']
        out.append({'d': r['d'], 'c': cum})
    return out


def _parse_dep(v):
    if pd.isna(v): return None
    s = str(v).strip()
    for fmt in ('%y/%m/%d', '%Y/%m/%d', '%Y-%m-%d'):
        try: return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError): pass
    try: return pd.to_datetime(s)
    except (ValueError, TypeError): return None


def compute_departure(df):
    df = df.copy()
    df['_area'] = df['團號'].apply(get_area)
    df['_dep']  = df['出發日期'].apply(_parse_dep)

    grouped = df.groupby([df['_dep'].dt.date, '_area']).size().unstack(fill_value=0)
    result = []
    for d in sorted(grouped.index):
        if not pd.notna(d): continue
        row = {'d': str(d), 'total': 0}
        for a in AREAS:
            v = int(grouped.loc[d, a]) if a in grouped.columns else 0
            row[a] = v
            row['total'] += v
        result.append(row)
    return result


def compute_demographics(df):
    df = df.copy()
    df['_lv']  = df['LEVEL'].apply(norm_level) if 'LEVEL' in df.columns else ''
    df['_age'] = pd.to_numeric(df['年齡'], errors='coerce') if '年齡' in df.columns else None

    gender = {}
    if '性別' in df.columns:
        gender = {str(k): int(v) for k, v in df['性別'].value_counts().items() if str(k) not in ('nan', '')}

    ski = {}
    if '滑雪類別' in df.columns:
        ski = {str(k): int(v) for k, v in df['滑雪類別'].value_counts().items() if str(k) not in ('nan', '')}

    level_dist = {k: int(v) for k, v in df['_lv'].value_counts().items() if k}

    age_dist = {}
    if df['_age'] is not None and df['_age'].notna().any():
        bins   = [0, 18, 25, 30, 35, 40, 50, 100]
        labels = ['18以下','19-25','26-30','31-35','36-40','41-50','51+']
        age_cut = pd.cut(df['_age'], bins=bins, labels=labels, right=True)
        age_dist = {str(k): int(v) for k, v in age_cut.value_counts().sort_index().items()}

    return {
        'gender':   gender,
        'ski_type': ski,
        'level':    level_dist,
        'age':      age_dist,
    }


def compute_area_level(df):
    df = df.copy()
    df['_area'] = df['團號'].apply(get_area)
    df['_lv']   = df['LEVEL'].apply(norm_level) if 'LEVEL' in df.columns else ''

    result = {a: {} for a in AREAS}
    for area in AREAS:
        sub = df[(df['_area'] == area) & (df['_lv'] != '')]
        for lv, cnt in sub['_lv'].value_counts().items():
            result[area][lv] = int(cnt)
    return result


def compute_addons(df):
    addon_cols = find_addon_cols(df.columns)
    result = {}
    for name, cols in addon_cols.items():
        if not cols:
            result[name] = 0
            continue
        result[name] = int(df.apply(lambda r: row_has_value(r, cols), axis=1).sum())
    return result


# ═══════════════════════════════════════════
# 回頭客
# ═══════════════════════════════════════════

def compute_returning(current_df, historical_dfs, historical_seasons, current_season):
    """
    跨年 person index（union-find）：身分證 / 護照 / 中文姓名+生日 任一相同 = 同一人
    輸出：
      stats: 本季回頭客率（與歷年比對）
      customers: 報名次數 >= 2 的旅客明細（聚合各次報名，不含個資）
    """
    # ── union-find person index ──
    key_to_pid = {}
    persons    = {}   # pid → {'name': str, 'signups': list}
    pid_counter = [0]

    def assign_person(keys):
        if not keys:
            return None
        found_pids = set()
        for k in keys:
            if k in key_to_pid:
                found_pids.add(key_to_pid[k])

        if not found_pids:
            pid_counter[0] += 1
            pid = pid_counter[0]
            persons[pid] = {'name': '', 'signups': []}
        elif len(found_pids) == 1:
            pid = next(iter(found_pids))
        else:
            # 多個 person 共有此 record 的 key → 合併
            pid = min(found_pids)
            for other in found_pids:
                if other == pid: continue
                persons[pid]['signups'].extend(persons[other]['signups'])
                if not persons[pid]['name'] and persons[other]['name']:
                    persons[pid]['name'] = persons[other]['name']
                del persons[other]
                for k_, v_ in list(key_to_pid.items()):
                    if v_ == other:
                        key_to_pid[k_] = pid
        for k in keys:
            key_to_pid[k] = pid
        return pid

    def process_df(df, season):
        df = df.copy()
        # 23/24 雪季當年僅有野澤一個雪場 → 強制全部歸野澤
        # 團號規則從 24/25 開始啟用
        if season == '23/24':
            df['_area'] = '野澤'
        else:
            df['_area'] = df['團號'].apply(get_area)
        df['_fee']  = pd.to_numeric(
            df['團費'].astype(str).str.replace(',', '').str.replace(' ', ''),
            errors='coerce',
        )
        for _, row in df.iterrows():
            keys = customer_keys(row)
            pid = assign_person(keys)
            if pid is None:   # 無任何 key 的記錄無法配對，跳過
                continue
            name = str(row.get('中文姓名', '') or '').strip()
            if name and not persons[pid]['name']:
                persons[pid]['name'] = name
            persons[pid]['signups'].append({
                'season': season,
                'area':   row.get('_area', '') or '未派團',
                'date':   str(row.get('出發日期', '') or '').strip(),
                'ski':    str(row.get('滑雪類別', '') or '').strip(),
                'fee':    int(row['_fee']) if pd.notna(row['_fee']) else 0,
            })

    # 處理所有資料：歷年 + 本季
    for season, df in zip(historical_seasons, historical_dfs):
        process_df(df, season)
    process_df(current_df, current_season)

    # ── 篩選次數 >= 2 的旅客 ──
    customers = []
    for pid, p in persons.items():
        if len(p['signups']) < 2:
            continue
        signups = sorted(p['signups'], key=lambda s: s['date'], reverse=True)
        customers.append({
            'name':          p['name'] or '?',
            'total_signups': len(signups),
            'total_fee':     sum(s['fee'] for s in signups),
            'signups':       signups,
        })
    # 依次數降冪、團費降冪、姓名排序
    customers.sort(key=lambda c: (-c['total_signups'], -c['total_fee'], c['name']))

    # ── 本季回頭客率（與歷年比對，原邏輯）──
    historical_year_count = defaultdict(int)
    for hist_df in historical_dfs:
        seen_this_year = set()
        for _, row in hist_df.iterrows():
            for k in customer_keys(row):
                if k not in seen_this_year:
                    seen_this_year.add(k)
                    historical_year_count[k] += 1

    cur_copy = current_df.copy()
    cur_copy['_area'] = cur_copy['團號'].apply(get_area)
    returning_count = 0
    by_area = defaultdict(int)
    times_dist = defaultdict(int)
    seen_visitors_this_season = set()
    seen_returning_visitors   = set()

    for _, row in cur_copy.iterrows():
        keys = customer_keys(row)
        if not keys: continue
        primary_key = keys[0]
        seen_visitors_this_season.add(primary_key)

        matched_years = 0
        for k in keys:
            matched_years = max(matched_years, historical_year_count.get(k, 0))
        if matched_years > 0:
            if primary_key not in seen_returning_visitors:
                seen_returning_visitors.add(primary_key)
                returning_count += 1
                area = row.get('_area', '')
                if area: by_area[area] += 1
                total_visits = matched_years + 1
                label = '4+' if total_visits >= 4 else str(total_visits)
                times_dist[label] += 1

    total_unique = len(seen_visitors_this_season)
    return {
        'count':        returning_count,
        'total_unique': total_unique,
        'pct':          round(returning_count / total_unique * 100, 1) if total_unique else 0,
        'by_area':      dict(by_area),
        'times_dist':   dict(times_dist),
        'customers':    customers,
    }


# ═══════════════════════════════════════════
# YoY 比較
# ═══════════════════════════════════════════

def _to_season_day(d, season_start_month=5):
    """報名日期 → 開賣後第幾天
    開賣日 = 雪季年的 5/1（25/26 雪季開賣日 = 2025/5/1）
    5/1 為 day 0
    """
    if pd.isna(d): return None
    if d.month >= season_start_month:
        season_start = pd.Timestamp(d.year, season_start_month, 1)
    else:
        season_start = pd.Timestamp(d.year - 1, season_start_month, 1)
    return (d - season_start).days


def _to_day_in_season(d, season_label, season_start_month=5):
    """報名日期 → 指定雪季的開賣後第幾天
    season_label: '25/26' 等
    5/1 之前的早鳥報名 → 歸到 day 0
    """
    if pd.isna(d): return None
    try:
        yy = int(str(season_label).split('/')[0]) + 2000
    except (ValueError, IndexError):
        return _to_season_day(d, season_start_month)
    season_start = pd.Timestamp(yy, season_start_month, 1)
    delta = (d - season_start).days
    return max(0, delta)   # 早鳥（< 5/1）全部歸 day 0


def compute_yoy(current_df, historical_dfs, historical_seasons, current_season):
    """
    產生與每個歷年的對比：
    - 同檔期累積：兩條線都截到本季當前累積天數（X 軸 = 開賣後第 N 天，5/1 起算）
      （5/1 之前的早鳥報名 → 全部歸到 day 0）
    - 同日比：按雪季月份排序（5→12→1→4），出發日同 月/日 對照
    回傳：{ cur_max_day: int, comparisons: [ {season, prev_total, same_period, same_date}, ... ] }
    """
    if not historical_dfs:
        return None

    # ── 本季資料預處理（用正確的雪季 5/1 為 day 0） ──
    cur_signup_days = (
        pd.to_datetime(current_df['報名日期'], errors='coerce')
          .dropna().apply(lambda d: _to_day_in_season(d, current_season))
    )
    cur_days_count = cur_signup_days.value_counts().sort_index()
    cur_max_day = int(cur_days_count.index.max()) if len(cur_days_count) else 0
    cur_total   = int(cur_days_count.sum())

    cur_dep_md = (
        current_df['出發日期'].apply(_parse_dep).dropna()
            .apply(lambda d: f"{d.month:02d}/{d.day:02d}")
    )
    cur_md_count = cur_dep_md.value_counts()

    # ── 雪季月份排序 key（5/1 起算）──
    def season_md_key(md):
        m, d = md.split('/')
        return ((int(m) - 5) % 12, int(d))

    comparisons = []
    for hist_df, season in zip(historical_dfs, historical_seasons):
        # 用該歷年雪季的 5/1 為 day 0（早於 5/1 的早鳥 → 全部歸 day 0）
        prev_signup_days = (
            pd.to_datetime(hist_df['報名日期'], errors='coerce')
              .dropna().apply(lambda d: _to_day_in_season(d, season))
        )
        prev_days_count = prev_signup_days.value_counts().sort_index()
        prev_total = int(prev_days_count.sum())

        # 兩條線都跑到本季當前 day（同檔期）
        cur_cum, prev_cum = [], []
        c = p = 0
        for d in range(cur_max_day + 1):
            c += int(cur_days_count.get(d, 0))
            p += int(prev_days_count.get(d, 0))
            cur_cum.append(c)
            prev_cum.append(p)

        prev_dep_md = (
            hist_df['出發日期'].apply(_parse_dep).dropna()
                .apply(lambda d: f"{d.month:02d}/{d.day:02d}")
        )
        prev_md_count = prev_dep_md.value_counts()

        all_md = sorted(set(cur_md_count.index) | set(prev_md_count.index), key=season_md_key)

        comparisons.append({
            'season': season,
            'prev_total': prev_total,
            'cur_total':  cur_total,
            'same_period': {
                'days':    list(range(cur_max_day + 1)),
                'current': cur_cum,
                'prev':    prev_cum,
            },
            'same_date': {
                'labels':  all_md,
                'current': [int(cur_md_count.get(m, 0))  for m in all_md],
                'prev':    [int(prev_md_count.get(m, 0)) for m in all_md],
            },
        })

    return {
        'cur_max_day': cur_max_day,
        'cur_total':   cur_total,
        'comparisons': comparisons,
    }


# ═══════════════════════════════════════════
# 主程式
# ═══════════════════════════════════════════

def run():
    print("=" * 50)
    print("SALES 銷量分析員 啟動")
    print("=" * 50)

    current_path, historical = find_xlsx_files()
    if not current_path:
        print("找不到當季原始大總表，跳過")
        return

    print(f"當季：{os.path.basename(current_path)}")
    print(f"歷年：{[s for s, _ in historical]}")

    print("\n載入當季資料...")
    current_df = load_xlsx(current_path)
    print(f"  {len(current_df)} 筆")

    print("\n載入歷年資料...")
    historical_dfs = []
    historical_seasons = []
    for season, path in historical:
        df = load_xlsx(path)
        historical_dfs.append(df)
        historical_seasons.append(season)
        print(f"  {season}: {len(df)} 筆")

    print("\n計算統計...")
    overview     = compute_overview(current_df)
    daily        = compute_daily_signup(current_df)
    cumulative   = compute_cumulative(daily)
    departure    = compute_departure(current_df)
    demographics = compute_demographics(current_df)
    area_level   = compute_area_level(current_df)
    addons       = compute_addons(current_df)

    # 先推算當季標籤（後面 output 也會用）
    today    = datetime.now()
    cur_year = today.year - 2000
    current_season = f"{cur_year}/{cur_year+1}" if today.month >= 5 else f"{cur_year-1}/{cur_year}"

    print("計算回頭客...")
    returning = compute_returning(current_df, historical_dfs, historical_seasons, current_season)
    print(f"  本季不重複旅客：{returning['total_unique']}")
    print(f"  本季回頭客：{returning['count']} 人 ({returning['pct']}%)")
    print(f"  跨年總次數 >= 2 的旅客：{len(returning['customers'])} 位")

    print("計算 YoY 對比...")
    yoy = compute_yoy(current_df, historical_dfs, historical_seasons, current_season)

    # current_season 已於 compute_returning 前推算

    output = {
        'updated':            datetime.now().strftime('%Y/%m/%d %H:%M'),
        'current_season':     current_season,
        'historical_seasons': historical_seasons,
        'overview':           overview,
        'daily_signup':       daily,
        'cumulative':         cumulative,
        'departure':          departure,
        'demographics':       demographics,
        'area_level':         area_level,
        'addons':             addons,
        'returning':          returning,
        'yoy':                yoy,
    }

    os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, separators=(',', ':'))

    size_kb = os.path.getsize(OUTPUT_JSON) / 1024
    print(f"\nOK 輸出 {os.path.basename(OUTPUT_JSON)} ({size_kb:.1f} KB)")
    print("=" * 50)


if __name__ == '__main__':
    run()
