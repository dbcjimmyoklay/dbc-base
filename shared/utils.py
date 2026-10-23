"""
shared/utils.py
DBC 影子中台 — 共用工具函數
"""

import re
import pandas as pd
from datetime import datetime
from .config import (
    WEEKDAY_MAP, ITEM_RULES, PURE_ITEMS,
    PRICE_TABLE, YUKIGUNI_PEAK, EQUIP_PRICE,
)


# ===== 基本清洗 =====

def clean_val(x):
    """清洗單一值：None/NaN/空字串/0 一律回傳空字串"""
    if x is None:
        return ''
    if hasattr(x, 'iloc'):
        x = x.iloc[0] if len(x) > 0 else ''
    try:
        if pd.isna(x):
            return ''
    except Exception:
        pass
    s = str(x).strip()
    if s.lower() in ('nan', 'none', ''):
        return ''
    if s in ('0', '0.0'):
        return ''
    if s.endswith('.0') and s[:-2].lstrip('-').isdigit():
        return s[:-2]
    return s


def clean_seq_no(x):
    """清洗序號欄位：保留 0（0 是有意義的代碼），只去除 None/NaN/空字串"""
    if x is None:
        return ''
    if hasattr(x, 'iloc'):
        x = x.iloc[0] if len(x) > 0 else ''
    try:
        if pd.isna(x):
            return ''
    except Exception:
        pass
    s = str(x).strip()
    if s.lower() in ('nan', 'none', ''):
        return ''
    if s.endswith('.0') and s[:-2].lstrip('-').isdigit():
        return s[:-2]
    return s


def safe_int(val):
    """轉整數；失敗或為 0 回傳空字串"""
    try:
        if val is None:
            return ''
        try:
            if pd.isna(val):
                return ''
        except Exception:
            pass
        v = pd.to_numeric(val, errors='coerce')
        if pd.isna(v):
            return ''
        i = int(v)
        return '' if i == 0 else i
    except Exception:
        return ''


def clean_order_no(val):
    """訂單編號去前導零"""
    try:
        return str(int(str(val).lstrip('0')))
    except Exception:
        return str(val)


def cn_to_num(s):
    """中文數字轉阿拉伯數字"""
    for k, v in {'一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
                 '六': '6', '七': '7', '八': '8', '九': '9', '零': '0'}.items():
        s = s.replace(k, v)
    return s


def has_value(row, col):
    """欄位是否有有效值"""
    if col not in row.index:
        return False
    v = row[col]
    return pd.notna(v) and v != 0 and v != ''


# ===== 日期工具 =====

def get_weekday(date):
    """datetime → 中文曜日"""
    return WEEKDAY_MAP.get(date.weekday(), '') if date else ''


def extract_date_from_group(g):
    """從科威團號字串中抽取日期（格式：YYMMDD）"""
    if pd.isna(g):
        return None
    m = re.search(r'(\d{6})', str(g))
    if m:
        d = m.group(1)
        try:
            return datetime(2000 + int(d[:2]), int(d[2:4]), int(d[4:6]))
        except Exception:
            return None
    return None


# ===== 雪場判斷 =====

def get_item(g):
    """從科威團號判斷雪場名稱"""
    if pd.isna(g):
        return '其他'
    s = str(g)
    for key, label in ITEM_RULES:
        if key in s:
            return label
    return '其他'


# ===== LEVEL 標準化 =====

def normalize_level(lv):
    """統一 LEVEL 格式為 lv0~lv4；空值回傳空字串"""
    if not lv or str(lv).strip() in ('', 'nan', 'None'):
        return ''
    s = str(lv).strip().lower()
    m = re.search(r'(\d+)', s)
    if m:
        return f"lv{int(m.group(1))}"
    return ''


# ===== 飲食代碼清洗 =====

def clean_diet(raw):
    """清洗飲食備註欄位，只保留有意義的標籤"""
    if raw is None:
        return ''
    if hasattr(raw, 'iloc'):
        raw = raw.iloc[0] if len(raw) > 0 else ''
    try:
        if pd.isna(raw):
            return ''
    except Exception:
        pass
    s = str(raw).strip()
    if not s or s.lower() in ('nan', 'none'):
        return ''

    CODE_MAP = {
        'VLML': '奶蛋素', 'NBML': '非牛肉餐', 'AVML': '東方素食',
        'CHML': '兒童餐',  'GFML': '無麩質',   'LCML': '低卡',
        'LSML': '低鹽',   'NLML': '無乳糖',
    }
    SKIP_CODES = {'NAML'}

    codes = re.findall(r'\b([A-Z]{4})\b', s)
    results, seen = [], set()
    for code in codes:
        if code in SKIP_CODES:
            continue
        label = CODE_MAP.get(code)
        if label and label not in seen:
            results.append(label)
            seen.add(label)

    if not results:
        free = re.sub(r'[A-Z]{4}[_\s－/／]*[^\s,，A-Z]*', '', s).strip().strip(',，／/').strip()
        if free and free not in ('', '無指定'):
            results.append(free)

    return '／'.join(results) if results else ''


# ===== OP備註解析（純課程用）=====

def parse_op_section(op, order_no):
    if not op or pd.isna(op):
        return ''
    try:
        ono = str(int(str(order_no).lstrip('0')))
    except Exception:
        ono = str(order_no)
    text = cn_to_num(str(op))
    m = re.search(rf'#\s*0*{re.escape(ono)}\b[\s\S]*?(?=\n#\d|\Z)', text)
    return m.group() if m else text


def op_has_section(op, order_no):
    """OP備註文字裡是否有這張訂單自己的 #訂單編號 段落"""
    if not op or pd.isna(op):
        return False
    try:
        ono = str(int(str(order_no).lstrip('0')))
    except Exception:
        ono = str(order_no)
    return bool(re.search(rf'#\s*0*{re.escape(ono)}\b', cn_to_num(str(op))))


def pick_op(own_op, group_op, order_no):
    """科威把整個團號的 OP備註全放在該團第一列（內含多張訂單的 #段落）。
    訂單自己第一列沒有自己的段落時，改從同團號合併的 OP備註找。"""
    if op_has_section(own_op, order_no):
        return own_op
    if op_has_section(group_op, order_no):
        return group_op
    return own_op


def parse_op_days(op, order_no):
    """從 OP備註解析上課天數（保留半天，例如 1.5、2.5；整數回傳 int）"""
    s = parse_op_section(op, order_no)
    for pat in [r'共\s*(\d+(?:\.\d+)?)\s*[日天晚泊]', r'上課\s*(\d+(?:\.\d+)?)\s*[日天]',
                r'(\d+(?:\.\d+)?)\s*[日天]\s*課程']:
        m = re.search(pat, s)
        if m:
            v = float(m.group(1))
            return int(v) if v.is_integer() else v
    return None


# OP備註天數/師生比 vs 團費 的一致性檢查（OP 為準，差異過大只標示、不改資料）
OP_FEE_ABS_TOL = 5000     # 差額超過此金額
OP_FEE_PCT_TOL = 0.15     # 且超過預期課程費的此比例，才視為異常


def check_op_fee(item, days, ratio, members, lesson_date=None, season_start_year=2026):
    """
    以 OP 的天數/師生比算預期課程費 =（單日價 + 湯澤旺季加價）× 天數，
    與實收（訂單團費合計 − 租借裝備）比較。裝備租借日數不一定等於上課日數，
    所以對每種裝備的租借日數（0~7）取最接近的組合。
    回傳 (差額, 預期課程費, 訂單團費合計, is_anomaly)；查不到價目回傳 None。
    """
    import itertools
    price = PRICE_TABLE.get(item, {}).get(ratio)
    if not price or not days:
        return None
    sur = (get_yukiguni_surcharge(lesson_date, season_start_year)
           if item == '湯澤純課' and lesson_date else 0)
    expected = (price + sur) * days
    total = sum(m['fee'] for m in members)
    cnt = {k: sum(1 for m in members if m['has_equip'].get(k)) for k in EQUIP_PRICE}
    flat = cnt.get('手套', 0) * EQUIP_PRICE['手套']
    per_day = [(k, cnt[k] * EQUIP_PRICE[k]) for k in ('衣褲', '護膝', '護臀', '4件組') if cnt.get(k)]
    best = None
    for dys in itertools.product(range(8), repeat=len(per_day)):
        eq = flat + sum(d * p for d, (_, p) in zip(dys, per_day))
        diff = total - eq - expected
        if best is None or abs(diff) < abs(best):
            best = diff
    best = best if best is not None else total - flat - expected
    anomaly = abs(best) > OP_FEE_ABS_TOL and abs(best) / expected > OP_FEE_PCT_TOL
    return best, expected, total, anomaly


def parse_op_ratio(op, order_no):
    """從 OP備註解析師生比"""
    s = parse_op_section(op, order_no)
    m = re.search(r'1\s*[對对:]\s*(\d+)', s)
    return f"1對{m.group(1)}" if m else None


# ===== 純課程費用反推 =====

def get_yukiguni_surcharge(lesson_date, year):
    """湯澤旺季附加費"""
    if not lesson_date:
        return 0
    for (m1, d1, m2, d2, surcharge) in YUKIGUNI_PEAK:
        y1 = year if m1 >= 12 else year + 1
        y2 = year if m2 >= 12 else year + 1
        if datetime(y1, m1, d1) <= lesson_date <= datetime(y2, m2, d2):
            return surcharge
    return 0


def calc_person_equip_fee(has_equip, days):
    """計算單人裝備費"""
    fee = 0
    for key, price in EQUIP_PRICE.items():
        if has_equip.get(key, False):
            fee += price if key == '手套' else price * days
    return fee


def infer_days_from_fee(item, order_members, lesson_date=None, season_start_year=2026):
    """從團費反推天數與師生比"""
    if item not in PRICE_TABLE:
        return None, None
    n_people = len(order_members)
    if n_people < 1 or n_people > 6:
        return None, None
    surcharge = (
        get_yukiguni_surcharge(lesson_date, season_start_year)
        if item == '湯澤純課' and lesson_date else 0
    )
    for ratio in range(1, 7):
        base_price = PRICE_TABLE[item].get(ratio, 0)
        day_price = base_price + surcharge
        if not day_price:
            continue
        for days in range(1, 8):
            total = sum(
                m['fee'] - calc_person_equip_fee(m['has_equip'], days)
                for m in order_members
            )
            if abs(total - day_price * days) < 1:
                return days, ratio
    return None, None


# ===== 欄位分類（從 df.columns 解析科威加購欄）=====

def classify_cols(cols):
    """
    解析科威原始欄位，找出：
      pure       → 純課程欄（含師生比）
      excl       → 專屬教練欄
      with_      → 續滑含教練欄（天數）
      without_   → 續滑不含教練欄（天數）
      four       → 4件組欄
      rent       → 租借裝備欄
      keep_city  → 保留市區住宿欄（天數+N，泊數不變）
    """
    pure, excl, with_, without_, four = {}, [], {}, {}, []
    rent = {'衣褲': [], '護膝': [], '護臀': [], '手套': []}
    keep_city = {}

    for col in cols:
        c = str(col)
        if '純課程' in c:
            m = re.search(r'1[對对]\s*(\d+)', c)
            if m:
                pure[col] = int(m.group(1))
        if '專屬教練' in c:
            excl.append(col)
        if ('續滑' in c or '加滑' in c) and '含教練' in c and '不含教練' not in c and '保留市區住宿' not in c:
            m = re.search(r'(\d+)\s*[日天]', c)
            if m:
                with_[col] = int(m.group(1))
        if ('續滑' in c or '加滑' in c) and '不含教練' in c and '保留市區住宿' not in c:
            m = re.search(r'(\d+)\s*[日天]', c)
            if m:
                without_[col] = int(m.group(1))
        if '保留市區住宿' in c:
            m = re.search(r'(\d+)\s*[日天]', c)
            days = int(m.group(1)) if m else 1
            keep_city[col] = days
        if '4件組' in c or '四件組' in c:
            four.append(col)
        if '租借' in c:
            if '雪衣' in c:
                rent['衣褲'].append(col)
            if '護膝' in c:
                rent['護膝'].append(col)
            if '護臀' in c:
                rent['護臀'].append(col)
            if '手套' in c:
                rent['手套'].append(col)

    return pure, excl, with_, without_, four, rent, keep_city


# ===== 檔案選擇（原始大總表*.xlsx）=====

def file_sort_key(path):
    """
    判斷「哪個原始大總表檔案最新」的排序 key。
    優先使用檔名裡的日期（如 原始大總表0819.xlsx → 08/19），
    只有在檔名無法解析時才退回用檔案修改時間（mtime）。

    原因：瀏覽器下載檔案時，有時會保留伺服器的 Last-Modified 標頭
    而非「實際下載時間」，單靠 mtime 判斷最新檔案並不可靠，
    曾造成系統誤用舊檔案跑出過期資料的問題。

    雪季以 7 月為切分點：8~12 月視為早半季，1~6 月視為晚半季（隔年），
    避免跨年比較（例如 12月 vs 隔年1月）時順序顛倒。
    """
    import re as _re
    import os as _os

    m = _re.search(r'原始大總表(\d{2})(\d{2})\.xlsx$', _os.path.basename(path))
    if m:
        mm, dd = int(m.group(1)), int(m.group(2))
        if 1 <= mm <= 12 and 1 <= dd <= 31:
            season_bucket = 0 if mm >= 7 else 1
            return (1, (season_bucket, mm, dd), _os.path.getmtime(path))
    return (0, (0, 0, 0), _os.path.getmtime(path))


# ===== 比對 Key =====

def build_record_key(r, seq=None):
    """
    建立唯一識別 key
    優先順序：旅客編號 > 訂編+中文姓名 > 訂編+英文姓名 > 訂編+序號
    """
    visitor_id = str(r.get('旅客編號', '')).strip()
    if visitor_id:
        return f"v_{visitor_id}"
    ding = str(r.get('訂編', '')).strip()
    cn   = str(r.get('中文姓名', '')).strip()
    en   = str(r.get('英文姓名', '')).strip()
    if cn:
        return f"{ding}_{cn}"
    if en:
        return f"{ding}_en_{en}"
    return f"{ding}_seq_{seq or 0}"
