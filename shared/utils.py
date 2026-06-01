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


def parse_op_days(op, order_no):
    """從 OP備註解析上課天數"""
    s = parse_op_section(op, order_no)
    for pat in [r'共\s*(\d+)\s*[日天晚泊]', r'上課\s*(\d+)\s*[日天]', r'(\d+)\s*[日天]\s*課程']:
        m = re.search(pat, s)
        if m:
            return int(m.group(1))
    return None


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
