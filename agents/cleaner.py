"""
agents/cleaner.py
CLEANER 清洗員
─────────────────────────────────────────────
職責：
  1. 讀取 DBC base 資料夾中最新的 原始大總表*.xlsx
  2. 清洗、整理、計算所有欄位
  3. 將結果序列化存成 .pkl，供其他工作員使用

輸出（cleaner_output.pkl）：
  all_records_flat  → list[dict]  純資料列（不含分隔列）
  all_sheet_rows    → list[dict]  含分隔列，供寫入 Google Sheet 用
  date_map          → dict        出發日期 → list[dict]
  DATE_START        → datetime
  DATE_END          → datetime
  headers           → list        科威原始欄位名稱（供 REPORTER 用）
  df_raw_records    → list[dict]  科威原始 df 轉成 records（供 REPORTER 用）
─────────────────────────────────────────────
"""

import os
import glob
import pickle
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict

# 使用相對 import（從 main.py 執行時）或絕對 import（直接執行時）
try:
    from shared.config import (
        PURE_ITEMS, ITEM_ORDER, WEEKDAY_MAP,
        ROOM_EXCLUDE_ITEMS,
    )
    from shared.utils import (
        get_item, get_weekday, extract_date_from_group,
        has_value, clean_val, safe_int, clean_order_no, clean_diet,
        normalize_level, classify_cols,
        parse_op_days, parse_op_ratio,
        infer_days_from_fee,
        build_record_key,
    )
except ImportError:
    import sys
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from shared.config import (
        PURE_ITEMS, ITEM_ORDER, WEEKDAY_MAP,
        ROOM_EXCLUDE_ITEMS,
    )
    from shared.utils import (
        get_item, get_weekday, extract_date_from_group,
        has_value, clean_val, safe_int, clean_order_no, clean_diet,
        normalize_level, classify_cols,
        parse_op_days, parse_op_ratio,
        infer_days_from_fee,
        build_record_key,
    )

BASE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_PATH = os.path.join(BASE_DIR, 'cleaner_output.pkl')


# ─────────────────────────────────────────────
# 主函數
# ─────────────────────────────────────────────

def run():
    print("=" * 50)
    print("CLEANER 清洗員 啟動")
    print("=" * 50)

    # ── 1. 找最新 xlsx ──
    # 本機優先；若本機找不到（例如在 GitHub Actions 執行），從 Google Drive 下載
    files = glob.glob(os.path.join(BASE_DIR, '原始大總表*.xlsx'))
    if not files:
        print("本機找不到原始大總表*.xlsx，嘗試從 Google Drive 下載...")
        from shared.drive import download_latest_xlsx
        download_latest_xlsx(dest_dir=BASE_DIR)
        files = glob.glob(os.path.join(BASE_DIR, '原始大總表*.xlsx'))
        if not files:
            raise FileNotFoundError("Google Drive 下載後仍找不到原始大總表*.xlsx")

    latest_file = max(files, key=os.path.getmtime)
    print(f"使用檔案：{os.path.basename(latest_file)}")

    # ── 2. 讀取 xlsx ──
    df_raw = pd.read_excel(latest_file, header=None)
    headers = df_raw.iloc[1].tolist()

    # 重複欄位改名（序號→序號_1/序號_2，分房備住→分房備住_1/分房備住_2）
    seen_headers = {}
    new_headers  = []
    for h in headers:
        h_str = str(h)
        if h_str in seen_headers:
            seen_headers[h_str] += 1
            new_headers.append(f"{h_str}_{seen_headers[h_str]}")
        else:
            seen_headers[h_str] = 1
            new_headers.append(h_str)

    df = df_raw.iloc[2:].copy()
    df.columns = new_headers
    df = df.reset_index(drop=True)
    df = df[df['團號'] != '數量小計'].copy()
    df['團號'] = df['團號'].ffill()
    df = df[df['訂單編號'].notna()].copy()
    df = df.reset_index(drop=True)
    print(f"讀取完成，共 {len(df)} 筆資料")

    # ── 3. 解析科威加購欄位 ──
    pure_cols, excl_cols, with_cols, without_cols, four_cols, rent_cols, keep_city_cols = \
        classify_cols(df.columns)
    if keep_city_cols:
        print(f"偵測到保留市區住宿欄位：{list(keep_city_cols.keys())}")

    # ── 4. 雪季偵測 ──
    try:
        dates = pd.to_datetime(df['出發日期'], errors='coerce').dropna()
        y = dates.min().year if dates.min().month >= 10 else dates.min().year - 1
        DATE_START = datetime(y, 12, 5)
        DATE_END   = datetime(y + 1, 3, 31)
    except Exception:
        DATE_START = datetime(2026, 12, 5)
        DATE_END   = datetime(2027, 3, 31)
    print(f"雪季：{DATE_START.strftime('%Y/%m/%d')} ~ {DATE_END.strftime('%Y/%m/%d')}")

    # ── 5. 純課程：OP 解析 + 費用反推 ──
    order_op_cache = {}
    for order_no, group in df.groupby('訂單編號', sort=False):
        first_row = group.iloc[0]
        item = get_item(first_row.get('團號', ''))
        if item not in PURE_ITEMS:
            continue
        op          = first_row.get('OP備註', '')
        days_val    = parse_op_days(op, order_no)
        ratio_val   = parse_op_ratio(op, order_no)
        tiansu      = str(days_val) if days_val else ''
        note        = ratio_val or ''
        is_excl     = any(has_value(first_row, c) for c in excl_cols)

        if is_excl:
            note = '專屬'
        elif not tiansu:
            members = []
            for _, mrow in group.iterrows():
                fee = safe_int(mrow.get('團費', 0))
                if fee == '':
                    fee = 0
                has_equip = {
                    '衣褲': any(has_value(mrow, c) for c in rent_cols['衣褲']),
                    '護膝': any(has_value(mrow, c) for c in rent_cols['護膝']),
                    '護臀': any(has_value(mrow, c) for c in rent_cols['護臀']),
                    '手套': any(has_value(mrow, c) for c in rent_cols['手套']),
                    '4件組': any(has_value(mrow, c) for c in four_cols),
                }
                members.append({'fee': fee, 'has_equip': has_equip})
            group_date = extract_date_from_group(first_row.get('團號', ''))
            inferred_days, n_people = infer_days_from_fee(
                item, members,
                lesson_date=group_date,
                season_start_year=DATE_START.year,
            )
            if inferred_days:
                tiansu = str(inferred_days)
                if not note and n_people:
                    note = f"1對{n_people}"
            else:
                tiansu = ''
                if not note:
                    note = '待確認'

        order_op_cache[str(order_no)] = (tiansu, note)

    # ── 6. 建立完整記錄 ──
    records = []
    for _, row in df.iterrows():
        group_no = row.get('團號', '')
        item     = get_item(group_no)
        is_pure  = item in PURE_ITEMS

        group_date = extract_date_from_group(group_no)
        if group_date is None:
            try:
                group_date = pd.to_datetime(row['出發日期'])
            except Exception:
                group_date = None
        try:
            group_date = None if pd.isna(group_date) else group_date
        except Exception:
            pass

        depart_date    = (group_date - timedelta(days=1)) if (is_pure and group_date) else group_date
        sort_date      = depart_date if depart_date else datetime(2099, 1, 1)
        depart_display = depart_date.strftime('%Y/%m/%d') if depart_date else ''

        order_no = row.get('訂單編號', '')

        if is_pure:
            pou_su = ''
            tiansu, note = order_op_cache.get(str(order_no), ('', ''))
        else:
            is_excl     = any(has_value(row, c) for c in excl_cols)
            add_with    = sum(n for col, n in with_cols.items()    if has_value(row, col))
            add_without = sum(n for col, n in without_cols.items() if has_value(row, col))
            add_keep    = sum(n for col, n in keep_city_cols.items() if has_value(row, col))
            pou_su      = 3 + add_with + add_without
            total_days  = 2 + add_with + add_without + add_keep
            has_含      = add_with > 0
            has_無      = (add_without + add_keep) > 0
            if has_含 and has_無:
                tiansu = f"{total_days}含無"
            elif has_含:
                tiansu = f"{total_days}含"
            elif has_無:
                tiansu = f"{total_days}無"
            else:
                tiansu = str(total_days)
            note = '專屬' if is_excl else ''

        records.append({
            '_sort_date'   : sort_date,
            '_item_order'  : ITEM_ORDER.get(item, 7),
            '_depart_date' : depart_date,
            '項目'         : item,
            '出發日期'     : depart_display,
            '泊數'         : str(pou_su) if pou_su != '' else '',
            '天數'         : tiansu,
            '註記'         : note,
            '訂編'         : clean_order_no(order_no),
            '中文姓名'     : clean_val(row.get('中文姓名', '')),
            '英文姓名'     : clean_val(row.get('英文姓名', '')),
            '性別'         : clean_val(row.get('性別', '')),
            '年齡'         : str(safe_int(row.get('年齡'))),
            '尺寸'         : '',
            '身高'         : str(safe_int(row.get('身高(CM)'))),
            '體重'         : str(safe_int(row.get('體重(KG)'))),
            '腳長'         : clean_val(row.get('腳長', '')),
            '雪板'         : '',
            'LEVEL'        : normalize_level(clean_val(row.get('LEVEL', ''))),
            '類別'         : clean_val(row.get('滑雪類別', '')),
            '衣褲'         : '1' if any(has_value(row, c) for c in rent_cols['衣褲']) else '',
            '護膝'         : '1' if any(has_value(row, c) for c in rent_cols['護膝']) else '',
            '護臀'         : '1' if any(has_value(row, c) for c in rent_cols['護臀']) else '',
            '手套'         : '1' if any(has_value(row, c) for c in rent_cols['手套']) else '',
            '4件組'        : '1' if any(has_value(row, c) for c in four_cols) else '',
            '教練'         : '',
            '助教'         : '',
            'checking'     : '',
            '備註'         : '',
            '旅客編號'     : clean_val(row.get('旅客編號', '')),
            '項目1'        : item,
            # 個資（只存本機）
            'OP備註'       : clean_val(row.get('OP備註', '')),
            '報名日期'     : clean_val(row.get('報名日期', '')),
            '團費'         : clean_val(row.get('團費', '')),
            '手機號碼'     : clean_val(row.get('手機號碼', '')),
            'EMail'        : clean_val(row.get('EMail', '')),
            '國籍'         : clean_val(row.get('國籍', '')),
            '身分證'       : clean_val(row.get('身分證', '')),
            '生日'         : clean_val(row.get('生日', '')),
            '護照號碼'     : clean_val(row.get('護照號碼', '')),
            '護照到期日'   : clean_val(row.get('護照到期日', '')),
            '飲食'         : clean_diet(row.get('飲食備註\n(定位備註)', '')),
            '房號'         : clean_val(row.get('標準分房(房號)', '')),
            '序號'         : clean_val(row.get('序號', '')),
            '_房務備註'    : clean_val(row.get('分房備住', '')),
        })

    # ── 7. 排序 ──
    records.sort(key=lambda x: (x['_sort_date'], x['_item_order'], x['訂編']))

    # ── 8. 按日期分組 ──
    date_map = defaultdict(list)
    for r in records:
        sd = r['_sort_date']
        if DATE_START <= sd <= DATE_END:
            date_map[sd].append(r)

    # ── 9. 產生含分隔列的完整列表（供寫入 Google Sheet 用）──
    def count_by_area(recs):
        counts = {'野澤': 0, '斑尾': 0, '湯澤': 0, '龍平': 0}
        for r in recs:
            itm = r.get('項目', '')
            if '野澤' in itm:   counts['野澤'] += 1
            elif '斑尾' in itm: counts['斑尾'] += 1
            elif '湯澤' in itm: counts['湯澤'] += 1
            elif '龍平' in itm: counts['龍平'] += 1
        return counts

    from shared.config import COURSE_COLS  # noqa

    all_sheet_rows  = []
    all_records_flat = []

    cur = DATE_START
    while cur <= DATE_END:
        day_records = date_map.get(cur, [])
        wd          = get_weekday(cur)
        date_str    = cur.strftime('%y/%m/%d')
        counts      = count_by_area(day_records)

        sep_row    = [''] * len(COURSE_COLS)
        sep_row[1] = date_str
        sep_row[2] = wd
        sep_row[3] = f"野:{counts['野澤']}"
        sep_row[4] = f"斑:{counts['斑尾']}"
        sep_row[5] = f"湯:{counts['湯澤']}"
        sep_row[6] = f"龍:{counts['龍平']}"
        all_sheet_rows.append({'_is_separator': True, '_data': sep_row})

        for local_idx, r in enumerate(day_records, 1):
            r['編號']     = str(local_idx)
            r['出發日期'] = date_str
            all_records_flat.append(r)
            all_sheet_rows.append({'_is_separator': False, '_data': r})

        if not day_records:
            all_sheet_rows.append({
                '_is_separator': False,
                '_data'        : {col: '' for col in COURSE_COLS},
            })

        cur += timedelta(days=1)

    print(f"整理完成：{len(all_records_flat)} 筆資料，{len(date_map)} 個出發日期")

    # ── 10. 序列化輸出 ──
    output = {
        'all_records_flat' : all_records_flat,
        'all_sheet_rows'   : all_sheet_rows,
        'date_map'         : dict(date_map),
        'DATE_START'       : DATE_START,
        'DATE_END'         : DATE_END,
        'headers'          : headers,           # 科威原始欄位名稱（供 REPORTER 用）
        'df_records'       : df.to_dict('records'),  # 科威原始資料（供 REPORTER 用）
        'generated_at'     : datetime.now().isoformat(),
    }

    with open(OUTPUT_PATH, 'wb') as f:
        pickle.dump(output, f)
    print(f"OK cleaner_output.pkl 已儲存")
    print(f"{'=' * 50}\n")

    return output


def load():
    """其他工作員呼叫此函數讀取 CLEANER 的輸出"""
    if not os.path.exists(OUTPUT_PATH):
        raise FileNotFoundError(
            "找不到 cleaner_output.pkl，請先執行 CLEANER 清洗員"
        )
    with open(OUTPUT_PATH, 'rb') as f:
        return pickle.load(f)


# ─────────────────────────────────────────────
if __name__ == '__main__':
    run()
