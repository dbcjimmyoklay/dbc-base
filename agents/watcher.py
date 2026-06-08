"""
agents/watcher.py
WATCHER 比對員
─────────────────────────────────────────────
職責：
  Part A │ 比對課程安排
         │  • 讀取 Google Sheet「課程安排26/27」現有資料
         │  • 與 CLEANER 最新資料比對
         │  • 偵測：新增 / 欄位異動 / 刪減
         │  • 寫入「異動紀錄」分頁（先清除已勾選TRUE的舊記錄）

  Part B │ 野澤訂房比對報告
         │  • 讀取「野澤當地訂房」試算表 26-27 分頁
         │  • 比對最新野澤旅客資料
         │  • 偵測：需要訂房的新訂單、同訂單新增人員、
         │          姓名異動、刪減旅客、出發日期/泊數異動
         │  • 在「野澤當地訂房」試算表寫入/更新「比對報告」分頁
─────────────────────────────────────────────
注意：WATCHER 只讀、只報告，不修改訂房表本體
─────────────────────────────────────────────
"""

import os
import re
import sys
from datetime import datetime

# ── import 路徑兼容（直接執行 or 從 main.py 執行）──
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shared.config import (
    SHEET_COURSE, SHEET_CHANGE, CHANGE_HEADER,
    AUTO_COLS, IMPORTANT_COLS, MANUAL_COLS,
    COURSE_COLS, COURSE_HEADERS,
    NOZAWA_BOOKING_SPREADSHEET_ID, NOZAWA_BOOKING_SHEET, NOZAWA_COMPARE_SHEET,
)
from shared.utils import normalize_level, build_record_key
from shared.gsheet import (
    get_spreadsheet, get_or_create_sheet,
    clear_rows_from, rebuild_filter, batch_update,
    req_cell_color, req_row_color, req_borders,
    req_col_width, req_checkbox, req_header_format,
)
import gspread
from google.oauth2.service_account import Credentials
from shared.config import CREDS_PATH, SCOPES

# ── 顏色常數 ──
C_NEW    = {'red': 1.0,  'green': 0.95, 'blue': 0.6}   # 新增　亮黃
C_DEL    = {'red': 1.0,  'green': 0.7,  'blue': 0.7}   # 刪減　紅
C_CHANGE = {'red': 1.0,  'green': 1.0,  'blue': 1.0}   # 異動　白
C_NEED   = {'red': 1.0,  'green': 0.98, 'blue': 0.7}   # 需訂房　淡黃
C_ADD    = {'red': 0.95, 'green': 1.0,  'blue': 0.85}  # 新增人員　淡綠
C_NAME   = {'red': 0.8,  'green': 0.9,  'blue': 1.0}   # 姓名異動　淡藍
C_DATE   = {'red': 1.0,  'green': 0.88, 'blue': 0.7}   # 日期/泊數異動　淡橘
C_HDR    = {'red': 0.15, 'green': 0.15, 'blue': 0.15}  # 區塊標題　深灰
C_WHITE  = {'red': 1.0,  'green': 1.0,  'blue': 1.0}


# ═══════════════════════════════════════════
# 工具函數
# ═══════════════════════════════════════════

def _val(s):
    """統一空值處理，用於比對"""
    if s is None:
        return ''
    return str(s).strip()


def _norm_date(s):
    """統一日期格式（去掉 20 前綴），例：2026/12/18 → 26/12/18"""
    s = _val(s)
    if s.startswith('20') and len(s) == 10:
        return s[2:]
    return s


def _norm_lv(s):
    """統一 LEVEL 格式"""
    return normalize_level(_val(s))


def _cells_equal(col, old_val, new_val):
    """
    智慧比對：判斷兩個值是否「實質相等」
    處理 LEVEL 格式、日期格式、空值等差異
    """
    o = _val(old_val)
    n = _val(new_val)

    # LEVEL 欄位統一格式後再比
    if col in ('LEVEL', 'Lv'):
        return _norm_lv(o) == _norm_lv(n)

    # 出發日期統一格式後再比
    if col == '出發日期':
        return _norm_date(o) == _norm_date(n)

    return o == n


# ═══════════════════════════════════════════
# Part A：課程安排比對
# ═══════════════════════════════════════════

def _build_existing_map(ws_course):
    """
    讀取現有課程安排，建立 key → fields 的索引
    結構：第1列=篩選按鈕，第2列=標題，第3列起=資料
    """
    existing_data = ws_course.get_all_values()
    existing_map  = {}

    if len(existing_data) <= 1:
        return existing_map

    header_row = existing_data[1]                          # 第2列是標題
    col_map    = {h: idx for idx, h in enumerate(header_row)}

    idx_visitor = col_map.get('旅客編號', -1)
    idx_ding    = col_map.get('訂編', -1)
    idx_name    = col_map.get('中文姓名', -1)
    idx_en      = col_map.get('英文姓名', -1)

    for row in existing_data[2:]:                          # 第3列起是資料
        if not row or len(row) < 2:
            continue

        def _get(idx):
            return row[idx].strip() if 0 <= idx < len(row) else ''

        visitor_id = _get(idx_visitor)
        ding       = _get(idx_ding)
        name       = _get(idx_name)
        en         = _get(idx_en)

        if visitor_id:
            key = f"v_{visitor_id}"
        elif name:
            key = f"{ding}_{name}"
        elif en:
            key = f"{ding}_en_{en}"
        else:
            continue

        fields = {h: row[i] if i < len(row) else ''
                  for h, i in col_map.items()}
        existing_map[key] = fields

    return existing_map


def _compare_course(all_records_flat, existing_map):
    """
    比對新舊課程安排，回傳 change_logs
    每筆 log：[異動日期, 異動類型, 項目, 出發日期, 訂編, 中文姓名, 異動欄位, 原值, 新值, 處理]
    """
    today_str    = datetime.now().strftime('%m/%d')
    change_logs  = []
    new_keys     = set()
    seq_map      = {}

    for r in all_records_flat:
        ding = r.get('訂編', '')
        seq_map[ding] = seq_map.get(ding, 0) + 1
        r['_seq'] = seq_map[ding]
        key = build_record_key(r, seq=r['_seq'])
        new_keys.add(key)

        item   = r.get('項目', '')
        depart = r.get('出發日期', '')
        ding_v = r.get('訂編', '')
        name_v = r.get('中文姓名', '')

        if key in existing_map:
            # ── 已存在：檢查 AUTO_COLS 中的 IMPORTANT_COLS ──
            old_fields = existing_map[key]
            for col in AUTO_COLS:
                if col not in IMPORTANT_COLS:
                    continue
                # Google Sheet 標題顯示的是 'Lv' 不是 'LEVEL'
                sheet_col = 'Lv' if col == 'LEVEL' else col
                old_val   = old_fields.get(sheet_col, '')
                new_val   = str(r.get(col, ''))
                if not _cells_equal(col, old_val, new_val):
                    change_logs.append([
                        today_str, '異動', item, depart,
                        ding_v, name_v,
                        col, _val(old_val), _val(new_val), '',
                    ])
        else:
            # ── 新增 ──
            change_logs.append([
                today_str, '新增', item, depart,
                ding_v, name_v, '', '', '', '',
            ])

    # ── 刪減：舊資料有、新資料沒有 ──
    for key in existing_map:
        if key not in new_keys:
            old = existing_map[key]
            change_logs.append([
                today_str, '刪減',
                old.get('項目', ''), old.get('出發日期', ''),
                old.get('訂編', ''), old.get('中文姓名', ''),
                '', '', '', '',
            ])

    return change_logs


def _write_change_log(sh, ws_change, change_logs):
    """
    寫入異動紀錄：
    1. 先清除「處理」欄已勾選（TRUE）的舊記錄
    2. 追加新異動
    3. 套用填色、格線、Checkbox
    """
    print("  清除已處理的異動紀錄...")
    existing_change = ws_change.get_all_values()
    n_cols          = len(CHANGE_HEADER)

    # 找「處理」欄的位置
    idx_done = -1
    if len(existing_change) > 1:
        try:
            idx_done = existing_change[1].index('處理')
        except ValueError:
            pass

    # 清除已勾選的記錄
    if idx_done >= 0 and len(existing_change) > 2:
        kept    = []
        removed = 0
        for row in existing_change[2:]:
            done_val = row[idx_done].strip().upper() if idx_done < len(row) else ''
            if done_val in ('TRUE', '1'):
                removed += 1
            else:
                kept.append(row)
        if removed > 0:
            clear_rows_from(sh, ws_change, start_row_index=2, n_cols=n_cols)
            if kept:
                ws_change.update(kept, 'A3', value_input_option='RAW')
            existing_change = existing_change[:2] + kept
            print(f"  已清除 {removed} 筆已處理記錄")

    if not change_logs:
        print("  無異動記錄")
        return

    # 確保第2列有標題
    if len(existing_change) < 2 or not existing_change[1]:
        ws_change.update([CHANGE_HEADER], 'A2', value_input_option='RAW')
        existing_change = (existing_change[:1] if existing_change else []) + [CHANGE_HEADER]

    # 追加位置
    next_row   = len(existing_change) + 1
    start_row  = next_row
    ws_change.update(change_logs, f'A{next_row}', value_input_option='RAW')
    print(f"  OK 寫入 {len(change_logs)} 筆異動記錄")

    # ── 格式 ──
    sid     = ws_change.id
    reqs    = []
    idx_type = CHANGE_HEADER.index('異動類型')

    # 欄寬
    from shared.config import CHANGE_COL_WIDTHS
    for ci, w in enumerate(CHANGE_COL_WIDTHS):
        reqs.append(req_col_width(sid, ci, w))

    # 每列填色
    for i, log in enumerate(change_logs):
        ridx = start_row + i          # 1-based
        t    = log[idx_type]
        color = C_NEW if t == '新增' else C_DEL if t == '刪減' else C_CHANGE
        reqs.append(req_row_color(sid, ridx, n_cols, color))

    # 格線
    total_rows = start_row - 1 + len(change_logs)
    reqs.append(req_borders(sid, 3, total_rows, 0, n_cols))

    # Checkbox（處理欄），從第3列到全部資料列再多200（預留空間）
    reqs.append(req_checkbox(sid, 2, total_rows + 200, idx_done))

    batch_update(sh, reqs)
    print("  OK 異動紀錄格式設定完成")

    # 重建篩選器
    rebuild_filter(sh, ws_change, total_rows - 2, n_cols)


# ═══════════════════════════════════════════
# Part B：野澤訂房比對
# ═══════════════════════════════════════════

def _read_booking_table(sh_booking):
    """
    讀取野澤訂房表 26-27 分頁
    支援「併房」：同一儲存格可有多個訂編（空格/換行/逗號分隔）

    回傳：
      existing_dings → set  訂房表所有看過的訂編（含併房展開）
      order_info     → dict ding → {出發日期, 泊數, 是併房, rows:[...]}
      placeholders   → dict ding → list[(佔位中文姓名, 佔位英文姓名)]
    """
    ws = sh_booking.worksheet(NOZAWA_BOOKING_SHEET)
    all_vals = ws.get_all_values()

    existing_dings = set()
    order_info     = {}
    placeholders   = {}

    if not all_vals:
        return existing_dings, order_info, placeholders

    col_map = {h.strip(): i for i, h in enumerate(all_vals[0]) if h.strip()}
    idx_ding  = col_map.get('訂編',   -1)
    idx_cn    = col_map.get('中文姓名', -1)
    idx_en    = col_map.get('英文姓名', -1)
    idx_date  = col_map.get('出發日期', -1)
    idx_nights= col_map.get('泊數',   -1)
    idx_adult = col_map.get('大人',   -1)
    idx_child = col_map.get('小人',   -1)
    idx_baby  = col_map.get('嬰兒',   -1)

    # 訂房表每張訂單只在第一列填 訂編+姓名，後面的空白列無法明確判斷屬於哪張訂單
    # → 只計算「有訂編 + 有姓名」的列；同訂編多人由我們之後展開
    for row in all_vals[1:]:
        def _g(idx):
            return row[idx].strip() if 0 <= idx < len(row) else ''

        ding_cell = _g(idx_ding)
        dings_in_row = re.findall(r'\d{4,5}', ding_cell)
        if not dings_in_row:
            continue

        cn      = _g(idx_cn)
        en      = _g(idx_en)
        dep     = _g(idx_date)
        nights  = _g(idx_nights)

        # 從「大人 / 小人 / 嬰兒」三欄相加得到該訂單實際人數
        def _to_int(v):
            try: return int(re.sub(r'\D', '', str(v))) if v else 0
            except: return 0
        n_adult = _to_int(_g(idx_adult)) if idx_adult >= 0 else 0
        n_child = _to_int(_g(idx_child)) if idx_child >= 0 else 0
        n_baby  = _to_int(_g(idx_baby))  if idx_baby  >= 0 else 0
        head_count = n_adult + n_child + n_baby

        is_ph       = bool(re.match(r'^旅客\w+$', cn))
        is_combined = len(dings_in_row) > 1

        for ding in dings_in_row:
            existing_dings.add(ding)
            if ding not in order_info:
                order_info[ding] = {
                    '出發日期': dep,
                    '泊數'   : nights,
                    '人數'   : head_count,   # 大人+小人+嬰兒
                    '是併房' : is_combined,
                    'rows'   : [],
                }
            else:
                if is_combined:
                    order_info[ding]['是併房'] = True
                # 若同訂編有多列各自填了人數，取最大
                if head_count > order_info[ding].get('人數', 0):
                    order_info[ding]['人數'] = head_count
            order_info[ding]['rows'].append({
                '中文姓名': cn,
                '英文姓名': en,
                '是佔位'  : is_ph,
            })
            if is_ph:
                placeholders.setdefault(ding, []).append((cn, en))

    return existing_dings, order_info, placeholders


def _compare_nozawa(all_records_flat, existing_dings, order_info, placeholders):
    """
    新版比對：產出 3 大區塊
      new_orders     → 新增訂單（按訂編去重，每訂單一列）
      removed_orders → 刪減訂單（訂房表有但新資料沒有）
      changes        → 異動（出發日期/泊數/人數）
      missing_names  → 補資料（訂房表佔位姓名 → 新資料真名）
    """
    # 最新野澤旅客（不含純課）
    nozawa = [
        r for r in all_records_flat
        if '野澤' in r.get('項目', '') and '純課' not in r.get('項目', '')
    ]

    # 新資料按訂編分組
    new_by_ding = {}
    for r in nozawa:
        d = str(r.get('訂編', '') or '').strip()
        if not d:
            continue
        new_by_ding.setdefault(d, []).append(r)

    new_dings = set(new_by_ding.keys())

    # ── 區塊 1a：新增訂單 ──
    new_orders = []
    for ding in sorted(new_dings - existing_dings,
                       key=lambda d: (_norm_date(new_by_ding[d][0].get('出發日期','')), d)):
        people = new_by_ding[ding]
        first  = people[0]
        new_orders.append({
            '訂編'   : ding,
            '出發日期': _norm_date(first.get('出發日期', '')),
            '泊數'   : str(first.get('泊數', '') or '').strip(),
            '天數'   : str(first.get('天數', '') or '').strip(),
            '中文姓名': first.get('中文姓名', ''),
            '英文姓名': first.get('英文姓名', ''),
            '人數'   : len(people),
            '備註'   : first.get('備註', '') or '',
        })

    # ── 區塊 1b：刪減訂單 ──
    removed_orders = []
    for ding in sorted(existing_dings - new_dings,
                       key=lambda d: (order_info[d].get('出發日期', ''), d)):
        info = order_info[ding]
        first_pax = info['rows'][0] if info['rows'] else {}
        removed_orders.append({
            '訂編'   : ding,
            '出發日期': info.get('出發日期', ''),
            '中文姓名': first_pax.get('中文姓名', ''),
            '英文姓名': first_pax.get('英文姓名', ''),
        })

    # ── 區塊 2：異動 ──
    changes = []
    for ding in sorted(new_dings & existing_dings,
                       key=lambda d: (order_info[d].get('出發日期', ''), d)):
        info = order_info[ding]
        people = new_by_ding[ding]
        first  = people[0]
        first_name = first.get('中文姓名', '') or ''

        old_dep = info.get('出發日期', '')
        new_dep = _norm_date(first.get('出發日期', ''))
        if old_dep and new_dep and old_dep != new_dep:
            changes.append({
                '訂編'   : ding,
                '中文姓名': first_name,
                '出發日期': new_dep,
                '類型'   : '出發日期',
                '原值'   : old_dep,
                '新值'   : new_dep,
            })

        old_nights = str(info.get('泊數', '') or '').strip()
        new_nights = str(first.get('泊數', '') or '').strip()
        if old_nights and new_nights and old_nights != new_nights:
            changes.append({
                '訂編'   : ding,
                '中文姓名': first_name,
                '出發日期': new_dep,
                '類型'   : '泊數',
                '原值'   : old_nights,
                '新值'   : new_nights,
            })

        # 人數變動：對比訂房表「大人+小人+嬰兒」欄位 vs 新資料人數
        # 併房訂單不比對（一格多訂單時無法明確分配人數）
        if not info.get('是併房'):
            old_count = info.get('人數', 0)
            new_count = len(people)
            if old_count > 0 and old_count != new_count:
                changes.append({
                    '訂編'   : ding,
                    '中文姓名': first_name,
                    '出發日期': new_dep,
                    '類型'   : '人數',
                    '原值'   : f'{old_count}人',
                    '新值'   : f'{new_count}人',
                })

    # ── 區塊 3：補資料 ──
    missing_names = []
    PLACEHOLDER_RE = re.compile(r'^旅客\w+$')
    for ding, ph_list in placeholders.items():
        if ding not in new_by_ding:
            continue
        # 從新資料抽出「不是佔位」的真名
        real = [p for p in new_by_ding[ding]
                if not PLACEHOLDER_RE.match(str(p.get('中文姓名', '') or '').strip())]
        if not real:
            continue
        # 整理真名 + 英文（佔位數 vs 真名數 也許不一致，全列出來）
        cn_names = '／'.join([p.get('中文姓名', '') for p in real if p.get('中文姓名')])
        en_names = '／'.join([p.get('英文姓名', '') for p in real if p.get('英文姓名')])
        missing_names.append({
            '訂編'    : ding,
            '出發日期': order_info[ding].get('出發日期', ''),
            '中文姓名': cn_names,
            '英文姓名': en_names,
            '佔位數'  : len(ph_list),
            '佔位姓名': '／'.join([n[0] for n in ph_list]),
        })
    missing_names.sort(key=lambda x: (x['出發日期'], x['訂編']))

    return new_orders, removed_orders, changes, missing_names


def _write_compare_report(sh_booking, new_orders, removed_orders, changes, missing_names):
    """
    在野澤訂房表試算表中寫入「比對報告」分頁
    新版三大區塊：
      ① 新增 / 刪減訂單
      ② 異動（出發日期 / 泊數 / 人數）
      ③ 補資料（佔位姓名 → 真實姓名）
    欄位：出發日期 / 泊數 / 天數 / 訂編 / 中文姓名 / 英文姓名 / 備註
    """
    print("  建立/更新 野澤訂房比對報告...")

    try:
        ws = sh_booking.worksheet(NOZAWA_COMPARE_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh_booking.add_worksheet(title=NOZAWA_COMPARE_SHEET, rows=500, cols=12)

    now_str = datetime.now().strftime('%Y/%m/%d %H:%M')
    sid     = ws.id

    # ── 統一欄位：出發日期 / 泊數 / 天數 / 訂編 / 中文姓名 / 英文姓名 / 備註 ──
    N_COLS  = 7
    HEADERS = ['出發日期', '泊數', '天數', '訂編', '中文姓名', '英文姓名', '備註']

    rows   = []
    colors = []

    def _add_title(text):
        rows.append([text] + [''] * (N_COLS - 1))
        colors.append('title')

    def _add_section(text, color_key='section'):
        rows.append([text] + [''] * (N_COLS - 1))
        colors.append(color_key)

    def _add_header():
        rows.append(HEADERS)
        colors.append('header')

    def _add_data(cols_list, color):
        rows.append(cols_list + [''] * max(0, N_COLS - len(cols_list)))
        colors.append(color)

    def _add_empty():
        rows.append([''] * N_COLS)
        colors.append(None)

    def _add_none():
        rows.append(['（無）'] + [''] * (N_COLS - 1))
        colors.append(None)

    # 依出發日期交替底色
    def _date_color_key(prev_date, cur_date, prev_color):
        if cur_date == prev_date:
            return prev_color
        return 'date_b' if prev_color == 'date_a' else 'date_a'

    # 更新時間
    rows.append([f'更新時間：{now_str}'] + [''] * (N_COLS - 1))
    colors.append(None)
    _add_empty()

    # ═══════════════════════════════════════════
    # 區塊 ①：新增 / 刪減訂單
    # ═══════════════════════════════════════════
    _add_title('═══ ① 新增 / 刪減 訂單 ═══')

    _add_section('▶ 新增訂單（需安排訂房）', 'section_add')
    if new_orders:
        _add_header()
        prev_date = None
        cur_color = 'date_b'   # 第一筆從 date_a 開始
        for x in new_orders:
            cur_color = _date_color_key(prev_date, x['出發日期'], cur_color)
            prev_date = x['出發日期']
            note = f"共{x['人數']}人" if x['人數'] > 1 else ''
            if x.get('備註'):
                note = (note + '  ' + x['備註']).strip()
            _add_data(
                [x['出發日期'], x['泊數'], x['天數'],
                 x['訂編'], x['中文姓名'], x['英文姓名'], note],
                cur_color,
            )
    else:
        _add_none()
    _add_empty()

    _add_section('▶ 刪減訂單（請確認是否取消訂房）', 'section_del')
    if removed_orders:
        _add_header()
        prev_date = None
        cur_color = 'date_b'
        for x in removed_orders:
            cur_color = _date_color_key(prev_date, x['出發日期'], cur_color)
            prev_date = x['出發日期']
            _add_data(
                [x['出發日期'], '', '',
                 x['訂編'], x['中文姓名'], x['英文姓名'], '請確認是否取消'],
                'del',
            )
    else:
        _add_none()
    _add_empty()

    # ═══════════════════════════════════════════
    # 區塊 ②：異動
    # ═══════════════════════════════════════════
    _add_title('═══ ② 異動（出發日期 / 泊數 / 人數）═══')
    if changes:
        _add_header()
        prev_date = None
        cur_color = 'date_b'
        for x in changes:
            dep = x.get('出發日期', '')
            cur_color = _date_color_key(prev_date, dep, cur_color)
            prev_date = dep
            note = f"{x['類型']}: {x['原值']} → {x['新值']}"
            _add_data(
                [dep, '', '', x['訂編'], x['中文姓名'], '', note],
                cur_color,
            )
    else:
        _add_none()
    _add_empty()

    # ═══════════════════════════════════════════
    # 區塊 ③：補資料
    # ═══════════════════════════════════════════
    _add_title('═══ ③ 補資料（訂房表佔位姓名 → 真實姓名）═══')
    if missing_names:
        _add_header()
        prev_date = None
        cur_color = 'date_b'
        for x in missing_names:
            dep = x.get('出發日期', '')
            cur_color = _date_color_key(prev_date, dep, cur_color)
            prev_date = dep
            note = f"原佔位：{x['佔位姓名']}（{x['佔位數']}位）"
            _add_data(
                [dep, '', '', x['訂編'], x['中文姓名'], x['英文姓名'], note],
                cur_color,
            )
    else:
        _add_none()

    # ── 清空舊內容，寫入新資料 ──
    sh_booking.batch_update({'requests': [{
        'updateCells': {
            'range' : {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 500,
                       'startColumnIndex': 0, 'endColumnIndex': N_COLS},
            'fields': 'userEnteredValue,userEnteredFormat',
        }
    }]})
    ws.update(rows, 'A1', value_input_option='RAW')

    # ── 格式設定 ──
    COLOR_MAP = {
        'title'      : C_HDR,
        'header'     : {'red': 0.3,  'green': 0.3,  'blue': 0.3},
        'section_add': {'red': 0.18, 'green': 0.49, 'blue': 0.20},   # 深綠（新增區段）
        'section_del': {'red': 0.70, 'green': 0.18, 'blue': 0.18},   # 深紅（刪減區段）
        'section'    : {'red': 0.20, 'green': 0.30, 'blue': 0.50},   # 深藍灰
        'date_a'     : {'red': 0.94, 'green': 0.97, 'blue': 1.00},   # 淡藍
        'date_b'     : {'red': 1.00, 'green': 0.97, 'blue': 0.88},   # 淡黃
        'del'        : {'red': 1.00, 'green': 0.90, 'blue': 0.90},   # 淡紅
    }
    TEXT_WHITE = {'red': 1.0, 'green': 1.0, 'blue': 1.0}
    BOLD_KEYS  = {'title', 'header', 'section_add', 'section_del', 'section'}

    fmt_reqs = []
    for i, (row_data, color_key) in enumerate(zip(rows, colors), 1):
        if color_key is None:
            continue
        color = COLOR_MAP.get(color_key, C_WHITE)
        is_dark = color_key in BOLD_KEYS
        fmt_reqs.append({
            'repeatCell': {
                'range': {'sheetId': sid, 'startRowIndex': i - 1, 'endRowIndex': i,
                          'startColumnIndex': 0, 'endColumnIndex': N_COLS},
                'cell' : {'userEnteredFormat': {
                    'backgroundColor': color,
                    'textFormat'     : {
                        'bold': is_dark,
                        'foregroundColor': TEXT_WHITE if is_dark else {'red': 0.1, 'green': 0.1, 'blue': 0.1},
                    },
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat)',
            }
        })

    # 欄寬：出發日期 / 泊數 / 天數 / 訂編 / 中文姓名 / 英文姓名 / 備註
    for ci, px in enumerate([72, 45, 45, 65, 80, 130, 160]):
        fmt_reqs.append(req_col_width(sid, ci, px))

    # 格線（全部資料範圍）
    fmt_reqs.append(req_borders(sid, 1, len(rows), 0, N_COLS))

    batch_update(sh_booking, fmt_reqs)

    total_sections = sum([
        len(new_orders), len(removed_orders),
        len(changes), len(missing_names),
    ])
    print(f"  OK 比對報告已更新（{total_sections} 項待確認）")


# ═══════════════════════════════════════════
# 主函數
# ═══════════════════════════════════════════

def run(cleaner_output=None):
    print("=" * 50)
    print("WATCHER 比對員 啟動")
    print("=" * 50)

    # ── 讀取 CLEANER 輸出 ──
    if cleaner_output is None:
        from agents.cleaner import load
        cleaner_output = load()

    all_records_flat = cleaner_output['all_records_flat']
    print(f"資料筆數：{len(all_records_flat)}")

    # ── 連線 Google Sheets ──
    print("連線 Google Sheets（DBC Base）...")
    sh_main = get_spreadsheet()

    ws_course = get_or_create_sheet(sh_main, SHEET_COURSE)
    ws_change = get_or_create_sheet(sh_main, SHEET_CHANGE, cols=12)

    # 確保異動紀錄有標題
    existing_change = ws_change.get_all_values()
    if len(existing_change) < 2 or not any(existing_change[1] if len(existing_change) > 1 else []):
        ws_change.update([CHANGE_HEADER], 'A2', value_input_option='RAW')

    # ── Part A：課程安排比對 ──
    print("\n[Part A] 課程安排比對...")
    existing_map = _build_existing_map(ws_course)
    print(f"  現有課程安排：{len(existing_map)} 筆")

    # 保護機制：有資料但讀取異常時停止
    if len(existing_change) > 10 and len(existing_map) == 0:
        print("警告：Google Sheet 有資料但讀取失敗，停止執行！")
        return

    change_logs = _compare_course(all_records_flat, existing_map)
    print(f"  偵測到 {len(change_logs)} 筆異動")
    _write_change_log(sh_main, ws_change, change_logs)

    # ── Part B：野澤訂房比對 ──
    print("\n[Part B] 野澤訂房比對...")
    creds      = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    gc         = gspread.authorize(creds)
    sh_booking = gc.open_by_key(NOZAWA_BOOKING_SPREADSHEET_ID)

    existing_dings, order_info, placeholders = _read_booking_table(sh_booking)
    print(f"  訂房表現有訂單：{len(existing_dings)} 個"
          f"（含併房展開）")

    new_orders, removed_orders, changes, missing_names = \
        _compare_nozawa(all_records_flat, existing_dings, order_info, placeholders)

    print(f"  新增訂單：{len(new_orders)} 個")
    print(f"  刪減訂單：{len(removed_orders)} 個")
    print(f"  異動    ：{len(changes)} 筆")
    print(f"  補資料  ：{len(missing_names)} 筆")

    try:
        _write_compare_report(
            sh_booking,
            new_orders, removed_orders, changes, missing_names,
        )
    except gspread.exceptions.APIError as e:
        msg = str(e)
        if '403' in msg or 'permission' in msg.lower():
            print(f"\n  ⚠️ 比對報告寫入失敗：服務帳號無權限編輯野澤訂房表")
            print(f"     請至野澤訂房表 → 右上「共用」→ 加入下列帳號為「編輯者」：")
            print(f"     dbc-sheet@gen-lang-client-0187113768.iam.gserviceaccount.com")
            print(f"     試算表 URL：https://docs.google.com/spreadsheets/d/{NOZAWA_BOOKING_SPREADSHEET_ID}/")
            print(f"     （其餘 WATCHER 結果已正常處理）")
        else:
            raise

    print(f"\n{'=' * 50}")
    print(f"WATCHER 完成！{datetime.now().strftime('%Y/%m/%d %H:%M')}")
    print(f"{'=' * 50}\n")


# ─────────────────────────────────────────────
if __name__ == '__main__':
    run()
