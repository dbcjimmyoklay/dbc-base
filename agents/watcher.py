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
    回傳：
      booking_map   → dict  key(訂編_中文姓名) → {欄位: 值}
      order_map     → dict  訂編 → list[旅客dict]  （同訂單所有人）
      col_map       → dict  欄位名稱 → 索引
    """
    ws = sh_booking.worksheet(NOZAWA_BOOKING_SHEET)
    all_vals = ws.get_all_values()

    booking_map = {}
    order_map   = {}
    col_map     = {}

    if not all_vals:
        return booking_map, order_map, col_map

    # 第1列是標題
    col_map = {h.strip(): i for i, h in enumerate(all_vals[0]) if h.strip()}

    idx_ding  = col_map.get('訂編',   -1)
    idx_cn    = col_map.get('中文姓名', -1)
    idx_en    = col_map.get('英文姓名', -1)
    idx_date  = col_map.get('出發日期', -1)
    idx_nights= col_map.get('泊數',   -1)

    for row in all_vals[1:]:
        def _g(idx):
            return row[idx].strip() if 0 <= idx < len(row) else ''

        ding = _g(idx_ding)
        if not ding or not ding.isdigit():
            continue                   # 略過空列或日期分隔列

        cn      = _g(idx_cn)
        en      = _g(idx_en)
        dep     = _g(idx_date)
        nights  = _g(idx_nights)
        key     = f"{ding}_{cn}" if cn else f"{ding}_en_{en}"

        record = {
            '訂編'   : ding,
            '中文姓名': cn,
            '英文姓名': en,
            '出發日期': dep,
            '泊數'   : nights,
            '_key'   : key,
        }
        booking_map[key] = record

        if ding not in order_map:
            order_map[ding] = []
        order_map[ding].append(record)

    return booking_map, order_map, col_map


def _compare_nozawa(all_records_flat, booking_map, order_map):
    """
    比對最新野澤旅客 vs 訂房表現有資料
    回傳五類清單：
      need_booking  → 全新訂單（訂編完全不在訂房表）
      need_add_pax  → 同訂單有新旅客（訂編存在但有人不在訂房表）
      name_changes  → 姓名異動
      deleted_pax   → 訂房表有、最新資料沒有
      date_changes  → 出發日期或泊數異動
    """
    # 最新野澤旅客（只取野澤，不含純課）
    nozawa_records = [
        r for r in all_records_flat
        if '野澤' in r.get('項目', '') and '純課' not in r.get('項目', '')
    ]

    # 新資料的訂編集合
    new_ding_set  = {r['訂編'] for r in nozawa_records}
    new_key_set   = set()

    need_booking  = []   # 全新訂單
    need_add_pax  = []   # 同訂單新增人員
    name_changes  = []   # 姓名異動
    date_changes  = []   # 日期/泊數異動

    for r in nozawa_records:
        ding    = r.get('訂編', '')
        cn      = r.get('中文姓名', '')
        en      = r.get('英文姓名', '')
        dep_new = _norm_date(r.get('出發日期', ''))
        nights_new = _val(r.get('泊數', ''))
        key     = f"{ding}_{cn}" if cn else f"{ding}_en_{en}"
        new_key_set.add(key)

        if ding not in order_map:
            # ── ① 全新訂單（訂編不存在）──
            # 同訂單第一次出現才加，避免重複
            if not any(x['訂編'] == ding for x in need_booking):
                need_booking.append({
                    '訂編'    : ding,
                    '出發日期': dep_new,
                    '泊數'    : nights_new,
                    '中文姓名': cn,
                    '英文姓名': en,
                    '天數'    : r.get('天數', ''),
                    '備註'    : r.get('備註', ''),
                    '_is_first': True,
                })
            else:
                # 同訂單後續人員也列出
                need_booking.append({
                    '訂編'    : ding,
                    '出發日期': dep_new,
                    '泊數'    : nights_new,
                    '中文姓名': cn,
                    '英文姓名': en,
                    '天數'    : r.get('天數', ''),
                    '備註'    : r.get('備註', ''),
                    '_is_first': False,
                })
        else:
            # 訂編存在，逐人比對
            booking_keys_for_order = {b['_key'] for b in order_map[ding]}
            if key not in booking_keys_for_order:
                # ── ② 同訂單新增人員 ──
                need_add_pax.append({
                    '訂編'    : ding,
                    '出發日期': dep_new,
                    '泊數'    : nights_new,
                    '中文姓名': cn,
                    '英文姓名': en,
                })
            else:
                # 已存在：檢查姓名和日期
                old = booking_map.get(key, {})
                cn_old  = _val(old.get('中文姓名', ''))
                en_old  = _val(old.get('英文姓名', ''))
                dep_old = _norm_date(old.get('出發日期', ''))
                nts_old = _val(old.get('泊數', ''))

                # ── ③ 姓名異動 ──
                if cn_old != _val(cn) or en_old != _val(en):
                    name_changes.append({
                        '訂編'    : ding,
                        '原中文'  : cn_old,
                        '新中文'  : _val(cn),
                        '原英文'  : en_old,
                        '新英文'  : _val(en),
                    })

                # ── ⑤ 出發日期 / 泊數異動 ──
                if dep_old and dep_old != dep_new:
                    date_changes.append({
                        '訂編'    : ding,
                        '中文姓名': cn,
                        '異動欄位': '出發日期',
                        '原值'   : dep_old,
                        '新值'   : dep_new,
                    })
                if nts_old and nts_old != nights_new:
                    date_changes.append({
                        '訂編'    : ding,
                        '中文姓名': cn,
                        '異動欄位': '泊數',
                        '原值'   : nts_old,
                        '新值'   : nights_new,
                    })

    # ── ④ 刪減：訂房表有，最新資料沒有 ──
    deleted_pax = []
    for key, old in booking_map.items():
        if old['訂編'] not in new_ding_set:
            deleted_pax.append(old)
        elif key not in new_key_set:
            deleted_pax.append(old)

    return need_booking, need_add_pax, name_changes, deleted_pax, date_changes


def _write_compare_report(sh_booking, need_booking, need_add_pax,
                           name_changes, deleted_pax, date_changes):
    """
    在野澤訂房表試算表中寫入「比對報告」分頁
    格式：區塊式，各類異動分開，有顏色標示
    """
    print("  建立/更新 野澤訂房比對報告...")

    # 取得或建立「比對報告」分頁
    try:
        ws = sh_booking.worksheet(NOZAWA_COMPARE_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh_booking.add_worksheet(title=NOZAWA_COMPARE_SHEET, rows=500, cols=12)

    now_str = datetime.now().strftime('%Y/%m/%d %H:%M')
    sid     = ws.id

    # ── 統一欄位順序：出發日期 / 泊數 / 天數 / 訂編 / 中文姓名 / 英文姓名 / 備註 ──
    N_COLS  = 7
    HEADERS = ['出發日期', '泊數', '天數', '訂編', '中文姓名', '英文姓名', '備註']

    rows   = []   # list of list[str]
    colors = []   # list of color key

    def _add_title(text):
        rows.append([text] + [''] * (N_COLS - 1))
        colors.append('title')

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

    # ── 同訂單去重：只保留首位旅客（with 人數計算）──
    def _dedupe_by_order(items):
        """同訂編只保留第一筆，並計算同訂單總人數"""
        order_count = {}
        for x in items:
            order_count[x['訂編']] = order_count.get(x['訂編'], 0) + 1
        seen = set()
        result = []
        for x in items:
            if x['訂編'] in seen:
                continue
            seen.add(x['訂編'])
            y = dict(x)
            y['_total_pax'] = order_count[x['訂編']]
            result.append(y)
        return result

    # ── 依出發日期排序 + 分組著色 ──
    def _date_color_key(date_str, group_idx):
        """單數日期 → date_a，雙數 → date_b（交替底色區隔不同日期）"""
        return 'date_a' if group_idx % 2 == 0 else 'date_b'

    # 更新時間
    rows.append([f'更新時間：{now_str}'] + [''] * (N_COLS - 1))
    colors.append(None)
    _add_empty()

    # ── ① 需要訂房（全新訂單）──
    _add_title('▼ 需要訂房（新訂單）')
    if need_booking:
        _add_header()
        dedup = _dedupe_by_order(need_booking)
        dedup.sort(key=lambda x: (x.get('出發日期', ''), x['訂編']))
        prev_date = None
        group_idx = -1
        for x in dedup:
            dep = x.get('出發日期', '')
            if dep != prev_date:
                group_idx += 1
                prev_date = dep
            c = _date_color_key(dep, group_idx)
            extra = f"共{x['_total_pax']}人" if x['_total_pax'] > 1 else ''
            note  = x.get('備註', '') or ''
            if extra:
                note = f"{extra}  {note}".strip()
            _add_data(
                [dep, x.get('泊數', ''), x.get('天數', ''),
                 x['訂編'], x['中文姓名'], x.get('英文姓名', ''), note],
                c,
            )
    else:
        _add_none()
    _add_empty()

    # ── ② 同訂單新增人員 ──
    _add_title('▼ 同訂單新增人員（已訂房，請通知旅館多訂房）')
    if need_add_pax:
        _add_header()
        dedup = _dedupe_by_order(need_add_pax)
        dedup.sort(key=lambda x: (x.get('出發日期', ''), x['訂編']))
        prev_date = None
        group_idx = -1
        for x in dedup:
            dep = x.get('出發日期', '')
            if dep != prev_date:
                group_idx += 1
                prev_date = dep
            c = _date_color_key(dep, group_idx)
            note = f"+{x['_total_pax']}人加入"
            _add_data(
                [dep, x.get('泊數', ''), '',
                 x['訂編'], x['中文姓名'], x.get('英文姓名', ''), note],
                c,
            )
    else:
        _add_none()
    _add_empty()

    # ── ③ 姓名需確認 ──
    _add_title('▼ 姓名需確認（請更新訂房表）')
    if name_changes:
        _add_header()
        for x in name_changes:
            note = f"原: {x.get('原中文','')} / {x.get('原英文','')}"
            _add_data(
                ['', '', '', x['訂編'], x.get('新中文', ''), x.get('新英文', ''), note],
                'name',
            )
    else:
        _add_none()
    _add_empty()

    # ── ④ 已刪減（請確認是否取消訂房）──
    _add_title('▼ 已刪減（請確認是否取消訂房）')
    if deleted_pax:
        _add_header()
        for x in deleted_pax:
            _add_data(
                [x.get('出發日期', ''), '', '',
                 x['訂編'], x['中文姓名'], x.get('英文姓名', ''), '刪減'],
                'del',
            )
    else:
        _add_none()
    _add_empty()

    # ── ⑤ 出發日期 / 泊數異動 ──
    _add_title('▼ 出發日期 / 泊數有異動（請確認住宿是否需調整）')
    if date_changes:
        _add_header()
        for x in date_changes:
            note = f"{x.get('異動欄位','')}: {x.get('原值','')} → {x.get('新值','')}"
            _add_data(
                ['', '', '', x['訂編'], x.get('中文姓名', ''), '', note],
                'date',
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
        'title' : C_HDR,
        'header': {'red': 0.3, 'green': 0.3, 'blue': 0.3},
        'date_a': {'red': 0.94, 'green': 0.97, 'blue': 1.00},   # 淡藍
        'date_b': {'red': 1.00, 'green': 0.97, 'blue': 0.88},   # 淡黃
        'name'  : C_NAME,
        'del'   : C_DEL,
        'date'  : C_DATE,
    }
    TEXT_WHITE = {'red': 1.0, 'green': 1.0, 'blue': 1.0}

    fmt_reqs = []
    for i, (row_data, color_key) in enumerate(zip(rows, colors), 1):
        if color_key is None:
            continue
        color = COLOR_MAP.get(color_key, C_WHITE)
        fmt_reqs.append({
            'repeatCell': {
                'range': {'sheetId': sid, 'startRowIndex': i - 1, 'endRowIndex': i,
                          'startColumnIndex': 0, 'endColumnIndex': N_COLS},
                'cell' : {'userEnteredFormat': {
                    'backgroundColor': color,
                    'textFormat'     : {
                        'bold': color_key in ('title', 'header'),
                        'foregroundColor': TEXT_WHITE if color_key in ('title', 'header') else {'red': 0.1, 'green': 0.1, 'blue': 0.1},
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
        len(need_booking), len(need_add_pax),
        len(name_changes), len(deleted_pax), len(date_changes),
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

    booking_map, order_map, _ = _read_booking_table(sh_booking)
    print(f"  訂房表現有旅客：{len(booking_map)} 筆，"
          f"共 {len(order_map)} 個訂單")

    need_booking, need_add_pax, name_changes, deleted_pax, date_changes = \
        _compare_nozawa(all_records_flat, booking_map, order_map)

    print(f"  新訂單需訂房：{len(set(x['訂編'] for x in need_booking))} 個訂單"
          f"（{len(need_booking)} 人）")
    print(f"  同訂單新增人員：{len(need_add_pax)} 人")
    print(f"  姓名異動：{len(name_changes)} 筆")
    print(f"  刪減旅客：{len(deleted_pax)} 人")
    print(f"  日期/泊數異動：{len(date_changes)} 筆")

    try:
        _write_compare_report(
            sh_booking,
            need_booking, need_add_pax, name_changes, deleted_pax, date_changes,
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
