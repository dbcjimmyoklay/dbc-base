"""
agents/checkin.py
CHECKIN 入住員
─────────────────────────────────────────────
職責：
  Step 1 │ 讀取「野澤當地訂房」試算表 26-27 分頁
         │  抓取每位旅客的：旅館名稱、房型、房間分配備註

  Step 2 │ 寫入三個入住單
         │  • 野澤入住單26/27  ← 同時帶入訂房表的旅館/房型資訊
         │  • 斑尾入住單26/27  ← 標準寫入，保留手動欄位
         │  • 龍平入住單26/27  ← 標準寫入，保留手動欄位
         │  手動欄位（房型/入住備註）永不覆蓋

  Step 3 │ 回寫「課程安排26/27」
         │  對有旅館資訊的野澤旅客，將「項目」欄更新為旅館名稱
         │  「項目1」欄保持原值（野澤），作為原始紀錄
─────────────────────────────────────────────
注意：
  • CHECKIN 依賴你已手動填好的野澤訂房表
  • 訂房表空白的旅客 → 入住表的項目維持「野澤」，房型/入住備註留空
  • 此 agent 應在你確認訂房後手動觸發，而非每日自動執行
─────────────────────────────────────────────
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shared.config import (
    SHEET_COURSE, COURSE_COLS, COURSE_HEADERS,
    CHECKIN_SHEETS, CHECKIN_COLS, CHECKIN_HEADERS,
    CHECKIN_MANUAL_COLS, CHECKIN_CENTER_COLS,
    CHECKIN_COL_WIDTH, COL_WIDTH,
    COLOR,
    NOZAWA_BOOKING_SPREADSHEET_ID, NOZAWA_BOOKING_SHEET,
    CREDS_PATH, SCOPES,
)
from shared.gsheet import (
    get_spreadsheet, get_sheet_by_gid,
    clear_rows_from, rebuild_filter, batch_update,
    req_row_color, req_cell_color,
    req_header_format, req_separator_format,
    req_center_align, req_borders,
    req_col_width, req_row_height,
)
import gspread
from google.oauth2.service_account import Credentials


# ── 欄位索引 ──
_CI_IDX  = {col: i for i, col in enumerate(CHECKIN_COLS)}
_CRS_IDX = {col: i for i, col in enumerate(COURSE_COLS)}


def _is_sep_row(row):
    """判斷是否為分隔列（第3格是曜日）"""
    return bool(row) and len(row) >= 3 and row[2] in ('一', '二', '三', '四', '五', '六', '日')


# ═══════════════════════════════════════════
# Step 1：讀取野澤訂房表
# ═══════════════════════════════════════════

def _read_booking_info(sh_booking):
    """
    讀取野澤訂房表 26-27，建立訂房資訊 map
    key   → 訂編_中文姓名（或 訂編_en_英文姓名）
    value → {
        '旅館'    : str,   ← 更新到課程安排「項目」欄
        '房型'    : str,   ← 更新到入住單「房型」欄
        '入住備註': str,   ← 更新到入住單「入住備註」欄（房間分配）
    }
    """
    ws       = sh_booking.worksheet(NOZAWA_BOOKING_SHEET)
    all_vals = ws.get_all_values()
    info_map = {}

    if not all_vals:
        return info_map

    # 第1列是表頭，用欄位名稱動態對應
    hrow    = all_vals[0]
    col_map = {h.strip(): i for i, h in enumerate(hrow) if h.strip()}

    idx_ding  = col_map.get('訂編',    -1)
    idx_cn    = col_map.get('中文姓名', -1)
    idx_en    = col_map.get('英文姓名', -1)
    idx_hotel = col_map.get('旅館',    -1)
    idx_room  = col_map.get('房型',    -1)
    idx_note  = col_map.get('房間分配', -1)

    for row in all_vals[1:]:
        def _g(idx):
            return row[idx].strip() if 0 <= idx < len(row) else ''

        ding  = _g(idx_ding)
        if not ding or not ding.isdigit():
            continue

        cn    = _g(idx_cn)
        en    = _g(idx_en)
        hotel = _g(idx_hotel)
        room  = _g(idx_room)
        note  = _g(idx_note)

        key = f"{ding}_{cn}" if cn else f"{ding}_en_{en}"
        info_map[key] = {
            '旅館'    : hotel,
            '房型'    : room,
            '入住備註': note,
        }

    hotel_count = sum(1 for v in info_map.values() if v['旅館'])
    print(f"  野澤訂房表：{len(info_map)} 筆，其中 {hotel_count} 筆已有旅館資訊")
    return info_map


# ═══════════════════════════════════════════
# Step 2：建立並寫入入住單
# ═══════════════════════════════════════════

def _read_existing_checkin(ws_ci):
    """
    讀取現有入住單，建立 key → {欄位: 值} 索引
    入住單結構：第1列=標題，第2列起=資料（與房務表不同）
    """
    all_vals = ws_ci.get_all_values()
    ci_map   = {}

    if not all_vals:
        return ci_map

    hrow    = all_vals[0]                        # 第1列是標題
    col_map = {h.strip(): i for i, h in enumerate(hrow) if h.strip()}

    idx_visitor = col_map.get('旅客編號', -1)
    idx_ding    = col_map.get('訂編',    -1)
    idx_cn      = col_map.get('中文姓名', -1)

    for row in all_vals[1:]:                     # 第2列起是資料
        def _g(idx):
            return row[idx].strip() if 0 <= idx < len(row) else ''

        visitor_id = _g(idx_visitor)
        ding       = _g(idx_ding)
        cn         = _g(idx_cn)

        if visitor_id:
            key = f"v_{visitor_id}"
        elif ding and cn:
            key = f"{ding}_{cn}"
        else:
            continue

        ci_map[key] = {h: row[i] if i < len(row) else ''
                       for h, i in col_map.items()}
    return ci_map


def _build_checkin_rows(all_sheet_rows, include_fn, existing_ci_map,
                        booking_info_map, is_nozawa):
    """
    建立單一入住單的寫入列
    is_nozawa=True 時從 booking_info_map 補旅館/房型資訊
    """
    write_rows = []
    n          = len(CHECKIN_COLS)

    for item in all_sheet_rows:
        if item['_is_separator']:
            sep    = [''] * n
            src    = item['_data']
            sep[0] = src[1]              # 日期（第1欄）
            sep[1] = src[2]              # 曜日（第2欄）
            write_rows.append(sep)
            continue

        r = item['_data']
        if not r.get('訂編', ''):
            continue

        item_val = r.get('項目', '')
        if not include_fn(item_val):
            continue

        ding       = r.get('訂編', '')
        cn         = r.get('中文姓名', '')
        en         = r.get('英文姓名', '')
        visitor_id = r.get('旅客編號', '')

        # ── key 與手動欄位保留 ──
        ci_key = f"v_{visitor_id}" if visitor_id else f"{ding}_{cn}"
        bk_key = f"{ding}_{cn}"    if cn          else f"{ding}_en_{en}"

        r_out = dict(r)              # 不污染原始 r

        # 保留現有手動欄位
        if ci_key in existing_ci_map:
            old = existing_ci_map[ci_key]
            for mc in CHECKIN_MANUAL_COLS:
                if mc in old and old[mc]:
                    r_out[mc] = old[mc]

        # ── 野澤：從訂房表補旅館/房型/備註 ──
        if is_nozawa and bk_key in booking_info_map:
            bk = booking_info_map[bk_key]
            hotel = bk.get('旅館', '')
            room  = bk.get('房型', '')
            note  = bk.get('入住備註', '')

            # 項目欄：若訂房表有旅館名稱則使用，否則保持「野澤」
            if hotel:
                r_out['項目'] = hotel

            # 房型/入住備註：訂房表有值才填（不覆蓋已有的手動值）
            if room and not r_out.get('房型', ''):
                r_out['房型'] = room
            if note and not r_out.get('入住備註', ''):
                r_out['入住備註'] = note

        row_data = [str(r_out.get(c, '') or '') for c in CHECKIN_COLS]
        write_rows.append(row_data)

    return write_rows


def _apply_checkin_format(sh, ws_ci, write_rows, include_fn):
    """套用入住單格式"""
    sid    = ws_ci.id
    n      = len(CHECKIN_COLS)
    reqs   = []

    # 標題列（第1列，index=0，row_idx=1）
    reqs.append(req_header_format(sid, 1, n))

    # 分隔列 + 整列底色
    for ridx, row in enumerate(write_rows, 2):   # 資料從第2列起
        if _is_sep_row(row):
            reqs.append(req_separator_format(sid, ridx, n))
            continue
        if all(v == '' for v in row):
            continue

        # 由 include_fn 推算雪場（入住單只有一個雪場，但保留邏輯彈性）
        item_val = row[0].strip() if row else ''
        if '斑尾' in item_val:
            bg = COLOR['row_blue']
        elif '湯澤' in item_val:
            bg = COLOR['row_orange']
        elif '龍平' in item_val:
            bg = COLOR['row_brown']
        else:
            bg = COLOR['white']
        reqs.append(req_row_color(sid, ridx, n, bg))

    total_rows = len(write_rows) + 1    # +1 標題列

    # 格線
    reqs.append(req_borders(sid, 1, total_rows, 0, n))

    # 置中
    col_idx_map = {c: i for i, c in enumerate(CHECKIN_COLS)}
    for cn, ci in col_idx_map.items():
        if cn in CHECKIN_CENTER_COLS:
            reqs.append(req_center_align(sid, 1, total_rows, ci))

    # 欄寬
    for ci, cn in enumerate(CHECKIN_COLS):
        px = CHECKIN_COL_WIDTH.get(cn, COL_WIDTH.get(cn, 70))
        reqs.append(req_col_width(sid, ci, px))

    # 列高
    reqs.append(req_row_height(sid, 0, 1, 30))
    reqs.append(req_row_height(sid, 1, total_rows, 24))

    batch_update(sh, reqs)


def _write_one_checkin(sh, ci_sheet, all_sheet_rows,
                       booking_info_map):
    """寫入單一入住單分頁"""
    name       = ci_sheet['name']
    gid        = ci_sheet['gid']
    include_fn = ci_sheet['include']
    is_nozawa  = '野澤' in name

    print(f"\n  寫入 {name}...")

    # 找到分頁（用 gid）
    ws_ci = get_sheet_by_gid(sh, gid)
    if ws_ci is None:
        ws_ci = sh.add_worksheet(title=name, rows=2000, cols=len(CHECKIN_COLS) + 2)
        print(f"    建立新分頁：{name}")

    # 讀取現有手動欄位
    existing_ci_map = _read_existing_checkin(ws_ci)
    print(f"    現有：{len(existing_ci_map)} 筆")

    # 建立寫入列
    write_rows = _build_checkin_rows(
        all_sheet_rows, include_fn, existing_ci_map,
        booking_info_map if is_nozawa else {},
        is_nozawa,
    )
    data_count = sum(
        1 for r in write_rows
        if r and not _is_sep_row(r) and any(v for v in r)
    )

    # 清除第2列以後（保留第1列標題）
    clear_rows_from(sh, ws_ci, start_row_index=1, n_cols=len(CHECKIN_COLS))

    # 寫入標題
    ws_ci.update([CHECKIN_HEADERS], 'A1', value_input_option='RAW')

    # 寫入資料（從第2列）
    if write_rows:
        ws_ci.update(write_rows, 'A2', value_input_option='RAW')

    # 凍結第1列
    ws_ci.freeze(rows=1)

    # 套用格式
    _apply_checkin_format(sh, ws_ci, write_rows, include_fn)

    # 重建篩選器（入住單從第1列起）
    rebuild_filter(sh, ws_ci, len(write_rows), len(CHECKIN_COLS), start_row=0)

    print(f"    OK {data_count} 筆資料")


# ═══════════════════════════════════════════
# Step 3：回寫課程安排（野澤旅館名稱）
# ═══════════════════════════════════════════

def _writeback_to_course(sh, ws_course, booking_info_map):
    """
    讀取課程安排現有內容，對野澤旅客更新「項目」欄為旅館名稱
    只更新有旅館資訊且尚未更新的列（避免覆蓋已手動修改的值）
    """
    print("\n  回寫課程安排（野澤旅館名稱）...")
    all_vals = ws_course.get_all_values()

    if len(all_vals) < 2:
        print("    課程安排無資料，略過")
        return

    hrow    = all_vals[1]                             # 第2列是標題
    col_map = {h.strip(): i for i, h in enumerate(hrow)}

    idx_item    = col_map.get('項目',    -1)
    idx_item1   = col_map.get('項目1',   -1)
    idx_ding    = col_map.get('訂編',    -1)
    idx_cn      = col_map.get('中文姓名', -1)
    idx_en      = col_map.get('英文姓名', -1)
    idx_visitor = col_map.get('旅客編號', -1)

    if idx_item < 0 or idx_ding < 0:
        print("    找不到「項目」或「訂編」欄，略過")
        return

    updates = []    # list of (row_number_1based, col_1based, value)

    for row_idx, row in enumerate(all_vals[2:], 3):  # 第3列起，row_idx 1-based
        def _g(idx):
            return row[idx].strip() if 0 <= idx < len(row) else ''

        ding       = _g(idx_ding)
        cn         = _g(idx_cn)
        en         = _g(idx_en)
        item_now   = _g(idx_item)
        item_orig  = _g(idx_item1)

        if not ding or not ding.isdigit():
            continue

        # 只處理原本是「野澤」的記錄
        if '野澤' not in item_orig and '野澤' not in item_now:
            continue

        bk_key = f"{ding}_{cn}" if cn else f"{ding}_en_{en}"
        if bk_key not in booking_info_map:
            continue

        hotel = booking_info_map[bk_key].get('旅館', '')
        if not hotel:
            continue

        # 只有當「項目」目前仍是「野澤」時才更新
        # （若已是旅館名稱或已手動改過，不覆蓋）
        if item_now == '野澤' and hotel != '野澤':
            updates.append((row_idx, idx_item + 1, hotel))

    if not updates:
        print("    無需更新（旅館資訊未變動）")
        return

    # 用 batch update 逐格更新（避免全表重寫）
    sid  = ws_course.id
    reqs = []
    for row_idx, col_idx_1based, value in updates:
        reqs.append({
            'updateCells': {
                'range': {
                    'sheetId'         : sid,
                    'startRowIndex'   : row_idx - 1,
                    'endRowIndex'     : row_idx,
                    'startColumnIndex': col_idx_1based - 1,
                    'endColumnIndex'  : col_idx_1based,
                },
                'rows'  : [{'values': [{'userEnteredValue': {'stringValue': value}}]}],
                'fields': 'userEnteredValue',
            }
        })

    batch_update(sh, reqs)
    print(f"    OK 更新 {len(updates)} 筆旅館名稱 → 課程安排「項目」欄")


# ═══════════════════════════════════════════
# 主函數
# ═══════════════════════════════════════════

def run(cleaner_output=None):
    print("=" * 50)
    print("CHECKIN 入住員 啟動")
    print("=" * 50)
    print("注意：本 agent 需要野澤訂房表已手動填寫完成")

    # ── 讀取 CLEANER 輸出 ──
    if cleaner_output is None:
        from agents.cleaner import load
        cleaner_output = load()

    all_sheet_rows = cleaner_output['all_sheet_rows']

    # ── Step 1：讀取野澤訂房表 ──
    print("\n[Step 1] 讀取野澤訂房表...")
    creds      = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    gc         = gspread.authorize(creds)
    sh_booking = gc.open_by_key(NOZAWA_BOOKING_SPREADSHEET_ID)
    booking_info_map = _read_booking_info(sh_booking)

    # ── Step 2：寫入三個入住單 ──
    print("\n[Step 2] 寫入入住單...")
    sh = get_spreadsheet()

    for ci_sheet in CHECKIN_SHEETS:
        _write_one_checkin(sh, ci_sheet, all_sheet_rows, booking_info_map)

    # ── Step 3：回寫課程安排 ──
    print("\n[Step 3] 回寫課程安排...")
    ws_course = sh.worksheet(SHEET_COURSE)
    _writeback_to_course(sh, ws_course, booking_info_map)

    print(f"\n{'=' * 50}")
    print(f"CHECKIN 完成！{datetime.now().strftime('%Y/%m/%d %H:%M')}")
    print(f"{'=' * 50}\n")


# ─────────────────────────────────────────────
if __name__ == '__main__':
    run()
