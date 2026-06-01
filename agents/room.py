"""
agents/room.py
ROOM 房務員
─────────────────────────────────────────────
職責：
  • 從 CLEANER 輸出取得最新資料
  • 讀取現有「房務表26/27」，保留手動欄位（房號/序號/備註）
  • 全部重寫房務表（清第3列以後，保留第1列篩選按鈕、第2列標題）
  • 套用格式：分隔列、整列底色、泊數異常標黃、格線、欄寬、置中
  • 重建篩選器

排除：湯澤純課不寫入房務表
─────────────────────────────────────────────
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shared.config import (
    SHEET_ROOM,
    ROOM_COLS, ROOM_HEADERS, ROOM_MANUAL_COLS,
    ROOM_CENTER_COLS, ROOM_EXCLUDE_ITEMS,
    ROOM_COL_WIDTH, COL_WIDTH,
    COLOR,
)
from shared.gsheet import (
    get_spreadsheet, get_or_create_sheet,
    clear_rows_from, rebuild_filter, batch_update,
    req_row_color, req_cell_color,
    req_header_format, req_separator_format,
    req_center_align, req_borders,
    req_col_width, req_row_height,
)


# ── 欄位索引（由 ROOM_COLS 決定）──
_COL_IDX = {col: i for i, col in enumerate(ROOM_COLS)}
IDX_ITEM   = _COL_IDX.get('項目',   -1)
IDX_POU    = _COL_IDX.get('泊數',   -1)
IDX_DATE   = _COL_IDX.get('出發日期', -1)


def _is_separator(row):
    """判斷某列是否為分隔列（曜日欄有值）"""
    if not row or len(row) < 3:
        return False
    return row[2] in ('一', '二', '三', '四', '五', '六', '日')


def _build_existing_map(ws_room):
    """
    讀取現有房務表，建立 (訂編_中文姓名) → {欄位: 值} 索引
    結構：第1列=篩選按鈕，第2列=標題，第3列起=資料
    """
    all_vals = ws_room.get_all_values()
    existing_map = {}

    if len(all_vals) <= 1:
        return existing_map

    hrow   = all_vals[1]                         # 第2列是標題
    cmap   = {h: i for i, h in enumerate(hrow)}
    idx_d  = cmap.get('訂編',   -1)
    idx_n  = cmap.get('中文姓名', -1)

    for row in all_vals[2:]:                     # 第3列起是資料
        ding = row[idx_d].strip() if 0 <= idx_d < len(row) else ''
        name = row[idx_n].strip() if 0 <= idx_n < len(row) else ''
        if not ding or not ding.isdigit():
            continue
        if not name:
            continue
        key = f"{ding}_{name}"
        existing_map[key] = {h: row[i] if i < len(row) else ''
                              for h, i in cmap.items()}
    return existing_map


def _build_write_rows(all_sheet_rows, existing_map):
    """
    建立要寫入的所有列（含分隔列）
    回傳：list of list[str]
    """
    write_rows = []
    n = len(ROOM_COLS)

    for item in all_sheet_rows:
        if item['_is_separator']:
            # ── 分隔列 ──
            src = item['_data']          # CLEANER 產生的 COURSE_COLS 長度列
            sep = [''] * n
            sep[1] = src[1]             # 日期（第2欄）
            sep[2] = src[2]             # 曜日（第3欄）
            # 房務表統計：野/斑/龍（不顯示湯澤）
            ya  = src[3].replace('野:', '') if len(src) > 3 and src[3].startswith('野:') else '0'
            ban = src[4].replace('斑:', '') if len(src) > 4 and src[4].startswith('斑:') else '0'
            lon = src[6].replace('龍:', '') if len(src) > 6 and src[6].startswith('龍:') else '0'
            sep[3] = f"野:{ya}"
            sep[4] = f"斑:{ban}"
            sep[5] = f"龍:{lon}"
            write_rows.append(sep)
            continue

        r = item['_data']

        # 略過空列（無資料日期的佔位列）
        if not r.get('訂編', ''):
            write_rows.append([''] * n)
            continue

        # 排除湯澤純課
        if r.get('項目', '') in ROOM_EXCLUDE_ITEMS:
            continue

        # 保留手動欄位
        key = f"{r.get('訂編','')}_{r.get('中文姓名','')}"
        if key in existing_map:
            old = existing_map[key]
            for mc in ROOM_MANUAL_COLS:
                if mc in old and old[mc]:
                    r[mc] = old[mc]

        # 備註初始值：若手動備註空，用科威的分房備住
        if not r.get('備註', '') and r.get('_房務備註', ''):
            r['備註'] = r['_房務備註']

        row_data = [str(r.get(c, '') or '') for c in ROOM_COLS]
        write_rows.append(row_data)

    # ── 補空列：若某天過濾後只剩分隔列，補一列空列 ──
    final = []
    i = 0
    while i < len(write_rows):
        row = write_rows[i]
        final.append(row)
        if _is_separator(row):
            next_i = i + 1
            if next_i >= len(write_rows) or _is_separator(write_rows[next_i]):
                final.append([''] * n)
        i += 1

    return final


def _apply_format(sh, ws_room, write_rows):
    """套用所有格式請求"""
    sid    = ws_room.id
    n      = len(ROOM_COLS)
    reqs   = []

    # ── 標題列（第2列，row_idx=2）黑底白字 ──
    reqs.append(req_header_format(sid, 2, n))

    # ── 分隔列 + 整列底色 + 泊數異常標黃 ──
    for ridx, row in enumerate(write_rows, 3):   # 資料從第3列(index=2)起
        if _is_separator(row):
            reqs.append(req_separator_format(sid, ridx, n))
            continue

        if all(v == '' for v in row):
            continue

        item_val = row[IDX_ITEM].strip() if IDX_ITEM >= 0 < len(row) else ''
        if not item_val:
            continue

        # 整列底色（依雪場）
        if '斑尾' in item_val:
            bg = COLOR['row_blue']
        elif '湯澤' in item_val:
            bg = COLOR['row_orange']
        elif '龍平' in item_val:
            bg = COLOR['row_brown']
        else:
            bg = COLOR['white']
        reqs.append(req_row_color(sid, ridx, n, bg))

        # 泊數 ≠ 3 → 單格亮黃
        if IDX_POU >= 0 and IDX_POU < len(row):
            pou_val = row[IDX_POU].strip()
            if pou_val and pou_val != '3':
                reqs.append(req_cell_color(sid, ridx, IDX_POU + 1, COLOR['yellow_bright']))

    total_rows = len(write_rows) + 2    # +2：篩選按鈕列 + 標題列

    # ── 格線（從第2列標題開始）──
    reqs.append(req_borders(sid, 2, total_rows, 0, n))

    # ── 置中 ──
    col_idx_map = {c: i for i, c in enumerate(ROOM_COLS)}
    for cn, ci in col_idx_map.items():
        if cn in ROOM_CENTER_COLS:
            reqs.append(req_center_align(sid, 2, total_rows, ci))

    # ── 欄寬 ──
    for ci, cn in enumerate(ROOM_COLS):
        px = ROOM_COL_WIDTH.get(cn, COL_WIDTH.get(cn, 80))
        reqs.append(req_col_width(sid, ci, px))

    # ── 列高：標題列30，其他24 ──
    reqs.append(req_row_height(sid, 1, 2, 30))     # 第2列（index=1）標題列
    reqs.append(req_row_height(sid, 2, total_rows, 24))

    batch_update(sh, reqs)
    print(f"  OK 格式設定完成")


def run(cleaner_output=None):
    print("=" * 50)
    print("ROOM 房務員 啟動")
    print("=" * 50)

    # ── 讀取 CLEANER 輸出 ──
    if cleaner_output is None:
        from agents.cleaner import load
        cleaner_output = load()

    all_sheet_rows   = cleaner_output['all_sheet_rows']
    all_records_flat = cleaner_output['all_records_flat']

    # ── 連線 Google Sheets ──
    print("連線 Google Sheets...")
    sh      = get_spreadsheet()
    ws_room = get_or_create_sheet(sh, SHEET_ROOM, rows=2000, cols=len(ROOM_COLS) + 2)

    # ── 讀取現有資料，保留手動欄位 ──
    print("讀取現有房務表...")
    existing_map = _build_existing_map(ws_room)
    print(f"  現有：{len(existing_map)} 筆")

    # ── 建立寫入列 ──
    write_rows = _build_write_rows(all_sheet_rows, existing_map)
    data_count = sum(1 for r in write_rows if r and not _is_separator(r) and any(v for v in r))
    print(f"  準備寫入：{data_count} 筆資料列")

    # ── 清除第3列以後（保留第1列篩選按鈕、第2列標題）──
    clear_rows_from(sh, ws_room, start_row_index=2, n_cols=len(ROOM_COLS))

    # ── 確保標題列 ──
    ws_room.update([ROOM_HEADERS], 'A2', value_input_option='RAW')

    # ── 從第3列開始寫入資料 ──
    if write_rows:
        ws_room.update(write_rows, 'A3', value_input_option='RAW')

    # ── 凍結前2列 ──
    ws_room.freeze(rows=2)

    # ── 套用格式 ──
    print("套用格式...")
    _apply_format(sh, ws_room, write_rows)

    # ── 重建篩選器 ──
    print("重建篩選器...")
    rebuild_filter(sh, ws_room, len(write_rows), len(ROOM_COLS))

    print(f"\n{'=' * 50}")
    print(f"ROOM 完成！共 {data_count} 筆  {datetime.now().strftime('%Y/%m/%d %H:%M')}")
    print(f"{'=' * 50}\n")


# ─────────────────────────────────────────────
if __name__ == '__main__':
    run()
