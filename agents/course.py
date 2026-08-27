"""
agents/course.py
COURSE 課程安排寫入員
─────────────────────────────────────────────
職責：
  • 從 CLEANER 輸出取得最新資料
  • 讀取現有「課程安排26/27」，保留手動欄位
    （教練/助教/checking/備註/項目旅館名稱）
  • 全部重寫課程安排（清第2列以後，保留第1列篩選按鈕）
  • 套用格式：分隔列、整列底色、單格填色、格線、欄寬、置中
  • 重建篩選器、隱藏旅客編號/項目1欄位
─────────────────────────────────────────────
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shared.config import (
    SHEET_COURSE,
    COURSE_COLS, COURSE_HEADERS,
    AUTO_COLS, MANUAL_COLS, HIDDEN_COLS,
    CENTER_COLS, COL_WIDTH,
    COLOR,
)
from shared.utils import build_record_key
from shared.gsheet import (
    get_spreadsheet, get_or_create_sheet,
    clear_rows_from, rebuild_filter, batch_update,
    req_row_color, req_cell_color,
    req_header_format, req_separator_format,
    req_center_align, req_borders,
    req_col_width, req_row_height,
    req_hide_col,
)

# ── 欄位索引 ──
_COL_IDX = {col: i for i, col in enumerate(COURSE_COLS)}
IDX_LEVEL = _COL_IDX.get('LEVEL', -1)
IDX_LEI   = _COL_IDX.get('類別',  -1)
IDX_TIAN  = _COL_IDX.get('天數',  -1)
IDX_ZHU   = _COL_IDX.get('註記',  -1)


def _is_separator(row):
    return bool(row) and len(row) >= 3 and row[2] in ('一', '二', '三', '四', '五', '六', '日')


def _build_existing_map(ws_course):
    """讀取現有課程安排，建立 key → {欄位:值} 索引"""
    all_vals     = ws_course.get_all_values()
    existing_map = {}

    if len(all_vals) <= 1:
        return existing_map

    hrow    = all_vals[1]                          # 第2列是標題
    col_map = {h: i for i, h in enumerate(hrow)}
    idx_v   = col_map.get('旅客編號', -1)
    idx_d   = col_map.get('訂編',    -1)
    idx_n   = col_map.get('中文姓名', -1)
    idx_en  = col_map.get('英文姓名', -1)

    for row in all_vals[2:]:
        def _g(idx):
            return row[idx].strip() if 0 <= idx < len(row) else ''

        visitor_id = _g(idx_v)
        ding       = _g(idx_d)
        name       = _g(idx_n)
        en         = _g(idx_en)

        if visitor_id:
            key = f"v_{visitor_id}"
        elif name:
            key = f"{ding}_{name}"
        elif en:
            key = f"{ding}_en_{en}"
        else:
            continue

        existing_map[key] = {h: row[i] if i < len(row) else ''
                             for h, i in col_map.items()}
    return existing_map


def _build_write_rows(all_sheet_rows, existing_map):
    """建立要寫入的所有列（含分隔列）"""
    write_rows = [COURSE_HEADERS]   # 第一列是標題
    n = len(COURSE_COLS)

    for item in all_sheet_rows:
        if item['_is_separator']:
            write_rows.append(item['_data'])
            continue

        # ⚠️ 這裡一定要淺複製（dict(...)），不能直接用 item['_data'] 的參照！
        # all_sheet_rows 是跟 CHECKIN／WATCHER／SALES 等其他工作員共用的同一份
        # cleaner_output 記憶體物件。若直接修改原始 dict，下面對「項目」欄的
        # 保護邏輯（改成旅館名稱）會外溢污染到後面才執行的 CHECKIN，
        # 導致已經配好房的訂單反而在 CHECKIN 被誤判排除（項目不再含地區名稱）。
        r = dict(item['_data'])

        # 空列（無資料日期）
        if not r.get('訂編', ''):
            write_rows.append([''] * n)
            continue

        key = build_record_key(r)

        if key in existing_map:
            old = existing_map[key]

            # 保留手動欄位（僅影響本次寫入 Sheet 的副本，不影響共用資料）
            for mc in MANUAL_COLS:
                if mc in old and old[mc]:
                    r[mc] = old[mc]

            # 保護項目欄：若現有值跟項目1不同，代表已手動改成旅館名稱
            old_item  = old.get('項目',  '').strip()
            old_item1 = old.get('項目1', '').strip()
            if old_item and old_item != old_item1:
                r['項目'] = old_item

        row_data = [str(r.get(col, '') or '') for col in COURSE_COLS]
        write_rows.append(row_data)

    return write_rows


def _apply_format(sh, ws_course, write_rows):
    """套用所有格式請求"""
    sid  = ws_course.id
    n    = len(COURSE_COLS)
    reqs = []

    # ── 標題列（第2列，write_rows[0] 對應 Google Sheet 第2列）──
    reqs.append(req_header_format(sid, 2, n))

    # ── 分隔列 + 整列底色 + 單格填色 ──
    for i, row in enumerate(write_rows[1:], 3):   # 第3列起是資料
        if _is_separator(row):
            reqs.append(req_separator_format(sid, i, n))
            continue

        if all(v == '' for v in row):
            continue

        # 取項目值（可能是 list 的第1欄）
        item_val = row[_COL_IDX.get('項目', 0)].strip() if row else ''
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
        reqs.append(req_row_color(sid, i, n, bg))

        # LEVEL 單格填色
        if IDX_LEVEL >= 0:
            lv = row[IDX_LEVEL].strip().lower()
            if lv in ('lv1', 'lv2'):
                reqs.append(req_cell_color(sid, i, IDX_LEVEL + 1, COLOR['yellow_light']))
            elif lv in ('lv3', 'lv4', 'lv5'):
                reqs.append(req_cell_color(sid, i, IDX_LEVEL + 1, COLOR['orange']))

        # 類別 = 雙板 → 藍綠
        if IDX_LEI >= 0 and row[IDX_LEI].strip() == '雙板':
            reqs.append(req_cell_color(sid, i, IDX_LEI + 1, COLOR['teal']))

        # 天數含「含」→ 亮黃
        if IDX_TIAN >= 0 and '含' in row[IDX_TIAN]:
            reqs.append(req_cell_color(sid, i, IDX_TIAN + 1, COLOR['yellow_bright']))

        # 註記 = 專屬 → 亮黃
        if IDX_ZHU >= 0 and row[IDX_ZHU].strip() == '專屬':
            reqs.append(req_cell_color(sid, i, IDX_ZHU + 1, COLOR['yellow_bright']))

    total_rows = len(write_rows) + 1   # +1 因從第2列起算

    # ── 格線 ──
    reqs.append(req_borders(sid, 2, total_rows, 0, n))

    # ── 置中 ──
    for col_name, col_idx in _COL_IDX.items():
        if col_name in CENTER_COLS:
            reqs.append(req_center_align(sid, 2, total_rows, col_idx))

    # ── 欄寬 ──
    for ci, cn in enumerate(COURSE_COLS):
        px = COL_WIDTH.get(cn, 80)
        reqs.append(req_col_width(sid, ci, px))

    # ── 列高 ──
    reqs.append(req_row_height(sid, 1, 2, 30))
    reqs.append(req_row_height(sid, 2, total_rows, 24))

    # ── 隱藏欄（旅客編號/項目1）──
    for col_name in HIDDEN_COLS:
        if col_name in _COL_IDX:
            reqs.append(req_hide_col(sid, _COL_IDX[col_name]))

    batch_update(sh, reqs)
    print(f"  OK 格式設定完成")


def run(cleaner_output=None):
    print("=" * 50)
    print("COURSE 課程安排寫入員 啟動")
    print("=" * 50)

    if cleaner_output is None:
        from agents.cleaner import load
        cleaner_output = load()

    all_sheet_rows = cleaner_output['all_sheet_rows']

    print("連線 Google Sheets...")
    sh        = get_spreadsheet()
    ws_course = get_or_create_sheet(sh, SHEET_COURSE, rows=2000, cols=len(COURSE_COLS) + 2)

    # 保護機制
    print("讀取現有課程安排...")
    existing_map = _build_existing_map(ws_course)
    print(f"  現有：{len(existing_map)} 筆")

    if ws_course.row_count > 10 and len(existing_map) == 0:
        print("警告：Google Sheet 有資料但讀取失敗，停止執行！")
        return

    # 建立寫入列
    write_rows = _build_write_rows(all_sheet_rows, existing_map)
    data_count = sum(
        1 for r in write_rows[1:]
        if r and not _is_separator(r) and any(v for v in r)
    )
    print(f"  準備寫入：{data_count} 筆資料列")

    # 清除第2列以後（保留第1列篩選按鈕）
    clear_rows_from(sh, ws_course, start_row_index=1, n_cols=len(COURSE_COLS))

    # 從第2列開始寫入（標題 + 資料）
    ws_course.update(write_rows, 'A2', value_input_option='RAW')

    # 凍結前2列
    ws_course.freeze(rows=2)

    # 格式
    print("套用格式...")
    _apply_format(sh, ws_course, write_rows)

    # 重建篩選器
    print("重建篩選器...")
    rebuild_filter(sh, ws_course, len(write_rows) - 1, len(COURSE_COLS))

    # ── 寫入後驗證：回頭讀取 Sheet，確認真的寫進去了 ──
    # 防止 Google API 靜默部分失敗（例如 batch_update 中途出錯、
    # 或前面步驟未拋出例外但實際沒寫完整），導致資料停留在舊狀態卻沒人發現。
    print("驗證寫入結果...")
    verify_map = _build_existing_map(ws_course)
    verify_count = len(verify_map)
    if verify_count < data_count * 0.9:   # 容許小幅誤差（去重、佔位資料等）
        raise RuntimeError(
            f"COURSE 寫入驗證失敗！預期寫入約 {data_count} 筆，"
            f"但回頭讀取 Sheet 只看到 {verify_count} 筆。"
            f"Google Sheets 可能沒有真的更新成功，請重新執行。"
        )
    print(f"  OK 驗證通過（Sheet 上實際有 {verify_count} 筆）")

    print(f"\n{'=' * 50}")
    print(f"COURSE 完成！共 {data_count} 筆  {datetime.now().strftime('%Y/%m/%d %H:%M')}")
    print(f"{'=' * 50}\n")


if __name__ == '__main__':
    run()
