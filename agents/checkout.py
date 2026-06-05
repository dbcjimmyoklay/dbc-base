"""
agents/checkout.py
CHECKOUT 送客員
─────────────────────────────────────────────
職責：
  Step 1 │ 讀取三個入住單，建立對照表
         │  野澤：旅客編號 → 旅館名稱（讓送客單顯示正確地點）
         │  全區：訂編+姓名 → 入住備註（供送客員判斷車票數）

  Step 2 │ 從 CLEANER 輸出計算各旅客退房日期
         │  退房日期 = 出發日期 + 泊數（天）
         │  分三區寫入送客單：野澤/斑尾/龍平

  手動欄位（退房備註）永不覆蓋

─────────────────────────────────────────────
送客單欄位：
  項目 | 退房日期 | 泊數 | 訂編 | 中文姓名 | 性別 | 年齡
  出發日期 | 入住備註 | 退房備註 | 旅客編號

分隔列（退房日期下）顯示：
  XX/XX（曜）退房 N人 ／ 下山巴士 N張 ／ 新幹線 N張

─────────────────────────────────────────────
觸發時機：
  每日 main.py 自動流程（COURSE 之後）
  單獨執行：python main.py --checkout
─────────────────────────────────────────────
"""

import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shared.config import (
    CHECKIN_SHEETS,
    CHECKOUT_SHEETS, CHECKOUT_COLS, CHECKOUT_HEADERS,
    CHECKOUT_MANUAL_COLS, CHECKOUT_CENTER_COLS,
    CHECKOUT_COL_WIDTH, COL_WIDTH,
    COLOR, WEEKDAY_MAP,
    CREDS_PATH, SCOPES,
)
from shared.gsheet import (
    get_spreadsheet, get_sheet_by_gid,
    clear_rows_from, rebuild_filter, batch_update,
    req_row_color,
    req_header_format, req_separator_format,
    req_center_align, req_borders,
    req_col_width, req_row_height,
)
import gspread


def _is_sep_row(row):
    """判斷是否為分隔列（col0=空, col2=曜日）"""
    return (
        bool(row) and len(row) >= 3
        and not str(row[0]).strip()
        and str(row[2]) in ('一', '二', '三', '四', '五', '六', '日')
    )


def _parse_date(s):
    s = str(s).strip()
    for fmt in ('%y/%m/%d', '%Y/%m/%d'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _fmt_date(d):
    return d.strftime('%y/%m/%d')


# ═══════════════════════════════════════════
# Step 1：從三個入住單讀取旅館名稱 + 入住備註
# ═══════════════════════════════════════════

def _read_checkin_info(sh):
    """
    掃描野澤/斑尾/龍平入住單，建立：
      hotel_map : 「訂編_中文姓名」 → 旅館名稱（主要用於野澤）
      note_map  : 「訂編_中文姓名」 → 入住備註（全區，供判斷車票數）
    """
    hotel_map = {}
    note_map  = {}

    for ci_cfg in CHECKIN_SHEETS:
        try:
            ws = get_sheet_by_gid(sh, ci_cfg['gid'])
            if not ws:
                continue
            all_vals = ws.get_all_values()
        except Exception as e:
            print(f"  讀取 {ci_cfg['name']} 失敗：{e}")
            continue

        if not all_vals:
            continue

        hrow    = all_vals[0]
        col_map = {h.strip(): i for i, h in enumerate(hrow)}
        idx_item = col_map.get('項目',    -1)
        idx_ding = col_map.get('訂編',    -1)
        idx_cn   = col_map.get('中文姓名', -1)
        idx_note = col_map.get('入住備註', -1)

        for row in all_vals[1:]:
            def _g(idx):
                return row[idx].strip() if 0 <= idx < len(row) else ''

            ding  = _g(idx_ding)
            if not ding or not ding.isdigit():
                continue   # 跳過分隔列或空列

            cn    = _g(idx_cn)
            hotel = _g(idx_item)
            note  = _g(idx_note)
            key   = f"{ding}_{cn}" if cn else ding

            if hotel:
                hotel_map[key] = hotel
            if note:
                note_map[key] = note

    print(f"  入住單資訊：旅館對照 {len(hotel_map)} 筆，入住備註 {len(note_map)} 筆")
    return hotel_map, note_map


# ═══════════════════════════════════════════
# Step 2：建立並寫入送客單
# ═══════════════════════════════════════════

def _read_existing_checkout(ws_co):
    """讀取現有退房備註（手動欄位），避免覆蓋"""
    all_vals = ws_co.get_all_values()
    co_map   = {}
    if not all_vals:
        return co_map

    hrow    = all_vals[0]
    col_map = {h.strip(): i for i, h in enumerate(hrow)}
    idx_v    = col_map.get('旅客編號', -1)
    idx_ding = col_map.get('訂編',    -1)
    idx_cn   = col_map.get('中文姓名', -1)
    idx_note = col_map.get('退房備註', -1)

    for row in all_vals[1:]:
        def _g(idx):
            return row[idx].strip() if 0 <= idx < len(row) else ''

        vid  = _g(idx_v)
        ding = _g(idx_ding)
        cn   = _g(idx_cn)
        note = _g(idx_note)

        if not note:
            continue

        key = f"v_{vid}" if vid else (f"{ding}_{cn}" if cn else ding)
        if key:
            co_map[key] = note

    return co_map


def _build_checkout_rows(all_sheet_rows, include_fn, existing_co_map,
                         hotel_map, note_map):
    """
    建立單一送客單的寫入列
    退房日期 = 出發日期 + 泊數，按日期升序排列
    分隔列 col 格式：
      col0=空, col1=退房日期, col2=曜日, col3=退房人數（供 Sheet 參考）
    """
    n = len(CHECKOUT_COLS)

    # ── Stage 1：按退房日期收集 ──
    date_wday  = {}   # {checkout_str: wday}
    date_order = []
    date_data  = {}   # {checkout_str: [row_list, ...]}

    for item in all_sheet_rows:
        if item['_is_separator']:
            continue

        r = item['_data']
        if not r.get('訂編', ''):
            continue

        item1 = str(r.get('項目1', '') or r.get('項目', '') or '').strip()
        if not include_fn(item1):
            continue

        depart_str = str(r.get('出發日期', '') or '').strip()
        nights_str = str(r.get('泊數', '') or '').strip()
        if not depart_str or not nights_str:
            continue

        try:
            nights = int(nights_str)
        except ValueError:
            continue

        depart_dt = _parse_date(depart_str)
        if not depart_dt:
            continue

        checkout_dt  = depart_dt + timedelta(days=nights)
        checkout_str = _fmt_date(checkout_dt)
        wday         = WEEKDAY_MAP[checkout_dt.weekday()]

        if checkout_str not in date_data:
            date_wday[checkout_str] = wday
            date_order.append(checkout_str)
            date_data[checkout_str] = []

        ding       = str(r.get('訂編', '') or '').strip()
        cn         = str(r.get('中文姓名', '') or '').strip()
        visitor_id = str(r.get('旅客編號', '') or '').strip()

        lookup_key = f"{ding}_{cn}" if cn else ding
        co_key     = f"v_{visitor_id}" if visitor_id else lookup_key

        # 旅館名稱：野澤從入住單取，其他用地區名
        hotel = hotel_map.get(lookup_key, '') or item1

        # 入住備註（從入住單帶入）
        ci_note = note_map.get(lookup_key, '')

        # 退房備註（手動，保留）
        manual_note = existing_co_map.get(co_key, '')

        row_out = [
            hotel,                                          # 0 項目
            checkout_str,                                   # 1 退房日期
            nights_str,                                     # 2 泊數
            ding,                                           # 3 訂編
            cn,                                             # 4 中文姓名
            str(r.get('性別', '') or ''),                   # 5 性別
            str(r.get('年齡', '') or ''),                   # 6 年齡
            depart_str,                                     # 7 出發日期
            ci_note,                                        # 8 入住備註
            manual_note,                                    # 9 退房備註
            visitor_id,                                     # 10 旅客編號
        ]
        date_data[checkout_str].append(row_out)

    # ── Stage 2：按日期排序輸出 ──
    date_order.sort()   # YY/MM/DD 字串可直接排序
    write_rows = []

    for checkout_str in date_order:
        rows = date_data.get(checkout_str, [])
        if not rows:
            continue

        total = len(rows)

        # 分隔列：col0=空, col1=退房日期, col2=曜日, col3=人數
        wday = date_wday[checkout_str]
        sep  = [''] * n
        sep[1] = checkout_str
        sep[2] = wday
        sep[3] = str(total)   # 便於在 Sheet 直接看到人數
        write_rows.append(sep)

        # 資料列：訂編排序（保持同日期同行程旅客相鄰）
        rows.sort(key=lambda r: r[3])   # 訂編
        write_rows.extend(rows)

    return write_rows


def _apply_checkout_format(sh, ws_co, write_rows):
    """套用送客單格式（標題/分隔/資料列顏色、欄寬、格線）"""
    sid  = ws_co.id
    n    = len(CHECKOUT_COLS)
    reqs = []

    reqs.append(req_header_format(sid, 1, n))

    for ridx, row in enumerate(write_rows, 2):
        if _is_sep_row(row):
            reqs.append(req_separator_format(sid, ridx, n))
            continue
        if all(v == '' for v in row):
            continue
        # 以 項目（col0）判斷地區底色
        item_val = str(row[0]).strip() if row else ''
        if '斑尾' in item_val:
            bg = COLOR['row_blue']
        elif '龍平' in item_val:
            bg = COLOR['row_brown']
        else:
            bg = COLOR['white']
        reqs.append(req_row_color(sid, ridx, n, bg))

    total_rows = len(write_rows) + 1

    reqs.append(req_borders(sid, 1, total_rows, 0, n))

    for ci, cn in enumerate(CHECKOUT_COLS):
        if cn in CHECKOUT_CENTER_COLS:
            reqs.append(req_center_align(sid, 1, total_rows, ci))

    for ci, cn in enumerate(CHECKOUT_COLS):
        px = CHECKOUT_COL_WIDTH.get(cn, COL_WIDTH.get(cn, 70))
        if px == 0:     # 旅客編號：隱藏欄（寬度 0 等效極窄）
            px = 25
        reqs.append(req_col_width(sid, ci, px))

    reqs.append(req_row_height(sid, 0, 1, 30))
    reqs.append(req_row_height(sid, 1, total_rows, 24))

    batch_update(sh, reqs)


def _write_one_checkout(sh, co_cfg, all_sheet_rows,
                        hotel_map, note_map):
    """寫入單一送客單分頁"""
    name       = co_cfg['name']
    include_fn = co_cfg['include']

    print(f"\n  寫入 {name}...")

    try:
        ws_co = sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws_co = sh.add_worksheet(title=name, rows=2000, cols=len(CHECKOUT_COLS) + 2)
        print(f"    建立新分頁：{name}")

    existing_co_map = _read_existing_checkout(ws_co)
    print(f"    現有退房備註：{len(existing_co_map)} 筆")

    write_rows = _build_checkout_rows(
        all_sheet_rows, include_fn, existing_co_map,
        hotel_map, note_map,
    )
    data_count = sum(1 for r in write_rows if r and not _is_sep_row(r) and any(v for v in r))

    clear_rows_from(sh, ws_co, start_row_index=1, n_cols=len(CHECKOUT_COLS))
    ws_co.update([CHECKOUT_HEADERS], 'A1', value_input_option='RAW')

    if write_rows:
        ws_co.update(write_rows, 'A2', value_input_option='RAW')

    ws_co.freeze(rows=1)
    _apply_checkout_format(sh, ws_co, write_rows)
    rebuild_filter(sh, ws_co, len(write_rows), len(CHECKOUT_COLS), start_row=0)

    print(f"    OK {data_count} 筆資料")


# ═══════════════════════════════════════════
# 主函數
# ═══════════════════════════════════════════

def run(cleaner_output=None):
    print("=" * 50)
    print("CHECKOUT 送客員 啟動")
    print("=" * 50)

    if cleaner_output is None:
        from agents.cleaner import load
        cleaner_output = load()

    all_sheet_rows = cleaner_output['all_sheet_rows']
    sh = get_spreadsheet()

    print("\n[Step 1] 讀取入住單資訊（旅館對照 + 入住備註）...")
    hotel_map, note_map = _read_checkin_info(sh)

    print("\n[Step 2] 寫入送客單...")
    for co_cfg in CHECKOUT_SHEETS:
        _write_one_checkout(sh, co_cfg, all_sheet_rows, hotel_map, note_map)

    print(f"\n{'=' * 50}")
    print(f"CHECKOUT 完成！{datetime.now().strftime('%Y/%m/%d %H:%M')}")
    print(f"{'=' * 50}\n")


if __name__ == '__main__':
    run()
