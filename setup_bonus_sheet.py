"""
setup_bonus_sheet.py
一次性建立「限時推薦獎金」分頁（DBC Base）＋ Portal 頁面權限欄位（使用者管理）
─────────────────────────────────────────────
用法：python setup_bonus_sheet.py            建立（分頁已有資料時不會覆蓋）
      python setup_bonus_sheet.py --force    強制重建分頁（會清掉現有內容！）

之後的日常更新：直接在 Google 試算表「限時推薦獎金」分頁編輯，Portal「限時推薦獎金」頁面即時讀取。
這支程式不在每日流程裡，不會改動你編輯的內容。
"""
import sys
import os
import warnings

import gspread
from google.oauth2.service_account import Credentials

warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from shared.config import CREDS_PATH, SCOPES, SPREADSHEET_ID, USER_MGMT_SPREADSHEET_ID  # noqa: E402

TAB = '限時推薦獎金'
PAGE_NAME = '限時推薦獎金'          # Portal 頁面名稱（要與 portal/index.html 的 PAGES 一致）
HEADERS = ['類型', '地區', '開始日期', '結束日期', '組數', '人數下限', '人數上限', '獎金', '狀態', '內部備註(不公開)']
N_PUBLIC = 9                         # A~I 會傳到 Portal；J（內部備註）不會
MAX_ROW = 300

STATUSES = ['開放中', '暫停', '已達成', '客人自行報名', '關閉']

# 目前公告（團客：出發日期；純課程：上課日期）
ROWS = [
    # 類型, 地區, 開始, 結束, 組數, 下限, 上限, 獎金, 狀態
    ['團客', '野澤', '2026/12/18', '', 1, 2, 4, 1000, '開放中'],
    ['團客', '野澤', '2026/12/22', '', 1, 2, 2, 1000, '開放中'],
    ['團客', '野澤', '2026/12/26', '', 1, 2, 2, 1000, '開放中'],
    ['團客', '野澤', '2026/12/27', '', 1, 2, 3, 1000, '開放中'],
    ['團客', '野澤', '2026/12/28', '', 1, 2, 2, 1000, '開放中'],
    ['團客', '野澤', '2026/12/29', '', 1, 2, 2, 1000, '開放中'],
    ['團客', '野澤', '2026/12/30', '', 1, 2, 3, 1000, '開放中'],
    ['團客', '斑尾', '2027/01/07', '', 3, 2, 4, 1000, '開放中'],
    ['團客', '斑尾', '2027/01/18', '', 2, 2, 4, 1000, '開放中'],
    ['團客', '龍平', '2027/02/04', '', 1, 2, 4, 1000, '開放中'],
    ['團客', '龍平', '2027/02/06', '2027/02/17', '', '', '', 1000, '開放中'],
] + [
    ['純課程', area, f, t, '', '', '', '10%', '開放中']
    for area in ('野澤', '斑尾', '湯澤')
    for (f, t) in (('2027/02/01', '2027/02/10'), ('2027/03/02', '2027/03/07'))
]

NOTES = [
    '使用說明（這一欄不會顯示在 Portal）',
    '每一列 = 一個獎金項目。Portal「限時推薦獎金」頁面會即時讀取這個分頁（A~I 欄）。',
    '類型：團客 / 純課程。團客的日期是「出發日期」，純課程的日期是「上課日期」。',
    '填 開始日期 + 組數 + 人數下限/上限 → 單日獎金，例如 12/18 1組(2-4人)。',
    '只填 開始日期（組數留空）→ 顯示「當天介紹成功都算」。',
    '填 開始日期 + 結束日期 → 區間，顯示「期間內介紹成功都算」。',
    '獎金：團客填每人金額（例 1000）；純課程填比例（例 10%）。',
    '狀態：開放中 → 顯示；暫停 → 顯示但標「暫停中」；已達成 / 客人自行報名 / 關閉 → 移到「已結束」並顯示原因。',
    '名額被用掉一部分時，直接把「組數」改小；全部用完再把狀態改成「已達成」或「客人自行報名」。',
    '日期過了會自動從頁面消失，不用手動刪除。',
    '純課程某個日期段只想開某一區，就只填那一區的那一列（每個「地區＋日期段」各自獨立一列）。',
    '人數下限/上限只是標準；實際成交人數依實際人頭計算。',
    '內部備註（J 欄）只有你看得到，不會傳到 Portal。',
]


def _rule(sid, text, rgb):
    return {'addConditionalFormatRule': {'index': 0, 'rule': {
        'ranges': [{'sheetId': sid, 'startRowIndex': 1, 'endRowIndex': MAX_ROW,
                    'startColumnIndex': 8, 'endColumnIndex': 9}],
        'booleanRule': {
            'condition': {'type': 'TEXT_EQ', 'values': [{'userEnteredValue': text}]},
            'format': {'backgroundColor': {'red': rgb[0], 'green': rgb[1], 'blue': rgb[2]}}}}}}


def _validation(sid, col, values):
    return {'setDataValidation': {
        'range': {'sheetId': sid, 'startRowIndex': 1, 'endRowIndex': MAX_ROW,
                  'startColumnIndex': col, 'endColumnIndex': col + 1},
        'rule': {'condition': {'type': 'ONE_OF_LIST', 'values': [{'userEnteredValue': v} for v in values]},
                 'showCustomUi': True, 'strict': False}}}


def setup_tab(gc, force=False):
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(TAB)
        existing = [r for r in ws.get_all_values()[1:] if any(c.strip() for c in r[:N_PUBLIC])]
        if existing and not force:
            print(f"  「{TAB}」已有 {len(existing)} 列資料，不覆蓋（要重建請加 --force）")
            return ws, False
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=TAB, rows=MAX_ROW, cols=14)
        print(f"  建立新分頁：{TAB}")
    if ws.row_count < MAX_ROW or ws.col_count < 14:
        ws.resize(rows=max(ws.row_count, MAX_ROW), cols=max(ws.col_count, 14))

    ws.update(range_name='A1', values=[HEADERS] + ROWS, value_input_option='USER_ENTERED')
    ws.update(range_name='L1', values=[[n] for n in NOTES], value_input_option='RAW')

    sid = ws.id
    dark = {'red': 0.2, 'green': 0.3, 'blue': 0.5}
    white = {'red': 1, 'green': 1, 'blue': 1}
    reqs = [
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': len(HEADERS)},
                        'cell': {'userEnteredFormat': {'backgroundColor': dark, 'horizontalAlignment': 'CENTER',
                                                       'textFormat': {'bold': True, 'foregroundColor': white}}},
                        'fields': 'userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)'}},
        {'updateSheetProperties': {'properties': {'sheetId': sid, 'gridProperties': {'frozenRowCount': 1}},
                                   'fields': 'gridProperties.frozenRowCount'}},
        # 日期欄：真正的日期格式
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 1, 'endRowIndex': MAX_ROW, 'startColumnIndex': 2, 'endColumnIndex': 4},
                        'cell': {'userEnteredFormat': {'numberFormat': {'type': 'DATE', 'pattern': 'yyyy/mm/dd'},
                                                       'horizontalAlignment': 'CENTER'}},
                        'fields': 'userEnteredFormat(numberFormat,horizontalAlignment)'}},
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 1, 'endRowIndex': MAX_ROW, 'startColumnIndex': 0, 'endColumnIndex': 2},
                        'cell': {'userEnteredFormat': {'horizontalAlignment': 'CENTER'}}, 'fields': 'userEnteredFormat.horizontalAlignment'}},
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 1, 'endRowIndex': MAX_ROW, 'startColumnIndex': 4, 'endColumnIndex': 9},
                        'cell': {'userEnteredFormat': {'horizontalAlignment': 'CENTER'}}, 'fields': 'userEnteredFormat.horizontalAlignment'}},
        _validation(sid, 0, ['團客', '純課程']),
        _validation(sid, 1, ['野澤', '斑尾', '龍平', '湯澤']),
        _validation(sid, 8, STATUSES),
        _rule(sid, '開放中', (0.78, 0.94, 0.80)),
        _rule(sid, '暫停', (1.0, 0.92, 0.65)),
        _rule(sid, '已達成', (0.85, 0.85, 0.85)),
        _rule(sid, '客人自行報名', (0.85, 0.85, 0.85)),
        _rule(sid, '關閉', (0.85, 0.85, 0.85)),
        # 欄寬
        *[{'updateDimensionProperties': {'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': i, 'endIndex': i + 1},
                                         'properties': {'pixelSize': px}, 'fields': 'pixelSize'}}
          for i, px in enumerate([70, 70, 100, 100, 60, 80, 80, 80, 110, 220, 20, 620])],
        # 說明欄樣式
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 11, 'endColumnIndex': 12},
                        'cell': {'userEnteredFormat': {'textFormat': {'bold': True}}}, 'fields': 'userEnteredFormat.textFormat'}},
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': len(NOTES), 'startColumnIndex': 11, 'endColumnIndex': 12},
                        'cell': {'userEnteredFormat': {'wrapStrategy': 'WRAP', 'verticalAlignment': 'TOP',
                                                       'backgroundColor': {'red': 1.0, 'green': 0.98, 'blue': 0.88}}},
                        'fields': 'userEnteredFormat(wrapStrategy,verticalAlignment,backgroundColor)'}},
    ]
    sh.batch_update({'requests': reqs})
    print(f"  OK 「{TAB}」已寫入 {len(ROWS)} 列（下拉選單、狀態顏色、說明都設好了）")
    return ws, True


def setup_permission(gc):
    """在使用者管理試算表的「PORTAL 頁面權限」加一欄，所有角色預設 O"""
    sh = gc.open_by_key(USER_MGMT_SPREADSHEET_ID)
    ws = sh.worksheet('PORTAL 頁面權限')
    vals = ws.get_all_values()
    header = vals[0]
    if PAGE_NAME in header:
        print(f"  權限表已有「{PAGE_NAME}」欄，略過")
        return
    col = len([h for h in header if h.strip()]) + 1          # 下一個空欄（1-based）
    if col > ws.col_count:
        ws.resize(cols=col)
    cells = [[PAGE_NAME]]
    n_roles = 0
    for r in vals[1:]:
        has_role = bool(r and r[0].strip())
        cells.append(['O' if has_role else ''])
        n_roles += 1 if has_role else 0
    ws.update(range_name=gspread.utils.rowcol_to_a1(1, col), values=cells, value_input_option='RAW')
    print(f"  OK 權限表新增「{PAGE_NAME}」欄（第 {col} 欄），{n_roles} 個角色都預設可看（O）")


def main():
    force = '--force' in sys.argv
    gc = gspread.authorize(Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES))
    print("[1/2] DBC Base「限時推薦獎金」分頁")
    setup_tab(gc, force)
    print("[2/2] Portal 頁面權限")
    setup_permission(gc)


if __name__ == '__main__':
    main()
