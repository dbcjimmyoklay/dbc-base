"""
重建異動紀錄表格式
執行一次，設定好所有格式、填色、Filter View、Checkbox
"""

import gspread
from google.oauth2.service_account import Credentials
import os

folder = os.path.dirname(os.path.abspath(__file__))
CREDS_PATH = os.path.join(folder, 'dbc-credentials.json')
SPREADSHEET_ID = '12p71jgMErzZYO4toU2LVELDPfwklHCyDARclldYImCc'
SHEET_NAME = '異動紀錄'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SPREADSHEET_ID)
ws = sh.worksheet(SHEET_NAME)
sid = ws.id

HEADER = ['異動日期', '異動類型', '項目', '訂編', '中文姓名', '異動欄位', '原值', '新值', '處理']

# ===== 讀取現有資料 =====
print("讀取現有異動紀錄...")
existing = ws.get_all_values()

# 重建標題列（如果欄位結構不對）
if not existing or existing[0] != HEADER:
    print("更新標題列...")
    # 保留舊資料，只更新標題
    if existing:
        old_data = existing[1:]  # 舊資料
    else:
        old_data = []
    ws.clear()
    ws.update([HEADER] + old_data, value_input_option='RAW')
    existing = ws.get_all_values()

total_rows = len(existing)
print(f"共 {total_rows} 列資料")

# ===== 欄位索引 =====
idx_type = HEADER.index('異動類型')
idx_done = HEADER.index('處理')

# ===== 顏色定義 =====
BLACK     = {'red': 0,    'green': 0,    'blue': 0}
WHITE     = {'red': 1,    'green': 1,    'blue': 1}
WHITE_TXT = {'red': 1,    'green': 1,    'blue': 1}
GRAY_LIGHT= {'red': 0.85, 'green': 0.85, 'blue': 0.85}
GRAY_TXT  = {'red': 0.6,  'green': 0.6,  'blue': 0.6}
COLOR_NEW = {'red': 1.0,  'green': 0.95, 'blue': 0.4}   # 亮黃（新增）
COLOR_DEL = {'red': 0.95, 'green': 0.6,  'blue': 0.6}   # 紅（刪減）
COLOR_CHG = {'red': 1.0,  'green': 1.0,  'blue': 1.0}   # 白（資料異動）
COLOR_DONE= {'red': 0.9,  'green': 0.9,  'blue': 0.9}   # 灰（已處理）

reqs = []

# ===== 1. 標題列格式 =====
reqs.append({'repeatCell': {
    'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1,
              'startColumnIndex': 0, 'endColumnIndex': len(HEADER)},
    'cell': {'userEnteredFormat': {
        'backgroundColor': BLACK,
        'textFormat': {'foregroundColor': WHITE_TXT, 'bold': True, 'fontSize': 11},
        'horizontalAlignment': 'CENTER',
        'verticalAlignment': 'MIDDLE'
    }},
    'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
}})

# ===== 2. 欄寬設定 =====
COL_WIDTHS = [65, 65, 65, 55, 90, 90, 90, 90, 65]
for ci, w in enumerate(COL_WIDTHS):
    reqs.append({'updateDimensionProperties': {
        'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': ci, 'endIndex': ci+1},
        'properties': {'pixelSize': w}, 'fields': 'pixelSize'
    }})

# 列高
reqs.append({'updateDimensionProperties': {
    'range': {'sheetId': sid, 'dimension': 'ROWS', 'startIndex': 0, 'endIndex': 1},
    'properties': {'pixelSize': 32}, 'fields': 'pixelSize'
}})
if total_rows > 1:
    reqs.append({'updateDimensionProperties': {
        'range': {'sheetId': sid, 'dimension': 'ROWS', 'startIndex': 1, 'endIndex': total_rows},
        'properties': {'pixelSize': 26}, 'fields': 'pixelSize'
    }})

# ===== 3. 各列填色（依異動類型和處理狀態）=====
for i, row in enumerate(existing[1:], 1):
    if not row or len(row) < len(HEADER): continue

    change_type = row[idx_type].strip() if idx_type < len(row) else ''
    is_done = str(row[idx_done]).strip().upper() in ('TRUE', '1', 'YES')

    if is_done:
        bg = COLOR_DONE
        txt_color = GRAY_TXT
        bold = False
    elif change_type == '新增':
        bg = COLOR_NEW
        txt_color = BLACK
        bold = True
    elif change_type == '刪減':
        bg = COLOR_DEL
        txt_color = BLACK
        bold = True
    else:
        bg = COLOR_CHG
        txt_color = BLACK
        bold = False

    reqs.append({'repeatCell': {
        'range': {'sheetId': sid, 'startRowIndex': i, 'endRowIndex': i+1,
                  'startColumnIndex': 0, 'endColumnIndex': len(HEADER)},
        'cell': {'userEnteredFormat': {
            'backgroundColor': bg,
            'textFormat': {'foregroundColor': txt_color, 'bold': bold},
            'verticalAlignment': 'MIDDLE'
        }},
        'fields': 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)'
    }})

# ===== 4. 處理欄設為 Checkbox =====
if total_rows > 1:
    reqs.append({'repeatCell': {
        'range': {'sheetId': sid, 'startRowIndex': 1, 'endRowIndex': max(total_rows, 500),
                  'startColumnIndex': idx_done, 'endColumnIndex': idx_done+1},
        'cell': {
            'dataValidation': {
                'condition': {'type': 'BOOLEAN'},
                'strict': True
            }
        },
        'fields': 'dataValidation'
    }})

# ===== 5. 格線 =====
if total_rows > 1:
    border_solid = {'style': 'SOLID', 'colorStyle': {'rgbColor': {'red': 0.7, 'green': 0.7, 'blue': 0.7}}}
    border_black = {'style': 'SOLID', 'colorStyle': {'rgbColor': BLACK}}
    reqs.append({'updateBorders': {
        'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': total_rows,
                  'startColumnIndex': 0, 'endColumnIndex': len(HEADER)},
        'top': border_black, 'bottom': border_black,
        'left': border_black, 'right': border_black,
        'innerHorizontal': border_solid,
        'innerVertical': border_solid,
    }})

# ===== 6. 凍結標題列 =====
ws.freeze(rows=1)

# ===== 7. 送出所有格式請求 =====
print("套用格式...")
for i in range(0, len(reqs), 500):
    sh.batch_update({'requests': reqs[i:i+500]})

# ===== 8. Filter View：待處理 / 全部紀錄 =====
print("建立 Filter View...")
meta = sh.fetch_sheet_metadata()
existing_fv = []
for s in meta.get('sheets', []):
    if s['properties']['sheetId'] == sid:
        existing_fv = s.get('filterViews', [])
        break

if existing_fv:
    sh.batch_update({'requests': [
        {'deleteFilterView': {'filterId': fv['filterViewId']}} for fv in existing_fv
    ]})

sh.batch_update({'requests': [
    # 待處理：處理欄是 FALSE（未勾）
    {'addFilterView': {'filter': {
        'title': '⚠ 待處理',
        'range': {'sheetId': sid, 'startRowIndex': 0, 'startColumnIndex': 0},
        'filterSpecs': [{
            'columnIndex': idx_done,
            'filterCriteria': {
                'condition': {'type': 'TEXT_EQ', 'values': [{'userEnteredValue': 'FALSE'}]}
            }
        }]
    }}},
    # 只看新增
    {'addFilterView': {'filter': {
        'title': '🟡 新增',
        'range': {'sheetId': sid, 'startRowIndex': 0, 'startColumnIndex': 0},
        'filterSpecs': [{
            'columnIndex': idx_type,
            'filterCriteria': {
                'condition': {'type': 'TEXT_EQ', 'values': [{'userEnteredValue': '新增'}]}
            }
        }]
    }}},
    # 只看刪減
    {'addFilterView': {'filter': {
        'title': '🔴 刪減',
        'range': {'sheetId': sid, 'startRowIndex': 0, 'startColumnIndex': 0},
        'filterSpecs': [{
            'columnIndex': idx_type,
            'filterCriteria': {
                'condition': {'type': 'TEXT_EQ', 'values': [{'userEnteredValue': '刪減'}]}
            }
        }]
    }}},
    # 全部紀錄
    {'addFilterView': {'filter': {
        'title': '📋 全部紀錄',
        'range': {'sheetId': sid, 'startRowIndex': 0, 'startColumnIndex': 0},
    }}}
]})

print("\n完成！異動紀錄表已升級：")
print("  🔴 刪減 → 紅色，最醒目")
print("  🟡 新增 → 黃色")
print("  ⬜ 資料異動 → 白底")
print("  ✅ 已處理（打勾）→ 灰色淡化")
print()
print("Filter View（資料→篩選器視圖）：")
print("  ⚠ 待處理  |  🟡 新增  |  🔴 刪減  |  📋 全部紀錄")

input("\n按Enter關閉")
