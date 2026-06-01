"""
建立房務表的 Filter View
執行一次即可，建立三個篩選視圖：野澤 / 斑尾 / 龍平
"""

import gspread
from google.oauth2.service_account import Credentials
import os

folder = os.path.dirname(os.path.abspath(__file__))
CREDS_PATH = os.path.join(folder, 'dbc-credentials.json')
SPREADSHEET_ID = '12p71jgMErzZYO4toU2LVELDPfwklHCyDARclldYImCc'
SHEET_NAME = '房務表26/27'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SPREADSHEET_ID)
ws = sh.worksheet(SHEET_NAME)
sheet_id = ws.id

# 找「項目」欄的位置
headers = ws.row_values(1)
try:
    item_col_idx = headers.index('項目')
    print(f"項目欄位置：第 {item_col_idx + 1} 欄（index {item_col_idx}）")
except ValueError:
    print("找不到項目欄位！")
    exit(1)

# 三個篩選視圖：用 TEXT_CONTAINS 篩選「野澤」「斑尾」「龍平」
FILTER_VIEWS = [
    {'name': '野澤', 'keyword': '野澤'},
    {'name': '斑尾', 'keyword': '斑尾'},
    {'name': '龍平', 'keyword': '龍平'},
]

# 先刪除現有的 Filter View
print("檢查現有 Filter View...")
spreadsheet = sh.fetch_sheet_metadata()
existing_filters = []
for s in spreadsheet.get('sheets', []):
    if s['properties']['sheetId'] == sheet_id:
        existing_filters = s.get('filterViews', [])
        break

if existing_filters:
    print(f"  刪除現有 {len(existing_filters)} 個 Filter View...")
    delete_reqs = [
        {'deleteFilterView': {'filterId': fv['filterViewId']}}
        for fv in existing_filters
    ]
    sh.batch_update({'requests': delete_reqs})

# 建立新的 Filter View
print("建立 Filter View...")
requests = []
for fv in FILTER_VIEWS:
    req = {
        'addFilterView': {
            'filter': {
                'title': fv['name'],
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 0,
                    'startColumnIndex': 0,
                },
                'filterSpecs': [
                    {
                        'columnIndex': item_col_idx,
                        'filterCriteria': {
                            'condition': {
                                'type': 'TEXT_CONTAINS',
                                'values': [{'userEnteredValue': fv['keyword']}]
                            }
                        }
                    }
                ]
            }
        }
    }
    requests.append(req)

sh.batch_update({'requests': requests})
print("完成！")
print()
print("使用方式：")
print("  Google Sheet 上方 → 資料 → 篩選器視圖 → 選擇 野澤/斑尾/龍平")

input("\n按Enter關閉")
