"""
setup_users.py
建立「使用者管理」分頁（一次性執行）
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from shared.gsheet import get_spreadsheet, batch_update
import gspread

SHEET_NAME = '使用者管理'

HEADERS = ['Email', '姓名', '角色', '狀態', '申請時間', '備註']

# 角色選項（之後可以增加）
ROLES = ['老闆', 'OP', '雪場主管', '教練', '其他']

def run():
    sh = get_spreadsheet()

    # 建立分頁
    try:
        ws = sh.worksheet(SHEET_NAME)
        print(f"分頁「{SHEET_NAME}」已存在")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=200, cols=10)
        print(f"建立新分頁：{SHEET_NAME}")

    sid = ws.id

    # 寫入標題列
    ws.update([HEADERS], 'A1', value_input_option='RAW')

    # 格式設定
    reqs = []

    # 標題列黑底白字
    reqs.append({
        'repeatCell': {
            'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1,
                      'startColumnIndex': 0, 'endColumnIndex': len(HEADERS)},
            'cell': {'userEnteredFormat': {
                'backgroundColor': {'red': 0, 'green': 0, 'blue': 0},
                'textFormat': {'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'bold': True},
                'horizontalAlignment': 'CENTER',
            }},
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
        }
    })

    # 欄寬
    for ci, px in enumerate([220, 100, 90, 80, 140, 200]):
        reqs.append({
            'updateDimensionProperties': {
                'range': {'sheetId': sid, 'dimension': 'COLUMNS',
                          'startIndex': ci, 'endIndex': ci + 1},
                'properties': {'pixelSize': px},
                'fields': 'pixelSize',
            }
        })

    # 「角色」欄下拉選單（C欄，index=2）
    reqs.append({
        'setDataValidation': {
            'range': {'sheetId': sid, 'startRowIndex': 1, 'endRowIndex': 200,
                      'startColumnIndex': 2, 'endColumnIndex': 3},
            'rule': {
                'condition': {
                    'type': 'ONE_OF_LIST',
                    'values': [{'userEnteredValue': r} for r in ROLES],
                },
                'showCustomUi': True,
                'strict': True,
            }
        }
    })

    # 「狀態」欄下拉選單（D欄，index=3）
    reqs.append({
        'setDataValidation': {
            'range': {'sheetId': sid, 'startRowIndex': 1, 'endRowIndex': 200,
                      'startColumnIndex': 3, 'endColumnIndex': 4},
            'rule': {
                'condition': {
                    'type': 'ONE_OF_LIST',
                    'values': [{'userEnteredValue': s} for s in ['待審核', '核准', '停用']],
                },
                'showCustomUi': True,
                'strict': True,
            }
        }
    })

    # 凍結標題列
    ws.freeze(rows=1)

    batch_update(sh, reqs)
    print("OK 使用者管理分頁設定完成")
    print()
    print("欄位說明：")
    print("  Email    → 使用者的 Google 帳號")
    print("  姓名     → 使用者填入的名字")
    print("  角色     → 你設定的：老闆 / OP / 雪場主管 / 教練 / 其他")
    print("  狀態     → 待審核 / 核准 / 停用")
    print("  申請時間 → 系統自動填入")
    print("  備註     → 你自己備注用")

if __name__ == '__main__':
    run()
