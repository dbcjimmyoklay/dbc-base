"""
setup_user_mgmt.py
設定「使用者管理」試算表（可重複執行，會覆蓋重建）
  - 分頁1：使用者清單
  - 分頁2：角色權限設定
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from shared.config import USER_MGMT_SPREADSHEET_ID
from shared.gsheet import get_client, batch_update
import gspread

def hex_to_rgb(h):
    h = h.lstrip('#')
    return {'red': int(h[0:2],16)/255, 'green': int(h[2:4],16)/255, 'blue': int(h[4:6],16)/255}

# ── 角色清單（依截圖）──
ROLES = ['系統創建者', '老闆', '區主管', '野澤主管', '斑尾主管', '湯澤主管', '龍平主管', 'OP']

# ── 功能欄位（依截圖欄位順序）──
FEATURES = [
    '課程安排', '房務表', '異動紀錄',
    '野澤入住單', '斑尾入住單', '龍平入住單',
    '野澤送客單', '斑尾送客單', '龍平送客單',
    '長野車單', '龍平車單',
    '教練需求報表', '銷量分析報表', '教練班表',
]

E = '可編輯'   # can edit
V = '只能檢視'  # view only
X = '✗'        # no access

# ── 權限矩陣（依截圖）──
ROLE_PERMS = {
    '系統創建者': [E, E, E,  E, E, E,  E, E, E,  E, E,  E, E, E],
    '老闆'      : [V, V, V,  V, V, V,  V, V, V,  V, V,  V, V, V],
    '區主管'    : [E, V, E,  V, V, V,  V, V, V,  V, V,  V, V, E],
    '野澤主管'  : [E, V, V,  E, V, V,  E, V, V,  V, V,  V, X, X],
    '斑尾主管'  : [V, V, V,  V, E, V,  V, E, V,  V, V,  V, X, X],
    '湯澤主管'  : [V, V, V,  V, V, V,  V, V, V,  V, V,  V, X, X],
    '龍平主管'  : [V, V, V,  V, V, E,  V, V, E,  V, E,  V, X, X],
    'OP'        : [V, V, V,  V, V, V,  V, V, V,  E, E,  V, V, V],
}

def run():
    gc = get_client()
    sh = gc.open_by_key(USER_MGMT_SPREADSHEET_ID)

    # ═══════════════════════════════════════
    # 分頁1：使用者清單
    # ═══════════════════════════════════════
    try:
        ws1 = sh.worksheet('使用者清單')
    except gspread.exceptions.WorksheetNotFound:
        try:
            ws1 = sh.worksheet('工作表1')
            ws1.update_title('使用者清單')
        except:
            ws1 = sh.add_worksheet(title='使用者清單', rows=300, cols=10)
    print("設定 分頁1：使用者清單...")

    USERS_HEADERS = ['Email', '姓名', '角色', '狀態', '申請時間', '核准時間', '核准者', '備註']
    ws1.clear()
    ws1.update([USERS_HEADERS], 'A1', value_input_option='RAW')
    sid1 = ws1.id
    reqs = []

    # 標題列
    reqs.append({'repeatCell': {
        'range': {'sheetId': sid1, 'startRowIndex': 0, 'endRowIndex': 1,
                  'startColumnIndex': 0, 'endColumnIndex': len(USERS_HEADERS)},
        'cell': {'userEnteredFormat': {
            'backgroundColor': hex_to_rgb('#1a1a2e'),
            'textFormat': {'foregroundColor': {'red':1,'green':1,'blue':1}, 'bold': True},
            'horizontalAlignment': 'CENTER',
        }},
        'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }})

    # 角色下拉
    reqs.append({'setDataValidation': {
        'range': {'sheetId': sid1, 'startRowIndex': 1, 'endRowIndex': 300,
                  'startColumnIndex': 2, 'endColumnIndex': 3},
        'rule': {
            'condition': {'type': 'ONE_OF_LIST',
                          'values': [{'userEnteredValue': r} for r in ROLES]},
            'showCustomUi': True, 'strict': True,
        }
    }})

    # 狀態下拉
    reqs.append({'setDataValidation': {
        'range': {'sheetId': sid1, 'startRowIndex': 1, 'endRowIndex': 300,
                  'startColumnIndex': 3, 'endColumnIndex': 4},
        'rule': {
            'condition': {'type': 'ONE_OF_LIST',
                          'values': [{'userEnteredValue': s} for s in ['待審核', '核准', '停用']]},
            'showCustomUi': True, 'strict': True,
        }
    }})

    # 狀態條件填色
    for status, color in [('核准','#d9ead3'), ('待審核','#fff2cc'), ('停用','#f4cccc')]:
        reqs.append({'addConditionalFormatRule': {
            'rule': {
                'ranges': [{'sheetId': sid1, 'startRowIndex': 1, 'endRowIndex': 300,
                            'startColumnIndex': 3, 'endColumnIndex': 4}],
                'booleanRule': {
                    'condition': {'type': 'TEXT_EQ', 'values': [{'userEnteredValue': status}]},
                    'format': {'backgroundColor': hex_to_rgb(color)},
                }
            }, 'index': 0,
        }})

    # 欄寬
    for ci, px in enumerate([230, 100, 100, 80, 140, 140, 80, 200]):
        reqs.append({'updateDimensionProperties': {
            'range': {'sheetId': sid1, 'dimension': 'COLUMNS', 'startIndex': ci, 'endIndex': ci+1},
            'properties': {'pixelSize': px}, 'fields': 'pixelSize',
        }})

    ws1.freeze(rows=1)
    print("  OK")

    # ═══════════════════════════════════════
    # 分頁2：角色權限設定
    # ═══════════════════════════════════════
    try:
        ws2 = sh.worksheet('角色權限設定')
        ws2.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws2 = sh.add_worksheet(title='角色權限設定', rows=20, cols=20)
    print("設定 分頁2：角色權限設定...")

    header_row = ['角色 \\ 功能'] + FEATURES
    data_rows  = [header_row]
    for role in ROLES:
        data_rows.append([role] + ROLE_PERMS[role])

    ws2.update(data_rows, 'A1', value_input_option='RAW')
    sid2 = ws2.id

    # 標題列
    reqs.append({'repeatCell': {
        'range': {'sheetId': sid2, 'startRowIndex': 0, 'endRowIndex': 1,
                  'startColumnIndex': 0, 'endColumnIndex': len(header_row)},
        'cell': {'userEnteredFormat': {
            'backgroundColor': hex_to_rgb('#1a1a2e'),
            'textFormat': {'foregroundColor': {'red':1,'green':1,'blue':1}, 'bold': True},
            'horizontalAlignment': 'CENTER',
        }},
        'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }})

    # 角色欄（A欄）
    reqs.append({'repeatCell': {
        'range': {'sheetId': sid2, 'startRowIndex': 1, 'endRowIndex': len(data_rows),
                  'startColumnIndex': 0, 'endColumnIndex': 1},
        'cell': {'userEnteredFormat': {
            'textFormat': {'bold': True},
            'backgroundColor': hex_to_rgb('#f8f9fa'),
        }},
        'fields': 'userEnteredFormat(textFormat,backgroundColor)',
    }})

    # 條件填色
    for symbol, color in [(E,'#d9ead3'), (V,'#e8f0fe'), (X,'#f4cccc')]:
        reqs.append({'addConditionalFormatRule': {
            'rule': {
                'ranges': [{'sheetId': sid2, 'startRowIndex': 1,
                            'endRowIndex': len(data_rows),
                            'startColumnIndex': 1, 'endColumnIndex': len(header_row)}],
                'booleanRule': {
                    'condition': {'type': 'TEXT_EQ', 'values': [{'userEnteredValue': symbol}]},
                    'format': {'backgroundColor': hex_to_rgb(color)},
                }
            }, 'index': 0,
        }})

    # 格線
    reqs.append({'updateBorders': {
        'range': {'sheetId': sid2, 'startRowIndex': 0, 'endRowIndex': len(data_rows),
                  'startColumnIndex': 0, 'endColumnIndex': len(header_row)},
        'top':    {'style':'SOLID','colorStyle':{'rgbColor':hex_to_rgb('#cccccc')}},
        'bottom': {'style':'SOLID','colorStyle':{'rgbColor':hex_to_rgb('#cccccc')}},
        'left':   {'style':'SOLID','colorStyle':{'rgbColor':hex_to_rgb('#cccccc')}},
        'right':  {'style':'SOLID','colorStyle':{'rgbColor':hex_to_rgb('#cccccc')}},
        'innerHorizontal': {'style':'SOLID','colorStyle':{'rgbColor':hex_to_rgb('#e0e0e0')}},
        'innerVertical':   {'style':'SOLID','colorStyle':{'rgbColor':hex_to_rgb('#e0e0e0')}},
    }})

    # 置中（權限欄）
    reqs.append({'repeatCell': {
        'range': {'sheetId': sid2, 'startRowIndex': 0, 'endRowIndex': len(data_rows),
                  'startColumnIndex': 1, 'endColumnIndex': len(header_row)},
        'cell': {'userEnteredFormat': {'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE'}},
        'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment)',
    }})

    # 欄寬
    reqs.append({'updateDimensionProperties': {
        'range': {'sheetId': sid2, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 1},
        'properties': {'pixelSize': 115}, 'fields': 'pixelSize',
    }})
    for ci in range(1, len(header_row)):
        reqs.append({'updateDimensionProperties': {
            'range': {'sheetId': sid2, 'dimension': 'COLUMNS', 'startIndex': ci, 'endIndex': ci+1},
            'properties': {'pixelSize': 78}, 'fields': 'pixelSize',
        }})

    # 列高
    for ri in range(len(data_rows)):
        reqs.append({'updateDimensionProperties': {
            'range': {'sheetId': sid2, 'dimension': 'ROWS', 'startIndex': ri, 'endIndex': ri+1},
            'properties': {'pixelSize': 28}, 'fields': 'pixelSize',
        }})

    ws2.freeze(rows=1, cols=1)

    # 批次送出
    for i in range(0, len(reqs), 500):
        sh.batch_update({'requests': reqs[i:i+500]})

    print("  OK")
    print("\n完成！")

if __name__ == '__main__':
    run()
