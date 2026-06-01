"""
shared/gsheet.py
DBC 影子中台 — Google Sheets 共用操作函數
"""

import gspread
from google.oauth2.service_account import Credentials
from .config import CREDS_PATH, SCOPES, SPREADSHEET_ID


# ===== 連線 =====

def get_client():
    """建立並回傳 gspread client"""
    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    return gspread.authorize(creds)


def get_spreadsheet():
    """取得 DBC Base 試算表物件"""
    return get_client().open_by_key(SPREADSHEET_ID)


def get_or_create_sheet(sh, title, rows=2000, cols=30):
    """取得分頁；不存在時自動建立"""
    try:
        return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
        print(f"  建立新分頁：{title}")
        return ws


def get_sheet_by_gid(sh, gid):
    """用 gid 找分頁（避免改名後重複建立）"""
    for ws in sh.worksheets():
        if ws.id == gid:
            return ws
    return None


# ===== 清除 =====

def clear_rows_from(sh, ws, start_row_index, n_cols):
    """
    清除指定列以後的內容與格式
    start_row_index：0-based（例：從第3列清 → start_row_index=2）
    """
    sh.batch_update({'requests': [{
        'updateCells': {
            'range': {
                'sheetId'        : ws.id,
                'startRowIndex'  : start_row_index,
                'endRowIndex'    : 2000,
                'startColumnIndex': 0,
                'endColumnIndex' : n_cols,
            },
            'fields': 'userEnteredValue,userEnteredFormat',
        }
    }]})


# ===== 篩選器 =====

def rebuild_filter(sh, ws, total_data_rows, n_cols, start_row=1):
    """
    重建分頁篩選器
    start_row=1 → 篩選從第2列開始（課程安排/房務表/異動紀錄）
    start_row=0 → 篩選從第1列開始（入住單）
    """
    sid      = ws.id
    last_row = total_data_rows + start_row + 1
    sh.batch_update({'requests': [{'clearBasicFilter': {'sheetId': sid}}]})
    sh.batch_update({'requests': [{
        'setBasicFilter': {
            'filter': {
                'range': {
                    'sheetId'         : sid,
                    'startRowIndex'   : start_row,
                    'endRowIndex'     : last_row,
                    'startColumnIndex': 0,
                    'endColumnIndex'  : n_cols,
                }
            }
        }
    }]})
    print(f"  OK 篩選器重建（第{start_row+1}列～第{last_row}列）")


# ===== 格式請求產生器 =====

def req_cell_color(sheet_id, row_idx, col_idx, color, n_cols=1):
    """單格（或多格）填色請求；row_idx/col_idx 皆 1-based"""
    return {
        'repeatCell': {
            'range': {
                'sheetId'         : sheet_id,
                'startRowIndex'   : row_idx - 1,
                'endRowIndex'     : row_idx,
                'startColumnIndex': col_idx - 1,
                'endColumnIndex'  : col_idx - 1 + n_cols,
            },
            'cell'  : {'userEnteredFormat': {'backgroundColor': color}},
            'fields': 'userEnteredFormat.backgroundColor',
        }
    }


def req_row_color(sheet_id, row_idx, n_cols, color):
    """整列填色請求；row_idx 1-based"""
    return req_cell_color(sheet_id, row_idx, 1, color, n_cols=n_cols)


def req_header_format(sheet_id, row_idx, n_cols):
    """標題列：黑底白字粗體置中"""
    return {
        'repeatCell': {
            'range': {
                'sheetId'         : sheet_id,
                'startRowIndex'   : row_idx - 1,
                'endRowIndex'     : row_idx,
                'startColumnIndex': 0,
                'endColumnIndex'  : n_cols,
            },
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor' : {'red': 0, 'green': 0, 'blue': 0},
                    'textFormat'      : {
                        'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
                        'bold'           : True,
                    },
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment'  : 'MIDDLE',
                }
            },
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
        }
    }


def req_separator_format(sheet_id, row_idx, n_cols):
    """分隔列：深灰底白字粗體"""
    return {
        'repeatCell': {
            'range': {
                'sheetId'         : sheet_id,
                'startRowIndex'   : row_idx - 1,
                'endRowIndex'     : row_idx,
                'startColumnIndex': 0,
                'endColumnIndex'  : n_cols,
            },
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2},
                    'textFormat'     : {
                        'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
                        'bold'           : True,
                    },
                }
            },
            'fields': 'userEnteredFormat(backgroundColor,textFormat)',
        }
    }


def req_center_align(sheet_id, start_row, end_row, col_idx):
    """欄位置中"""
    return {
        'repeatCell': {
            'range': {
                'sheetId'         : sheet_id,
                'startRowIndex'   : start_row - 1,
                'endRowIndex'     : end_row,
                'startColumnIndex': col_idx,
                'endColumnIndex'  : col_idx + 1,
            },
            'cell': {
                'userEnteredFormat': {
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment'  : 'MIDDLE',
                }
            },
            'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment)',
        }
    }


def req_borders(sheet_id, start_row, end_row, start_col, end_col):
    """黑色格線"""
    solid = {'style': 'SOLID', 'colorStyle': {'rgbColor': {'red': 0, 'green': 0, 'blue': 0}}}
    gray  = {'style': 'SOLID', 'colorStyle': {'rgbColor': {'red': 0.8, 'green': 0.8, 'blue': 0.8}}}
    return {
        'updateBorders': {
            'range': {
                'sheetId'         : sheet_id,
                'startRowIndex'   : start_row - 1,
                'endRowIndex'     : end_row,
                'startColumnIndex': start_col,
                'endColumnIndex'  : end_col,
            },
            'top'            : solid,
            'bottom'         : solid,
            'left'           : solid,
            'right'          : solid,
            'innerHorizontal': gray,
            'innerVertical'  : gray,
        }
    }


def req_col_width(sheet_id, col_idx, pixel_size):
    """欄寬設定；col_idx 0-based"""
    return {
        'updateDimensionProperties': {
            'range'     : {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                           'startIndex': col_idx, 'endIndex': col_idx + 1},
            'properties': {'pixelSize': pixel_size},
            'fields'    : 'pixelSize',
        }
    }


def req_row_height(sheet_id, start_idx, end_idx, pixel_size):
    """列高設定；index 0-based"""
    return {
        'updateDimensionProperties': {
            'range'     : {'sheetId': sheet_id, 'dimension': 'ROWS',
                           'startIndex': start_idx, 'endIndex': end_idx},
            'properties': {'pixelSize': pixel_size},
            'fields'    : 'pixelSize',
        }
    }


def req_hide_col(sheet_id, col_idx):
    """隱藏欄位；col_idx 0-based"""
    return {
        'updateDimensionProperties': {
            'range'     : {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                           'startIndex': col_idx, 'endIndex': col_idx + 1},
            'properties': {'hiddenByUser': True},
            'fields'    : 'hiddenByUser',
        }
    }


def req_checkbox(sheet_id, start_row, end_row, col_idx):
    """設定 Checkbox 資料驗證；index 0-based"""
    return {
        'repeatCell': {
            'range': {
                'sheetId'         : sheet_id,
                'startRowIndex'   : start_row,
                'endRowIndex'     : end_row,
                'startColumnIndex': col_idx,
                'endColumnIndex'  : col_idx + 1,
            },
            'cell': {
                'dataValidation': {
                    'condition'   : {'type': 'BOOLEAN'},
                    'strict'      : True,
                    'showCustomUi': True,
                }
            },
            'fields': 'dataValidation',
        }
    }


# ===== 批次送出 =====

def batch_update(sh, requests, chunk=500):
    """分批送出格式請求（每批最多 500 個）"""
    for i in range(0, len(requests), chunk):
        sh.batch_update({'requests': requests[i:i + chunk]})
