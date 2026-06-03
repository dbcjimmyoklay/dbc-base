"""
shared/config.py
DBC 影子中台 — 所有設定、常數、欄位定義集中在此
"""

import os

# ===== 路徑 =====
BASE_DIR   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CREDS_PATH = os.path.join(BASE_DIR, 'dbc-credentials.json')

# ===== PORTAL =====
APPS_SCRIPT_URL  = 'https://script.google.com/macros/s/AKfycbw2-7v48isKdkrJyZNCJEI1UEFfbL3lupGZwUPO4vHyt9B3AjkxKaV3ZHQlvz_LenOI/exec'
PORTAL_URL       = 'https://dbcjimmyoklay.github.io/dbc-base'
GOOGLE_CLIENT_ID = '278003648368-dd08o1b24knf3tlvhfp7rustnq3cs2gc.apps.googleusercontent.com'

# ===== Google Drive =====
# 你放置「原始大總表*.xlsx」的 Drive 資料夾 ID
# 開啟資料夾網址後，最後一段路徑就是 ID
# 例：https://drive.google.com/drive/folders/1ABC...XYZ → ID = 1ABC...XYZ
DRIVE_XLSX_FOLDER_ID = '1futMHPcOQ4-HeLbKxNoBRarSPycKbZDG'

# ===== Google Sheets — DBC Base（主表）=====
SPREADSHEET_ID = '12p71jgMErzZYO4toU2LVELDPfwklHCyDARclldYImCc'

SHEET_COURSE  = '課程安排26/27'
SHEET_ROOM    = '房務表26/27'
SHEET_CHANGE  = '異動紀錄'

# 入住單（gid 已在 CHECKIN_SHEETS 定義）
# 送客單
CHECKOUT_SHEETS = [
    {'name': '野澤送客單26/27', 'include': lambda item: '野澤' in item and '純課' not in item},
    {'name': '斑尾送客單26/27', 'include': lambda item: '斑尾' in item and '純課' not in item},
    {'name': '龍平送客單26/27', 'include': lambda item: '龍平' in item},
]

# 車單
SHEET_NAGANO_BUS = '長野車單'
SHEET_RYONGPYONG_BUS = '龍平車單'

# ===== Google Sheets — 野澤當地訂房（獨立）=====
NOZAWA_BOOKING_SPREADSHEET_ID = '1eNXXQgCEqHxd2VmzbGe3Z-Gac2PNBgQlSAKx19uqQ6s'
NOZAWA_BOOKING_SHEET          = '26-27'
NOZAWA_BOOKING_GID            = 2053936646
NOZAWA_COMPARE_SHEET          = '比對報告'

# ===== Google Sheets — 教練班表（獨立）=====
COACH_SCHEDULE_SPREADSHEET_ID = '1UsOty8fY5FZ5duOYLscMDYFH_lsKTfDhmuk_n3VU9B8'
COACH_SCHEDULE_SHEET          = '26/27 教練班表'   # 當季分頁名稱

# ===== Google Sheets — 使用者管理（獨立）=====
USER_MGMT_SPREADSHEET_ID = '1b2YR6VpDT_KYOExWLw0Gsy9D0Ise-kBcKtVDW8zrfPM'
USER_MGMT_SHEET          = '工作表1'   # 之後可改名

CHECKIN_SHEETS = [
    {
        'name'   : '野澤入住單26/27',
        'gid'    : 2070775881,
        'include': lambda item: '野澤' in item and '純課' not in item,
    },
    {
        'name'   : '斑尾入住單26/27',
        'gid'    : 460046160,
        'include': lambda item: '斑尾' in item and '純課' not in item,
    },
    {
        'name'   : '龍平入住單26/27',
        'gid'    : 343688459,
        'include': lambda item: '龍平' in item,
    },
]

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# ===== 欄位定義 =====

# 自動欄（科威來的，每次覆蓋）
AUTO_COLS = {
    '項目', '出發日期', '泊數', '天數', '註記', '訂編',
    '中文姓名', '英文姓名', '性別', '年齡',
    '尺寸', '身高', '體重', '腳長', '雪板',
    'LEVEL', '類別', '衣褲', '護膝', '護臀', '手套', '4件組',
}

# 手動欄（人員填的，永不覆蓋）
MANUAL_COLS = {'教練', '助教', 'checking', '備註', '配房序'}

# 異動要通知的重要欄位
IMPORTANT_COLS = {
    'Lv', '出發日期', '天數', '類別', '衣褲', '護膝', '護臀', '手套', '4件組', '項目', '泊數',
}

# 課程安排欄位順序
COURSE_COLS = [
    '編號', '項目', '出發日期', '泊數', '天數', '註記', '訂編',
    '中文姓名', '英文姓名', '性別', '年齡',
    '尺寸', '身高', '體重', '腳長', '雪板',
    'LEVEL', '類別', '衣褲', '護膝', '護臀', '手套', '4件組',
    '教練', '助教', 'checking', '備註',
    '旅客編號', '項目1', '配房序',  # 隱藏欄位
]

# 課程安排表頭顯示名稱（LEVEL→Lv，4件組→四件）
COURSE_HEADERS = [
    '編號', '項目', '出發日期', '泊數', '天數', '註記', '訂編',
    '中文姓名', '英文姓名', '性別', '年齡',
    '尺寸', '身高', '體重', '腳長', '雪板',
    'Lv', '類別', '衣褲', '護膝', '護臀', '手套', '四件',
    '教練', '助教', 'checking', '備註',
    '旅客編號', '項目1', '配房序',
]

# 隱藏欄位
HIDDEN_COLS = {'旅客編號', '項目1', '配房序'}

# 房務表
ROOM_COLS = [
    '項目', '出發日期', '泊數', '訂編',
    '中文姓名', '英文姓名', '性別', '年齡',
    '房號', '序號', '備註', '飲食', 'OP備註',
]
ROOM_HEADERS     = ROOM_COLS
ROOM_MANUAL_COLS = {'房號', '序號', '備註'}
ROOM_IMPORTANT_COLS = {'出發日期', '泊數', '項目', '飲食'}
ROOM_CENTER_COLS = {
    '項目', '出發日期', '泊數', '訂編',
    '中文姓名', '英文姓名', '性別', '年齡',
    '房號', '序號',
}
ROOM_EXCLUDE_ITEMS = {'湯澤純課'}

# 入住單
CHECKIN_COLS = [
    '項目', '出發日期', '泊數', '天數', '訂編',
    '中文姓名', '英文姓名', '性別', '年齡',
    '房號', '房型', '序號', '分房備註', '入住備註', '飲食', 'OP備註',
]
CHECKIN_HEADERS     = CHECKIN_COLS
CHECKIN_MANUAL_COLS = {'房號', '房型', '序號', '入住備註'}  # 分房備註/飲食/OP備註 自動帶入
CHECKIN_IMPORTANT_COLS = {'出發日期', '泊數', '天數', '項目'}
CHECKIN_CENTER_COLS = {
    '項目', '出發日期', '泊數', '訂編',
    '中文姓名', '英文姓名', '性別', '年齡',
    '房號', '序號',
}

# 異動紀錄表頭
CHANGE_HEADER = [
    '異動日期', '異動類型', '項目', '出發日期', '訂編', '中文姓名',
    '異動欄位', '原值', '新值', '處理',
]

# ===== 雪場對應規則 =====
ITEM_RULES = [
    ('ZN', '野澤純課'), ('ZB', '斑尾純課'), ('ZY', '湯澤純課'),
    ('YS', '龍平'), ('NT', '野澤'), ('BT', '斑尾'),
    ('N',  '野澤'), ('B',  '斑尾'),
]
PURE_ITEMS  = {'野澤純課', '斑尾純課', '湯澤純課'}
RESORT_NAME = {'野澤純課': '野澤', '斑尾純課': '斑尾', '湯澤純課': '湯澤'}
ITEM_ORDER  = {'野澤': 1, '野澤純課': 2, '斑尾': 3, '斑尾純課': 4, '湯澤純課': 5, '龍平': 6}

# ===== 價目表 =====
PRICE_TABLE = {
    '野澤純課': {1: 16500, 2: 17500, 3: 19500, 4: 19500, 5: 21000, 6: 21000},
    '斑尾純課': {1: 14500, 2: 15000, 3: 16000, 4: 17000, 5: 18000, 6: 19000},
    '湯澤純課': {1: 12500, 2: 13000, 3: 14000, 4: 14000, 5: 16000, 6: 16000},
}
YUKIGUNI_PEAK = [
    (1, 12, 1, 15, 1000), (1, 21, 2, 10, 1000), (2, 25, 2, 25, 1000),
    (2, 26, 3,  2, 2000), (3,  3, 3,  7, 1000),
]
EQUIP_PRICE = {'衣褲': 1000, '護膝': 200, '護臀': 200, '4件組': 1800, '手套': 400}

# ===== 曜日對照 =====
WEEKDAY_MAP = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}

# ===== 填色定義 =====
COLOR = {
    'yellow_light' : {'red': 1.0,        'green': 0.851,      'blue': 0.4},    # lv1/lv2 #ffd966
    'yellow_bright': {'red': 1.0,        'green': 1.0,        'blue': 0.0},    # 專屬/含  #ffff00
    'orange'       : {'red': 1.0,        'green': 0.6,        'blue': 0.2},    # lv3/lv4
    'teal'         : {'red': 0.2,        'green': 0.8,        'blue': 0.8},    # 雙板
    'white'        : {'red': 1.0,        'green': 1.0,        'blue': 1.0},
    'separator'    : {'red': 0.2,        'green': 0.2,        'blue': 0.2},    # 分隔列黑底
    'header'       : {'red': 0.0,        'green': 0.0,        'blue': 0.0},    # 標題黑底
    'row_blue'     : {'red': 0.8,        'green': 0.9,        'blue': 1.0},    # 斑尾 淡藍
    'row_orange'   : {'red': 1.0,        'green': 0.9,        'blue': 0.8},    # 湯澤 淡橘
    'row_brown'    : {'red': 0xcd / 255, 'green': 0xa5 / 255, 'blue': 0x81 / 255},  # 龍平 #cda581
    'change_new'   : {'red': 1.0,        'green': 0.95,       'blue': 0.6},    # 新增 亮黃
    'change_del'   : {'red': 1.0,        'green': 0.7,        'blue': 0.7},    # 刪減 紅
}

# ===== 欄寬定義 =====
COL_WIDTH = {
    '編號': 30, '泊數': 30, '天數': 30, '性別': 30, '年齡': 30,
    '尺寸': 30, '身高': 30, '體重': 30, '腳長': 30, '雪板': 30,
    'Lv': 30, 'LEVEL': 30, '類別': 30,
    '衣褲': 30, '護膝': 30, '護臀': 30, '手套': 30, '四件': 30, '4件組': 30,
    '註記': 35, '訂編': 35,
    '項目': 60, '出發日期': 60, '中文姓名': 60,
    '英文姓名': 150,
    '教練': 70, '助教': 70, 'checking': 70,
    '備註': 70,
    '房號': 40, '序號': 40,
    '飲食': 80,
    'OP備註': 120,
}
ROOM_COL_WIDTH = {
    '項目': 40, '泊數': 40, '訂編': 40, '性別': 40, '年齡': 40,
    '房號': 40, '序號': 40,
    '出發日期': 70, '中文姓名': 70,
    '英文姓名': 130,
    '備註': 130,
    '飲食': 80,
    'OP備註': 120,
}
CHECKIN_COL_WIDTH = {
    '項目': 70, '出發日期': 65, '泊數': 40,
    '訂編': 45, '中文姓名': 70, '英文姓名': 130,
    '性別': 35, '年齡': 35,
    '房號': 45, '房型': 80, '序號': 40,
    '分房備註': 130, '入住備註': 130, '飲食': 80, 'OP備註': 130,
}
CHANGE_COL_WIDTHS = [65, 65, 65, 70, 55, 80, 80, 80, 80, 65]

# ===== 置中欄位 =====
CENTER_COLS = {
    '編號', '出發日期', '曜日', '泊數', '天數', '註記', '訂編',
    '中文姓名', '性別', '年齡',
    '尺寸', '身高', '體重', '腳長', '雪板', 'LEVEL', '類別',
    '衣褲', '護膝', '護臀', '手套', '專屬', '4件組',
    '教練', '助教', 'checking', '備註',
}

# ===== 教練需求設定 =====
COACH_CAPS = {'野澤': 14, '斑尾': 8, '湯澤': 10, '龍平': 10}
COACH_THRESHOLDS = {
    '野澤': [(6, 'green'), (9, 'yellow'), (11, 'orange'), (999, 'red')],
    '斑尾': [(4, 'green'), (6, 'yellow'), (7,  'orange'), (999, 'red')],
    '湯澤': [(5, 'green'), (7, 'yellow'), (9,  'orange'), (999, 'red')],
    '龍平': [(5, 'green'), (7, 'yellow'), (9,  'orange'), (999, 'red')],
}
