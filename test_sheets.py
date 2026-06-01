"""
測試 Google Sheets API 連線
"""
import gspread
from google.oauth2.service_account import Credentials
import os

folder = os.path.dirname(os.path.abspath(__file__))
creds_path = os.path.join(folder, 'dbc-credentials.json')

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

try:
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    gc = gspread.authorize(creds)
    
    # 開啟你的 Google Sheet
    SPREADSHEET_ID = '12p71jgMErzZYO4toU2LVELDPfwklHCyDARclldYImCc'
    sh = gc.open_by_key(SPREADSHEET_ID)
    
    print("連線成功！")
    print(f"試算表名稱：{sh.title}")
    print(f"分頁清單：")
    for ws in sh.worksheets():
        print(f"  - {ws.title}")
        
except Exception as e:
    print(f"連線失敗：{e}")

input("\n按Enter關閉")
