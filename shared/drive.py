"""
shared/drive.py
DBC 影子中台 — Google Drive 下載工具
─────────────────────────────────────────────
用途：從指定 Drive 資料夾下載最新的「原始大總表*.xlsx」
      供 GitHub Actions 在雲端執行時取得原始檔案
─────────────────────────────────────────────
"""

import os
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials
from .config import CREDS_PATH, SCOPES, DRIVE_XLSX_FOLDER_ID, BASE_DIR


def download_latest_xlsx(dest_dir=None):
    """
    從 DRIVE_XLSX_FOLDER_ID 資料夾下載最新的「原始大總表*.xlsx」
    回傳下載後的本機路徑
    """
    if not DRIVE_XLSX_FOLDER_ID:
        raise ValueError(
            "請在 shared/config.py 填入 DRIVE_XLSX_FOLDER_ID\n"
            "（Google Drive 資料夾 ID，用來放置原始大總表 xlsx）"
        )

    dest_dir = dest_dir or BASE_DIR
    creds    = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    service  = build('drive', 'v3', credentials=creds)

    # 列出資料夾內所有 xlsx 檔，依修改時間排序
    query  = (
        f"'{DRIVE_XLSX_FOLDER_ID}' in parents"
        " and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
        " and trashed=false"
        " and name contains '原始大總表'"
    )
    result = service.files().list(
        q=query,
        orderBy='modifiedTime desc',
        pageSize=1,
        fields='files(id, name, modifiedTime)',
    ).execute()

    files = result.get('files', [])
    if not files:
        raise FileNotFoundError(
            f"Drive 資料夾（{DRIVE_XLSX_FOLDER_ID}）中找不到「原始大總表*.xlsx」\n"
            "請先上傳最新的原始大總表至該資料夾"
        )

    file_info = files[0]
    file_id   = file_info['id']
    file_name = file_info['name']
    modified  = file_info.get('modifiedTime', '')
    print(f"  找到檔案：{file_name}（最後修改：{modified[:10]}）")

    # 下載到本機
    dest_path = os.path.join(dest_dir, file_name)
    request   = service.files().get_media(fileId=file_id)
    fh        = io.FileIO(dest_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            print(f"  下載中... {int(status.progress() * 100)}%")
    fh.close()

    print(f"  OK 下載完成：{file_name}")
    return dest_path
