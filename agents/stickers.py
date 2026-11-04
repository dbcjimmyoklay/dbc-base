"""
agents/stickers.py
STICKER 旅客貼紙員
─────────────────────────────────────────────
職責：
  每日流程讀完原始大總表後，把「野澤 / 斑尾 / 龍平」團客的貼紙名單寫進一份「獨立」的
  貼紙試算表（shared/config.py 的 STICKER_SPREADSHEET_ID），完全不動 DBC Base。

貼紙內容（70 x 29.7 mm、A4 每頁 3 欄 x 10 排）：
  新版（龍平每張、野澤/斑尾 每人的第 1 張）：
      第 1 行  出發日期  中文姓名
      第 2 行  英文姓名
      下方一排 5 個 55x55 的正方形空框：尺寸 / 衣褲 / 膝 / 臀 / 手套（手寫）
  舊版（野澤/斑尾 每人的第 2、3 張）：
      第 1 行  出發日期  #訂編
      第 2 行  中文姓名
      第 3 行  英文姓名
  張數：野澤 / 斑尾 每人 3 張（同一排橫向 3 張），龍平 每人 1 張（一排 3 個人）

貼紙試算表的分頁：
  總覽                       待印人數 / 張數 / 預估 A4 頁數 ＋「已全部印完」勾選格
  [列印用] 野澤 / 斑尾 / 龍平  各區一個分頁，只放 待印 / 需重印 的人，格式與舊 Apps Script 完全相同
                              （野澤、斑尾：1 人 1 列 x 3 張；龍平：1 列 3 人、每人 1 張）
  名單                       所有旅客與狀態（待印 / 需重印 / 待補資料 / 已印 / 已取消）

狀態規則：
  待印      新增、還沒印過
  需重印    已印過，但出發日期 / 姓名 / 區域後來有異動（舊貼紙作廢）
  待補資料  中文或英文姓名空白、或還是「旅客N」佔位名 → 不進列印分頁，補齊後自動變待印
  已印      已確認印過（資料沒變）
  已取消    新的原始大總表已經沒有這個人（保留紀錄，不印）
  也可以在「名單」的狀態欄手動輸入：  已印（標記為印過）／ 重印（強制重印）

標記已印：印完後到「總覽」勾選「該區」的格子（各區獨立），下一次跑每日流程就會把該區列印分頁上的人標記為已印。

⚠ 列印格式（寬 293px、高 124px、Arial 11、置中、自動換行）是實測過的數值，勿動。
"""

import os
import re
import sys
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from shared.config import CREDS_PATH, SCOPES, STICKER_SPREADSHEET_ID   # noqa: E402

# ═══════════════════════════════════════════
# 列印版面（實測過，勿動）
# ═══════════════════════════════════════════
COL_WIDTH   = 293      # 每欄寬 px
ROW_HEIGHT  = 124      # 每列高 px（A4 恰好 10 列）
COLS_PER_ROW = 3       # 每列 3 張貼紙
FONT_NAME   = 'Arial'
FONT_SIZE   = 11
ROWS_PER_PAGE = 10

COPIES = {'野澤': 3, '斑尾': 3, '龍平': 1}     # 每人幾張
AREA_ORDER = {'野澤': 0, '斑尾': 1, '龍平': 2}

TAB_SUMMARY = '總覽'
TAB_MASTER  = '名單'
TAB_PRINT   = {a: f'[列印用] {a} 70x29.7mm貼紙' for a in COPIES}      # 每區各一個列印分頁
TAB_PRINT_OLD = ['[列印用] 野澤斑尾 70x29.7mm貼紙']                  # 舊版合併分頁，會自動刪除
CHECK_CELLS = {'野澤': 'B4', '斑尾': 'B5', '龍平': 'B6'}              # 總覽的「已印完」勾選格（各區獨立）

ST_PENDING, ST_REPRINT, ST_MISSING, ST_PRINTED, ST_CANCEL = '待印', '需重印', '待補資料', '已印', '已取消'

HEADERS = ['狀態', '區域', '出發日期', '訂編', '中文姓名', '英文姓名', '張數',
           '新增日期', '列印日期', '備註', 'key', '印單版本', '已印版本']
H = {name: i for i, name in enumerate(HEADERS)}

COLORS = {
    ST_PENDING: (1.00, 0.95, 0.70),   # 淡黃
    ST_REPRINT: (1.00, 0.80, 0.60),   # 橘
    ST_MISSING: (0.96, 0.80, 0.80),   # 淡紅
    ST_PRINTED: (0.90, 0.96, 0.90),   # 淡綠
    ST_CANCEL : (0.90, 0.90, 0.90),   # 灰
}


# ═══════════════════════════════════════════
# 資料整理
# ═══════════════════════════════════════════

def _s(v):
    return '' if v is None else str(v).strip()


def _fmt_date(dt):
    try:
        return dt.strftime('%Y/%m/%d')
    except Exception:
        return _s(dt)


def collect_travelers(records):
    """從 CLEANER 輸出取出貼紙名單（只含 野澤 / 斑尾 / 龍平 團客，不含純課與湯澤）"""
    out = {}
    for r in records:
        area = _s(r.get('項目'))
        if area not in COPIES:
            continue
        dt = r.get('_depart_date')
        date = _fmt_date(dt) if dt is not None else ''
        ono = _s(r.get('訂編'))
        cn, en = _s(r.get('中文姓名')), _s(r.get('英文姓名'))
        pid = _s(r.get('旅客編號')) or cn
        key = f"{ono}|{pid}"
        if key in out:                       # 極少數重複 key：加序避免覆蓋
            n = 2
            while f"{key}#{n}" in out:
                n += 1
            key = f"{key}#{n}"
        missing = []
        if not cn or re.match(r'^旅客\w*$', cn):
            missing.append('中文姓名空白/佔位')
        if not en:
            missing.append('英文姓名空白')
        if not date:
            missing.append('出發日期空白')
        out[key] = {'key': key, 'area': area, 'date': date, 'ono': ono,
                    'cn': cn, 'en': en, 'missing': missing,
                    'sort': (AREA_ORDER[area], date, ono, cn)}
    return out


def _version(p):
    return f"{p['area']}|{p['date']}|{p['cn']}|{p['en']}"


def sticker_text(p):
    """貼紙文字：3 行（出發日期 #訂編 / 中文 / 英文）"""
    return f"{p['date']}  #{p['ono']}\n{p['cn']}\n{p['en']}"


# ═══════════════════════════════════════════
# 試算表操作
# ═══════════════════════════════════════════

def _open(sheet_id):
    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(sheet_id)


def _retry(fn, tries=6):
    """Google Sheets 每分鐘寫入次數有上限（429）：等一下再試，最多 6 次"""
    import time
    for k in range(tries):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            if '429' not in str(e) or k == tries - 1:
                raise
            print(f"    （Google 寫入太頻繁，等 {20 * (k + 1)} 秒後重試）")
            time.sleep(20 * (k + 1))


def _get_or_create(sh, title, rows=100, cols=13):
    try:
        return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)


def _read_master(ws):
    vals = ws.get_all_values()
    if not vals or vals[0][:len(HEADERS)] != HEADERS:
        return {}
    out = {}
    for row in vals[1:]:
        row = row + [''] * (len(HEADERS) - len(row))
        d = {h: row[i] for h, i in H.items()}
        if d['key']:
            out[d['key']] = d
    return out


def _read_checkboxes(ws):
    """讀各區「印完了」勾選格，回傳被勾選的區域 set。
    讀取失敗直接報錯（不能默默當成沒勾，否則會漏標記已印）"""
    from gspread.utils import ValueRenderOption
    out = set()
    for area, cell in CHECK_CELLS.items():
        v = ws.acell(cell, value_render_option=ValueRenderOption.unformatted).value
        if v is True or str(v).upper() == 'TRUE':
            out.add(area)
    return out


def merge(current, existing, ticked_areas, today):
    """合併「今天的名單」與「上次的名單」，回傳依排序的資料列 dict 清單。
    ticked_areas：被勾選「印完了」的區域，這些區域上次印單上的人標記為已印"""
    rows = {}
    marked = 0
    # 1. 使用者勾選「印完了」→ 該區上次印單上的人，標記為已印
    if ticked_areas:
        for e in existing.values():
            if e['印單版本'] and e['區域'] in ticked_areas:
                e['已印版本'] = e['印單版本']
                e['列印日期'] = today
                e['印單版本'] = ''
                e['狀態'] = ST_PRINTED
                marked += 1

    for key, p in current.items():
        e = existing.get(key)
        ver = _version(p)
        printed_ver = e['已印版本'] if e else ''
        printed_date = e['列印日期'] if e else ''
        first_seen = (e['新增日期'] or today) if e else today
        note = ''
        manual = e['狀態'] if e else ''

        # 手動指定
        if manual == '已印' and not printed_ver:
            printed_ver, printed_date = ver, printed_date or today
        if manual == '重印':
            printed_ver = ''
            note = '手動標記重印'

        if p['missing']:
            status = ST_MISSING
            note = '、'.join(p['missing'])
            if printed_ver and printed_ver != ver:
                note += '（已印過，資料有異動）'
        elif printed_ver:
            if printed_ver == ver:
                status = ST_PRINTED
            else:
                status = ST_REPRINT
                old = printed_ver.split('|')
                note = f"已印：{old[1]} {old[2]} {old[3]} → 現在：{p['date']} {p['cn']} {p['en']}"
                if old[0] != p['area']:
                    note = f"區域 {old[0]}→{p['area']}；" + note
        else:
            status = ST_REPRINT if manual == '重印' else ST_PENDING
            if manual == '重印':
                note = '手動標記重印'

        rows[key] = {
            '狀態': status, '區域': p['area'], '出發日期': p['date'], '訂編': p['ono'],
            '中文姓名': p['cn'], '英文姓名': p['en'], '張數': COPIES[p['area']],
            '新增日期': first_seen, '列印日期': printed_date, '備註': note,
            'key': key, '印單版本': '', '已印版本': printed_ver,
            '_sort': p['sort'],
        }

    # 2. 新資料已經沒有的人 → 已取消（保留紀錄）
    for key, e in existing.items():
        if key in rows:
            continue
        rows[key] = {
            '狀態': ST_CANCEL, '區域': e['區域'], '出發日期': e['出發日期'], '訂編': e['訂編'],
            '中文姓名': e['中文姓名'], '英文姓名': e['英文姓名'],
            '張數': e['張數'], '新增日期': e['新增日期'], '列印日期': e['列印日期'],
            '備註': '新的原始大總表已無此人' + ('（已印過）' if e['已印版本'] else ''),
            'key': key, '印單版本': '', '已印版本': e['已印版本'],
            '_sort': (AREA_ORDER.get(e['區域'], 9), e['出發日期'], e['訂編'], e['中文姓名']),
        }

    ordered = sorted(rows.values(), key=lambda r: r['_sort'])
    # 3. 進列印分頁的人：記下「這次印單上的版本」，下次勾選時才知道印的是什麼
    for r in ordered:
        if r['狀態'] in (ST_PENDING, ST_REPRINT):
            r['印單版本'] = f"{r['區域']}|{r['出發日期']}|{r['中文姓名']}|{r['英文姓名']}"
    return ordered, marked


# ═══════════════════════════════════════════
# 寫入
# ═══════════════════════════════════════════

def _rgb(t):
    return {'red': t[0], 'green': t[1], 'blue': t[2]}


# ═══════════════════════════════════════════
# 新版貼紙版面：每張貼紙 = 7 欄 x 5 列的小格子（總寬/總高仍是 293 x 124，A4 分頁完全不變）
#   plain 貼紙：整張合併成 1 格，置中（與舊版一模一樣）
#   box   貼紙：[上留白][日期+中文 / 英文 兩行][空隙][一排 5 個正方形空框][下留白]
# ═══════════════════════════════════════════
# ⚠ 1. Google 匯出 PDF 做「符合頁寬」時，每多 1 個欄位就多算約 1px 寬（實測：3 欄排距 83.75pt、
#      21 欄 82.07pt）→ 每張貼紙拆成 7 欄，欄寬總和要減 6px 抵銷，列印縮放才會與舊版完全相同。
# ⚠ 2. 印表機印不到紙張邊緣（實測：左邊約 6.5mm、上方約 5mm 被切掉；「邊界 0」只是軟體設定）。
#      所以最左／最右那張貼紙的內容要離頁面側邊 >= 8mm；每頁第 1 排離上緣 >= 6.4mm、第 10 排離下緣 >= 6.4mm。
#      中間幾排沒有這個限制。三種排版只差上下留白，總高一律 124px。
SUB_COLS = [36, 43, 43, 43, 43, 43, 36]   # 左右留白 8.6mm + 5 個 43px(約 10.3mm)正方形；加總 = COL_WIDTH - (欄數 - 1)
_TEXT_H, _GAP_H, _BOX_H = 36, 6, 43       # 文字區(2 行) / 文字與框的空隙 / 框框高度（正方形）
PAD_FIRST, PAD_LAST, PAD_NORMAL = (27, 12), (12, 27), (19, 20)   # (上留白, 下留白)：每頁第 1 排 / 第 10 排 / 其他排
N_SUB = 5
BOX_LABELS = ['尺寸', '衣褲', '膝', '臀', '手套']   # 框框左上角小字，下方留白手寫
assert sum(SUB_COLS) == COL_WIDTH - (len(SUB_COLS) - 1)
assert all(sum(pad) + _TEXT_H + _GAP_H + _BOX_H == ROW_HEIGHT for pad in (PAD_FIRST, PAD_LAST, PAD_NORMAL))


def _sub_rows(r):
    """第 r 排貼紙（從 0 起算）的 5 個小列高度；A4 每頁 10 排，依在頁面中的位置決定上下留白"""
    k = r % ROWS_PER_PAGE
    top, bot = PAD_FIRST if k == 0 else PAD_LAST if k == ROWS_PER_PAGE - 1 else PAD_NORMAL
    return [top, _TEXT_H, _GAP_H, _BOX_H, bot]


def box_text(p):
    """box 貼紙文字：第 1 行 出發日期 + 中文姓名，第 2 行 英文姓名（不含訂編）"""
    return f"{p['date']}  {p['cn']}\n{p['en']}"


# Helvetica/Arial 字寬（1000 單位），用來估算英文姓名在 11pt 下的寬度
_AFM = {**dict(zip('ABCDEFGHIJKLMNOPQRSTUVWXYZ',
        [667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611])),
        **dict(zip('abcdefghijklmnopqrstuvwxyz',
        [556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500])),
        ' ': 278, '-': 333, '.': 278, ',': 278, "'": 191}
SAFE_TEXT_W = 215          # 文字可用寬度(px)：5 個框的總寬，左右各離邊 >= 8.6mm，避開印表機不可列印區


def _text_px(s, pt=FONT_SIZE):
    return sum(_AFM.get(ch, 556 if ch.isdigit() else 600) for ch in s) / 1000 * pt * 96 / 72


def _fit_runs(text, line_idx):
    """第 line_idx 行（0 起算）比安全寬度寬時，回傳 textFormatRuns 讓那一行縮小字體；不需要則回傳 None"""
    lines = text.split('\n')
    w = _text_px(lines[line_idx])
    if w <= SAFE_TEXT_W:
        return None
    fs = max(7, int(FONT_SIZE * SAFE_TEXT_W / w))
    start = sum(len(l) + 1 for l in lines[:line_idx])
    base = {'fontFamily': FONT_NAME, 'fontSize': FONT_SIZE}
    return [{'startIndex': 0, 'format': base},
            {'startIndex': start, 'format': {'fontFamily': FONT_NAME, 'fontSize': fs}}]


def _write_layout_tab(sh, title, sticker_rows):
    """sticker_rows：每個元素是一排貼紙（最多 3 張），每張 = ('box'|'plain', 文字) 或 None（空位）。
    第 1 排貼紙就是分頁最上面，不可有標題列。"""
    n = max(len(sticker_rows), 1)
    nrows, ncols = n * N_SUB, COLS_PER_ROW * len(SUB_COLS)
    ws = _get_or_create(sh, title, rows=nrows, cols=ncols)
    # 先取消所有合併、清掉所有格式（框線等）。ws.clear() 只清內容，舊版面的框線會殘留疊在新版面上
    _retry(lambda: sh.batch_update({'requests': [
        {'unmergeCells': {'range': {'sheetId': ws.id}}},
        {'updateCells': {'range': {'sheetId': ws.id}, 'fields': 'userEnteredFormat'}},
    ]}))
    ws.clear()
    ws.resize(rows=nrows, cols=ncols)
    sid = ws.id

    grid = [[''] * ncols for _ in range(nrows)]
    merges, boxes, rich = [], [], []
    for r, row in enumerate(sticker_rows):
        r0 = r * N_SUB
        for s, st in enumerate(row[:COLS_PER_ROW]):
            if not st:
                continue
            kind, text = st
            c0 = s * len(SUB_COLS)
            if kind == 'box':
                grid[r0 + 1][c0] = text
                runs = _fit_runs(text, 1)                       # 英文姓名那一行太寬 → 縮小字體
                if runs:
                    rich.append((r0 + 1, c0, text, runs))
                merges.append((r0 + 1, r0 + 2, c0, c0 + len(SUB_COLS)))          # 文字區（2 行）
                for k, lab in enumerate(BOX_LABELS):
                    grid[r0 + 3][c0 + 1 + k] = lab
                boxes.append((r0 + 3, c0 + 1, c0 + 1 + len(BOX_LABELS)))
            else:
                grid[r0][c0] = text
                if s == COLS_PER_ROW - 1:                       # 最右邊那張貼紙貼近頁面右緣，同樣避開不可列印區
                    runs = _fit_runs(text, 2)
                    if runs:
                        rich.append((r0, c0, text, runs))
                merges.append((r0, r0 + N_SUB, c0, c0 + len(SUB_COLS)))          # 整張合併、置中

    # 保護：Google 只依「有內容的範圍」做符合頁寬。若第 3 張貼紙的位置整頁都沒東西（例如龍平只剩 1~2 人），
    # 縮放會變大、貼紙全部對不齊 → 在最右邊放一個空白字元，讓範圍永遠涵蓋 3 張貼紙的寬度
    if not any(len(row) > 2 and row[2] for row in sticker_rows):
        grid[0][ncols - 1] = ' '

    _retry(lambda: ws.update(range_name='A1', values=grid, value_input_option='RAW'))

    rng = lambda r1, r2, c1, c2: {'sheetId': sid, 'startRowIndex': r1, 'endRowIndex': r2,
                                  'startColumnIndex': c1, 'endColumnIndex': c2}
    reqs = []
    for c in range(ncols):
        reqs.append({'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': c, 'endIndex': c + 1},
            'properties': {'pixelSize': SUB_COLS[c % len(SUB_COLS)]}, 'fields': 'pixelSize'}})
    for r in range(nrows):
        reqs.append({'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'ROWS', 'startIndex': r, 'endIndex': r + 1},
            'properties': {'pixelSize': _sub_rows(r // N_SUB)[r % N_SUB]}, 'fields': 'pixelSize'}})
    reqs.append({'repeatCell': {
        'range': rng(0, nrows, 0, ncols),
        'cell': {'userEnteredFormat': {
            'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE', 'wrapStrategy': 'WRAP',
            'textFormat': {'fontFamily': FONT_NAME, 'fontSize': FONT_SIZE},
            'backgroundColor': {'red': 1, 'green': 1, 'blue': 1}}},
        'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,wrapStrategy,textFormat,backgroundColor)'}})
    reqs.append({'updateSheetProperties': {
        'properties': {'sheetId': sid, 'gridProperties': {'hideGridlines': True}},
        'fields': 'gridProperties.hideGridlines'}})
    for (r1, r2, c1, c2) in merges:
        reqs.append({'mergeCells': {'range': rng(r1, r2, c1, c2), 'mergeType': 'MERGE_ALL'}})
    black = {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}}
    for (r, c1, c2) in boxes:
        reqs.append({'updateBorders': {'range': rng(r, r + 1, c1, c2), 'top': black, 'bottom': black,
                                       'left': black, 'right': black, 'innerVertical': black}})
        reqs.append({'repeatCell': {
            'range': rng(r, r + 1, c1, c2),
            'cell': {'userEnteredFormat': {'horizontalAlignment': 'LEFT', 'verticalAlignment': 'TOP',
                                           'textFormat': {'fontFamily': FONT_NAME, 'fontSize': 8}}},
            'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat)'}})
    for (r, c, text, runs) in rich:
        reqs.append({'updateCells': {
            'range': rng(r, r + 1, c, c + 1),
            'rows': [{'values': [{'userEnteredValue': {'stringValue': text}, 'textFormatRuns': runs}]}],
            'fields': 'userEnteredValue,textFormatRuns'}})
    for i in range(0, len(reqs), 1000):
        _retry(lambda: sh.batch_update({'requests': reqs[i:i + 1000]}))
    if rich:
        print(f"    {title}：{len(rich)} 張貼紙因英文姓名太長而縮小字體")
    return ws


def _print_rows(persons):
    """依區域回傳「貼紙排」清單，交給 _write_layout_tab。
    野澤 / 斑尾：每人 1 排 = 新版貼紙 + 舊版 2 張（含訂編）；龍平：每排 3 人，每人 1 張新版貼紙。"""
    out = {}
    for area, copies in COPIES.items():
        ps = [p for p in persons if p['area'] == area]
        if copies == 1:
            out[area] = [[('box', box_text(p)) for p in ps[i:i + COLS_PER_ROW]]
                         for i in range(0, len(ps), COLS_PER_ROW)]
        else:
            out[area] = [[('box', box_text(p)), ('plain', sticker_text(p)), ('plain', sticker_text(p))]
                         for p in ps]
    return out


def _write_master(sh, ordered):
    ws = _get_or_create(sh, TAB_MASTER, rows=max(len(ordered) + 1, 2), cols=len(HEADERS))
    ws.clear()
    ws.resize(rows=max(len(ordered) + 1, 2), cols=len(HEADERS))
    values = [HEADERS] + [[r[h] for h in HEADERS] for r in ordered]
    ws.update(range_name='A1', values=values, value_input_option='RAW')
    sid = ws.id
    n = len(ordered) + 1
    reqs = [
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1},
                        'cell': {'userEnteredFormat': {
                            'backgroundColor': {'red': 0.2, 'green': 0.3, 'blue': 0.5},
                            'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}}},
                        'fields': 'userEnteredFormat(backgroundColor,textFormat)'}},
        {'updateSheetProperties': {'properties': {'sheetId': sid, 'gridProperties': {'frozenRowCount': 1}},
                                   'fields': 'gridProperties.frozenRowCount'}},
        {'setBasicFilter': {'filter': {'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': n,
                                                  'startColumnIndex': 0, 'endColumnIndex': len(HEADERS)}}}},
    ]
    # 依狀態把連續相同狀態的列一起上色
    i = 0
    while i < len(ordered):
        j = i
        while j + 1 < len(ordered) and ordered[j + 1]['狀態'] == ordered[i]['狀態']:
            j += 1
        reqs.append({'repeatCell': {
            'range': {'sheetId': sid, 'startRowIndex': i + 1, 'endRowIndex': j + 2,
                      'startColumnIndex': 0, 'endColumnIndex': len(HEADERS)},
            'cell': {'userEnteredFormat': {'backgroundColor': _rgb(COLORS[ordered[i]['狀態']])}},
            'fields': 'userEnteredFormat.backgroundColor'}})
        i = j + 1
    widths = [70, 50, 90, 60, 80, 190, 45, 85, 85, 340, 130, 200, 200]
    for ci, px in enumerate(widths):
        reqs.append({'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': ci, 'endIndex': ci + 1},
            'properties': {'pixelSize': px}, 'fields': 'pixelSize'}})
    sh.batch_update({'requests': reqs})
    return ws


def _write_summary(sh, ordered, marked, src_name, today, now_str):
    ws = _get_or_create(sh, TAB_SUMMARY, rows=40, cols=6)
    ws.clear()

    def cnt(status, area=None, new_only=False):
        return sum(1 for r in ordered if r['狀態'] == status and (area is None or r['區域'] == area)
                   and (not new_only or r['新增日期'] == today))

    areas = ['野澤', '斑尾', '龍平']
    pend = {a: cnt(ST_PENDING, a) + cnt(ST_REPRINT, a) for a in areas}
    per_page = {a: ROWS_PER_PAGE * (COLS_PER_ROW if COPIES[a] == 1 else 1) for a in areas}   # 每頁幾人

    def line(label, fn):
        vals = [fn(a) for a in areas]
        return [label] + vals + [sum(vals)]

    def est(a):
        n, pp = pend[a], per_page[a]
        how = '1 人 1 張，每頁 30 人' if COPIES[a] == 1 else '1 人 1 列 × 3 張，每頁 10 人'
        return [f'[列印用] {a}（{how}）', n, f'{n / pp:.1f} 頁（再 {(-n) % pp} 人可湊滿整頁）']

    rows = [
        ['🖨️ 旅客貼紙總覽'],
        [f'更新時間：{now_str}　｜　資料來源：{src_name}' + (f'　｜　本次依勾選，標記 {marked} 人為已印' if marked else '')],
        [],
        ['✅ 野澤 的貼紙印完了 → 勾選右邊格子（下次跑每日流程時，會把「野澤」列印分頁上的人標記為已印）', False],
        ['✅ 斑尾 的貼紙印完了 → 勾選右邊格子（下次跑每日流程時，會把「斑尾」列印分頁上的人標記為已印）', False],
        ['✅ 龍平 的貼紙印完了 → 勾選右邊格子（下次跑每日流程時，會把「龍平」列印分頁上的人標記為已印）', False],
        ['⚠ 各區獨立勾選，只勾你印完的那一區。請印完後立刻勾；如果勾選前又跑過每日流程，請先確認該區列印分頁內容再勾'],
        [],
        ['項目', '野澤', '斑尾', '龍平', '合計'],
        line('待印（人）', lambda a: pend[a]),
        line('　其中今天新增（人）', lambda a: cnt(ST_PENDING, a, True)),
        line('　其中需重印（資料異動，人）', lambda a: cnt(ST_REPRINT, a)),
        line('待印貼紙張數', lambda a: pend[a] * COPIES[a]),
        line('待補資料（姓名空白/佔位，不會進列印分頁）', lambda a: cnt(ST_MISSING, a)),
        line('已印（人）', lambda a: cnt(ST_PRINTED, a)),
        line('已取消（新資料已無此人）', lambda a: cnt(ST_CANCEL, a)),
        line('全部（人）', lambda a: sum(1 for r in ordered if r['區域'] == a)),
        [],
        ['A4 估算', '人數', '頁數（不滿一頁的部分，下一批會接著累積）'],
        est('野澤'), est('斑尾'), est('龍平'),
        [],
        ['列印設定（每次相同）：檔案 → 列印 → 目前的工作表 → A4 → 頁面方向「縱向」→ 縮放「符合頁寬」→ 邊界「自訂」全部填 0 → 取消頁首頁尾與格線 → 預覽總頁數要和上表一致'],
        ['名單的狀態欄可手動輸入：「已印」＝標記為已印過；「重印」＝強制重印'],
    ]
    ws.update(range_name='A1', values=rows, value_input_option='RAW')
    sid = ws.id
    reqs = [
        {'setDataValidation': {
            'range': {'sheetId': sid, 'startRowIndex': 3, 'endRowIndex': 6, 'startColumnIndex': 1, 'endColumnIndex': 2},
            'rule': {'condition': {'type': 'BOOLEAN'}, 'strict': True}}},
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1},
                        'cell': {'userEnteredFormat': {'textFormat': {'bold': True, 'fontSize': 16}}},
                        'fields': 'userEnteredFormat.textFormat'}},
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 3, 'endRowIndex': 6},
                        'cell': {'userEnteredFormat': {'backgroundColor': _rgb((1.0, 0.95, 0.70)),
                                                       'textFormat': {'bold': True}}},
                        'fields': 'userEnteredFormat(backgroundColor,textFormat)'}},
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 8, 'endRowIndex': 9},
                        'cell': {'userEnteredFormat': {'backgroundColor': {'red': 0.2, 'green': 0.3, 'blue': 0.5},
                                                       'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}}},
                        'fields': 'userEnteredFormat(backgroundColor,textFormat)'}},
        {'repeatCell': {'range': {'sheetId': sid, 'startRowIndex': 18, 'endRowIndex': 19},
                        'cell': {'userEnteredFormat': {'backgroundColor': {'red': 0.2, 'green': 0.3, 'blue': 0.5},
                                                       'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}}},
                        'fields': 'userEnteredFormat(backgroundColor,textFormat)'}},
        {'updateDimensionProperties': {'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 1},
                                       'properties': {'pixelSize': 430}, 'fields': 'pixelSize'}},
        {'updateDimensionProperties': {'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': 1, 'endIndex': 5},
                                       'properties': {'pixelSize': 90}, 'fields': 'pixelSize'}},
    ]
    sh.batch_update({'requests': reqs})
    return ws


def _order_tabs(sh):
    want = [TAB_SUMMARY] + [TAB_PRINT[a] for a in COPIES] + [TAB_MASTER]
    reqs = []
    for i, t in enumerate(want):
        try:
            reqs.append({'updateSheetProperties': {
                'properties': {'sheetId': sh.worksheet(t).id, 'index': i}, 'fields': 'index'}})
        except gspread.exceptions.WorksheetNotFound:
            pass
    if reqs:
        sh.batch_update({'requests': reqs})


# ═══════════════════════════════════════════
# 主函數
# ═══════════════════════════════════════════

def run(cleaner_output=None, sheet_id=None, src_name=''):
    print("=" * 50)
    print("STICKER 旅客貼紙員 啟動")
    print("=" * 50)

    sheet_id = sheet_id or STICKER_SPREADSHEET_ID
    if not sheet_id:
        print("  略過：尚未設定貼紙試算表（shared/config.py 的 STICKER_SPREADSHEET_ID）")
        return None

    if cleaner_output is None:
        from agents.cleaner import load
        cleaner_output = load()
    records = cleaner_output['all_records_flat']
    if not src_name:
        src_name = cleaner_output.get('source_file', '') or ''

    current = collect_travelers(records)
    print(f"  貼紙名單：{len(current)} 人（野澤/斑尾/龍平團客）")
    if len(current) < 50:
        # 保護機制：名單異常少（資料讀取失敗/檔案錯誤）時不動貼紙試算表，避免把大家標成「已取消」
        raise RuntimeError(f"貼紙名單只有 {len(current)} 人，疑似資料異常，已停止更新（請確認大總表）")

    sh = _open(sheet_id)
    today = datetime.now().strftime('%Y/%m/%d')
    now_str = datetime.now().strftime('%Y/%m/%d %H:%M')

    ws_sum = _get_or_create(sh, TAB_SUMMARY, rows=40, cols=6)
    ws_master = _get_or_create(sh, TAB_MASTER, rows=2, cols=len(HEADERS))
    existing = _read_master(ws_master)
    ticked = _read_checkboxes(ws_sum) if existing else set()
    print(f"  上次名單：{len(existing)} 人" + (f"｜偵測到「印完了」勾選：{'、'.join(sorted(ticked))}" if ticked else ''))

    ordered, marked = merge(current, existing, ticked, today)
    pending = [r for r in ordered if r['狀態'] in (ST_PENDING, ST_REPRINT)]
    people = [current[r['key']] for r in pending]
    rows_by_area = _print_rows(people)

    _write_master(sh, ordered)
    for area in COPIES:
        _write_layout_tab(sh, TAB_PRINT[area], rows_by_area[area])
    for old in TAB_PRINT_OLD + ['[預覽] 新版貼紙']:   # 刪除舊版合併分頁與預覽分頁
        try:
            sh.del_worksheet(sh.worksheet(old))
        except gspread.exceptions.WorksheetNotFound:
            pass
    _write_summary(sh, ordered, marked, src_name, today, now_str)
    _order_tabs(sh)

    from collections import Counter
    c = Counter(r['狀態'] for r in ordered)
    print("  狀態：" + '、'.join(f"{k} {v}" for k, v in c.items()))
    for area in COPIES:
        n = sum(1 for r in pending if r['區域'] == area)
        print(f"  列印分頁 {area}：{len(rows_by_area[area])} 排貼紙（{n} 人 x {COPIES[area]} 張）")
    print(f"  OK 貼紙試算表已更新{'（本次標記 ' + str(marked) + ' 人為已印）' if marked else ''}")
    return {'pending': len(pending), 'marked': marked}


if __name__ == '__main__':
    run()
