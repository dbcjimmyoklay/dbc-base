"""售價與成本總表產生器。
用法: python cost_report.py [原始大總表xlsx] [--xlsx]
  預設寫入 Google 試算表(ID 存在 cost_sheet_id.txt,第一次會自動建立);加 --xlsx 則改輸出 Excel 檔。
  不給檔名則取資料夾內最新的 原始大總表MMDD.xlsx
只讀必要欄位，不讀任何個資。重新產生時會保留你在 Excel 填過的單價、匯率與租借價格。
"""
import sys, re, glob, os, warnings
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter as L

warnings.filterwarnings('ignore')
BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, '售價成本總表.xlsx')

# ---------- 分類 ----------
CATS = {  # key: (sheet, 項目 list)
    ('野澤', True):  ['野澤住宿(1泊)', '東京住宿(1泊)', '交通 機場>雪場', '交通 雪場>市區', '雪票(1日)', '雪靴(1日)', '下山巴士', '機票', '保險'],
    ('野澤', False): ['野澤住宿(1泊)', '東京住宿(1泊)', '交通 機場>雪場', '交通 雪場>市區', '雪票(1日)', '雪靴(1日)', '下山巴士', '保險'],
    ('斑尾', True):  ['斑尾住宿(1泊)', '東京住宿(1泊)', '交通 機場>雪場', '交通 雪場>市區', '雪票(1日)', '三點set', '下山接駁', '機票', '保險'],
    ('斑尾', False): ['斑尾住宿(1泊)', '東京住宿(1泊)', '交通 機場>雪場', '交通 雪場>市區', '雪票(1日)', '三點set', '下山接駁', '保險'],
    ('龍平', True):  ['龍平住宿+早餐', '首爾住宿+早餐', '交通 機場>雪場', '交通 雪場>市區', '保護費', '三點set+雪票', '機票', '保險'],
    ('龍平', False): ['龍平住宿+早餐', '首爾住宿+早餐', '交通 機場>雪場', '交通 雪場>市區', '保護費', '三點set+雪票', '保險'],
}
# 每個項目的數量算法: lodge 雪場泊數 / city 市區泊數 / days 雪票天數 / pax 每人1
RULE = {'野澤住宿(1泊)': 'lodge', '斑尾住宿(1泊)': 'lodge', '龍平住宿+早餐': 'lodge',
        '東京住宿(1泊)': 'city', '首爾住宿+早餐': 'city',
        '雪票(1日)': 'days', '雪靴(1日)': 'days', '三點set': 'days', '三點set+雪票': 'days'}
RULE_TXT = {'lodge': '標準3泊 + 續滑天數(保留市區住宿者不加)', 'city': '標準1泊 + 市區續住;取消者標準為0',
            'days': '標準2日 + 續滑天數', 'pax': '每人1'}
SHEET = lambda r, a: f"{r}{'含' if a else '不含'}機票"
GRP = re.compile(r'^S4-(J|K)-(N|B|Y)([TS])?([0A])(\d{6})([A-Z]?)$')
PURE = re.compile(r'^S4-Z([YNB])\d{6}')
PURE_NAME = {'Y': '湯澤', 'N': '野澤', 'B': '斑尾'}


def classify(g):
    m = PURE.match(g)
    if m:
        return ('純課程', PURE_NAME[m.group(1)])
    m = GRP.match(g)
    if m:
        area, r, _, flag, _, _ = m.groups()
        reg = {('J', 'N'): '野澤', ('J', 'B'): '斑尾', ('K', 'Y'): '龍平'}.get((area, r))
        if reg:
            return ('團客', reg, flag == 'A')
    return None  # 例外


# ---------- 讀資料 ----------
EXT = {'野澤-續滑1日含教練/人': 1, '野澤-續滑2日含教練/人': 2, '斑尾-續滑1日含教練/人': 1, '斑尾-續滑2日含教練/人': 2,
       '斑尾-續滑3日含教練/人': 3, '龍平-續滑1日不含教練/人': 1, '龍平-續滑1日含教練/人': 1, '續滑1日不含教練/人': 1,
       '續滑2日不含教練/人': 2, '續滑3日不含教練/人': 3, '野澤-雪場續滑-不含教練(保留市區住宿)': 1}
KEEP_COL = '野澤-雪場續滑-不含教練(保留市區住宿)'
CITY_EXT = {'東京續住1日/人': 1, '東京續住2日/人': 2, '東京續住3日/人': 3, '首爾續住1晚/人': 1, '首爾續住2晚/人': 2}
RENT = ['租借-手套', '租借-雪衣雪褲', '租借-護膝', '租借-護臀']


def load(path):
    want = ['團號', 'OP備註', '出發日期', '報名日期', '訂單編號', '旅客編號', '中文姓名', 'LEVEL', '滑雪類別', '團費',
            '旅客訂單備註', '分房備住', '分房備住.1'] + list(EXT) + list(CITY_EXT) + RENT
    d = pd.read_excel(path, header=1, dtype=str, usecols=lambda c: c in want)
    d['團號'] = d['團號'].ffill()
    d = d[d['團號'].str.match(r'^(S4|DBC)-', na=False)]
    d['OP備註'] = d['OP備註'].where(d['OP備註'].notna(), None)
    d['OP備註'] = d.groupby('團號')['OP備註'].transform(lambda s: s.ffill().bfill())
    d['出發日期'] = d.groupby('團號')['出發日期'].transform(lambda s: s.ffill().bfill())
    d = d[d['中文姓名'].fillna('') != '取消']
    d = d[d['旅客編號'].notna() | d['中文姓名'].notna()].copy()
    num = lambda c: pd.to_numeric(d[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0) if c in d else 0
    d['fee'] = num('團費')
    ext = sum(num(c) * v for c, v in EXT.items())
    txt = (d.get('分房備住', '').fillna('') + d.get('分房備住.1', '').fillna('') + d.get('旅客訂單備註', '').fillna(''))
    keep = (num(KEEP_COL) > 0) | txt.str.contains('保留市區')
    cancel = (d.get('分房備住', '').fillna('') + d.get('分房備住.1', '').fillna('')).str.contains('取消東京|取消首爾')
    d['days'] = 2 + ext
    d['lodge'] = 3 + ext.where(~keep, 0)
    d['city'] = (~cancel).astype(int) + sum(num(c) * v for c, v in CITY_EXT.items())
    return d


def pure_days(d):
    """純課程租借天數: 從 OP備註 '#訂單 ... 共N日' 取得"""
    mp = {}
    for op in d.loc[d['團號'].str.match(PURE), 'OP備註'].dropna().unique():
        for m in re.finditer(r'#(\d+)(.*?)(?=#\d+|$)', op, re.S):
            n = re.search(r'共(\d+)日', m.group(2))
            if n:
                mp[int(m.group(1))] = int(n.group(1))
    return mp


def rental_rev(d, price, days):
    q = lambda c: pd.to_numeric(d[c], errors='coerce').fillna(0) if c in d else 0
    return q('租借-手套') * price['手套'] + (q('租借-雪衣雪褲') * price['雪衣雪褲'] + q('租借-護臀') * price['護臀'] + q('租借-護膝') * price['護膝']) * days


# ---------- 讀舊檔使用者輸入 ----------
def old_inputs():
    inp = {'設定': {}, 'items': {}}
    if not os.path.exists(OUT):
        return inp
    try:
        wb = load_workbook(OUT)
        if '設定' in wb:
            for r in wb['設定'].iter_rows(min_row=1, max_row=15, max_col=2, values_only=True):
                if r[0] and r[1] is not None:
                    inp['設定'][r[0]] = r[1]
        if '設定' in wb:
            for r in wb['設定'].iter_rows(min_row=15, max_col=3, values_only=True):
                if r[0] in ('野澤', '斑尾', '龍平') and isinstance(r[2], (int, float)):
                    inp['items'][(r[0], r[1])] = r[2]
        for ws in wb.worksheets:  # 舊版格式:單價填在各分類分頁
            for r in ws.iter_rows(min_row=5, max_row=16, max_col=2, values_only=True):
                if r[0] and isinstance(r[1], (int, float)) and ws.title[:2] in ('野澤', '斑尾', '龍平'):
                    inp['items'].setdefault((ws.title[:2], r[0]), r[1])
    except Exception:
        pass
    return inp


YEL = PatternFill('solid', fgColor='FFF2A8')
HDR = PatternFill('solid', fgColor='1F3A4A')
GRY = PatternFill('solid', fgColor='E8EEF1')
BOLD = Font(bold=True)
WH = Font(bold=True, color='FFFFFF')


def hdr(ws, row, vals, col=1):
    for i, v in enumerate(vals):
        c = ws.cell(row, col + i, v); c.fill = HDR; c.font = WH; c.alignment = Alignment(horizontal='center', wrap_text=True)


# ---------- Google 試算表 ----------
ID_FILE = os.path.join(BASE, 'cost_sheet_id.txt')
SHARE_TO = ['jimmyoklay@gmail.com']


def open_gsheet():
    import gspread
    from google.oauth2.service_account import Credentials
    sys.path.insert(0, BASE)
    from shared.config import CREDS_PATH, SCOPES
    gc = gspread.authorize(Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES))
    if os.path.exists(ID_FILE):
        return gc.open_by_key(open(ID_FILE).read().strip())
    sh = gc.create('DBC 售價成本總表')
    for m in SHARE_TO:
        sh.share(m, perm_type='user', role='writer', notify=False)
    open(ID_FILE, 'w').write(sh.id)
    print('已建立 Google 試算表:', sh.url)
    return sh


def old_inputs_gs(sh):
    inp = {'設定': {}, 'items': {}}
    try:
        rows = sh.worksheet('設定').get_all_values(value_render_option='UNFORMATTED_VALUE')
    except Exception:
        return inp
    for i, r in enumerate(rows):
        r = r + ['', '', '']
        if i < 14 and r[0] and isinstance(r[1], (int, float)):
            inp['設定'][r[0]] = r[1]
        if i >= 15 and r[0] in ('野澤', '斑尾', '龍平') and isinstance(r[2], (int, float)):
            inp['items'][(r[0], r[1])] = r[2]
    return inp


def push_gsheet(sh, wb):
    import gspread
    from gspread.utils import rowcol_to_a1
    date_re = re.compile(r'^\d{4}/\d{2}/\d{2}$')
    have = {w.title: w for w in sh.worksheets()}
    order = [w.title for w in wb.worksheets]
    for t in order:
        if t not in have:
            have[t] = sh.add_worksheet(t, rows=200, cols=40)
    for i, t in enumerate(order):
        sh.batch_update({'requests': [{'updateSheetProperties': {'properties': {'sheetId': have[t].id, 'index': i}, 'fields': 'index'}}]})
    if 'Sheet1' in have and 'Sheet1' not in order:
        sh.del_worksheet(have['Sheet1'])
    reqs = []
    for wso in wb.worksheets:
        g = have[wso.title]
        nr, nc = wso.max_row, wso.max_column
        if g.row_count < nr + 5 or g.col_count < nc + 2:
            g.resize(rows=max(g.row_count, nr + 5), cols=max(g.col_count, nc + 2))
        g.clear()
        vals = []
        for row in wso.iter_rows(min_row=1, max_row=nr, max_col=nc):
            line = []
            for c in row:
                v = c.value
                if v is None: line.append('')
                elif isinstance(v, str) and date_re.match(v): line.append("'" + v)
                else: line.append(v)
            vals.append(line)
        g.update(range_name='A1', values=vals, value_input_option='USER_ENTERED')
        sid = g.id
        full = {'sheetId': sid}
        rng = lambda r1, c1, r2, c2: {'sheetId': sid, 'startRowIndex': r1, 'endRowIndex': r2, 'startColumnIndex': c1, 'endColumnIndex': c2}
        reqs.append({'repeatCell': {'range': full, 'cell': {'userEnteredFormat': {}}, 'fields': 'userEnteredFormat'}})
        rgb = lambda h: {k: int(h[a:a + 2], 16) / 255 for k, a in (('red', 0), ('green', 2), ('blue', 4))}
        def key(c):
            fill = str(c.fill.fgColor.rgb)[-6:] if c.fill is not None and c.fill.fill_type == 'solid' else None
            f = c.font
            fc = f.color.rgb[-6:] if f.color is not None and f.color.type == 'rgb' and isinstance(f.color.rgb, str) else None
            ha = c.alignment.horizontal; wr = bool(c.alignment.wrap_text)
            bd = c.border is not None and c.border.left is not None and c.border.left.style is not None
            nf = c.number_format if c.number_format != 'General' else None
            return (fill, bool(f.bold), bool(f.italic), int(f.size or 11), fc, ha, wr, bd, nf)
        DEF = (None, False, False, 11, None, None, False, False, None)
        def fmt_of(k):
            fill, b, it, sz, fc, ha, wr, bd, nf = k
            f = {}
            if fill: f['backgroundColor'] = rgb(fill)
            tf = {'bold': b, 'italic': it, 'fontSize': sz}
            if fc: tf['foregroundColor'] = rgb(fc)
            f['textFormat'] = tf
            f['verticalAlignment'] = 'MIDDLE'
            if ha: f['horizontalAlignment'] = ha.upper()
            if wr: f['wrapStrategy'] = 'WRAP'
            if bd:
                side = {'style': 'SOLID', 'color': rgb('C5D0D6')}
                f['borders'] = {'top': side, 'bottom': side, 'left': side, 'right': side}
            if nf: f['numberFormat'] = {'type': 'NUMBER', 'pattern': nf}
            return f
        for row in wso.iter_rows(min_row=1, max_row=nr, max_col=nc):
            run = None
            for c in list(row) + [None]:
                k = key(c) if c is not None else None
                if run and (c is None or k != run[0] or c.column != run[2] + 1):
                    reqs.append({'repeatCell': {'range': rng(run[1] - 1, run[2] - run[3], run[1], run[2]), 'cell': {'userEnteredFormat': fmt_of(run[0])},
                                                'fields': 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment,wrapStrategy,borders,numberFormat)'}})
                    run = None
                if c is not None and k != DEF:
                    if run: run[2] = c.column; run[3] += 1
                    else: run = [k, c.row, c.column, 1]
        for ci in range(1, nc + 1):
            w = wso.column_dimensions[L(ci)].width if L(ci) in wso.column_dimensions and wso.column_dimensions[L(ci)].width else 14
            reqs.append({'updateDimensionProperties': {'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': ci - 1, 'endIndex': ci}, 'properties': {'pixelSize': int(w * 7.5 + 8)}, 'fields': 'pixelSize'}})
        for rn, dim in wso.row_dimensions.items():
            if dim.height:
                reqs.append({'updateDimensionProperties': {'range': {'sheetId': sid, 'dimension': 'ROWS', 'startIndex': rn - 1, 'endIndex': rn}, 'properties': {'pixelSize': int(dim.height * 1.3)}, 'fields': 'pixelSize'}})
        tc = wso.sheet_properties.tabColor
        props = {'sheetId': sid, 'gridProperties': {'hideGridlines': True, 'frozenRowCount': 4 if wso.title[:2] in ('野澤', '斑尾', '龍平') else 0}}
        fields = 'gridProperties.hideGridlines,gridProperties.frozenRowCount'
        if tc is not None and isinstance(tc.rgb, str):
            props['tabColor'] = rgb(tc.rgb[-6:]); fields += ',tabColor'
        reqs.append({'updateSheetProperties': {'properties': props, 'fields': fields}})
    for i in range(0, len(reqs), 200):
        sh.batch_update({'requests': reqs[i:i + 200]})
    print('已更新 Google 試算表:', sh.url)


# ---------- 美化 ----------
from openpyxl.styles import Border, Side
NAVY = '1F3A4A'; MONEY = '#,##0;[Red]-#,##0'
REG_LIGHT = {'野澤': 'DCE9F5', '斑尾': 'DDF0E8', '龍平': 'FBE5D0', '純課程': 'EDE7F6', '總覽': 'DCE6EC', '設定': 'FFF4CC', '例外': 'EEEEEE'}
REG_DARK = {'野澤': '2F6DA3', '斑尾': '2E8B6A', '龍平': 'C8651B', '純課程': '6B4FA0', '總覽': NAVY, '設定': '9A6B00', '例外': '666666'}
_t = Side(style='thin', color='C5D0D6')
BOX = Border(left=_t, right=_t, top=_t, bottom=_t)


def _paint(ws, r1, c1, r2, c2, fill=None, bold=None, size=None, color=None, fmt=None, border=False, halign=None, italic=None):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            x = ws.cell(r, c)
            if fill: x.fill = PatternFill('solid', fgColor=fill)
            if bold is not None or size or color or italic is not None:
                f = x.font
                col = color or (f.color.rgb[-6:] if f.color is not None and f.color.type == 'rgb' and isinstance(f.color.rgb, str) else None)
                x.font = Font(bold=f.bold if bold is None else bold, italic=f.italic if italic is None else italic,
                              size=size or f.size, color=col)
            if fmt: x.number_format = fmt
            if border: x.border = BOX
            if halign: x.alignment = Alignment(horizontal=halign, vertical='center', wrap_text=x.alignment.wrap_text)


def _title(ws, key, text_cols):
    ws.row_dimensions[1].height = 30
    _paint(ws, 1, 1, 1, text_cols, fill=REG_LIGHT[key])
    _paint(ws, 1, 1, 1, 1, bold=True, size=16, color=REG_DARK[key])


def _band(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        if (r - r1) % 2 == 1:
            _paint(ws, r, c1, r, c2, fill='F5F8FA')


def beautify(wb, META):
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        key = ws.title if ws.title in REG_LIGHT else ws.title[:2]
        ws.sheet_properties.tabColor = REG_DARK.get(key, NAVY)
    for name, m in META.items():
        if 'reg' not in m:
            continue
        ws = wb[name]; reg = m['reg']; n, tr, rr, R = m['n'], m['tr'], m['rr'], m['R']
        _title(ws, reg, 8)
        _paint(ws, 2, 1, 2, 1, bold=True); _paint(ws, 2, 2, 2, 2, bold=True, size=12, fmt='#,##0')
        ws.row_dimensions[4].height = 34
        _band(ws, 5, tr - 1, 1, 7)
        _paint(ws, 5, 1, tr, 7, border=True, fmt=MONEY)
        _paint(ws, 5, 2, tr - 1, 2, fill='E6F0FA', color='1F5F99')
        _paint(ws, 5, 6, tr, 6, fmt='0.00')
        _paint(ws, 5, 8, tr - 1, 8, color='7A8A94', italic=True, size=10)
        _paint(ws, tr, 1, tr, 7, fill='DDE6EB', bold=True)
        _paint(ws, rr + 1, 1, rr + 10, 3, border=True, fmt=MONEY)
        for k in ('計入收入', '總成本'):
            _paint(ws, R[k], 1, R[k], 3, bold=True)
        _paint(ws, R['最後結果'], 1, R['最後結果'], 3, fill='CDE8D3', bold=True, size=12, color='1E6B34')
        for k in ('每人 計入收入', '每人 總成本', '每人 結果'):
            _paint(ws, R[k], 1, R[k], 3, fill='F0F7F2')
        _paint(ws, R['每人 結果'], 1, R['每人 結果'], 3, bold=True, color='1E6B34')
        ws.row_dimensions[rr].height = 30
        g0, f1, l1, lc = m['gs'], m['first'], m['last'], m['last_col']
        _paint(ws, g0 - 1, 1, g0 - 1, 1, bold=True, size=12, color=REG_DARK[reg])
        ws.row_dimensions[g0].height = 34
        if l1 >= f1:
            _paint(ws, f1, 1, l1, lc, border=True, fmt=MONEY)
            _band(ws, f1, l1, 1, lc)
            _paint(ws, f1, 2, l1, 3, halign='center')
            _paint(ws, f1, lc - 3, l1, lc, fill='EAF5EC')
            _paint(ws, f1, lc - 5, l1, lc - 4, bold=True)
        ws.column_dimensions['A'].width = 32
        for c in range(2, lc + 1):
            ws.column_dimensions[L(c)].width = 16
    # 純課程
    m = META['純課程']; ws = wb['純課程']
    _title(ws, '純課程', 8)
    ws.row_dimensions[3].height = 30
    _paint(ws, 4, 1, 7, 8, border=True, fmt=MONEY)
    _paint(ws, 4, 1, 6, 1, bold=True)
    _paint(ws, 7, 1, 7, 8, fill='DDE6EB', bold=True)
    _paint(ws, 4, 5, 6, 8, fill='EAF5EC')
    _paint(ws, 7, 5, 7, 8, fill='CDE8D3', bold=True, color='1E6B34')
    ws.row_dimensions[m['gs']].height = 28
    _paint(ws, m['gs'] - 1, 1, m['gs'] - 1, 1, bold=True, size=12, color=REG_DARK['純課程'])
    if m['last'] >= m['first']:
        _paint(ws, m['first'], 1, m['last'], 8, border=True, fmt=MONEY)
        _band(ws, m['first'], m['last'], 1, 8)
        _paint(ws, m['first'], 2, m['last'], 3, halign='center')
        _paint(ws, m['first'], 7, m['last'], 8, fill='EAF5EC')
    ws.column_dimensions['A'].width = 24
    for c in range(2, 9):
        ws.column_dimensions[L(c)].width = 18
    # 總覽
    m = META['總覽']; ws = wb['總覽']
    _title(ws, '總覽', 10)
    _paint(ws, 2, 1, 2, 1, italic=True, color='5D6B74')
    ws.row_dimensions[4].height = 30
    t = m['total']
    _paint(ws, 5, 1, t, 10, border=True, fmt=MONEY)
    _band(ws, 5, t - 1, 1, 10)
    _paint(ws, 5, 2, t, 3, halign='center')
    for i, nm in enumerate(m['rows']):
        k = nm[:2] if nm[:2] in REG_LIGHT else nm
        _paint(ws, 5 + i, 1, 5 + i, 1, fill=REG_LIGHT[k], bold=True)
    _paint(ws, 5, 7, t, 8, fill='EAF5EC', bold=True)
    _paint(ws, 5, 9, t, 10, fill='F4FAF6')
    _paint(ws, t, 1, t, 10, fill='DDE6EB', bold=True)
    _paint(ws, t, 7, t, 10, fill='CDE8D3', bold=True, color='1E6B34')
    _paint(ws, t + 2, 1, t + 2, 1, italic=True, color='5D6B74')
    ws.column_dimensions['A'].width = 18
    for c in range(2, 11):
        ws.column_dimensions[L(c)].width = 16
    # 設定
    m = META['設定']; ws = wb['設定']
    _title(ws, '設定', 5)
    _paint(ws, 2, 1, 9, 1, bold=True)
    _paint(ws, 2, 2, 9, 2, border=True, bold=True, halign='right')
    for r, f in ((2, '0.0000'), (3, '0.0000'), (4, '0.0%'), (5, '0.0000'), (6, '#,##0'), (7, '#,##0'), (8, '#,##0'), (9, '#,##0')):
        ws.cell(r, 2).number_format = f
    notes = {2: '野澤、斑尾成本換算用', 3: '收入 × 此係數(刷卡手續費)', 4: '(收入 − 成本) × 稅率', 5: '龍平成本換算用',
             6: '手套:每人一次', 7: '雪衣雪褲:每人每日', 8: '護臀:每人每日', 9: '護膝:每人每日'}
    for r, t_ in notes.items():
        ws.cell(r, 3, t_)
    _paint(ws, 2, 3, 9, 3, italic=True, color='7A8A94', size=10)
    _paint(ws, 11, 1, 12, 1, italic=True, color='7A8A94', size=10)
    _paint(ws, 14, 1, 14, 1, bold=True, size=12, color=REG_DARK['設定'])
    ws.row_dimensions[15].height = 28
    for r in range(m['first'], m['last'] + 1):
        if ws.cell(r, 1).value:
            _paint(ws, r, 1, r, 5, border=True)
            _paint(ws, r, 3, r, 3, fmt='#,##0')
            _paint(ws, r, 4, r, 4, fmt='#,##0', fill='F5F8FA', bold=True)
            _paint(ws, r, 5, r, 5, halign='center', color='5D6B74')
            _paint(ws, r, 1, r, 1, fill=REG_LIGHT[ws.cell(r, 1).value], bold=True)
    for col, w in (('A', 22), ('B', 24), ('C', 26), ('D', 16), ('E', 10)):
        ws.column_dimensions[col].width = w
    # 例外
    ws = wb['例外']
    _title(ws, '例外', 9)
    ws.row_dimensions[3].height = 26
    _paint(ws, 4, 1, ws.max_row, 9, border=True)
    _band(ws, 4, ws.max_row, 1, 9)
    _paint(ws, 4, 9, ws.max_row, 9, fmt=MONEY)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    to_xlsx = '--xlsx' in sys.argv
    src = args[0] if args else sorted(glob.glob(os.path.join(BASE, '原始大總表[0-9][0-9][0-9][0-9].xlsx')))[-1]
    print('讀取', os.path.basename(src))
    d = load(src)
    gsh = None if to_xlsx else open_gsheet()
    inp = old_inputs_gs(gsh) if gsh else old_inputs()
    S = inp['設定']
    fx = S.get('匯率 日幣→台幣', 0.21); fxk = S.get('匯率 韓幣→台幣', 0.024); fee = S.get('刷卡手續費係數', 0.9815); tax = S.get('稅率', 0.05)
    price = {k: S.get(f'租借 {k}', v) for k, v in {'雪衣雪褲': 1000, '護臀': 200, '護膝': 200, '手套': 400}.items()}

    cls = d['團號'].map(classify)
    d['cls'] = cls
    exc = d[cls.isna()]
    d = d[cls.notna()].copy()
    pdays = pure_days(d)
    isp = d['cls'].map(lambda c: c[0] == '純課程')
    d['rdays'] = d['days']
    d.loc[isp, 'rdays'] = d.loc[isp, '訂單編號'].map(lambda o: pdays.get(int(o), 1) if pd.notna(o) else 1)
    unmatched = int((isp & ~d['訂單編號'].map(lambda o: pd.notna(o) and int(o) in pdays)).sum())
    d['rent'] = rental_rev(d, price, d['rdays'])

    wb = Workbook(); wb.remove(wb.active)
    META = {}
    META['設定'] = {}
    ov = wb.create_sheet('總覽')
    st = wb.create_sheet('設定')
    # 設定
    st['A1'] = '設定(黃色格子可修改)'; st['A1'].font = BOLD
    rows = [('匯率 日幣→台幣', fx), ('刷卡手續費係數', fee), ('稅率', tax), ('匯率 韓幣→台幣', fxk),
            ('租借 手套', price['手套']), ('租借 雪衣雪褲', price['雪衣雪褲']), ('租借 護臀', price['護臀']), ('租借 護膝', price['護膝'])]
    for i, (k, v) in enumerate(rows, start=2):
        if k:
            st.cell(i, 1, k); c = st.cell(i, 2, v); c.fill = YEL
    st['A11'] = '租借單價(台幣):手套 每人一次;雪衣雪褲、護臀、護膝 每人每日。全套兩天=400+2×(1000+200+200)=3,200。'
    st['A12'] = '團客租借天數 = 雪票天數;純課程租借天數 = OP備註「共N日」,找不到時以1日計。'
    st.column_dimensions['A'].width = 24
    FX, FEE, TAX, FXK = '設定!$B$2', '設定!$B$3', '設定!$B$4', '設定!$B$5'

    st['A14'] = '成本單價(野澤、斑尾為日幣;龍平為韓幣)— 填一次,含機票/不含機票的分頁都會引用'; st['A14'].font = BOLD
    hdr(st, 15, ['區域', '項目', '單價(外幣)', '單價(台幣)', '幣別'])
    PRICE = {}
    r0 = 16
    for reg in ('野澤', '斑尾', '龍平'):
        for it in CATS[(reg, True)]:
            st.cell(r0, 1, reg); st.cell(r0, 2, it)
            c = st.cell(r0, 3, inp['items'].get((reg, it))); c.fill = YEL
            PRICE[(reg, it)] = f'設定!$C${r0}'
            st.cell(r0, 4, f'=IF(C{r0}="","",C{r0}*{FXK if reg == "龍平" else FX})'); st.cell(r0, 5, '韓幣' if reg == '龍平' else '日幣')
            r0 += 1
        r0 += 1
    META['設定']['first'] = 16; META['設定']['last'] = r0 - 2

    summary = []  # (名稱, sheet, 結果cell地址...)
    for (reg, air), items in CATS.items():
        name = SHEET(reg, air)
        CUR = '韓幣' if reg == '龍平' else '日幣'; FXR = FXK if reg == '龍平' else FX
        ws = wb.create_sheet(name)
        sub = d[d['cls'].map(lambda c: c[0] == '團客' and c[1] == reg and c[2] == air)]
        n = len(items)
        ws['A1'] = f'{name}(團號 {"含A" if air else "無A"})'; ws['A1'].font = Font(bold=True, size=14)
        ws['A2'] = '人數'; ws['B2'] = None  # 後面填公式
        hdr(ws, 4, ['項目', f'單價({CUR},在設定分頁填)', '數量', f'{CUR}總價', '台幣總價', '每人數量', '每人台幣', '數量算法'])
        gstart = 5 + n + 18  # 團別明細表標題列
        g = sub.groupby('團號')
        first = gstart + 1; last = gstart + len(g)
        qcol = {it: 6 + i for i, it in enumerate(items)}  # 團別表各項目欄
        for i, it in enumerate(items):
            r = 5 + i
            ws.cell(r, 1, it)
            ws.cell(r, 2, f'={PRICE[(reg, it)]}')
            col = L(qcol[it])
            ws.cell(r, 3, f'=SUM({col}{first}:{col}{last})' if len(g) else 0)
            ws.cell(r, 4, f'=B{r}*C{r}'); ws.cell(r, 5, f'=D{r}*{FXR}')
            ws.cell(r, 6, f'=IF($B$2=0,0,C{r}/$B$2)'); ws.cell(r, 7, f'=IF($B$2=0,0,E{r}/$B$2)')
            ws.cell(r, 8, RULE_TXT[RULE.get(it, 'pax')])
        tr = 5 + n
        ws.cell(tr, 1, '總成本').font = BOLD
        ws.cell(tr, 4, f'=SUM(D5:D{tr-1})'); ws.cell(tr, 5, f'=SUM(E5:E{tr-1})'); ws.cell(tr, 7, f'=SUM(G5:G{tr-1})')
        ws['B2'] = f'=SUM(C{first}:C{last})' if len(g) else 0
        # 結果區
        rr = tr + 2
        hdr(ws, rr, ['結果(台幣)', '版本一(不計租借收入)', '版本二(含租借收入)'])
        lab = ['總收入(原始團費)', '租借收入(依租借欄位估算)', '計入收入', '刷卡後收入(×係數)', '總成本', '稅 (計入收入−總成本)×稅率', '最後結果',
               '每人 計入收入', '每人 總成本', '每人 結果']
        R = {k: rr + 1 + i for i, k in enumerate(lab)}
        for k, r in R.items():
            ws.cell(r, 1, k)
        tot_rev = f'SUM(D{first}:D{last})' if len(g) else '0'; tot_rent = f'SUM(E{first}:E{last})' if len(g) else '0'
        for col in 'BC':
            ws[f'{col}{R["總收入(原始團費)"]}'] = f'={tot_rev}'
            ws[f'{col}{R["租借收入(依租借欄位估算)"]}'] = f'={tot_rent}'
            ws[f'{col}{R["計入收入"]}'] = f'={col}{R["總收入(原始團費)"]}' + (f'-{col}{R["租借收入(依租借欄位估算)"]}' if col == 'B' else '')
            ws[f'{col}{R["刷卡後收入(×係數)"]}'] = f'={col}{R["計入收入"]}*{FEE}'
            ws[f'{col}{R["總成本"]}'] = f'=$E${tr}'
            ws[f'{col}{R["稅 (計入收入−總成本)×稅率"]}'] = f'=({col}{R["計入收入"]}-{col}{R["總成本"]})*{TAX}'
            ws[f'{col}{R["最後結果"]}'] = f'={col}{R["刷卡後收入(×係數)"]}-{col}{R["總成本"]}-{col}{R["稅 (計入收入−總成本)×稅率"]}'
            ws[f'{col}{R["每人 計入收入"]}'] = f'=IF($B$2=0,0,{col}{R["計入收入"]}/$B$2)'
            ws[f'{col}{R["每人 總成本"]}'] = f'=IF($B$2=0,0,{col}{R["總成本"]}/$B$2)'
            ws[f'{col}{R["每人 結果"]}'] = f'=IF($B$2=0,0,{col}{R["最後結果"]}/$B$2)'
        for c in ('A', 'B', 'C'):
            ws[f'{c}{R["最後結果"]}'].font = BOLD; ws[f'{c}{R["最後結果"]}'].fill = GRY
        # 團別明細
        ws.cell(gstart - 1, 1, '團別明細(數量來自原始資料;成本用上方單價計算)').font = BOLD
        hdr(ws, gstart, ['團號', '出發日', '人數', '總收入', '租借收入(估)'] + items +
            [f'{CUR}成本', '台幣成本', '版本一結果', '版本二結果', '每人 版本一', '每人 版本二'])
        k = len(items); cj, ct, c1, c2, p1, p2 = [6 + k + i for i in range(6)]
        for i, (gn, x) in enumerate(g):
            r = first + i
            ws.cell(r, 1, gn); ws.cell(r, 2, x['出發日期'].iloc[0]); ws.cell(r, 3, len(x))
            ws.cell(r, 4, float(x['fee'].sum())); ws.cell(r, 5, float(x['rent'].sum()))
            for it in items:
                rule = RULE.get(it, 'pax')
                v = {'lodge': x['lodge'].sum(), 'city': x['city'].sum(), 'days': x['days'].sum()}.get(rule, len(x))
                ws.cell(r, qcol[it], float(v))
            ws.cell(r, cj, '=' + '+'.join(f'{L(qcol[it])}{r}*$B${5+j}' for j, it in enumerate(items)))
            ws.cell(r, ct, f'={L(cj)}{r}*{FXR}')
            for cc, rev in ((c1, f'(D{r}-E{r})'), (c2, f'D{r}')):
                ws.cell(r, cc, f'={rev}*{FEE}-{L(ct)}{r}-({rev}-{L(ct)}{r})*{TAX}')
            ws.cell(r, p1, f'={L(c1)}{r}/C{r}'); ws.cell(r, p2, f'={L(c2)}{r}/C{r}')
        for cidx in range(1, 40):
            ws.column_dimensions[L(cidx)].width = 14
        ws.column_dimensions['A'].width = 30; ws.column_dimensions['H'].width = 14
        for r in range(5, rr + 12):
            for cidx in range(2, 8):
                ws.cell(r, cidx).number_format = '#,##0'
        for r in range(first, last + 1):
            for cidx in range(4, p2 + 1):
                ws.cell(r, cidx).number_format = '#,##0'
        ws.freeze_panes = 'A5'
        META[name] = dict(reg=reg, n=n, tr=tr, rr=rr, R=R, gs=gstart, first=first, last=last, last_col=p2)
        summary.append((name, R, len(sub), sub['團號'].nunique()))

    # 純課程
    wp = wb.create_sheet('純課程')
    pc = d[d['cls'].map(lambda c: c[0] == '純課程')]
    wp['A1'] = '純課程(只扣刷卡手續費,無成本項目)'; wp['A1'].font = Font(bold=True, size=14)
    hdr(wp, 3, ['區域', '人數', '總收入', '租借收入(估)', '版本一(扣租借×係數)', '版本二(×係數)', '每人 版本一', '每人 版本二'])
    gs = 11
    g = pc.groupby('團號'); first = gs + 1; last = gs + len(g)
    for i, rg in enumerate(['湯澤', '野澤', '斑尾']):
        r = 4 + i
        wp.cell(r, 1, rg)
        for c, col in ((2, 'D'), (3, 'E'), (4, 'F')):
            wp.cell(r, c, f'=SUMIF($B${first}:$B${last},$A{r},{col}${first}:{col}${last})')
        wp.cell(r, 5, f'=(C{r}-D{r})*{FEE}'); wp.cell(r, 6, f'=C{r}*{FEE}')
        wp.cell(r, 7, f'=IF(B{r}=0,0,E{r}/B{r})'); wp.cell(r, 8, f'=IF(B{r}=0,0,F{r}/B{r})')
    wp.cell(7, 1, '合計').font = BOLD
    for c in range(2, 7):
        wp.cell(7, c, f'=SUM({L(c)}4:{L(c)}6)')
    wp.cell(7, 7, '=IF(B7=0,0,E7/B7)'); wp.cell(7, 8, '=IF(B7=0,0,F7/B7)')
    wp.cell(gs - 1, 1, '團別明細').font = BOLD
    hdr(wp, gs, ['團號', '區域', '出發日', '人數', '總收入', '租借收入(估)', '版本一', '版本二'])
    for i, (gn, x) in enumerate(g):
        r = first + i
        wp.cell(r, 1, gn); wp.cell(r, 2, x['cls'].iloc[0][1]); wp.cell(r, 3, x['出發日期'].iloc[0]); wp.cell(r, 4, len(x))
        wp.cell(r, 5, float(x['fee'].sum())); wp.cell(r, 6, float(x['rent'].sum()))
        wp.cell(r, 7, f'=(E{r}-F{r})*{FEE}'); wp.cell(r, 8, f'=E{r}*{FEE}')
    # 純課程明細 人數欄對應: 區域彙總用 D(人數)E(收入)F(租借) → 調整 SUMIF 欄位
    for i in range(3):
        r = 4 + i
        for c, col in ((2, 'D'), (3, 'E'), (4, 'F')):
            wp.cell(r, c, f'=SUMIF($B${first}:$B${last},$A{r},{col}${first}:{col}${last})')
    for r in range(4, last + 1):
        for c in range(2, 9):
            wp.cell(r, c).number_format = '#,##0'
    for c in range(1, 9):
        wp.column_dimensions[L(c)].width = 18
    META['純課程'] = dict(gs=gs, first=first, last=last)

    # 例外
    we = wb.create_sheet('例外')
    we['A1'] = '無法辨識或客製團(NC、DBC-NC 等),保留原始資料,不計算'; we['A1'].font = BOLD
    cols = ['團號', '出發日期', '報名日期', '訂單編號', '旅客編號', '中文姓名', 'LEVEL', '滑雪類別', '團費']
    hdr(we, 3, cols)
    for i, row in enumerate(exc[cols].itertuples(index=False), start=4):
        for j, v in enumerate(row, start=1):
            we.cell(i, j, None if pd.isna(v) else v)
    for c in range(1, 10):
        we.column_dimensions[L(c)].width = 16

    # 總覽
    ov['A1'] = '售價與成本總覽(台幣)'; ov['A1'].font = Font(bold=True, size=14)
    ov['A2'] = f'資料來源:{os.path.basename(src)}。成本單價請到各分類分頁的黃色格子填寫,匯率與租借價格在「設定」。'
    hdr(ov, 4, ['分類', '團數', '人數', '計入收入 版一', '計入收入 版二', '總成本', '結果 版一', '結果 版二', '每人結果 版一', '每人結果 版二'])
    r = 5
    for name, R, npax, ngrp in summary:
        q = f"'{name}'!"
        ov.cell(r, 1, name); ov.cell(r, 2, ngrp); ov.cell(r, 3, f'={q}B2')
        ov.cell(r, 4, f'={q}B{R["計入收入"]}'); ov.cell(r, 5, f'={q}C{R["計入收入"]}'); ov.cell(r, 6, f'={q}B{R["總成本"]}')
        ov.cell(r, 7, f'={q}B{R["最後結果"]}'); ov.cell(r, 8, f'={q}C{R["最後結果"]}')
        ov.cell(r, 9, f'=IF(C{r}=0,0,G{r}/C{r})'); ov.cell(r, 10, f'=IF(C{r}=0,0,H{r}/C{r})')
        r += 1
    ov.cell(r, 1, '純課程'); ov.cell(r, 2, int(pc['團號'].nunique())); ov.cell(r, 3, "='純課程'!B7")
    ov.cell(r, 4, "='純課程'!C7-'純課程'!D7"); ov.cell(r, 5, "='純課程'!C7"); ov.cell(r, 6, 0)
    ov.cell(r, 7, "='純課程'!E7"); ov.cell(r, 8, "='純課程'!F7")
    ov.cell(r, 9, f'=IF(C{r}=0,0,G{r}/C{r})'); ov.cell(r, 10, f'=IF(C{r}=0,0,H{r}/C{r})')
    r += 1
    ov.cell(r, 1, '合計').font = BOLD
    for c in range(2, 9):
        ov.cell(r, c, f'=SUM({L(c)}5:{L(c)}{r-1})')
    ov.cell(r, 9, f'=IF(C{r}=0,0,G{r}/C{r})'); ov.cell(r, 10, f'=IF(C{r}=0,0,H{r}/C{r})')
    for rr_ in range(5, r + 1):
        for c in range(2, 11):
            ov.cell(rr_, c).number_format = '#,##0'
    ov.cell(r + 2, 1, f'例外(不計算)共 {len(exc)} 列、{exc["團號"].nunique()} 團,見「例外」分頁。純課程租借天數找不到 OP備註的有 {unmatched} 人(以1日計)。')
    ov.column_dimensions['A'].width = 16
    for c in range(2, 11):
        ov.column_dimensions[L(c)].width = 15
    META['總覽'] = dict(total=r, rows=[x[0] for x in summary] + ['純課程'])
    wb.move_sheet('總覽', offset=-wb.index(wb['總覽']))
    beautify(wb, META)
    if to_xlsx:
        wb.save(OUT); print('完成', OUT)
    else:
        push_gsheet(gsh, wb)
    print(d.groupby(d['cls'].map(lambda c: c if c[0] == '純課程' else c)).agg(團=('團號', 'nunique'), 人=('團號', 'size'), 收入=('fee', 'sum'), 租借=('rent', 'sum')).round(0).to_string())
    print('例外', len(exc), exc['團號'].unique().tolist()[:12], '| 純課程找不到天數', unmatched)


if __name__ == '__main__':
    main()
