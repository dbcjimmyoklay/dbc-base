"""
agents/reporter.py
REPORTER 報表員
─────────────────────────────────────────────
職責（原 Worker 5 + Worker 6 合併）：
  報表1 │ 每日教練需求報表 → dbc_coach_report.html
        │  • 依出發日期計算各雪場每天需要幾位教練
        │  • 純課程：每訂單1位教練 × 天數
        │  • 團客：按 Level 分組，每組÷6無條件進位
        │  • 不含教練的續滑天數不計入
        │  • 顯示：列表視圖 + 日曆視圖

  報表2 │ 銷量分析報表 → dbc_sales_report.html
        │  • 從科威原始資料統計
        │  • 總覽、報名趨勢、出發分布、客群分析、加購統計
─────────────────────────────────────────────
"""

import os
import sys
import json
import re
import math
from datetime import datetime, timedelta
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from shared.config import (
    PURE_ITEMS, COACH_CAPS, COACH_THRESHOLDS,
    EQUIP_PRICE,
)
from shared.utils import normalize_level

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ═══════════════════════════════════════════
# 報表1：每日教練需求
# ═══════════════════════════════════════════

AREAS = ['野澤', '斑尾', '湯澤', '龍平']
AREA_MAP = {
    '野澤純課': '野澤', '斑尾純課': '斑尾', '湯澤純課': '湯澤',
    '野澤': '野澤', '斑尾': '斑尾', '龍平': '龍平',
}
WEEKDAY_ZH = ['一', '二', '三', '四', '五', '六', '日']


def _lv_group(lv):
    """Level → 分組（g0 / g12 / g345）"""
    if not lv or lv == 'lv0':
        return 'g0'
    m = re.search(r'(\d+)', lv)
    n = int(m.group(1)) if m else 0
    return 'g12' if n <= 2 else 'g345'


def _parse_coach_days(tiansu):
    """
    從天數欄解析含教練的上課天數
    「無」= 不含教練續滑 → 教練只算基本2天
    「含」= 含教練續滑   → 用總天數
    純課程               → 直接用天數
    """
    if not tiansu or str(tiansu).strip() == '':
        return 2
    s = str(tiansu).strip()
    m = re.search(r'(\d+)', s)
    if not m:
        return 2
    n = int(m.group(1))
    if '無' in s and '含' not in s:
        return 2          # 不含教練的續滑，教練只上基本2天
    return n


def _coach_color(area, n):
    """教練人數 → 警示顏色"""
    if n == 0:
        return 'none'
    for limit, color in COACH_THRESHOLDS[area]:
        if n <= limit:
            return color
    return 'red'


def _calc_daily_coach(all_records_flat, DATE_START, DATE_END):
    """計算雪季每天每個雪場的教練需求"""
    # daily[date_str][area] = {'pure_orders': set, 'g0': int, 'g12': int, 'g345': int}
    daily = defaultdict(lambda: defaultdict(
        lambda: {'pure_orders': set(), 'g0': 0, 'g12': 0, 'g345': 0}
    ))

    for r in all_records_flat:
        item = r.get('項目', '')
        if not item or item == '其他':
            continue
        area = AREA_MAP.get(item, '')
        if not area:
            continue

        dep_str = r.get('出發日期', '')
        if not dep_str:
            continue
        try:
            parts = dep_str.split('/')
            dep = datetime(2000 + int(parts[0]), int(parts[1]), int(parts[2]))
        except Exception:
            continue

        ono    = str(r.get('訂編', '')).strip()
        lv     = r.get('LEVEL', '') or 'lv0'
        grp    = _lv_group(lv)
        tiansu = str(r.get('天數', '')).strip()

        if item in PURE_ITEMS:
            days = _parse_coach_days(tiansu)
            if days == 0:
                days = 2
            for d in range(1, days + 1):
                ds = (dep + timedelta(days=d)).strftime('%Y/%m/%d')
                daily[ds][area]['pure_orders'].add(ono)
        else:
            days = _parse_coach_days(tiansu)
            if days == 0:
                days = 2
            for d in range(1, days + 1):
                ds = (dep + timedelta(days=d)).strftime('%Y/%m/%d')
                daily[ds][area][grp] += 1

    return daily


def _calc_coaches(d):
    """單一(日期,雪場)資料 → 教練人數、分類"""
    g0    = d.get('g0',  0)
    g12   = d.get('g12', 0)
    g345  = d.get('g345', 0)
    pure  = len(d.get('pure_orders', set()))
    team  = (
        (math.ceil(g0  / 6) if g0   > 0 else 0) +
        (math.ceil(g12 / 6) if g12  > 0 else 0) +
        (math.ceil(g345/ 6) if g345 > 0 else 0)
    )
    return pure + team, pure, team, g0, g12, g345


def _build_coach_json(all_records_flat, DATE_START, DATE_END):
    """建立教練需求 JSON 資料（供 HTML 使用）"""
    daily = _calc_daily_coach(all_records_flat, DATE_START, DATE_END)
    rows  = []
    cur   = DATE_START

    while cur <= DATE_END:
        ds  = cur.strftime('%Y/%m/%d')
        wd  = WEEKDAY_ZH[cur.weekday()]
        row = {'date': ds, 'wd': wd, 'areas': {}}
        total_jp = total_kr = total_all = 0

        for area in AREAS:
            d = daily[ds].get(area, {})
            coaches, pure, team, g0, g12, g345 = _calc_coaches(d)
            color = _coach_color(area, coaches)
            cap   = COACH_CAPS[area]
            row['areas'][area] = {
                'c': coaches, 'pure': pure, 'team': team,
                'g0': g0, 'g12': g12, 'g345': g345,
                'color': color, 'cap': cap,
            }
            if area in ('野澤', '斑尾', '湯澤'):
                total_jp += coaches
            else:
                total_kr += coaches
            total_all += coaches

        row['total_jp'] = total_jp
        row['total_kr'] = total_kr
        row['total']    = total_all
        rows.append(row)
        cur += timedelta(days=1)

    return json.dumps(rows, ensure_ascii=False)


COACH_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>DBC 每日教練需求</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700&family=DM+Mono:wght@400;500&display=swap');
:root{--bg:#0a0f1e;--s1:#111827;--s2:#1a2235;--bd:#1e2d45;--tx:#e2e8f0;--mu:#64748b;
  --ny:#ffffff;--nb:#7ec8f5;--nt:#ffb347;--nl:#c8864a;
  --green:#10b981;--yellow:#f59e0b;--orange:#f97316;--red:#ef4444;--none:#334155;}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Noto Sans TC',sans-serif;font-size:15px}
.hd{background:linear-gradient(135deg,#0f172a,#1e3a5f 50%,#0f172a);border-bottom:1px solid var(--bd);padding:14px 20px;position:sticky;top:0;z-index:200}
.hd-in{max-width:1400px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap}
.logo{display:flex;align-items:center;gap:10px}
.logo-i{width:30px;height:30px;background:#3b82f6;border-radius:7px;display:flex;align-items:center;justify-content:center}
.logo-t h1{font-size:15px;font-weight:700}.logo-t p{font-size:10px;color:var(--mu)}
.badge{background:var(--s2);border:1px solid var(--bd);border-radius:20px;padding:4px 12px;font-size:11px;color:#06b6d4;font-family:'DM Mono',monospace}
.main{max-width:1400px;margin:0 auto;padding:16px}
.view-toggle{display:flex;gap:6px;margin-bottom:14px}
.vt-btn{padding:7px 20px;border-radius:7px;font-size:13px;cursor:pointer;border:1px solid var(--bd);color:var(--mu);background:transparent;transition:all .15s;font-family:'Noto Sans TC',sans-serif}
.vt-btn.active{background:#3b82f6;border-color:#3b82f6;color:#fff;font-weight:500}
.legend{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:14px;padding:12px 16px;background:var(--s1);border:1px solid var(--bd);border-radius:8px;align-items:center}
.cal-wrap{display:flex;flex-direction:column;gap:24px}
.cal-month{background:var(--s1);border:1px solid var(--bd);border-radius:10px;overflow:hidden}
.cal-month-title{padding:12px 16px;font-size:15px;font-weight:700;background:var(--s2);border-bottom:1px solid var(--bd);color:#3b82f6}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);border-left:1px solid var(--bd)}
.cal-dow{padding:8px 4px;text-align:center;font-size:11px;color:var(--mu);border-bottom:1px solid var(--bd);border-right:1px solid var(--bd);background:var(--s2)}
.cal-day{border-right:1px solid var(--bd);border-bottom:1px solid var(--bd);min-height:90px;padding:6px;position:relative}
.cal-day.empty{background:rgba(0,0,0,.15)}.cal-day.weekend{background:rgba(245,158,11,.03)}.cal-day.has-data{background:rgba(30,45,69,.4)}
.cal-date{font-size:12px;color:var(--mu);margin-bottom:6px;font-family:'DM Mono',monospace}
.cal-badges{display:flex;flex-direction:column;gap:3px}
.cal-badge{display:flex;align-items:center;gap:4px;font-size:11px;border-radius:4px;padding:2px 5px}
.cal-badge-dot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
.cal-badge-n{font-family:'DM Mono',monospace;font-weight:700;font-size:12px}
.cal-badge.bg-green{background:rgba(16,185,129,.1)}.cal-badge.bg-yellow{background:rgba(245,158,11,.1)}
.cal-badge.bg-orange{background:rgba(249,115,22,.1)}.cal-badge.bg-red{background:rgba(239,68,68,.1)}
.legend-item{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--mu)}
.leg-dot{width:10px;height:10px;border-radius:50%}
.filters{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
.filter-btn{padding:6px 16px;border-radius:6px;font-size:12px;cursor:pointer;border:1px solid var(--bd);color:var(--mu);background:transparent;transition:all .15s}
.filter-btn.active{border-color:#3b82f6;color:#3b82f6;background:rgba(59,130,246,.1)}
.tbl-wrap{border-radius:10px;border:1px solid var(--bd);overflow:auto;max-height:calc(100vh - 180px)}
table{width:100%;border-collapse:collapse;min-width:600px}
thead{position:sticky;top:0;z-index:100}
thead th{background:#0d1829;padding:12px 16px;text-align:center;font-size:13px;color:var(--mu);font-weight:500;border-bottom:1px solid var(--bd);white-space:nowrap}
thead tr:first-child th{border-bottom:none}thead tr:last-child th{border-bottom:2px solid var(--bd);font-size:12px}
tbody tr{border-bottom:1px solid var(--bd);transition:background .1s}
tbody tr:hover{background:var(--s2)}tbody tr.has-data{background:rgba(30,45,69,.3)}
td{padding:12px 16px;text-align:center;font-family:'DM Mono',monospace;font-size:17px}
td.date-col{text-align:left;font-family:'Noto Sans TC',sans-serif;color:var(--mu);font-size:14px;white-space:nowrap}
td.wd-col{font-size:14px;color:var(--mu);padding:12px 6px}
.coach-badge{display:inline-flex;align-items:center;justify-content:center;min-width:48px;height:40px;border-radius:8px;font-weight:700;font-size:22px;padding:0 12px}
.cb-green{background:rgba(16,185,129,.15);color:var(--green)}.cb-yellow{background:rgba(245,158,11,.15);color:var(--yellow)}
.cb-orange{background:rgba(249,115,22,.15);color:var(--orange)}.cb-red{background:rgba(239,68,68,.15);color:var(--red)}
.cb-none{color:var(--mu);opacity:.3;font-size:16px}
.month-header td{background:rgba(15,23,42,.95);color:#3b82f6;font-size:13px;font-weight:600;text-align:left;padding:8px 16px;border-bottom:1px solid #1e3a5f}
</style></head><body>
<div class="hd"><div class="hd-in">
  <div class="logo"><div class="logo-i">👨‍🏫</div><div class="logo-t"><h1>每日教練需求</h1><p>DAILY COACH REQUIREMENT</p></div></div>
  <div class="badge" id="upd">—</div>
</div></div>
<div class="main">
<div class="view-toggle">
  <button class="vt-btn active" onclick="setView('table',this)">📋 列表</button>
  <button class="vt-btn" onclick="setView('calendar',this)">📅 日曆</button>
</div>
<div id="table-view">
<div class="legend">
  <span style="font-size:11px;color:var(--mu);margin-right:4px">教練人數警示：</span>
  <div class="legend-item"><div class="leg-dot" style="background:var(--green)"></div>充足</div>
  <div class="legend-item"><div class="leg-dot" style="background:var(--yellow)"></div>注意</div>
  <div class="legend-item"><div class="leg-dot" style="background:var(--orange)"></div>緊張</div>
  <div class="legend-item"><div class="leg-dot" style="background:var(--red)"></div>超載</div>
  <div style="margin-left:auto;font-size:10px;color:var(--mu)">上限：野澤14 斑尾8 湯澤10 龍平10</div>
</div>
<div class="filters">
  <button class="filter-btn active" onclick="setFilter('all',this)">全部</button>
  <button class="filter-btn" onclick="setFilter('has',this)">有課日</button>
  <button class="filter-btn" onclick="setFilter('warn',this)">注意/緊張/超載</button>
  <button class="filter-btn" onclick="setFilter('nozawa',this)">野澤</button>
  <button class="filter-btn" onclick="setFilter('madarao',this)">斑尾</button>
  <button class="filter-btn" onclick="setFilter('yuzawa',this)">湯澤</button>
  <button class="filter-btn" onclick="setFilter('ryongpyong',this)">龍平</button>
</div>
<div class="tbl-wrap"><table id="coach-table">
<thead>
  <tr>
    <th rowspan="2" style="text-align:left">日期</th>
    <th rowspan="2" style="padding:10px 4px">曜</th>
    <th style="border-bottom:2px solid var(--ny)">野澤</th>
    <th style="border-bottom:2px solid var(--nb)">斑尾</th>
    <th style="border-bottom:2px solid var(--nt)">湯澤</th>
    <th style="border-bottom:2px solid var(--nl)">龍平</th>
    <th rowspan="2">日本<br><span style="font-size:9px;color:var(--mu)">(野+斑+湯)</span></th>
    <th rowspan="2">韓國<br><span style="font-size:9px;color:var(--mu)">(龍平)</span></th>
    <th rowspan="2">合計</th>
  </tr>
  <tr><th></th><th></th><th></th><th></th></tr>
</thead>
<tbody id="tbody"></tbody>
</table></div>
</div>
<div id="cal-view" style="display:none">
  <div class="legend">
    <span style="font-size:11px;color:var(--mu);margin-right:4px">教練人數警示：</span>
    <div class="legend-item"><div class="leg-dot" style="background:var(--green)"></div>充足</div>
    <div class="legend-item"><div class="leg-dot" style="background:var(--yellow)"></div>注意</div>
    <div class="legend-item"><div class="leg-dot" style="background:var(--orange)"></div>緊張</div>
    <div class="legend-item"><div class="leg-dot" style="background:var(--red)"></div>超載</div>
  </div>
  <div class="cal-wrap" id="cal-wrap"></div>
</div>
</div>
<script>
const D=COACH_DATA_PLACEHOLDER;
const AREAS=['野澤','斑尾','湯澤','龍平'];
document.getElementById('upd').textContent='更新：'+new Date().toLocaleDateString('zh-TW');
function badge(n,color){if(n===0)return`<div class="coach-badge cb-none">—</div>`;return`<div class="coach-badge cb-${color}">${n}</div>`;}
let currentFilter='all';
function setFilter(f,btn){currentFilter=f;document.querySelectorAll('.filter-btn').forEach(b=>b.classList.remove('active'));btn.classList.add('active');renderTable();}
function shouldShow(row){
  if(currentFilter==='all')return true;if(currentFilter==='has')return row.total>0;
  if(currentFilter==='warn')return AREAS.some(a=>{const c=row.areas[a]?.color;return c==='yellow'||c==='orange'||c==='red';});
  const aMap={nozawa:'野澤',madarao:'斑尾',yuzawa:'湯澤',ryongpyong:'龍平'};
  const area=aMap[currentFilter];return area&&(row.areas[area]?.c||0)>0;
}
function renderTable(){
  const tbody=document.getElementById('tbody');tbody.innerHTML='';let prevMonth='';
  D.forEach(row=>{
    if(!shouldShow(row))return;
    const month=row.date.slice(0,7);
    if(month!==prevMonth){const mtr=document.createElement('tr');mtr.className='month-header';mtr.innerHTML=`<td colspan="9">${month.slice(2).replace('-','/')}</td>`;tbody.appendChild(mtr);prevMonth=month;}
    const tr=document.createElement('tr');tr.className=(row.total>0?'has-data ':'')+(row.wd==='六'||row.wd==='日'?'weekend':'');
    let cols='';AREAS.forEach(a=>{const d=row.areas[a]||{c:0,color:'none'};cols+=`<td>${badge(d.c,d.color)}</td>`;});
    const jpC=row.total_jp>25?'red':row.total_jp>20?'orange':row.total_jp>15?'yellow':'';
    tr.innerHTML=`<td class="date-col">${row.date.slice(5)}</td><td class="wd-col" style="color:${row.wd==='六'||row.wd==='日'?'#f59e0b':'var(--mu)'}">${row.wd}</td>${cols}
    <td style="font-size:16px;font-weight:700;color:${jpC?'var(--'+jpC+')':'#06b6d4'}">${row.total_jp||'—'}</td>
    <td style="font-size:16px;font-weight:700;color:${row.total_kr>0?'#8b5cf6':'var(--mu)'}">${row.total_kr||'—'}</td>
    <td style="font-size:18px;font-weight:700">${row.total||'—'}</td>`;
    tbody.appendChild(tr);
  });
}
renderTable();
function setView(v,btn){document.querySelectorAll('.vt-btn').forEach(b=>b.classList.remove('active'));btn.classList.add('active');document.getElementById('table-view').style.display=v==='table'?'':'none';document.getElementById('cal-view').style.display=v==='calendar'?'':'none';if(v==='calendar'&&!calRendered)renderCalendar();}
let calRendered=false;const DOW=['日','一','二','三','四','五','六'];
const COLOR_MAP={'green':'var(--green)','yellow':'var(--yellow)','orange':'var(--orange)','red':'var(--red)','none':'var(--mu)'};
function renderCalendar(){
  calRendered=true;const wrap=document.getElementById('cal-wrap');wrap.innerHTML='';
  const months={};D.forEach(row=>{const m=row.date.replace(/\//g,'-').slice(0,7);if(!months[m])months[m]=[];months[m].push(row);});
  Object.entries(months).forEach(([month,rows])=>{
    const[y,mo]=month.split('-').map(Number);const firstDay=new Date(y,mo-1,1).getDay();const dim=new Date(y,mo,0).getDate();
    const dayMap={};rows.forEach(r=>{dayMap[parseInt(r.date.replace(/\//g,'-').slice(8))]=r;});
    const div=document.createElement('div');div.className='cal-month';
    div.innerHTML=`<div class="cal-month-title">${String(y).slice(2)}/${String(mo).padStart(2,'0')}</div>
<div class="cal-grid">${DOW.map(d=>`<div class="cal-dow">${d}</div>`).join('')}${Array(firstDay).fill('<div class="cal-day empty"></div>').join('')}
${Array.from({length:dim},(_,i)=>{const day=i+1,row=dayMap[day],date=new Date(y,mo-1,day),dow=date.getDay(),isW=dow===0||dow===6,hasD=row&&row.total>0;
let badges='';if(row){AREAS.forEach(a=>{const d=row.areas[a];if(!d||d.c===0)return;const c=d.color;badges+=`<div class="cal-badge bg-${c}"><div class="cal-badge-dot" style="background:${COLOR_MAP[c]}"></div><span style="color:var(--mu);font-size:10px;flex-shrink:0">${a.slice(0,1)}</span><span class="cal-badge-n" style="color:${COLOR_MAP[c]}">${d.c}</span></div>`;});}
return`<div class="cal-day${isW?' weekend':''}${hasD?' has-data':''}"><div class="cal-date">${day}</div><div class="cal-badges">${badges}</div></div>`;}).join('')}
</div>`;wrap.appendChild(div);});
}
</script></body></html>"""


# ═══════════════════════════════════════════
# 報表2：銷量分析
# ═══════════════════════════════════════════

def _build_sales_json(df_records, all_records_flat, headers):
    """建立銷量分析 JSON 資料"""
    df = pd.DataFrame(df_records)

    def _get_area(g):
        if pd.isna(g): return '其他'
        s = str(g)
        if 'ZN' in s: return '野澤'
        if 'ZB' in s: return '斑尾'
        if 'ZY' in s: return '湯澤'
        if 'YS' in s or 'NT' in s: return '龍平'
        if 'BT' in s: return '斑尾'
        if 'N' in s: return '野澤'
        if 'B' in s: return '斑尾'
        return '其他'

    df['_area']      = df['團號'].apply(_get_area)
    df['_lv']        = df['LEVEL'].apply(lambda x: normalize_level(str(x) if pd.notna(x) else ''))
    df['_signup_dt'] = pd.to_datetime(df['報名日期'], errors='coerce')
    df['_fee']       = pd.to_numeric(
        df['團費'].astype(str).str.replace(',', '').str.replace(' ', ''),
        errors='coerce',
    )
    df['_age'] = pd.to_numeric(df['年齡'], errors='coerce')

    # 各區人次（從 all_records_flat，與課程安排一致）
    area_type = {}
    for r in all_records_flat:
        item = r.get('項目', '')
        if '野澤' in item:   area = '野澤'
        elif '斑尾' in item: area = '斑尾'
        elif '湯澤' in item: area = '湯澤'
        elif '龍平' in item: area = '龍平'
        else: continue
        is_pure = '純課' in item
        if area not in area_type:
            area_type[area] = {'pure': 0, 'group': 0, 'total': 0}
        area_type[area]['pure'  if is_pure else 'group'] += 1
        area_type[area]['total'] += 1

    # 每日報名（各區）
    daily_area = df.groupby([df['_signup_dt'].dt.date, '_area']).size().unstack(fill_value=0)
    daily_list = []
    for d in sorted(daily_area.index):
        if not pd.notna(d): continue
        row = {'d': str(d)}
        for a in AREAS:
            row[a] = int(daily_area.loc[d, a]) if a in daily_area.columns else 0
        row['c'] = sum(row.get(a, 0) for a in AREAS)
        daily_list.append(row)

    # 累積
    cum = df.groupby(df['_signup_dt'].dt.date).size().sort_index().cumsum()
    cum_list = [{'d': str(d), 'c': int(c)} for d, c in cum.items() if pd.notna(d)]

    # 每週
    df['_week'] = df['_signup_dt'].dt.to_period('W').astype(str)
    weekly_area = df.groupby(['_week', '_area']).size().unstack(fill_value=0)
    weekly_list = []
    for w in sorted(weekly_area.index):
        if w in ('NaT', 'nan'): continue
        row = {'w': w}
        for a in AREAS:
            row[a] = int(weekly_area.loc[w, a]) if a in weekly_area.columns else 0
        row['c'] = sum(row.get(a, 0) for a in AREAS)
        weekly_list.append(row)

    # 出發月×區域
    area_depart = {}
    for r in all_records_flat:
        dep = r.get('出發日期', '')
        if not dep or len(dep) < 5: continue
        mon  = '20' + dep[:2] + '-' + dep[3:5]
        item = r.get('項目', '')
        area = ('野澤' if '野澤' in item else '斑尾' if '斑尾' in item
                else '湯澤' if '湯澤' in item else '龍平' if '龍平' in item else '')
        if not area: continue
        if mon not in area_depart: area_depart[mon] = {}
        area_depart[mon][area] = area_depart[mon].get(area, 0) + 1

    months_sorted = sorted(area_depart.keys())
    stacked_vals  = [[area_depart.get(m, {}).get(a, 0) for a in AREAS] for m in months_sorted]

    # Level 分布
    level_dist = {k: int(v) for k, v in df['_lv'].value_counts().items() if k}

    # 各區 Level
    area_level = {a: {} for a in AREAS}
    for r in all_records_flat:
        item = r.get('項目', '')
        area = ('野澤' if '野澤' in item else '斑尾' if '斑尾' in item
                else '湯澤' if '湯澤' in item else '龍平' if '龍平' in item else '')
        if not area: continue
        lv = normalize_level(r.get('LEVEL', ''))
        if not lv: continue
        area_level[area][lv] = area_level[area].get(lv, 0) + 1

    # 加購統計
    def _hv(row, cols):
        for c in cols:
            if c in row.index:
                v = row[c]
                if pd.notna(v) and v != 0 and str(v).strip() not in ('', '0'): return True
        return False

    _with    = [c for c in headers if ('續滑' in str(c) or '加滑' in str(c)) and '含' in str(c) and '不含' not in str(c) and '保留' not in str(c)]
    _without = [c for c in headers if ('續滑' in str(c) or '加滑' in str(c)) and '不含' in str(c) and '保留' not in str(c)]
    _excl    = [c for c in headers if '專屬' in str(c)]
    _rc      = [c for c in headers if '租借' in str(c) and '雪衣' in str(c)]
    _rk      = [c for c in headers if '租借' in str(c) and '護膝' in str(c)]
    _rh      = [c for c in headers if '租借' in str(c) and '護臀' in str(c)]
    _rg      = [c for c in headers if '租借' in str(c) and '手套' in str(c)]

    addons = {
        '續滑含教練': int(df.apply(lambda r: _hv(r, _with),    axis=1).sum()),
        '續滑不含'  : int(df.apply(lambda r: _hv(r, _without), axis=1).sum()),
        '專屬教練'  : int(df.apply(lambda r: _hv(r, _excl),    axis=1).sum()),
        '租借衣褲'  : int(df.apply(lambda r: _hv(r, _rc),      axis=1).sum()),
        '租借護膝'  : int(df.apply(lambda r: _hv(r, _rk),      axis=1).sum()),
        '租借護臀'  : int(df.apply(lambda r: _hv(r, _rh),      axis=1).sum()),
        '租借手套'  : int(df.apply(lambda r: _hv(r, _rg),      axis=1).sum()),
    }

    # 年齡分布
    age_bins   = [0, 18, 25, 30, 35, 40, 50, 100]
    age_labels = ['18以下', '19-25', '26-30', '31-35', '36-40', '41-50', '51+']
    age_cut    = pd.cut(df['_age'], bins=age_bins, labels=age_labels, right=True)
    age_dist   = {str(k): int(v) for k, v in age_cut.value_counts().sort_index().items()}

    area_from_records = {}
    for r in all_records_flat:
        item = r.get('項目', '')
        area = ('野澤' if '野澤' in item else '斑尾' if '斑尾' in item
                else '湯澤' if '湯澤' in item else '龍平' if '龍平' in item else '其他')
        area_from_records[area] = area_from_records.get(area, 0) + 1

    monthly = df.groupby(df['_signup_dt'].dt.to_period('M').astype(str)).size()

    s = {
        'pax'         : len(df),
        'orders'      : int(df['訂單編號'].nunique()),
        'revenue'     : int(df['_fee'].sum()),
        'avg'         : int(df['_fee'].mean()),
        'female'      : int((df['性別'] == '女').sum()),
        'male'        : int((df['性別'] == '男').sum()),
        'female_pct'  : round(df[df['性別'] == '女'].shape[0] / max(df['性別'].isin(['男', '女']).sum(), 1) * 100, 1),
        'male_pct'    : round(df[df['性別'] == '男'].shape[0] / max(df['性別'].isin(['男', '女']).sum(), 1) * 100, 1),
        'beginner_pct': round(df[df['_lv'] == 'lv0'].shape[0]  / max((df['_lv'] != '').sum(), 1) * 100, 1),
        'board_pct'   : round(df[df['滑雪類別'] == '單板'].shape[0] / max(df['滑雪類別'].isin(['單板', '雙板']).sum(), 1) * 100, 1),
        'updated'     : datetime.now().strftime('%Y/%m/%d %H:%M'),
    }

    return json.dumps({
        's'          : s,
        'area'       : {k: int(v) for k, v in area_from_records.items() if k != '其他'},
        'area_type'  : area_type,
        'daily'      : daily_list,
        'weekly'     : weekly_list,
        'monthly'    : {k: int(v) for k, v in monthly.items() if k not in ('NaT', 'nan')},
        'cum'        : cum_list,
        'stacked'    : {'months': months_sorted, 'areas': AREAS, 'vals': stacked_vals},
        'level'      : level_dist,
        'area_level' : area_level,
        'gender'     : {str(k): int(v) for k, v in df['性別'].value_counts().items() if str(k) not in ('nan', '')},
        'ski'        : {str(k): int(v) for k, v in df['滑雪類別'].value_counts().items() if str(k) not in ('nan', '')},
        'age'        : age_dist,
        'addons'     : addons,
    }, ensure_ascii=False)


SALES_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>DBC 銷量分析</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700&family=DM+Mono:wght@400;500&display=swap');
:root{--bg:#0a0f1e;--s1:#111827;--s2:#1a2235;--bd:#1e2d45;--ac:#3b82f6;--gr:#10b981;--tx:#e2e8f0;--mu:#64748b;--ny:#3b82f6;--nb:#06b6d4;--nt:#f59e0b;--nl:#8b5cf6;}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Noto Sans TC',sans-serif}
.hd{background:linear-gradient(135deg,#0f172a,#1e3a5f 50%,#0f172a);border-bottom:1px solid var(--bd);padding:14px 20px;position:sticky;top:0;z-index:100}
.hd-in{max-width:1280px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap}
.logo{display:flex;align-items:center;gap:10px}
.logo-i{width:30px;height:30px;background:var(--ac);border-radius:7px;display:flex;align-items:center;justify-content:center}
.logo-t h1{font-size:15px;font-weight:700}.logo-t p{font-size:10px;color:var(--mu);font-family:'DM Mono',monospace}
.badge{background:var(--s2);border:1px solid var(--bd);border-radius:20px;padding:4px 12px;font-size:11px;color:var(--nb);font-family:'DM Mono',monospace}
.main{max-width:1280px;margin:0 auto;padding:18px 16px 48px}
.sec{font-size:11px;font-weight:500;color:var(--mu);text-transform:uppercase;letter-spacing:1px;margin:18px 0 12px;display:flex;align-items:center;gap:8px}
.sec::after{content:'';flex:1;height:1px;background:var(--bd)}
.card{background:var(--s1);border:1px solid var(--bd);border-radius:10px;padding:16px;margin-bottom:14px}
.ct{font-size:12px;font-weight:500;margin-bottom:14px;display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}.cw{position:relative}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:14px}.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}
.overview{display:grid;grid-template-columns:auto 1fr auto;gap:20px;align-items:center}
.ov-total{text-align:center;padding-right:20px;border-right:1px solid var(--bd)}
.ov-total .big{font-size:48px;font-weight:700;font-family:'DM Mono',monospace;line-height:1}
.ov-total .lbl{font-size:11px;color:var(--mu);margin-top:4px}
.ov-areas{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px}
.area-card{background:var(--s2);border:1px solid var(--bd);border-left:3px solid var(--ac);border-radius:8px;padding:12px}
.area-name{font-size:13px;font-weight:600;margin-bottom:6px}
.area-total{font-size:20px;font-weight:700;font-family:'DM Mono',monospace;line-height:1;margin-bottom:6px}
.area-bar{height:5px;background:var(--s1);border-radius:3px;margin-bottom:6px;overflow:hidden;display:flex;gap:1px}
.area-bar-g{height:100%;border-radius:2px}
.area-split{display:flex;justify-content:space-between;font-size:10px;color:var(--mu)}
.ov-rev{text-align:center;padding-left:20px;border-left:1px solid var(--bd)}
.ov-rev .rev-v{font-size:32px;font-weight:700;font-family:'DM Mono',monospace;color:var(--gr);line-height:1}
.ov-rev .rev-l{font-size:11px;color:var(--mu);margin-top:4px}
.tabs{display:flex;gap:4px;flex-wrap:wrap}
.tab{padding:4px 11px;border-radius:5px;font-size:11px;cursor:pointer;border:1px solid transparent;color:var(--mu);background:transparent}
.tab.active{background:var(--ac);border-color:var(--ac);color:#fff;font-weight:500}
.mini-kpi{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:10px;margin-bottom:14px}
.mk{background:var(--s2);border:1px solid var(--bd);border-radius:8px;padding:12px;position:relative;overflow:hidden}
.mk::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--mkc,var(--ac))}
.mk-l{font-size:10px;color:var(--mu);text-transform:uppercase;letter-spacing:.6px;margin-bottom:5px}
.mk-v{font-size:20px;font-weight:700;font-family:'DM Mono',monospace;line-height:1}
.mk-s{font-size:10px;color:var(--mu);margin-top:3px}
.addon-g{display:grid;grid-template-columns:repeat(auto-fill,minmax(105px,1fr));gap:8px}
.addon-c{background:var(--s2);border:1px solid var(--bd);border-radius:8px;padding:11px;text-align:center}
.addon-n{font-size:19px;font-weight:700;font-family:'DM Mono',monospace;color:var(--ac)}
.addon-l{font-size:10px;color:var(--mu);margin-top:3px}.addon-p{font-size:9px;color:var(--mu);font-family:'DM Mono',monospace}
@media(max-width:768px){.overview{grid-template-columns:1fr;gap:14px}.g2,.g3{grid-template-columns:1fr}}
</style></head><body>
<div class="hd"><div class="hd-in">
  <div class="logo"><div class="logo-i">🏔</div><div class="logo-t"><h1>DBC 銷量分析</h1><p>SEASON REPORT</p></div></div>
  <div class="badge" id="upd">—</div>
</div></div>
<div class="main">
  <div class="sec">總覽</div>
  <div class="card"><div class="overview">
    <div class="ov-total"><div class="big" id="ov-pax">—</div><div class="lbl">總報名人次</div></div>
    <div class="ov-areas" id="ov-areas"></div>
    <div class="ov-rev"><div class="rev-v" id="ov-rev">—</div><div class="rev-l">總營收 NT$</div></div>
  </div></div>
  <div class="sec">累積報名人次</div>
  <div class="card">
    <div class="ct"><span class="dot" style="background:#10b981"></span>累積報名曲線
      <div class="tabs" style="margin-left:auto">
        <button class="tab active" onclick="setCum('all',this)">全期</button>
        <button class="tab" onclick="setCum('2m',this)">近2月</button>
        <button class="tab" onclick="setCum('1m',this)">近1月</button>
      </div>
    </div>
    <div class="cw"><canvas id="cumC" height="90"></canvas></div>
  </div>
  <div class="sec">每日新增報名</div>
  <div class="card">
    <div class="ct"><span class="dot" style="background:#f59e0b"></span>每日新增
      <div class="tabs" style="margin-left:auto" id="daily-tabs">
        <button class="tab active" onclick="setDaily('all','all',this)">全部</button>
        <button class="tab" onclick="setDaily('all','野澤',this)">野澤</button>
        <button class="tab" onclick="setDaily('all','斑尾',this)">斑尾</button>
        <button class="tab" onclick="setDaily('all','湯澤',this)">湯澤</button>
        <button class="tab" onclick="setDaily('all','龍平',this)">龍平</button>
        <button class="tab" onclick="setDaily('3m','all',this)">近3月</button>
      </div>
    </div>
    <div class="cw"><canvas id="dailyC" height="140"></canvas></div>
  </div>
  <div class="sec">每週新增報名</div>
  <div class="card">
    <div class="ct"><span class="dot" style="background:#8b5cf6"></span>每週新增
      <div class="tabs" style="margin-left:auto" id="weekly-tabs">
        <button class="tab active" onclick="setWeekly('all',this)">全部</button>
        <button class="tab" onclick="setWeekly('野澤',this)">野澤</button>
        <button class="tab" onclick="setWeekly('斑尾',this)">斑尾</button>
        <button class="tab" onclick="setWeekly('湯澤',this)">湯澤</button>
        <button class="tab" onclick="setWeekly('龍平',this)">龍平</button>
      </div>
    </div>
    <div class="cw"><canvas id="weeklyC" height="140"></canvas></div>
  </div>
  <div class="sec">出發日期分析</div>
  <div class="g2">
    <div class="card" style="margin-bottom:0"><div class="ct"><span class="dot" style="background:#06b6d4"></span>各區出發人次</div><div id="area-bars"></div></div>
    <div class="card" style="margin-bottom:0"><div class="ct"><span class="dot" style="background:#3b82f6"></span>出發月份 × 各區</div><div class="cw"><canvas id="stackC" height="200"></canvas></div></div>
  </div>
  <div class="sec" style="margin-top:14px">客群分析</div>
  <div class="mini-kpi" id="demo-kpi"></div>
  <div class="g3">
    <div class="card" style="margin-bottom:0"><div class="ct"><span class="dot" style="background:#06b6d4"></span>Level 分布</div><div class="cw"><canvas id="lvC" height="200"></canvas></div></div>
    <div class="card" style="margin-bottom:0"><div class="ct"><span class="dot" style="background:#3b82f6"></span>年齡分布</div><div class="cw"><canvas id="ageC" height="200"></canvas></div></div>
    <div class="card" style="margin-bottom:0"><div class="ct"><span class="dot" style="background:#8b5cf6"></span>各區Level
      <div class="tabs" style="margin-left:auto">
        <button class="tab active" onclick="setAreaLv('野澤',this)">野澤</button>
        <button class="tab" onclick="setAreaLv('斑尾',this)">斑尾</button>
        <button class="tab" onclick="setAreaLv('湯澤',this)">湯澤</button>
        <button class="tab" onclick="setAreaLv('龍平',this)">龍平</button>
      </div>
    </div><div class="cw"><canvas id="areaLvC" height="170"></canvas></div></div>
  </div>
  <div class="sec" style="margin-top:14px">加購 & 租借</div>
  <div class="card"><div class="ct"><span class="dot" style="background:#10b981"></span>加購項目人次</div><div class="addon-g" id="addon-g"></div></div>
</div>
<script>
const D=REPORT_DATA_PLACEHOLDER;
const AC={'野澤':'#ffffff','斑尾':'#7ec8f5','湯澤':'#ffb347','龍平':'#c8864a'};
const LC={'lv0':'#64748b','lv1':'#3b82f6','lv2':'#06b6d4','lv3':'#f59e0b','lv4':'#ef4444'};
const LL={'lv0':'lv0 初學','lv1':'lv1 入門','lv2':'lv2 初中','lv3':'lv3 中級','lv4':'lv4 高級'};
const AREAS=['野澤','斑尾','湯澤','龍平'];
function co(h,a){const r=parseInt(h.slice(1,3),16),g=parseInt(h.slice(3,5),16),b=parseInt(h.slice(5,7),16);return`rgba(${r},${g},${b},${a})`}
const fmt=n=>n>=10000?(n/10000).toFixed(1)+'萬':n.toLocaleString();
const CD={responsive:true,plugins:{legend:{display:false},tooltip:{backgroundColor:'#1a2235',borderColor:'#1e2d45',borderWidth:1,titleColor:'#e2e8f0',bodyColor:'#94a3b8',padding:10,cornerRadius:8}},scales:{x:{grid:{color:'#1e2d45'},ticks:{color:'#64748b',font:{size:10},maxTicksLimit:14}},y:{grid:{color:'#1e2d45'},ticks:{color:'#64748b',font:{size:10}}}}};
document.getElementById('ov-pax').textContent=D.s.pax.toLocaleString();
document.getElementById('ov-rev').textContent=fmt(D.s.revenue);
document.getElementById('upd').textContent='更新：'+D.s.updated;
const og=document.getElementById('ov-areas'),tot=D.s.pax;
AREAS.forEach(a=>{const at=D.area_type[a]||{pure:0,group:0,total:0};if(!at.total)return;const c=AC[a];const aPct=(at.total/tot*100).toFixed(1);
og.innerHTML+=`<div class="area-card" style="border-left-color:${c}"><div class="area-name" style="color:${c}">${a} <span style="font-size:10px;color:var(--mu);font-weight:400">${aPct}%</span></div><div class="area-total">${at.total.toLocaleString()}<span style="font-size:11px;color:var(--mu);font-weight:400"> 人</span></div><div class="area-bar"><div class="area-bar-g" style="flex:${at.group};background:${c};opacity:.9"></div><div class="area-bar-g" style="flex:${at.pure};background:${c};opacity:.35"></div></div><div class="area-split"><span style="color:${c}">團${at.group} (${at.total>0?(at.group/at.total*100).toFixed(0):0}%)</span><span style="opacity:.6">純${at.pure}</span></div></div>`;});
let cumChart=null;
function setCum(range,btn){document.querySelectorAll('.tab').forEach(t=>{if(t.onclick&&t.onclick.toString().includes('setCum'))t.classList.remove('active')});btn.classList.add('active');
let pts=D.cum;if(range!=='all'){const days=range==='1m'?30:60;const cut=new Date(pts[pts.length-1].d);cut.setDate(cut.getDate()-days);pts=pts.filter(p=>new Date(p.d)>=cut);}
if(cumChart)cumChart.destroy();cumChart=new Chart(document.getElementById('cumC'),{type:'line',data:{labels:pts.map(p=>p.d.slice(5)),datasets:[{data:pts.map(p=>p.c),borderColor:'#10b981',backgroundColor:co('#10b981',.1),fill:true,tension:.4,borderWidth:2,pointRadius:0}]},options:{...CD}});}
setCum('all',document.querySelector('.tab'));
let dailyChart=null;
function setDaily(range,area,btn){document.querySelectorAll('#daily-tabs .tab').forEach(t=>t.classList.remove('active'));btn.classList.add('active');
let pts=D.daily;if(range==='3m'){const cut=new Date(pts[pts.length-1].d);cut.setDate(cut.getDate()-90);pts=pts.filter(p=>new Date(p.d)>=cut);}
if(dailyChart)dailyChart.destroy();const labels=pts.map(p=>p.d.slice(5));const stacked=area==='all';
const datasets=area==='all'?AREAS.map(a=>({label:a,data:pts.map(p=>p[a]||0),backgroundColor:co(AC[a],.8),borderRadius:1,borderSkipped:false})):[{label:area,data:pts.map(p=>p[area]||0),backgroundColor:co(AC[area],.8),borderRadius:2,borderSkipped:false}];
dailyChart=new Chart(document.getElementById('dailyC'),{type:'bar',data:{labels,datasets},options:{...CD,plugins:{...CD.plugins,legend:{display:stacked,position:'bottom',labels:{color:'#94a3b8',font:{size:10},boxWidth:10,padding:10}}},scales:{x:{...CD.scales.x,stacked},y:{...CD.scales.y,stacked}}}});}
setDaily('all','all',document.querySelectorAll('#daily-tabs .tab')[0]);
let weeklyChart=null;
function setWeekly(area,btn){document.querySelectorAll('#weekly-tabs .tab').forEach(t=>t.classList.remove('active'));btn.classList.add('active');
const pts=D.weekly;const stacked=area==='all';
const datasets=area==='all'?AREAS.map(a=>({label:a,data:pts.map(p=>p[a]||0),backgroundColor:co(AC[a],.8),borderRadius:1,borderSkipped:false})):[{label:area,data:pts.map(p=>p[area]||0),backgroundColor:co(AC[area],.8),borderRadius:2,borderSkipped:false}];
if(weeklyChart)weeklyChart.destroy();weeklyChart=new Chart(document.getElementById('weeklyC'),{type:'bar',data:{labels:pts.map(p=>p.w.slice(5,12)),datasets},options:{...CD,plugins:{...CD.plugins,legend:{display:stacked,position:'bottom',labels:{color:'#94a3b8',font:{size:10},boxWidth:10,padding:10}}},scales:{x:{...CD.scales.x,stacked},y:{...CD.scales.y,stacked}}}});}
setWeekly('all',document.querySelectorAll('#weekly-tabs .tab')[0]);
const areaBars=document.getElementById('area-bars');
AREAS.forEach(a=>{const at=D.area_type[a]||{total:0};if(!at.total)return;const pct=(at.total/tot*100).toFixed(1);const c=AC[a];
areaBars.innerHTML+=`<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px"><div style="font-size:11px;color:var(--mu);min-width:45px;text-align:right">${a}</div><div style="flex:1;height:22px;background:var(--s2);border-radius:5px;overflow:hidden"><div style="width:${pct}%;height:100%;background:${c};opacity:.8;display:flex;align-items:center;padding-left:8px;font-size:10px;font-family:'DM Mono',monospace;color:#000">${at.total.toLocaleString()}</div></div><div style="font-size:10px;color:var(--mu);font-family:'DM Mono',monospace;min-width:36px">${pct}%</div></div>`;});
const sM=D.stacked.months.map(m=>m.slice(2).replace('-','/'));
new Chart(document.getElementById('stackC'),{type:'bar',data:{labels:sM,datasets:D.stacked.areas.map((a,ai)=>({label:a,data:D.stacked.vals.map(r=>r[ai]),backgroundColor:co(AC[a]||'#64748b',.8),borderRadius:2,borderSkipped:false}))},options:{...CD,plugins:{...CD.plugins,legend:{display:true,position:'bottom',labels:{color:'#94a3b8',font:{size:10},boxWidth:10,padding:10}}},scales:{x:{...CD.scales.x,stacked:true},y:{...CD.scales.y,stacked:true}}}});
const dk=document.getElementById('demo-kpi');
[{l:'女性',v:D.s.female,s:D.s.female_pct+'%',c:'#ec4899'},{l:'男性',v:D.s.male,s:D.s.male_pct+'%',c:'#3b82f6'},{l:'單板',v:D.ski['單板']||0,s:D.s.board_pct+'%',c:'#06b6d4'},{l:'雙板',v:D.ski['雙板']||0,s:(100-D.s.board_pct).toFixed(1)+'%',c:'#f59e0b'},{l:'初學(lv0)',v:D.level['lv0']||0,s:D.s.beginner_pct+'%',c:'#8b5cf6'}].forEach(item=>{
dk.innerHTML+=`<div class="mk" style="--mkc:${item.c}"><div class="mk-l">${item.l}</div><div class="mk-v" style="color:${item.c}">${item.v.toLocaleString()}</div><div class="mk-s">${item.s}</div></div>`;});
const lvK=Object.keys(D.level),lvV=Object.values(D.level);
new Chart(document.getElementById('lvC'),{type:'bar',data:{labels:lvK.map(k=>LL[k]||k),datasets:[{data:lvV,backgroundColor:lvK.map(k=>co(LC[k]||'#64748b',.8)),borderRadius:3,borderSkipped:false}]},options:{...CD}});
const ageK=Object.keys(D.age),ageV=Object.values(D.age);
new Chart(document.getElementById('ageC'),{type:'bar',data:{labels:ageK,datasets:[{data:ageV,backgroundColor:ageV.map((_,i)=>co('#f59e0b',.4+i/ageK.length*.6)),borderRadius:3,borderSkipped:false}]},options:{...CD,indexAxis:'y'}});
let aLvChart=null;
function setAreaLv(area,btn){document.querySelectorAll('.tab').forEach(t=>{if(t.onclick&&t.onclick.toString().includes('setAreaLv'))t.classList.remove('active')});btn.classList.add('active');
const d=D.area_level[area]||{};const keys=['lv0','lv1','lv2','lv3','lv4'];
if(aLvChart)aLvChart.destroy();aLvChart=new Chart(document.getElementById('areaLvC'),{type:'bar',data:{labels:keys.map(k=>LL[k]||k),datasets:[{data:keys.map(k=>d[k]||0),backgroundColor:keys.map(k=>co(LC[k]||'#64748b',.8)),borderRadius:3,borderSkipped:false}]},options:{...CD}});}
setAreaLv('野澤',document.querySelectorAll('.tab')[document.querySelectorAll('.tab').length-4]);
const ag=document.getElementById('addon-g');
Object.entries(D.addons).forEach(([name,val])=>{const pct=(val/D.s.pax*100).toFixed(1);ag.innerHTML+=`<div class="addon-c"><div class="addon-n">${val.toLocaleString()}</div><div class="addon-l">${name}</div><div class="addon-p">${pct}% 人次</div></div>`;});
</script></body></html>"""


# ═══════════════════════════════════════════
# 主函數
# ═══════════════════════════════════════════

def run(cleaner_output=None):
    print("=" * 50)
    print("REPORTER 報表員 啟動")
    print("=" * 50)

    if cleaner_output is None:
        from agents.cleaner import load
        cleaner_output = load()

    all_records_flat = cleaner_output['all_records_flat']
    DATE_START       = cleaner_output['DATE_START']
    DATE_END         = cleaner_output['DATE_END']
    headers          = cleaner_output['headers']
    df_records       = cleaner_output['df_records']

    # ── 報表1：每日教練需求 ──
    print("\n[報表1] 每日教練需求...")
    try:
        coach_json = _build_coach_json(all_records_flat, DATE_START, DATE_END)
        coach_html = COACH_HTML_TEMPLATE.replace('COACH_DATA_PLACEHOLDER', coach_json)
        coach_path = os.path.join(BASE_DIR, 'dbc_coach_report.html')
        with open(coach_path, 'w', encoding='utf-8') as f:
            f.write(coach_html)
        print(f"  OK → dbc_coach_report.html")
    except Exception as e:
        import traceback
        print(f"  教練需求報表生成失敗：{e}")
        traceback.print_exc()

    # ── 報表2：銷量分析 ──
    print("\n[報表2] 銷量分析...")
    try:
        sales_json = _build_sales_json(df_records, all_records_flat, headers)
        sales_html = SALES_HTML_TEMPLATE.replace('REPORT_DATA_PLACEHOLDER', sales_json)
        sales_path = os.path.join(BASE_DIR, 'dbc_sales_report.html')
        with open(sales_path, 'w', encoding='utf-8') as f:
            f.write(sales_html)
        print(f"  OK → dbc_sales_report.html")
    except Exception as e:
        import traceback
        print(f"  銷量報表生成失敗：{e}")
        traceback.print_exc()

    print(f"\n{'=' * 50}")
    print(f"REPORTER 完成！{datetime.now().strftime('%Y/%m/%d %H:%M')}")
    print(f"{'=' * 50}\n")


# ─────────────────────────────────────────────
if __name__ == '__main__':
    run()
