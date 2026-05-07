#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""COT Dashboard Generator v13"""

import math
import pandas as pd
import json
import webbrowser
from pathlib import Path
from datetime import datetime

BASE_DIR    = Path(__file__).parent
DATA_FILE   = BASE_DIR / "data" / "COT_OVERVIEW.xlsx"
OUTPUT_FILE = BASE_DIR / "output" / "cot_dashboard.html"

SKIP_SHEETS  = {'overview', 'info'}
CHART_WEEKS  = 260
SPARK_WEEKS  = 26
HISTORY      = 52

COLOR_LS = '#4a9eff'
COLOR_CM = '#20d483'
COLOR_ST = '#f0515a'

REPORTS = [
    {'id':'usda_crop',  'name':'USDA Crop Progress', 'sched':'Пн 22:00 (кві-лис)', 'tag':None},
    {'id':'eia_petrol', 'name':'EIA Petroleum',      'sched':'Ср 16:30',            'tag':None},
    {'id':'usda_exp',   'name':'USDA Export Sales',  'sched':'Чт 14:30',            'tag':None},
    {'id':'cot_cftc',   'name':'COT Report (CFTC)',  'sched':'Пт 21:30',            'tag':None},
    {'id':'usda_wasde', 'name':'USDA WASDE',         'sched':'~12 число, 18:00',    'tag':'Місячний'},
    {'id':'usda_oil',   'name':'USDA Oilseeds',      'sched':'з WASDE, 18:15',      'tag':'Місячний'},
]

REPORT_RELEVANCE = {
    'usda_crop': {
        'CORN':'direct','WHEAT':'direct','SOYBEAN':'direct',
        'SOYBEAN_MEAL':'direct','SOYBEAN_OIL':'direct',
        'COTTON':'direct','RICE':'direct','COFFEE':'direct',
        'COCOA':'direct','SUGAR':'direct','OJ':'direct',
        'CATTLE':'indirect','LUMBER':'indirect','_default':'none',
    },
    'eia_petrol': {
        'WTI_CRUDE':'direct','BRENT':'direct','NAT_GAS':'direct',
        'SP500':'indirect','NASDAQ':'indirect','DOW_30':'indirect',
        'RUSSELL2K':'indirect','VIX':'indirect','CAD':'indirect','_default':'none',
    },
    'usda_exp': {
        'WHEAT':'direct','CORN':'direct','SOYBEAN':'direct',
        'SOYBEAN_MEAL':'direct','SOYBEAN_OIL':'direct',
        'COTTON':'direct','RICE':'direct','COFFEE':'direct',
        'COCOA':'indirect','SUGAR':'indirect','OJ':'indirect','_default':'none',
    },
    'cot_cftc': {'_default':'indirect'},
    'usda_wasde': {
        'WHEAT':'direct','CORN':'direct','SOYBEAN':'direct',
        'SOYBEAN_MEAL':'direct','SOYBEAN_OIL':'direct',
        'COTTON':'direct','RICE':'direct','SUGAR':'direct',
        'COFFEE':'direct','COCOA':'direct','OJ':'direct','CATTLE':'direct',
        'LUMBER':'indirect','SP500':'indirect','NASDAQ':'indirect',
        'DOW_30':'indirect','RUSSELL2K':'indirect','VIX':'indirect',
        'WTI_CRUDE':'indirect','BRENT':'indirect','NAT_GAS':'indirect',
        'AUD':'indirect','CAD':'indirect','NZD':'indirect','_default':'none',
    },
    'usda_oil': {
        'SOYBEAN':'direct','SOYBEAN_MEAL':'direct','SOYBEAN_OIL':'direct',
        'CORN':'indirect','WHEAT':'indirect','COTTON':'indirect',
        'WTI_CRUDE':'indirect','BRENT':'indirect','_default':'none',
    },
}

def get_relevance(instrument_sid, report_id):
    mapping = REPORT_RELEVANCE.get(report_id, {})
    return mapping.get(instrument_sid, mapping.get('_default', 'none'))

DISPLAY = {
    'SP500':'S&P 500','DOW_30':'DOW 30','RUSSELL2K':'RUSSELL 2K',
    'NAT_GAS':'NAT GAS','WTI_CRUDE':'WTI CRUDE',
    'SOYBEAN_MEAL':'SOYBEAN MEAL','SOYBEAN_OIL':'SOYBEAN OIL',
    'BTC_MICRO':'BTC MICRO','BTC_NANO':'BTC NANO',
    'ETH_CASH':'ETH CASH','ETH_MICRO':'ETH MICRO','ETH_NANO':'ETH NANO',
}
CATEGORIES = {
    'Валюти':  ['DXY','EUR','GBP','JPY','AUD','CAD','CHF','NZD'],
    'Метали':  ['GOLD','SILVER','COPPER','PLATINUM','PALLADIUM'],
    'Індекси': ['SP500','NASDAQ','DOW_30','RUSSELL2K','VIX'],
    'Енергія': ['WTI_CRUDE','BRENT','NAT_GAS'],
    'Агро':    ['CORN','WHEAT','SOYBEAN','SOYBEAN_MEAL','SOYBEAN_OIL',
                'SUGAR','COFFEE','COCOA','COTTON','RICE','OJ','LUMBER','CATTLE'],
    'Крипто':  ['BTC','BTC_MICRO','BTC_NANO','ETH_CASH','ETH_MICRO','ETH_NANO'],
}
OVERVIEW_TO_SID = {
    'DXY':'DXY','EUR':'EUR','GBP':'GBP','CAD':'CAD','JPY':'JPY',
    'AUD':'AUD','NZD':'NZD','CHF':'CHF',
    'GOLD':'GOLD','SILVER':'SILVER','COPPER':'COPPER',
    'PLATINUM':'PLATINUM','PALLADIUM':'PALLADIUM',
    'SP500':'SP500','S&P500':'SP500','NASDAQ':'NASDAQ',
    'DOW_30':'DOW_30','DOW 30':'DOW_30','RUSSELL2K':'RUSSELL2K','VIX':'VIX',
    'WTI_CRUDE':'WTI_CRUDE','WTI CRUDE':'WTI_CRUDE',
    'BRENT':'BRENT','NAT_GAS':'NAT_GAS','NAT GAS':'NAT_GAS',
    'CORN':'CORN','WHEAT':'WHEAT','SOYBEAN':'SOYBEAN',
    'SOYBEAN_MEAL':'SOYBEAN_MEAL','SOYBEAN MEAL':'SOYBEAN_MEAL',
    'SOYBEAN_OIL':'SOYBEAN_OIL','SOYBEAN OIL':'SOYBEAN_OIL',
    'SUGAR':'SUGAR','COFFEE':'COFFEE','COCOA':'COCOA','COTTON':'COTTON',
    'RICE':'RICE','OJ':'OJ','LUMBER':'LUMBER','CATTLE':'CATTLE',
    'BTC':'BTC','BTC_NANO':'BTC_NANO','BTC NANO':'BTC_NANO',
    'BTC_MICRO':'BTC_MICRO','BTC MICRO':'BTC_MICRO',
    'ETH_CASH':'ETH_CASH','ETH CASH':'ETH_CASH',
    'ETH_MICRO':'ETH_MICRO','ETH MICRO':'ETH_MICRO',
    'ETH_NANO':'ETH_NANO','ETH NANO':'ETH_NANO',
}
COL = {
    'date':1,'ls_cl':4,'ls_cs':5,'ls_pct':6,'ls_net':8,
    'cm_cl':11,'cm_cs':12,'cm_pct':13,'cm_net':15,
    'st_cl':18,'st_cs':19,'st_pct':20,'st_net':22,
}
DATA_START_ROW = 5


def disp(n): return DISPLAY.get(n, n)
def sid(n):  return n.replace(' ','_').replace('&','n').replace('/','_')
def to_num(s): return pd.to_numeric(s, errors='coerce').fillna(0).round(2).tolist()

def norm_pct(vals):
    if not vals: return vals
    nz = [abs(v) for v in vals if v != 0]
    if nz and max(nz) <= 1.5: return [round(v*100,1) for v in vals]
    return [round(v,1) for v in vals]

def fv(n, short=False, sign=False):
    try: n = int(round(float(n)))
    except: return '—'
    if short:
        if abs(n)>=1_000_000: body=f"{abs(n)/1_000_000:.1f}M"
        elif abs(n)>=1_000:   body=f"{abs(n)/1_000:.0f}K"
        else:                 body=str(abs(n))
    else:
        body=f"{abs(n):,}".replace(',','\u202f')
    if n>0:   return ('+'if sign else '')+body
    elif n<0: return '-'+body
    return body

def fv_full(n, sign=False):
    try: n=int(round(float(n)))
    except: return '—'
    body=f"{abs(n):,}".replace(',','\u202f')
    if n>0:   return ('+'if sign else '')+body
    elif n<0: return '-'+body
    return body

def fp(n, signed=False):
    try:
        v=float(n); s='+'if(signed and v>0)else ''
        return f"{s}{v:.1f}%"
    except: return '—'

def pct_change(chg, net):
    try:
        prev = float(net) - float(chg)
        if abs(prev) < 1: return '—'
        return fp(float(chg)/abs(prev)*100, signed=True)
    except: return '—'

def cc(n):
    try:
        v=float(n)
        return 'g'if v>0 else('r'if v<0 else 'd')
    except: return 'd'

def ar(n):
    try:
        v=float(n)
        return '▲'if v>0 else('▼'if v<0 else'—')
    except: return '—'

def calc_cot_index(series, weeks=None):
    s=list(series[-weeks:]) if weeks and len(series)>=weeks else list(series)
    if len(s)<2: return 50.0
    cur=s[-1]; mn=min(s); mx=max(s)
    if mx==mn: return 50.0
    return round((cur-mn)/(mx-mn)*100,1)

def compute_delta(vals):
    if not vals: return []
    result=[0.0]
    for i in range(1,len(vals)):
        result.append(round(float(vals[i])-float(vals[i-1]),0))
    return result

def find_oi_col(df):
    n=len(df)
    for idx in range(23,min(50,df.shape[1])):
        vals=pd.to_numeric(df.iloc[:,idx],errors='coerce').fillna(0)
        if (vals.abs()>100).sum()>n*0.4 and vals.abs().mean()>1000:
            return idx
    return 23


def make_sparkline(series, color, h=38):
    data=[float(v) for v in (series or [])[-SPARK_WEEKS:]]
    n=len(data)
    if n<3: return ''
    W=200; H=h; mn=min(data); mx=max(data); rng=mx-mn
    if rng==0: return ''
    def px(i): return round(i/(n-1)*W,1)
    def py(v): return round((1-(v-mn)/rng)*H,1)
    pts=[(px(i),py(v)) for i,v in enumerate(data)]
    line=' '.join(f"{x},{y}"for x,y in pts)
    zy=max(0,min(H,py(0))); area=f"0,{zy} "+line+f" {W},{zy}"
    lx,ly=pts[-1]
    return(
        f'<svg viewBox="0 0 {W} {H}" preserveAspectRatio="none" '
        f'xmlns="http://www.w3.org/2000/svg" '
        f'style="width:100%;height:{H}px;display:block;margin-top:8px">'
        f'<polygon points="{area}" fill="{color}" opacity="0.14"/>'
        f'<polyline points="{line}" fill="none" stroke="{color}" '
        f'stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>'
        f'<circle cx="{lx}" cy="{ly}" r="2.5" fill="{color}"/>'
        f'</svg>'
    )


OVERVIEW_TABLE = []

def read_overview(xl):
    global OVERVIEW_TABLE
    result={}; OVERVIEW_TABLE=[]
    try:
        raw=xl.parse('overview',header=None)
        def safe_date(cell):
            v=pd.to_datetime(cell,errors='coerce')
            return v.strftime('%d.%m.%Y') if pd.notna(v) else '—'
        try:
            rep_date=safe_date(raw.iloc[1,2]); today_date=safe_date(raw.iloc[1,5])
        except:
            rep_date='—'; today_date='—'
        OVERVIEW_TABLE.append(('_meta',rep_date,today_date))
        current_group=''
        for i in range(4,len(raw)):
            row=raw.iloc[i]
            asset=str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            if not asset or asset=='nan': continue
            def safe(c):
                v=pd.to_numeric(row.iloc[c],errors='coerce')
                return float(v) if pd.notna(v) else None
            cot_ls=safe(4)
            if cot_ls is None:
                current_group=asset
                OVERVIEW_TABLE.append(('_group',asset))
                continue
            s=OVERVIEW_TO_SID.get(asset,asset)
            net_ls=safe(2); net_cm=safe(3)
            chg_ls=safe(10); chg_cm=safe(12)
            oi_chg_pct=safe(14)
            sm_div=safe(8); sm_div_3m=safe(18); sm_div_6m=safe(19)
            result[s]={
                'cot_ls_all':round(cot_ls*100,1),
                'cot_cm_all':round((safe(5) or 0)*100,1),
                'cot_st_all':round((safe(6) or 0)*100,1),
                'sm_div':    round(sm_div or 0,3),
                'sm_div_3m': round(sm_div_3m or 0,3),
                'sm_div_6m': round(sm_div_6m or 0,3),
            }
            OVERVIEW_TABLE.append({
                'asset':asset,'sid':s,'group':current_group,
                'net_ls':net_ls,'net_cm':net_cm,
                'cot_ls':round(cot_ls*100,1),
                'cot_cm':round((safe(5) or 0)*100,1),
                'cot_st':round((safe(6) or 0)*100,1),
                'chg_ls':chg_ls,'chg_cm':chg_cm,
                'oi_chg_pct':oi_chg_pct,
                'sm_div':sm_div,'sm_div_3m':sm_div_3m,'sm_div_6m':sm_div_6m,
            })
    except Exception as e:
        print(f"  ⚠  overview: {e}")
    return result


def read_sheet(xl, name, overview):
    try:
        raw=xl.parse(name,header=None)
        if raw.shape[0]<DATA_START_ROW+2 or raw.shape[1]<20: return None
        df=raw.iloc[DATA_START_ROW:].reset_index(drop=True).copy()
        dates_raw=pd.to_datetime(df.iloc[:,COL['date']],errors='coerce')
        valid=dates_raw.notna()&(dates_raw.dt.year>2000)
        df=df[valid].copy(); df['_dt']=dates_raw[valid].values
        if df.empty: return None
        df=df.sort_values('_dt').reset_index(drop=True)
        oi_col=find_oi_col(df)
        def gc(idx):
            return to_num(df.iloc[:,idx]) if idx<df.shape[1] else [0.0]*len(df)
        ls_net=gc(COL['ls_net']); cm_net=gc(COL['cm_net']); st_net=gc(COL['st_net'])
        ls_cl=gc(COL['ls_cl']); ls_cs=gc(COL['ls_cs'])
        cm_cl=gc(COL['cm_cl']); cm_cs=gc(COL['cm_cs'])
        ls_pct=norm_pct(gc(COL['ls_pct']))
        cm_pct=norm_pct(gc(COL['cm_pct']))
        st_pct=norm_pct(gc(COL['st_pct']))
        oi_all=gc(oi_col)
        all_dates=df['_dt'].dt.strftime('%d.%m.%Y').tolist()
        i0=-1; i1=-2 if len(all_dates)>1 else -1
        ov=overview.get(sid(name),{})
        def cot_idx(series,all_key):
            return{
                'all':ov.get(all_key,calc_cot_index(series)),
                '3y':calc_cot_index(series,156),
                '1y':calc_cot_index(series,52),
                '6m':calc_cot_index(series,26),
                '3m':calc_cot_index(series,13),
            }

        def stats(net_s, cl_s=None, cs_s=None):
            def mm(series):
                s = [float(v) for v in series if v != 0]
                s1y = s[-52:] if len(s) >= 52 else s
                return {
                    'max_all': max(s) if s else 0,
                    'min_all': min(s) if s else 0,
                    'max_1y':  max(s1y) if s1y else 0,
                    'min_1y':  min(s1y) if s1y else 0,
                }
            res = mm(net_s)
            res['cl'] = mm(cl_s) if cl_s is not None else None
            res['cs'] = mm(cs_s) if cs_s is not None else None
            return res

        N=min(CHART_WEEKS,len(df))
        chart_df=df.tail(N).reset_index(drop=True)
        def gcc(idx):
            return to_num(chart_df.iloc[:,idx]) if idx<chart_df.shape[1] else [0.0]*N
        ls_w=gcc(COL['ls_net']); cm_w=gcc(COL['cm_net']); st_w=gcc(COL['st_net'])
        oi_w=gcc(oi_col)
        chart={
            'dates':chart_df['_dt'].dt.strftime('%d.%m.%Y').tolist(),
            'ls':ls_w,'cm':cm_w,'st':st_w,'oi':oi_w,
            'ld':compute_delta(ls_w),'cd':compute_delta(cm_w),'sd':compute_delta(st_w),
        }
        hist_df=df.tail(HISTORY).reset_index(drop=True)
        def gch(idx):
            return to_num(hist_df.iloc[:,idx]) if idx<hist_df.shape[1] else [0.0]*len(hist_df)
        def gch_sm(idx, fallback_key):
            if idx < hist_df.shape[1]:
                vals = pd.to_numeric(hist_df.iloc[:,idx], errors='coerce').tolist()
                if any(v is not None and not (v != v) for v in vals):
                    return vals
            return [ov.get(fallback_key, None)] * len(hist_df)

        hist={
            'dates':hist_df['_dt'].dt.strftime('%d.%m.%Y').tolist(),
            'ls_cl':gch(COL['ls_cl']),'ls_cs':gch(COL['ls_cs']),'ls_net':gch(COL['ls_net']),
            'cm_cl':gch(COL['cm_cl']),'cm_cs':gch(COL['cm_cs']),'cm_net':gch(COL['cm_net']),
            'st_cl':gch(COL['st_cl']),'st_cs':gch(COL['st_cs']),'st_net':gch(COL['st_net']),
            'oi':gch(oi_col),
            'sm_div_row':    gch_sm(57, 'sm_div'),
            'sm_div_6m_row': gch_sm(58, 'sm_div_6m'),
            'sm_div_3m_row': gch_sm(59, 'sm_div_3m'),
        }
        oi_cur=float(oi_all[i0]); oi_prev=float(oi_all[i1])
        oi_pct=round((oi_cur-oi_prev)/abs(oi_prev)*100,2) if oi_prev!=0 else 0.0
        ls_chg=round(float(ls_net[i0])-float(ls_net[i1]),0)
        cm_chg=round(float(cm_net[i0])-float(cm_net[i1]),0)
        st_chg=round(float(st_net[i0])-float(st_net[i1]),0)
        oi_chg=round(oi_cur-oi_prev,0)
        # OI Capacity: поточний OI відносно максимального (all-time)
        _oi_stats = stats(oi_all)
        oi_capacity = round(oi_cur / _oi_stats['max_all'] * 100, 1) if _oi_stats.get('max_all', 0) > 0 else 50.0
        return{
            'name':name,'display':disp(name),'sid':sid(name),
            'chart':chart,'hist':hist,
            'stats_ls': stats(ls_net, ls_cl, ls_cs),
            'stats_cm': stats(cm_net, cm_cl, cm_cs),
            'stats_st': stats(st_net),
            'stats_oi': stats(oi_all),
            'cot_idx':{
                'ls':cot_idx(ls_net,'cot_ls_all'),
                'cm':cot_idx(cm_net,'cot_cm_all'),
                'st':cot_idx(st_net,'cot_st_all'),
            },
            'sm':{
                'div':ov.get('sm_div',0.0),'div_3m':ov.get('sm_div_3m',0.0),'div_6m':ov.get('sm_div_6m',0.0),
            },
            'spark':{'ls':ls_net,'cm':cm_net,'st':st_net,'oi':oi_all},
            'oi_capacity': oi_capacity,
            'cur':{
                'date':all_dates[i0],
                'ls_net':ls_net[i0],'cm_net':cm_net[i0],'st_net':st_net[i0],
                'ls_pct':ls_pct[i0],'cm_pct':cm_pct[i0],'st_pct':st_pct[i0],
                'ls_cl':ls_cl[i0],'ls_cs':ls_cs[i0],
                'cm_cl':cm_cl[i0],'cm_cs':cm_cs[i0],
                'oi':oi_cur,'oi_pct':oi_pct,
                'ls_chg':ls_chg,'cm_chg':cm_chg,'st_chg':st_chg,'oi_chg':oi_chg,
                'ls_chg_pct':pct_change(ls_chg,ls_net[i0]),
                'cm_chg_pct':pct_change(cm_chg,cm_net[i0]),
                'st_chg_pct':pct_change(st_chg,st_net[i0]),
                'oi_chg_pct':fp(oi_pct,signed=True),
            }
        }
    except Exception as e:
        print(f"  ❌  {name}: {e}"); return None


def load_all():
    print(f"📂  {DATA_FILE}")
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"\n❌  Файл не знайдено: {DATA_FILE}")
    xl=pd.ExcelFile(DATA_FILE)
    print("    Читаємо overview...")
    overview=read_overview(xl)
    sheets=[s for s in xl.sheet_names if s not in SKIP_SHEETS]
    print(f"    Вкладок: {len(sheets)}\n")
    data={}
    for s in sheets:
        r=read_sheet(xl,s,overview)
        if r:
            data[s]=r
            print(f"  ✓  {s:20s}  {r['cur']['date']}  "
                  f"LS={fv(r['cur']['ls_net'],True,True):>8s}  "
                  f"COTls={r['cot_idx']['ls']['all']:>5.1f}%")
        else:
            print(f"  ✗  {s}")
    print(f"\n✅  Завантажено: {len(data)} інструментів\n")
    return data


# ================================================================
# 🎯  GAUGE ІНДИКАТОР (дугоподібний)
# ================================================================
def make_gauge_svg(value, color, size=74):
    """
    Дугоподібний індикатор 0–100%.
    Дуга: від нижнього-лівого (7:20) до нижнього-правого (4:40),
    проходить через верхню точку. Загальний кут: 240° за годинниковою у SVG.
    """
    value = max(0.0, min(100.0, float(value)))
    cx = cy = size / 2
    r = size * 0.40

    # SVG кути (за годинниковою від правого = 0°)
    START_SVG = 140.0   # ~7:20 (нижній лівий)
    SWEEP     = 240.0   # градусів за годинниковою

    def pt(svg_deg):
        a = math.radians(svg_deg)
        return (round(cx + r * math.cos(a), 2), round(cy + r * math.sin(a), 2))

    s  = pt(START_SVG)               # 0%
    e  = pt(START_SVG + SWEEP)       # 100%  (= 380° mod 360° = 20°)

    val_sweep = value / 100.0 * SWEEP
    v = pt(START_SVG + val_sweep)

    # Фонова дуга (ціла 240°, CW, large-arc=1, sweep=1)
    bg = f"M{s[0]},{s[1]} A{r:.1f},{r:.1f} 0 1,1 {e[0]},{e[1]}"

    # Активна дуга (0 → value)
    if value > 0:
        vl = 1 if val_sweep > 180 else 0
        fg = f"M{s[0]},{s[1]} A{r:.1f},{r:.1f} 0 {vl},1 {v[0]},{v[1]}"
    else:
        fg = None

    # Текст по центру дуги (зсунутий трохи вниз)
    tx = round(cx, 1)
    ty = round(cy + r * 0.22, 1)

    return (
        f'<svg viewBox="0 0 {size} {size}" width="{size}" height="{size}" '
        f'xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;display:block">'
        # Фонова дуга
        f'<path d="{bg}" stroke="#252d48" stroke-width="3" fill="none" stroke-linecap="round"/>'
        # Активна дуга
        + (f'<path d="{fg}" stroke="{color}" stroke-width="3" fill="none" stroke-linecap="round"/>' if fg else '')
        # Мітка на поточній позиції
        + f'<circle cx="{v[0]}" cy="{v[1]}" r="4" fill="{color}"/>'
        # Крапка початку (0%)
        f'<circle cx="{s[0]}" cy="{s[1]}" r="2.5" fill="#252d48" stroke="{color}" stroke-width="1"/>'
        # Число по центру
        f'<text x="{tx}" y="{ty + 5}" text-anchor="middle" dominant-baseline="middle" '
        f'font-family="Courier New,monospace" font-size="15" font-weight="bold" fill="{color}">'
        f'{int(round(value))}</text>'
        f'</svg>'
    )


def gauge_color(value, oi=False):
    """Колір gauge: 0-15% = червоний, 15-85% = білий, 85-100% = зелений. OI — завжди білий."""
    if oi: return '#dde2ee'
    v = float(value)
    if v < 15:  return '#f0515a'
    if v > 85:  return '#20d483'
    return '#dde2ee'


# ================================================================
# 🎨  ТАБЛИЦЯ — єдина таблиця для синхронізації колонок MM + Data
# ================================================================
def intensity_bg(val, max_abs):
    if max_abs==0: return ''
    try:
        v=float(val); ratio=min(abs(v)/max_abs,1.0); op=0.10+ratio*0.67
        if v>0:   return f' style="background:rgba(32,212,131,{op:.2f})"'
        elif v<0: return f' style="background:rgba(240,81,90,{op:.2f})"'
    except: pass
    return ''


def make_hist_table(hist, stats_ls, stats_cm, stats_st, stats_oi, sm):
    """
    Єдина таблиця: MAX/MIN рядки у першому tbody,
    дані — у другому tbody. Колонки синхронізовані через спільний colgroup.
    OI — лише одна колонка (All time).
    """
    n = len(hist['dates'])
    if n == 0:
        return '<p style="padding:12px;color:#8090b0">Немає даних</p>'

    def maxabs(lst):
        vals = [abs(v) for v in lst if v != 0]
        return max(vals) if vals else 1

    m_ls_cl = maxabs(hist['ls_cl']); m_ls_cs = maxabs(hist['ls_cs'])
    m_cm_cl = maxabs(hist['cm_cl']); m_cm_cs = maxabs(hist['cm_cs'])
    m_st_cl = maxabs(hist['st_cl']); m_st_cs = maxabs(hist['st_cs'])

    # Кольори для заголовків
    ls_bg  = 'background:rgba(74,158,255,.18);color:#fff;border-left:2px solid rgba(74,158,255,.7)'
    cm_bg  = 'background:rgba(32,212,131,.15);color:#fff;border-left:2px solid rgba(32,212,131,.7)'
    st_bg  = 'background:rgba(240,81,90,.15);color:#fff;border-left:2px solid rgba(240,81,90,.7)'
    oi_bg  = 'background:rgba(160,170,192,.1);color:#fff'
    sm_bg  = 'background:var(--bg3);color:var(--d)'
    sub_plain = 'background:var(--bg3);color:var(--d)'
    sub_ls = f'{sub_plain};border-left:2px solid rgba(74,158,255,.5)'
    sub_cm = f'{sub_plain};border-left:2px solid rgba(32,212,131,.5)'
    sub_st = f'{sub_plain};border-left:2px solid rgba(240,81,90,.5)'

    # Colgroup — 14 колонок (ДАТА + LS×3 + CM×3 + ST×3 + OI + SM×3)
    colgroup = (
        '<colgroup>'
        '<col style="width:84px">'    # 0: ДАТА / лейбл MM
        '<col><col><col>'              # 1-3: LS CHG L / CHG S / NET
        '<col><col><col>'              # 4-6: CM CHG L / CHG S / NET
        '<col><col><col>'              # 7-9: ST CHG L / CHG S / NET
        '<col style="width:82px">'    # 10: OI
        '<col style="width:50px">'    # 11: SM DIV
        '<col style="width:50px">'    # 12: SM 6M
        '<col style="width:50px">'    # 13: SM 3M
        '</colgroup>'
    )

    # Thead — спільний для MM та Data секцій
    thead = (
        '<thead>'
        # Рядок 1: групові назви
        '<tr class="th-row1">'
        '<th class="th-corner"></th>'
        f'<th colspan="3" class="th-group" style="{ls_bg}">LARGE SPECULATORS</th>'
        f'<th colspan="3" class="th-group" style="{cm_bg}">COMMERCIALS</th>'
        f'<th colspan="3" class="th-group" style="{st_bg}">SMALL TRADERS</th>'
        f'<th class="th-group" style="{oi_bg}">OI</th>'
        f'<th colspan="3" class="th-group sm-th-group" style="{sm_bg}">SM DIV</th>'
        '</tr>'
        # Рядок 2: підзаголовки колонок
        '<tr class="th-row2">'
        '<th class="th-date th-left">ДАТА</th>'
        f'<th style="{sub_ls}">CHG L</th>'
        f'<th style="{sub_plain}">CHG S</th>'
        f'<th class="sep-r" style="{sub_plain}">NET POS</th>'
        f'<th style="{sub_cm}">CHG L</th>'
        f'<th style="{sub_plain}">CHG S</th>'
        f'<th class="sep-r" style="{sub_plain}">NET POS</th>'
        f'<th style="{sub_st}">CHG L</th>'
        f'<th style="{sub_plain}">CHG S</th>'
        f'<th class="sep-r" style="{sub_plain}">NET POS</th>'
        f'<th style="{oi_bg}" class="th-oi">All</th>'
        '<th class="sm-th">All</th><th class="sm-th">6M</th><th class="sm-th">3M</th>'
        '</tr>'
        '</thead>'
    )

    # Хелпери для MM рядків
    def mm_v(v, cls=''):
        return f'<td class="mm-val {cls}">{fv_full(v, sign=True)}</td>'

    def sm_mini(v):
        if v is None: return '<td class="sm-td d">–</td>'
        try:
            f2 = float(v)
            c = 'g' if f2 > 0 else ('r' if f2 < 0 else 'd')
            return f'<td class="sm-td {c}">{f2:+.2f}</td>'
        except:
            return '<td class="sm-td d">–</td>'

    def grp(st, key, col):
        """3 клітинки: CHG L, CHG S, NET для однієї групи"""
        cl_d = st.get('cl') or {}
        cs_d = st.get('cs') or {}
        return (
            mm_v(cl_d.get(key, 0), col) +
            mm_v(cs_d.get(key, 0), col) +
            mm_v(st.get(key, 0), col)
        )

    # MM рядки (окремий tbody)
    mm_rows = (
        '<tr class="mm-row">'
        '<td class="mm-lbl">MAX</td>'
        + grp(stats_ls,'max_all','g') + grp(stats_cm,'max_all','g') + grp(stats_st,'max_all','g')
        + mm_v(stats_oi['max_all'], '')
        + sm_mini(sm['div']) + sm_mini(sm['div_6m']) + sm_mini(sm['div_3m'])
        + '</tr>'

        '<tr class="mm-row">'
        '<td class="mm-lbl">MIN</td>'
        + grp(stats_ls,'min_all','r') + grp(stats_cm,'min_all','r') + grp(stats_st,'min_all','r')
        + mm_v(stats_oi['min_all'], '')
        + '<td class="sm-td"></td><td class="sm-td"></td><td class="sm-td"></td>'
        + '</tr>'

        '<tr class="mm-row mm-yr">'
        '<td class="mm-lbl">MAX 1Y</td>'
        + grp(stats_ls,'max_1y','g') + grp(stats_cm,'max_1y','g') + grp(stats_st,'max_1y','g')
        + mm_v(stats_oi['max_1y'], '')
        + '<td class="sm-td"></td><td class="sm-td"></td><td class="sm-td"></td>'
        + '</tr>'

        '<tr class="mm-row mm-yr">'
        '<td class="mm-lbl">MIN 1Y</td>'
        + grp(stats_ls,'min_1y','r') + grp(stats_cm,'min_1y','r') + grp(stats_st,'min_1y','r')
        + mm_v(stats_oi['min_1y'], '')
        + '<td class="sm-td"></td><td class="sm-td"></td><td class="sm-td"></td>'
        + '</tr>'
    )
    mm_tbody = f'<tbody class="mm-tbody">{mm_rows}</tbody>'

    # Дані рядки (окремий tbody)
    def td_chg(v, maxv):
        cls = cc(v)
        txt = fv_full(v, sign=True) if v != 0 else '—'
        return f'<td class="{cls}"{intensity_bg(v, maxv)}>{txt}</td>'

    def td_net(v, extra=''):
        return f'<td class="{cc(v)}{extra}">{fv_full(v, sign=True)}</td>'

    def sm_cell(v):
        if v is None or (isinstance(v, float) and v != v):
            return '<td class="sm-td d"></td>'
        try:
            f2 = float(v)
            if f2 == 0: return '<td class="sm-td d"></td>'
            c = 'g' if f2 > 0 else 'r'
            return f'<td class="sm-td {c}">{f2:+.3f}</td>'
        except:
            return '<td class="sm-td d"></td>'

    sm_div_val    = hist.get('sm_div_row',    [None]*n)
    sm_div_6m_val = hist.get('sm_div_6m_row', [None]*n)
    sm_div_3m_val = hist.get('sm_div_3m_row', [None]*n)

    data_rows = []
    for i in range(n-1, -1, -1):
        row_idx = n-1-i
        row = (
            f'<tr data-row="{row_idx}">'
            f'<td class="date-col">{hist["dates"][i]}</td>'
            + td_chg(hist['ls_cl'][i], m_ls_cl)
            + td_chg(hist['ls_cs'][i], m_ls_cs)
            + td_net(hist['ls_net'][i], ' sep-r')
            + td_chg(hist['cm_cl'][i], m_cm_cl)
            + td_chg(hist['cm_cs'][i], m_cm_cs)
            + td_net(hist['cm_net'][i], ' sep-r')
            + td_chg(hist['st_cl'][i], m_st_cl)
            + td_chg(hist['st_cs'][i], m_st_cs)
            + td_net(hist['st_net'][i], ' sep-r')
            + f'<td class="t">{fv_full(hist["oi"][i])}</td>'
            + sm_cell(sm_div_val[i])
            + sm_cell(sm_div_6m_val[i])
            + sm_cell(sm_div_3m_val[i])
            + '</tr>'
        )
        data_rows.append(row)

    data_tbody = f'<tbody class="data-tbody">{"".join(data_rows)}</tbody>'

    return (
        '<table class="ht">'
        + colgroup + thead + mm_tbody + data_tbody
        + '</table>'
    )


def sm_bar(val, label):
    v=float(val) if val else 0.0
    pct=min(max((v+1)/2*100,0),100)
    color='#20d483'if v>0 else('#f0515a'if v<0 else'#6a7290')
    cls='g'if v>0 else('r'if v<0 else'd')
    return(
        f'<div class="sm-row">'
        f'<div class="sm-lbl">{label}</div>'
        f'<div class="sm-bar-bg"><div class="sm-mk" style="left:{pct:.0f}%;background:{color}"></div></div>'
        f'<div class="sm-val {cls}">{v:+.2f}</div>'
        f'</div>'
    )


def make_reports_panel(s_id):
    rows = []
    for rpt in REPORTS:
        rid = rpt['id']
        rel = get_relevance(s_id, rid)
        tag = f'<span class="rpt-tag">{rpt["tag"]}</span>' if rpt['tag'] else ''
        if rel == 'direct':
            icon = '<span class="rel-icon rel-d" title="Прямий вплив">●</span>'
        elif rel == 'indirect':
            icon = '<span class="rel-icon rel-i" title="Непрямий / через COT">■</span>'
        else:
            icon = '<span class="rel-icon rel-n" title="Не впливає">○</span>'
        row_cls = 'rpt-row' if rel != 'none' else 'rpt-row rpt-dim'
        rows.append(
            f'<div class="{row_cls}" id="rpt_{s_id}_{rid}">'
            f'<div class="rpt-rel">{icon}</div>'
            f'<div class="rpt-info">'
            f'<div class="rpt-name">{rpt["name"]}{tag}</div>'
            f'<div class="rpt-sched">{rpt["sched"]}</div>'
            f'</div>'
            f'<div class="rpt-btns">'
            f'<button class="rb rb-l" onclick="setRpt(\'{s_id}\',\'{rid}\',\'long\')" title="LONG">L</button>'
            f'<button class="rb rb-n active" onclick="setRpt(\'{s_id}\',\'{rid}\',\'neutral\')" title="NEUTRAL">N</button>'
            f'<button class="rb rb-s" onclick="setRpt(\'{s_id}\',\'{rid}\',\'short\')" title="SHORT">S</button>'
            f'</div>'
            f'</div>'
        )
    return (
        f'<div class="panel rpt-panel" id="rpts_{s_id}">'
        f'<div class="plbl-row">'
        f'<div class="plbl" style="margin:0">ЗВІТИ</div>'
        f'<div class="rel-legend">'
        f'<span class="rel-icon rel-d">●</span>прямий&nbsp;'
        f'<span class="rel-icon rel-i">■</span>непрямий'
        f'</div></div>'
        + ''.join(rows)
        + f'</div>'
    )


def make_metric_card(lbl, val, chg, chg_pct, spark_series, spark_color,
                     oi=False, gauge_val=50.0, sub_text='COT Index: —'):
    """
    Картка метрики:
    - Лейбл вгорі зліва (повернуто)
    - Велике число ліворуч
    - Gauge індикатор праворуч
    - Sparkline внизу
    """
    spark = make_sparkline(spark_series, spark_color)
    try: chg_int = int(round(float(chg)))
    except: chg_int = 0
    chg_cls = cc(chg)
    if chg_int > 0:
        bg = 'background:rgba(32,212,131,.20);border-radius:3px;padding:2px 7px;display:inline-block;'
    elif chg_int < 0:
        bg = 'background:rgba(240,81,90,.20);border-radius:3px;padding:2px 7px;display:inline-block;'
    else:
        bg = 'padding:2px 7px;display:inline-block;'
    # Значення
    if oi:
        val_str = f'<span class="t">{fv(val)}</span>'
    else:
        val_str = f'<span class="{cc(val)}">{fv(val, sign=True)}</span>'
    # Gauge
    gcol = gauge_color(gauge_val, oi=oi)
    g_svg = make_gauge_svg(gauge_val, gcol)
    return (
        f'<div class="mc">'
        # Лейбл — зліва вгорі (повернуто до попереднього стилю)
        f'<div class="mc-lbl">{lbl}</div>'
        f'<div class="mc-inner">'
        f'<div class="mc-left">'
        f'<div class="mc-val">{val_str}</div>'
        f'<div class="mc-chg-wrap"><span class="{chg_cls}" style="{bg}">'
        f'{ar(chg)} {fv_full(abs(chg_int))}'
        f'<span class="mc-wtag"> за тиждень</span></span></div>'
        f'<div class="mc-pct {chg_cls}">{chg_pct}</div>'
        f'<div class="mc-sub">{sub_text}</div>'
        f'</div>'
        f'<div class="mc-right">{g_svg}</div>'
        f'</div>'
        f'{spark}'
        f'</div>'
    )


def analysis_row(group_label, group_color, net, cl, cs, chg, chg_pct):
    """
    Панель аналізу: прибрано 'ЗМІНА ЗА ТИЖДЕНЬ',
    збільшено шрифти підписів.
    """
    dc = 'g' if net > 0 else 'r'
    return (
        f'<div class="arow"><div class="arow-body">'
        f'<div class="arow-left">'
        # CHG LONG і CHG SHORT в 2 колонки (без рядка "ЗМІНА ЗА ТИЖДЕНЬ")
        f'<div class="arow-grid2">'
        f'<div class="ag-item">'
        f'<span class="ag-lbl">CHG LONG</span>'
        f'<span class="{cc(cl)} ag-val">{fv_full(cl, sign=True)}</span>'
        f'</div>'
        f'<div class="ag-item">'
        f'<span class="ag-lbl">CHG SHORT</span>'
        f'<span class="{cc(cs)} ag-val">{fv_full(cs, sign=True)}</span>'
        f'</div>'
        f'</div>'
        # Δ NET — окремий рядок
        f'<div class="arow-dnet">'
        f'<span class="ag-lbl">Δ NET</span>'
        f'<span class="{cc(chg)} ag-val-net">{fv_full(chg, sign=True)}'
        f'<span class="ag-pct"> ({chg_pct})</span></span>'
        f'</div>'
        f'</div>'
        f'<div class="arow-right">'
        f'<div class="ar-glbl" style="color:{group_color}">{group_label}</div>'
        f'<div class="ar-net {dc}">{fv_full(net, sign=True)}</div>'
        f'</div>'
        f'</div></div>'
    )


def make_instrument_view(d):
    c=d['cur']; s=d['sid']; sm=d['sm']

    mc_ls=make_metric_card('LARGE SPEC (NETTO)',    c['ls_net'],c['ls_chg'],c['ls_chg_pct'],
                           d['spark']['ls'],COLOR_LS,
                           gauge_val=d['cot_idx']['ls']['all'],
                           sub_text=f"COT Index: {fp(d['cot_idx']['ls']['all'])}")
    mc_cm=make_metric_card('COMMERCIALS (NETTO)',   c['cm_net'],c['cm_chg'],c['cm_chg_pct'],
                           d['spark']['cm'],COLOR_CM,
                           gauge_val=d['cot_idx']['cm']['all'],
                           sub_text=f"COT Index: {fp(d['cot_idx']['cm']['all'])}")
    mc_st=make_metric_card('SMALL TRADERS (NETTO)',c['st_net'],c['st_chg'],c['st_chg_pct'],
                           d['spark']['st'],COLOR_ST,
                           gauge_val=d['cot_idx']['st']['all'],
                           sub_text=f"COT Index: {fp(d['cot_idx']['st']['all'])}")
    mc_oi=make_metric_card('OPEN INTEREST',         c['oi'],    c['oi_chg'],c['oi_chg_pct'],
                           d['spark']['oi'],'#a0aac0',oi=True,
                           gauge_val=d.get('oi_capacity', 50.0),
                           sub_text=f"зміна: {fv(int(c['oi_chg']),True,sign=True)}")

    analysis_panel=(
        f'<div class="panel">'
        +analysis_row('LARGE SPEC', COLOR_LS,c['ls_net'],c['ls_cl'],c['ls_cs'],c['ls_chg'],c['ls_chg_pct'])
        +analysis_row('COMMERCIALS',COLOR_CM,c['cm_net'],c['cm_cl'],c['cm_cs'],c['cm_chg'],c['cm_chg_pct'])
        +f'</div>'
    )
    sm_panel=(
        f'<div class="panel sm-panel">'
        f'<div class="plbl">SM DIVERGENCE</div>'
        +sm_bar(sm['div'],'All time')+sm_bar(sm['div_6m'],'6 months')+sm_bar(sm['div_3m'],'3 months')
        +f'<div class="sm-hint">+ бичача &nbsp;|&nbsp; − ведмежа</div>'
        +f'</div>'
    )
    reports_panel = make_reports_panel(s)

    ini=d['cot_idx']['ls']['all']
    ini_pos=min(max(ini,0),100)
    ini_color='#f0515a'if ini<15 else('#20d483'if ini>85 else'#dde2ee')
    ini_label='екстрем. шорт'if ini<15 else('екстрем. лонг'if ini>85 else'нейтральна зона')
    cot_json=json.dumps(d['cot_idx'],ensure_ascii=False)
    pct_panel=(
        f'<div class="panel pct-panel">'
        f'<div class="plbl">ПЕРЦЕНТИЛЬ (COT INDEX)</div>'
        f'<div class="pct-sel-row">'
        f'<div class="psel-group">'
        f'<button class="psel active" data-p="ls" onclick="pctSel(this,\'{s}\')">LS</button>'
        f'<button class="psel" data-p="cm" onclick="pctSel(this,\'{s}\')">CM</button>'
        f'<button class="psel" data-p="st" onclick="pctSel(this,\'{s}\')">ST</button>'
        f'</div><div class="psel-sep"></div>'
        f'<div class="psel-group">'
        f'<button class="pper active" data-per="all" onclick="pperSel(this,\'{s}\')">All</button>'
        f'<button class="pper" data-per="3y" onclick="pperSel(this,\'{s}\')">3Y</button>'
        f'<button class="pper" data-per="1y" onclick="pperSel(this,\'{s}\')">1Y</button>'
        f'<button class="pper" data-per="6m" onclick="pperSel(this,\'{s}\')">6M</button>'
        f'<button class="pper" data-per="3m" onclick="pperSel(this,\'{s}\')">3M</button>'
        f'</div></div>'
        f'<div class="pct-val-row">'
        f'<span id="pctval_{s}" style="font-size:16px;font-weight:bold;color:{ini_color}">{fp(ini)}</span>'
        f'<span id="pctlbl_{s}" style="font-size:11px;color:#8090b0;margin-left:8px;">— {ini_label}</span>'
        f'</div>'
        f'<div class="pbar-wrap"><div class="pbar-bg">'
        f'<div class="pbar-lo"></div><div class="pbar-hi"></div>'
        f'<div class="ptick" style="left:15%"></div><div class="ptick" style="left:85%"></div>'
        f'<div class="pbar-mk" id="pctmk_{s}" style="left:{ini_pos:.1f}%"></div>'
        f'</div>'
        f'<div class="ptick-labels">'
        f'<span class="ptlbl" style="left:15%">15%</span>'
        f'<span id="pctcur_{s}" class="ptlbl ptlbl-cur" style="left:{ini_pos:.1f}%">{fp(ini)}</span>'
        f'<span class="ptlbl" style="left:85%">85%</span>'
        f'</div></div>'
        f'<div class="pbar-lb"><span>0%</span><span>50%</span><span>100%</span></div>'
        f'</div>'
        f'<script>_ci["{s}"]={cot_json};</script>'
    )

    chart_json=json.dumps(d['chart'],ensure_ascii=False)
    chart_block=(
        f'<div class="chartbox">'
        f'<div class="chartbox-hdr">'
        f'<div class="plbl" style="margin:0">ЧИСТІ ПОЗИЦІЇ</div>'
        f'<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">'
        f'<div class="period-btns">'
        f'<button class="per-btn active" data-per="1y" onclick="setChartPer(this,\'{s}\')">1 рік</button>'
        f'<button class="per-btn" data-per="3y" onclick="setChartPer(this,\'{s}\')">3 роки</button>'
        f'<button class="per-btn" data-per="5y" onclick="setChartPer(this,\'{s}\')">5 років</button>'
        f'</div>'
        f'<div class="chart-leg">'
        f'<span><span class="ll" style="background:{COLOR_LS}"></span>Large Spec</span>'
        f'<span><span class="ll" style="background:{COLOR_CM}"></span>Commercials</span>'
        f'<span class="ll-dash" style="border-top-color:{COLOR_ST}"></span>'
        f'<span style="margin-left:4px">Small Traders</span>'
        f'</div></div></div>'
        f'<div class="cw"><canvas id="cv_{s}"></canvas></div>'
        f'</div>'
        f'<script>_cd["{s}"]={chart_json};</script>'
    )
    bar_block=(
        f'<div class="bar-charts-grid chartbox">'
        f'<div class="bar-wrap"><div class="bar-lbl" style="color:{COLOR_LS}">LARGE SPEC — WEEKLY ΔNet</div>'
        f'<div class="bar-cw"><canvas id="barcv_ls_{s}"></canvas></div></div>'
        f'<div class="bar-wrap"><div class="bar-lbl" style="color:{COLOR_CM}">COMMERCIALS — WEEKLY ΔNet</div>'
        f'<div class="bar-cw"><canvas id="barcv_cm_{s}"></canvas></div></div>'
        f'<div class="bar-wrap"><div class="bar-lbl" style="color:{COLOR_ST}">SMALL TRADERS — WEEKLY ΔNet</div>'
        f'<div class="bar-cw"><canvas id="barcv_st_{s}"></canvas></div></div>'
        f'</div>'
    )

    # Єдина таблиця з синхронізованими колонками
    table_html = make_hist_table(
        d['hist'], d['stats_ls'], d['stats_cm'], d['stats_st'], d['stats_oi'], sm
    )
    table_block=(
        f'<div class="htable-wrap">'
        f'<div class="htable-hdr"><span>ТИЖНЕВА СТАТИСТИКА ПОЗИЦІЙ</span>'
        f'<div class="hsel">'
        f'<button class="hbtn active" data-n="10" onclick="setHist(this,\'{s}\')">10</button>'
        f'<button class="hbtn" data-n="26" onclick="setHist(this,\'{s}\')">26</button>'
        f'<button class="hbtn" data-n="52" onclick="setHist(this,\'{s}\')">52</button>'
        f'</div></div>'
        f'<div class="htable-scroll">{table_html}</div>'
        f'</div>'
    )

    mid = f'<div class="mid">{analysis_panel}{sm_panel}{reports_panel}{pct_panel}</div>'

    return(
        f'<div class="iview" id="iv_{s}" data-sid="{s}">'
        f'<div class="report-tabs">'
        f'<span class="rtab-lbl">ТИП ЗВІТУ:</span>'
        f'<button class="rtab active">Legacy Report</button>'
        f'<button class="rtab disabled" title="Файл не завантажено">TFF Report</button>'
        f'<button class="rtab disabled" title="Файл не завантажено">Disaggregated</button>'
        f'</div>'
        +f'<div class="mcards">{mc_ls}{mc_cm}{mc_st}{mc_oi}</div>'
        +mid
        +bar_block+chart_block+table_block
        +f'</div>'
    )


def make_overview_tab():
    rows_html=[]; rep_date='—'; today_date='—'
    for item in OVERVIEW_TABLE:
        if isinstance(item,tuple) and item[0]=='_meta':
            rep_date=item[1]; today_date=item[2]; continue
        if isinstance(item,tuple) and item[0]=='_group':
            rows_html.append(f'<tr class="ov-group"><td colspan="12">{item[1]}</td></tr>')
            continue
        if not isinstance(item,dict): continue
        d=item
        def fnum(v,sign=False,pct=False):
            if v is None: return '<span class="d">—</span>'
            if pct:
                v2=v*100 if abs(v)<=1.5 else v
                s='+' if(sign and v2>0)else ''
                cls='g' if v2>0 else('r' if v2<0 else 'd')
                return f'<span class="{cls}">{s}{v2:.1f}%</span>'
            try:
                nv=int(round(float(v))); body=f"{abs(nv):,}".replace(',','\u202f')
                s='+' if(sign and nv>0)else('-' if nv<0 else '')
                cls='g' if nv>0 else('r' if nv<0 else 'd')
                return f'<span class="{cls}">{s}{body}</span>'
            except: return '<span class="d">—</span>'
        def pct_bar(v,lo=0.15,hi=0.85):
            if v is None: return ''
            pct=min(max(v/100,0),1)
            color='#f0515a' if pct<lo else('#20d483' if pct>hi else '#4a9eff')
            return(f'<div class="ov-bar-bg">'
                   f'<div class="ov-bar-fill" style="width:{pct*100:.1f}%;background:{color}"></div></div>')
        def sm_fmt(v):
            if v is None: return '<span class="d">—</span>'
            cls='g' if float(v)>0 else('r' if float(v)<0 else 'd')
            return f'<span class="{cls}">{float(v):+.2f}</span>'
        rows_html.append(
            f'<tr class="ov-row">'
            f'<td class="ov-asset">{d["asset"]}</td>'
            f'<td>{fnum(d["net_ls"],sign=True)}</td>'
            f'<td>{fnum(d["chg_ls"],sign=True)}</td>'
            f'<td>{fnum(d["net_cm"],sign=True)}</td>'
            f'<td>{fnum(d["chg_cm"],sign=True)}</td>'
            f'<td>{fnum(d["oi_chg_pct"],pct=True)}</td>'
            f'<td><div class="ov-cot-cell">{pct_bar(d["cot_ls"])}'
            f'<span class="ov-cot-val">{d["cot_ls"]:.0f}%</span></div></td>'
            f'<td><div class="ov-cot-cell">{pct_bar(d["cot_cm"])}'
            f'<span class="ov-cot-val">{d["cot_cm"]:.0f}%</span></div></td>'
            f'<td><div class="ov-cot-cell">{pct_bar(d["cot_st"])}'
            f'<span class="ov-cot-val">{d["cot_st"]:.0f}%</span></div></td>'
            f'<td>{sm_fmt(d["sm_div"])}</td>'
            f'<td>{sm_fmt(d["sm_div_6m"])}</td>'
            f'<td>{sm_fmt(d["sm_div_3m"])}</td>'
            f'</tr>'
        )
    thead=(
        '<thead><tr>'
        '<th class="ov-asset">ASSET</th>'
        '<th>NET LS</th><th>CHG LS</th>'
        '<th>NET CM</th><th>CHG CM</th><th>%OI CHG</th>'
        f'<th style="color:{COLOR_LS}">COT LS</th>'
        f'<th style="color:{COLOR_CM}">COT CM</th>'
        f'<th style="color:{COLOR_ST}">COT ST</th>'
        '<th>SM DIV</th><th>SM 6M</th><th>SM 3M</th>'
        '</tr></thead>'
    )
    return(
        f'<div class="ov-meta">Звіт: <b>{rep_date}</b> &nbsp;|&nbsp; Оновлено: {today_date}</div>'
        f'<div class="ov-scroll">'
        f'<table class="ov-table">{thead}<tbody>{"".join(rows_html)}</tbody></table>'
        f'</div>'
    )


def generate_html(data):
    all_dates=[d['cur']['date'] for d in data.values()]
    report_date=max(all_dates) if all_dates else '—'
    updated=datetime.now().strftime('%d.%m.%Y %H:%M')

    cat_tabs=[]; cat_sects=[]; first_cat=True
    for cat,instruments in CATEGORIES.items():
        available=[i for i in instruments if i in data]
        if not available: continue
        act=' active'if first_cat else''
        cat_tabs.append(
            f'<button class="ctab{act}" data-c="{cat}" onclick="selCat(\'{cat}\')">'
            f'{cat}<span class="tc">({len(available)})</span></button>'
        )
        inst_btns=''.join(
            f'<button class="itab" data-cat="{cat}" data-i="{i}"'
            f' onclick="selInst(\'{cat}\',\'{i}\')">{disp(i)}</button>'
            for i in available
        )
        views=''.join(make_instrument_view(data[i]) for i in available)
        cat_sects.append(
            f'<div class="catsec{act}" id="cs_{cat}">'
            f'<div class="itabs" id="itabs_{cat}">{inst_btns}</div>'
            f'<div class="iviews" id="iv_{cat}">{views}</div>'
            f'</div>'
        )
        first_cat=False

    ov_html=make_overview_tab()

    # DASHBOARD — перші 3 букви синього, 3 зеленого, 3 червоного
    dashboard_text = (
        '<span class="dash-b">DAS</span>'
        '<span class="dash-g">HBO</span>'
        '<span class="dash-r">ARD</span>'
    )

    return(
        HTML_HEAD
        +f'<header class="hdr">'
        +f'<div><div class="hdr-t">COT {dashboard_text}</div>'
        +f'<div class="hdr-s">COMMITMENTS OF TRADERS — CFTC LEGACY REPORT</div></div>'
        +f'<div style="display:flex;align-items:center;gap:16px">'
        +f'<div class="hdr-r">Звіт: <b>{report_date}</b><br>Оновлено: {updated}</div>'
        +f'<button class="auth-btn" id="authBtn" onclick="openAuth()">УВІЙТИ</button>'
        +f'</div></header>'
        +f'<div class="main-tabs">'
        +f'<button class="mtab active" data-mt="cot" onclick="selMain(\'cot\')">COT Dashboard</button>'
        +f'<button class="mtab" data-mt="ov" onclick="selMain(\'ov\')">Overview</button>'
        +f'</div>'
        +f'<div class="main-sec active" id="ms_cot">'
        +f'<div class="ctabs">'+''.join(cat_tabs)+f'</div>'
        +''.join(cat_sects)
        +f'</div>'
        +f'<div class="main-sec" id="ms_ov"><div style="padding:16px 24px">{ov_html}</div></div>'
        +f'<div class="footer">COT DASHBOARD &bull; CFTC LEGACY &bull; {updated}</div>'
        +AUTH_MODAL_HTML+HTML_FOOT
    )


AUTH_MODAL_HTML = """
<div class="auth-overlay" id="authOverlay" onclick="if(event.target===this)closeAuth()">
  <div class="auth-box">
    <button class="auth-close" onclick="closeAuth()">✕</button>
    <div class="auth-tabs">
      <div class="auth-tab active" onclick="authTab('login')">ВХІД</div>
      <div class="auth-tab" onclick="authTab('reg')">РЕЄСТРАЦІЯ</div>
    </div>
    <div id="auth-logged" style="display:none">
      <div class="auth-user">Ви увійшли як <b id="auth-email-display"></b></div>
      <button class="auth-submit" style="background:var(--r);margin-top:8px" onclick="doLogout()">ВИЙТИ</button>
    </div>
    <div id="auth-loggedout">
      <div id="at-login">
        <div class="auth-field"><label>EMAIL</label><input type="email" id="al-email" placeholder="your@email.com"></div>
        <div class="auth-field"><label>ПАРОЛЬ</label><input type="password" id="al-pass" placeholder="••••••••" onkeydown="if(event.key==='Enter')doAuthLogin()"></div>
        <button class="auth-submit" onclick="doAuthLogin()">УВІЙТИ</button>
        <div class="auth-msg" id="al-msg"></div>
      </div>
      <div id="at-reg" style="display:none">
        <div class="auth-field"><label>EMAIL</label><input type="email" id="ar-email" placeholder="your@email.com"></div>
        <div class="auth-field"><label>ПАРОЛЬ</label><input type="password" id="ar-pass" placeholder="мінімум 6 символів"></div>
        <div class="auth-field"><label>ПАРОЛЬ ЩЕ РАЗ</label><input type="password" id="ar-pass2" placeholder="••••••••" onkeydown="if(event.key==='Enter')doAuthReg()"></div>
        <button class="auth-submit" onclick="doAuthReg()">ЗАРЕЄСТРУВАТИСЬ</button>
        <div class="auth-msg" id="ar-msg"></div>
      </div>
    </div>
  </div>
</div>
"""

HTML_HEAD = f"""<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>COT Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{{
  --bg:#1a1e2d;--bg2:#21263a;--bg3:#282f47;--bd:#343d5a;
  --g:#20d483;--r:#f0515a;--b:#4a9eff;
  --accent:#f59420;
  --t:#dde2ee;--d:#8090b0;
  --f:'Courier New',Courier,monospace;
  --hdr-h:50px;
}}
*,*::before,*::after{{margin:0;padding:0;box-sizing:border-box;}}
html,body{{background:var(--bg);color:var(--t);font-family:var(--f);font-size:13px;}}
.t{{color:var(--t);}}

/* ── HEADER ── */
.hdr{{height:var(--hdr-h);padding:0 24px;background:var(--bg2);
  border-bottom:1px solid var(--bd);display:flex;align-items:center;
  justify-content:space-between;position:sticky;top:0;z-index:300;}}
.hdr-t{{font-size:17px;font-weight:bold;color:#fff;letter-spacing:2px;}}
/* DASHBOARD кольорові букви */
.dash-b{{color:#4a9eff;font-style:normal;}}
.dash-g{{color:#20d483;font-style:normal;}}
.dash-r{{color:#f0515a;font-style:normal;}}
.hdr-s{{font-size:10px;color:var(--d);margin-top:2px;letter-spacing:1px;}}
.hdr-r{{text-align:right;font-size:11px;color:var(--d);line-height:2;}}
.hdr-r b{{color:var(--t);}}

.auth-btn{{padding:6px 16px;border:1px solid var(--g);border-radius:3px;
  background:transparent;color:var(--g);font-family:var(--f);font-size:11px;
  cursor:pointer;letter-spacing:1px;transition:all .15s;}}
.auth-btn:hover{{background:rgba(32,212,131,.15);}}
.auth-btn.logged{{border-color:var(--b);color:var(--b);}}

/* ── AUTH ── */
.auth-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);
  z-index:1000;align-items:center;justify-content:center;}}
.auth-overlay.open{{display:flex;}}
.auth-box{{background:var(--bg2);border:1px solid var(--bd);border-radius:8px;padding:32px;width:320px;position:relative;}}
.auth-close{{position:absolute;top:12px;right:14px;background:none;border:none;color:var(--d);cursor:pointer;font-size:16px;font-family:var(--f);}}
.auth-tabs{{display:flex;margin-bottom:20px;border-bottom:1px solid var(--bd);}}
.auth-tab{{flex:1;padding:7px;text-align:center;cursor:pointer;font-size:11px;color:var(--d);letter-spacing:1px;border-bottom:2px solid transparent;margin-bottom:-1px;}}
.auth-tab.active{{color:#fff;border-bottom-color:var(--g);}}
.auth-field{{margin-bottom:12px;}}
.auth-field label{{display:block;font-size:9px;color:var(--d);letter-spacing:.8px;margin-bottom:4px;}}
.auth-field input{{width:100%;background:var(--bg);border:1px solid var(--bd);border-radius:4px;padding:9px 11px;color:var(--t);font-family:var(--f);font-size:12px;outline:none;}}
.auth-field input:focus{{border-color:var(--g);}}
.auth-submit{{width:100%;padding:10px;background:var(--g);color:#000;border:none;border-radius:4px;cursor:pointer;font-family:var(--f);font-size:12px;font-weight:bold;margin-top:4px;}}
.auth-msg{{font-size:11px;padding:7px 10px;border-radius:4px;margin-top:10px;text-align:center;display:none;}}
.auth-msg.err{{background:rgba(240,81,90,.15);border:1px solid var(--r);color:var(--r);}}
.auth-msg.ok{{background:rgba(32,212,131,.15);border:1px solid var(--g);color:var(--g);}}
.auth-user{{font-size:11px;color:var(--d);text-align:center;margin-bottom:12px;padding:10px;background:var(--bg3);border-radius:4px;}}
.auth-user b{{color:var(--t);}}

/* ── MAIN TABS ── */
.main-tabs{{display:flex;gap:0;padding:0 24px;background:var(--bg2);border-bottom:2px solid var(--bd);position:sticky;top:var(--hdr-h);z-index:290;}}
.mtab{{padding:10px 20px;border:none;background:transparent;color:#b0bcd4;font-family:var(--f);font-size:12px;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-2px;letter-spacing:.5px;}}
.mtab:hover{{color:var(--t);}}.mtab.active{{color:var(--accent);border-bottom-color:var(--accent);font-weight:bold;}}
.main-sec{{display:none;}}.main-sec.active{{display:block;}}

/* ── CATEGORY TABS — оранжевий акцент, яскравіші неактивні ── */
.ctabs{{display:flex;gap:5px;padding:8px 24px;background:var(--bg2);border-bottom:1px solid var(--bd);flex-wrap:wrap;}}
.ctab{{padding:5px 14px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;
  color:#c0cce0;font-family:var(--f);font-size:12px;background:transparent;transition:color .15s,border-color .15s;}}
.ctab:hover{{border-color:var(--accent);color:#fff;}}
.ctab.active{{background:var(--accent);color:#000;border-color:var(--accent);font-weight:bold;}}
.tc{{opacity:.5;font-size:9px;margin-left:3px;}}

/* ── INSTRUMENT TABS — оранжевий акцент, яскравіші неактивні ── */
.catsec{{display:none;}}.catsec.active{{display:block;}}
.itabs{{display:flex;gap:4px;padding:7px 24px;background:var(--bg);border-bottom:1px solid var(--bd);flex-wrap:wrap;position:sticky;top:calc(var(--hdr-h) + 42px);z-index:200;}}
.itab{{padding:4px 12px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;
  color:#b0bcd4;font-family:var(--f);font-size:11px;background:transparent;transition:all .15s;}}
.itab:hover{{border-color:var(--accent);color:#fff;}}
.itab.active{{background:var(--bg3);color:var(--accent);border-color:var(--accent);font-weight:bold;}}

.iviews{{padding:16px 24px;}}.iview{{display:none;}}.iview.active{{display:block;}}

/* ── REPORT TABS ── */
.report-tabs{{display:flex;align-items:center;gap:6px;margin-bottom:14px;flex-wrap:wrap;}}
.rtab-lbl{{font-size:9px;color:var(--d);letter-spacing:1px;}}
.rtab{{padding:4px 12px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:11px;background:transparent;}}
.rtab.active{{background:var(--accent);color:#000;border-color:var(--accent);font-weight:bold;}}
.rtab.disabled{{opacity:.3;cursor:not-allowed;}}

/* ── METRIC CARDS — лейбл зліва, gauge справа ── */
.mcards{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px;}}
.mc{{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:12px 14px;overflow:hidden;min-width:0;display:flex;flex-direction:column;}}
/* Лейбл повернуто вліво вгорі */
.mc-lbl{{font-size:9px;color:#fff;letter-spacing:.6px;margin-bottom:5px;text-align:left;opacity:.85;}}
/* Рядок: число (ліво) + gauge (право) */
.mc-inner{{display:flex;align-items:center;gap:8px;}}
.mc-left{{flex:1;min-width:0;}}
.mc-right{{flex-shrink:0;display:flex;align-items:center;justify-content:center;}}
.mc-val{{font-size:clamp(18px,2.5vw,34px);font-weight:bold;line-height:1.1;}}
.mc-chg-wrap{{margin-top:6px;font-size:12px;}}
.mc-wtag{{font-size:9px;color:var(--d);margin-left:3px;}}
.mc-pct{{font-size:10px;margin-top:2px;opacity:.85;}}
.mc-sub{{font-size:10px;color:var(--d);margin-top:3px;}}

/* ── ANALYSIS PANEL — без "ЗМІНА ЗА ТИЖДЕНЬ", більші підписи ── */
.mid{{display:grid;grid-template-columns:1fr 160px 190px 1fr;gap:8px;margin-bottom:12px;}}
.panel{{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:12px 14px;}}
.plbl{{font-size:9px;color:#fff;letter-spacing:.5px;margin-bottom:10px;}}
.arow{{margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid var(--bd);}}
.arow:last-child{{margin:0;padding:0;border:none;}}
.arow-body{{display:flex;gap:10px;align-items:center;}}
.arow-left{{flex:1;min-width:0;display:flex;flex-direction:column;gap:6px;}}
.arow-right{{width:140px;flex-shrink:0;display:flex;flex-direction:column;justify-content:center;align-items:flex-end;border-left:1px solid var(--bd);padding-left:10px;}}
/* "ЗМІНА ЗА ТИЖДЕНЬ" прибрано з Python коду */
.ar-glbl{{font-size:11px;font-weight:bold;letter-spacing:.8px;margin-bottom:3px;}}
.ar-net{{font-size:clamp(20px,2.2vw,30px);font-weight:bold;text-align:right;}}
/* 2 колонки CHG L / CHG S */
.arow-grid2{{display:grid;grid-template-columns:1fr 1fr;gap:3px;}}
/* Δ NET — окремий рядок */
.arow-dnet{{display:flex;flex-direction:column;align-items:center;padding-top:6px;border-top:1px solid var(--bd);}}
.arow-dnet .ag-lbl{{font-size:10px;color:var(--d);letter-spacing:.4px;margin-bottom:2px;}}
.ag-val-net{{font-size:clamp(19px,2.2vw,28px);font-weight:bold;line-height:1.2;}}
.ag-item{{display:flex;flex-direction:column;gap:3px;}}
.ag-lbl{{font-size:10px;color:#c0ccd8;letter-spacing:.4px;font-weight:bold;}}
.ag-val{{font-size:clamp(17px,2vw,26px);font-weight:bold;line-height:1.2;}}
.ag-pct{{font-size:10px;opacity:.75;font-weight:normal;}}

/* ── SM PANEL ── */
.sm-panel{{display:flex;flex-direction:column;justify-content:space-between;}}
.sm-row{{margin-bottom:8px;}}
.sm-lbl{{font-size:9px;color:var(--d);margin-bottom:3px;}}
.sm-bar-bg{{background:var(--bg3);border-radius:10px;height:8px;position:relative;overflow:hidden;}}
.sm-mk{{position:absolute;top:1px;width:8px;height:6px;border-radius:3px;transform:translateX(-50%);}}
.sm-val{{font-size:12px;font-weight:bold;margin-top:3px;}}
.sm-hint{{font-size:8px;color:var(--d);margin-top:6px;line-height:1.5;border-top:1px solid var(--bd);padding-top:6px;}}

/* ── REPORTS PANEL ── */
.rpt-panel{{display:flex;flex-direction:column;overflow:hidden;}}
.plbl-row{{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;}}
.rel-legend{{font-size:8px;color:var(--d);display:flex;align-items:center;gap:3px;}}
.rel-icon{{font-size:10px;margin-right:1px;line-height:1;}}
.rel-d{{color:#20d483;}}.rel-i{{color:#f0a030;}}.rel-n{{color:#343d5a;}}
.rpt-row{{display:flex;align-items:center;padding:4px 0;border-bottom:1px solid var(--bd);gap:4px;}}
.rpt-row:last-child{{border:none;padding-bottom:0;}}
.rpt-dim{{opacity:.35;}}
.rpt-rel{{width:12px;flex-shrink:0;text-align:center;}}
.rpt-info{{flex:1;min-width:0;}}
.rpt-name{{font-size:9px;color:var(--t);font-weight:bold;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.rpt-sched{{font-size:8px;color:var(--d);margin-top:1px;}}
.rpt-tag{{display:inline-block;font-size:7px;padding:0 4px;border-radius:2px;background:rgba(74,158,255,.2);color:var(--b);margin-left:4px;vertical-align:middle;}}
.rpt-btns{{display:flex;gap:2px;flex-shrink:0;}}
.rb{{width:20px;height:20px;border-radius:3px;border:1px solid var(--bd);background:var(--bg3);color:var(--d);font-family:var(--f);font-size:9px;font-weight:bold;cursor:pointer;transition:all .1s;line-height:20px;text-align:center;padding:0;}}
.rb:hover{{border-color:var(--t);color:var(--t);}}
.rb-l.active{{background:rgba(32,212,131,.25);border-color:var(--g);color:var(--g);}}
.rb-s.active{{background:rgba(240,81,90,.20);border-color:var(--r);color:var(--r);}}
.rb-n.active{{background:var(--bg3);border-color:var(--d);color:var(--d);}}

/* ── PCT PANEL ── */
.pct-sel-row{{display:flex;gap:5px;align-items:center;margin-bottom:10px;flex-wrap:wrap;}}
.psel-group{{display:flex;gap:3px;}}
.psel-sep{{width:1px;height:16px;background:var(--bd);margin:0 3px;}}
.psel,.pper{{padding:2px 8px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:10px;background:transparent;}}
.psel:hover,.pper:hover{{border-color:var(--accent);color:#fff;}}
.psel.active,.pper.active{{background:var(--bg3);color:var(--accent);border-color:var(--accent);}}
.pct-val-row{{margin-bottom:8px;}}
.pbar-wrap{{position:relative;margin-bottom:3px;}}
.pbar-bg{{background:var(--bg3);border-radius:3px;height:18px;position:relative;overflow:hidden;}}
.pbar-lo{{position:absolute;left:0;top:0;height:100%;background:rgba(240,81,90,.3);width:15%;}}
.pbar-hi{{position:absolute;right:0;top:0;height:100%;background:rgba(32,212,131,.3);width:15%;}}
.ptick{{position:absolute;top:0;width:2px;height:100%;background:rgba(255,255,255,.25);}}
.pbar-mk{{position:absolute;top:2px;width:4px;height:14px;background:var(--g);border-radius:2px;transform:translateX(-50%);transition:left .3s;}}
.ptick-labels{{position:relative;height:16px;margin-top:2px;}}
.ptlbl{{position:absolute;transform:translateX(-50%);font-size:8px;color:var(--d);transition:left .3s;}}
.ptlbl-cur{{color:var(--t);font-weight:bold;}}
.pbar-lb{{display:flex;justify-content:space-between;font-size:8px;color:var(--d);margin-top:12px;}}

/* ── CHART BOXES ── */
.chartbox{{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:14px 16px;margin-bottom:12px;}}
.chartbox-hdr{{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:6px;}}
.period-btns{{display:flex;gap:3px;}}
.per-btn{{padding:2px 9px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:10px;background:transparent;}}
.per-btn:hover{{border-color:var(--accent);color:#fff;}}
.per-btn.active{{background:var(--bg3);color:var(--accent);border-color:var(--accent);}}
.chart-leg{{display:flex;gap:10px;font-size:10px;color:var(--d);align-items:center;flex-wrap:wrap;}}
.ll{{display:inline-block;width:14px;height:2px;border-radius:1px;vertical-align:middle;margin-right:4px;}}
.ll-dash{{display:inline-block;width:14px;height:0;border-top:2px dashed;vertical-align:middle;}}
.cw{{height:140px;position:relative;}}
.bar-charts-grid{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;}}
.bar-lbl{{font-size:8px;letter-spacing:.5px;margin-bottom:6px;}}
.bar-cw{{height:80px;position:relative;}}

/* ── ЄДИНА ТАБЛИЦЯ (MM + Data) ── */
.htable-wrap{{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;overflow:hidden;margin-bottom:12px;}}
.htable-hdr{{padding:7px 14px;border-bottom:1px solid var(--bd);font-size:10px;color:#fff;letter-spacing:.5px;display:flex;align-items:center;justify-content:space-between;}}
.hsel{{display:flex;gap:4px;}}
.hbtn{{padding:2px 10px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:11px;background:transparent;}}
.hbtn.active{{background:var(--bg3);color:var(--accent);border-color:var(--accent);}}
.htable-scroll{{overflow-x:auto;}}

/* Спільна таблиця */
table.ht{{width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap;table-layout:fixed;}}
table.ht th{{padding:4px 8px;border-bottom:1px solid var(--bd);font-weight:normal;font-size:9px;letter-spacing:.5px;text-align:right;overflow:hidden;}}
table.ht .th-corner{{text-align:left;background:var(--bg3);}}
table.ht .th-date{{text-align:left;background:var(--bg3);}}
table.ht .th-left{{text-align:left;}}
table.ht .th-group{{text-align:center;}}
table.ht .th-oi{{text-align:center;}}
table.ht .sm-th{{text-align:center;font-size:8px;color:var(--d);}}
table.ht .sm-th-group{{font-size:8px;}}
table.ht td{{padding:4px 8px;border-bottom:1px solid var(--bg3);text-align:right;overflow:hidden;}}
table.ht .date-col{{text-align:left;color:var(--d);background:var(--bg3);}}
table.ht tr:hover td{{background:rgba(52,61,90,.5)!important;}}
table.ht .sep-r{{border-right:1px solid var(--bd);}}
table.ht .sm-td{{text-align:center;font-size:10px;padding:4px 6px;}}

/* MM tbody — верхня секція Max/Min */
table.ht tbody.mm-tbody{{border-bottom:3px solid var(--bd);}}
table.ht tbody.mm-tbody td{{background:var(--bg3);text-align:right;border-bottom:1px solid rgba(52,61,90,.8);padding:4px 8px;}}
table.ht tbody.mm-tbody .mm-lbl{{text-align:left;font-size:8px;color:var(--d);letter-spacing:.5px;font-weight:bold;}}
table.ht tbody.mm-tbody .mm-val{{text-align:right;font-size:10px;}}
table.ht tbody.mm-tbody tr.mm-yr td{{opacity:.78;}}

/* OVERVIEW TABLE */
.ov-meta{{padding:10px 0 8px;font-size:11px;color:var(--d);}}.ov-meta b{{color:var(--t);}}
.ov-scroll{{overflow-x:auto;}}
.ov-table{{width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap;}}
.ov-table th{{padding:6px 10px;background:var(--bg3);border-bottom:1px solid var(--bd);color:var(--d);font-weight:normal;font-size:9px;letter-spacing:.5px;text-align:right;}}
.ov-table th:first-child{{text-align:left;}}
.ov-table td{{padding:5px 10px;border-bottom:1px solid var(--bg3);text-align:right;}}
.ov-table .ov-asset{{text-align:left;color:var(--t);font-weight:bold;}}
.ov-table .ov-group td{{background:var(--bg3);color:var(--d);font-size:8px;letter-spacing:1px;padding:4px 10px;text-align:left;}}
.ov-table tr:hover td{{background:rgba(52,61,90,.4)!important;}}
.ov-cot-cell{{display:flex;align-items:center;gap:6px;justify-content:flex-end;}}
.ov-bar-bg{{width:50px;height:5px;background:var(--bg3);border-radius:2px;overflow:hidden;flex-shrink:0;}}
.ov-bar-fill{{height:100%;border-radius:2px;}}
.ov-cot-val{{font-size:10px;color:var(--t);min-width:30px;text-align:right;}}

.g{{color:var(--g);}}.r{{color:var(--r);}}.d{{color:var(--d);}}
.footer{{text-align:center;padding:14px;color:var(--d);font-size:9px;letter-spacing:1px;border-top:1px solid var(--bd);margin-top:4px;}}

@media(max-width:640px){{
  :root{{--hdr-h:56px;}}
  .hdr{{padding:8px 12px;height:auto;min-height:var(--hdr-h);flex-wrap:wrap;}}
  .hdr-t{{font-size:14px;}}
  .ctabs{{padding:6px 12px;}}.ctab{{padding:4px 10px;font-size:11px;}}
  .itabs{{padding:5px 12px;}}.itab{{padding:3px 9px;font-size:10px;}}
  .iviews{{padding:10px 12px;}}
  .mcards{{grid-template-columns:1fr 1fr;gap:8px;}}
  .mc{{padding:10px 12px;}}.mc-val{{font-size:clamp(16px,5vw,26px);}}
  .mid{{grid-template-columns:1fr 1fr;gap:8px;}}
  .mid > .panel:first-child{{grid-column:1/-1;}}
  .bar-charts-grid{{grid-template-columns:1fr;}}
  .arow-right{{width:110px;padding-left:8px;}}
  .ar-net{{font-size:clamp(14px,4vw,20px);}}
  .ag-val{{font-size:clamp(11px,3.5vw,15px);}}
  table.ht{{font-size:10px;}}table.ht th{{padding:3px 5px;font-size:8px;}}table.ht td{{padding:3px 5px;}}
  .auth-box{{width:90vw;padding:24px 20px;}}
}}
</style>
</head>
<body>
<script>const _cd={{}};const _ci={{}};const Charts={{}};const BarChts={{}};</script>
"""

HTML_FOOT = f"""
<script>
const CL_LS='{COLOR_LS}', CL_CM='{COLOR_CM}', CL_ST='{COLOR_ST}';
const CurPer={{}};
let _loggedIn=false, _userEmail='';

function selCat(cat){{
  document.querySelectorAll('.ctab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.catsec').forEach(s=>s.classList.remove('active'));
  document.querySelector('[data-c="'+cat+'"]').classList.add('active');
  document.getElementById('cs_'+cat).classList.add('active');
  const first=document.querySelector('[data-cat="'+cat+'"]');
  if(first) selInst(cat,first.dataset.i);
}}
function selInst(cat,key){{
  document.querySelectorAll('[data-cat="'+cat+'"]').forEach(b=>b.classList.remove('active'));
  const btn=document.querySelector('[data-cat="'+cat+'"][data-i="'+key+'"]');
  if(btn) btn.classList.add('active');
  document.getElementById('iv_'+cat).querySelectorAll('.iview').forEach(v=>v.classList.remove('active'));
  const sid=key.replaceAll(' ','_').replaceAll('&','n').replaceAll('/','_');
  const view=document.getElementById('iv_'+sid);
  if(view){{
    view.classList.add('active');
    filterRows(sid,10);
    const n=CurPer[sid]||52;
    setTimeout(()=>{{drawMainChart(sid,n);drawBarsFor(sid,n);}},30);
    loadRptStances(sid);
  }}
}}
function selMain(mt){{
  document.querySelectorAll('.mtab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.main-sec').forEach(s=>s.classList.remove('active'));
  document.querySelector('[data-mt="'+mt+'"]').classList.add('active');
  document.getElementById('ms_'+mt).classList.add('active');
  if(mt==='cot'){{const fc=document.querySelector('.ctab');if(fc)selCat(fc.dataset.c);}}
}}
function setChartPer(btn,sid){{
  btn.closest('.period-btns').querySelectorAll('.per-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const per=btn.dataset.per; const n=per==='1y'?52:per==='3y'?156:9999;
  CurPer[sid]=n; drawMainChart(sid,n); drawBarsFor(sid,n);
}}
function drawMainChart(sid,nWeeks){{
  const cv=document.getElementById('cv_'+sid); if(!cv) return;
  if(Charts[sid]){{Charts[sid].destroy();delete Charts[sid];}}
  const d=_cd[sid]; if(!d) return;
  const n=Math.min(nWeeks,d.dates.length);
  Charts[sid]=new Chart(cv.getContext('2d'),{{
    type:'line',
    data:{{labels:d.dates.slice(-n),datasets:[
      {{label:'Large Spec',   data:d.ls.slice(-n),borderColor:CL_LS,backgroundColor:CL_LS+'22',borderWidth:1.5,pointRadius:0,tension:.3,fill:true}},
      {{label:'Commercials',  data:d.cm.slice(-n),borderColor:CL_CM,backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3}},
      {{label:'Small Traders',data:d.st.slice(-n),borderColor:CL_ST,backgroundColor:'transparent',borderWidth:1,  pointRadius:0,tension:.3,borderDash:[3,3]}},
    ]}},
    options:{{responsive:true,maintainAspectRatio:false,animation:false,
      interaction:{{mode:'index',intersect:false}},
      plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,
        titleColor:'#dde2ee',bodyColor:'#dde2ee',
        titleFont:{{family:'Courier New',size:10}},bodyFont:{{family:'Courier New',size:10}},
        callbacks:{{label:ctx=>' '+ctx.dataset.label+': '+fmtFull(ctx.parsed.y)}}}}}},
      scales:{{
        x:{{display:true,ticks:{{color:'#8090b0',font:{{family:'Courier New',size:8}},maxTicksLimit:8,
          callback:function(v,i){{return i%Math.ceil(n/8)===0?this.getLabelForValue(v):'';}}}},
          grid:{{display:false}},border:{{display:false}}}},
        y:{{display:true,grid:{{color:'rgba(52,61,90,.8)',lineWidth:.5}},
          ticks:{{color:'#8090b0',font:{{family:'Courier New',size:9}},maxTicksLimit:4,callback:v=>fmtV(v,true)}},
          border:{{display:false}}}}
      }}
    }}
  }});
}}
function drawBarsFor(sid,nWeeks){{
  const d=_cd[sid]; if(!d) return;
  const n=Math.min(nWeeks,d.dates.length); const dates=d.dates.slice(-n);
  drawOneBar('barcv_ls_'+sid,dates,d.ld.slice(-n),CL_LS,'rgba(240,81,90,.75)');
  drawOneBar('barcv_cm_'+sid,dates,d.cd.slice(-n),CL_CM,'rgba(240,81,90,.75)');
  drawOneBar('barcv_st_'+sid,dates,d.sd.slice(-n),CL_ST,'rgba(15,18,28,.9)');
}}
function drawOneBar(cvId,dates,data,baseColor,negColor){{
  const cv=document.getElementById(cvId); if(!cv) return;
  const key='b_'+cvId;
  if(BarChts[key]){{BarChts[key].destroy();delete BarChts[key];}}
  const colors=data.map(v=>v>=0?baseColor+'cc':negColor);
  BarChts[key]=new Chart(cv.getContext('2d'),{{
    type:'bar',
    data:{{labels:dates,datasets:[{{data:data,backgroundColor:colors,borderWidth:0,borderRadius:1}}]}},
    options:{{responsive:true,maintainAspectRatio:false,animation:false,
      plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,
        titleColor:'#dde2ee',bodyColor:'#dde2ee',
        titleFont:{{family:'Courier New',size:9}},bodyFont:{{family:'Courier New',size:9}},
        callbacks:{{label:ctx=>fmtFull(ctx.parsed.y)}}}}}},
      scales:{{x:{{display:false}},y:{{display:true,
        grid:{{color:'rgba(52,61,90,.6)',lineWidth:.5}},
        ticks:{{color:'#8090b0',font:{{family:'Courier New',size:8}},maxTicksLimit:3,callback:v=>fmtV(v,true)}},
        border:{{display:false}}}}}}
    }}
  }});
}}
function pctSel(btn,sid){{
  btn.closest('.psel-group').querySelectorAll('.psel').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); updatePctBar(sid);
}}
function pperSel(btn,sid){{
  btn.closest('.psel-group').querySelectorAll('.pper').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); updatePctBar(sid);
}}
function updatePctBar(sid){{
  const view=document.getElementById('iv_'+sid); if(!view) return;
  const p=view.querySelector('.psel.active')?.dataset.p||'ls';
  const per=view.querySelector('.pper.active')?.dataset.per||'all';
  const val=_ci[sid]?.[p]?.[per]??50;
  const pos=Math.min(Math.max(val,0),100);
  const col=val<15?'#f0515a':val>85?'#20d483':'#dde2ee';
  const lbl=val<15?'— екстрем. шорт':val>85?'— екстрем. лонг':'— нейтральна зона';
  const mk=document.getElementById('pctmk_'+sid);
  const valEl=document.getElementById('pctval_'+sid);
  const lblEl=document.getElementById('pctlbl_'+sid);
  const curEl=document.getElementById('pctcur_'+sid);
  if(mk)mk.style.left=pos+'%';
  if(valEl){{valEl.style.color=col;valEl.textContent=val.toFixed(1)+'%';}}
  if(lblEl)lblEl.textContent=lbl;
  if(curEl){{curEl.style.left=pos+'%';curEl.textContent=val.toFixed(1)+'%';}}
}}
function setHist(btn,sid){{
  const n=parseInt(btn.dataset.n);
  btn.closest('.htable-hdr').querySelectorAll('.hbtn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); filterRows(sid,n);
}}
/* Фільтруємо лише рядки data-tbody (не mm-tbody) */
function filterRows(sid,n){{
  const view=document.getElementById('iv_'+sid); if(!view) return;
  view.querySelectorAll('table.ht tbody.data-tbody tr').forEach(tr=>{{
    tr.style.display=parseInt(tr.dataset.row)<n?'':'none';
  }});
}}
function setRpt(sid,rptId,stance){{
  const row=document.getElementById('rpt_'+sid+'_'+rptId); if(!row) return;
  row.querySelectorAll('.rb').forEach(b=>b.classList.remove('active'));
  const cls=stance==='long'?'.rb-l':stance==='short'?'.rb-s':'.rb-n';
  row.querySelector(cls)?.classList.add('active');
  localStorage.setItem('rpt_'+sid+'_'+rptId,stance);
  if(_loggedIn){{
    fetch('/api/rpt_stance',{{method:'POST',headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{instrument:sid,report:rptId,stance:stance}})}}).catch(()=>{{}});
  }}
}}
function loadRptStances(sid){{
  const rptIds={json.dumps([r['id'] for r in REPORTS])};
  rptIds.forEach(rptId=>{{
    const saved=localStorage.getItem('rpt_'+sid+'_'+rptId)||'neutral';
    applyRptStance(sid,rptId,saved);
  }});
  if(_loggedIn){{
    fetch('/api/rpt_stances?instrument='+sid).then(r=>r.json()).then(data=>{{
      if(data&&data.stances){{
        Object.entries(data.stances).forEach(([rptId,stance])=>{{
          applyRptStance(sid,rptId,stance);
          localStorage.setItem('rpt_'+sid+'_'+rptId,stance);
        }});
      }}
    }}).catch(()=>{{}});
  }}
}}
function applyRptStance(sid,rptId,stance){{
  const row=document.getElementById('rpt_'+sid+'_'+rptId); if(!row) return;
  row.querySelectorAll('.rb').forEach(b=>b.classList.remove('active'));
  const cls=stance==='long'?'.rb-l':stance==='short'?'.rb-s':'.rb-n';
  row.querySelector(cls)?.classList.add('active');
}}
function fmtV(n,short=false){{
  if(n===null||isNaN(n))return'—'; n=Math.round(n);
  if(short){{
    if(Math.abs(n)>=1e6)return(n/1e6).toFixed(1)+'M';
    if(Math.abs(n)>=1e3)return(n/1e3).toFixed(0)+'K';
    return''+n;
  }}
  const sign=n>0?'+':n<0?'-':'';
  return sign+Math.abs(n).toLocaleString('uk-UA');
}}
function fmtFull(n){{
  if(n===null||isNaN(n))return'—'; n=Math.round(n);
  const sign=n>0?'+':n<0?'-':'';
  return sign+Math.abs(n).toLocaleString('uk-UA');
}}
function openAuth(){{document.getElementById('authOverlay').classList.add('open');}}
function closeAuth(){{document.getElementById('authOverlay').classList.remove('open');}}
function authTab(t){{
  document.querySelectorAll('.auth-tab').forEach((el,i)=>
    el.classList.toggle('active',(i===0&&t==='login')||(i===1&&t==='reg')));
  document.getElementById('at-login').style.display=t==='login'?'':'none';
  document.getElementById('at-reg').style.display=t==='reg'?'':'none';
  ['al-msg','ar-msg'].forEach(id=>{{const el=document.getElementById(id);if(el)el.style.display='none';}});
}}
function showAuthMsg(id,text,isErr){{
  const el=document.getElementById(id);
  el.textContent=text;el.className='auth-msg '+(isErr?'err':'ok');el.style.display='block';
}}
async function doAuthLogin(){{
  const email=document.getElementById('al-email').value.trim();
  const pass=document.getElementById('al-pass').value;
  if(!email||!pass){{showAuthMsg('al-msg','Заповніть всі поля',true);return;}}
  try{{
    const res=await fetch('/api/login',{{method:'POST',headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{email,password:pass}})}});
    const data=await res.json();
    if(data.ok){{setLoggedIn(data.email);closeAuth();}}
    else showAuthMsg('al-msg',data.error||'Помилка входу',true);
  }}catch(e){{showAuthMsg('al-msg','Сервер недоступний',true);}}
}}
async function doAuthReg(){{
  const email=document.getElementById('ar-email').value.trim();
  const pass=document.getElementById('ar-pass').value;
  const pass2=document.getElementById('ar-pass2').value;
  if(!email||!pass){{showAuthMsg('ar-msg','Заповніть всі поля',true);return;}}
  if(pass!==pass2){{showAuthMsg('ar-msg','Паролі не збігаються',true);return;}}
  try{{
    const res=await fetch('/api/register',{{method:'POST',headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{email,password:pass}})}});
    const data=await res.json();
    if(data.ok){{showAuthMsg('ar-msg','Успішно! Тепер увійдіть.',false);setTimeout(()=>authTab('login'),1500);}}
    else showAuthMsg('ar-msg',data.error||'Помилка реєстрації',true);
  }}catch(e){{showAuthMsg('ar-msg','Сервер недоступний',true);}}
}}
async function doLogout(){{await fetch('/api/logout',{{method:'POST'}});setLoggedOut();closeAuth();}}
function setLoggedIn(email){{
  _loggedIn=true;_userEmail=email;
  const btn=document.getElementById('authBtn');
  if(btn){{btn.textContent='● '+email.split('@')[0].toUpperCase().substring(0,8);btn.classList.add('logged');}}
  const logged=document.getElementById('auth-logged');
  const loggedout=document.getElementById('auth-loggedout');
  const emailDisplay=document.getElementById('auth-email-display');
  if(emailDisplay)emailDisplay.textContent=email;
  if(logged)logged.style.display='';if(loggedout)loggedout.style.display='none';
  const activeView=document.querySelector('.iview.active');
  if(activeView)loadRptStances(activeView.dataset.sid);
}}
function setLoggedOut(){{
  _loggedIn=false;_userEmail='';
  const btn=document.getElementById('authBtn');
  if(btn){{btn.textContent='УВІЙТИ';btn.classList.remove('logged');}}
  const logged=document.getElementById('auth-logged');
  const loggedout=document.getElementById('auth-loggedout');
  if(logged)logged.style.display='none';if(loggedout)loggedout.style.display='';
}}
fetch('/api/me').then(r=>r.json()).then(d=>{{if(d.logged_in)setLoggedIn(d.email);}}).catch(()=>{{}});
const firstCat=document.querySelector('.ctab');
if(firstCat)selCat(firstCat.dataset.c);
</script>
</body>
</html>
"""


def main():
    print()
    print("="*55)
    print("   COT Dashboard Generator v13")
    print("="*55)
    print()
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    try:
        data = load_all()
    except FileNotFoundError as e:
        print(e); return
    if not data:
        print("❌  Дані порожні."); return
    print("🔧  Генеруємо HTML...")
    html = generate_html(data)
    OUTPUT_FILE.write_text(html, encoding='utf-8')
    kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"✅  Збережено: {OUTPUT_FILE}  ({kb:.0f} KB)")
    print("🌐  Відкриваємо браузер...")
    import os
    uri = OUTPUT_FILE.as_uri()
    opened = webbrowser.open(uri)
    if not opened:
        try:
            os.startfile(str(OUTPUT_FILE))
        except Exception:
            pass
    print(f"   Файл: {OUTPUT_FILE}")
    print("\n✨  Готово!\n")

if __name__ == '__main__':
    main()