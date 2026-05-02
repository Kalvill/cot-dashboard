#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""COT Dashboard Generator v7"""

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
    'DOW_30':'DOW_30','DOW 30':'DOW_30',
    'RUSSELL2K':'RUSSELL2K','VIX':'VIX',
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
    'date':1,
    'ls_cl':4,'ls_cs':5,'ls_pct':6,'ls_net':8,
    'cm_cl':11,'cm_cs':12,'cm_pct':13,'cm_net':15,
    'st_cl':18,'st_cs':19,'st_pct':20,'st_net':22,
}
DATA_START_ROW = 5


# ================================================================
# 🔧  УТИЛІТИ
# ================================================================
def disp(n): return DISPLAY.get(n, n)
def sid(n):  return n.replace(' ','_').replace('&','n').replace('/','_')

def to_num(s):
    return pd.to_numeric(s, errors='coerce').fillna(0).round(2).tolist()

def norm_pct(vals):
    if not vals: return vals
    nz = [abs(v) for v in vals if v != 0]
    if nz and max(nz) <= 1.5:
        return [round(v*100,1) for v in vals]
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
    """Повне число з пробілом-роздільником тисяч: +78 560"""
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


# ================================================================
# 🎨  SVG SPARKLINE
# ================================================================
def make_sparkline(series, h=38):
    data=[float(v) for v in (series or [])[-SPARK_WEEKS:]]
    n=len(data)
    if n<3: return ''
    W=200; H=h
    mn=min(data); mx=max(data); rng=mx-mn
    if rng==0: return ''
    def px(i): return round(i/(n-1)*W,1)
    def py(v): return round((1-(v-mn)/rng)*H,1)
    pts=[(px(i),py(v)) for i,v in enumerate(data)]
    line=' '.join(f"{x},{y}"for x,y in pts)
    zy=max(0,min(H,py(0)))
    area=f"0,{zy} "+line+f" {W},{zy}"
    lx,ly=pts[-1]
    color='#20d483'if data[-1]>=0 else'#f0515a'
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


# ================================================================
# 📥  OVERVIEW
# ================================================================
def read_overview(xl):
    result={}
    try:
        raw=xl.parse('overview',header=None)
        for i in range(4,len(raw)):
            row=raw.iloc[i]
            asset=str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            if not asset or asset=='nan': continue
            cot_ls=pd.to_numeric(row.iloc[4],errors='coerce')
            if pd.isna(cot_ls): continue
            def safe(c):
                v=pd.to_numeric(row.iloc[c],errors='coerce')
                return float(v) if pd.notna(v) else 0.0
            s=OVERVIEW_TO_SID.get(asset,asset)
            result[s]={
                'cot_ls_all':round(safe(4)*100,1),
                'cot_cm_all':round(safe(5)*100,1),
                'cot_st_all':round(safe(6)*100,1),
                'sm_div':    round(safe(8),3),
                'sm_div_3m': round(safe(18),3),
                'sm_div_6m': round(safe(19),3),
            }
    except Exception as e:
        print(f"  ⚠  overview: {e}")
    return result


# ================================================================
# 📥  ОСНОВНИЙ АРКУШ
# ================================================================
def read_sheet(xl,name,overview):
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

        N=min(CHART_WEEKS,len(df))
        chart_df=df.tail(N).reset_index(drop=True)
        def gcc(idx):
            return to_num(chart_df.iloc[:,idx]) if idx<chart_df.shape[1] else [0.0]*N
        ls_w=gcc(COL['ls_net']); cm_w=gcc(COL['cm_net']); st_w=gcc(COL['st_net'])
        oi_w=gcc(oi_col)
        chart={
            'dates':chart_df['_dt'].dt.strftime('%d.%m.%Y').tolist(),
            'ls':ls_w,'cm':cm_w,'st':st_w,'oi':oi_w,
            'ld':compute_delta(ls_w),
            'cd':compute_delta(cm_w),
            'sd':compute_delta(st_w),
        }

        hist_df=df.tail(HISTORY).reset_index(drop=True)
        def gch(idx):
            return to_num(hist_df.iloc[:,idx]) if idx<hist_df.shape[1] else [0.0]*len(hist_df)
        hist={
            'dates': hist_df['_dt'].dt.strftime('%d.%m.%Y').tolist(),
            'ls_cl': gch(COL['ls_cl']), 'ls_cs': gch(COL['ls_cs']),
            'ls_net':gch(COL['ls_net']),
            'cm_cl': gch(COL['cm_cl']), 'cm_cs': gch(COL['cm_cs']),
            'cm_net':gch(COL['cm_net']),
            'st_cl': gch(COL['st_cl']), 'st_cs': gch(COL['st_cs']),
            'st_net':gch(COL['st_net']),
            'oi':    gch(oi_col),
        }

        oi_cur=float(oi_all[i0]); oi_prev=float(oi_all[i1])
        oi_pct=round((oi_cur-oi_prev)/abs(oi_prev)*100,2) if oi_prev!=0 else 0.0

        return{
            'name':name,'display':disp(name),'sid':sid(name),
            'chart':chart,'hist':hist,
            'cot_idx':{
                'ls':cot_idx(ls_net,'cot_ls_all'),
                'cm':cot_idx(cm_net,'cot_cm_all'),
                'st':cot_idx(st_net,'cot_st_all'),
            },
            'sm':{
                'div':   ov.get('sm_div',0.0),
                'div_3m':ov.get('sm_div_3m',0.0),
                'div_6m':ov.get('sm_div_6m',0.0),
            },
            'spark':{'ls':ls_net,'cm':cm_net,'st':st_net,'oi':oi_all},
            'cur':{
                'date':  all_dates[i0],
                'ls_net':ls_net[i0],'cm_net':cm_net[i0],'st_net':st_net[i0],
                'ls_pct':ls_pct[i0],'cm_pct':cm_pct[i0],'st_pct':st_pct[i0],
                'ls_cl': ls_cl[i0], 'ls_cs': ls_cs[i0],
                'cm_cl': cm_cl[i0], 'cm_cs': cm_cs[i0],
                'oi':    oi_cur,'oi_pct':oi_pct,
                'ls_chg':round(float(ls_net[i0])-float(ls_net[i1]),0),
                'cm_chg':round(float(cm_net[i0])-float(cm_net[i1]),0),
                'st_chg':round(float(st_net[i0])-float(st_net[i1]),0),
                'oi_chg':round(oi_cur-oi_prev,0),
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
# 🎨  ТАБЛИЦЯ (без % NET, повні числа, інтенсивність кольору)
# ================================================================
def intensity_bg(val, max_abs):
    if max_abs==0: return ''
    try:
        v=float(val)
        ratio=min(abs(v)/max_abs,1.0)
        op=0.10+ratio*0.67
        if v>0:   return f' style="background:rgba(32,212,131,{op:.2f})"'
        elif v<0: return f' style="background:rgba(240,81,90,{op:.2f})"'
    except: pass
    return ''

def make_hist_table(hist):
    n=len(hist['dates'])
    if n==0: return '<p style="padding:12px;color:#8090b0">Немає даних</p>'

    def maxabs(lst):
        vals=[abs(v) for v in lst if v!=0]
        return max(vals) if vals else 1

    m_ls_cl=maxabs(hist['ls_cl']); m_ls_cs=maxabs(hist['ls_cs'])
    m_cm_cl=maxabs(hist['cm_cl']); m_cm_cs=maxabs(hist['cm_cs'])
    m_st_cl=maxabs(hist['st_cl']); m_st_cs=maxabs(hist['st_cs'])

    rows=[]
    for i in range(n-1,-1,-1):
        row_idx=n-1-i
        def td_chg(v,maxv):
            cls=cc(v); txt=fv_full(v,sign=True) if v!=0 else '—'
            return f'<td class="{cls}"{intensity_bg(v,maxv)}>{txt}</td>'
        def td_net(v,extra=''):
            return f'<td class="{cc(v)}{extra}">{fv_full(v,sign=True)}</td>'
        row=(
            f'<tr data-row="{row_idx}">'
            f'<td class="date-col">{hist["dates"][i]}</td>'
            +td_chg(hist['ls_cl'][i],m_ls_cl)
            +td_chg(hist['ls_cs'][i],m_ls_cs)
            +td_net(hist['ls_net'][i],' sep-r')
            +td_chg(hist['cm_cl'][i],m_cm_cl)
            +td_chg(hist['cm_cs'][i],m_cm_cs)
            +td_net(hist['cm_net'][i],' sep-r')
            +td_chg(hist['st_cl'][i],m_st_cl)
            +td_chg(hist['st_cs'][i],m_st_cs)
            +td_net(hist['st_net'][i],' sep-r')
            +f'<td class="d">{fv_full(hist["oi"][i])}</td>'
            +'</tr>'
        )
        rows.append(row)

    return(
        '<table class="ht">'
        '<thead>'
        '<tr>'
        '<th rowspan="2" class="th-left th-date">ДАТА</th>'
        '<th colspan="3" class="th-group">LARGE SPECULATORS</th>'
        '<th colspan="3" class="th-group">COMMERCIALS</th>'
        '<th colspan="3" class="th-group">SMALL TRADERS</th>'
        '<th class="th-group">OI</th>'
        '</tr>'
        '<tr>'
        '<th>CHG L</th><th>CHG S</th><th class="sep-r">NET POS</th>'
        '<th>CHG L</th><th>CHG S</th><th class="sep-r">NET POS</th>'
        '<th>CHG L</th><th>CHG S</th><th class="sep-r">NET POS</th>'
        '<th></th>'
        '</tr>'
        '</thead>'
        '<tbody>'+''.join(rows)+'</tbody>'
        '</table>'
    )


# ================================================================
# 🃏  SM DIV бар
# ================================================================
def sm_bar(val,label):
    v=float(val) if val else 0.0
    pct=min(max((v+1)/2*100,0),100)
    color='#20d483'if v>0 else('#f0515a'if v<0 else'#6a7290')
    cls='g'if v>0 else('r'if v<0 else'd')
    return(
        f'<div class="sm-row">'
        f'<div class="sm-lbl">{label}</div>'
        f'<div class="sm-bar-bg">'
        f'<div class="sm-mk" style="left:{pct:.0f}%;background:{color}"></div>'
        f'</div>'
        f'<div class="sm-val {cls}">{v:+.2f}</div>'
        f'</div>'
    )


# ================================================================
# 🃏  METRIC CARD
# ================================================================
def make_metric_card(lbl,val,chg_str,chg_cls,sub,spark_series):
    spark=make_sparkline(spark_series)
    return(
        f'<div class="mc">'
        f'<div class="mc-lbl">{lbl}</div>'
        f'<div class="mc-val {cc(val)}">{fv(val,sign=True)}</div>'
        f'<div class="mc-chg {chg_cls}">{chg_str}</div>'
        f'<div class="mc-sub">{sub}</div>'
        f'{spark}'
        f'</div>'
    )


# ================================================================
# 🃏  АНАЛІЗ ПОЗИЦІОНУВАННЯ (повні числа, рівномірно)
# ================================================================
def analysis_row(lbl, net, cl, cs, chg):
    dw='ЛОНГ'if net>0 else'ШОРТ'
    dc='g'   if net>0 else'r'
    return(
        f'<div class="arow">'
        f'<div class="arow-lbl">{lbl}</div>'
        f'<div class="arow-dir {dc}">{dw}'
        f'<span class="arow-net">{fv_full(net,sign=True)}</span>'
        f'</div>'
        f'<div class="arow-grid">'
        f'<div class="ag-item">'
        f'<span class="ag-lbl">CHG LONG</span>'
        f'<span class="{cc(cl)} ag-val">{fv_full(cl,sign=True)}</span>'
        f'</div>'
        f'<div class="ag-item">'
        f'<span class="ag-lbl">CHG SHORT</span>'
        f'<span class="{cc(-cs)} ag-val">{fv_full(-cs,sign=True)}</span>'
        f'</div>'
        f'<div class="ag-item">'
        f'<span class="ag-lbl">Δ NET</span>'
        f'<span class="{cc(chg)} ag-val">{fv_full(chg,sign=True)}</span>'
        f'</div>'
        f'</div>'
        f'</div>'
    )


# ================================================================
# 🃏  КАРТКА ІНСТРУМЕНТУ
# ================================================================
def make_instrument_view(d):
    c=d['cur']; s=d['sid']; sm=d['sm']

    mc_ls=make_metric_card('LARGE SPEC (NETTO)',   c['ls_net'],
        f'{ar(c["ls_chg"])} {fv(abs(int(c["ls_chg"])),True)}',
        cc(c['ls_chg']),f'COT Index: {fp(d["cot_idx"]["ls"]["all"])}',d['spark']['ls'])
    mc_cm=make_metric_card('COMMERCIALS (NETTO)',  c['cm_net'],
        f'{ar(c["cm_chg"])} {fv(abs(int(c["cm_chg"])),True)}',
        cc(c['cm_chg']),f'COT Index: {fp(d["cot_idx"]["cm"]["all"])}',d['spark']['cm'])
    mc_st=make_metric_card('SMALL TRADERS (NETTO)',c['st_net'],
        f'{ar(c["st_chg"])} {fv(abs(int(c["st_chg"])),True)}',
        cc(c['st_chg']),f'COT Index: {fp(d["cot_idx"]["st"]["all"])}',d['spark']['st'])
    mc_oi=make_metric_card('OPEN INTEREST',        c['oi'],
        f'{ar(c["oi_pct"])} {fp(c["oi_pct"])}',
        cc(c['oi_pct']),f'зміна: {fv(int(c["oi_chg"]),True,sign=True)}',d['spark']['oi'])

    analysis_panel=(
        f'<div class="panel">'
        f'<div class="plbl">АНАЛІЗ ПОЗИЦІОНУВАННЯ</div>'
        +analysis_row('LARGE SPEC',  c['ls_net'],c['ls_cl'],c['ls_cs'],c['ls_chg'])
        +analysis_row('COMMERCIALS', c['cm_net'],c['cm_cl'],c['cm_cs'],c['cm_chg'])
        +f'</div>'
    )

    sm_panel=(
        f'<div class="panel sm-panel">'
        f'<div class="plbl">SM DIVERGENCE</div>'
        +sm_bar(sm['div'],   'All time')
        +sm_bar(sm['div_6m'],'6 months')
        +sm_bar(sm['div_3m'],'3 months')
        +f'<div class="sm-hint">+ бичача &nbsp;|&nbsp; − ведмежа розбіжність</div>'
        +f'</div>'
    )

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
        f'</div>'
        f'<div class="psel-sep"></div>'
        f'<div class="psel-group">'
        f'<button class="pper active" data-per="all" onclick="pperSel(this,\'{s}\')">All</button>'
        f'<button class="pper" data-per="3y" onclick="pperSel(this,\'{s}\')">3Y</button>'
        f'<button class="pper" data-per="1y" onclick="pperSel(this,\'{s}\')">1Y</button>'
        f'<button class="pper" data-per="6m" onclick="pperSel(this,\'{s}\')">6M</button>'
        f'<button class="pper" data-per="3m" onclick="pperSel(this,\'{s}\')">3M</button>'
        f'</div>'
        f'</div>'
        f'<div class="pct-val-row">'
        f'<span id="pctval_{s}" style="font-size:16px;font-weight:bold;color:{ini_color}">'
        f'{fp(ini)}</span>'
        f'<span id="pctlbl_{s}" style="font-size:11px;color:#8090b0;margin-left:8px;">'
        f'— {ini_label}</span>'
        f'</div>'
        f'<div class="pbar-wrap">'
        f'<div class="pbar-bg">'
        f'<div class="pbar-lo"></div>'
        f'<div class="pbar-hi"></div>'
        f'<div class="ptick" style="left:15%"></div>'
        f'<div class="ptick" style="left:85%"></div>'
        f'<div class="pbar-mk" id="pctmk_{s}" style="left:{ini_pos:.1f}%"></div>'
        f'</div>'
        f'<div class="ptick-labels">'
        f'<span class="ptlbl" style="left:15%">15%</span>'
        f'<span id="pctcur_{s}" class="ptlbl ptlbl-cur" style="left:{ini_pos:.1f}%">{fp(ini)}</span>'
        f'<span class="ptlbl" style="left:85%">85%</span>'
        f'</div>'
        f'</div>'
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
        f'<span><span class="ll" style="background:#20d483"></span>Large Spec</span>'
        f'<span><span class="ll" style="background:#f0515a"></span>Commercials</span>'
        f'<span class="ll-dash"></span><span style="margin-left:4px">Small Traders</span>'
        f'</div></div>'
        f'</div>'
        f'<div class="cw"><canvas id="cv_{s}"></canvas></div>'
        f'</div>'
        f'<script>_cd["{s}"]={chart_json};</script>'
    )

    bar_block=(
        f'<div class="bar-charts-grid chartbox">'
        f'<div class="bar-wrap">'
        f'<div class="bar-lbl">LARGE SPEC — WEEKLY ΔNet</div>'
        f'<div class="bar-cw"><canvas id="barcv_ls_{s}"></canvas></div>'
        f'</div>'
        f'<div class="bar-wrap">'
        f'<div class="bar-lbl">COMMERCIALS — WEEKLY ΔNet</div>'
        f'<div class="bar-cw"><canvas id="barcv_cm_{s}"></canvas></div>'
        f'</div>'
        f'<div class="bar-wrap">'
        f'<div class="bar-lbl">SMALL TRADERS — WEEKLY ΔNet</div>'
        f'<div class="bar-cw"><canvas id="barcv_st_{s}"></canvas></div>'
        f'</div>'
        f'</div>'
    )

    hist_html=make_hist_table(d['hist'])
    table_block=(
        f'<div class="htable-wrap">'
        f'<div class="htable-hdr">'
        f'<span>ТИЖНЕВА СТАТИСТИКА ПОЗИЦІЙ</span>'
        f'<div class="hsel">'
        f'<button class="hbtn active" data-n="10" onclick="setHist(this,\'{s}\')">10</button>'
        f'<button class="hbtn" data-n="26" onclick="setHist(this,\'{s}\')">26</button>'
        f'<button class="hbtn" data-n="52" onclick="setHist(this,\'{s}\')">52</button>'
        f'</div></div>'
        f'<div class="htable-scroll">{hist_html}</div>'
        f'</div>'
    )

    return(
        f'<div class="iview" id="iv_{s}" data-sid="{s}">'
        f'<div class="report-tabs">'
        f'<span class="rtab-lbl">ТИП ЗВІТУ:</span>'
        f'<button class="rtab active">Legacy Report</button>'
        f'<button class="rtab disabled" title="Файл не завантажено">TFF Report</button>'
        f'<button class="rtab disabled" title="Файл не завантажено">Disaggregated</button>'
        f'</div>'
        +f'<div class="mcards">{mc_ls}{mc_cm}{mc_st}{mc_oi}</div>'
        +f'<div class="mid">{analysis_panel}{sm_panel}{pct_panel}</div>'
        +chart_block+bar_block+table_block
        +f'</div>'
    )


# ================================================================
# 🌐  ЗБІРКА HTML
# ================================================================
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

    return(
        HTML_HEAD
        +f'<header class="hdr">'
        +f'<div><div class="hdr-t">COT <em>DASHBOARD</em></div>'
        +f'<div class="hdr-s">COMMITMENTS OF TRADERS — CFTC LEGACY REPORT</div></div>'
        +f'<div class="hdr-r">Звіт: <b>{report_date}</b><br>Оновлено: {updated}</div>'
        +f'</header>'
        # Категорії — НЕ sticky, прокручуються разом зі сторінкою
        +f'<div class="ctabs">'+''.join(cat_tabs)+f'</div>'
        +''.join(cat_sects)
        +f'<div class="footer">COT DASHBOARD &bull; CFTC LEGACY &bull; {updated}</div>'
        +HTML_FOOT
    )


# ================================================================
# 🖼️  CSS + HEAD
# ================================================================
HTML_HEAD="""<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>COT Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{
  --bg:#1a1e2d;--bg2:#21263a;--bg3:#282f47;--bd:#343d5a;
  --g:#20d483;--r:#f0515a;--b:#4a9eff;
  --t:#dde2ee;--d:#8090b0;
  --f:'Courier New',Courier,monospace;
  --hdr-h:50px;
}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box;}
html,body{background:var(--bg);color:var(--t);font-family:var(--f);font-size:13px;}

/* ── HEADER ── sticky завжди */
.hdr{
  height:var(--hdr-h);padding:0 24px;
  background:var(--bg2);border-bottom:1px solid var(--bd);
  display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:300;
}
.hdr-t{font-size:17px;font-weight:bold;color:#fff;letter-spacing:2px;}
.hdr-t em{color:var(--g);font-style:normal;}
.hdr-s{font-size:10px;color:var(--d);margin-top:2px;letter-spacing:1px;}
.hdr-r{text-align:right;font-size:11px;color:var(--d);line-height:2;}
.hdr-r b{color:var(--t);}

/* ── CATEGORY TABS ── НЕ sticky, прокручуються */
.ctabs{
  display:flex;gap:5px;padding:8px 24px;
  background:var(--bg2);border-bottom:1px solid var(--bd);flex-wrap:wrap;
}
.ctab{padding:5px 14px;border:1px solid var(--bd);border-radius:3px;
  cursor:pointer;color:var(--d);font-family:var(--f);font-size:12px;
  background:transparent;transition:all .15s;}
.ctab:hover{border-color:var(--g);color:var(--t);}
.ctab.active{background:var(--g);color:#000;border-color:var(--g);font-weight:bold;}
.tc{opacity:.4;font-size:9px;margin-left:3px;}

/* ── INSTRUMENT TABS ── sticky одразу під header */
.catsec{display:none;}.catsec.active{display:block;}
.itabs{
  display:flex;gap:4px;padding:7px 24px;
  background:var(--bg);border-bottom:1px solid var(--bd);flex-wrap:wrap;
  position:sticky;top:var(--hdr-h);z-index:200;
}
.itab{padding:4px 12px;border:1px solid var(--bd);border-radius:3px;
  cursor:pointer;color:var(--d);font-family:var(--f);font-size:11px;
  background:transparent;transition:all .12s;}
.itab:hover{border-color:var(--b);color:var(--t);}
.itab.active{background:var(--bg3);color:#fff;border-color:var(--b);}

/* ── VIEWS ── */
.iviews{padding:16px 24px;}
.iview{display:none;}.iview.active{display:block;}

/* ── REPORT BUTTONS ── */
.report-tabs{display:flex;align-items:center;gap:6px;margin-bottom:14px;flex-wrap:wrap;}
.rtab-lbl{font-size:9px;color:var(--d);letter-spacing:1px;}
.rtab{padding:4px 12px;border:1px solid var(--bd);border-radius:3px;
  cursor:pointer;color:var(--d);font-family:var(--f);font-size:11px;
  background:transparent;transition:all .12s;}
.rtab.active{background:var(--g);color:#000;border-color:var(--g);font-weight:bold;}
.rtab.disabled{opacity:.3;cursor:not-allowed;}

/* ── METRIC CARDS ── */
.mcards{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px;}
.mc{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:14px 16px;overflow:hidden;min-width:0;}
.mc-lbl{font-size:9px;color:#fff;letter-spacing:.5px;margin-bottom:6px;white-space:nowrap;}
.mc-val{font-size:clamp(16px,2.5vw,26px);font-weight:bold;letter-spacing:.3px;
  word-break:break-all;line-height:1.15;}
.mc-chg{font-size:11px;margin-top:5px;}
.mc-sub{font-size:10px;color:var(--d);margin-top:3px;}

/* ── MID 3-COL ── */
.mid{display:grid;grid-template-columns:1fr 200px 1fr;gap:10px;margin-bottom:12px;}
.panel{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:14px 16px;}
.plbl{font-size:9px;color:#fff;letter-spacing:.5px;margin-bottom:10px;}

/* ── АНАЛІЗ ── рівномірна сітка 3 колонки */
.arow{margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid var(--bd);}
.arow:last-child{margin:0;padding:0;border:none;}
.arow-lbl{font-size:8px;color:var(--d);letter-spacing:.5px;margin-bottom:4px;}
.arow-dir{font-size:14px;font-weight:bold;margin-bottom:8px;}
.arow-net{font-size:12px;margin-left:8px;opacity:.85;font-weight:normal;}
/* Сітка 3 рівні колонки */
.arow-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;}
.ag-item{display:flex;flex-direction:column;gap:2px;}
.ag-lbl{font-size:8px;color:#fff;letter-spacing:.3px;}
.ag-val{font-size:11px;font-weight:bold;}

/* ── SM DIV ── */
.sm-panel{display:flex;flex-direction:column;justify-content:space-between;}
.sm-row{margin-bottom:10px;}
.sm-lbl{font-size:9px;color:var(--d);margin-bottom:3px;}
.sm-bar-bg{background:var(--bg3);border-radius:10px;height:8px;position:relative;overflow:hidden;}
.sm-mk{position:absolute;top:1px;width:8px;height:6px;border-radius:3px;transform:translateX(-50%);}
.sm-val{font-size:12px;font-weight:bold;margin-top:3px;}
.sm-hint{font-size:8px;color:var(--d);margin-top:8px;line-height:1.5;
  border-top:1px solid var(--bd);padding-top:8px;}

/* ── COT PERCENTILE ── */
.pct-sel-row{display:flex;gap:5px;align-items:center;margin-bottom:10px;flex-wrap:wrap;}
.psel-group{display:flex;gap:3px;}
.psel-sep{width:1px;height:16px;background:var(--bd);margin:0 3px;}
.psel,.pper{padding:2px 8px;border:1px solid var(--bd);border-radius:3px;
  cursor:pointer;color:var(--d);font-family:var(--f);font-size:10px;
  background:transparent;transition:all .1s;}
.psel:hover,.pper:hover{border-color:var(--b);color:var(--t);}
.psel.active,.pper.active{background:var(--bg3);color:#fff;border-color:var(--b);}
.pct-val-row{margin-bottom:8px;}
.pbar-wrap{position:relative;margin-bottom:3px;}
.pbar-bg{background:var(--bg3);border-radius:3px;height:18px;position:relative;overflow:hidden;}
.pbar-lo{position:absolute;left:0;top:0;height:100%;background:rgba(240,81,90,.3);width:15%;}
.pbar-hi{position:absolute;right:0;top:0;height:100%;background:rgba(32,212,131,.3);width:15%;}
.ptick{position:absolute;top:0;width:2px;height:100%;background:rgba(255,255,255,.25);}
.pbar-mk{position:absolute;top:2px;width:4px;height:14px;background:var(--g);border-radius:2px;transform:translateX(-50%);transition:left .3s;}
.ptick-labels{position:relative;height:16px;margin-top:2px;}
.ptlbl{position:absolute;transform:translateX(-50%);font-size:8px;color:var(--d);transition:left .3s;}
.ptlbl-cur{color:var(--t);font-weight:bold;}
.pbar-lb{display:flex;justify-content:space-between;font-size:8px;color:var(--d);margin-top:12px;}

/* ── CHART ── */
.chartbox{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:14px 16px;margin-bottom:12px;}
.chartbox-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:6px;}
.period-btns{display:flex;gap:3px;}
.per-btn{padding:2px 9px;border:1px solid var(--bd);border-radius:3px;
  cursor:pointer;color:var(--d);font-family:var(--f);font-size:10px;
  background:transparent;transition:all .1s;}
.per-btn:hover{border-color:var(--b);color:var(--t);}
.per-btn.active{background:var(--bg3);color:#fff;border-color:var(--b);}
.chart-leg{display:flex;gap:10px;font-size:10px;color:var(--d);align-items:center;flex-wrap:wrap;}
.ll{display:inline-block;width:14px;height:2px;border-radius:1px;vertical-align:middle;margin-right:4px;}
.ll-dash{display:inline-block;width:14px;height:0;border-top:2px dashed #4a9eff;vertical-align:middle;}
.cw{height:140px;position:relative;}

/* ── BAR CHARTS ── */
.bar-charts-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;}
.bar-lbl{font-size:8px;color:#fff;letter-spacing:.5px;margin-bottom:6px;}
.bar-cw{height:80px;position:relative;}

/* ── HISTORY TABLE ── */
.htable-wrap{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;overflow:hidden;margin-bottom:12px;}
.htable-hdr{padding:8px 14px;border-bottom:1px solid var(--bd);
  font-size:10px;color:#fff;letter-spacing:.5px;
  display:flex;align-items:center;justify-content:space-between;}
.hsel{display:flex;gap:4px;}
.hbtn{padding:2px 10px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;
  color:var(--d);font-family:var(--f);font-size:11px;background:transparent;}
.hbtn.active{background:var(--bg3);color:#fff;border-color:var(--b);}
.htable-scroll{overflow-x:auto;}
table.ht{width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap;}
table.ht th{padding:5px 9px;background:var(--bg3);border-bottom:1px solid var(--bd);
  color:#fff;font-weight:normal;font-size:9px;letter-spacing:.5px;text-align:right;}
table.ht .th-left{text-align:left;}
table.ht .th-date{min-width:80px;max-width:90px;width:88px;}
table.ht .th-group{text-align:center;border-left:1px solid var(--bd);}
table.ht td{padding:5px 9px;border-bottom:1px solid var(--bg3);text-align:right;}
table.ht .date-col{text-align:left;color:var(--d);width:88px;}
table.ht tr:hover td{background:rgba(52,61,90,.5)!important;}
table.ht .sep-r{border-right:1px solid var(--bd);}

.g{color:var(--g);}.r{color:var(--r);}.d{color:var(--d);}
.footer{text-align:center;padding:14px;color:var(--d);font-size:9px;
  letter-spacing:1px;border-top:1px solid var(--bd);margin-top:4px;}

/* ──────── MOBILE ──────── */
@media(max-width:640px){
  :root{--hdr-h:56px;}
  body{font-size:12px;}
  .hdr{padding:8px 12px;flex-wrap:wrap;height:auto;min-height:var(--hdr-h);}
  .hdr-t{font-size:14px;}
  .hdr-r{font-size:10px;}

  .ctabs{padding:6px 12px;gap:4px;}
  .ctab{padding:4px 10px;font-size:11px;}

  .itabs{padding:5px 12px;top:var(--hdr-h);}
  .itab{padding:3px 9px;font-size:10px;}

  .iviews{padding:10px 12px;}
  /* Картки: 2 колонки на мобілці */
  .mcards{grid-template-columns:1fr 1fr;gap:8px;}
  .mc{padding:10px 12px;}
  .mc-lbl{font-size:8px;}
  .mc-val{font-size:clamp(14px,5vw,22px);}
  .mc-chg{font-size:10px;}
  .mc-sub{font-size:9px;}

  /* Середня секція: одна колонка */
  .mid{grid-template-columns:1fr;gap:8px;}
  .sm-panel{order:3;}
  .pct-panel{order:2;}

  /* Bar charts: одна колонка */
  .bar-charts-grid{grid-template-columns:1fr;}

  /* Рядок аналізу */
  .arow-grid{grid-template-columns:1fr 1fr 1fr;}
  .ag-val{font-size:10px;}

  /* Легенда графіку */
  .chart-leg{font-size:9px;gap:6px;}
  .period-btns .per-btn{font-size:9px;padding:2px 7px;}

  /* Таблиця: менший шрифт */
  table.ht{font-size:10px;}
  table.ht th{padding:4px 6px;font-size:8px;}
  table.ht td{padding:4px 6px;}
  .date-col{width:72px!important;}

  .report-tabs{flex-wrap:wrap;gap:4px;}
  .rtab{font-size:10px;padding:3px 9px;}
}
@media(max-width:380px){
  .mcards{grid-template-columns:1fr 1fr;}
  .mc-val{font-size:clamp(12px,6vw,18px);}
}
</style>
</head>
<body>
<script>const _cd={}; const _ci={}; const Charts={}; const BarChts={};</script>
"""

HTML_FOOT="""
<script>
const CurPer={};

function selCat(cat){
  document.querySelectorAll('.ctab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.catsec').forEach(s=>s.classList.remove('active'));
  document.querySelector('[data-c="'+cat+'"]').classList.add('active');
  document.getElementById('cs_'+cat).classList.add('active');
  const first=document.querySelector('[data-cat="'+cat+'"]');
  if(first) selInst(cat,first.dataset.i);
}

function selInst(cat,key){
  document.querySelectorAll('[data-cat="'+cat+'"]').forEach(b=>b.classList.remove('active'));
  const btn=document.querySelector('[data-cat="'+cat+'"][data-i="'+key+'"]');
  if(btn) btn.classList.add('active');
  const container=document.getElementById('iv_'+cat);
  container.querySelectorAll('.iview').forEach(v=>v.classList.remove('active'));
  const sid=key.replaceAll(' ','_').replaceAll('&','n').replaceAll('/','_');
  const view=document.getElementById('iv_'+sid);
  if(view){
    view.classList.add('active');
    filterRows(sid,10);
    const n=CurPer[sid]||52;
    setTimeout(()=>{drawMainChart(sid,n);drawBarsFor(sid,n);},30);
  }
}

function setChartPer(btn,sid){
  btn.closest('.period-btns').querySelectorAll('.per-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const per=btn.dataset.per;
  const n=per==='1y'?52:per==='3y'?156:9999;
  CurPer[sid]=n;
  drawMainChart(sid,n);
  drawBarsFor(sid,n);
}

function drawMainChart(sid,nWeeks){
  const cv=document.getElementById('cv_'+sid); if(!cv) return;
  if(Charts[sid]){Charts[sid].destroy();delete Charts[sid];}
  const d=_cd[sid]; if(!d) return;
  const n=Math.min(nWeeks,d.dates.length);
  Charts[sid]=new Chart(cv.getContext('2d'),{
    type:'line',
    data:{
      labels:d.dates.slice(-n),
      datasets:[
        {label:'Large Spec',   data:d.ls.slice(-n),borderColor:'#20d483',backgroundColor:'rgba(32,212,131,.07)',borderWidth:1.5,pointRadius:0,tension:.3,fill:true},
        {label:'Commercials',  data:d.cm.slice(-n),borderColor:'#f0515a',backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3},
        {label:'Small Traders',data:d.st.slice(-n),borderColor:'#4a9eff',backgroundColor:'transparent',borderWidth:1,  pointRadius:0,tension:.3,borderDash:[3,3]},
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,animation:false,
      interaction:{mode:'index',intersect:false},
      plugins:{
        legend:{display:false},
        tooltip:{
          backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,
          titleColor:'#dde2ee',bodyColor:'#dde2ee',
          titleFont:{family:'Courier New',size:10},bodyFont:{family:'Courier New',size:10},
          callbacks:{label:ctx=>' '+ctx.dataset.label+': '+fmtFull(ctx.parsed.y)}
        }
      },
      scales:{
        x:{display:true,
          ticks:{color:'#8090b0',font:{family:'Courier New',size:8},
            maxTicksLimit:8,
            callback:function(v,i){return i%Math.ceil(n/8)===0?this.getLabelForValue(v):'';}},
          grid:{display:false},border:{display:false}},
        y:{display:true,
          grid:{color:'rgba(52,61,90,.8)',lineWidth:.5},
          ticks:{color:'#8090b0',font:{family:'Courier New',size:9},maxTicksLimit:4,callback:v=>fmtV(v,true)},
          border:{display:false}}
      }
    }
  });
}

function drawBarsFor(sid,nWeeks){
  const d=_cd[sid]; if(!d) return;
  const n=Math.min(nWeeks,d.dates.length);
  const dates=d.dates.slice(-n);
  drawOneBar('barcv_ls_'+sid,dates,d.ld.slice(-n));
  drawOneBar('barcv_cm_'+sid,dates,d.cd.slice(-n));
  drawOneBar('barcv_st_'+sid,dates,d.sd.slice(-n));
}

function drawOneBar(cvId,dates,data){
  const cv=document.getElementById(cvId); if(!cv) return;
  const key='b_'+cvId;
  if(BarChts[key]){BarChts[key].destroy();delete BarChts[key];}
  const colors=data.map(v=>v>=0?'rgba(32,212,131,.75)':'rgba(240,81,90,.75)');
  BarChts[key]=new Chart(cv.getContext('2d'),{
    type:'bar',
    data:{labels:dates,datasets:[{data:data,backgroundColor:colors,borderWidth:0,borderRadius:1}]},
    options:{
      responsive:true,maintainAspectRatio:false,animation:false,
      plugins:{
        legend:{display:false},
        tooltip:{
          backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,
          titleColor:'#dde2ee',bodyColor:'#dde2ee',
          titleFont:{family:'Courier New',size:9},bodyFont:{family:'Courier New',size:9},
          callbacks:{label:ctx=>fmtFull(ctx.parsed.y)}
        }
      },
      scales:{
        x:{display:false},
        y:{display:true,
          grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},
          ticks:{color:'#8090b0',font:{family:'Courier New',size:8},maxTicksLimit:3,callback:v=>fmtV(v,true)},
          border:{display:false}}
      }
    }
  });
}

function pctSel(btn,sid){
  btn.closest('.psel-group').querySelectorAll('.psel').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); updatePctBar(sid);
}
function pperSel(btn,sid){
  btn.closest('.psel-group').querySelectorAll('.pper').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); updatePctBar(sid);
}
function updatePctBar(sid){
  const view=document.getElementById('iv_'+sid); if(!view) return;
  const p  =view.querySelector('.psel.active')?.dataset.p  ||'ls';
  const per=view.querySelector('.pper.active')?.dataset.per||'all';
  const val=_ci[sid]?.[p]?.[per]??50;
  const pos=Math.min(Math.max(val,0),100);
  const col=val<15?'#f0515a':val>85?'#20d483':'#dde2ee';
  const lbl=val<15?'— екстрем. шорт':val>85?'— екстрем. лонг':'— нейтральна зона';
  const mk=document.getElementById('pctmk_'+sid);
  const valEl=document.getElementById('pctval_'+sid);
  const lblEl=document.getElementById('pctlbl_'+sid);
  const curEl=document.getElementById('pctcur_'+sid);
  if(mk)    mk.style.left=pos+'%';
  if(valEl){valEl.style.color=col;valEl.textContent=val.toFixed(1)+'%';}
  if(lblEl) lblEl.textContent=lbl;
  if(curEl){curEl.style.left=pos+'%';curEl.textContent=val.toFixed(1)+'%';}
}

function setHist(btn,sid){
  const n=parseInt(btn.dataset.n);
  btn.closest('.htable-hdr').querySelectorAll('.hbtn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); filterRows(sid,n);
}
function filterRows(sid,n){
  const view=document.getElementById('iv_'+sid); if(!view) return;
  view.querySelectorAll('table.ht tbody tr').forEach(tr=>{
    tr.style.display=parseInt(tr.dataset.row)<n?'':'none';
  });
}

function fmtV(n,short=false){
  if(n===null||isNaN(n)) return'—';
  n=Math.round(n);
  if(short){
    if(Math.abs(n)>=1e6) return(n/1e6).toFixed(1)+'M';
    if(Math.abs(n)>=1e3) return(n/1e3).toFixed(0)+'K';
    return''+n;
  }
  const sign=n>0?'+':n<0?'-':'';
  return sign+Math.abs(n).toLocaleString('uk-UA');
}
function fmtFull(n){
  if(n===null||isNaN(n)) return'—';
  n=Math.round(n);
  const sign=n>0?'+':n<0?'-':'';
  return sign+Math.abs(n).toLocaleString('uk-UA');
}

const firstCat=document.querySelector('.ctab');
if(firstCat) selCat(firstCat.dataset.c);
</script>
</body>
</html>
"""


# ================================================================
# 🚀  ЗАПУСК
# ================================================================
def main():
    print()
    print("="*55)
    print("   COT Dashboard Generator v7")
    print("="*55)
    print()
    OUTPUT_FILE.parent.mkdir(parents=True,exist_ok=True)
    try:
        data=load_all()
    except FileNotFoundError as e:
        print(e); return
    if not data:
        print("❌  Дані порожні."); return
    print("🔧  Генеруємо HTML...")
    html=generate_html(data)
    OUTPUT_FILE.write_text(html,encoding='utf-8')
    kb=OUTPUT_FILE.stat().st_size/1024
    print(f"✅  Збережено: {OUTPUT_FILE}  ({kb:.0f} KB)")
    print("🌐  Відкриваємо браузер...")
    webbrowser.open(OUTPUT_FILE.as_uri())
    print("\n✨  Готово!\n")

if __name__=='__main__':
    main()