#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""COT Dashboard Generator v15 — Legacy + TFF + Disaggregated"""

import math, pandas as pd, json, webbrowser
from pathlib import Path
from datetime import datetime

BASE_DIR    = Path(__file__).parent
DATA_FILE   = BASE_DIR / "data" / "COT_OVERVIEW.xlsx"
TFF_FILE    = BASE_DIR / "data" / "COT_TFF_REPORTS.xlsx"
DISAG_FILE  = BASE_DIR / "data" / "COT_DISAGRAGATE_REPORTS.xlsx"
OUTPUT_FILE = BASE_DIR / "output" / "cot_dashboard.html"
# v60: сезонність — читається локально при генерації, жодних мережевих запитів
SEASON_FILE = BASE_DIR / "output" / "seasonality_test.json"
SEASON_SIDS = {'EUR'}   # блок сезонності показуємо ТІЛЬКИ для цих інструментів

SKIP_SHEETS      = {'overview', 'info'}
TFF_SKIP_SHEETS  = {'tff', 'info', 'lkup'}
DISAG_SKIP_SHEETS= {'disagragate', 'info', 'lkup'}

TFF_SHEET_TO_SID = {
    'EURO':'EUR','DOW 30':'DOW_30','RUSSELL 2K':'RUSSELL2K',
    'ETH CASH':'ETH_CASH','ETH NANO':'ETH_NANO','ETH MICRO':'ETH_MICRO',
    'BTC NANO':'BTC_NANO','BTC MICRO':'BTC_MICRO',
    'NAT GAS':'NAT_GAS','WTI CRUDE':'WTI_CRUDE',
    'S&P500':'SP500','NASDAQ':'NASDAQ','VIX':'VIX',
}
DISAG_SHEET_TO_SID = {
    'NAT GAS':'NAT_GAS','WTI CRUDE':'WTI_CRUDE',
    'SOYBEAN MEAL':'SOYBEAN_MEAL','SOYBEAN OIL':'SOYBEAN_OIL',
}

CHART_WEEKS=260; SPARK_WEEKS=26; HISTORY=52

COLOR_LS='#4a9eff'; COLOR_CM='#20d483'; COLOR_ST='#f0515a'
TFF_COLOR_LEV='#f0b429'; TFF_COLOR_AM='#4a9eff'; TFF_COLOR_DL='#20d483'

# Disaggregated кольори (аналогічно TFF)
DISAG_COLOR_MM ='#a78bfa'  # Managed Money    — фіолетовий
DISAG_COLOR_PM ='#20d483'  # Prod/Merchant    — зелений
DISAG_COLOR_SD ='#f0b429'  # Swap Dealers     — жовтий

TFF_COL={'date':1,'am_cl':11,'am_cs':12,'am_net':15,'dl_cl':4,'dl_cs':5,'dl_net':8,'lev_cl':18,'lev_cs':19,'lev_net':22,'oi':34}
TFF_DATA_START = 19  # перший рядок даних TFF = pandas індекс 19
# M_Money: cl=4, cs=5, net=8 | Prod/Merch: cl=11, cs=12, net=15
# Swap Dealers: cl=18, cs=19, net=22 | OI=34
DISAG_COL={'date':1,'mm_cl':4,'mm_cs':5,'mm_net':8,'pm_cl':11,'pm_cs':12,'pm_net':15,'sd_cl':18,'sd_cs':19,'sd_net':22,'oi':34}
DISAG_DATA_START=19  # v53: перший рядок даних — pandas-індекс 19 (Excel row 20).
                     # Було 20 — найсвіжіший тиждень щоразу відкидався.

CROP_FILE = BASE_DIR / "data" / "ALL_Crops_Dashboard.xlsx"

CROP_META = {
    'Corn':         {'emoji':'🌽','sid':'CORN',        'color':'#f59420','tv':'ZC1!',
                     'stages':[('Planted',1,2,3),('Emerged',4,5,6),('Silked',7,8,9),('Dough',10,11,12),('Dent',13,14,15),('Mature',16,17,18),('Harvested',19,20,21)]},
    'Soybeans':     {'emoji':'🫘','sid':'SOYBEAN',      'color':'#20d483','tv':'ZS1!',
                     'stages':[('Planted',1,2,3),('Emerged',4,5,6),('Blooming',7,8,9),('Setting Pods',10,11,12),('Dropping Leaves',13,14,15),('Harvested',16,17,18)]},
    'Spring_Wheat': {'emoji':'🌾','sid':'SPRING WHEAT','color':'#a78bfa','tv':'MWE1!',
                     'stages':[('Planted',1,2,3),('Emerged',4,5,6),('Headed',7,8,9),('Harvested',10,11,12)]},
    'Cotton':       {'emoji':'🌿','sid':'COTTON',       'color':'#22d3ee','tv':'CT1!',
                     'stages':[('Planted',1,2,3),('Squaring',4,5,6),('Blooming',7,8,9),('Bolls Open',10,11,12),('Harvested',13,14,15)]},
    'Rice':         {'emoji':'🌾','sid':'RICE',         'color':'#f0b429','tv':'ZR1!',
                     'stages':[('Planted',1,2,3),('Emerged',4,5,6),('Headed',7,8,9),('Harvested',10,11,12)]},
}


REPORTS=[
    {'id':'usda_crop', 'name':'USDA Crop Progress','sched':'Пн 22:00 (кві-лис)','tag':None},
    {'id':'eia_petrol','name':'EIA Petroleum',     'sched':'Ср 16:30',           'tag':None},
    {'id':'usda_exp',  'name':'USDA Export Sales', 'sched':'Чт 14:30',           'tag':None},
    {'id':'cot_cftc',  'name':'COT Report (CFTC)', 'sched':'Пт 21:30',           'tag':None},
    {'id':'usda_wasde','name':'USDA WASDE',        'sched':'~12 число, 18:00',   'tag':'Місячний'},
    {'id':'usda_oil',  'name':'USDA Oilseeds',     'sched':'з WASDE, 18:15',     'tag':'Місячний'},
]
REPORT_RELEVANCE={
    'usda_crop':  {'CORN':'direct','WHEAT':'direct','SOYBEAN':'direct','SOYBEAN_MEAL':'direct','SOYBEAN_OIL':'direct','COTTON':'direct','RICE':'direct','COFFEE':'direct','COCOA':'direct','SUGAR':'direct','OJ':'direct','CATTLE':'indirect','LUMBER':'indirect','_default':'none'},
    'eia_petrol': {'WTI_CRUDE':'direct','BRENT':'direct','NAT_GAS':'direct','SP500':'indirect','NASDAQ':'indirect','DOW_30':'indirect','RUSSELL2K':'indirect','VIX':'indirect','CAD':'indirect','_default':'none'},
    'usda_exp':   {'WHEAT':'direct','CORN':'direct','SOYBEAN':'direct','SOYBEAN_MEAL':'direct','SOYBEAN_OIL':'direct','COTTON':'direct','RICE':'direct','COFFEE':'direct','COCOA':'indirect','SUGAR':'indirect','OJ':'indirect','_default':'none'},
    'cot_cftc':   {'_default':'indirect'},
    'usda_wasde': {'WHEAT':'direct','CORN':'direct','SOYBEAN':'direct','SOYBEAN_MEAL':'direct','SOYBEAN_OIL':'direct','COTTON':'direct','RICE':'direct','SUGAR':'direct','COFFEE':'direct','COCOA':'direct','OJ':'direct','CATTLE':'direct','LUMBER':'indirect','SP500':'indirect','NASDAQ':'indirect','DOW_30':'indirect','RUSSELL2K':'indirect','VIX':'indirect','WTI_CRUDE':'indirect','BRENT':'indirect','NAT_GAS':'indirect','AUD':'indirect','CAD':'indirect','NZD':'indirect','_default':'none'},
    'usda_oil':   {'SOYBEAN':'direct','SOYBEAN_MEAL':'direct','SOYBEAN_OIL':'direct','CORN':'indirect','WHEAT':'indirect','COTTON':'indirect','WTI_CRUDE':'indirect','BRENT':'indirect','_default':'none'},
}
def get_relevance(sid_,rid): m=REPORT_RELEVANCE.get(rid,{}); return m.get(sid_,m.get('_default','none'))

DISPLAY={'SP500':'S&P 500','DOW_30':'DOW 30','RUSSELL2K':'RUSSELL 2K','NAT_GAS':'NAT GAS','WTI_CRUDE':'WTI CRUDE','SOYBEAN_MEAL':'SOYBEAN MEAL','SOYBEAN_OIL':'SOYBEAN OIL','BTC_MICRO':'BTC MICRO','BTC_NANO':'BTC NANO','ETH_CASH':'ETH CASH','ETH_MICRO':'ETH MICRO','ETH_NANO':'ETH NANO'}
# sid → символ TradingView. Інструменти, яких тут немає, отримують ту саму праву
# колонку, але із заглушкою .tv-empty замість віджета (сітка завжди двоколонкова).
# Беремо відкриті спот/CFD-тікери (TVC/CAPITALCOM), крипту — зі спота Bitstamp:
# біржові ф'ючерсні контракти (COMEX/NYMEX/CBOT/ICEUS) закриті для безкоштовних віджетів.
# SOYBEAN_MEAL, SOYBEAN_OIL, RICE, OJ, LUMBER, CATTLE свідомо відсутні — адекватного
# відкритого відповідника немає, а проксі вводив би в оману; вони йдуть без віджета.
TV_SYMBOL={
 'DXY':'CAPITALCOM:DXY','EUR':'FX:EURUSD','GBP':'FX:GBPUSD',
 'JPY':'FX_IDC:JPYUSD','AUD':'FX:AUDUSD','CAD':'FX_IDC:CADUSD',
 'CHF':'FX_IDC:CHFUSD','NZD':'FX:NZDUSD',
 'GOLD':'TVC:GOLD','SILVER':'TVC:SILVER','COPPER':'TVC:COPPER',
 'PLATINUM':'TVC:PLATINUM','PALLADIUM':'TVC:PALLADIUM',
 'SP500':'AMEX:SPY','NASDAQ':'NASDAQ:QQQ','DOW_30':'AMEX:DIA',
 'RUSSELL2K':'AMEX:IWM','VIX':'CAPITALCOM:VIX',
 'WTI_CRUDE':'TVC:USOIL','BRENT':'TVC:UKOIL','NAT_GAS':'TVC:NATURALGAS',
 'CORN':'CAPITALCOM:CORN','WHEAT':'CAPITALCOM:WHEAT','SOYBEAN':'CAPITALCOM:SOYBEAN',
 'SUGAR':'CAPITALCOM:SUGAR','COFFEE':'CAPITALCOM:COFFEE','COCOA':'CAPITALCOM:COCOA',
 'COTTON':'CAPITALCOM:COTTON',
 'BTC':'BITSTAMP:BTCUSD','BTC_MICRO':'BITSTAMP:BTCUSD',
 'BTC_NANO':'BITSTAMP:BTCUSD','ETH_CASH':'BITSTAMP:ETHUSD',
 'ETH_MICRO':'BITSTAMP:ETHUSD','ETH_NANO':'BITSTAMP:ETHUSD',
}
CATEGORIES={
    'Валюти':  ['DXY','EUR','GBP','JPY','AUD','CAD','CHF','NZD'],
    'Метали':  ['GOLD','SILVER','COPPER','PLATINUM','PALLADIUM'],
    'Індекси': ['SP500','NASDAQ','DOW_30','RUSSELL2K','VIX'],
    'Енергія': ['WTI_CRUDE','BRENT','NAT_GAS'],
    'Агро':    ['CORN','WHEAT','SOYBEAN','SOYBEAN_MEAL','SOYBEAN_OIL','SUGAR','COFFEE','COCOA','COTTON','RICE','OJ','LUMBER','CATTLE'],
    'Крипто':  ['BTC','BTC_MICRO','BTC_NANO','ETH_CASH','ETH_MICRO','ETH_NANO'],
}
OVERVIEW_TO_SID={'DXY':'DXY','EUR':'EUR','GBP':'GBP','CAD':'CAD','JPY':'JPY','AUD':'AUD','NZD':'NZD','CHF':'CHF','GOLD':'GOLD','SILVER':'SILVER','COPPER':'COPPER','PLATINUM':'PLATINUM','PALLADIUM':'PALLADIUM','SP500':'SP500','S&P500':'SP500','NASDAQ':'NASDAQ','DOW_30':'DOW_30','DOW 30':'DOW_30','RUSSELL2K':'RUSSELL2K','VIX':'VIX','WTI_CRUDE':'WTI_CRUDE','WTI CRUDE':'WTI_CRUDE','BRENT':'BRENT','NAT_GAS':'NAT_GAS','NAT GAS':'NAT_GAS','CORN':'CORN','WHEAT':'WHEAT','SOYBEAN':'SOYBEAN','SOYBEAN_MEAL':'SOYBEAN_MEAL','SOYBEAN MEAL':'SOYBEAN_MEAL','SOYBEAN_OIL':'SOYBEAN_OIL','SOYBEAN OIL':'SOYBEAN_OIL','SUGAR':'SUGAR','COFFEE':'COFFEE','COCOA':'COCOA','COTTON':'COTTON','RICE':'RICE','OJ':'OJ','LUMBER':'LUMBER','CATTLE':'CATTLE','BTC':'BTC','BTC_NANO':'BTC_NANO','BTC NANO':'BTC_NANO','BTC_MICRO':'BTC_MICRO','BTC MICRO':'BTC_MICRO','ETH_CASH':'ETH_CASH','ETH CASH':'ETH_CASH','ETH_MICRO':'ETH_MICRO','ETH MICRO':'ETH_MICRO','ETH_NANO':'ETH_NANO','ETH NANO':'ETH_NANO'}
COL={'date':1,'ls_cl':4,'ls_cs':5,'ls_pct':6,'ls_net':8,'cm_cl':11,'cm_cs':12,'cm_pct':13,'cm_net':15,'st_cl':18,'st_cs':19,'st_pct':20,'st_net':22}
DATA_START_ROW=5

def disp(n): return DISPLAY.get(n,n)
def sid(n):  return n.replace(' ','_').replace('&','n').replace('/','_')
def to_num(s): return pd.to_numeric(s,errors='coerce').fillna(0).round(2).tolist()
def norm_pct(vals):
    if not vals: return vals
    nz=[abs(v) for v in vals if v!=0]
    if nz and max(nz)<=1.5: return [round(v*100,1) for v in vals]
    return [round(v,1) for v in vals]
def fv(n,short=False,sign=False):
    try: n=int(round(float(n)))
    except: return '—'
    if short:
        if abs(n)>=1_000_000: body=f"{abs(n)/1_000_000:.1f}M"
        elif abs(n)>=1_000:   body=f"{abs(n)/1_000:.0f}K"
        else:                 body=str(abs(n))
    else: body=f"{abs(n):,}".replace(',','\u202f')
    if n>0:   return ('+'if sign else '')+body
    elif n<0: return '-'+body
    return body
def fv_full(n,sign=False):
    try: n=int(round(float(n)))
    except: return '—'
    body=f"{abs(n):,}".replace(',','\u202f')
    if n>0:   return ('+'if sign else '')+body
    elif n<0: return '-'+body
    return body
def fp(n,signed=False):
    try: v=float(n);s='+'if(signed and v>0)else''; return f"{s}{v:.1f}%"
    except: return '—'
def pct_change(chg,net):
    try:
        prev=float(net)-float(chg)
        if abs(prev)<1: return '—'
        return fp(float(chg)/abs(prev)*100,signed=True)
    except: return '—'
def cc(n):
    try: v=float(n); return 'g'if v>0 else('r'if v<0 else'd')
    except: return 'd'
def ar(n):
    try: v=float(n); return '▲'if v>0 else('▼'if v<0 else'—')
    except: return '—'
def calc_cot_index(series,weeks=None):
    s=list(series[-weeks:]) if weeks and len(series)>=weeks else list(series)
    if len(s)<2: return 50.0
    cur=s[-1];mn=min(s);mx=max(s)
    if mx==mn: return 50.0
    return round((cur-mn)/(mx-mn)*100,1)
def compute_delta(vals):
    if not vals: return []
    r=[0.0]
    for i in range(1,len(vals)): r.append(round(float(vals[i])-float(vals[i-1]),0))
    return r
OI_COL_FIXED = 24  # колонка Y "Open Interest" — однакова на всіх аркушах

def find_oi_col(df):
    # v22 FIX: раніше евристика вимагала >40% заповнених рядків, через що
    # BRENT/COPPER/BTC*/ETH* (коротка історія OI, 10-20% рядків) падали
    # на порожню колонку 23 і показували OI=0.
    # OI завжди у фіксованій колонці 24 — перевіряємо її першою.
    n=len(df)
    if OI_COL_FIXED < df.shape[1]:
        vals=pd.to_numeric(df.iloc[:,OI_COL_FIXED],errors='coerce').fillna(0)
        if (vals.abs()>100).sum() > 0 and vals.abs().max() > 1000:
            return OI_COL_FIXED
    # фолбек — стара евристика зі зниженим порогом
    for idx in range(23,min(50,df.shape[1])):
        vals=pd.to_numeric(df.iloc[:,idx],errors='coerce').fillna(0)
        if (vals.abs()>100).sum()>n*0.05 and vals.abs().mean()>100: return idx
    return OI_COL_FIXED if OI_COL_FIXED < df.shape[1] else 23
def make_sparkline(series,color,h=38):
    data=[float(v) for v in (series or [])[-SPARK_WEEKS:]]
    n=len(data)
    if n<3: return ''
    W=200;H=h;mn=min(data);mx=max(data);rng=mx-mn
    if rng==0: return ''
    def px(i): return round(i/(n-1)*W,1)
    def py(v): return round((1-(v-mn)/rng)*H,1)
    pts=[(px(i),py(v)) for i,v in enumerate(data)]
    line=' '.join(f"{x},{y}" for x,y in pts)
    zy=max(0,min(H,py(0)));area=f"0,{zy} "+line+f" {W},{zy}"
    lx,ly=pts[-1]
    return(f'<svg viewBox="0 0 {W} {H}" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:{H}px;display:block;margin-top:8px">'
           f'<polygon points="{area}" fill="{color}" opacity="0.14"/>'
           f'<polyline points="{line}" fill="none" stroke="{color}" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>'
           f'<circle cx="{lx}" cy="{ly}" r="2.5" fill="{color}"/></svg>')

def make_gauge_svg(value,color,size=74,label='COT INDEX'):
    # v27 gauge: велике число всередині дуги, підпис під дугою
    value=max(0.0,min(100.0,float(value)))
    cx=size/2; cy=size*0.44; r=size*0.37
    START_SVG=140.0;SWEEP=240.0
    def pt(d):
        a=math.radians(d); return (round(cx+r*math.cos(a),2),round(cy+r*math.sin(a),2))
    s=pt(START_SVG);e=pt(START_SVG+SWEEP)
    vs=value/100.0*SWEEP;v=pt(START_SVG+vs)
    bg=f"M{s[0]},{s[1]} A{r:.1f},{r:.1f} 0 1,1 {e[0]},{e[1]}"
    fg=f"M{s[0]},{s[1]} A{r:.1f},{r:.1f} 0 {1 if vs>180 else 0},1 {v[0]},{v[1]}" if value>0 else None
    # v28: розмір числа підбирається так, щоб воно завжди вміщалось у дугу.
    # Ціле число (макс. 3 символи "100") дозволяє більший шрифт, ніж "100.0".
    val_txt=f"{value:.0f}"
    inner_w=2*(r-1.5-2.0)                        # діаметр мінус обведення і відступ
    fs_fit =inner_w/(max(len(val_txt),1)*0.60)   # Courier: ширина символу ~0.6em
    val_fs =round(min(size*0.34,fs_fit),1)       # більше ніж було, але без виходу за дугу
    lbl_fs =round(size*0.155,1)                  # більший підпис
    val_y  =round(cy+val_fs*0.35,1)
    lbl_y  =round(cy+r+lbl_fs*0.80,1)
    return(f'<svg viewBox="0 0 {size} {size}" width="{size}" height="{size}" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;display:block">'
           f'<path d="{bg}" stroke="#252d48" stroke-width="3" fill="none" stroke-linecap="round"/>'
           +(f'<path d="{fg}" stroke="{color}" stroke-width="3" fill="none" stroke-linecap="round"/>' if fg else '')
           +f'<circle cx="{v[0]}" cy="{v[1]}" r="3.5" fill="{color}"/>'
           f'<circle cx="{s[0]}" cy="{s[1]}" r="2.5" fill="#252d48" stroke="{color}" stroke-width="1"/>'
           f'<text x="{cx}" y="{val_y}" text-anchor="middle" font-family="Courier New,monospace" font-size="{val_fs}" font-weight="bold" fill="{color}">{val_txt}</text>'
           f'<text x="{cx}" y="{lbl_y}" text-anchor="middle" font-family="Courier New,monospace" font-size="{lbl_fs}" fill="{color}" opacity="0.75">{label}</text></svg>')
def gauge_color(value,oi=False):
    if oi: return '#dde2ee'
    v=float(value)
    if v<15: return '#f0515a'
    if v>85: return '#20d483'
    return '#dde2ee'

def make_metric_card(lbl,val,chg,chg_pct,spark_series,spark_color,oi=False,gauge_val=50.0,sub_text='COT Index: —',ranked_val=None,lbl_color='#fff',bar_id=None,border_color='var(--bd)'):
    # v61: усі картки (включно з OPEN INTEREST) мають бар тижневої зміни — спарклайн не рендериться.
    # Гілка лишається як фолбек для викликів без bar_id.
    spark='' if bar_id else make_sparkline(spark_series,spark_color)
    try: ci=int(round(float(chg)))
    except: ci=0
    cc_=cc(chg)
    bg=('background:rgba(32,212,131,.20);border-radius:3px;padding:2px 7px;display:inline-block;' if ci>0
        else 'background:rgba(240,81,90,.20);border-radius:3px;padding:2px 7px;display:inline-block;' if ci<0
        else 'padding:2px 7px;display:inline-block;')
    val_str=(f'<span class="t">{fv(val)}</span>' if oi else f'<span class="{cc(val)}">{fv(val,sign=True)}</span>')
    gcol=gauge_color(gauge_val,oi=oi); g=make_gauge_svg(gauge_val,gcol,size=72,label='COT INDEX')
    # COT INDEX Ranked (M) — другий gauge поруч
    if ranked_val is not None and not oi:
        rgcol=gauge_color(ranked_val)
        rg=make_gauge_svg(ranked_val,rgcol,size=58,label='RANK')
        gauge_col=f'<div class="mc-right mc-gauges"><div class="mc-gauge-wrap">{g}</div><div class="mc-gauge-wrap">{rg}</div></div>'
    else:
        gauge_col=f'<div class="mc-right">{g}</div>'
    # v59: WEEKLY ΔNet всередині картки — фіксовані 26 тижнів (не залежить від перемикача періоду)
    bar_block=(f'<div class="mc-bar"><canvas id="{bar_id}"></canvas></div>' if bar_id else '')
    # v60: кольорова рамка картки за групою (напівпрозорий hex), для OI — стандартний --bd
    return(f'<div class="mc" style="border:1px solid {border_color}">'
           f'<div class="mc-lbl" style="color:{lbl_color}">{lbl}</div>'
           f'<div class="mc-inner"><div class="mc-left">'
           f'<div class="mc-val">{val_str}</div>'
           f'<div class="mc-chg-wrap"><span class="{cc_}" style="{bg}">{ar(chg)} {fv_full(abs(ci))}<span class="mc-wtag"> за тиждень</span></span></div>'
           f'<div class="mc-pct {cc_}">{chg_pct}</div>'
           +(f'<div class="mc-sub">{sub_text}</div>' if sub_text else '')
           +f'</div>'+gauge_col+f'</div>{spark}{bar_block}</div>')

def analysis_row(group_label,group_color,net,cl,cs,chg,chg_pct,idx=None):
    # v26 legacy analysis — вигляд як TFF: метрики зліва, gauges COT INDEX справа
    dc='g'if net>0 else'r'
    idx=idx or {}
    g_all=idx.get('all',50.0);g_3y=idx.get('3y',50.0);g_1y=idx.get('1y',50.0)
    col_all=gauge_color(g_all);col_3y=gauge_color(g_3y);col_1y=gauge_color(g_1y)
    gauge_all=make_gauge_svg(g_all,col_all,size=62,label='COT ALL')
    gauge_3y =make_gauge_svg(g_3y, col_3y, size=62,label='COT 3Y')
    gauge_1y =make_gauge_svg(g_1y, col_1y, size=62,label='COT 1Y')
    # v60: лівий кольоровий акцент рядка за групою
    return(f'<div class="tff-a-row" style="border-left:2px solid {group_color}">'
           f'<div class="tff-a-left">'
           f'<div class="tff-a-name" style="color:{group_color}">{group_label}</div>'
           f'<div class="tff-a-metrics">'
           f'<div class="tff-ag-item"><span class="ag-lbl">CHG LONG</span>'
           f'<span class="{cc(cl)} ag-val">{fv_full(cl,sign=True)}</span></div>'
           f'<div class="tff-ag-item"><span class="ag-lbl">CHG SHORT</span>'
           f'<span class="{cc(cs)} ag-val">{fv_full(cs,sign=True)}</span></div>'
           f'<div class="tff-ag-item"><span class="ag-lbl">Δ NET</span>'
           f'<span class="{cc(chg)} ag-val">{fv_full(chg,sign=True)}'
           f'<span class="ag-pct"> ({chg_pct})</span></span></div>'
           f'<div class="tff-ag-item"><span class="ag-lbl">NET POS</span>'
           f'<span class="{dc} ag-val ag-bignet">{fv_full(net,sign=True)}</span></div>'
           f'</div></div>'
           f'<div class="tff-a-gauges">'
           f'<div class="tff-a-gwrap">{gauge_all}</div>'
           f'<div class="tff-a-gwrap">{gauge_3y}</div>'
           f'<div class="tff-a-gwrap">{gauge_1y}</div>'
           f'</div></div>')

def make_reports_panel(s_id):
    rows=[]
    for rpt in REPORTS:
        rid=rpt['id'];rel=get_relevance(s_id,rid)
        tag=f'<span class="rpt-tag">{rpt["tag"]}</span>' if rpt['tag'] else ''
        icon=('<span class="rel-icon rel-d">●</span>' if rel=='direct' else
              '<span class="rel-icon rel-i">■</span>' if rel=='indirect' else
              '<span class="rel-icon rel-n">○</span>')
        rc='rpt-row' if rel!='none' else 'rpt-row rpt-dim'
        rows.append(f'<div class="{rc}" id="rpt_{s_id}_{rid}"><div class="rpt-rel">{icon}</div>'
                    f'<div class="rpt-info"><div class="rpt-name">{rpt["name"]}{tag}</div>'
                    f'<div class="rpt-sched">{rpt["sched"]}</div></div>'
                    f'<div class="rpt-btns">'
                    f'<button class="rb rb-l" onclick="setRpt(\'{s_id}\',\'{rid}\',\'long\')">L</button>'
                    f'<button class="rb rb-n active" onclick="setRpt(\'{s_id}\',\'{rid}\',\'neutral\')">N</button>'
                    f'<button class="rb rb-s" onclick="setRpt(\'{s_id}\',\'{rid}\',\'short\')">S</button>'
                    f'</div></div>')
    return(f'<div class="panel rpt-panel" id="rpts_{s_id}">'
           f'<div class="plbl-row"><div class="plbl" style="margin:0">ЗВІТИ</div>'
           f'<div class="rel-legend"><span class="rel-icon rel-d">●</span>прямий&nbsp;<span class="rel-icon rel-i">■</span>непрямий</div></div>'
           +''.join(rows)+'</div>')

def sm_bar(val,label):
    v=float(val) if val else 0.0
    pct=min(max((v+1)/2*100,0),100)
    color='#20d483'if v>0 else('#f0515a'if v<0 else'#6a7290')
    cls='g'if v>0 else('r'if v<0 else'd')
    return(f'<div class="sm-row"><div class="sm-lbl">{label}</div>'
           f'<div class="sm-bar-bg"><div class="sm-mk" style="left:{pct:.0f}%;background:{color}"></div></div>'
           f'<div class="sm-val {cls}">{v:+.2f}</div></div>')

def intensity_bg(val,max_abs):
    if max_abs==0: return ''
    try:
        v=float(val);ratio=min(abs(v)/max_abs,1.0);op=0.10+ratio*0.67
        if v>0:   return f' style="background:rgba(32,212,131,{op:.2f})"'
        elif v<0: return f' style="background:rgba(240,81,90,{op:.2f})"'
    except: pass
    return ''

OVERVIEW_TABLE=[]
def read_overview(xl):
    global OVERVIEW_TABLE
    result={};OVERVIEW_TABLE=[]
    try:
        raw=xl.parse('overview',header=None)
        def safe_date(c):
            v=pd.to_datetime(c,errors='coerce'); return v.strftime('%d.%m.%Y') if pd.notna(v) else '—'
        try: rep_date=safe_date(raw.iloc[1,2]);today_date=safe_date(raw.iloc[1,5])
        except: rep_date='—';today_date='—'
        OVERVIEW_TABLE.append(('_meta',rep_date,today_date))
        OV_GROUP_UA={'CURRENCIES':'ВАЛЮТИ','METALS':'МЕТАЛИ','METALAS':'МЕТАЛИ',
                     'INDEXES':'ІНДЕКСИ','ENERGY':'ЕНЕРГІЯ','SOFTS':'СОФТИ',
                     'GRAINS':'ЗЕРНОВІ','CRYPTO':'КРИПТО'}
        cur_group=''
        for i in range(4,len(raw)):
            row=raw.iloc[i]
            grp0=str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            asset=str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            # Група — текстова назва у колонці A, коли колонка B порожня
            if (not asset or asset=='nan') and grp0 and grp0!='nan' and not grp0.replace('.','').replace(',','').isdigit():
                cur_group=OV_GROUP_UA.get(grp0.upper(),grp0)
                OVERVIEW_TABLE.append(('_group',cur_group));continue
            if not asset or asset=='nan': continue
            def safe(c):
                v=pd.to_numeric(row.iloc[c],errors='coerce'); return float(v) if pd.notna(v) else None
            cot_ls=safe(4)
            if cot_ls is None: continue  # рядок без даних (напр. ETH CASH) — пропускаємо
            s_=OVERVIEW_TO_SID.get(asset,asset)
            sm_div=safe(8);sm_div_3m=safe(18);sm_div_6m=safe(19)
            crowded=str(row.iloc[20]).strip() if pd.notna(row.iloc[20]) else '—'
            cm_lead=str(row.iloc[13]).strip() if pd.notna(row.iloc[13]) else '—'
            result[s_]={'cot_ls_all':round(cot_ls*100,1),'cot_cm_all':round((safe(5) or 0)*100,1),'cot_st_all':round((safe(6) or 0)*100,1),'sm_div':round(sm_div or 0,3),'sm_div_3m':round(sm_div_3m or 0,3),'sm_div_6m':round(sm_div_6m or 0,3)}
            OVERVIEW_TABLE.append({'asset':asset,'sid':s_,'group':cur_group,'net_ls':safe(2),'net_cm':safe(3),'chg_pct_ls':safe(9),'chg_pct_cm':safe(11),'cot_ls':round(cot_ls*100,1),'cot_cm':round((safe(5) or 0)*100,1),'cot_st':round((safe(6) or 0)*100,1),'chg_ls':safe(10),'chg_cm':safe(12),'oi_chg_pct':safe(14),'sm_div':sm_div,'sm_div_3m':sm_div_3m,'sm_div_6m':sm_div_6m,'crowded':crowded,'cm_lead':cm_lead})
    except Exception as e: print(f"  ⚠  overview: {e}")
    return result

def read_sheet(xl,name,overview):
    try:
        raw=xl.parse(name,header=None)
        if raw.shape[0]<DATA_START_ROW+2 or raw.shape[1]<20: return None
        df=raw.iloc[DATA_START_ROW:].reset_index(drop=True).copy()
        dates_raw=pd.to_datetime(df.iloc[:,COL['date']],errors='coerce')
        valid=dates_raw.notna()&(dates_raw.dt.year>2000)
        df=df[valid].copy();df['_dt']=dates_raw[valid].values
        if df.empty: return None
        df=df.sort_values('_dt').reset_index(drop=True)
        oi_col=find_oi_col(df)
        def gc(idx): return to_num(df.iloc[:,idx]) if idx<df.shape[1] else [0.0]*len(df)
        ls_net=gc(COL['ls_net']);cm_net=gc(COL['cm_net']);st_net=gc(COL['st_net'])
        ls_cl=gc(COL['ls_cl']);ls_cs=gc(COL['ls_cs'])
        cm_cl=gc(COL['cm_cl']);cm_cs=gc(COL['cm_cs'])
        st_cl=gc(COL['st_cl']);st_cs=gc(COL['st_cs'])
        ls_pct=norm_pct(gc(COL['ls_pct']));cm_pct=norm_pct(gc(COL['cm_pct']));st_pct=norm_pct(gc(COL['st_pct']))
        oi_all=gc(oi_col);all_dates=df['_dt'].dt.strftime('%d.%m.%Y').tolist()
        i0=-1;i1=-2 if len(all_dates)>1 else -1
        ov=overview.get(sid(name),{})
        def cot_idx(series,all_key):
            return{'all':ov.get(all_key,calc_cot_index(series)),'3y':calc_cot_index(series,156),'1y':calc_cot_index(series,52),'6m':calc_cot_index(series,26),'3m':calc_cot_index(series,13)}
        # COT INDEX Ranked (M) з колонок 87-101 (значення 0-1 → ×100)
        def gm(ci):
            if ci>=df.shape[1]: return None
            v=pd.to_numeric(df.iloc[valid_idx[0] if hasattr(valid_idx,'__len__') else 0, ci],errors='coerce')
            return round(float(v)*100,1) if pd.notna(v) else None
        # Беремо значення COT INDEX(M) з останнього рядка відсортованого df
        def gcm(ci):
            # COT INDEX Ranked(M): у файлі значення зберігаються як частка 0..1.
            # ЗАХИСТ: якщо на позиції випадково опиниться дата/timestamp або
            # інше сміття (як було у старому файлі до появи ranked-колонок),
            # pd.to_numeric дасть наносекунди (~1e15) -> відсікаємо діапазоном.
            if ci>=df.shape[1]: return None
            try:
                raw=df.iloc[-1, ci]
                # дати/timestamp одразу відкидаємо
                if isinstance(raw, pd.Timestamp): return None
                v=pd.to_numeric(raw, errors='coerce')
                if not pd.notna(v): return None
                v=float(v)
                if v<=0: return None
                if v<=1.0:      pct=round(v*100,1)   # частка -> %
                elif v<=100.0:  pct=round(v,1)        # вже у %
                else:           return None           # аномалія (timestamp тощо)
                return pct
            except: return None
        # Колонки (0-based): All Time: LS=87,CM=88,ST=89 | 3y: 90,91,92 | 1y: 93,94,95 | 6mo: 96,97,98 | 3mo: 99,100,101
        cot_idx_m={'ls':{'all':gcm(87),'3y':gcm(90),'1y':gcm(93),'6m':gcm(96),'3m':gcm(99)},
                   'cm':{'all':gcm(88),'3y':gcm(91),'1y':gcm(94),'6m':gcm(97),'3m':gcm(100)},
                   'st':{'all':gcm(89),'3y':gcm(92),'1y':gcm(95),'6m':gcm(98),'3m':gcm(101)}}
        def stats(net_s,cl_s=None,cs_s=None):
            def mm(s):
                sv=[float(v) for v in s if v!=0];s1y=sv[-52:] if len(sv)>=52 else sv
                return{'max_all':max(sv)if sv else 0,'min_all':min(sv)if sv else 0,'max_1y':max(s1y)if s1y else 0,'min_1y':min(s1y)if s1y else 0}
            r=mm(net_s);r['cl']=mm(cl_s) if cl_s is not None else None;r['cs']=mm(cs_s) if cs_s is not None else None;return r
        N=min(CHART_WEEKS,len(df));cdf=df.tail(N).reset_index(drop=True)
        def gcc(idx): return to_num(cdf.iloc[:,idx]) if idx<cdf.shape[1] else [0.0]*N
        ls_w=gcc(COL['ls_net']);cm_w=gcc(COL['cm_net']);st_w=gcc(COL['st_net']);oi_w=gcc(oi_col)
        # v61: 'oid' — тижнева зміна OI (як ld/cd/sd), для бару в картці OPEN INTEREST
        chart={'dates':cdf['_dt'].dt.strftime('%d.%m.%Y').tolist(),'ls':ls_w,'cm':cm_w,'st':st_w,'oi':oi_w,'ld':compute_delta(ls_w),'cd':compute_delta(cm_w),'sd':compute_delta(st_w),'oid':compute_delta(oi_w)}
        hdf=df.tail(HISTORY).reset_index(drop=True)
        def gch(idx): return to_num(hdf.iloc[:,idx]) if idx<hdf.shape[1] else [0.0]*len(hdf)
        def gch_pct(idx):
            # %-колонки: у файлі частки 0..1 -> множимо на 100 (сирі значення, без округлення to_num)
            if idx>=hdf.shape[1]: return [None]*len(hdf)
            _v=pd.to_numeric(hdf.iloc[:,idx],errors='coerce')
            return [round(float(x)*100,1) if pd.notna(x) else None for x in _v]
        def gch_sm(idx,fkey):
            if idx<hdf.shape[1]:
                vals=pd.to_numeric(hdf.iloc[:,idx],errors='coerce').tolist()
                if any(v is not None and not(v!=v) for v in vals): return vals
            return [ov.get(fkey,None)]*len(hdf)
        hist={'dates':hdf['_dt'].dt.strftime('%d.%m.%Y').tolist(),'ls_cl':gch(COL['ls_cl']),'ls_cs':gch(COL['ls_cs']),'ls_net':gch(COL['ls_net']),'cm_cl':gch(COL['cm_cl']),'cm_cs':gch(COL['cm_cs']),'cm_net':gch(COL['cm_net']),'st_cl':gch(COL['st_cl']),'st_cs':gch(COL['st_cs']),'st_net':gch(COL['st_net']),'oi':gch(oi_col),'ls_pct_row':gch_pct(6),'ls_oich_row':gch_pct(7),'cm_pct_row':gch_pct(13),'cm_oich_row':gch_pct(14),'st_pct_row':gch_pct(20),'st_oich_row':gch_pct(21),'sm_div_row':gch_sm(57,'sm_div'),'sm_div_6m_row':gch_sm(58,'sm_div_6m'),'sm_div_3m_row':gch_sm(59,'sm_div_3m')}
        oi_cur=float(oi_all[i0]);oi_prev=float(oi_all[i1])
        oi_pct=round((oi_cur-oi_prev)/abs(oi_prev)*100,2) if oi_prev!=0 else 0.0
        ls_chg=round(float(ls_net[i0])-float(ls_net[i1]),0)
        cm_chg=round(float(cm_net[i0])-float(cm_net[i1]),0)
        st_chg=round(float(st_net[i0])-float(st_net[i1]),0)
        oi_chg=round(oi_cur-oi_prev,0)
        _oi_st=stats(oi_all)
        oi_cap=round(oi_cur/_oi_st['max_all']*100,1) if _oi_st.get('max_all',0)>0 else 50.0
        return{'name':name,'display':disp(name),'sid':sid(name),'table':_tbl_payload(df),'chart':chart,'hist':hist,'stats_ls':stats(ls_net,ls_cl,ls_cs),'stats_cm':stats(cm_net,cm_cl,cm_cs),'stats_st':stats(st_net,st_cl,st_cs),'stats_oi':stats(oi_all),'cot_idx':{'ls':cot_idx(ls_net,'cot_ls_all'),'cm':cot_idx(cm_net,'cot_cm_all'),'st':cot_idx(st_net,'cot_st_all')},'cot_idx_m':cot_idx_m,'sm':{'div':ov.get('sm_div',0.0),'div_3m':ov.get('sm_div_3m',0.0),'div_6m':ov.get('sm_div_6m',0.0)},'spark':{'ls':ls_net,'cm':cm_net,'st':st_net,'oi':oi_all},'oi_capacity':oi_cap,'cur':{'date':all_dates[i0],'ls_net':ls_net[i0],'cm_net':cm_net[i0],'st_net':st_net[i0],'ls_pct':ls_pct[i0],'cm_pct':cm_pct[i0],'st_pct':st_pct[i0],'ls_cl':ls_cl[i0],'ls_cs':ls_cs[i0],'cm_cl':cm_cl[i0],'cm_cs':cm_cs[i0],'oi':oi_cur,'oi_pct':oi_pct,'ls_chg':ls_chg,'cm_chg':cm_chg,'st_chg':st_chg,'oi_chg':oi_chg,'ls_chg_pct':pct_change(ls_chg,ls_net[i0]),'cm_chg_pct':pct_change(cm_chg,cm_net[i0]),'st_chg_pct':pct_change(st_chg,st_net[i0]),'oi_chg_pct':fp(oi_pct,signed=True)}}
    except Exception as e: print(f"  ❌  {name}: {e}"); return None

def load_all():
    print(f"📂  {DATA_FILE}")
    if not DATA_FILE.exists(): raise FileNotFoundError(f"\n❌  Файл не знайдено: {DATA_FILE}")
    xl=pd.ExcelFile(DATA_FILE)
    print("    Читаємо overview...")
    overview=read_overview(xl)
    sheets=[s for s in xl.sheet_names if s not in SKIP_SHEETS]
    print(f"    Вкладок: {len(sheets)}\n")
    data={}
    for s in sheets:
        r=read_sheet(xl,s,overview)
        if r: data[s]=r; print(f"  ✓  {s:20s}  {r['cur']['date']}  LS={fv(r['cur']['ls_net'],True,True):>8s}  COTls={r['cot_idx']['ls']['all']:>5.1f}%")
        else: print(f"  ✗  {s}")
    print(f"\n✅  Legacy: {len(data)} інструментів\n")
    return data

# ================================================================
# TFF — читання даних (незмінно з v14)
# ================================================================
def detect_tff_cols(raw):
    n_cols=raw.shape[1];col_text=[]
    for ci in range(n_cols):
        parts=[]
        for ri in range(min(8,len(raw))):
            v=str(raw.iloc[ri,ci]).strip().lower()
            if v not in('','nan'): parts.append(v)
        col_text.append(' '.join(parts))
    def find(pats_list):
        for ci,text in enumerate(col_text):
            for pats in pats_list:
                if all(p in text for p in pats): return ci
        return None
    r={}
    r['dl_cl']=find([('dealer','change','long'),('dealer','chg','long')])
    r['dl_cs']=find([('dealer','change','short'),('dealer','chg','short')])
    r['dl_net']=find([('dealer','net'),('dealer','nett')])
    r['am_cl']=find([('asset','change','long'),('asset','chg','long')])
    r['am_cs']=find([('asset','change','short'),('asset','chg','short')])
    r['am_net']=find([('asset','net'),('asset mgr','net'),('asset_mgr','net')])
    r['lev_cl']=find([('lev','change','long'),('lev','chg','long'),('leveraged','change','long')])
    r['lev_cs']=find([('lev','change','short'),('lev','chg','short'),('leveraged','change','short')])
    r['lev_net']=find([('lev','net'),('leveraged','net'),('lev money','net')])
    found=sum(1 for v in r.values() if v is not None)
    if found>=6: print(f"      ✓ TFF колонки знайдено ({found}/9)"); return r
    print(f"      ⚠ TFF заголовки не знайдено ({found}/9), використовую TFF_COL"); return {}

def read_tff_sheet(xl_tff,name):
    try:
        raw=xl_tff.parse(name,header=None)
        if raw.shape[0]<TFF_DATA_START+2 or raw.shape[1]<10: return None
        auto=detect_tff_cols(raw);cols=TFF_COL.copy()
        if auto: cols.update({k:v for k,v in auto.items() if v is not None})
        df=raw.iloc[TFF_DATA_START:].reset_index(drop=True).copy()
        dates_raw=pd.to_datetime(df.iloc[:,cols['date']],errors='coerce')
        valid=dates_raw.notna()&(dates_raw.dt.year>2000)
        df=df[valid].copy();df['_dt']=dates_raw[valid].values
        if df.empty: return None
        df=df.sort_values('_dt').reset_index(drop=True)
        def gc(key):
            idx=cols.get(key,-1)
            if idx<0 or idx>=df.shape[1]: return [0.0]*len(df)
            return to_num(df.iloc[:,idx])
        dl_cl=gc('dl_cl');dl_cs=gc('dl_cs');dl_net=gc('dl_net')
        am_cl=gc('am_cl');am_cs=gc('am_cs');am_net=gc('am_net')
        lev_cl=gc('lev_cl');lev_cs=gc('lev_cs');lev_net=gc('lev_net')
        oi_col_idx=cols.get('oi',-1)
        if oi_col_idx<0 or oi_col_idx>=df.shape[1]: oi_col_idx=find_oi_col(df)
        oi_ci=oi_col_idx
        oi_all=to_num(df.iloc[:,oi_ci]) if oi_ci<df.shape[1] else [0.0]*len(df)
        all_dates=df['_dt'].dt.strftime('%d.%m.%Y').tolist()
        i0=-1;i1=-2 if len(all_dates)>1 else -1
        def nd(net): return round(float(net[i0])-float(net[i1]),0)
        lev_chg=nd(lev_net);am_chg=nd(am_net);dl_chg=nd(dl_net);oi_chg=nd(oi_all)
        oi_cur=float(oi_all[i0]);oi_prev=float(oi_all[i1])
        oi_pct=round((oi_cur-oi_prev)/abs(oi_prev)*100,2) if oi_prev!=0 else 0.0
        def cot(s): return{'all':calc_cot_index(s),'3y':calc_cot_index(s,156),'1y':calc_cot_index(s,52),'6m':calc_cot_index(s,26),'3m':calc_cot_index(s,13)}
        def stats(net_s,cl_s=None,cs_s=None):
            def mm(s):
                sv=[float(v) for v in s if v!=0];s1y=sv[-52:] if len(sv)>=52 else sv
                return{'max_all':max(sv)if sv else 0,'min_all':min(sv)if sv else 0,'max_1y':max(s1y)if s1y else 0,'min_1y':min(s1y)if s1y else 0}
            r=mm(net_s);r['cl']=mm(cl_s) if cl_s is not None else None;r['cs']=mm(cs_s) if cs_s is not None else None;return r
        N=min(CHART_WEEKS,len(df));cdf=df.tail(N).reset_index(drop=True)
        def gcc(key):
            idx=cols.get(key,-1)
            if idx<0 or idx>=cdf.shape[1]: return [0.0]*N
            return to_num(cdf.iloc[:,idx])
        lev_w=gcc('lev_net');am_w=gcc('am_net');dl_w=gcc('dl_net')
        oi_w=to_num(cdf.iloc[:,oi_ci]) if oi_ci<cdf.shape[1] else [0.0]*N
        # v61: 'oi_d' — тижнева зміна OI, для бару в картці OPEN INTEREST
        chart={'dates':cdf['_dt'].dt.strftime('%d.%m.%Y').tolist(),'lev':lev_w,'am':am_w,'dl':dl_w,'oi':oi_w,'lev_d':compute_delta(lev_w),'am_d':compute_delta(am_w),'dl_d':compute_delta(dl_w),'oi_d':compute_delta(oi_w)}
        hdf=df.tail(HISTORY).reset_index(drop=True)
        def gch(key):
            idx=cols.get(key,-1)
            if idx<0 or idx>=hdf.shape[1]: return [0.0]*len(hdf)
            return to_num(hdf.iloc[:,idx])
        hist={'dates':hdf['_dt'].dt.strftime('%d.%m.%Y').tolist(),'lev_cl':gch('lev_cl'),'lev_cs':gch('lev_cs'),'lev_net':gch('lev_net'),'am_cl':gch('am_cl'),'am_cs':gch('am_cs'),'am_net':gch('am_net'),'dl_cl':gch('dl_cl'),'dl_cs':gch('dl_cs'),'dl_net':gch('dl_net'),'oi':to_num(hdf.iloc[:,oi_ci]) if oi_ci<hdf.shape[1] else [0.0]*len(hdf)}
        oi_st=stats(oi_all)
        oi_cap=round(oi_cur/oi_st['max_all']*100,1) if oi_st.get('max_all',0)>0 else 50.0
        return{'name':name,'sid':sid(name),'table':_tbl_payload(df,TFF_TBL_COLS),'chart':chart,'hist':hist,'stats':{'lev':stats(lev_net,lev_cl,lev_cs),'am':stats(am_net,am_cl,am_cs),'dl':stats(dl_net,dl_cl,dl_cs),'oi':oi_st},'cot_idx':{'lev':cot(lev_net),'am':cot(am_net),'dl':cot(dl_net)},'spark':{'lev':lev_net,'am':am_net,'dl':dl_net,'oi':oi_all},'oi_capacity':oi_cap,'cur':{'date':all_dates[i0],'lev_net':lev_net[i0],'lev_chg':lev_chg,'lev_cl':lev_cl[i0],'lev_cs':lev_cs[i0],'lev_chg_pct':pct_change(lev_chg,lev_net[i0]),'am_net':am_net[i0],'am_chg':am_chg,'am_cl':am_cl[i0],'am_cs':am_cs[i0],'am_chg_pct':pct_change(am_chg,am_net[i0]),'dl_net':dl_net[i0],'dl_chg':dl_chg,'dl_cl':dl_cl[i0],'dl_cs':dl_cs[i0],'dl_chg_pct':pct_change(dl_chg,dl_net[i0]),'oi':oi_cur,'oi_chg':oi_chg,'oi_chg_pct':fp(oi_pct,signed=True)}}
    except Exception as e: print(f"  ❌ TFF {name}: {e}"); return None

def load_tff_data():
    if not TFF_FILE.exists(): print(f"  ⚠  TFF файл не знайдено: {TFF_FILE}"); return {}
    print(f"\n📂  {TFF_FILE}")
    xl=pd.ExcelFile(TFF_FILE)
    sheets=[s for s in xl.sheet_names if s.lower() not in TFF_SKIP_SHEETS]
    print(f"    TFF вкладок: {len(sheets)}\n")
    tff={}
    for sheet_name in sheets:
        mapped_sid=TFF_SHEET_TO_SID.get(sheet_name,sid(sheet_name))
        r=read_tff_sheet(xl,sheet_name)
        if r:
            r['sid']=mapped_sid;tff[mapped_sid]=r
            print(f"  ✓  TFF {sheet_name:20s} → {mapped_sid:12s}  {r['cur']['date']}  AM={fv(r['cur']['am_net'],True,True):>8s}")
        else: print(f"  ✗  TFF {sheet_name}")
    print(f"\n✅  TFF: {len(tff)} інструментів\n")
    return tff

# ================================================================
# DISAGGREGATED — читання даних
# ================================================================
def read_disag_sheet(xl_d, name):
    try:
        raw=xl_d.parse(name, header=None)
        if raw.shape[0]<DISAG_DATA_START+2 or raw.shape[1]<20: return None
        df=raw.iloc[DISAG_DATA_START:].reset_index(drop=True).copy()
        dates_raw=pd.to_datetime(df.iloc[:,DISAG_COL['date']],errors='coerce')
        valid=dates_raw.notna()&(dates_raw.dt.year>2000)
        df=df[valid].copy();df['_dt']=dates_raw[valid].values
        if df.empty: return None
        df=df.sort_values('_dt').reset_index(drop=True)
        def gc(key):
            idx=DISAG_COL.get(key,-1)
            if idx<0 or idx>=df.shape[1]: return [0.0]*len(df)
            return to_num(df.iloc[:,idx])
        mm_cl=gc('mm_cl');mm_cs=gc('mm_cs');mm_net=gc('mm_net')
        pm_cl=gc('pm_cl');pm_cs=gc('pm_cs');pm_net=gc('pm_net')
        sd_cl=gc('sd_cl');sd_cs=gc('sd_cs');sd_net=gc('sd_net')
        oi_idx=DISAG_COL.get('oi',34)
        oi_all=to_num(df.iloc[:,oi_idx]) if oi_idx<df.shape[1] else [0.0]*len(df)
        all_dates=df['_dt'].dt.strftime('%d.%m.%Y').tolist()
        i0=-1;i1=-2 if len(all_dates)>1 else -1
        def nd(net): return round(float(net[i0])-float(net[i1]),0)
        mm_chg=nd(mm_net);pm_chg=nd(pm_net);sd_chg=nd(sd_net);oi_chg=nd(oi_all)
        oi_cur=float(oi_all[i0]);oi_prev=float(oi_all[i1])
        oi_pct=round((oi_cur-oi_prev)/abs(oi_prev)*100,2) if oi_prev!=0 else 0.0
        def cot(s): return{'all':calc_cot_index(s),'3y':calc_cot_index(s,156),'1y':calc_cot_index(s,52),'6m':calc_cot_index(s,26),'3m':calc_cot_index(s,13)}
        def stats(net_s,cl_s=None,cs_s=None):
            def mm(s):
                sv=[float(v) for v in s if v!=0];s1y=sv[-52:] if len(sv)>=52 else sv
                return{'max_all':max(sv)if sv else 0,'min_all':min(sv)if sv else 0,'max_1y':max(s1y)if s1y else 0,'min_1y':min(s1y)if s1y else 0}
            r=mm(net_s);r['cl']=mm(cl_s) if cl_s is not None else None;r['cs']=mm(cs_s) if cs_s is not None else None;return r
        N=min(CHART_WEEKS,len(df));cdf=df.tail(N).reset_index(drop=True)
        def gcc(key):
            idx=DISAG_COL.get(key,-1)
            if idx<0 or idx>=cdf.shape[1]: return [0.0]*N
            return to_num(cdf.iloc[:,idx])
        mm_w=gcc('mm_net');pm_w=gcc('pm_net');sd_w=gcc('sd_net')
        oi_w=to_num(cdf.iloc[:,oi_idx]) if oi_idx<cdf.shape[1] else [0.0]*N
        # v61: 'oi_d' — тижнева зміна OI, для бару в картці OPEN INTEREST
        chart={'dates':cdf['_dt'].dt.strftime('%d.%m.%Y').tolist(),'mm':mm_w,'pm':pm_w,'sd':sd_w,'oi':oi_w,'mm_d':compute_delta(mm_w),'pm_d':compute_delta(pm_w),'sd_d':compute_delta(sd_w),'oi_d':compute_delta(oi_w)}
        hdf=df.tail(HISTORY).reset_index(drop=True)
        def gch(key):
            idx=DISAG_COL.get(key,-1)
            if idx<0 or idx>=hdf.shape[1]: return [0.0]*len(hdf)
            return to_num(hdf.iloc[:,idx])
        hist={'dates':hdf['_dt'].dt.strftime('%d.%m.%Y').tolist(),'mm_cl':gch('mm_cl'),'mm_cs':gch('mm_cs'),'mm_net':gch('mm_net'),'pm_cl':gch('pm_cl'),'pm_cs':gch('pm_cs'),'pm_net':gch('pm_net'),'sd_cl':gch('sd_cl'),'sd_cs':gch('sd_cs'),'sd_net':gch('sd_net'),'oi':to_num(hdf.iloc[:,oi_idx]) if oi_idx<hdf.shape[1] else [0.0]*len(hdf)}
        oi_st=stats(oi_all)
        oi_cap=round(oi_cur/oi_st['max_all']*100,1) if oi_st.get('max_all',0)>0 else 50.0
        return{'name':name,'sid':sid(name),'table':_tbl_payload(df,DISAG_TBL_COLS),'chart':chart,'hist':hist,'stats':{'mm':stats(mm_net,mm_cl,mm_cs),'pm':stats(pm_net,pm_cl,pm_cs),'sd':stats(sd_net,sd_cl,sd_cs),'oi':oi_st},'cot_idx':{'mm':cot(mm_net),'pm':cot(pm_net),'sd':cot(sd_net)},'spark':{'mm':mm_net,'pm':pm_net,'sd':sd_net,'oi':oi_all},'oi_capacity':oi_cap,'cur':{'date':all_dates[i0],'mm_net':mm_net[i0],'mm_chg':mm_chg,'mm_cl':mm_cl[i0],'mm_cs':mm_cs[i0],'mm_chg_pct':pct_change(mm_chg,mm_net[i0]),'pm_net':pm_net[i0],'pm_chg':pm_chg,'pm_cl':pm_cl[i0],'pm_cs':pm_cs[i0],'pm_chg_pct':pct_change(pm_chg,pm_net[i0]),'sd_net':sd_net[i0],'sd_chg':sd_chg,'sd_cl':sd_cl[i0],'sd_cs':sd_cs[i0],'sd_chg_pct':pct_change(sd_chg,sd_net[i0]),'oi':oi_cur,'oi_chg':oi_chg,'oi_chg_pct':fp(oi_pct,signed=True)}}
    except Exception as e: print(f"  ❌ DISAG {name}: {e}"); return None

def load_disag_data():
    if not DISAG_FILE.exists(): print(f"  ⚠  DISAG файл не знайдено: {DISAG_FILE}"); return {}
    print(f"\n📂  {DISAG_FILE}")
    xl=pd.ExcelFile(DISAG_FILE)
    sheets=[s for s in xl.sheet_names if s.lower() not in DISAG_SKIP_SHEETS]
    print(f"    DISAG вкладок: {len(sheets)}\n")
    disag={}
    for sheet_name in sheets:
        mapped_sid=DISAG_SHEET_TO_SID.get(sheet_name,sid(sheet_name))
        r=read_disag_sheet(xl,sheet_name)
        if r:
            r['sid']=mapped_sid;disag[mapped_sid]=r
            print(f"  ✓  DISAG {sheet_name:20s} → {mapped_sid:12s}  {r['cur']['date']}  MM={fv(r['cur']['mm_net'],True,True):>8s}")
        else: print(f"  ✗  DISAG {sheet_name}")
    print(f"\n✅  DISAG: {len(disag)} інструментів\n")
    return disag

# ================================================================
# TFF UI — незмінно з v14
# ================================================================
def make_tff_metric_cards(tff,s):
    c=tff['cur']
    mc_lev=make_metric_card('LEV MONEY (NETTO)',c['lev_net'],c['lev_chg'],c['lev_chg_pct'],tff['spark']['lev'],TFF_COLOR_LEV,gauge_val=tff['cot_idx']['lev']['all'],sub_text="",lbl_color=TFF_COLOR_LEV,bar_id=f'mcbar_tff_lev_{s}',border_color=TFF_COLOR_LEV+'66')
    mc_am =make_metric_card('ASSET MGR (NETTO)',c['am_net'],c['am_chg'],c['am_chg_pct'],tff['spark']['am'],TFF_COLOR_AM,gauge_val=tff['cot_idx']['am']['all'],sub_text="",lbl_color=TFF_COLOR_AM,bar_id=f'mcbar_tff_am_{s}',border_color=TFF_COLOR_AM+'66')
    mc_dl =make_metric_card('DEALER (NETTO)',c['dl_net'],c['dl_chg'],c['dl_chg_pct'],tff['spark']['dl'],TFF_COLOR_DL,gauge_val=tff['cot_idx']['dl']['all'],sub_text="",lbl_color=TFF_COLOR_DL,bar_id=f'mcbar_tff_dl_{s}',border_color=TFF_COLOR_DL+'66')
    mc_oi =make_metric_card('OPEN INTEREST',c['oi'],c['oi_chg'],c['oi_chg_pct'],tff['spark']['oi'],'#a0aac0',oi=True,gauge_val=tff.get('oi_capacity',50.0),sub_text=f"зміна: {fv(int(c['oi_chg']),True,sign=True)}",bar_id=f'mcbar_tff_oi_{s}')
    return f'<div class="mcards">{mc_am}{mc_lev}{mc_dl}{mc_oi}</div>'

def make_tff_analysis(tff):
    # v24 tff analysis — рядки вліво + gauges COT INDEX (ALL / 1Y) справа
    c=tff['cur'];ci=tff['cot_idx']
    def row_html(label,color,net,cl,cs,chg,chg_pct,idx):
        dc='g'if net>0 else'r'
        g_all=idx.get('all',50.0);g_3y=idx.get('3y',50.0);g_1y=idx.get('1y',50.0)
        col_all=gauge_color(g_all);col_3y=gauge_color(g_3y);col_1y=gauge_color(g_1y)
        gauge_all=make_gauge_svg(g_all,col_all,size=62,label='COT ALL')
        gauge_3y =make_gauge_svg(g_3y, col_3y, size=62,label='COT 3Y')
        gauge_1y =make_gauge_svg(g_1y, col_1y, size=62,label='COT 1Y')
        # v60: лівий кольоровий акцент рядка за групою
        return(f'<div class="tff-a-row" style="border-left:2px solid {color}">'
               f'<div class="tff-a-left">'
               f'<div class="tff-a-name" style="color:{color}">{label}</div>'
               f'<div class="tff-a-metrics">'
               f'<div class="tff-ag-item"><span class="ag-lbl">CHG LONG</span>'
               f'<span class="{cc(cl)} ag-val">{fv_full(cl,sign=True)}</span></div>'
               f'<div class="tff-ag-item"><span class="ag-lbl">CHG SHORT</span>'
               f'<span class="{cc(cs)} ag-val">{fv_full(cs,sign=True)}</span></div>'
               f'<div class="tff-ag-item"><span class="ag-lbl">Δ NET</span>'
               f'<span class="{cc(chg)} ag-val">{fv_full(chg,sign=True)}'
               f'<span class="ag-pct"> ({chg_pct})</span></span></div>'
               f'<div class="tff-ag-item"><span class="ag-lbl">NET POS</span>'
               f'<span class="{dc} ag-val ag-bignet">{fv_full(net,sign=True)}</span></div>'
               f'</div></div>'
               f'<div class="tff-a-gauges">'
               f'<div class="tff-a-gwrap">{gauge_all}</div>'
               f'<div class="tff-a-gwrap">{gauge_3y}</div>'
               f'<div class="tff-a-gwrap">{gauge_1y}</div>'
               f'</div></div>')
    return(f'<div class="panel tff-analysis-panel">'
           +row_html('ASSET MGR',TFF_COLOR_AM,c['am_net'],c['am_cl'],c['am_cs'],c['am_chg'],c['am_chg_pct'],ci['am'])
           +row_html('LEV MONEY',TFF_COLOR_LEV,c['lev_net'],c['lev_cl'],c['lev_cs'],c['lev_chg'],c['lev_chg_pct'],ci['lev'])
           +row_html('DEALER',TFF_COLOR_DL,c['dl_net'],c['dl_cl'],c['dl_cs'],c['dl_chg'],c['dl_chg_pct'],ci['dl'])
           +f'</div>')

def make_tff_pct_panel(tff,s):
    ci=tff['cot_idx'];ini=ci['am']['all']
    ini_pos=min(max(ini,0),100)
    ini_color='#f0515a'if ini<15 else('#20d483'if ini>85 else'#dde2ee')
    ini_lbl='екстрем. шорт'if ini<15 else('екстрем. лонг'if ini>85 else'нейтральна зона')
    cj=json.dumps(ci,ensure_ascii=False)
    return(f'<div class="panel pct-panel"><div class="plbl">ПЕРЦЕНТИЛЬ (COT INDEX)</div>'
           f'<div class="pct-sel-row"><div class="psel-group">'
           f'<button class="psel active" data-p="lev" onclick="pctTffSel(this,\'{s}\')">LEV</button>'
           f'<button class="psel" data-p="am"  onclick="pctTffSel(this,\'{s}\')">AM</button>'
           f'<button class="psel" data-p="dl"  onclick="pctTffSel(this,\'{s}\')">DL</button>'
           f'</div><div class="psel-sep"></div><div class="psel-group">'
           f'<button class="pper active" data-per="all" onclick="pperTffSel(this,\'{s}\')">All</button>'
           f'<button class="pper" data-per="3y"  onclick="pperTffSel(this,\'{s}\')">3Y</button>'
           f'<button class="pper" data-per="1y"  onclick="pperTffSel(this,\'{s}\')">1Y</button>'
           f'<button class="pper" data-per="6m"  onclick="pperTffSel(this,\'{s}\')">6M</button>'
           f'<button class="pper" data-per="3m"  onclick="pperTffSel(this,\'{s}\')">3M</button>'
           f'</div></div>'
           f'<div class="pct-val-row"><span id="tff_pctval_{s}" style="font-size:16px;font-weight:bold;color:{ini_color}">{fp(ini)}</span>'
           f'<span id="tff_pctlbl_{s}" style="font-size:11px;color:#8090b0;margin-left:8px;">— {ini_lbl}</span></div>'
           f'<div class="pbar-wrap"><div class="pbar-bg"><div class="pbar-lo"></div><div class="pbar-hi"></div>'
           f'<div class="ptick" style="left:15%"></div><div class="ptick" style="left:85%"></div>'
           f'<div class="pbar-mk" id="tff_pctmk_{s}" style="left:{ini_pos:.1f}%"></div></div>'
           f'<div class="ptick-labels"><span class="ptlbl" style="left:15%">15%</span>'
           f'<span id="tff_pctcur_{s}" class="ptlbl ptlbl-cur" style="left:{ini_pos:.1f}%">{fp(ini)}</span>'
           f'<span class="ptlbl" style="left:85%">85%</span></div></div>'
           f'<div class="pbar-lb"><span>0%</span><span>50%</span><span>100%</span></div></div>'
           f'<script>_ci_tff["{s}"]={cj};</script>')

def make_tff_chart_block(tff,s):
    cj=json.dumps(tff['chart'],ensure_ascii=False)
    return(f'<div class="chartbox"><div class="chartbox-hdr"><div class="plbl" style="margin:0">ЧИСТІ ПОЗИЦІЇ</div>'
           f'<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap"><div class="period-btns">'
           f'<button class="per-btn active" data-per="1y" onclick="setTffChartPer(this,\'{s}\')">1 рік</button>'
           f'<button class="per-btn" data-per="3y" onclick="setTffChartPer(this,\'{s}\')">3 роки</button>'
           f'<button class="per-btn" data-per="5y" onclick="setTffChartPer(this,\'{s}\')">5 років</button>'
           f'</div><div class="chart-leg">'
           f'<span><span class="ll" style="background:{TFF_COLOR_AM}"></span>Asset Mgr</span>'
           f'<span><span class="ll" style="background:{TFF_COLOR_LEV}"></span>Lev Money</span>'
           f'<span class="ll-dash" style="border-top-color:{TFF_COLOR_DL}"></span><span style="margin-left:4px">Dealer</span>'
           f'</div></div></div><div class="cw"><canvas id="tff_cv_{s}"></canvas></div></div>'
           f'<script>_tff["{s}"]={cj};</script>')

def make_tff_bar_block(tff,s):
    return(f'<div class="bar-charts-grid chartbox">'
           f'<div class="bar-wrap"><div class="bar-lbl" style="color:{TFF_COLOR_AM}">ASSET MGR — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="tff_barcv_am_{s}"></canvas></div></div>'
           f'<div class="bar-wrap"><div class="bar-lbl" style="color:{TFF_COLOR_LEV}">LEV MONEY — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="tff_barcv_lev_{s}"></canvas></div></div>'
           f'<div class="bar-wrap"><div class="bar-lbl" style="color:{TFF_COLOR_DL}">DEALER — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="tff_barcv_dl_{s}"></canvas></div></div>'
           f'</div>')

def make_tff_hist_table(tff,table_id):
    hist=tff['hist'];stats=tff['stats'];n=len(hist['dates'])
    if n==0: return '<p style="padding:12px;color:#8090b0">Немає даних</p>'
    def maxabs(lst): vals=[abs(v) for v in lst if v!=0]; return max(vals) if vals else 1
    m={'lev_cl':maxabs(hist['lev_cl']),'lev_cs':maxabs(hist['lev_cs']),'am_cl':maxabs(hist['am_cl']),'am_cs':maxabs(hist['am_cs']),'dl_cl':maxabs(hist['dl_cl']),'dl_cs':maxabs(hist['dl_cs'])}
    def hc(color):
        r,g,b=int(color[1:3],16),int(color[3:5],16),int(color[5:7],16)
        return f'background:rgba({r},{g},{b},.18);color:#fff;border-left:2px solid {color}88'
    def sc(color): return f'background:var(--bg3);color:var(--d);border-left:2px solid {color}55'
    sp='background:var(--bg3);color:var(--d)';oi_bg='background:rgba(160,170,192,.1);color:#fff'
    colgroup='<colgroup><col style="width:84px"><col><col><col><col><col><col><col><col><col><col style="width:82px"></colgroup>'
    thead=(f'<thead><tr class="th-row1"><th class="th-corner"></th>'
           f'<th colspan="3" class="th-group" style="{hc(TFF_COLOR_AM)}">ASSET MGR</th>'
           f'<th colspan="3" class="th-group" style="{hc(TFF_COLOR_LEV)}">LEV MONEY</th>'
           f'<th colspan="3" class="th-group" style="{hc(TFF_COLOR_DL)}">DEALER</th>'
           f'<th class="th-group" style="{oi_bg}">OI</th></tr>'
           f'<tr class="th-row2"><th class="th-date th-left">ДАТА</th>'
           f'<th style="{sc(TFF_COLOR_AM)}">CHG L</th><th style="{sp}">CHG S</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{sc(TFF_COLOR_LEV)}">CHG L</th><th style="{sp}">CHG S</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{sc(TFF_COLOR_DL)}">CHG L</th><th style="{sp}">CHG S</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{oi_bg}" class="th-oi">All</th></tr></thead>')
    def mm_v(v,cls=''): return f'<td class="mm-val {cls}">{fv_full(v,sign=True)}</td>'
    def grp(st,key,col):
        cl_d=st.get('cl') or{};cs_d=st.get('cs') or{}
        return mm_v(cl_d.get(key,0),col)+mm_v(cs_d.get(key,0),col)+mm_v(st.get(key,0),col)
    mm_rows=('<tr class="mm-row"><td class="mm-lbl">MAX</td>'+grp(stats['am'],'max_all','g')+grp(stats['lev'],'max_all','g')+grp(stats['dl'],'max_all','g')+mm_v(stats['oi']['max_all'],'')+'</tr>'
             '<tr class="mm-row"><td class="mm-lbl">MIN</td>'+grp(stats['am'],'min_all','r')+grp(stats['lev'],'min_all','r')+grp(stats['dl'],'min_all','r')+mm_v(stats['oi']['min_all'],'')+'</tr>'
             '<tr class="mm-row mm-yr"><td class="mm-lbl">MAX 1Y</td>'+grp(stats['am'],'max_1y','g')+grp(stats['lev'],'max_1y','g')+grp(stats['dl'],'max_1y','g')+mm_v(stats['oi']['max_1y'],'')+'</tr>'
             '<tr class="mm-row mm-yr"><td class="mm-lbl">MIN 1Y</td>'+grp(stats['am'],'min_1y','r')+grp(stats['lev'],'min_1y','r')+grp(stats['dl'],'min_1y','r')+mm_v(stats['oi']['min_1y'],'')+'</tr>')
    mm_tbody=f'<tbody class="mm-tbody">{mm_rows}</tbody>'
    def td_chg(v,mv): cls=cc(v);txt=fv_full(v,sign=True) if v!=0 else'—'; return f'<td class="{cls}"{intensity_bg(v,mv)}>{txt}</td>'
    def td_net(v,ex=''): return f'<td class="{cc(v)}{ex}">{fv_full(v,sign=True)}</td>'
    rows=[]
    for i in range(n-1,-1,-1):
        ri=n-1-i
        rows.append(f'<tr data-row="{ri}"><td class="date-col">{hist["dates"][i]}</td>'
                    +td_chg(hist['am_cl'][i],m['am_cl'])+td_chg(hist['am_cs'][i],m['am_cs'])+td_net(hist['am_net'][i],' sep-r')
                    +td_chg(hist['lev_cl'][i],m['lev_cl'])+td_chg(hist['lev_cs'][i],m['lev_cs'])+td_net(hist['lev_net'][i],' sep-r')
                    +td_chg(hist['dl_cl'][i],m['dl_cl'])+td_chg(hist['dl_cs'][i],m['dl_cs'])+td_net(hist['dl_net'][i],' sep-r')
                    +f'<td class="t">{fv_full(hist["oi"][i])}</td></tr>')
    data_tbody=f'<tbody class="data-tbody">{"".join(rows)}</tbody>'
    return f'<table class="ht" id="{table_id}">'+colgroup+thead+mm_tbody+data_tbody+'</table>'

# ================================================================
# v60: СЕЗОННІСТЬ (тільки EUR)
# Числа беремо з output/seasonality_test.json (його створює test_seasonality.py)
# і вбудовуємо прямо в HTML. Немає файлу / немає символа → блок не рендериться,
# решта дашборду працює як є.
# ================================================================
MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
_SEASON_CACHE = None

SEASON_DAYS = 365   # довжина денних кривих у presets/current

def _season_curve(v):
    """Валідує денну криву: список рівно на SEASON_DAYS точок (числа або None).
    Усе інше → None (крива просто не потрапляє у графік)."""
    if not (isinstance(v, list) and len(v) == SEASON_DAYS): return None
    out = []
    for x in v:
        if x is None: out.append(None)
        elif isinstance(x, (int, float)) and not (isinstance(x, float) and math.isnan(x)):
            out.append(round(float(x), 4))
        else: out.append(None)
    return out

def _is_preset_key(k):
    """Ключ пресету має вигляд '<N>y' — 'current'/'all_years' сюди не потрапляють."""
    k = str(k)
    return len(k) > 1 and k[-1] == 'y' and k[:-1].isdigit()

def _preset_sort(k):
    return int(str(k)[:-1])

def load_seasonality():
    """Читає SEASON_FILE → {symbol: {...}} з такими ключами:
      'periods'     — {'<N>y': {'avg':[12],'prob':[12],'from':<рік>}} для кожного пресету;
      'default'     — ключ пресету за замовчуванням ('10y', якщо є);
      'months'      — {рік: [12]} по ВСІХ роках (monthly.all_years або об'єднання пресетів);
      'current'     — місячні дані поточного року (monthly.current);
      'pre_daily'   — {'<N>y': [365]} денні криві середніх (блок presets) для лінії «Сер. NY»;
      'cur_daily'   — денна крива поточного року;
      'years_curves'— денні криві по кожному завершеному року.
    Будь-яка проблема (немає файлу, битий JSON, неповні дані) → символ просто
    не потрапляє у словник. Результат кешується на час запуску."""
    global _SEASON_CACHE
    if _SEASON_CACHE is not None: return _SEASON_CACHE
    _SEASON_CACHE = {}
    try:
        raw = json.loads(SEASON_FILE.read_text(encoding='utf-8'))
    except (OSError, ValueError) as e:
        print(f"ℹ️   Сезонність: {SEASON_FILE.name} недоступний ({e.__class__.__name__}) — блок пропущено")
        return _SEASON_CACHE
    for it in (raw if isinstance(raw, list) else [raw]):
        if not isinstance(it, dict): continue
        sym = it.get('symbol'); mon = it.get('monthly')
        if not sym or not isinstance(mon, dict): continue
        # ── місячні пресети: avg/prob + межа вибірки ──
        periods = {}; months = {}
        for pk in sorted((k for k in mon if _is_preset_key(k)), key=_preset_sort):
            blk = mon[pk]
            if not isinstance(blk, dict): continue
            avg = blk.get('avg'); prob = blk.get('prob')
            if not (isinstance(avg,list) and len(avg)==12 and isinstance(prob,list) and len(prob)==12):
                continue
            yrs = blk.get('years') if isinstance(blk.get('years'), dict) else {}
            ylist = sorted((int(y) for y in yrs if str(y).isdigit()))
            if not ylist: continue
            periods[str(pk)] = {'avg': avg, 'prob': prob, 'from': ylist[0]}
            for y, vals in yrs.items():
                if str(y).isdigit() and isinstance(vals, list) and len(vals) == 12:
                    months.setdefault(str(y), vals)
        if not periods: continue
        # всі роки помісячно; якщо блока немає — лишається об'єднання пресетів
        allm = mon.get('all_years')
        if isinstance(allm, dict):
            for y, vals in allm.items():
                if str(y).isdigit() and isinstance(vals, list) and len(vals) == 12:
                    months[str(y)] = vals
        default = '10y' if '10y' in periods else sorted(periods, key=_preset_sort)[0]
        # ── денні криві для графіка ──
        pre_daily = {}
        raw_pre = it.get('presets')
        if isinstance(raw_pre, dict):
            for pk, pv in raw_pre.items():
                if not _is_preset_key(pk): continue
                cur = _season_curve(pv)
                if cur is not None: pre_daily[str(pk)] = cur
        cur_d = it.get('current') if isinstance(it.get('current'), dict) else {}
        cur_daily = None
        cd_series = _season_curve(cur_d.get('data'))
        if cd_series is not None:
            cur_daily = {'year': cur_d.get('year'), 'data': cd_series}
        ycurves = {}
        raw_yc = it.get('years_curves')
        if isinstance(raw_yc, dict):
            for yk, yv in raw_yc.items():
                if not str(yk).isdigit(): continue
                c = _season_curve(yv)
                if c is not None: ycurves[str(yk)] = c
        _SEASON_CACHE[sym] = {
            'periods': periods, 'default': default, 'months': months,
            'current': mon.get('current') if isinstance(mon.get('current'), dict) else {},
            'pre_daily': pre_daily, 'cur_daily': cur_daily, 'years_curves': ycurves,
        }
    if _SEASON_CACHE:
        for sym, v in sorted(_SEASON_CACHE.items()):
            print(f"📅  Сезонність {sym}: пресети {', '.join(sorted(v['periods'], key=_preset_sort))} "
                  f"(типовий {v['default']}), років у таблиці {len(v['months'])}, "
                  f"денних кривих по роках {len(v['years_curves'])}")
    return _SEASON_CACHE

def _sn_ret(v, star=False):
    """Клітинка дохідності: фон/текст за знаком, None → '--'."""
    if v is None or (isinstance(v,float) and math.isnan(v)):
        return '<td class="sn-na">--</td>'
    mark = '*' if star else ''
    if v > 0:   return f'<td class="sn-p">+{v:.2f}{mark}</td>'
    if v < 0:   return f'<td class="sn-n">{v:.2f}{mark}</td>'
    return f'<td class="sn-z">0.00{mark}</td>'

def _sn_prob(v):
    """Клітинка ймовірності: >60% зелена, <40% червона, 40–60% нейтральна."""
    if v is None or (isinstance(v,float) and math.isnan(v)):
        return '<td class="sn-na">--</td>'
    cls = 'sn-p' if v > 60 else ('sn-n' if v < 40 else 'sn-z')
    return f'<td class="{cls}">{round(v):.0f}%</td>'

def _preset_lbl(pk):
    """'10y' → '10Y' для кнопки, '10 РОКІВ' для заголовка."""
    n = _preset_sort(pk)
    return f'{n}Y', f'{n} РОКІВ'

def make_season_box(s, prefix=''):
    """Блок сезонності: [Таблиця|Графік] + перемикач періоду + таблиця + Chart.js-графік.
    prefix ('' | 'tff_' | 'dg_') робить id унікальними між секціями одного інструмента.
    Рядки Probability/Average return, заголовок, мітка межі вибірки і приглушення
    рядків перемикаються на клієнті — дані всіх пресетів вбудовані одразу.
    '' — якщо даних немає."""
    d = load_seasonality().get(s) if s in SEASON_SIDS else None
    if not d: return ''
    key = f'{prefix}{s}'
    periods = d['periods']; dflt = d['default']
    cur = d['current']; cur_year = cur.get('year')
    cur_data = cur.get('data') if isinstance(cur.get('data'), list) else []
    inc = cur.get('incomplete_month')   # номер місяця 1..12
    head = ''.join(f'<th>{m}</th>' for m in MONTHS_SHORT)
    p0 = periods[dflt]
    # sn-stat — два підсумкові рядки; їх вміст переписує JS при зміні періоду
    rows = ['<tr class="sn-stat" id="snprobr_%s"><td class="sn-y">Probability %%</td>%s</tr>'
            % (key, ''.join(_sn_prob(v) for v in p0['prob'])),
            '<tr class="sn-stat" id="snavgr_%s"><td class="sn-y">Average return%%</td>%s</tr>'
            % (key, ''.join(_sn_ret(v) for v in p0['avg'])),
            '<tr class="sn-sep"><td colspan="13"></td></tr>']
    has_star = False
    if cur_year and len(cur_data) == 12:
        cells = []
        for i, v in enumerate(cur_data):
            star = (inc is not None and i == inc-1 and v is not None)
            has_star = has_star or star
            cells.append(_sn_ret(v, star))
        # поточний рік — поза пресетом, показуємо завжди (data-y="0" виводить його
        # з-під фільтра пресету: він не приглушується і не ховається у згорнутому стані)
        rows.append(f'<tr class="sn-cur" data-y="0"><td class="sn-y">{cur_year}</td>'
                    + ''.join(cells) + '</tr>')
    # Початковий стан (типовий пресет, згорнуто) рендеримо одразу в HTML — без спалаху
    # всіх 27 рядків до першого snApply(). Далі стан перемикає тільки JS.
    months = d['months']; y_from = p0['from']
    for y in sorted((int(y) for y in months), reverse=True):
        vals = months[str(y)]
        in_pre = y >= y_from
        cls = ('' if in_pre else ' class="sn-out"') if y != y_from else ' class="sn-brd"'
        hide = '' if in_pre else ' style="display:none"'
        rows.append(f'<tr data-y="{y}"{cls}{hide}><td class="sn-y">{y}'
                    f'<span class="sn-mark">◄ від цього року рахується середнє</span></td>'
                    + ''.join(_sn_ret(v) for v in vals) + '</tr>')
    note = '<div class="sn-note">* місяць незавершений</div>' if has_star else ''
    more = (f'<button class="sn-more" id="snmore_{key}" '
            f'onclick="snToggleAll(this,\'{key}\')">Показати всі роки ▼</button>')
    table_pane = (f'<div class="sn-pane sn-pane-t" id="snpane_tbl_{key}">'
                  f'<table class="sn" id="sntbl_{key}"><thead><tr><th class="sn-y">Year</th>{head}</tr></thead>'
                  f'<tbody>{"".join(rows)}</tbody></table>{note}{more}</div>')
    # Перемикач періоду середнього — окремий рядок під [Таблиця][Графік]
    pbtns = ''.join(
        f'<button class="psel{" active" if pk==dflt else ""}" data-p="{pk}" '
        f'onclick="snSetPeriod(this,\'{key}\',\'{s}\')">{_preset_lbl(pk)[0]}</button>'
        for pk in sorted(periods, key=_preset_sort))
    per_row = (f'<div class="sn-per"><span class="sn-per-lbl">ПЕРІОД:</span>'
               f'<div class="psel-group" id="snper_{key}">{pbtns}</div></div>')
    title = (f'<div class="plbl" id="sntitle_{key}" style="margin:0">'
             f'СЕЗОННІСТЬ ({_preset_lbl(dflt)[1]})</div>')
    # v61: графік малюється ліниво — при першому перемиканні на вкладку «Графік».
    # Рядок кнопок ліній (snlines_) наповнює JS у момент ініціалізації графіка,
    # щоб порядок кнопок гарантовано збігався з порядком datasets.
    pre_daily = d.get('pre_daily') or {}
    cur_daily = d.get('cur_daily')
    if not (pre_daily or cur_daily):
        return f'<div class="sn-hdr">{title}</div>{per_row}{table_pane}'
    chart_pane = (f'<div class="sn-pane sn-chart" id="snpane_chart_{key}" style="display:none">'
                  f'<div class="sn-lines" id="snlines_{key}"></div>'
                  f'<div class="sn-cw"><canvas id="sncv_{key}"></canvas></div></div>')
    sel = (f'<div class="sn-sel psel-group">'
           f'<button class="psel active" data-sv="tbl" onclick="seasonView(this,\'{key}\',\'{s}\',\'tbl\')">Таблиця</button>'
           f'<button class="psel" data-sv="chart" onclick="seasonView(this,\'{key}\',\'{s}\',\'chart\')">Графік</button>'
           f'</div>')
    # Дані вбудовуємо один раз — у Legacy-секції (вона рендериться завжди),
    # TFF/DISAG беруть той самий _season[sid]. Патерн той самий, що у _cd / _ci.
    data_js = ''
    if prefix == '':
        payload = {'per': {pk: {'prob': v['prob'], 'avg': v['avg'], 'from': v['from']}
                           for pk, v in periods.items()},
                   'dflt': dflt,
                   'pre': pre_daily,
                   'cur_year': (cur_daily or {}).get('year'),
                   'cur': (cur_daily or {}).get('data'),
                   'years': d.get('years_curves') or {}}
        data_js = f'<script>_season["{s}"]={json.dumps(payload,ensure_ascii=False)};</script>'
    return (f'<div class="sn-hdr">{title}{sel}</div>{per_row}'
            f'{table_pane}{chart_pane}{data_js}')

def make_tv_col(s,prefix=''):
    """v58: блок TradingView зверху секції — віджет або заглушка.
    prefix: '' | 'tff_' | 'dg_' — щоб id контейнерів не конфліктували між секціями.
    Заглушка не має data-tvsym, тому лінива ініціалізація її не чіпає.
    v60: якщо для інструмента є сезонність — верх ділиться навпіл
    (таблиця сезонності | TradingView), інакше графік лишається на всю ширину."""
    box_id=f'tv_{prefix}{s}'
    tv_sym=TV_SYMBOL.get(s)
    if tv_sym:
        inner=f'<div class="tv-box" id="{box_id}" data-tvsym="{tv_sym}"></div>'
    else:
        inner=(f'<div class="tv-box tv-empty" id="{box_id}"><div class="tv-empty-msg">'
               '<div class="tv-empty-ico">📉</div>'
               '<div>Графік недоступний</div>'
               '<div class="tv-empty-sub">Немає відкритого символу TradingView '
               'для цього інструмента</div>'
               f'<button class="tv-pick-btn" onclick="tvPick(\'{box_id}\')">Вибрати символ</button>'
               '</div></div>')
    season=make_season_box(s,prefix)
    if season:
        return (f'<div class="top-split"><div class="season-box">{season}</div>'
                f'<div class="tv-half">{inner}</div></div>')
    return f'<div class="top-split one-col"><div class="tv-half">{inner}</div></div>'

def make_tff_view(tff,s,reports_panel_html):
    cards=make_tff_metric_cards(tff,s);analysis=make_tff_analysis(tff)
    pct=make_tff_pct_panel(tff,s);chart=make_tff_chart_block(tff,s)
    bars=make_tff_bar_block(tff,s);tbl_id=f'tff_tbl_{s}';tbl=make_tff_hist_table(tff,tbl_id)
    # v52: верстка вкладки TABLE, обрізана по OPEN INTEREST
    table_block=(f'<div class="htable-wrap"><div class="htable-hdr"><span>ТИЖНЕВА СТАТИСТИКА ПОЗИЦІЙ</span>'
                 f'<div class="hsel">'
                 f'<button class="hbtn active" data-n="10" onclick="setMiniHistT(this,\'{s}\')">10</button>'
                 f'<button class="hbtn" data-n="26" onclick="setMiniHistT(this,\'{s}\')">26</button>'
                 f'<button class="hbtn" data-n="52" onclick="setMiniHistT(this,\'{s}\')">52</button>'
                 f'</div></div><div class="tb-scroll tb-mini" id="mini_tff_{s}"></div></div>')
    # v58: та сама структура, що й у Legacy — TradingView на всю ширину зверху,
    # під ним картки | analysis+перцентиль. Bars, chart і таблиця — під сіткою.
    lg_top=(make_tv_col(s,'tff_')
            +f'<div class="lg-grid">'
             f'<div class="lg-cards">{cards}</div>'
             f'<div class="lg-side">{analysis}{pct}</div>'
             f'</div>')
    return(f'<div class="rpt-sec" id="rpt_tff_{s}" style="display:none">'+lg_top+bars+chart+table_block+'</div>')

# ================================================================
# DISAGGREGATED UI — аналогічно TFF, різні назви учасників
# ================================================================
def make_disag_metric_cards(dg,s):
    c=dg['cur']
    mc_mm=make_metric_card('MAN MONEY (NETTO)',c['mm_net'],c['mm_chg'],c['mm_chg_pct'],dg['spark']['mm'],DISAG_COLOR_MM,gauge_val=dg['cot_idx']['mm']['all'],sub_text="",lbl_color=DISAG_COLOR_MM,bar_id=f'mcbar_dg_mm_{s}',border_color=DISAG_COLOR_MM+'66')
    mc_pm=make_metric_card('PROD/MERCH (NETTO)',c['pm_net'],c['pm_chg'],c['pm_chg_pct'],dg['spark']['pm'],DISAG_COLOR_PM,gauge_val=dg['cot_idx']['pm']['all'],sub_text="",lbl_color=DISAG_COLOR_PM,bar_id=f'mcbar_dg_pm_{s}',border_color=DISAG_COLOR_PM+'66')
    mc_sd=make_metric_card('SWAP DEALERS (NETTO)',c['sd_net'],c['sd_chg'],c['sd_chg_pct'],dg['spark']['sd'],DISAG_COLOR_SD,gauge_val=dg['cot_idx']['sd']['all'],sub_text="",lbl_color=DISAG_COLOR_SD,bar_id=f'mcbar_dg_sd_{s}',border_color=DISAG_COLOR_SD+'66')
    mc_oi=make_metric_card('OPEN INTEREST',c['oi'],c['oi_chg'],c['oi_chg_pct'],dg['spark']['oi'],'#a0aac0',oi=True,gauge_val=dg.get('oi_capacity',50.0),sub_text=f"зміна: {fv(int(c['oi_chg']),True,sign=True)}",bar_id=f'mcbar_dg_oi_{s}')
    return f'<div class="mcards">{mc_mm}{mc_pm}{mc_sd}{mc_oi}</div>'

def make_disag_analysis(dg):
    # v24 disag analysis — рядки вліво + gauges COT INDEX (ALL/1Y) справа
    c=dg['cur'];ci=dg['cot_idx']
    def row_html(label,color,net,cl,cs,chg,chg_pct,idx):
        dc='g'if net>0 else'r'
        g_all=idx.get('all',50.0);g_3y=idx.get('3y',50.0);g_1y=idx.get('1y',50.0)
        col_all=gauge_color(g_all);col_3y=gauge_color(g_3y);col_1y=gauge_color(g_1y)
        gauge_all=make_gauge_svg(g_all,col_all,size=62,label='COT ALL')
        gauge_3y =make_gauge_svg(g_3y, col_3y, size=62,label='COT 3Y')
        gauge_1y =make_gauge_svg(g_1y, col_1y, size=62,label='COT 1Y')
        # v60: лівий кольоровий акцент рядка за групою
        return(f'<div class="tff-a-row" style="border-left:2px solid {color}">'
               f'<div class="tff-a-left">'
               f'<div class="tff-a-name" style="color:{color}">{label}</div>'
               f'<div class="tff-a-metrics">'
               f'<div class="tff-ag-item"><span class="ag-lbl">CHG LONG</span>'
               f'<span class="{cc(cl)} ag-val">{fv_full(cl,sign=True)}</span></div>'
               f'<div class="tff-ag-item"><span class="ag-lbl">CHG SHORT</span>'
               f'<span class="{cc(cs)} ag-val">{fv_full(cs,sign=True)}</span></div>'
               f'<div class="tff-ag-item"><span class="ag-lbl">Δ NET</span>'
               f'<span class="{cc(chg)} ag-val">{fv_full(chg,sign=True)}'
               f'<span class="ag-pct"> ({chg_pct})</span></span></div>'
               f'<div class="tff-ag-item"><span class="ag-lbl">NET POS</span>'
               f'<span class="{dc} ag-val ag-bignet">{fv_full(net,sign=True)}</span></div>'
               f'</div></div>'
               f'<div class="tff-a-gauges">'
               f'<div class="tff-a-gwrap">{gauge_all}</div>'
               f'<div class="tff-a-gwrap">{gauge_3y}</div>'
               f'<div class="tff-a-gwrap">{gauge_1y}</div>'
               f'</div></div>')
    return(f'<div class="panel tff-analysis-panel">'
           +row_html('MAN MONEY',DISAG_COLOR_MM,c['mm_net'],c['mm_cl'],c['mm_cs'],c['mm_chg'],c['mm_chg_pct'],ci['mm'])
           +row_html('PROD/MERCH',DISAG_COLOR_PM,c['pm_net'],c['pm_cl'],c['pm_cs'],c['pm_chg'],c['pm_chg_pct'],ci['pm'])
           +row_html('SWAP DEALERS',DISAG_COLOR_SD,c['sd_net'],c['sd_cl'],c['sd_cs'],c['sd_chg'],c['sd_chg_pct'],ci['sd'])
           +f'</div>')

def make_disag_pct_panel(dg,s):
    ci=dg['cot_idx'];ini=ci['mm']['all']
    ini_pos=min(max(ini,0),100)
    ini_color='#f0515a'if ini<15 else('#20d483'if ini>85 else'#dde2ee')
    ini_lbl='екстрем. шорт'if ini<15 else('екстрем. лонг'if ini>85 else'нейтральна зона')
    cj=json.dumps(ci,ensure_ascii=False)
    return(f'<div class="panel pct-panel"><div class="plbl">ПЕРЦЕНТИЛЬ (COT INDEX)</div>'
           f'<div class="pct-sel-row"><div class="psel-group">'
           f'<button class="psel active" data-p="mm" onclick="pctDgSel(this,\'{s}\')">MM</button>'
           f'<button class="psel" data-p="pm"  onclick="pctDgSel(this,\'{s}\')">PM</button>'
           f'<button class="psel" data-p="sd"  onclick="pctDgSel(this,\'{s}\')">SD</button>'
           f'</div><div class="psel-sep"></div><div class="psel-group">'
           f'<button class="pper active" data-per="all" onclick="pperDgSel(this,\'{s}\')">All</button>'
           f'<button class="pper" data-per="3y"  onclick="pperDgSel(this,\'{s}\')">3Y</button>'
           f'<button class="pper" data-per="1y"  onclick="pperDgSel(this,\'{s}\')">1Y</button>'
           f'<button class="pper" data-per="6m"  onclick="pperDgSel(this,\'{s}\')">6M</button>'
           f'<button class="pper" data-per="3m"  onclick="pperDgSel(this,\'{s}\')">3M</button>'
           f'</div></div>'
           f'<div class="pct-val-row"><span id="dg_pctval_{s}" style="font-size:16px;font-weight:bold;color:{ini_color}">{fp(ini)}</span>'
           f'<span id="dg_pctlbl_{s}" style="font-size:11px;color:#8090b0;margin-left:8px;">— {ini_lbl}</span></div>'
           f'<div class="pbar-wrap"><div class="pbar-bg"><div class="pbar-lo"></div><div class="pbar-hi"></div>'
           f'<div class="ptick" style="left:15%"></div><div class="ptick" style="left:85%"></div>'
           f'<div class="pbar-mk" id="dg_pctmk_{s}" style="left:{ini_pos:.1f}%"></div></div>'
           f'<div class="ptick-labels"><span class="ptlbl" style="left:15%">15%</span>'
           f'<span id="dg_pctcur_{s}" class="ptlbl ptlbl-cur" style="left:{ini_pos:.1f}%">{fp(ini)}</span>'
           f'<span class="ptlbl" style="left:85%">85%</span></div></div>'
           f'<div class="pbar-lb"><span>0%</span><span>50%</span><span>100%</span></div></div>'
           f'<script>_ci_dg["{s}"]={cj};</script>')

def make_disag_chart_block(dg,s):
    cj=json.dumps(dg['chart'],ensure_ascii=False)
    return(f'<div class="chartbox"><div class="chartbox-hdr"><div class="plbl" style="margin:0">ЧИСТІ ПОЗИЦІЇ</div>'
           f'<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap"><div class="period-btns">'
           f'<button class="per-btn active" data-per="1y" onclick="setDgChartPer(this,\'{s}\')">1 рік</button>'
           f'<button class="per-btn" data-per="3y" onclick="setDgChartPer(this,\'{s}\')">3 роки</button>'
           f'<button class="per-btn" data-per="5y" onclick="setDgChartPer(this,\'{s}\')">5 років</button>'
           f'</div><div class="chart-leg">'
           f'<span><span class="ll" style="background:{DISAG_COLOR_MM}"></span>Man Money</span>'
           f'<span><span class="ll" style="background:{DISAG_COLOR_PM}"></span>Prod/Merch</span>'
           f'<span class="ll-dash" style="border-top-color:{DISAG_COLOR_SD}"></span><span style="margin-left:4px">Swap Dealers</span>'
           f'</div></div></div><div class="cw"><canvas id="dg_cv_{s}"></canvas></div></div>'
           f'<script>_dg["{s}"]={cj};</script>')

def make_disag_bar_block(dg,s):
    return(f'<div class="bar-charts-grid chartbox">'
           f'<div class="bar-wrap"><div class="bar-lbl" style="color:{DISAG_COLOR_MM}">MAN MONEY — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="dg_barcv_mm_{s}"></canvas></div></div>'
           f'<div class="bar-wrap"><div class="bar-lbl" style="color:{DISAG_COLOR_PM}">PROD/MERCH — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="dg_barcv_pm_{s}"></canvas></div></div>'
           f'<div class="bar-wrap"><div class="bar-lbl" style="color:{DISAG_COLOR_SD}">SWAP DEALERS — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="dg_barcv_sd_{s}"></canvas></div></div>'
           f'</div>')

def make_disag_hist_table(dg,table_id):
    hist=dg['hist'];stats=dg['stats'];n=len(hist['dates'])
    if n==0: return '<p style="padding:12px;color:#8090b0">Немає даних</p>'
    def maxabs(lst): vals=[abs(v) for v in lst if v!=0]; return max(vals) if vals else 1
    m={'mm_cl':maxabs(hist['mm_cl']),'mm_cs':maxabs(hist['mm_cs']),'pm_cl':maxabs(hist['pm_cl']),'pm_cs':maxabs(hist['pm_cs']),'sd_cl':maxabs(hist['sd_cl']),'sd_cs':maxabs(hist['sd_cs'])}
    def hc(color):
        r,g,b=int(color[1:3],16),int(color[3:5],16),int(color[5:7],16)
        return f'background:rgba({r},{g},{b},.18);color:#fff;border-left:2px solid {color}88'
    def sc(color): return f'background:var(--bg3);color:var(--d);border-left:2px solid {color}55'
    sp='background:var(--bg3);color:var(--d)';oi_bg='background:rgba(160,170,192,.1);color:#fff'
    colgroup='<colgroup><col style="width:84px"><col><col><col><col><col><col><col><col><col><col style="width:82px"></colgroup>'
    thead=(f'<thead><tr class="th-row1"><th class="th-corner"></th>'
           f'<th colspan="3" class="th-group" style="{hc(DISAG_COLOR_MM)}">MAN MONEY</th>'
           f'<th colspan="3" class="th-group" style="{hc(DISAG_COLOR_PM)}">PROD/MERCH</th>'
           f'<th colspan="3" class="th-group" style="{hc(DISAG_COLOR_SD)}">SWAP DEALERS</th>'
           f'<th class="th-group" style="{oi_bg}">OI</th></tr>'
           f'<tr class="th-row2"><th class="th-date th-left">ДАТА</th>'
           f'<th style="{sc(DISAG_COLOR_MM)}">CHG L</th><th style="{sp}">CHG S</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{sc(DISAG_COLOR_PM)}">CHG L</th><th style="{sp}">CHG S</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{sc(DISAG_COLOR_SD)}">CHG L</th><th style="{sp}">CHG S</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{oi_bg}" class="th-oi">All</th></tr></thead>')
    def mm_v(v,cls=''): return f'<td class="mm-val {cls}">{fv_full(v,sign=True)}</td>'
    def grp(st,key,col):
        cl_d=st.get('cl') or{};cs_d=st.get('cs') or{}
        return mm_v(cl_d.get(key,0),col)+mm_v(cs_d.get(key,0),col)+mm_v(st.get(key,0),col)
    mm_rows=('<tr class="mm-row"><td class="mm-lbl">MAX</td>'+grp(stats['mm'],'max_all','g')+grp(stats['pm'],'max_all','g')+grp(stats['sd'],'max_all','g')+mm_v(stats['oi']['max_all'],'')+'</tr>'
             '<tr class="mm-row"><td class="mm-lbl">MIN</td>'+grp(stats['mm'],'min_all','r')+grp(stats['pm'],'min_all','r')+grp(stats['sd'],'min_all','r')+mm_v(stats['oi']['min_all'],'')+'</tr>'
             '<tr class="mm-row mm-yr"><td class="mm-lbl">MAX 1Y</td>'+grp(stats['mm'],'max_1y','g')+grp(stats['pm'],'max_1y','g')+grp(stats['sd'],'max_1y','g')+mm_v(stats['oi']['max_1y'],'')+'</tr>'
             '<tr class="mm-row mm-yr"><td class="mm-lbl">MIN 1Y</td>'+grp(stats['mm'],'min_1y','r')+grp(stats['pm'],'min_1y','r')+grp(stats['sd'],'min_1y','r')+mm_v(stats['oi']['min_1y'],'')+'</tr>')
    mm_tbody=f'<tbody class="mm-tbody">{mm_rows}</tbody>'
    def td_chg(v,mv): cls=cc(v);txt=fv_full(v,sign=True) if v!=0 else'—'; return f'<td class="{cls}"{intensity_bg(v,mv)}>{txt}</td>'
    def td_net(v,ex=''): return f'<td class="{cc(v)}{ex}">{fv_full(v,sign=True)}</td>'
    rows=[]
    for i in range(n-1,-1,-1):
        ri=n-1-i
        rows.append(f'<tr data-row="{ri}"><td class="date-col">{hist["dates"][i]}</td>'
                    +td_chg(hist['mm_cl'][i],m['mm_cl'])+td_chg(hist['mm_cs'][i],m['mm_cs'])+td_net(hist['mm_net'][i],' sep-r')
                    +td_chg(hist['pm_cl'][i],m['pm_cl'])+td_chg(hist['pm_cs'][i],m['pm_cs'])+td_net(hist['pm_net'][i],' sep-r')
                    +td_chg(hist['sd_cl'][i],m['sd_cl'])+td_chg(hist['sd_cs'][i],m['sd_cs'])+td_net(hist['sd_net'][i],' sep-r')
                    +f'<td class="t">{fv_full(hist["oi"][i])}</td></tr>')
    data_tbody=f'<tbody class="data-tbody">{"".join(rows)}</tbody>'
    return f'<table class="ht" id="{table_id}">'+colgroup+thead+mm_tbody+data_tbody+'</table>'

def make_disag_view(dg,s,reports_panel_html):
    cards=make_disag_metric_cards(dg,s);analysis=make_disag_analysis(dg)
    pct=make_disag_pct_panel(dg,s);chart=make_disag_chart_block(dg,s)
    bars=make_disag_bar_block(dg,s);tbl_id=f'dg_tbl_{s}';tbl=make_disag_hist_table(dg,tbl_id)
    # v53: верстка вкладки TABLE, обрізана по OPEN INTEREST
    table_block=(f'<div class="htable-wrap"><div class="htable-hdr"><span>ТИЖНЕВА СТАТИСТИКА ПОЗИЦІЙ</span>'
                 f'<div class="hsel">'
                 f'<button class="hbtn active" data-n="10" onclick="setMiniHistG(this,\'{s}\')">10</button>'
                 f'<button class="hbtn" data-n="26" onclick="setMiniHistG(this,\'{s}\')">26</button>'
                 f'<button class="hbtn" data-n="52" onclick="setMiniHistG(this,\'{s}\')">52</button>'
                 f'</div></div><div class="tb-scroll tb-mini" id="mini_dg_{s}"></div></div>')
    # v58: аналогічно TFF — TradingView на всю ширину зверху, під ним cards | analysis+pct
    lg_top=(make_tv_col(s,'dg_')
            +f'<div class="lg-grid">'
             f'<div class="lg-cards">{cards}</div>'
             f'<div class="lg-side">{analysis}{pct}</div>'
             f'</div>')
    return(f'<div class="rpt-sec" id="rpt_dg_{s}" style="display:none">'+lg_top+bars+chart+table_block+'</div>')

# ================================================================
# LEGACY таблиця (незмінно з v14)
# ================================================================
def make_hist_table(hist,stats_ls,stats_cm,stats_st,stats_oi,sm):
    n=len(hist['dates'])
    if n==0: return '<p style="padding:12px;color:#8090b0">Немає даних</p>'
    def maxabs(lst): vals=[abs(v) for v in lst if v!=0]; return max(vals) if vals else 1
    m_ls_cl=maxabs(hist['ls_cl']);m_ls_cs=maxabs(hist['ls_cs'])
    m_cm_cl=maxabs(hist['cm_cl']);m_cm_cs=maxabs(hist['cm_cs'])
    m_st_cl=maxabs(hist['st_cl']);m_st_cs=maxabs(hist['st_cs'])
    ls_bg='background:rgba(74,158,255,.18);color:#fff;border-left:2px solid rgba(74,158,255,.7)'
    cm_bg='background:rgba(32,212,131,.15);color:#fff;border-left:2px solid rgba(32,212,131,.7)'
    st_bg='background:rgba(240,81,90,.15);color:#fff;border-left:2px solid rgba(240,81,90,.7)'
    oi_bg='background:rgba(160,170,192,.1);color:#fff';sm_bg='background:var(--bg3);color:var(--d)'
    sp='background:var(--bg3);color:var(--d)'
    sub_ls=f'{sp};border-left:2px solid rgba(74,158,255,.5)'
    sub_cm=f'{sp};border-left:2px solid rgba(32,212,131,.5)'
    sub_st=f'{sp};border-left:2px solid rgba(240,81,90,.5)'
    _grpcols='<col><col><col style="width:56px"><col style="width:56px"><col>'
    colgroup='<colgroup><col style="width:84px">'+_grpcols*3+'<col style="width:82px"><col style="width:50px"><col style="width:50px"><col style="width:50px"></colgroup>'
    thead=(f'<thead><tr class="th-row1"><th class="th-corner"></th>'
           f'<th colspan="5" class="th-group" style="{ls_bg}">LARGE SPECULATORS</th>'
           f'<th colspan="5" class="th-group" style="{cm_bg}">COMMERCIALS</th>'
           f'<th colspan="5" class="th-group" style="{st_bg}">SMALL TRADERS</th>'
           f'<th class="th-group" style="{oi_bg}">OI</th>'
           f'<th colspan="3" class="th-group sm-th-group" style="{sm_bg}">SM DIV</th></tr>'
           f'<tr class="th-row2"><th class="th-date th-left">ДАТА</th>'
           f'<th style="{sub_ls}">CHG L</th><th style="{sp}">CHG S</th><th style="{sp}">%N/OI</th><th style="{sp}">%OIΔ</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{sub_cm}">CHG L</th><th style="{sp}">CHG S</th><th style="{sp}">%N/OI</th><th style="{sp}">%OIΔ</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{sub_st}">CHG L</th><th style="{sp}">CHG S</th><th style="{sp}">%N/OI</th><th style="{sp}">%OIΔ</th><th class="sep-r" style="{sp}">NET POS</th>'
           f'<th style="{oi_bg}" class="th-oi">All</th>'
           f'<th class="sm-th">All</th><th class="sm-th">6M</th><th class="sm-th">3M</th></tr></thead>')
    def mm_v(v,cls=''): return f'<td class="mm-val {cls}">{fv_full(v,sign=True)}</td>'
    def sm_mini(v):
        if v is None: return '<td class="sm-td d">–</td>'
        try: f2=float(v);c='g'if f2>0 else('r'if f2<0 else'd'); return f'<td class="sm-td {c}">{f2:+.2f}</td>'
        except: return '<td class="sm-td d">–</td>'
    def grp(st,key,col):
        cl_d=st.get('cl') or{};cs_d=st.get('cs') or{}
        return mm_v(cl_d.get(key,0),col)+mm_v(cs_d.get(key,0),col)+'<td class="mm-val"></td><td class="mm-val"></td>'+mm_v(st.get(key,0),col)
    mm_rows=('<tr class="mm-row"><td class="mm-lbl">MAX</td>'+grp(stats_ls,'max_all','g')+grp(stats_cm,'max_all','g')+grp(stats_st,'max_all','g')+mm_v(stats_oi['max_all'],'')+sm_mini(sm['div'])+sm_mini(sm['div_6m'])+sm_mini(sm['div_3m'])+'</tr>'
             '<tr class="mm-row"><td class="mm-lbl">MIN</td>'+grp(stats_ls,'min_all','r')+grp(stats_cm,'min_all','r')+grp(stats_st,'min_all','r')+mm_v(stats_oi['min_all'],'')+'<td class="sm-td"></td><td class="sm-td"></td><td class="sm-td"></td></tr>'
             '<tr class="mm-row mm-yr"><td class="mm-lbl">MAX 1Y</td>'+grp(stats_ls,'max_1y','g')+grp(stats_cm,'max_1y','g')+grp(stats_st,'max_1y','g')+mm_v(stats_oi['max_1y'],'')+'<td class="sm-td"></td><td class="sm-td"></td><td class="sm-td"></td></tr>'
             '<tr class="mm-row mm-yr"><td class="mm-lbl">MIN 1Y</td>'+grp(stats_ls,'min_1y','r')+grp(stats_cm,'min_1y','r')+grp(stats_st,'min_1y','r')+mm_v(stats_oi['min_1y'],'')+'<td class="sm-td"></td><td class="sm-td"></td><td class="sm-td"></td></tr>')
    mm_tbody=f'<tbody class="mm-tbody">{mm_rows}</tbody>'
    def td_chg(v,mv): cls=cc(v);txt=fv_full(v,sign=True) if v!=0 else'—'; return f'<td class="{cls}"{intensity_bg(v,mv)}>{txt}</td>'
    def td_net(v,ex=''): return f'<td class="{cc(v)}{ex}">{fv_full(v,sign=True)}</td>'
    def sm_cell(v):
        if v is None or(isinstance(v,float) and v!=v): return '<td class="sm-td d"></td>'
        try:
            f2=float(v)
            if f2==0: return '<td class="sm-td d"></td>'
            c='g'if f2>0 else'r'; return f'<td class="sm-td {c}">{f2:+.3f}</td>'
        except: return '<td class="sm-td d"></td>'
    def td_pct(v):
        # %-клітинка: аномалії (ділення на ~0 у файлі) ховаємо
        try: f2=float(v)
        except: return '<td class="pctc d">—</td>'
        if f2!=f2 or abs(f2)>999: return '<td class="pctc d">—</td>'
        _st=''
        if f2>30:    _st=' style="background:rgba(32,212,131,.30)"'
        elif f2<-30: _st=' style="background:rgba(240,81,90,.30)"'
        return f'<td class="pctc {cc(f2)}"{_st}>{f2:+.0f}%</td>'
    lsp=hist.get('ls_pct_row',[None]*n);lso=hist.get('ls_oich_row',[None]*n)
    cmp_=hist.get('cm_pct_row',[None]*n);cmo=hist.get('cm_oich_row',[None]*n)
    stp=hist.get('st_pct_row',[None]*n);sto=hist.get('st_oich_row',[None]*n)
    sv=hist.get('sm_div_row',[None]*n);s6=hist.get('sm_div_6m_row',[None]*n);s3=hist.get('sm_div_3m_row',[None]*n)
    rows=[]
    for i in range(n-1,-1,-1):
        ri=n-1-i
        rows.append(f'<tr data-row="{ri}"><td class="date-col">{hist["dates"][i]}</td>'
                    +td_chg(hist['ls_cl'][i],m_ls_cl)+td_chg(hist['ls_cs'][i],m_ls_cs)+td_pct(lsp[i])+td_pct(lso[i])+td_net(hist['ls_net'][i],' sep-r')
                    +td_chg(hist['cm_cl'][i],m_cm_cl)+td_chg(hist['cm_cs'][i],m_cm_cs)+td_pct(cmp_[i])+td_pct(cmo[i])+td_net(hist['cm_net'][i],' sep-r')
                    +td_chg(hist['st_cl'][i],m_st_cl)+td_chg(hist['st_cs'][i],m_st_cs)+td_pct(stp[i])+td_pct(sto[i])+td_net(hist['st_net'][i],' sep-r')
                    +f'<td class="t">{fv_full(hist["oi"][i])}</td>'
                    +sm_cell(sv[i])+sm_cell(s6[i])+sm_cell(s3[i])+'</tr>')
    data_tbody=f'<tbody class="data-tbody">{"".join(rows)}</tbody>'
    return '<table class="ht ht-legacy">'+colgroup+thead+mm_tbody+data_tbody+'</table>'

# ================================================================
# INSTRUMENT VIEW
# ================================================================
def _make_ranked_section(s, cot_m):
    """Вбудована секція COT INDEX RANKED (M) всередині pct-panel"""
    if not cot_m: return ''
    # Якщо жодного валідного значення немає у всіх групах — не показуємо панель
    _has_any = any(
        (cot_m.get(g,{}) or {}).get(p) is not None
        for g in ('ls','cm','st')
        for p in ('all','3y','1y','6m','3m')
    )
    if not _has_any: return ''
    ini = cot_m.get('ls',{}).get('all') or 50.0
    ini_pos=min(max(ini,0),100)
    ini_color='#f0515a'if ini<15 else('#20d483'if ini>85 else'#dde2ee')
    ini_lbl='екстрем. шорт'if ini<15 else('екстрем. лонг'if ini>85 else'нейтральна зона')
    return(f'<div class="panel pct-panel" id="pctm_panel_{s}">'
           f'<div class="plbl" style="margin-bottom:6px">COT INDEX RANKED (M)</div>'
           f'<div class="pct-sel-row"><div class="psel-group">'
           f'<button class="psm active" data-p="ls" onclick="pctMSel(this,\'{s}\')">LS</button>'
           f'<button class="psm" data-p="cm" onclick="pctMSel(this,\'{s}\')">CM</button>'
           f'<button class="psm" data-p="st" onclick="pctMSel(this,\'{s}\')">ST</button>'
           f'</div><div class="psel-sep"></div><div class="psel-group">'
           f'<button class="ppm active" data-per="all" onclick="pperMSel(this,\'{s}\')">All</button>'
           f'<button class="ppm" data-per="3y"  onclick="pperMSel(this,\'{s}\')">3Y</button>'
           f'<button class="ppm" data-per="1y"  onclick="pperMSel(this,\'{s}\')">1Y</button>'
           f'<button class="ppm" data-per="6m"  onclick="pperMSel(this,\'{s}\')">6M</button>'
           f'<button class="ppm" data-per="3m"  onclick="pperMSel(this,\'{s}\')">3M</button>'
           f'</div></div>'
           f'<div class="pct-val-row">'
           f'<span id="pctmval_{s}" style="font-size:16px;font-weight:bold;color:{ini_color}">{ini:.1f}%</span>'
           f'<span id="pctmcls_{s}" style="font-size:11px;color:#8090b0;margin-left:8px;">— {ini_lbl}</span>'
           f'</div>'
           f'<div class="pbar-wrap"><div class="pbar-bg"><div class="pbar-lo"></div><div class="pbar-hi"></div>'
           f'<div class="ptick" style="left:15%"></div><div class="ptick" style="left:85%"></div>'
           f'<div class="pbar-mk" id="pctmmk_{s}" style="left:{ini_pos:.1f}%;background:#e8a838"></div></div>'
           f'<div class="ptick-labels"><span class="ptlbl" style="left:15%">15%</span>'
           f'<span id="pctmcur_{s}" class="ptlbl ptlbl-cur" style="left:{ini_pos:.1f}%">{ini:.1f}%</span>'
           f'<span class="ptlbl" style="left:85%">85%</span></div></div>'
           f'<div class="pbar-lb"><span>0%</span><span>50%</span><span>100%</span></div></div>')

def _make_ranked_block(s, cot_m, default_group='ls'):
    """Блок COT INDEX Ranked (M) під перцентилем"""
    if not cot_m or all(v is None for v in cot_m.get(default_group,{}).values()):
        return ''
    ini_group = default_group
    ini = cot_m.get(ini_group,{}).get('all') or 50.0
    ini_pos=min(max(ini,0),100)
    ini_color='#f0515a'if ini<15 else('#20d483'if ini>85 else'#dde2ee')
    ini_lbl='екстрем. шорт'if ini<15 else('екстрем. лонг'if ini>85 else'нейтральна зона')
    return(f'<div class="panel pct-panel" id="pctm_panel_{s}" style="margin-top:8px">'
           f'<div class="plbl">ПЕРЦЕНТИЛЬ (COT INDEX RANKED M)</div>'
           f'<div class="pct-sel-row"><div class="psel-group">'
           f'<button class="psel active" data-p="ls" onclick="pctMSel(this,\'{s}\')">LS</button>'
           f'<button class="psel" data-p="cm" onclick="pctMSel(this,\'{s}\')">CM</button>'
           f'<button class="psel" data-p="st" onclick="pctMSel(this,\'{s}\')">ST</button>'
           f'</div><div class="psel-sep"></div><div class="psel-group">'
           f'<button class="pper active" data-per="all" onclick="pperMSel(this,\'{s}\')">All</button>'
           f'<button class="pper" data-per="3y"  onclick="pperMSel(this,\'{s}\')">3Y</button>'
           f'<button class="pper" data-per="1y"  onclick="pperMSel(this,\'{s}\')">1Y</button>'
           f'<button class="pper" data-per="6m"  onclick="pperMSel(this,\'{s}\')">6M</button>'
           f'<button class="pper" data-per="3m"  onclick="pperMSel(this,\'{s}\')">3M</button>'
           f'</div></div>'
           f'<div class="pct-val-row">'
           f'<span id="pctmval_{s}" style="font-size:16px;font-weight:bold;color:{ini_color}">{ini:.1f}%</span>'
           f'<span id="pctmcls_{s}" style="font-size:11px;color:#8090b0;margin-left:8px;">— {ini_lbl}</span>'
           f'</div>'
           f'<div class="pbar-wrap"><div class="pbar-bg"><div class="pbar-lo"></div><div class="pbar-hi"></div>'
           f'<div class="ptick" style="left:15%"></div><div class="ptick" style="left:85%"></div>'
           f'<div class="pbar-mk" id="pctmmk_{s}" style="left:{ini_pos:.1f}%"></div></div>'
           f'<div class="ptick-labels"><span class="ptlbl" style="left:15%">15%</span>'
           f'<span id="pctmcur_{s}" class="ptlbl ptlbl-cur" style="left:{ini_pos:.1f}%">{ini:.1f}%</span>'
           f'<span class="ptlbl" style="left:85%">85%</span></div></div>'
           f'<div class="pbar-lb"><span>0%</span><span>50%</span><span>100%</span></div></div>')

def make_instrument_view(d,tff=None,disag=None):
    c=d['cur'];s=d['sid'];sm=d['sm']
    has_tff=tff is not None;has_disag=disag is not None
    mc_ls=make_metric_card('LARGE SPEC (NETTO)',c['ls_net'],c['ls_chg'],c['ls_chg_pct'],d['spark']['ls'],COLOR_LS,gauge_val=d['cot_idx']['ls']['all'],sub_text="",ranked_val=d.get('cot_idx_m',{}).get('ls',{}).get('all'),lbl_color=COLOR_LS,bar_id=f'mcbar_ls_{s}',border_color=COLOR_LS+'66')
    mc_cm=make_metric_card('COMMERCIALS (NETTO)',c['cm_net'],c['cm_chg'],c['cm_chg_pct'],d['spark']['cm'],COLOR_CM,gauge_val=d['cot_idx']['cm']['all'],sub_text="",ranked_val=d.get('cot_idx_m',{}).get('cm',{}).get('all'),lbl_color=COLOR_CM,bar_id=f'mcbar_cm_{s}',border_color=COLOR_CM+'66')
    mc_st=make_metric_card('SMALL TRADERS (NETTO)',c['st_net'],c['st_chg'],c['st_chg_pct'],d['spark']['st'],COLOR_ST,gauge_val=d['cot_idx']['st']['all'],sub_text="",ranked_val=d.get('cot_idx_m',{}).get('st',{}).get('all'),lbl_color=COLOR_ST,bar_id=f'mcbar_st_{s}',border_color=COLOR_ST+'66')
    mc_oi=make_metric_card('OPEN INTEREST',c['oi'],c['oi_chg'],c['oi_chg_pct'],d['spark']['oi'],'#a0aac0',oi=True,gauge_val=d.get('oi_capacity',50.0),sub_text=f"зміна: {fv(int(c['oi_chg']),True,sign=True)}",bar_id=f'mcbar_oi_{s}')
    # v26 analysis calls — передаємо cot_idx для gauges, обгортка як у TFF
    _ci=d['cot_idx']
    analysis_panel=(f'<div class="panel tff-analysis-panel">'
                    +analysis_row('LARGE SPEC',COLOR_LS,c['ls_net'],c['ls_cl'],c['ls_cs'],c['ls_chg'],c['ls_chg_pct'],_ci['ls'])
                    +analysis_row('COMMERCIALS',COLOR_CM,c['cm_net'],c['cm_cl'],c['cm_cs'],c['cm_chg'],c['cm_chg_pct'],_ci['cm'])
                    +f'</div>')
    # v55: панель SM DIVERGENCE прибрана з mid — pct_combined займає всю ширину.
    # sm_bar() лишається: він ще потрібен для колонок SM DIV у Legacy-таблиці.
    # v16: панель ЗВІТИ видалена з усіх видів
    ini=d['cot_idx']['ls']['all'];ini_pos=min(max(ini,0),100)
    ini_color='#f0515a'if ini<15 else('#20d483'if ini>85 else'#dde2ee')
    ini_lbl='екстрем. шорт'if ini<15 else('екстрем. лонг'if ini>85 else'нейтральна зона')
    cj=json.dumps(d['cot_idx'],ensure_ascii=False)
    pct_panel=(f'<div class="panel pct-panel"><div class="plbl">ПЕРЦЕНТИЛЬ (COT INDEX)</div>'
               f'<div class="pct-sel-row"><div class="psel-group">'
               f'<button class="psel active" data-p="ls" onclick="pctSel(this,\'{s}\')">LS</button>'
               f'<button class="psel" data-p="cm" onclick="pctSel(this,\'{s}\')">CM</button>'
               f'<button class="psel" data-p="st" onclick="pctSel(this,\'{s}\')">ST</button>'
               f'</div><div class="psel-sep"></div><div class="psel-group">'
               f'<button class="pper active" data-per="all" onclick="pperSel(this,\'{s}\')">All</button>'
               f'<button class="pper" data-per="3y" onclick="pperSel(this,\'{s}\')">3Y</button>'
               f'<button class="pper" data-per="1y" onclick="pperSel(this,\'{s}\')">1Y</button>'
               f'<button class="pper" data-per="6m" onclick="pperSel(this,\'{s}\')">6M</button>'
               f'<button class="pper" data-per="3m" onclick="pperSel(this,\'{s}\')">3M</button>'
               f'</div></div>'
               f'<div class="pct-val-row"><span id="pctval_{s}" style="font-size:16px;font-weight:bold;color:{ini_color}">{fp(ini)}</span>'
               f'<span id="pctlbl_{s}" style="font-size:11px;color:#8090b0;margin-left:8px;">— {ini_lbl}</span></div>'
               f'<div class="pbar-wrap"><div class="pbar-bg"><div class="pbar-lo"></div><div class="pbar-hi"></div>'
               f'<div class="ptick" style="left:15%"></div><div class="ptick" style="left:85%"></div>'
               f'<div class="pbar-mk" id="pctmk_{s}" style="left:{ini_pos:.1f}%"></div></div>'
               f'<div class="ptick-labels"><span class="ptlbl" style="left:15%">15%</span>'
               f'<span id="pctcur_{s}" class="ptlbl ptlbl-cur" style="left:{ini_pos:.1f}%">{fp(ini)}</span>'
               f'<span class="ptlbl" style="left:85%">85%</span></div></div>'
               f'<div class="pbar-lb"><span>0%</span><span>50%</span><span>100%</span></div></div>'
                   + f'<script>_ci["{s}"]={cj};_ci_m["{s}"]={json.dumps(d.get("cot_idx_m",{}),ensure_ascii=False)};</script>')
    cj2=json.dumps(d['chart'],ensure_ascii=False)
    chart_block=(f'<div class="chartbox"><div class="chartbox-hdr"><div class="plbl" style="margin:0">ЧИСТІ ПОЗИЦІЇ</div>'
                 f'<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap"><div class="period-btns">'
                 f'<button class="per-btn active" data-per="1y" onclick="setChartPer(this,\'{s}\')">1 рік</button>'
                 f'<button class="per-btn" data-per="3y" onclick="setChartPer(this,\'{s}\')">3 роки</button>'
                 f'<button class="per-btn" data-per="5y" onclick="setChartPer(this,\'{s}\')">5 років</button>'
                 f'</div><div class="chart-leg">'
                 f'<span><span class="ll" style="background:{COLOR_LS}"></span>Large Spec</span>'
                 f'<span><span class="ll" style="background:{COLOR_CM}"></span>Commercials</span>'
                 f'<span class="ll-dash" style="border-top-color:{COLOR_ST}"></span><span style="margin-left:4px">Small Traders</span>'
                 f'</div></div></div><div class="cw"><canvas id="cv_{s}"></canvas></div></div>'
                 f'<script>_cd["{s}"]={cj2};</script>')
    bar_block=(f'<div class="bar-charts-grid chartbox">'
               f'<div class="bar-wrap"><div class="bar-lbl" style="color:{COLOR_LS}">LARGE SPEC — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="barcv_ls_{s}"></canvas></div></div>'
               f'<div class="bar-wrap"><div class="bar-lbl" style="color:{COLOR_CM}">COMMERCIALS — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="barcv_cm_{s}"></canvas></div></div>'
               f'<div class="bar-wrap"><div class="bar-lbl" style="color:{COLOR_ST}">SMALL TRADERS — WEEKLY ΔNet</div><div class="bar-cw"><canvas id="barcv_st_{s}"></canvas></div></div>'
               f'</div>')
    # v44: та сама верстка, що й на вкладці TABLE, обрізана по OPEN INTEREST
    table_block=(f'<div class="htable-wrap"><div class="htable-hdr"><span>ТИЖНЕВА СТАТИСТИКА ПОЗИЦІЙ</span>'
                 f'<div class="hsel">'
                 f'<button class="hbtn active" data-n="10" onclick="setMiniHist(this,\'{s}\')">10</button>'
                 f'<button class="hbtn" data-n="26" onclick="setMiniHist(this,\'{s}\')">26</button>'
                 f'<button class="hbtn" data-n="52" onclick="setMiniHist(this,\'{s}\')">52</button>'
                 f'</div></div><div class="tb-scroll tb-mini" id="mini_tbl_{s}"></div></div>')
    # Ranked блок — окрема панель під стандартним перцентилем
    ranked_panel = _make_ranked_section(s, d.get('cot_idx_m',{}))
    # Обгортаємо обидві pct-панелі у flex-колонку для 4-го grid-item
    pct_combined = (f'<div class="pct-combined">'
                    + pct_panel
                    + ranked_panel
                    + '</div>')
    # v58: TradingView — окремим блоком на всю ширину зверху; під ним двоколонкова
    # сітка: ліворуч картки 2×2, праворуч analysis + перцентиль. Bars/chart/table — нижче.
    lg_top=(make_tv_col(s)
            +f'<div class="lg-grid">'
             f'<div class="lg-cards"><div class="mcards">{mc_ls}{mc_cm}{mc_st}{mc_oi}</div></div>'
             f'<div class="lg-side">{analysis_panel}{pct_combined}</div>'
             f'</div>')
    legacy_sec=(f'<div class="rpt-sec" id="rpt_legacy_{s}">'
                +lg_top+bar_block+chart_block+table_block+'</div>')
    tff_sec  =make_tff_view(tff,s,make_reports_panel(s)) if has_tff else ''
    disag_sec=make_disag_view(disag,s,make_reports_panel(s)) if has_disag else ''
    # v22: приховуємо недоступні звіти повністю (замість disabled-кнопки)
    tff_btn=(f'<button class="rtab" data-rtype="tff" onclick="switchReport(\'{s}\',\'tff\')">TFF Report</button>'
             if has_tff else '')
    dg_btn=(f'<button class="rtab" data-rtype="dg" onclick="switchReport(\'{s}\',\'dg\')">Disaggregated</button>'
            if has_disag else '')
    CROP_LINK_SIDS = {'CORN','WHEAT','SOYBEAN','SOYBEAN_MEAL','SOYBEAN_OIL','COTTON','RICE'}
    crop_sheet_map = {'CORN':'corn','WHEAT':'springwheat','SOYBEAN':'soybeans','SOYBEAN_MEAL':'soybeans','SOYBEAN_OIL':'soybeans','COTTON':'cotton','RICE':'rice'}
    if s in CROP_LINK_SIDS:
        crop_cid = crop_sheet_map.get(s, s.lower())
        crop_btn = (f'<button class="rtab" data-rtype="crop" onclick="switchReport(\'{s}\',\'crop\')">Crop Progress</button>')
        # Вбудована crop секція всередині iview
        crop_sec = (f'<div class="rpt-sec" id="rpt_crop_{s}" style="display:none">'
                    f'<div id="crop_embed_{s}"></div>'
                    f'</div>')
    else:
        crop_btn = ''
        crop_sec = ''
    return(f'<div class="iview" id="iv_{s}" data-sid="{s}">'
           f'<div class="report-tabs"><span class="rtab-lbl">ТИП ЗВІТУ:</span>'
           f'<button class="rtab active" data-rtype="legacy" onclick="switchReport(\'{s}\',\'legacy\')">Legacy Report</button>'
           +tff_btn+dg_btn+crop_btn+
           f'</div>'+legacy_sec+tff_sec+disag_sec+crop_sec+'</div>')

# ================================================================
# OVERVIEW TAB (незмінно)
# ================================================================
def make_overview_tab():
    # v21 overview — компактна, кольорові фони, перемикач COT-періодів, CROWDED
    rows_html=[];rep_date='—';today_date='—'
    # збираємо cot-періоди для JS-перемикача: { sid: {ls:{all,3y,1y,6m,3m}, cm:{...}, st:{...}} }
    cot_periods={}
    # v49: максимуми по стовпцях — для пропорційної сили фону
    def _colmax(key):
        vals=[]
        for _it in OVERVIEW_TABLE:
            if not isinstance(_it,dict): continue
            v=_it.get(key)
            try:
                if v is not None: vals.append(abs(float(v)))
            except: pass
        return max(vals) if vals else 0.0
    _MX_OI=_colmax('oi_chg_pct'); _MX_AL=_colmax('acc_ls'); _MX_AC=_colmax('acc_cm')
    for item in OVERVIEW_TABLE:
        if isinstance(item,tuple) and item[0]=='_meta': rep_date=item[1];today_date=item[2];continue
        if isinstance(item,tuple) and item[0]=='_group': rows_html.append(f'<tr class="ov-group"><td colspan="19">{item[1]}</td></tr>');continue
        if not isinstance(item,dict): continue
        d=item
        s_=d.get('sid','')
        ci=d.get('cot_idx')
        if ci:
            cot_periods[s_]={
                'ls':ci.get('ls',{}),'cm':ci.get('cm',{}),'st':ci.get('st',{})
            }
        def fnum(v,sign=False):
            if v is None: return '<span class="d">—</span>'
            try:
                nv=int(round(float(v)));body=f"{abs(nv):,}".replace(',','\u202f')
                s2='+' if(sign and nv>0)else('-' if nv<0 else'');cls='g'if nv>0 else('r'if nv<0 else'd')
                return f'<span class="{cls}">{s2}{body}</span>'
            except: return '<span class="d">—</span>'
        def cell_bg(v,sign=True):
            # білі цифри на зелен/червон фоні (як фото2)
            if v is None: return '<td class="ov-num d">—</td>'
            try:
                nv=int(round(float(v)))
            except: return '<td class="ov-num d">—</td>'
            body=f"{abs(nv):,}".replace(',','\u202f')
            s2='+' if(nv>0)else('-' if nv<0 else'')
            if nv>0:   cls='ov-bg-g'
            elif nv<0: cls='ov-bg-r'
            else:      cls='ov-bg-0'
            return f'<td class="ov-num {cls}">{s2}{body}</td>'
        def pctcell_bg(v):
            # частка 0..1 -> %; білі цифри на фоні
            if v is None: return '<td class="ov-num d">—</td>'
            try: v2=float(v)*100
            except: return '<td class="ov-num d">—</td>'
            if abs(v2)>999: return '<td class="ov-num d">—</td>'
            s2='+' if v2>0 else ''
            if v2>0:   cls='ov-bg-g'
            elif v2<0: cls='ov-bg-r'
            else:      cls='ov-bg-0'
            return f'<td class="ov-num {cls}">{s2}{v2:.1f}%</td>'
        def pct_bar(v,lo=15,hi=85):
            if v is None: return ''
            pct=min(max(v/100,0),1);color='#20d483'if v<20 else('#f0515a'if v>80 else'#4a9eff')
            return f'<div class="ov-bar-bg"><div class="ov-bar-fill" style="width:{pct*100:.1f}%;background:{color}"></div></div>'
        def cot_td(sid_,grp):
            # клітинка COT з data-атрибутами всіх періодів для JS-перемикача
            v_all=d.get('cot_'+grp)
            per=d.get('cot_idx',{}).get(grp,{}) if d.get('cot_idx') else {}
            da=per.get('all',v_all);d3y=per.get('3y');d1y=per.get('1y');d6=per.get('6m');d3=per.get('3m')
            def _f(x):
                try: return f'{float(x):.0f}'
                except: return ''
            init=v_all if v_all is not None else 0
            _cls = 'ov-cot-lo' if init>80 else ('ov-cot-hi' if init<20 else '')
            return (f'<td class="ov-cot-cell-td" data-cotgrp="{grp}" '
                    f'data-all="{_f(da)}" data-3y="{_f(d3y)}" data-1y="{_f(d1y)}" '
                    f'data-6m="{_f(d6)}" data-3m="{_f(d3)}">'
                    f'<div class="ov-cot-cell">{pct_bar(init)}'
                    f'<span class="ov-cot-val {_cls}">{init:.0f}%</span></div></td>')
        def sm_fmt(v):
            if v is None: return '<span class="d">—</span>'
            cls='g'if float(v)>0 else('r'if float(v)<0 else'd'); return f'<span class="{cls}">{float(v):+.2f}</span>'
        def crowded_fmt(cw):
            if not cw or cw in('—','nan','None'): return '<span class="d">—</span>'
            low=cw.lower()
            if 'very' in low: return '<span class="ov-crowd ov-crowd-vc">Very Crowded</span>'
            if 'crowd' in low: return '<span class="ov-crowd ov-crowd-c">Crowded</span>'
            return f'<span class="d">{cw}</span>'
        def lead_fmt(v):
            if not v or str(v).strip() in ('—','-','nan','None',''):
                return '<span class="d">—</span>'
            return '<span class="ov-lead">YES</span>'
        def sm_td(v):
            # колір за порогом +-0.25, фон слабкий і пропорційний модулю
            if v is None: return '<td class="ov-num d">—</td>'
            try: f=float(v)
            except: return '<td class="ov-num d">—</td>'
            cls='g' if f>0.25 else ('r' if f<-0.25 else 'd')
            st=''
            if abs(f)>0.25:
                op=min(abs(f),1.0)*0.14
                rgb='32,212,131' if f>0 else '240,81,90'
                st=f' style="background:rgba({rgb},{op:.3f})"'
            return f'<td class="ov-num {cls}"{st}>{f:+.2f}</td>'
        def num_td_bg(v,mx):
            # число зі знаком, фон пропорційний модулю відносно максимуму стовпця
            if v is None: return '<td class="ov-num d">—</td>'
            try: nv=int(round(float(v)))
            except: return '<td class="ov-num d">—</td>'
            body=f"{abs(nv):,}".replace(',','\u202f')
            s2='+' if nv>0 else ('-' if nv<0 else '')
            cls='g' if nv>0 else ('r' if nv<0 else 'd')
            st=''
            if nv!=0 and mx:
                op=min(abs(nv)/mx,1.0)*0.42
                rgb='32,212,131' if nv>0 else '240,81,90'
                st=f' style="background:rgba({rgb},{op:.3f})"'
            return f'<td class="ov-num {cls}"{st}>{s2}{body}</td>'
        def pct_td_bg(v,mx):
            # відсоток, фон пропорційний модулю відносно максимуму стовпця
            if v is None: return '<td class="ov-num d">—</td>'
            try: f=float(v)
            except: return '<td class="ov-num d">—</td>'
            if abs(f)*100>9999: return '<td class="ov-num d">—</td>'
            cls='g' if f>0 else ('r' if f<0 else 'd')
            s2='+' if f>0 else ''
            st=''
            if f!=0 and mx:
                op=min(abs(f)/mx,1.0)*0.42
                rgb='32,212,131' if f>0 else '240,81,90'
                st=f' style="background:rgba({rgb},{op:.3f})"'
            return f'<td class="ov-num {cls}"{st}>{s2}{f*100:.1f}%</td>'
        def pct_td_plain(v):
            # відсоток без фону, колір за знаком
            if v is None: return '<td class="ov-num d">—</td>'
            try: f=float(v)*100
            except: return '<td class="ov-num d">—</td>'
            if abs(f)>9999: return '<td class="ov-num d">—</td>'
            cls='g' if f>0 else ('r' if f<0 else 'd')
            s2='+' if f>0 else ''
            return f'<td class="ov-num {cls}">{s2}{f:.1f}%</td>'
        def pct_td30(v):
            # без фону: >+30% зелений, <-30% червоний, решта сірий
            if v is None: return '<td class="ov-num d">—</td>'
            try: f=float(v)*100
            except: return '<td class="ov-num d">—</td>'
            if abs(f)>999: return '<td class="ov-num d">—</td>'
            cls='g' if f>30 else ('r' if f<-30 else 'd')
            s2='+' if f>0 else ''
            return f'<td class="ov-num {cls}">{s2}{f:.1f}%</td>'
        rows_html.append(f'<tr class="ov-row">'
                         f'<td class="ov-idx"></td>'
                         f'<td class="ov-asset"><span class="ov-fav" data-fav="{d["sid"]}" onclick="event.stopPropagation();ovToggleFav(this)">☆</span>'
                         f'<span class="ov-asset-link" onclick="ovGoTable(\'{d["sid"]}\')">{d["asset"]}</span></td>'
                         f'{fnum_td(d["net_ls"])}'
                         f'{fnum_td(d["net_cm"])}'
                         f'{cot_td(d["sid"],"ls")}'
                         f'{cot_td(d["sid"],"cm")}'
                         f'{cot_td(d["sid"],"st")}'
                         f'{sm_td(d["sm_div"])}'
                         f'{pct_td30(d.get("chg_pct_ls"))}'
                         f'{fnum_td(d["chg_ls"])}'
                         f'{pct_td30(d.get("chg_pct_cm"))}'
                         f'{fnum_td(d["chg_cm"])}'
                         f'<td>{lead_fmt(d.get("cm_lead"))}</td>'
                         f'{pct_td_plain(d.get("oi_chg_pct"))}'
                         f'{fnum_td(d.get("acc_ls"))}'
                         f'{fnum_td(d.get("acc_cm"))}'
                         f'{sm_td(d["sm_div_6m"])}'
                         f'{sm_td(d["sm_div_3m"])}'
                         f'<td>{crowded_fmt(d.get("crowded"))}</td></tr>')
    # v27: клікабельні заголовки для сортування
    thead=(f'<thead><tr>'
           f'<th class="ov-idx-th">#</th>'
           f'<th class="ov-asset ov-sortable" data-col="1" data-stype="reset" onclick="ovSort(this)" title="Скинути сортування">ASSET</th>'
           f'<th class="ov-sortable" data-col="2" onclick="ovSort(this)">NET LS</th>'
           f'<th class="ov-sortable" data-col="3" onclick="ovSort(this)">NET CM</th>'
           f'<th class="ov-sortable" data-col="4" data-stype="cot" onclick="ovSort(this)" style="color:{COLOR_LS}">COT LS</th>'
           f'<th class="ov-sortable" data-col="5" data-stype="cot" onclick="ovSort(this)" style="color:{COLOR_CM}">COT CM</th>'
           f'<th class="ov-sortable" data-col="6" data-stype="cot" onclick="ovSort(this)" style="color:{COLOR_ST}">COT ST</th>'
           f'<th class="ov-sortable" data-col="7" onclick="ovSort(this)">SM DIV</th>'
           f'<th class="ov-sortable" data-col="8" onclick="ovSort(this)">CHG %LS</th>'
           f'<th class="ov-sortable" data-col="9" onclick="ovSort(this)">CHG LS</th>'
           f'<th class="ov-sortable" data-col="10" onclick="ovSort(this)">CHG %CM</th>'
           f'<th class="ov-sortable" data-col="11" onclick="ovSort(this)">CHG CM</th>'
           f'<th class="ov-sortable" data-col="12" data-stype="crowd" onclick="ovSort(this)">CM LEAD</th>'
           f'<th class="ov-sortable" data-col="13" onclick="ovSort(this)">%OI CHG</th>'
           f'<th class="ov-sortable" data-col="14" onclick="ovSort(this)">ACC LS</th>'
           f'<th class="ov-sortable" data-col="15" onclick="ovSort(this)">ACC CM</th>'
           f'<th class="ov-sortable" data-col="16" onclick="ovSort(this)">SM DIV 6M</th>'
           f'<th class="ov-sortable" data-col="17" onclick="ovSort(this)">SM DIV 3M</th>'
           f'<th class="ov-sortable" data-col="18" data-stype="crowd" onclick="ovSort(this)">CROWDED ATH</th>'
           f'</tr></thead>')
    # дані для SM DIV графіків
    sm_chart_data = []
    for item in OVERVIEW_TABLE:
        if not isinstance(item, dict): continue
        sm_chart_data.append({
            'label': item.get('asset',''),
            'div':   item.get('sm_div',   0) or 0,
            'div_6m':item.get('sm_div_6m',0) or 0,
            'div_3m':item.get('sm_div_3m',0) or 0,
        })
    import json as _json
    sm_json = _json.dumps(sm_chart_data, ensure_ascii=False)

    sm_parts = [
        '<div class="ov-meta">Звіт: <b>' + rep_date + '</b> &nbsp;|&nbsp; Оновлено: ' + today_date + '</div>',
        # перемикач періодів COT
        '<div class="ov-per-row"><span class="ov-per-lbl">COT INDEX ПЕРІОД:</span>',
        '<button class="ov-per active" data-per="all" onclick="ovSetPer(this)">Весь час</button>',
        '<button class="ov-per" data-per="3y" onclick="ovSetPer(this)">3 роки</button>',
        '<button class="ov-per" data-per="1y" onclick="ovSetPer(this)">1 рік</button>',
        '<button class="ov-per" data-per="6m" onclick="ovSetPer(this)">6 міс</button>',
        '<button class="ov-per" data-per="3m" onclick="ovSetPer(this)">3 міс</button>',
        '<div class="ov-zoom">',
        '<button class="ov-zb" onclick="ovZoom(-1)" title="Дрібніше">&minus;</button>',
        '<span class="ov-zl" id="ovZoomLbl">17px</span>',
        '<button class="ov-zb" onclick="ovZoom(1)" title="Крупніше">+</button>',
        '<button class="ov-zb" onclick="ovZoomReset()" title="Скинути масштаб">СКИНУТИ</button>',
        '</div>',
        '</div>',
        '<div class="ov-scroll"><table class="ov-table">' + thead + '<tbody>' + ''.join(rows_html) + '</tbody></table></div>',
        '<div class="ov-sm-chart-wrap">',
        '<div class="ov-sm-tabs">',
        '<button class="ov-sm-tab active" onclick="selSmTab(this,\'div\')">SM DIV</button>',
        '<button class="ov-sm-tab" onclick="selSmTab(this,\'div_6m\')">SM 6M</button>',
        '<button class="ov-sm-tab" onclick="selSmTab(this,\'div_3m\')">SM 3M</button>',
        '</div>',
        '<div class="ov-sm-cv-wrap"><canvas id="ovSmChart"></canvas></div>',
        '</div>',
        '<script>window._ovSmInit=' + sm_json + ';</script>',
    ]
    return ''.join(sm_parts)

# ================================================================
# v54 — Overview для TFF і Disaggregated
# ================================================================
# (файл, аркуш, кольори трьох груп)
OV2_SOURCES = [
    ('tff', 'TFF',           lambda: TFF_FILE,   'Overview',
     [TFF_COLOR_LEV, TFF_COLOR_AM, TFF_COLOR_DL]),
    ('dg',  'DISAGGREGATED', lambda: DISAG_FILE, 'OVERVIEW',
     [DISAG_COLOR_MM, DISAG_COLOR_PM, DISAG_COLOR_SD]),
]
OV2_HDR_ROW  = 3    # рядок із підписами колонок (0-based)
OV2_DATA_ROW = 4    # перший рядок даних


def read_overview2(path, sheet):
    """Читає аркуш Overview з TFF/Disaggregated. Повертає (labels, rows)."""
    try:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
    except Exception as e:
        print(f"  ⚠  {path.name}/{sheet}: {e}")
        return [], []
    labels = []
    for c in range(2, 15):
        v = raw.iloc[OV2_HDR_ROW, c] if c < raw.shape[1] else None
        labels.append(str(v).strip() if pd.notna(v) else '')
    rows = []
    for i in range(OV2_DATA_ROW, len(raw)):
        asset = raw.iloc[i, 1]
        if pd.isna(asset): continue
        asset = str(asset).strip()
        if not asset or asset == 'nan': continue
        vals = []
        for c in range(2, 15):
            v = pd.to_numeric(raw.iloc[i, c], errors='coerce') if c < raw.shape[1] else None
            vals.append(float(v) if pd.notna(v) else None)
        if all(v is None for v in vals): continue
        rows.append({'asset': asset, 'v': vals})
    return labels, rows


def _ov2_num(v):
    """Ціле зі знаком, колір за знаком, без фону."""
    if v is None: return '<td class="ov-num d">—</td>'
    try: nv = int(round(float(v)))
    except: return '<td class="ov-num d">—</td>'
    body = f"{abs(nv):,}".replace(',', '\u202f')
    s2 = '+' if nv > 0 else ('-' if nv < 0 else '')
    cls = 'g' if nv > 0 else ('r' if nv < 0 else 'd')
    return f'<td class="ov-num {cls}">{s2}{body}</td>'


def _ov2_pct30(v):
    """Відсоток: >+30% зелений, <-30% червоний, решта сірий."""
    if v is None: return '<td class="ov-num d">—</td>'
    try: f = float(v) * 100
    except: return '<td class="ov-num d">—</td>'
    if abs(f) > 9999: return '<td class="ov-num d">—</td>'
    cls = 'g' if f > 30 else ('r' if f < -30 else 'd')
    s2 = '+' if f > 0 else ''
    return f'<td class="ov-num {cls}">{s2}{f:.1f}%</td>'


def _ov2_pct(v):
    """Відсоток, колір за знаком."""
    if v is None: return '<td class="ov-num d">—</td>'
    try: f = float(v) * 100
    except: return '<td class="ov-num d">—</td>'
    if abs(f) > 9999: return '<td class="ov-num d">—</td>'
    cls = 'g' if f > 0 else ('r' if f < 0 else 'd')
    s2 = '+' if f > 0 else ''
    return f'<td class="ov-num {cls}">{s2}{f:.1f}%</td>'


def _ov2_cot(v):
    """COT INDEX: смуга + значення. <20% зелений, >80% червоний."""
    if v is None: return '<td class="ov-num d">—</td>'
    try: p = max(0.0, min(100.0, float(v) * 100))
    except: return '<td class="ov-num d">—</td>'
    color = '#20d483' if p < 20 else ('#f0515a' if p > 80 else '#4a9eff')
    cls = 'ov-cot-hi' if p < 20 else ('ov-cot-lo' if p > 80 else '')
    return (f'<td><div class="ov-cot-cell">'
            f'<div class="ov-bar-bg"><div class="ov-bar-fill" '
            f'style="width:{p:.1f}%;background:{color}"></div></div>'
            f'<span class="ov-cot-val {cls}">{p:.0f}%</span></div></td>')


def make_overview2_tab(labels, rows, colors):
    """Таблиця Overview для TFF / Disaggregated."""
    if not rows:
        return '<p style="padding:24px;color:#8090b0">Немає даних</p>'
    # порядок колонок аркуша: NET x3, COT x3, (Chg%, Chg) x3, %OI Chg
    order = [(0,'num',None), (1,'num',None), (2,'num',None),
             (3,'cot',0),    (4,'cot',1),    (5,'cot',2),
             (6,'p30',None), (7,'num',None),
             (8,'p30',None), (9,'num',None),
             (10,'p30',None),(11,'num',None),
             (12,'pct',None)]
    th = ['<th class="ov-idx-th">#</th>',
          '<th class="ov-asset ov-sortable" data-col="1" data-stype="reset" '
          'onclick="ovSort(this)" title="Скинути сортування">ASSET</th>']
    for n, (ix, kind, gi) in enumerate(order):
        lbl = labels[ix] if ix < len(labels) and labels[ix] else f'C{ix}'
        st = ' data-stype="cot"' if kind == 'cot' else ''
        style = f' style="color:{colors[gi]}"' if gi is not None else ''
        th.append(f'<th class="ov-sortable" data-col="{n+2}"{st} '
                  f'onclick="ovSort(this)"{style}>{lbl}</th>')
    thead = '<thead><tr>' + ''.join(th) + '</tr></thead>'

    body = []
    for r in rows:
        tds = ['<td class="ov-idx"></td>',
               f'<td class="ov-asset"><span class="ov-asset-link" '
               f'onclick="ovGoTable(\'{sid(r["asset"])}\')">{r["asset"]}</span></td>']
        for ix, kind, _gi in order:
            v = r['v'][ix] if ix < len(r['v']) else None
            tds.append(_ov2_cot(v) if kind == 'cot' else
                       _ov2_pct30(v) if kind == 'p30' else
                       _ov2_pct(v) if kind == 'pct' else _ov2_num(v))
        body.append('<tr class="ov-row">' + ''.join(tds) + '</tr>')

    return ('<div class="ov-scroll"><table class="ov-table">' + thead
            + '<tbody>' + ''.join(body) + '</tbody></table></div>')


OV2_CSS = """
<style>
/* v54 — перемикач джерела Overview */
.ov-src{display:flex;gap:4px;margin:0 0 10px 0;}
.ov-srcb{padding:5px 20px;border:1px solid var(--bd);border-radius:3px;background:transparent;
  color:#b0bcd4;font-family:var(--f);font-size:12px;cursor:pointer;letter-spacing:1px;}
.ov-srcb:hover{border-color:var(--accent);color:#fff;}
.ov-srcb.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);font-weight:bold;}
.ov-sec{display:none;}
.ov-sec.active{display:block;}
</style>
"""

OV2_JS = """
<script>
// v54 — перемикання джерела Overview
function ovSrcSet(src,btn){
  document.querySelectorAll('.ov-srcb').forEach(function(b){b.classList.remove('active');});
  if(btn)btn.classList.add('active');
  document.querySelectorAll('.ov-sec').forEach(function(s){s.classList.remove('active');});
  const sec=document.getElementById('ovsec_'+src);
  if(sec)sec.classList.add('active');
  if(window.ovApplyZoom)ovApplyZoom();
  if(window.ovRenumber)ovRenumber();
  if(window.ovLoadFavs)ovLoadFavs();
}
</script>
"""


def make_overview_all():
    """Перемикач LEGACY / TFF / DISAGGREGATED + три секції."""
    secs = [('leg', 'LEGACY', make_overview_tab())]
    for key, title, getf, sheet, colors in OV2_SOURCES:
        path = getf()
        if not path.exists(): continue
        labels, rows = read_overview2(path, sheet)
        if not rows: continue
        print(f"  ✓  Overview {title}: {len(rows)} активів")
        secs.append((key, title, make_overview2_tab(labels, rows, colors)))

    if len(secs) == 1:
        return secs[0][2]

    btns = ''.join(
        f'<button class="ov-srcb{" active" if i == 0 else ""}" '
        f'onclick="ovSrcSet(\'{k}\',this)">{t}</button>'
        for i, (k, t, _h) in enumerate(secs))
    panes = ''.join(
        f'<div class="ov-sec{" active" if i == 0 else ""}" id="ovsec_{k}">{h}</div>'
        for i, (k, _t, h) in enumerate(secs))
    return f'<div class="ov-src">{btns}</div>{panes}'


def fnum_td(v):
    # NET-клітинка без фону (звичайна кольорова цифра)
    if v is None: return '<td class="ov-num d">—</td>'
    try:
        nv=int(round(float(v)));body=f"{abs(nv):,}".replace(',','\u202f')
        s2='+' if nv>0 else('-' if nv<0 else'');cls='g'if nv>0 else('r'if nv<0 else'd')
        return f'<td class="ov-num {cls}">{s2}{body}</td>'
    except: return '<td class="ov-num d">—</td>'



# ================================================================
# v29 — TABLE TAB (повна тижнева таблиця по інструменту)
# ================================================================
# 0 = вся історія. Якщо HTML стане завеликим — постав 520 (10 років)
# або 260 (5 років): це найпростіший спосіб зменшити розмір файлу.
TABLE_MAX_WEEKS = 0

# Групи колонок таблиці.
# Формат: (назва групи, колір, [(індекс колонки в Excel 0-based, підпис, тип)])
# Типи: int   — ціле число без знаку (Long/Short/OI)
#       chg   — зміна за тиждень (знак + градієнт фону)
#       net   — чиста позиція (знак, роздільник справа)
#       pct   — відсоток (у файлі частка 0..1)
#       cot   — COT INDEX (частка 0..1 -> %, порогові кольори 15/85)
#       ratio — коефіцієнт -1..1 (SM DIV)
#       txt   — текст (CM LEAD / CROWDED ATH)
TBL_GROUPS = [
    ('LARGE SPECULATORS', COLOR_LS, [
        (2,'LONG','grad'),(3,'SHORT','grad'),(4,'CHG LONG','chg'),(5,'CHG SHORT','chg'),
        (6,'%NET/OI','pct'),(7,'%OI CHG','pct'),(8,'NET LS','net')]),
    ('COMMERCIALS', COLOR_CM, [
        (9,'LONG','grad'),(10,'SHORT','grad'),(11,'CHG LONG','chg'),(12,'CHG SHORT','chg'),
        (13,'%NET/OI','pct'),(14,'%OI CHG','pct'),(15,'NET CM','net')]),
    ('SMALL TRADERS', COLOR_ST, [
        (16,'LONG','grad'),(17,'SHORT','grad'),
        (20,'%NET/OI','pct'),(21,'%OI CHG','pct'),(22,'NET ST','net')]),
    ('OPEN INTEREST', '#a0aac0', [
        (23,'%OI CHG','pct'),(24,'OPEN INT','oi')]),
    ('COT INDEX (ALL)', '#f0b429', [(25,'LS','cot'),(26,'CM','cot'),(27,'ST','cot')]),
    ('COT 5Y',          '#f0b429', [(28,'LS','cot'),(29,'CM','cot'),(30,'ST','cot')]),
    ('COT 3Y',          '#f0b429', [(31,'LS','cot'),(32,'CM','cot'),(33,'ST','cot')]),
    ('COT 1Y',          '#f0b429', [(34,'LS','cot'),(35,'CM','cot'),(36,'ST','cot')]),
    ('COT 6M',          '#f0b429', [(37,'LS','cot'),(38,'CM','cot'),(39,'ST','cot')]),
    ('COT 3M',          '#f0b429', [(40,'LS','cot'),(41,'CM','cot'),(42,'ST','cot')]),
    ('SM / FLOW / CYCLE', '#a78bfa', [
        (57,'SM DIV','ratio'),(58,'SM DIV 6M','ratio'),(59,'SM DIV 3M','ratio'),
        (61,'FLOW LS','chg'),(62,'VELOCITY CM','chg'),(64,'RISK CAP','cot'),
        (65,'CM LEAD','txt'),(66,'ACCEL LS','chg'),(67,'ACCEL CM','chg'),
        (68,'CROWDED ATH','txt')]),
]
TBL_COLS = [c for _g, _c, cols in TBL_GROUPS for c in cols]
# частки 0..1 та коефіцієнти зберігаємо як ціле ×1000 (компактніше в JSON)
TBL_SCALE = {'pct':1000, 'cot':1000, 'ratio':1000}


def _tbl_payload(df, colspec=None):
    """Готує компактний колонковий payload для вкладки TABLE.

    df — той самий DataFrame, що вже відфільтрований і відсортований у read_sheet().
    colspec — набір колонок (TBL_COLS для Legacy, TFF_TBL_COLS для TFF).
    Зберігаємо у зворотному порядку (найновіші зверху) — так рендер бере просто зріз.
    """
    if colspec is None: colspec = TBL_COLS
    try:
        work = df
        # відкидаємо порожні/#N/A рядки (немає ані Long, ані net у Large Spec)
        if work.shape[1] > 8:
            _keep = (pd.to_numeric(work.iloc[:, 2], errors='coerce').notna() |
                     pd.to_numeric(work.iloc[:, 8], errors='coerce').notna())
            work = work[_keep]
        if TABLE_MAX_WEEKS and len(work) > TABLE_MAX_WEEKS:
            work = work.tail(TABLE_MAX_WEEKS)
        dates = work['_dt'].dt.strftime('%d.%m.%Y').tolist()[::-1]
        cols = []
        n = len(work)
        for ci, _lbl, kind in colspec:
            if ci >= work.shape[1]:
                cols.append([None] * n); continue
            s = work.iloc[:, ci]
            if kind == 'txt':
                out = []
                for v in s:
                    t = str(v).strip() if pd.notna(v) else ''
                    out.append(t if t and t not in ('nan','—','-','#N/A','0') else None)
            else:
                mul = TBL_SCALE.get(kind, 1)
                nums = pd.to_numeric(s, errors='coerce')
                out = []
                for x in nums:
                    if pd.isna(x):
                        out.append(None); continue
                    fx = float(x)
                    if abs(fx) > 1e12:      # захист від timestamp-сміття
                        out.append(None); continue
                    out.append(int(round(fx * mul)))
            cols.append(out[::-1])
        return {'d': dates, 'c': cols}
    except Exception as e:
        print(f"    ⚠  _tbl_payload: {e}")
        return {'d': [], 'c': [[] for _ in colspec]}



# ================================================================
# v51 — TFF TABLE (колонки B..W + AH..BH)
# ================================================================
# У файлі стовпці BB..BH не мають заголовків — перейменуй тут за потреби.
# Назви взято з аркушів Disaggregated, де ці ж колонки підписані в рядку 3.
TFF_EXTRA_LABELS = {53:'VELOC DL', 54:'VELOC AM', 55:'VELOC LEV', 56:'RISK CAP',
                    57:'ACCEL DL', 58:'ACCEL AM', 59:'ACCEL LEV'}

def _tff_grp(base, name, color, net_lbl):
    """7 стандартних колонок групи TFF, base — індекс колонки Long"""
    return (name, color, [
        (base + 0, 'LONG', 'grad'), (base + 1, 'SHORT', 'grad'),
        (base + 2, 'CHG LONG', 'chg'), (base + 3, 'CHG SHORT', 'chg'),
        (base + 4, '%NET/OI', 'pct'), (base + 5, '%OI CHG', 'pct'),
        (base + 6, net_lbl, 'net')])

TFF_TBL_GROUPS = [
    _tff_grp(16, 'LEV MONEY',  TFF_COLOR_LEV, 'NET LEV'),
    _tff_grp(9,  'ASSET MGR',  TFF_COLOR_AM,  'NET AM'),
    _tff_grp(2,  'DEALER',     TFF_COLOR_DL,  'NET DL'),
    ('OPEN INTEREST', '#a0aac0', [(33, '%OI CHG', 'pct'), (34, 'OPEN INT', 'oi')]),
    ('COT INDEX (ALL)', '#f0b429', [(35,'DL','cot'),(36,'AM','cot'),(37,'LEV','cot')]),
    ('COT 5Y',          '#f0b429', [(38,'DL','cot'),(39,'AM','cot'),(40,'LEV','cot')]),
    ('COT 3Y',          '#f0b429', [(41,'DL','cot'),(42,'AM','cot'),(43,'LEV','cot')]),
    ('COT 1Y',          '#f0b429', [(44,'DL','cot'),(45,'AM','cot'),(46,'LEV','cot')]),
    ('COT 6M',          '#f0b429', [(47,'DL','cot'),(48,'AM','cot'),(49,'LEV','cot')]),
    ('COT 3M',          '#f0b429', [(50,'DL','cot'),(51,'AM','cot'),(52,'LEV','cot')]),
    ('EXTRA', '#a78bfa', [(53, TFF_EXTRA_LABELS[53], 'chg'), (54, TFF_EXTRA_LABELS[54], 'chg'),
                          (55, TFF_EXTRA_LABELS[55], 'chg'), (56, TFF_EXTRA_LABELS[56], 'cot'),
                          (57, TFF_EXTRA_LABELS[57], 'chg'), (58, TFF_EXTRA_LABELS[58], 'chg'),
                          (59, TFF_EXTRA_LABELS[59], 'chg')]),
]
TFF_TBL_COLS = [c for _g, _c, cols in TFF_TBL_GROUPS for c in cols]


def _tbl_assets_bar(dataset):
    """Селектор активів для одного джерела. Повертає (html, перший sid)."""
    GRP_COLOR = {'Валюти':'#4a9eff','Метали':'#f0b429','Індекси':'#a78bfa',
                 'Енергія':'#f59420','Агро':'#20d483','Крипто':'#22d3ee'}
    sel_parts = []
    first_sid = None
    for cat, instruments in CATEGORIES.items():
        avail = [i for i in instruments if i in dataset]
        if not avail: continue
        col = GRP_COLOR.get(cat, '#8090b0')
        btns = []
        for i in avail:
            s = dataset[i]['sid']
            if first_sid is None: first_sid = s
            btns.append(f'<button class="tb-a" data-sid="{s}" '
                        f'onclick="tblSel(\'{s}\',this)">{disp(i)}</button>')
        sel_parts.append(f'<div class="tb-grow">'
                         f'<div class="tb-grp" style="color:{col}">{cat.upper()}</div>'
                         f'<div class="tb-gbtns">' + ''.join(btns) + '</div></div>')
    return '<div class="tb-assets">' + ''.join(sel_parts) + '</div>', first_sid


def _tbl_heads(groups):
    """Шапка таблиці + скорочена шапка до OPEN INTEREST + spec колонок."""
    r1 = '<th class="tb-corner" rowspan="2">ДАТА</th>'
    r2 = ''
    m1 = r1; m2 = ''
    mini_cols = 0
    MINI_LAST = 'OPEN INTEREST'
    in_mini = True
    for gname, gcol, cols in groups:
        r_, g_, b_ = int(gcol[1:3],16), int(gcol[3:5],16), int(gcol[5:7],16)
        gh = (f'<th colspan="{len(cols)}" class="tb-g" '
              f'style="background:rgba({r_},{g_},{b_},.18);color:#fff;'
              f'border-left:2px solid {gcol}99">{gname}</th>')
        r1 += gh
        if in_mini: m1 += gh
        for k, (_ci, lbl, kind) in enumerate(cols):
            bl = f'border-left:2px solid {gcol}55;' if k == 0 else ''
            if kind == 'net':   # NET-колонка обведена рамкою кольору групи
                bl = f'border-left:2px solid {gcol};border-right:2px solid {gcol};'
            th = f'<th class="tb-s" style="{bl}">{lbl}</th>'
            r2 += th
            if in_mini:
                m2 += th; mini_cols += 1
        if gname == MINI_LAST: in_mini = False
    thead = f'<thead><tr class="tb-r1">{r1}</tr><tr class="tb-r2">{r2}</tr></thead>'
    mini_thead = f'<thead><tr class="tb-r1">{m1}</tr><tr class="tb-r2">{m2}</tr></thead>'
    spec = []
    for _g, gcol, cols in groups:
        for k, (_ci, _lbl, kind) in enumerate(cols):
            spec.append({'k': kind, 's': 1 if k == len(cols) - 1 else 0, 'c': gcol,
                         'i': 1 if _lbl == 'SHORT' else 0,
                         'n': 1 if _lbl == '%NET/OI' else 0,
                         'l': _lbl,
                         'ct': 1 if _g.startswith('COT') else 0})
    return thead, mini_thead, mini_cols, spec


# ================================================================
# v53 — DISAGGREGATED TABLE (структура ідентична TFF)
# ================================================================
DISAG_TBL_GROUPS = [
    _tff_grp(2,  'MAN MONEY',    DISAG_COLOR_MM, 'NET MM'),
    _tff_grp(9,  'PROD/MERCH',   DISAG_COLOR_PM, 'NET PM'),
    _tff_grp(16, 'SWAP DEALERS', DISAG_COLOR_SD, 'NET SD'),
    ('OPEN INTEREST', '#a0aac0', [(33, '%OI CHG', 'pct'), (34, 'OPEN INT', 'oi')]),
    ('COT INDEX (ALL)', '#f0b429', [(35,'MM','cot'),(36,'PM','cot'),(37,'SD','cot')]),
    ('COT 5Y',          '#f0b429', [(38,'MM','cot'),(39,'PM','cot'),(40,'SD','cot')]),
    ('COT 3Y',          '#f0b429', [(41,'MM','cot'),(42,'PM','cot'),(43,'SD','cot')]),
    ('COT 1Y',          '#f0b429', [(44,'MM','cot'),(45,'PM','cot'),(46,'SD','cot')]),
    ('COT 6M',          '#f0b429', [(47,'MM','cot'),(48,'PM','cot'),(49,'SD','cot')]),
    ('COT 3M',          '#f0b429', [(50,'MM','cot'),(51,'PM','cot'),(52,'SD','cot')]),
    # BB..BH підписані в рядку 3 аркуша Disaggregated
    ('EXTRA', '#a78bfa', [(53,'VELOC MM','chg'),(54,'VELOC PM','chg'),(55,'VELOC SD','chg'),
                          (56,'RISK CAP','cot'),(57,'ACCEL MM','chg'),(58,'ACCEL PM','chg'),
                          (59,'ACCEL SD','chg')]),
]
DISAG_TBL_COLS = [c for _g, _c, cols in DISAG_TBL_GROUPS for c in cols]


def make_table_tab(data, tff_data=None, dg_data=None):
    """Вкладка TABLE — перемикач джерела, селектор активів, повна таблиця."""
    if not data:
        return '<p style="padding:24px;color:#8090b0">Немає даних</p>'
    tff_data = tff_data or {}
    dg_data  = dg_data or {}

    asset_bar,  first_sid  = _tbl_assets_bar(data)
    asset_barT, first_sidT = _tbl_assets_bar(tff_data)
    asset_barG, first_sidG = _tbl_assets_bar(dg_data)

    thead,  mini_thead, mini_cols, spec  = _tbl_heads(TBL_GROUPS)
    theadT, mini_theadT, mini_colsT, specT = _tbl_heads(TFF_TBL_GROUPS)
    theadG, mini_theadG, mini_colsG, specG = _tbl_heads(DISAG_TBL_GROUPS)

    def _fit_cols(groups):
        n = 1
        for gname, _c, cols in groups:
            n += len(cols)
            if gname == 'OPEN INTEREST': break
        return n

    payloads = []
    for key, d in data.items():
        tb = d.get('table')
        if not tb: continue
        payloads.append('_tbl["%s"]=%s;' % (d['sid'], json.dumps(tb, separators=(',', ':'))))
        payloads.append('_tblName["%s"]=%s;' % (d['sid'], json.dumps(d['display'], ensure_ascii=False)))
    for key, d in tff_data.items():
        tb = d.get('table')
        if not tb: continue
        payloads.append('_tblT["%s"]=%s;' % (d['sid'], json.dumps(tb, separators=(',', ':'))))
        payloads.append('_tblNameT["%s"]=%s;' % (d['sid'], json.dumps(disp(key), ensure_ascii=False)))
    for key, d in dg_data.items():
        tb = d.get('table')
        if not tb: continue
        payloads.append('_tblG["%s"]=%s;' % (d['sid'], json.dumps(tb, separators=(',', ':'))))
        payloads.append('_tblNameG["%s"]=%s;' % (d['sid'], json.dumps(disp(key), ensure_ascii=False)))

    head_js = ('<script>const _tbl={};const _tblName={};'
               'const _tblT={};const _tblNameT={};'
               'const _tblG={};const _tblNameG={};let _tblSrc=\'leg\';'
               'window._TBL_SPEC_G=' + json.dumps(specG, separators=(',', ':')) + ';'
               'window._TBL_THEAD_G=' + json.dumps(theadG, ensure_ascii=False) + ';'
               'window._TBL_ASSETS_G=' + json.dumps(asset_barG, ensure_ascii=False) + ';'
               'window._TBL_FITCOLS_G=' + str(_fit_cols(DISAG_TBL_GROUPS)) + ';'
               'window._TBL_MINI_THEAD_G=' + json.dumps(mini_theadG, ensure_ascii=False) + ';'
               'window._TBL_MINI_COLS_G=' + str(mini_colsG) + ';'
               'window._TBL_FIRST_G=' + json.dumps(first_sidG or '') + ';'
               'window._TBL_SPEC_L=' + json.dumps(spec, separators=(',', ':')) + ';'
               'window._TBL_SPEC_T=' + json.dumps(specT, separators=(',', ':')) + ';'
               'window._TBL_THEAD_L=' + json.dumps(thead, ensure_ascii=False) + ';'
               'window._TBL_THEAD_T=' + json.dumps(theadT, ensure_ascii=False) + ';'
               'window._TBL_ASSETS_L=' + json.dumps(asset_bar, ensure_ascii=False) + ';'
               'window._TBL_ASSETS_T=' + json.dumps(asset_barT, ensure_ascii=False) + ';'
               'window._TBL_FITCOLS_L=' + str(_fit_cols(TBL_GROUPS)) + ';'
               'window._TBL_FITCOLS_T=' + str(_fit_cols(TFF_TBL_GROUPS)) + ';'
               'window._TBL_MINI_THEAD=' + json.dumps(mini_thead, ensure_ascii=False) + ';'
               'window._TBL_MINI_COLS=' + str(mini_cols) + ';'
               'window._TBL_MINI_THEAD_T=' + json.dumps(mini_theadT, ensure_ascii=False) + ';'
               'window._TBL_MINI_COLS_T=' + str(mini_colsT) + ';'
               'window._TBL_FIRST_L=' + json.dumps(first_sid or '') + ';'
               'window._TBL_FIRST_T=' + json.dumps(first_sidT or '') + ';'
               'window._TBL_FIRST=' + json.dumps(first_sid or '') + ';</script>')
    data_js = '<script>' + ''.join(payloads) + '</script>'

    src_sw = ('<div class="tb-src">'
              '<button class="tb-srcb active" data-src="leg" '
              'onclick="tblSrcSet(\'leg\',this)">LEGACY</button>'
              + ('<button class="tb-srcb" data-src="tff" '
                 'onclick="tblSrcSet(\'tff\',this)">TFF</button>' if tff_data else '')
              + ('<button class="tb-srcb" data-src="dg" '
                 'onclick="tblSrcSet(\'dg\',this)">DISAGGREGATED</button>' if dg_data else '')
              + '</div>')

    per_btns = ''.join(
        f'<button class="tb-per{" active" if p=="6M" else ""}" data-n="{n}" '
        f'onclick="tblPer(this)">{p}</button>'
        for p, n in [('6M',26),('1Y',52),('2Y',104),('3Y',156),('5Y',260),('ВСЕ',999999)])

    body = (
        '<div class="tb-wrap">'
        + src_sw + asset_bar +
        '<div class="tb-hdr">'
        '<div class="tb-hl"><span class="tb-name" id="tblName">—</span>'
        '<button class="tb-dash" id="tblDashBtn" onclick="tblGoDash()">Dashboard</button>'
        '<span class="tb-meta" id="tblMeta"></span>'
        '<span class="tb-info"><span id="tblSmDiv">SM DIV —</span>'
        '<span class="tb-crowd tb-cw-n" id="tblCrowd">—</span></span></div>'
        '<div class="tb-zoom">'
        '<button class="tb-zb" onclick="tblZoom(-1)" title="Дрібніше">&minus;</button>'
        '<span class="tb-zl" id="tblZoomLbl">авто</span>'
        '<button class="tb-zb" onclick="tblZoom(1)" title="Крупніше">+</button>'
        '<button class="tb-zb tb-za" onclick="tblZoomAuto()" title="Підігнати під ширину екрана">АВТО</button>'
        '</div>'
        '<div class="tb-pers">' + per_btns + '</div>'
        # v60: кнопки горизонтального гортання таблиці (миша/трекпад працюють як раніше)
        '<div class="tb-hscroll">'
        '<button class="tbl-scroll-btn" onclick="tblScroll(\'tblScrollBox\',-1)" '
        'title="Гортати вліво">&#9664;</button>'
        '<button class="tbl-scroll-btn" onclick="tblScroll(\'tblScrollBox\',1)" '
        'title="Гортати вправо">&#9654;</button>'
        '</div>'
        '</div>'
        '<div class="tb-scroll" id="tblScrollBox"><table class="dt" id="dtTable">'
        + thead +
        '<tbody class="tb-stats" id="dtStats"></tbody>'
        '<tbody id="dtBody"></tbody>'
        '</table></div></div>'
    )
    return TBL_CSS + head_js + body + data_js + TBL_JS


TBL_CSS = """
<style>
/* ── v29 TABLE TAB ── */
.tb-wrap{padding:12px 16px;}
.tb-src{display:flex;gap:4px;margin-bottom:8px;}
.tb-srcb{padding:5px 20px;border:1px solid var(--bd);border-radius:3px;background:transparent;
  color:#b0bcd4;font-family:var(--f);font-size:12px;cursor:pointer;letter-spacing:1px;}
.tb-srcb:hover{border-color:var(--accent);color:#fff;}
.tb-srcb.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);font-weight:bold;}
.tb-assets{display:flex;flex-wrap:wrap;align-items:flex-start;gap:10px 22px;
  padding:9px 12px;
  background:var(--bg2);border:1px solid var(--bd);border-radius:5px;margin-bottom:10px;}
.tb-grow{flex:0 0 auto;}
.tb-grp{display:block;font-size:9px;font-weight:bold;letter-spacing:1.2px;margin:0 0 4px 0;}
.tb-gbtns{display:flex;flex-wrap:wrap;gap:3px;}
.tb-a{padding:2px 8px;border:1px solid transparent;border-radius:3px;background:transparent;
  color:#b0bcd4;font-family:var(--f);font-size:11px;cursor:pointer;transition:all .12s;}
.tb-a:hover{color:#fff;border-color:var(--bd);}
.tb-a.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);font-weight:bold;}
.tb-hdr{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;
  padding:9px 14px;background:var(--bg2);border:1px solid var(--bd);border-radius:5px 5px 0 0;
  border-bottom:none;}
.tb-hl{display:flex;align-items:baseline;gap:10px;}
.tb-name{font-size:18px;font-weight:bold;color:#fff;letter-spacing:1px;}
.tb-meta{font-size:9px;color:var(--d);}
.tb-pers{display:flex;gap:3px;}
.tb-zoom{display:flex;align-items:center;gap:3px;margin-right:10px;}
.tb-zb{min-width:24px;padding:3px 8px;border:1px solid var(--bd);border-radius:3px;
  background:transparent;color:#b0bcd4;font-family:var(--f);font-size:11px;cursor:pointer;}
.tb-zb:hover{border-color:var(--accent);color:#fff;}
.tb-za{font-size:9px;letter-spacing:.5px;}
.tb-zl{font-size:10px;color:var(--d);min-width:52px;text-align:center;}
.tb-per{padding:3px 11px;border:1px solid var(--bd);border-radius:3px;background:transparent;
  color:#b0bcd4;font-family:var(--f);font-size:10px;cursor:pointer;}
.tb-per:hover{border-color:var(--accent);color:#fff;}
.tb-per.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);font-weight:bold;}
/* v60: кнопки горизонтального гортання таблиці */
.tb-hscroll{display:flex;gap:3px;margin-left:10px;}
.tbl-scroll-btn{padding:3px 12px;border:1px solid var(--bd);border-radius:3px;background:transparent;
  color:#b0bcd4;font-family:var(--f);font-size:11px;cursor:pointer;}
.tbl-scroll-btn:hover{border-color:var(--accent);color:#fff;}
.tb-scroll{overflow:auto;max-height:calc(100vh - 210px);background:#07080c;
  border:1px solid var(--bd);border-radius:0 0 5px 5px;}
/* v38: розмір шрифту виставляє JS (tblFit), відступи в em — масштабуються разом */
table.dt{border-collapse:separate;border-spacing:0;font-size:18px;white-space:nowrap;
  -webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;}
table.dt th{padding:.42em .85em;font-weight:bold;font-size:.78em;letter-spacing:.6px;
  text-align:right;background:#12141b;position:sticky;z-index:4;
  text-shadow:none;}
table.dt tr.tb-r1 th{top:0;text-align:center;}
table.dt tr.tb-r2 th{top:38px;color:#fff;border-bottom:1px solid var(--bd);}
table.dt th.tb-corner{left:0;z-index:6;top:0;text-align:left;color:#fff;
  border-right:1px solid var(--bd);border-bottom:1px solid var(--bd);}
table.dt td{padding:.34em .84em;text-align:right;border-bottom:1px solid rgba(52,61,90,.45);
  border-right:1px solid rgba(128,144,176,.10);color:#75809a;
  font-weight:bold;text-shadow:none;}
table.dt td.tb-date{position:sticky;left:0;z-index:3;background:#12141b;color:var(--d);
  text-align:left;border-right:1px solid var(--bd);}
table.dt td.tb-sep{border-right:1px solid var(--bd);}
/* v31: без цих правил `table.dt td{color:#fff}` перебиває класи .g/.r/.t */
table.dt td.d{color:var(--d);}
table.dt td.g{color:var(--g);}
table.dt td.r{color:var(--r);}
table.dt td.t{color:#75809a;}
table.dt tbody.tb-stats td{background:#12141b;font-size:.9em;
  border-bottom:1px solid rgba(52,61,90,.9);}
table.dt tbody.tb-stats tr:last-child td{border-bottom:3px solid var(--bd);}
table.dt tbody.tb-stats td.tb-date{font-size:.8em;letter-spacing:.5px;font-weight:bold;}
table.dt tbody#dtBody tr:hover td{box-shadow:inset 0 1px 0 rgba(255,255,255,.25),
  inset 0 -1px 0 rgba(255,255,255,.25);}
table.dt tbody#dtBody tr:hover td.tb-date{box-shadow:inset 0 1px 0 rgba(255,255,255,.25),
  inset 0 -1px 0 rgba(255,255,255,.25),inset 3px 0 0 var(--accent);}
.tb-badge{font-size:.72em;padding:.12em .55em;border-radius:10px;font-weight:bold;}
/* v44 */
.tb-mini{max-height:none;border:none;border-radius:0;}
.tb-dash{padding:5px 14px;border:1px solid var(--accent);border-radius:3px;background:transparent;
  color:var(--accent);font-family:var(--f);font-size:12px;cursor:pointer;letter-spacing:.5px;}
.tb-dash:hover{background:rgba(245,148,32,.15);}
.tb-info{display:flex;align-items:center;gap:12px;font-size:14px;font-weight:bold;}
.tb-crowd{font-size:12px;padding:3px 11px;border-radius:10px;font-weight:bold;}
.tb-cw-c{background:rgba(240,180,41,.15);border:1px solid #f0b429;color:#f0b429;}
.tb-cw-v{background:rgba(240,81,90,.18);border:1px solid #f0515a;color:#f0515a;}
.tb-cw-n{color:var(--d);}
.tb-bd-y{background:rgba(32,212,131,.15);border:1px solid #20d483;color:#20d483;}
.tb-bd-c{background:rgba(240,180,41,.15);border:1px solid #f0b429;color:#f0b429;}
.tb-bd-v{background:rgba(240,81,90,.18);border:1px solid #f0515a;color:#f0515a;}
</style>
"""

TBL_JS = """
<script>
let _tblCur='',_tblN=26;

function tblFmtInt(v){if(v==null)return'—';const s=v<0?'-':'';
  return s+Math.abs(v).toLocaleString('uk-UA');}
function tblFmtSign(v){if(v==null)return'—';if(v===0)return'—';
  const s=v>0?'+':'-';return s+Math.abs(v).toLocaleString('uk-UA');}
function tblFmtPct(v){if(v==null)return'—';const f=v/10;
  if(!isFinite(f)||Math.abs(f)>99999)return'—';return(f>0?'+':'')+f.toFixed(1)+'%';}
function tblFmtCot(v){if(v==null)return'—';const f=v/10;
  if(!isFinite(f)||f<-999||f>999)return'—';return f.toFixed(0)+'%';}
function tblFmtRatio(v){if(v==null)return'—';const f=v/1000;
  if(!isFinite(f))return'—';return(f>0?'+':'')+f.toFixed(2);}
function tblCls(v){return v>0?'g':(v<0?'r':'d');}

// градієнт фону для колонок CHG — як у тижневій таблиці Legacy
function tblBg(v,mx){
  if(v==null||v===0||!mx)return'';
  const o=(0.10+Math.min(Math.abs(v)/mx,1)*0.62).toFixed(2);
  return v>0?' style="background:rgba(32,212,131,'+o+')"'
            :' style="background:rgba(240,81,90,'+o+')"';
}

// ── КОЛІРНА СХЕМА v34 ───────────────────────────────────────────
// TB_SCALE: 'pct' — позиція = перцентиль (скільки % тижнів були нижчими);
//           'lin' — позиція = лінійна частка між історичними min і max.
const TB_SCALE='lin';
// Смуги відкалібровані по референсу: межі 12 / 24 / 36 % діапазону min..max.
// Сіра зона навмисно НЕ по центру — LONG/SHORT мають скошений розподіл,
// більшість історії лежить у нижній третині діапазону.
// [верхня межа позиції у %, [r,g,b]]
const TB_BANDS=[[ 12,[240, 81, 90]],  // максимально червоний
                [ 24,[196,120,132]],  // червоний
                [ 36,[117,128,154]],  // сірий
                [ 48,[122,187,166]],  // слабо зелений
                [ 60,[ 70,200,142]],  // зелений
                [100,[ 32,212,131]]]; // максимально зелений
// межі сірої зони — використовуються ще й для фону
const TB_NEU_LO=24, TB_NEU_HI=36;
const _TB_NEU='#75809a';
// смуги для CHG: |Δ| у % від історичного максимуму по модулю
const TB_CHG_BANDS=[4,19,60];
// фон вмикається лише після TB_BG_START (частка 0..1) і росте до TB_BG_MAX
const TB_BG_START=0.65, TB_BG_MAX=0.50;

function _tbHex(c){return'#'+c.map(x=>x.toString(16).padStart(2,'0')).join('');}
// перцентильний ранг: яка частка тижнів мала значення менше за v (0..1)
function tblRank(v,arr){
  let lo=0,hi=arr.length;
  while(lo<hi){const m=(lo+hi)>>1;if(arr[m]<v)lo=m+1;else hi=m;}
  return arr.length>1?lo/(arr.length-1):0.5;
}
// позиція значення в історії колонки, 0..100
function tblPos(v,rg){
  if(!rg)return 50;
  if(TB_SCALE==='lin'){
    if(rg.mx==null||rg.mn==null||rg.mx===rg.mn)return 50;
    return Math.max(0,Math.min(100,(v-rg.mn)/(rg.mx-rg.mn)*100));
  }
  if(!rg.srt||rg.srt.length<3)return 50;
  return tblRank(v,rg.srt)*100;
}
// колір цифри за смугами; inv=1 для SHORT (дзеркально)
function tblBandColor(p,inv){
  const x=inv?100-p:p;
  for(let i=0;i<TB_BANDS.length;i++)if(x<=TB_BANDS[i][0])return _tbHex(TB_BANDS[i][1]);
  return _tbHex(TB_BANDS[TB_BANDS.length-1][1]);
}
// колір цифри для CHG за величиною зміни
function tblChgColor(r){
  const a=Math.abs(r);
  if(a<TB_CHG_BANDS[0])return _TB_NEU;              // < 4%  — сірий
  const g=r>0;
  if(a<TB_CHG_BANDS[1])return _tbHex(g?[122,187,166]:[196,120,132]);  // 4–19%  слабий
  if(a<TB_CHG_BANDS[2])return _tbHex(g?[ 70,200,142]:[224, 97,110]);  // 19–60% звичайний
  return _tbHex(g?[32,212,131]:[240,81,90]);                          // > 60%  максимальний
}
// фон для LONG/SHORT/NET: усередині сірої зони фону немає,
// далі росте пропорційно віддаленню до відповідного краю діапазону
const TB_POS_BG_MAX=0.20, TB_POS_BG_GAMMA=1.7;
function tblBgPos(x){
  let str,pos;
  if(x<TB_NEU_LO){str=(TB_NEU_LO-x)/TB_NEU_LO;pos=false;}
  else if(x>TB_NEU_HI){str=(x-TB_NEU_HI)/(100-TB_NEU_HI);pos=true;}
  else return'';
  const o=TB_POS_BG_MAX*Math.pow(str,TB_POS_BG_GAMMA);
  if(o<0.012)return'';
  return'background:rgba('+(pos?'32,212,131':'240,81,90')+','+o.toFixed(3)+');';
}
// фон-підсвітка: str 0..1 — наскільки близько до краю; pos=true -> зелений
function tblBgFill(str,pos){
  if(!(str>TB_BG_START))return'';
  const o=((str-TB_BG_START)/(1-TB_BG_START)*TB_BG_MAX).toFixed(2);
  return'background:rgba('+(pos?'32,212,131':'240,81,90')+','+o+');';
}
// ── SM DIV: смуги у % від |1.00| ──
const TB_SMD_BANDS=[30,60,80];
const TB_SMD_BG_MAX=0.30, TB_SMD_BG_GAMMA=1.5;
function tblSmColor(p){
  const a=Math.abs(p);
  if(a<TB_SMD_BANDS[0])return _TB_NEU;
  const g=p>0;
  // рівні розведені і за відтінком, і за світлістю
  if(a<TB_SMD_BANDS[1])return _tbHex(g?[ 61,122,102]:[125, 74, 84]);  // темний приглушений
  if(a<TB_SMD_BANDS[2])return _tbHex(g?[ 49,181,127]:[192, 74, 88]);  // середній
  return _tbHex(g?[34,238,146]:[255,85,96]);                          // яскравий
}
function tblBgSm(p){
  const a=Math.min(Math.abs(p),100)/100;
  const o=TB_SMD_BG_MAX*Math.pow(a,TB_SMD_BG_GAMMA);
  if(o<0.012)return'';
  return'background:rgba('+(p>0?'32,212,131':'240,81,90')+','+o.toFixed(3)+');';
}
// ── COT INDEX: інвертовані смуги (низький = зелено, високий = червоно) ──
const TB_COT_BG_MAX=0.25, TB_COT_BG_GAMMA=1.6;
function tblCotColor(p){
  p=Math.max(0,Math.min(100,p));   // короткі періоди інколи виходять за 0..100
  if(p<20)return _tbHex([ 34,238,146]);   // яскравий зелений
  if(p<40)return _tbHex([ 61,143,107]);   // темний приглушений зелений
  if(p<60)return _TB_NEU;                 // сірий
  if(p<80)return _tbHex([156, 76, 88]);   // темний приглушений червоний
  return _tbHex([255,85,96]);             // яскравий червоний
}
function tblBgCot(p){
  p=Math.max(0,Math.min(100,p));
  const a=Math.min(Math.abs(p-50)/50,1);
  const o=TB_COT_BG_MAX*Math.pow(a,TB_COT_BG_GAMMA);
  if(o<0.012)return'';
  return'background:rgba('+(p<50?'32,212,131':'240,81,90')+','+o.toFixed(3)+');';
}
// фон для CHG: непрозорість лінійно від |Δ| / історичний максимум |Δ|.
// Тінт є майже всюди, слабкий для звичайних тижнів і сильний для рідкісних.
const TB_CHG_BG_MAX=0.45, TB_CHG_BG_GAMMA=1.0;
function tblBgChg(v,rg){
  const mxa=(rg&&rg.mxa)?rg.mxa:0;
  if(!mxa)return'';
  const o=TB_CHG_BG_MAX*Math.pow(Math.min(Math.abs(v)/mxa,1),TB_CHG_BG_GAMMA);
  if(o<0.012)return'';
  return'background:rgba('+(v>0?'32,212,131':'240,81,90')+','+o.toFixed(3)+');';
}
// Фон для %NET/OI та %OI CHG. Шкала — частка від «робастного максимуму»
// (90-й перцентиль |v| за всю історію): звичайний максимум у %OI CHG
// буває 18873% і занулив би решту комірок.
const TB_PCT_BG_MAX=0.32;
// %NET/OI майже не змінюється тиждень до тижня, тому в нього окрема,
// набагато слабша шкала — інакше колонка стає суцільною плямою
const TB_PCTN_BG_MAX=0.15, TB_PCTN_BG_GAMMA=1.8;
function tblBgPct(v,rg,soft){
  if(!rg||!rg.p90)return'';
  const r=Math.min(Math.abs(v)/rg.p90,1);
  const o=soft?TB_PCTN_BG_MAX*Math.pow(r,TB_PCTN_BG_GAMMA):TB_PCT_BG_MAX*r;
  if(o<0.012)return'';
  return'background:rgba('+(v>0?'32,212,131':'240,81,90')+','+o.toFixed(3)+');';
}
// %NET/OI та %OI CHG: >+30% зелений, <-30% червоний, решта — дефолт
function tblPctCls(v){const f=v/10;return f>30?'g':(f<-30?'r':'t');}

function tblCell(kind,v,mx,sep,rg,gc,inv,soft,ct){
  const sp=sep?' tb-sep':'';
  const nb=' border-left:2px solid '+(gc||'#343d5a')+';border-right:2px solid '+(gc||'#343d5a')+';';
  if(kind==='txt'){
    if(!v)return'<td class="d'+sp+'">—</td>';
    const low=String(v).toLowerCase();
    let c='tb-bd-y';
    if(low.indexOf('very')>=0)c='tb-bd-v';
    else if(low.indexOf('crowd')>=0)c='tb-bd-c';
    return'<td class="'+sp.trim()+'"><span class="tb-badge '+c+'">'+v+'</span></td>';
  }
  if(v==null)return'<td class="d'+sp+'"'+(kind==='net'?' style="'+nb+'"':'')+'>—</td>';
  if(kind==='int')  return'<td class="t'+sp+'">'+tblFmtInt(v)+'</td>';
  if(kind==='grad'||kind==='net'){
    const p=tblPos(v,rg);
    const col=tblBandColor(p,inv);
    const bg=tblBgPos(inv?100-p:p);
    const bold=(kind==='net')?'font-weight:bold;':'';
    const brd=(kind==='net')?nb:'';
    const txt=(kind==='net')?tblFmtSign(v):tblFmtInt(v);
    return'<td class="'+sp.trim()+'" style="'+bold+brd+bg+'color:'+col+'">'+txt+'</td>';
  }
  if(kind==='oi'){
    const ath=(rg&&rg.mx!=null&&v>=rg.mx);
    return'<td class="'+(ath?'g':'t')+sp+'"'+(ath?' style="font-weight:bold"':'')
          +'>'+tblFmtInt(v)+'</td>';
  }
  if(kind==='chg'){
    const mxa=(rg&&rg.mxa)?rg.mxa:1;
    const r=v/mxa*100;
    const bg=tblBgChg(v,rg);
    return'<td class="'+sp.trim()+'" style="'+bg+'color:'+tblChgColor(r)+'">'
          +tblFmtSign(v)+'</td>';
  }
  if(kind==='pct')  return'<td class="'+tblPctCls(v)+sp+'" style="'+tblBgPct(v,rg,soft)
                          +'">'+tblFmtPct(v)+'</td>';
  if(kind==='ratio'){
    const p=v/10;                       // v = коефіцієнт x1000 -> % від |1.00|
    return'<td class="'+sp.trim()+'" style="'+tblBgSm(p)+'color:'+tblSmColor(p)+'">'
          +tblFmtRatio(v)+'</td>';
  }
  if(kind==='cot'){
    const p=v/10;
    if(ct)  // справжній COT INDEX — інвертовані смуги
      return'<td class="'+sp.trim()+'" style="'+tblBgCot(p)+'color:'+tblCotColor(p)+'">'
            +tblFmtCot(v)+'</td>';
    const cl=v>850?'g':(v<150?'r':'t');   // RISK CAP — стара логіка
    return'<td class="'+cl+sp+'">'+tblFmtCot(v)+'</td>';
  }
  return'<td'+sp+'>'+v+'</td>';
}

// підсумкові рядки: MAX/MIN за весь час, MAX/MIN за 5 років, середнє за 13 тижнів
// ── v51: два джерела даних — Legacy (leg) і TFF ──
// суфікс глобальних змінних поточного джерела: _L / _T / _G
function _tblSfx(){return _tblSrc==='tff'?'_T':(_tblSrc==='dg'?'_G':'_L');}
function _tblVar(base){return window[base+_tblSfx()];}
function tblD(){return _tblSrc==='tff'?_tblT:(_tblSrc==='dg'?_tblG:_tbl);}
function tblNm(){return _tblSrc==='tff'?_tblNameT:(_tblSrc==='dg'?_tblNameG:_tblName);}
function tblSpec(){return _tblVar('_TBL_SPEC');}
function tblFitCols(){return _tblVar('_TBL_FITCOLS')||22;}
function tblSrcSet(src,btn){
  if(_tblSrc===src)return;
  _tblSrc=src;
  document.querySelectorAll('.tb-srcb').forEach(function(b){b.classList.remove('active');});
  if(btn)btn.classList.add('active');
  const ab=document.querySelector('.tb-assets');
  if(ab)ab.outerHTML=_tblVar('_TBL_ASSETS');
  const tbl=document.getElementById('dtTable');
  const th=tbl?tbl.querySelector('thead'):null;
  if(th)th.outerHTML=_tblVar('_TBL_THEAD');
  _tblCur='';
  const first=_tblVar('_TBL_FIRST');
  if(first&&tblD()[first])tblSel(first,null);
}

// Підсумкові рядки. Кожен рядок цілком одного кольору:
// MAX(ALL) насичений зелений, MAX(5Y) приглушений; MIN — дзеркально червоним.
const TB_STAT_ROWS=[['MAX (ALL)',0,'max','#20d483'],
                    ['MIN (ALL)',0,'min','#f0515a'],
                    ['MAX (5Y)',260,'max','#7abba6'],
                    ['MIN (5Y)',260,'min','#c47884']];
function tblStatsRows(d,LIM,S){
  if(!S)S=tblSpec();
  const N=d.d.length;
  if(LIM==null)LIM=S.length;
  let html='';
  for(const[lbl,win,mode,rcol]of TB_STAT_ROWS){
    const lim=win?Math.min(win,N):N;
    let tds='<td class="tb-date" style="color:'+rcol+'">'+lbl+'</td>';
    for(let ci=0;ci<LIM;ci++){
      const k=S[ci].k,sep=S[ci].s?' tb-sep':'';
      const brd=(k==='net')?'border-left:2px solid '+S[ci].c
                +';border-right:2px solid '+S[ci].c+';':'';
      if(k==='txt'){
        tds+='<td class="'+sep.trim()+'" style="'+brd+'color:'+rcol+'">—</td>';continue;
      }
      const col=d.c[ci];let acc=null;
      for(let i=0;i<lim;i++){
        const v=col[i];if(v==null)continue;
        if(mode==='max')acc=(acc==null||v>acc)?v:acc;
        else acc=(acc==null||v<acc)?v:acc;
      }
      let txt='—';
      if(acc!=null){
        if(k==='int'||k==='oi'||k==='grad')txt=tblFmtInt(acc);
        else if(k==='pct')txt=tblFmtPct(acc);
        else if(k==='cot')txt=tblFmtCot(acc);
        else if(k==='ratio')txt=tblFmtRatio(acc);
        else txt=tblFmtSign(acc);
      }
      tds+='<td class="'+sep.trim()+'" style="'+brd+'color:'+rcol+'">'+txt+'</td>';
    }
    html+='<tr>'+tds+'</tr>';
  }
  return html;
}

// ── Автомасштаб: підбирає розмір шрифту так, щоб колонки до OPEN INTEREST
// включно вміщались у видиму ширину на будь-якому екрані ──
const TB_FIT_COLS=22, TB_FIT_BASE=18, TB_FIT_MIN=8, TB_FIT_MAX=32;
// null = автопідбір під ширину екрана, число = ручний розмір шрифту у px
let _tblZoom=null;
try{const _z=localStorage.getItem('tbl_zoom');if(_z)_tblZoom=parseFloat(_z);}catch(e){}

// підбирає розмір так, щоб ДАТА … OPEN INTEREST вмістились у видиму ширину
function tblAutoSize(tbl,sc){
  tbl.style.fontSize=TB_FIT_BASE+'px';
  const row=tbl.querySelector('tbody#dtBody tr')||tbl.querySelector('tbody tr');
  if(!row)return TB_FIT_BASE;
  let w=0;const n=Math.min(tblFitCols(),row.children.length);
  for(let i=0;i<n;i++)w+=row.children[i].getBoundingClientRect().width;
  const avail=sc.clientWidth-2;
  if(!(w>0&&avail>0))return TB_FIT_BASE;
  return Math.max(TB_FIT_MIN,Math.min(TB_FIT_MAX,TB_FIT_BASE*avail/w));
}
function tblFit(){
  const tbl=document.getElementById('dtTable'),sc=document.querySelector('.tb-scroll');
  if(!tbl||!sc)return;
  let fs=(_tblZoom!=null&&isFinite(_tblZoom))?_tblZoom:tblAutoSize(tbl,sc);
  // ціле число пікселів — інакше гліфи лягають на субпіксельну сітку і «милять»
  fs=Math.max(TB_FIT_MIN,Math.min(TB_FIT_MAX,Math.round(fs)));
  tbl.style.fontSize=Math.round(fs)+'px';
  const lbl=document.getElementById('tblZoomLbl');
  if(lbl)lbl.textContent=(_tblZoom!=null?'':'авто ')+Math.round(fs)+'px';
  // другий рядок шапки має липнути рівно під першим
  // саме th.tb-g, бо перший th рядка — кутова «ДАТА» з rowspan=2
  const h1=tbl.querySelector('tr.tb-r1 th.tb-g');
  if(h1){
    const h=h1.getBoundingClientRect().height;
    tbl.querySelectorAll('tr.tb-r2 th').forEach(function(th){th.style.top=h+'px';});
  }
}
function tblZoom(step){
  const tbl=document.getElementById('dtTable');
  let cur=_tblZoom;
  if(cur==null&&tbl)cur=parseFloat(tbl.style.fontSize)||TB_FIT_BASE;
  _tblZoom=Math.max(TB_FIT_MIN,Math.min(TB_FIT_MAX,(cur||TB_FIT_BASE)+step));
  try{localStorage.setItem('tbl_zoom',_tblZoom);}catch(e){}
  tblFit();
}
// v60: горизонтальне гортання таблиці кнопками (dir: -1 вліво, 1 вправо)
function tblScroll(id,dir){
  const el=document.getElementById(id);
  if(!el)return;
  el.scrollBy({left:dir*400,behavior:'smooth'});
}
function tblZoomAuto(){
  _tblZoom=null;
  try{localStorage.removeItem('tbl_zoom');}catch(e){}
  tblFit();
}
window.addEventListener('resize',function(){
  clearTimeout(window._tbFitT);
  window._tbFitT=setTimeout(tblFit,150);
});

// Розрахунок діапазонів по колонках. Спільний для повної таблиці й міні-версії.
function tblCalc(d,n,LIM,S){
  if(!S)S=tblSpec();
  const N=d.d.length,mx=[],rg=[];
  for(let ci=0;ci<LIM;ci++){
    const sp=S[ci],col=d.c[ci];
    let m=0;
    if(sp.k==='chg'){for(let i=0;i<n;i++){const v=col[i];if(v!=null&&Math.abs(v)>m)m=Math.abs(v);}}
    mx.push(m||1);
    if(sp.k!=='net'&&sp.k!=='oi'&&sp.k!=='grad'&&sp.k!=='chg'&&sp.k!=='pct'){rg.push(null);continue;}
    const arr=[];
    for(let i=0;i<N;i++){const v=col[i];if(v!=null)arr.push(v);}
    if(!arr.length){rg.push(null);continue;}
    let mxa=0;for(let i=0;i<arr.length;i++){const a=Math.abs(arr[i]);if(a>mxa)mxa=a;}
    let asrt=null,p90=0;
    if(sp.k==='chg'||sp.k==='pct'){
      asrt=arr.map(function(x){return Math.abs(x);}).sort(function(a,b){return a-b;});
      p90=asrt[Math.floor(0.9*(asrt.length-1))]||mxa;
    }
    arr.sort(function(a,b){return a-b;});
    rg.push({mn:arr[0],mx:arr[arr.length-1],srt:arr,mxa:mxa||1,asrt:asrt,p90:p90});
  }
  return{mx:mx,rg:rg};
}
function tblRowsHtml(d,n,LIM,c,S){
  if(!S)S=tblSpec();
  const parts=[];
  for(let i=0;i<n;i++){
    parts.push('<tr><td class="tb-date">'+d.d[i]+'</td>');
    for(let ci=0;ci<LIM;ci++)
      parts.push(tblCell(S[ci].k,d.c[ci][i],c.mx[ci],S[ci].s,c.rg[ci],S[ci].c,S[ci].i,S[ci].n,S[ci].ct));
    parts.push('</tr>');
  }
  return parts.join('');
}
function tblRender(){
  const d=tblD()[_tblCur];
  const body=document.getElementById('dtBody');
  const stats=document.getElementById('dtStats');
  if(!d||!body)return;
  const S=tblSpec(),N=d.d.length,n=Math.min(_tblN,N);
  const c=tblCalc(d,n,S.length);
  stats.innerHTML=tblStatsRows(d,S.length);
  body.innerHTML=tblRowsHtml(d,n,S.length,c);

  const nm=document.getElementById('tblName');
  const mt=document.getElementById('tblMeta');
  if(nm)nm.textContent=tblNm()[_tblCur]||_tblCur;
  if(mt)mt.textContent='';
  tblHeadInfo(d);
  requestAnimationFrame(tblFit);
}

// ── Шапка: кнопка дашборда, SM DIV, CROWDED ATH ──
function tblColIdx(label){
  const S=tblSpec();
  for(let i=0;i<S.length;i++)if(S[i].l===label)return i;
  return -1;
}
function tblHeadInfo(d){
  const btn=document.getElementById('tblDashBtn');
  if(btn)btn.textContent=(tblNm()[_tblCur]||_tblCur)+' Dashboard';
  // у TFF немає колонок SM DIV / CROWDED ATH — ховаємо блок цілком
  const info=document.querySelector('.tb-info');
  if(info)info.style.display=(_tblSrc==='leg'?'':'none');
  const sm=document.getElementById('tblSmDiv'),cw=document.getElementById('tblCrowd');
  if(sm){
    const i=tblColIdx('SM DIV');
    const v=(i>=0&&d.c[i])?d.c[i][0]:null;
    if(v==null){sm.textContent='SM DIV —';sm.style.color='var(--d)';}
    else{
      const f=v/1000;
      sm.textContent='SM DIV '+(f>0?'+':'')+f.toFixed(2);
      sm.style.color=f>0?'#20d483':(f<0?'#f0515a':'#75809a');
    }
  }
  if(cw){
    const i=tblColIdx('CROWDED ATH');
    const t=(i>=0&&d.c[i])?d.c[i][0]:null;
    cw.textContent=t?t:'—';
    cw.className='tb-crowd '+(t?(String(t).toLowerCase().indexOf('very')>=0?'tb-cw-v':'tb-cw-c'):'tb-cw-n');
  }
}
function tblGoDash(){
  if(typeof ovGoInstrument==='function')ovGoInstrument(_tblCur);
}

// ── Міні-таблиця на вкладці інструмента (та сама верстка, до OPEN INTEREST) ──
const _miniN={};
function tblMiniRender(sid,n){
  const box=document.getElementById('mini_tbl_'+sid),d=_tbl[sid];
  if(!box||!d)return;
  const LIM=window._TBL_MINI_COLS||21,N=d.d.length;
  n=Math.min(n||10,N);
  _miniN[sid]=n;
  const SL=window._TBL_SPEC_L;
  const c=tblCalc(d,n,LIM,SL);
  box.innerHTML='<table class="dt">'+(window._TBL_MINI_THEAD||'')
    +'<tbody class="tb-stats">'+tblStatsRows(d,LIM,SL)+'</tbody>'
    +'<tbody>'+tblRowsHtml(d,n,LIM,c,SL)+'</tbody></table>';
  requestAnimationFrame(function(){tblFitEl(box.querySelector('table'),box,LIM+1);});
}
// ── Міні-таблиця TFF (LEV MONEY / ASSET MGR / DEALER / OPEN INTEREST) ──
const _miniNT={};
function tblMiniRenderT(sid,n){
  const box=document.getElementById('mini_tff_'+sid),d=(typeof _tblT!=='undefined')?_tblT[sid]:null;
  if(!box||!d)return;
  const ST=window._TBL_SPEC_T;
  if(!ST)return;
  const LIM=window._TBL_MINI_COLS_T||23,N=d.d.length;
  n=Math.min(n||10,N);
  _miniNT[sid]=n;
  const c=tblCalc(d,n,LIM,ST);
  box.innerHTML='<table class="dt">'+(window._TBL_MINI_THEAD_T||'')
    +'<tbody class="tb-stats">'+tblStatsRows(d,LIM,ST)+'</tbody>'
    +'<tbody>'+tblRowsHtml(d,n,LIM,c,ST)+'</tbody></table>';
  requestAnimationFrame(function(){tblFitEl(box.querySelector('table'),box,LIM+1);});
}
// ── Міні-таблиця Disaggregated ──
const _miniNG={};
function tblMiniRenderG(sid,n){
  const box=document.getElementById('mini_dg_'+sid),d=(typeof _tblG!=='undefined')?_tblG[sid]:null;
  if(!box||!d)return;
  const SG=window._TBL_SPEC_G;
  if(!SG)return;
  const LIM=window._TBL_MINI_COLS_G||23,N=d.d.length;
  n=Math.min(n||10,N);
  _miniNG[sid]=n;
  const c=tblCalc(d,n,LIM,SG);
  box.innerHTML='<table class="dt">'+(window._TBL_MINI_THEAD_G||'')
    +'<tbody class="tb-stats">'+tblStatsRows(d,LIM,SG)+'</tbody>'
    +'<tbody>'+tblRowsHtml(d,n,LIM,c,SG)+'</tbody></table>';
  requestAnimationFrame(function(){tblFitEl(box.querySelector('table'),box,LIM+1);});
}
function setMiniHistG(btn,sid){
  btn.parentNode.querySelectorAll('.hbtn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  tblMiniRenderG(sid,parseInt(btn.dataset.n));
}
function setMiniHistT(btn,sid){
  btn.parentNode.querySelectorAll('.hbtn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  tblMiniRenderT(sid,parseInt(btn.dataset.n));
}
function setMiniHist(btn,sid){
  btn.parentNode.querySelectorAll('.hbtn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  tblMiniRender(sid,parseInt(btn.dataset.n));
}
function tblFitEl(tbl,sc,cols){
  if(!tbl||!sc)return;
  tbl.style.fontSize=TB_FIT_BASE+'px';
  const row=tbl.querySelector('tbody:last-child tr');
  if(row){
    let w=0;const n=Math.min(cols,row.children.length);
    for(let i=0;i<n;i++)w+=row.children[i].getBoundingClientRect().width;
    const avail=sc.clientWidth-2;
    if(w>0&&avail>0){
      let fs=Math.round(TB_FIT_BASE*avail/w);
      fs=Math.max(TB_FIT_MIN,Math.min(TB_FIT_MAX,fs));
      tbl.style.fontSize=fs+'px';
    }
  }
  const h1=tbl.querySelector('tr.tb-r1 th.tb-g');
  if(h1){
    const h=h1.getBoundingClientRect().height;
    tbl.querySelectorAll('tr.tb-r2 th').forEach(function(th){th.style.top=h+'px';});
  }
}

function tblSel(sid,btn){
  _tblCur=sid;
  document.querySelectorAll('.tb-a').forEach(b=>b.classList.remove('active'));
  if(btn)btn.classList.add('active');
  else{const b=document.querySelector('.tb-a[data-sid="'+sid+'"]');if(b)b.classList.add('active');}
  tblRender();
  const sc=document.querySelector('.tb-scroll');if(sc)sc.scrollTop=0;
}
function tblPer(btn){
  document.querySelectorAll('.tb-per').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  _tblN=parseInt(btn.dataset.n);
  tblRender();
}
window.tblInit=function(){
  if(_tblCur)return;
  const sid=window._TBL_FIRST;
  if(sid&&tblD()[sid])tblSel(sid,null);
};
</script>
"""


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

HTML_HEAD = """<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=1440">
<title>COT Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://s3.tradingview.com/tv.js"></script>
<style>
:root{--bg:#1a1e2d;--bg2:#21263a;--bg3:#282f47;--bd:#343d5a;--g:#20d483;--r:#f0515a;--b:#4a9eff;--accent:#f59420;--t:#dde2ee;--d:#8090b0;--f:'Courier New',Courier,monospace;--hdr-h:50px;}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box;}
html,body{background:var(--bg);color:var(--t);font-family:var(--f);font-size:13px;}
.t{color:var(--t);}
.hdr{height:var(--hdr-h);padding:0 24px;background:var(--bg2);border-bottom:1px solid var(--bd);display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:300;}
.hdr-left{display:flex;flex-direction:column;gap:2px;}
.hdr-t{font-size:17px;font-weight:bold;color:#fff;letter-spacing:2px;}
.dash-b{color:#4a9eff;}.dash-g{color:#20d483;}.dash-r{color:#f0515a;}
.hdr-s{font-size:10px;color:var(--d);letter-spacing:1px;}
.hdr-center{position:absolute;left:50%;transform:translateX(-50%);text-align:center;pointer-events:none;}
.hdr-bc{font-size:11px;color:var(--d);}
.hdr-bc-pill{display:inline-flex;align-items:center;gap:2px;background:rgba(52,61,90,.75);border:1px solid var(--bd);border-radius:20px;padding:3px 12px;backdrop-filter:blur(4px);}
.hdr-bc span{color:#dde2ee;font-weight:bold;}
.hdr-bc .bc-sep{color:#4a5580;margin:0 5px;font-size:12px;}
.hdr-r{text-align:right;font-size:11px;color:var(--d);line-height:2;}.hdr-r b{color:var(--t);}
.tff-badge{display:inline-block;font-size:8px;padding:1px 5px;border-radius:2px;background:rgba(232,168,56,.2);color:#e8a838;border:1px solid #e8a83855;margin-left:6px;vertical-align:middle;}
.dg-badge{display:inline-block;font-size:8px;padding:1px 5px;border-radius:2px;background:rgba(167,139,250,.2);color:#a78bfa;border:1px solid #a78bfa55;margin-left:4px;vertical-align:middle;}
.auth-btn{padding:6px 16px;border:1px solid var(--g);border-radius:3px;background:transparent;color:var(--g);font-family:var(--f);font-size:11px;cursor:pointer;letter-spacing:1px;transition:all .15s;}
.auth-btn:hover{background:rgba(32,212,131,.15);}.auth-btn.logged{border-color:var(--b);color:var(--b);}
.sync-btn{padding:6px 10px;border:1px solid var(--bd);border-radius:3px;background:transparent;color:var(--d);font-family:var(--f);font-size:13px;cursor:pointer;transition:all .15s;}
.sync-btn:hover{border-color:var(--b);color:var(--b);}
.auth-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:1000;align-items:center;justify-content:center;}
.auth-overlay.open{display:flex;}
.auth-box{background:var(--bg2);border:1px solid var(--bd);border-radius:8px;padding:32px;width:320px;position:relative;}
.auth-close{position:absolute;top:12px;right:14px;background:none;border:none;color:var(--d);cursor:pointer;font-size:16px;font-family:var(--f);}
.auth-tabs{display:flex;margin-bottom:20px;border-bottom:1px solid var(--bd);}
.auth-tab{flex:1;padding:7px;text-align:center;cursor:pointer;font-size:11px;color:var(--d);letter-spacing:1px;border-bottom:2px solid transparent;margin-bottom:-1px;}
.auth-tab.active{color:#fff;border-bottom-color:var(--g);}
.auth-field{margin-bottom:12px;}.auth-field label{display:block;font-size:9px;color:var(--d);letter-spacing:.8px;margin-bottom:4px;}
.auth-field input{width:100%;background:var(--bg);border:1px solid var(--bd);border-radius:4px;padding:9px 11px;color:var(--t);font-family:var(--f);font-size:12px;outline:none;}
.auth-field input:focus{border-color:var(--g);}
.auth-submit{width:100%;padding:10px;background:var(--g);color:#000;border:none;border-radius:4px;cursor:pointer;font-family:var(--f);font-size:12px;font-weight:bold;margin-top:4px;}
.auth-msg{font-size:11px;padding:7px 10px;border-radius:4px;margin-top:10px;text-align:center;display:none;}
.auth-msg.err{background:rgba(240,81,90,.15);border:1px solid var(--r);color:var(--r);}
.auth-msg.ok{background:rgba(32,212,131,.15);border:1px solid var(--g);color:var(--g);}
.auth-user{font-size:11px;color:var(--d);text-align:center;margin-bottom:12px;padding:10px;background:var(--bg3);border-radius:4px;}.auth-user b{color:var(--t);}
.main-tabs{display:flex;gap:0;padding:0 24px;background:var(--bg2);border-bottom:2px solid var(--bd);}
.mtab{padding:10px 20px;border:none;background:transparent;color:#b0bcd4;font-family:var(--f);font-size:12px;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-2px;letter-spacing:.5px;}
.mtab:hover{color:var(--t);}.mtab.active{color:var(--accent);border-bottom-color:var(--accent);font-weight:bold;}
.main-sec{display:none;}.main-sec.active{display:block;}
.ctabs{display:flex;gap:5px;padding:8px 24px;background:var(--bg2);border-bottom:1px solid var(--bd);flex-wrap:wrap;}
.ctab{padding:5px 14px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#c0cce0;font-family:var(--f);font-size:12px;background:transparent;}
.ctab:hover{border-color:var(--accent);color:#fff;}.ctab.active{background:var(--accent);color:#000;border-color:var(--accent);font-weight:bold;}.tc{opacity:.5;font-size:9px;margin-left:3px;}
.catsec{display:none;}.catsec.active{display:block;}
.itabs{display:flex;gap:4px;padding:7px 24px;background:var(--bg);border-bottom:1px solid var(--bd);flex-wrap:wrap;}
.itab{padding:4px 12px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:11px;background:transparent;}
.itab:hover{border-color:var(--accent);color:#fff;}.itab.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);font-weight:bold;}
.iviews{padding:16px 24px;}.iview{display:none;}.iview.active{display:block;}
.report-tabs{display:flex;align-items:center;gap:6px;margin-bottom:14px;flex-wrap:wrap;}
.rtab-lbl{font-size:9px;color:var(--d);letter-spacing:1px;}
.rtab{padding:4px 12px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:11px;background:transparent;transition:all .15s;}
.rtab:hover:not(.disabled){border-color:var(--accent);color:#fff;}.rtab.active{background:var(--accent);color:#000;border-color:var(--accent);font-weight:bold;}.rtab.disabled{opacity:.3;cursor:not-allowed;}
.mcards{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px;}
.mc{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:10px 14px 12px;overflow:hidden;min-width:0;display:flex;flex-direction:column;}
.mc-lbl{font-size:9px;color:#fff;letter-spacing:.6px;margin-bottom:0;line-height:1.1;opacity:.85;}
.mc-inner{display:flex;align-items:flex-start;gap:8px;margin-top:0;}.mc-left{flex:1;min-width:0;}.mc-right{flex-shrink:0;}
.mc-val{font-size:clamp(18px,2.5vw,34px);font-weight:bold;line-height:1;margin-top:0;}
/* v59: WEEKLY ΔNet всередині картки — фіксовані 26 тижнів */
.mc-bar{height:130px;position:relative;margin-top:8px;border-top:1px solid var(--bd);padding-top:6px;}
.mc-chg-wrap{margin-top:6px;font-size:12px;}.mc-wtag{font-size:9px;color:var(--d);margin-left:3px;}
.mc-pct{font-size:10px;margin-top:2px;opacity:.85;}.mc-sub{font-size:10px;color:var(--d);margin-top:3px;}
.mid{display:grid;grid-template-columns:1fr 180px 1fr;gap:8px;margin-bottom:12px;}
/* v26 mid layout */
/* v55: без sm-panel — pct_combined на всю ширину */
.mid-nopanel{grid-template-columns:1fr;}
/* v58: TradingView на всю ширину зверху, під ним двоколонкова сітка
   (картки 2×2 ліворуч | analysis + перцентиль праворуч) */
/* v60: верх секції — сітка навпіл (сезонність | TradingView).
   Без даних сезонності .one-col робить графік на всю ширину, як було. */
.top-split{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px;}
.top-split.one-col{grid-template-columns:1fr;}
.season-box{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:12px 14px;overflow:auto;overflow-x:auto;display:flex;flex-direction:column;}
/* v61: заголовок + перемикач [Таблиця|Графік] в один рядок; панелі ділять решту висоти */
.sn-hdr{display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;margin-bottom:8px;}
.sn-pane{flex:1 1 auto;min-height:0;}
/* прокрутка живе у панелі таблиці — заголовок і перемикачі лишаються на місці */
.sn-pane-t{overflow-x:auto;min-width:0;}
/* v63: рядок вибору періоду середнього */
.sn-per{display:flex;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap;}
.sn-per-lbl{font-size:10px;color:var(--d);letter-spacing:.5px;}
/* v63: розгортання таблиці — у розгорнутому стані блок не росте, а скролиться */
.season-box.sn-open{max-height:520px;overflow-y:auto;}
.sn-more{margin-top:8px;padding:3px 10px;border:1px solid var(--bd);border-radius:3px;background:transparent;
         color:#b0bcd4;font-family:var(--f);font-size:10px;cursor:pointer;}
.sn-more:hover{border-color:var(--accent);color:#fff;}
.sn-chart{display:flex;flex-direction:column;min-height:300px;}
.sn-lines{display:flex;flex-wrap:wrap;gap:3px;margin-bottom:6px;}
.sn-cw{position:relative;flex:1 1 auto;min-height:240px;}
/* кнопка-тумблер лінії: кольорова крапка + рік; неактивна — приглушена */
.snl{display:inline-flex;align-items:center;gap:4px;padding:2px 7px;border:1px solid var(--bd);border-radius:3px;
     cursor:pointer;color:#8090b0;font-family:var(--f);font-size:9px;background:transparent;opacity:.55;}
.snl:hover{color:#fff;}
.snl.active{opacity:1;color:#dde2ee;background:var(--bg3);}
.snl-dot{width:9px;height:2px;border-radius:1px;flex-shrink:0;}
.tv-half{position:relative;min-height:470px;}
/* v60: таблиця сезонності — 13 компактних колонок, кольори як у Legacy-таблиці */
/* v62: крупніша й контрастніша, за зразком table.ht — колір несе фон, текст білий */
/* min-width 740px = 78px (Year) + 12×55px: 6 символів Courier@12px ("+0.35*") + падінги.
   Вужче за це — прокрутка у .sn-pane-t, замість обрізаних значень. */
table.sn{width:100%;min-width:740px;border-collapse:collapse;font-size:12px;table-layout:fixed;white-space:nowrap;}
table.sn th,table.sn td{padding:4px 6px;text-align:center;border-bottom:1px solid var(--bd);
                        border-right:1px solid rgba(128,144,176,.13);overflow:hidden;}
table.sn th{font-size:11px;font-weight:normal;color:var(--d);}
table.sn td{color:#fff;}
/* перша колонка: фон як у .date-col легасі-таблиці; довгі підписи переносяться, а не обрізаються */
table.sn .sn-y{width:78px;text-align:left;font-size:11px;color:#b0bcd4;background:var(--bg3);
               border-right:1px solid var(--bd);white-space:normal;word-break:break-word;line-height:1.25;}
table.sn td.sn-p{background:rgba(32,212,131,.22);}
table.sn td.sn-n{background:rgba(240,81,90,.22);}
table.sn td.sn-z,table.sn td.sn-na{color:var(--d);}
table.sn tr.sn-stat{background:rgba(128,144,176,.10);}
table.sn tr.sn-sep td{height:3px;padding:0;border-bottom:2px solid var(--bd);border-right:none;}
table.sn tr.sn-cur{background:rgba(245,148,32,.07);}
table.sn tr.sn-cur td.sn-y{border-left:2px solid var(--accent);color:#fff;}
/* v63: рік поза поточним пресетом — приглушений (видно, хто входить у середнє) */
table.sn tbody tr.sn-out{opacity:.55;}
/* v63: межа вибірки — акцентний бордер на рядку + підпис праворуч від номера року.
   Підпис абсолютний, щоб 78px колонка не розсувалась і сітка не їхала. */
table.sn tbody tr.sn-brd td.sn-y{border-left:2px solid var(--accent);color:#fff;}
table.sn td.sn-y{position:relative;}
.sn-mark{display:none;position:absolute;left:100%;top:50%;transform:translateY(-50%);
         margin-left:6px;white-space:nowrap;font-size:9px;line-height:1.4;padding:1px 6px;
         border:1px solid var(--accent);border-radius:3px;background:var(--bg3);
         color:var(--accent);pointer-events:none;z-index:3;}
table.sn tbody tr.sn-brd .sn-mark{display:inline-block;}
/* hover рядка — той самий патерн, що у table.ht (фон клітинок не чіпаємо) */
table.sn tbody tr:hover td{box-shadow:inset 0 1px 0 rgba(255,255,255,.25),inset 0 -1px 0 rgba(255,255,255,.25);}
table.sn tbody tr.sn-sep:hover td{box-shadow:none;}
.sn-note{font-size:10px;color:var(--d);margin-top:6px;}
.lg-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;align-items:stretch;}
.lg-cards{min-width:0;}
.lg-cards .mcards{grid-template-columns:repeat(2,1fr);margin-bottom:0;}
.lg-side{min-width:0;height:100%;display:flex;flex-direction:column;gap:10px;}
/* власні нижні відступи панелей всередині колонки не потрібні — інтервал дає gap */
.lg-side>.tff-analysis-panel{margin-bottom:0;}
/* v59: нижні межі колонок вирівнюються — остання панель правої колонки добирає різницю */
.lg-side>*:last-child{flex:1;}
.pct-combined{display:flex;flex-direction:column;gap:8px;min-height:0;}
.pct-combined>*:last-child{flex:1;}
.tv-box{position:absolute;inset:0;border:1px solid var(--bd);border-radius:5px;overflow:hidden;}
/* v57: заглушка для інструментів без відкритого символу TradingView */
.tv-empty{display:flex;align-items:center;justify-content:center;background:var(--bg2);}
.tv-empty-msg{text-align:center;color:var(--d);font-size:11px;line-height:1.8;}
.tv-empty-ico{font-size:28px;opacity:.4;margin-bottom:8px;}
.tv-empty-sub{font-size:9px;opacity:.7;margin-top:4px;}
.tv-pick-btn{margin-top:12px;padding:5px 14px;border:1px solid var(--bd);border-radius:3px;background:transparent;color:#b0bcd4;font-family:var(--f);font-size:10px;cursor:pointer;}
.tv-pick-btn:hover{border-color:var(--accent);color:#fff;}
.panel{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:12px 14px;}
.plbl{font-size:9px;color:#fff;letter-spacing:.5px;margin-bottom:10px;}
.arow{margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid var(--bd);}.arow:last-child{margin:0;padding:0;border:none;}
.arow-body{display:flex;gap:10px;align-items:center;}
.arow-left{flex:1;min-width:0;display:flex;flex-direction:column;gap:6px;}
.arow-right{width:140px;flex-shrink:0;display:flex;flex-direction:column;justify-content:center;align-items:flex-end;border-left:1px solid var(--bd);padding-left:10px;}
.ar-glbl{font-size:11px;font-weight:bold;letter-spacing:.8px;margin-bottom:3px;}
.ar-net{font-size:clamp(20px,2.2vw,30px);font-weight:bold;text-align:right;}
.arow-grid2{display:grid;grid-template-columns:1fr 1fr;gap:3px;}
.arow-dnet{display:flex;flex-direction:column;align-items:center;padding-top:6px;border-top:1px solid var(--bd);}
.arow-dnet .ag-lbl{font-size:10px;color:var(--d);margin-bottom:2px;}
.ag-val-net{font-size:clamp(19px,2.2vw,28px);font-weight:bold;line-height:1.2;}
.ag-item{display:flex;flex-direction:column;gap:3px;}
.ag-lbl{font-size:10px;color:#c0ccd8;letter-spacing:.4px;font-weight:bold;}
.ag-val{font-size:clamp(17px,2vw,26px);font-weight:bold;line-height:1.2;}
.ag-pct{font-size:10px;opacity:.75;font-weight:normal;}
.tff-analysis-panel{display:flex;flex-direction:column;background:var(--bg2);border:1px solid var(--bd);border-radius:5px;margin-bottom:12px;overflow:hidden;padding:0;}
/* v24 tff analysis css */
/* v26 gauge align */
.tff-a-row{display:flex;align-items:center;justify-content:flex-start;padding:12px 16px;border-bottom:1px solid var(--bd);gap:16px;}
.tff-a-row:last-child{border-bottom:none;}
.tff-a-left{flex:0 0 620px;max-width:620px;min-width:0;}
.tff-a-name{font-size:11px;font-weight:bold;letter-spacing:1px;margin-bottom:8px;}
.tff-a-metrics{display:flex;align-items:flex-end;gap:22px;flex-wrap:nowrap;}
.tff-ag-item{display:flex;flex-direction:column;gap:3px;min-width:110px;}
.tff-a-metrics .ag-lbl{font-size:9px;color:#c0ccd8;letter-spacing:.4px;font-weight:bold;}
.tff-a-metrics .ag-val{font-size:clamp(16px,1.8vw,22px);font-weight:bold;line-height:1.15;}
.tff-a-metrics .ag-bignet{font-size:clamp(18px,2vw,26px);}
.tff-a-metrics .ag-pct{font-size:10px;opacity:.75;font-weight:normal;}
.tff-a-gauges{display:flex;gap:10px;flex-shrink:0;align-items:center;border-left:1px solid var(--bd);padding-left:24px;}
.tff-a-gwrap{display:flex;flex-direction:column;align-items:center;}
.tff-mid{display:grid;grid-template-columns:1fr;gap:8px;margin-bottom:12px;}
.sm-panel{display:flex;flex-direction:column;justify-content:space-between;}
.sm-row{margin-bottom:8px;}.sm-lbl{font-size:9px;color:var(--d);margin-bottom:3px;}
.sm-bar-bg{background:var(--bg3);border-radius:10px;height:8px;position:relative;overflow:hidden;}
.sm-mk{position:absolute;top:1px;width:8px;height:6px;border-radius:3px;transform:translateX(-50%);}
.sm-val{font-size:12px;font-weight:bold;margin-top:3px;}
.sm-hint{font-size:8px;color:var(--d);margin-top:6px;line-height:1.5;border-top:1px solid var(--bd);padding-top:6px;}
.rpt-panel{display:flex;flex-direction:column;overflow:hidden;}
.plbl-row{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;}
.rel-legend{font-size:8px;color:var(--d);display:flex;align-items:center;gap:3px;}
.rel-icon{font-size:10px;margin-right:1px;line-height:1;}.rel-d{color:#20d483;}.rel-i{color:#f0a030;}.rel-n{color:#343d5a;}
.rpt-row{display:flex;align-items:center;padding:4px 0;border-bottom:1px solid var(--bd);gap:4px;}.rpt-row:last-child{border:none;padding-bottom:0;}
.rpt-dim{opacity:.35;}.rpt-rel{width:12px;flex-shrink:0;text-align:center;}.rpt-info{flex:1;min-width:0;}
.rpt-name{font-size:9px;color:var(--t);font-weight:bold;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.rpt-sched{font-size:8px;color:var(--d);margin-top:1px;}
.rpt-tag{display:inline-block;font-size:7px;padding:0 4px;border-radius:2px;background:rgba(74,158,255,.2);color:var(--b);margin-left:4px;vertical-align:middle;}
.rpt-btns{display:flex;gap:2px;flex-shrink:0;}
.rb{width:20px;height:20px;border-radius:3px;border:1px solid var(--bd);background:var(--bg3);color:var(--d);font-family:var(--f);font-size:9px;font-weight:bold;cursor:pointer;line-height:20px;text-align:center;padding:0;}
.rb:hover{border-color:var(--t);color:var(--t);}.rb-l.active{background:rgba(32,212,131,.25);border-color:var(--g);color:var(--g);}.rb-s.active{background:rgba(240,81,90,.20);border-color:var(--r);color:var(--r);}.rb-n.active{background:var(--bg3);border-color:var(--d);color:var(--d);}
.pct-sel-row{display:flex;gap:5px;align-items:center;margin-bottom:10px;flex-wrap:wrap;}.psel-group{display:flex;gap:3px;}.psel-sep{width:1px;height:16px;background:var(--bd);margin:0 3px;}
.psel,.pper,.psm,.ppm{padding:2px 8px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:10px;background:transparent;}
.psel:hover,.pper:hover,.psm:hover,.ppm:hover{border-color:var(--accent);color:#fff;}.psel.active,.pper.active,.psm.active,.ppm.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);}
.pct-val-row{margin-bottom:8px;}
.pbar-wrap{position:relative;margin-bottom:3px;}.pbar-bg{background:var(--bg3);border-radius:3px;height:18px;position:relative;overflow:hidden;}
.pbar-lo{position:absolute;left:0;top:0;height:100%;background:rgba(240,81,90,.3);width:15%;}.pbar-hi{position:absolute;right:0;top:0;height:100%;background:rgba(32,212,131,.3);width:15%;}
.ptick{position:absolute;top:0;width:2px;height:100%;background:rgba(255,255,255,.25);}
.pbar-mk{position:absolute;top:2px;width:4px;height:14px;background:var(--g);border-radius:2px;transform:translateX(-50%);transition:left .3s;}
.ptick-labels{position:relative;height:16px;margin-top:2px;}.ptlbl{position:absolute;transform:translateX(-50%);font-size:8px;color:var(--d);transition:left .3s;}.ptlbl-cur{color:var(--t);font-weight:bold;}
.pbar-lb{display:flex;justify-content:space-between;font-size:8px;color:var(--d);margin-top:12px;}
.chartbox{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:14px 16px;margin-bottom:12px;}
.chartbox-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:6px;}
.period-btns{display:flex;gap:3px;}.per-btn{padding:2px 9px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:10px;background:transparent;}
.per-btn:hover{border-color:var(--accent);color:#fff;}.per-btn.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);}
.chart-leg{display:flex;gap:10px;font-size:10px;color:var(--d);align-items:center;flex-wrap:wrap;}
.ll{display:inline-block;width:14px;height:2px;border-radius:1px;vertical-align:middle;margin-right:4px;}
.ll-dash{display:inline-block;width:14px;height:0;border-top:2px dashed;vertical-align:middle;}
.cw{height:140px;position:relative;}.bar-charts-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;}.bar-lbl{font-size:8px;letter-spacing:.5px;margin-bottom:6px;}.bar-cw{height:170px;position:relative;}
.htable-wrap{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;overflow:hidden;margin-bottom:12px;}
.htable-hdr{padding:7px 14px;border-bottom:1px solid var(--bd);font-size:10px;color:#fff;letter-spacing:.5px;display:flex;align-items:center;justify-content:space-between;}
.hsel{display:flex;gap:4px;}.hbtn{padding:2px 10px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:11px;background:transparent;}
.hbtn.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);}.htable-scroll{overflow-x:auto;}
table.ht{width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap;table-layout:fixed;}
table.ht th{padding:4px 8px;border-bottom:1px solid var(--bd);font-weight:normal;font-size:9px;letter-spacing:.5px;text-align:right;overflow:hidden;}
table.ht .th-corner{text-align:left;background:var(--bg3);}.th-date{text-align:left;background:var(--bg3);}.th-left{text-align:left;}.th-group{text-align:center;}.th-oi{text-align:center;}.sm-th{text-align:center;font-size:8px;color:var(--d);}.sm-th-group{font-size:8px;}
table.ht td{padding:4px 8px;border-bottom:1px solid var(--bg3);text-align:right;overflow:hidden;}
/* v27 hover */
/* виділення рядка БЕЗ зміни кольорів клітинок (фон не чіпаємо) */
table.ht tbody tr:hover td,
.ov-table tbody tr:hover td,
.cp-sum-table tbody tr:hover td,
table.cp-wt tbody tr:hover td{
  box-shadow:inset 0 1px 0 rgba(255,255,255,.25),inset 0 -1px 0 rgba(255,255,255,.25);
}
table.ht tbody tr:hover td:first-child,
.ov-table tbody tr:hover td:first-child,
.cp-sum-table tbody tr:hover td:first-child,
table.cp-wt tbody tr:hover td:first-child{
  box-shadow:inset 0 1px 0 rgba(255,255,255,.25),inset 0 -1px 0 rgba(255,255,255,.25),inset 3px 0 0 var(--accent);
}
/* v27 th align */
/* вища специфічність: правило 'table.ht th' перекривало .th-group */
table.ht th.th-group,table.ht th.th-oi,table.ht th.sm-th{text-align:center;}
table.ht th.th-corner,table.ht th.th-date,table.ht th.th-left{text-align:left;}
.tff-a-gauges{gap:8px!important;}

table.ht .date-col{text-align:left;color:var(--d);background:var(--bg3);}

table.ht .sep-r{border-right:1px solid var(--bd);}.sm-td{text-align:center;font-size:10px;padding:4px 6px;}
table.ht.ht-legacy tbody td:nth-child(2){border-left:2px solid rgba(74,158,255,.45);}
table.ht.ht-legacy tbody td:nth-child(7){border-left:2px solid rgba(32,212,131,.45);}
table.ht.ht-legacy tbody td:nth-child(12){border-left:2px solid rgba(240,81,90,.45);}
table.ht[id^="tff_tbl_"] tbody td:nth-child(2){border-left:2px solid rgba(74,158,255,.45);}
table.ht[id^="tff_tbl_"] tbody td:nth-child(5){border-left:2px solid rgba(240,180,41,.45);}
table.ht[id^="tff_tbl_"] tbody td:nth-child(8){border-left:2px solid rgba(32,212,131,.45);}
table.ht[id^="dg_tbl_"] tbody td:nth-child(2){border-left:2px solid rgba(167,139,250,.45);}
table.ht[id^="dg_tbl_"] tbody td:nth-child(5){border-left:2px solid rgba(32,212,131,.45);}
table.ht[id^="dg_tbl_"] tbody td:nth-child(8){border-left:2px solid rgba(240,180,41,.45);}
table.ht .pctc{font-size:10px;}
table.ht.ht-legacy td{color:#fff;}
table.ht.ht-legacy td.date-col,table.ht.ht-legacy .mm-lbl{color:var(--d);}
table.ht.ht-legacy tbody td{border-right:1px solid rgba(128,144,176,.13);}
table.ht.ht-legacy tbody td.sep-r,table.ht.ht-legacy tbody td.date-col{border-right:1px solid var(--bd);}
/* v22 tff/dg table — оформлення як у Legacy */
table.ht[id^="tff_tbl_"] td,table.ht[id^="dg_tbl_"] td{color:#fff;}
table.ht[id^="tff_tbl_"] td.date-col,table.ht[id^="dg_tbl_"] td.date-col,
table.ht[id^="tff_tbl_"] .mm-lbl,table.ht[id^="dg_tbl_"] .mm-lbl{color:var(--d);}
table.ht[id^="tff_tbl_"] tbody td,table.ht[id^="dg_tbl_"] tbody td{border-right:1px solid rgba(128,144,176,.13);}
table.ht[id^="tff_tbl_"] tbody td.sep-r,table.ht[id^="dg_tbl_"] tbody td.sep-r,
table.ht[id^="tff_tbl_"] tbody td.date-col,table.ht[id^="dg_tbl_"] tbody td.date-col{border-right:1px solid var(--bd);}

table.ht tbody.mm-tbody{border-bottom:3px solid var(--bd);}
table.ht tbody.mm-tbody td{background:var(--bg3);text-align:right;border-bottom:1px solid rgba(52,61,90,.8);padding:4px 8px;}
table.ht tbody.mm-tbody .mm-lbl{text-align:left;font-size:8px;color:var(--d);letter-spacing:.5px;font-weight:bold;}
table.ht tbody.mm-tbody .mm-val{text-align:right;font-size:10px;}
table.ht tbody.mm-tbody tr.mm-yr td{opacity:.78;}
.ov-meta{padding:10px 0 8px;font-size:11px;color:var(--d);}.ov-meta b{color:var(--t);}
.ov-scroll{overflow-x:auto;}.ov-table{width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap;}
.ov-table th{padding:6px 10px;background:var(--bg3);border-bottom:1px solid var(--bd);color:var(--d);font-weight:normal;font-size:9px;letter-spacing:.5px;text-align:center;}
.ov-table th:first-child{text-align:left;}.ov-table td{padding:5px 10px;border-bottom:1px solid var(--bg3);text-align:right;}
.ov-table .ov-asset{text-align:left;color:var(--t);font-weight:bold;}
.ov-table .ov-group td{background:var(--bg3);color:var(--d);font-size:8px;letter-spacing:1px;padding:4px 10px;text-align:left;}

.ov-cot-cell{display:flex;align-items:center;gap:6px;justify-content:flex-end;}
/* v21 overview css */
/* v49: без обмежувача — таблиця на всю ширину екрана */
#ms_ov>div{max-width:none;margin:0;}
/* v48: базовий кегль ставить JS (ovApplyZoom), відступи в em */
.ov-table{font-size:17px;}
.ov-table th{padding:.5em .62em;font-size:.68em;}
.ov-table td{padding:.42em .62em;}
.ov-idx{color:#5a6482;text-align:right;font-size:.72em;padding-right:.3em!important;}
.ov-idx-th{color:#5a6482;text-align:right;width:1px;white-space:nowrap;}
.ov-zoom{display:flex;align-items:center;gap:3px;margin-left:18px;}
.ov-zb{min-width:24px;padding:3px 9px;border:1px solid var(--bd);border-radius:3px;
  background:transparent;color:#b0bcd4;font-family:var(--f);font-size:10px;cursor:pointer;}
.ov-zb:hover{border-color:var(--accent);color:#fff;}
.ov-zl{font-size:10px;color:var(--d);min-width:40px;text-align:center;}
.ov-num{text-align:right;font-variant-numeric:tabular-nums;}
.ov-bg-g{background:rgba(32,212,131,.85);color:#0b0d12!important;font-weight:bold;border-radius:2px;}
.ov-bg-r{background:rgba(240,81,90,.85);color:#fff!important;font-weight:bold;border-radius:2px;}
.ov-bg-0{color:var(--d);}
.ov-bg-g span,.ov-bg-r span{color:inherit!important;}
.ov-per-row{display:flex;align-items:center;gap:5px;padding:4px 0 10px;flex-wrap:wrap;}
.ov-per-lbl{font-size:9px;color:var(--d);letter-spacing:.6px;margin-right:4px;}
.ov-per{padding:3px 11px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:10px;background:transparent;}
.ov-per:hover{border-color:var(--accent);color:#fff;}
.ov-per.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);font-weight:bold;}
.ov-crowd{font-size:9px;padding:2px 8px;border-radius:10px;font-weight:bold;white-space:nowrap;}
.ov-lead{font-size:9px;padding:2px 8px;border-radius:10px;font-weight:bold;white-space:nowrap;
  background:rgba(32,212,131,.15);border:1px solid #20d483;color:#20d483;}
.ov-crowd-c{background:rgba(240,180,41,.15);border:1px solid #f0b429;color:#f0b429;}
.ov-crowd-vc{background:rgba(240,81,90,.18);border:1px solid #f0515a;color:#f0515a;}
.ov-sm-cv-wrap{height:340px!important;}
/* v22 overview css */
.ov-cot-val{font-size:.8em!important;font-weight:bold;min-width:2.6em!important;}
.ov-cot-hi{color:#20d483!important;}
.ov-cot-lo{color:#f0515a!important;}
.ov-fav{cursor:pointer;color:#4a5580;margin-right:7px;font-size:13px;user-select:none;transition:color .15s,transform .15s;display:inline-block;}
.ov-fav:hover{color:#f0b429;transform:scale(1.2);}
.ov-fav.on{color:#f0b429;text-shadow:0 0 6px rgba(240,180,41,.6);}
/* v23 fav — підсвічується тільки зірка, рядок без змін */
/* v27 sort */
.ov-table th.ov-sortable{cursor:pointer;user-select:none;transition:color .15s;}
.ov-table th.ov-sortable:hover{color:var(--accent)!important;}
.ov-table th.ov-sort-desc::after{content:' ▼';font-size:7px;opacity:.85;}
.ov-table th.ov-sort-asc::after{content:' ▲';font-size:7px;opacity:.85;}
.ov-asset-link{cursor:pointer;transition:color .15s;}
.ov-asset-link:hover{color:var(--accent);text-decoration:underline;}


.mc-gauges{display:flex;flex-direction:row;gap:8px;flex-shrink:0;align-items:flex-start;}
.mc-gauge-wrap{display:flex;flex-direction:column;align-items:center;}
.ov-sm-chart-wrap{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;margin-top:16px;overflow:hidden;}
.ov-sm-tabs{display:flex;gap:0;padding:8px 16px 0;border-bottom:1px solid var(--bd);}
.ov-sm-tab{padding:6px 18px;border:none;background:transparent;color:#b0bcd4;font-family:var(--f);font-size:11px;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px;letter-spacing:.5px;}
.ov-sm-tab:hover{color:#fff;}.ov-sm-tab.active{color:var(--accent);border-bottom-color:var(--accent);font-weight:bold;}
.ov-sm-cv-wrap{padding:14px 16px 16px;height:220px;position:relative;}.ov-bar-bg{width:50px;height:5px;background:var(--bg3);border-radius:2px;overflow:hidden;flex-shrink:0;}.ov-bar-fill{height:100%;border-radius:2px;}.ov-cot-val{font-size:10px;color:var(--t);min-width:30px;text-align:right;}
.g{color:var(--g);}.r{color:var(--r);}.d{color:var(--d);}
.footer{text-align:center;padding:14px;color:var(--d);font-size:9px;letter-spacing:1px;border-top:1px solid var(--bd);margin-top:4px;}
.cp-wrap{padding:16px 24px;}
.cp-tabs{display:flex;gap:0;padding:0;border-bottom:2px solid var(--bd);margin-bottom:16px;}
.cp-tab{padding:8px 20px;border:none;background:transparent;color:#b0bcd4;font-family:var(--f);font-size:12px;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-2px;letter-spacing:.3px;}
.cp-tab:hover{color:#fff;}.cp-tab.active{color:#fff;font-weight:bold;}
.cp-panel{display:none;}.cp-panel.active{display:block;}
.cp-panel-hdr{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;padding:12px 16px;background:var(--bg2);border:1px solid var(--bd);border-radius:5px;}
.cp-stages-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px;margin-bottom:12px;}
.cp-stage-card{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:12px 14px;}
.cp-stage-card.cp-dim{opacity:.35;}
.cp-stage-name{font-size:9px;color:var(--d);letter-spacing:.5px;margin-bottom:6px;text-transform:uppercase;}
.cp-stage-cur{font-size:30px;font-weight:bold;line-height:1.1;margin-bottom:8px;}
.cp-stage-bars{margin-bottom:8px;}
.cp-bar-row{display:flex;align-items:center;gap:6px;margin-bottom:4px;}
.cp-bar-lbl{font-size:8px;color:var(--d);width:26px;flex-shrink:0;}
.cp-bar-bg{flex:1;height:6px;background:var(--bg3);border-radius:3px;overflow:hidden;}
.cp-bar-fill{height:100%;border-radius:3px;}
.cp-bar-val{font-size:9px;width:32px;text-align:right;flex-shrink:0;}
.cp-stage-meta{display:flex;align-items:center;gap:6px;flex-wrap:wrap;border-top:1px solid var(--bd);padding-top:6px;margin-top:2px;}
.cp-stage-lw,.cp-stage-ly{font-size:9px;color:var(--d);}
.cp-vs-badge{font-size:10px;padding:2px 8px;border-radius:12px;font-weight:bold;}
.cp-vs-badge.g{background:rgba(32,212,131,.15);border:1px solid #20d483;color:#20d483;}
.cp-vs-badge.r{background:rgba(240,81,90,.15);border:1px solid #f0515a;color:#f0515a;}
.cp-vs-badge.d{color:var(--d);font-size:9px;}
.cp-chart-wrap{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:14px 16px;margin-bottom:12px;}
.cp-chart-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;font-size:10px;color:#fff;letter-spacing:.5px;}
.cp-chart-leg{display:flex;gap:12px;font-size:10px;color:var(--d);align-items:center;}
.cp-summary-wrap{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;overflow:hidden;margin-top:4px;}
.cp-summary-hdr{padding:8px 14px;font-size:9px;color:#fff;letter-spacing:.5px;border-bottom:1px solid var(--bd);}
.cp-sum-table{width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap;}
.cp-sum-table th{padding:5px 10px;background:var(--bg3);font-size:9px;color:var(--d);font-weight:normal;text-align:right;letter-spacing:.3px;}
.cp-sum-table th:first-child{text-align:left;}
.cp-sum-table td{padding:5px 10px;border-bottom:1px solid var(--bg3);text-align:right;}

.cp-sum-group td{background:var(--bg3);font-weight:bold;font-size:10px;padding:6px 10px;text-align:left;border-left:3px solid transparent;}
.cp-sum-stage{color:var(--d);font-size:10px;text-align:left;}
.cp-mini-bar{display:inline-block;width:44px;height:4px;background:var(--bg3);border-radius:2px;overflow:hidden;margin-right:5px;vertical-align:middle;}
.cp-mini-bar div{height:100%;border-radius:2px;}
.cp-badge-wrap{display:inline-flex;align-items:center;justify-content:center;}
.cp-badge{display:inline-flex;align-items:center;justify-content:center;font-size:8px;padding:1px 6px;border-radius:10px;}
.cp-badge.g{background:rgba(32,212,131,.15);border:1px solid #20d48355;color:#20d483;}
.cp-badge.r{background:rgba(240,81,90,.15);border:1px solid #f0515a55;color:#f0515a;}
@media(max-width:640px){
  :root{--hdr-h:56px;}.hdr{padding:8px 12px;height:auto;min-height:var(--hdr-h);flex-wrap:wrap;}
  .mcards{grid-template-columns:1fr 1fr;gap:8px;}.mid{grid-template-columns:1fr 1fr;gap:8px;}
  .mid>div:first-child{grid-column:1/-1;}.tff-mid{grid-template-columns:1fr;}
  /* v58: на вузькому екрані сітка розпадається в одну колонку */
  .lg-grid{grid-template-columns:1fr;}
  /* v60: на вузькому екрані сезонність і графік ідуть один під одним */
  .top-split{grid-template-columns:1fr;}.tv-half{min-height:320px;}
  /* v24 mobile */
  .tff-a-row{flex-direction:column;align-items:stretch;gap:10px;}
  .tff-a-gauges{border-left:none;border-top:1px solid var(--bd);padding-left:0;padding-top:10px;justify-content:center;}
  .tff-a-metrics{gap:14px;}
  .bar-charts-grid{grid-template-columns:1fr;}.arow-right{width:110px;padding-left:8px;}
  table.ht{font-size:10px;}.auth-box{width:90vw;padding:24px 20px;}
}
</style>
</head>
<body>
<script>const _cd={};const _ci={};const _ci_m={};const _tff={};const _ci_tff={};const _dg={};const _ci_dg={};const _cpData={};const _season={};const SeasonCharts={};const Charts={};const BarChts={};const TffCharts={};const TffBarChts={};const DgCharts={};const DgBarChts={};</script>
"""

def generate_html(data, tff_data=None, disag_data=None, crop_data=None):
    if tff_data is None: tff_data={}
    # v21: enrich overview з cot_idx періодів (беремо з data по sid)
    # індекси ACCEL LS / ACCEL CM у payload таблиці (колонки BO / BP)
    try:
        _I_ACC_LS = next(i for i,(ci,_l,_k) in enumerate(TBL_COLS) if ci == 66)
        _I_ACC_CM = next(i for i,(ci,_l,_k) in enumerate(TBL_COLS) if ci == 67)
    except StopIteration:
        _I_ACC_LS = _I_ACC_CM = None
    for _it in OVERVIEW_TABLE:
        if isinstance(_it, dict):
            _d = data.get(_it.get('sid'))
            if _d and _d.get('cot_idx'):
                _it['cot_idx'] = _d['cot_idx']
            _tb = (_d or {}).get('table')
            if _tb and _I_ACC_LS is not None:
                def _last(ix):
                    try:
                        col = _tb['c'][ix]
                        return col[0] if col else None
                    except Exception:
                        return None
                _it['acc_ls'] = _last(_I_ACC_LS)
                _it['acc_cm'] = _last(_I_ACC_CM)
    if disag_data is None: disag_data={}
    all_dates=[d['cur']['date'] for d in data.values()]
    report_date=max(all_dates) if all_dates else '—'
    updated=datetime.now().strftime('%d.%m.%Y %H:%M')
    cat_tabs=[];cat_sects=[];first_cat=True
    for cat,instruments in CATEGORIES.items():
        available=[i for i in instruments if i in data]
        if not available: continue
        act=' active'if first_cat else''
        cat_tabs.append(f'<button class="ctab{act}" data-c="{cat}" onclick="selCat(\'{cat}\')">{cat}<span class="tc">({len(available)})</span></button>')
        inst_btns=''.join(f'<button class="itab" data-cat="{cat}" data-i="{i}" onclick="selInst(\'{cat}\',\'{i}\')">{disp(i)}</button>' for i in available)
        views=''.join(make_instrument_view(data[i], tff_data.get(i), disag_data.get(i)) for i in available)
        cat_sects.append(f'<div class="catsec{act}" id="cs_{cat}"><div class="itabs" id="itabs_{cat}">{inst_btns}</div><div class="iviews" id="iv_{cat}">{views}</div></div>')
        first_cat=False
    ov_html=make_overview_all()
    tbl_html=make_table_tab(data, tff_data, disag_data)
    db='<span class="dash-b">DAS</span><span class="dash-g">HBO</span><span class="dash-r">ARD</span>'
    badge=f' <span class="tff-badge">{len(tff_data)} TFF</span>' if tff_data else ''
    dg_badge=f' <span class="dg-badge">{len(disag_data)} DG</span>' if disag_data else ''
    return(HTML_HEAD
           +f'<header class="hdr">'
           +f'<div class="hdr-left"><div class="hdr-t">COT {db}</div>'
           +f'<div class="hdr-s">COMMITMENTS OF TRADERS — CFTC{badge}{dg_badge}</div></div>'
           +f'<div class="hdr-center"><div class="hdr-bc" id="hdrBreadcrumb"></div></div>'
           +f'<div style="display:flex;align-items:center;gap:16px">'
           +f'<div class="hdr-r">Звіт: <b>{report_date}</b><br>Оновлено: {updated}</div>'
           +f'<button class="sync-btn" onclick="openSyncModal()" title="Синхронізація налаштувань">⇄</button>'
           +f'<button class="auth-btn" id="authBtn" onclick="openAuth()">УВІЙТИ</button>'
           +f'</div></header>'
           +f'<div class="main-tabs">'
           +f'<button class="mtab" data-mt="ov" onclick="selMain(\'ov\')">Overview</button>'
           +f'<button class="mtab" data-mt="tbl" onclick="selMain(\'tbl\')">Table</button>'
           +f'<button class="mtab active" data-mt="cot" onclick="selMain(\'cot\')">COT Dashboard</button>'
           +f'<button class="mtab" data-mt="crop" onclick="selMain(\'crop\')">Crop Progress</button></div>'
           +f'<div class="main-sec" id="ms_ov"><div style="padding:16px 24px">{ov_html}</div></div>'
           +f'<div class="main-sec" id="ms_tbl">{tbl_html}</div>'
           +f'<div class="main-sec active" id="ms_cot"><div class="ctabs">{"".join(cat_tabs)}</div>{"".join(cat_sects)}</div>'
           +f'<div class="main-sec" id="ms_crop">{make_crop_tab(crop_data) if crop_data else ""}</div>'
           +f'<div class="footer">COT DASHBOARD &bull; CFTC LEGACY + TFF + DISAGGREGATED &bull; {updated}</div>'
           +AUTH_MODAL_HTML+OV2_CSS+OV2_JS+HTML_FOOT)

HTML_FOOT = """
<script>
const CL_LS='#4a9eff',CL_CM='#20d483',CL_ST='#f0515a';
const TFF_LEV='#e8a838',TFF_AM='#4a9eff',TFF_DL='#20d483';
const DG_MM='#a78bfa',DG_PM='#20d483',DG_SD='#f0b429';
const CL_OI='#a0aac0';   // v61: OPEN INTEREST — сірі бари тижневої зміни
const CurPer={};
let _loggedIn=false,_userEmail='';

function selCat(cat){
  document.querySelectorAll('.ctab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.catsec').forEach(s=>s.classList.remove('active'));
  document.querySelector('[data-c="'+cat+'"]').classList.add('active');
  document.getElementById('cs_'+cat).classList.add('active');
  _bcCat=cat; updateBreadcrumb();
  const first=document.querySelector('[data-cat="'+cat+'"]');
  if(first) selInst(cat,first.dataset.i);
}
/* ==================== TradingView ==================== */
/* Вихідний HTML заглушок, які користувач замінив віджетом через tvPick —
   потрібен, щоб tvDestroyAll могла повернути контейнер у початковий стан. */
const _tvEmptyHTML={};
function tvMount(box,symbol,allowChange){
  box.dataset.loaded='1';
  new TradingView.widget({
    symbol:symbol,
    theme:"dark",
    interval:"D",
    range:"12M",
    style:"1",
    locale:"uk",
    timezone:"Europe/Warsaw",
    toolbar_bg:"#1a1e2d",
    withdateranges:true,
    hide_side_toolbar:true,
    allow_symbol_change:allowChange,
    autosize:true,
    container_id:box.id
  });
}
/* Лінива ініціалізація: тільки контейнери з data-tvsym і ще не завантажені.
   Заглушки (.tv-empty) не мають data-tvsym, тому сюди не потрапляють. */
function tvInit(boxId){
  const b=document.getElementById(boxId);
  if(!b||!b.dataset.tvsym||b.dataset.loaded==='1')return;
  if(typeof TradingView==='undefined')return;
  tvMount(b,b.dataset.tvsym,false);
}
/* Кнопка "Вибрати символ" на заглушці: стартуємо з EURUSD, далі користувач
   змінює символ сам. Вибір НЕ зберігається між перезавантаженнями. */
function tvPick(boxId){
  const b=document.getElementById(boxId);
  if(!b||typeof TradingView==='undefined')return;
  if(_tvEmptyHTML[boxId]===undefined)_tvEmptyHTML[boxId]=b.innerHTML;
  b.classList.remove('tv-empty');
  b.innerHTML='';
  tvMount(b,"FX:EURUSD",true);
}
/* Знищуємо всі живі віджети, щоб не накопичувались iframe із WebSocket.
   Контейнер повертається у стан, з якого ініціалізація спрацює знову:
   звичайний — порожній з data-tvsym, створений через tvPick — назад у заглушку. */
function tvDestroyAll(){
  document.querySelectorAll('.tv-box').forEach(b=>{
    if(b.dataset.loaded!=='1')return;
    b.innerHTML='';
    delete b.dataset.loaded;
    if(!b.dataset.tvsym){
      if(_tvEmptyHTML[b.id]!==undefined)b.innerHTML=_tvEmptyHTML[b.id];
      b.classList.add('tv-empty');
    }
  });
}

function selInst(cat,key){
  tvDestroyAll();
  document.querySelectorAll('[data-cat="'+cat+'"]').forEach(b=>b.classList.remove('active'));
  const btn=document.querySelector('[data-cat="'+cat+'"][data-i="'+key+'"]');
  if(btn) btn.classList.add('active');
  document.getElementById('iv_'+cat).querySelectorAll('.iview').forEach(v=>v.classList.remove('active'));
  const s=key.replaceAll(' ','_').replaceAll('&','n').replaceAll('/','_');
  const view=document.getElementById('iv_'+s);
  if(view){
    view.classList.add('active');
    _bcInst=key; _bcRpt='Legacy Report'; updateBreadcrumb();
    switchReport(s,'legacy');
    loadRptStances(s);
  }
}
function selMain(mt){
  document.querySelectorAll('.mtab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.main-sec').forEach(s=>s.classList.remove('active'));
  document.querySelector('[data-mt="'+mt+'"]').classList.add('active');
  document.getElementById('ms_'+mt).classList.add('active');
  if(mt==='cot'){const fc=document.querySelector('.ctab');if(fc)selCat(fc.dataset.c);}
  if(mt==='ov'){if(window._ovSmInit)initOvSmChart(_ovSmKey||'div');if(window.ovApplyZoom)ovApplyZoom();}
  if(mt==='tbl'){if(window.tblInit)tblInit();}
  if(mt==='crop'){const f=document.querySelector('.cp-tab.active');if(f)f.click();}
}

function switchReport(sid,type){
  tvDestroyAll();
  const view=document.getElementById('iv_'+sid);if(!view)return;
  view.querySelectorAll('.rpt-sec').forEach(s=>s.style.display='none');
  const sec=document.getElementById('rpt_'+type+'_'+sid);
  if(sec) sec.style.display='';
  view.querySelectorAll('.rtab[data-rtype]').forEach(b=>b.classList.remove('active'));
  const activeTab=view.querySelector('.rtab[data-rtype="'+type+'"]');
  if(activeTab) activeTab.classList.add('active');
  _bcRpt=type==='tff'?'TFF Report':type==='dg'?'Disaggregated':'Legacy Report';
  updateBreadcrumb();
  const n=CurPer[sid]||52;
  if(type==='legacy'){
    filterRows(sid,10);
    if(typeof tblMiniRender==='function')
      tblMiniRender(sid,(typeof _miniN!=='undefined'&&_miniN[sid])||10);
    setTimeout(()=>{drawMainChart(sid,n);drawBarsFor(sid,n);drawCardBars(sid);},30);
    // v55: TradingView — ліниво, один раз на секцію (патерн crop_embed_)
    tvInit('tv_'+sid);
  } else if(type==='tff'){
    filterTffRows(sid,10);
    if(typeof tblMiniRenderT==='function')
      tblMiniRenderT(sid,(typeof _miniNT!=='undefined'&&_miniNT[sid])||10);
    setTimeout(()=>{drawTffChart(sid,n);drawTffBars(sid,n);drawTffCardBars(sid);},30);
    tvInit('tv_tff_'+sid);
  } else if(type==='dg'){
    filterDgRows(sid,10);
    if(typeof tblMiniRenderG==='function')
      tblMiniRenderG(sid,(typeof _miniNG!=='undefined'&&_miniNG[sid])||10);
    setTimeout(()=>{drawDgChart(sid,n);drawDgBars(sid,n);drawDgCardBars(sid);},30);
    tvInit('tv_dg_'+sid);
  } else if(type==='crop'){
    // Копіюємо Crop Progress контент прямо в iview
    const embed=document.getElementById('crop_embed_'+sid);
    if(embed&&!embed.dataset.loaded){
      const cid=({'CORN':'corn','WHEAT':'springwheat','SOYBEAN':'soybeans',
                  'SOYBEAN_MEAL':'soybeans','SOYBEAN_OIL':'soybeans',
                  'COTTON':'cotton','RICE':'rice'})[sid]||sid.toLowerCase();
      const src=document.getElementById('cp_'+cid);
      if(src){
        embed.innerHTML=src.innerHTML;
        embed.dataset.loaded='1';
        embed.dataset.cid=cid;
        // Ініціалізуємо графіки після рендеру
        setTimeout(()=>{
          if(window.selCrop)selCrop(cid);
        },100);
      }
    }
  }
}

// ── Legacy Charts ──
function setChartPer(btn,sid){
  btn.closest('.period-btns').querySelectorAll('.per-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const per=btn.dataset.per;const n=per==='1y'?52:per==='3y'?156:9999;
  CurPer[sid]=n;drawMainChart(sid,n);drawBarsFor(sid,n);
}
function drawMainChart(sid,nWeeks){
  const cv=document.getElementById('cv_'+sid);if(!cv)return;
  if(Charts[sid]){Charts[sid].destroy();delete Charts[sid];}
  const d=_cd[sid];if(!d)return;
  const n=Math.min(nWeeks,d.dates.length);
  Charts[sid]=new Chart(cv.getContext('2d'),{type:'line',data:{labels:d.dates.slice(-n),datasets:[
    {label:'Large Spec',data:d.ls.slice(-n),borderColor:CL_LS,backgroundColor:CL_LS+'22',borderWidth:1.5,pointRadius:0,tension:.3,fill:true},
    {label:'Commercials',data:d.cm.slice(-n),borderColor:CL_CM,backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3},
    {label:'Small Traders',data:d.st.slice(-n),borderColor:CL_ST,backgroundColor:'transparent',borderWidth:1,pointRadius:0,tension:.3,borderDash:[3,3]},
  ]},options:{responsive:true,maintainAspectRatio:false,animation:false,
    interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:false},tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',titleFont:{family:'Courier New',size:10},bodyFont:{family:'Courier New',size:10},callbacks:{label:ctx=>' '+ctx.dataset.label+': '+fmtFull(ctx.parsed.y)}}},
    scales:{x:{display:true,ticks:{color:'#8090b0',font:{family:'Courier New',size:8},maxTicksLimit:8,callback:function(v,i){return i%Math.ceil(n/8)===0?this.getLabelForValue(v):'';}},grid:{display:false},border:{display:false}},
      y:{display:true,grid:{color:'rgba(52,61,90,.8)',lineWidth:.5},ticks:{color:'#8090b0',font:{family:'Courier New',size:9},maxTicksLimit:4,callback:v=>fmtV(v,true)},border:{display:false}}}}});
}
function drawBarsFor(sid,nWeeks){
  const d=_cd[sid];if(!d)return;
  const n=Math.min(nWeeks,d.dates.length);const dates=d.dates.slice(-n);
  drawOneBar('barcv_ls_'+sid,dates,d.ld.slice(-n),CL_LS,'rgba(240,81,90,.75)');
  drawOneBar('barcv_cm_'+sid,dates,d.cd.slice(-n),CL_CM,'rgba(240,81,90,.75)');
  drawOneBar('barcv_st_'+sid,dates,d.sd.slice(-n),CL_ST,'rgba(15,18,28,.9)');
}
// v59: бари всередині карток — ЗАВЖДИ 26 тижнів, не залежать від перемикача періоду
function drawCardBars(sid){
  const d=_cd[sid];if(!d)return;
  const dates=d.dates.slice(-26);
  drawOneBar('mcbar_ls_'+sid,dates,d.ld.slice(-26),CL_LS,'rgba(240,81,90,.75)');
  drawOneBar('mcbar_cm_'+sid,dates,d.cd.slice(-26),CL_CM,'rgba(240,81,90,.75)');
  drawOneBar('mcbar_st_'+sid,dates,d.sd.slice(-26),CL_ST,'rgba(15,18,28,.9)');
  // v61: OPEN INTEREST — сірий бар зростання, червоний спад
  if(d.oid)drawOneBar('mcbar_oi_'+sid,dates,d.oid.slice(-26),CL_OI,'rgba(240,81,90,.75)');
}
function drawOneBar(cvId,dates,data,baseColor,negColor){
  const cv=document.getElementById(cvId);if(!cv)return;
  const key='b_'+cvId;if(BarChts[key]){BarChts[key].destroy();delete BarChts[key];}
  const colors=data.map(v=>v>=0?baseColor+'cc':negColor);
  BarChts[key]=new Chart(cv.getContext('2d'),{type:'bar',data:{labels:dates,datasets:[{data:data,backgroundColor:colors,borderWidth:0,borderRadius:1}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      plugins:{legend:{display:false},tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',titleFont:{family:'Courier New',size:9},bodyFont:{family:'Courier New',size:9},callbacks:{label:ctx=>fmtFull(ctx.parsed.y)}}},
      scales:{x:{display:false},y:{display:true,grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},ticks:{color:'#8090b0',font:{family:'Courier New',size:8},maxTicksLimit:3,callback:v=>fmtV(v,true)},border:{display:false}}}}});
}

// ── TFF Charts ──
function setTffChartPer(btn,sid){
  btn.closest('.period-btns').querySelectorAll('.per-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const per=btn.dataset.per;const n=per==='1y'?52:per==='3y'?156:9999;
  CurPer[sid]=n;drawTffChart(sid,n);drawTffBars(sid,n);
}
function drawTffChart(sid,nWeeks){
  const cv=document.getElementById('tff_cv_'+sid);if(!cv)return;
  const key='tff_'+sid;if(TffCharts[key]){TffCharts[key].destroy();delete TffCharts[key];}
  const d=_tff[sid];if(!d)return;
  const n=Math.min(nWeeks,d.dates.length);
  TffCharts[key]=new Chart(cv.getContext('2d'),{type:'line',data:{labels:d.dates.slice(-n),datasets:[
    {label:'Lev Money',data:d.lev.slice(-n),borderColor:TFF_LEV,backgroundColor:TFF_LEV+'22',borderWidth:1.5,pointRadius:0,tension:.3,fill:true},
    {label:'Asset Mgr',data:d.am.slice(-n), borderColor:TFF_AM, backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3},
    {label:'Dealer',   data:d.dl.slice(-n), borderColor:TFF_DL, backgroundColor:'transparent',borderWidth:1,  pointRadius:0,tension:.3,borderDash:[3,3]},
  ]},options:{responsive:true,maintainAspectRatio:false,animation:false,
    interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:false},tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',titleFont:{family:'Courier New',size:10},bodyFont:{family:'Courier New',size:10},callbacks:{label:ctx=>' '+ctx.dataset.label+': '+fmtFull(ctx.parsed.y)}}},
    scales:{x:{display:true,ticks:{color:'#8090b0',font:{family:'Courier New',size:8},maxTicksLimit:8,callback:function(v,i){return i%Math.ceil(n/8)===0?this.getLabelForValue(v):'';}},grid:{display:false},border:{display:false}},
      y:{display:true,grid:{color:'rgba(52,61,90,.8)',lineWidth:.5},ticks:{color:'#8090b0',font:{family:'Courier New',size:9},maxTicksLimit:4,callback:v=>fmtV(v,true)},border:{display:false}}}}});
}
function drawTffBars(sid,nWeeks){
  const d=_tff[sid];if(!d)return;
  const n=Math.min(nWeeks,d.dates.length);const dates=d.dates.slice(-n);
  drawOneTffBar('tff_barcv_lev_'+sid,dates,d.lev_d.slice(-n),TFF_LEV);
  drawOneTffBar('tff_barcv_am_'+sid, dates,d.am_d.slice(-n), TFF_AM);
  drawOneTffBar('tff_barcv_dl_'+sid, dates,d.dl_d.slice(-n), TFF_DL);
}
// v59: TFF бари всередині карток — ЗАВЖДИ 26 тижнів
function drawTffCardBars(sid){
  const d=_tff[sid];if(!d)return;
  const dates=d.dates.slice(-26);
  drawOneTffBar('mcbar_tff_lev_'+sid,dates,d.lev_d.slice(-26),TFF_LEV);
  drawOneTffBar('mcbar_tff_am_'+sid, dates,d.am_d.slice(-26), TFF_AM);
  drawOneTffBar('mcbar_tff_dl_'+sid, dates,d.dl_d.slice(-26), TFF_DL);
  // v61: OPEN INTEREST — сірий бар зростання, червоний спад
  if(d.oi_d)drawOneTffBar('mcbar_tff_oi_'+sid,dates,d.oi_d.slice(-26),CL_OI);
}
function drawOneTffBar(cvId,dates,data,color){
  const cv=document.getElementById(cvId);if(!cv)return;
  const key='tb_'+cvId;if(TffBarChts[key]){TffBarChts[key].destroy();delete TffBarChts[key];}
  const colors=data.map(v=>v>=0?color+'cc':'rgba(240,81,90,.75)');
  TffBarChts[key]=new Chart(cv.getContext('2d'),{type:'bar',data:{labels:dates,datasets:[{data:data,backgroundColor:colors,borderWidth:0,borderRadius:1}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      plugins:{legend:{display:false},tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',titleFont:{family:'Courier New',size:9},bodyFont:{family:'Courier New',size:9},callbacks:{label:ctx=>fmtFull(ctx.parsed.y)}}},
      scales:{x:{display:false},y:{display:true,grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},ticks:{color:'#8090b0',font:{family:'Courier New',size:8},maxTicksLimit:3,callback:v=>fmtV(v,true)},border:{display:false}}}}});
}

// ── DISAG Charts ──
function setDgChartPer(btn,sid){
  btn.closest('.period-btns').querySelectorAll('.per-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const per=btn.dataset.per;const n=per==='1y'?52:per==='3y'?156:9999;
  CurPer['dg_'+sid]=n;drawDgChart(sid,n);drawDgBars(sid,n);
}
function drawDgChart(sid,nWeeks){
  const cv=document.getElementById('dg_cv_'+sid);if(!cv)return;
  const key='dg_'+sid;if(DgCharts[key]){DgCharts[key].destroy();delete DgCharts[key];}
  const d=_dg[sid];if(!d)return;
  const n=Math.min(nWeeks,d.dates.length);
  DgCharts[key]=new Chart(cv.getContext('2d'),{type:'line',data:{labels:d.dates.slice(-n),datasets:[
    {label:'Man Money', data:d.mm.slice(-n),borderColor:DG_MM,backgroundColor:DG_MM+'22',borderWidth:1.5,pointRadius:0,tension:.3,fill:true},
    {label:'Prod/Merch',data:d.pm.slice(-n),borderColor:DG_PM,backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3},
    {label:'Swap Dlrs', data:d.sd.slice(-n),borderColor:DG_SD,backgroundColor:'transparent',borderWidth:1,  pointRadius:0,tension:.3,borderDash:[3,3]},
  ]},options:{responsive:true,maintainAspectRatio:false,animation:false,
    interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:false},tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',titleFont:{family:'Courier New',size:10},bodyFont:{family:'Courier New',size:10},callbacks:{label:ctx=>' '+ctx.dataset.label+': '+fmtFull(ctx.parsed.y)}}},
    scales:{x:{display:true,ticks:{color:'#8090b0',font:{family:'Courier New',size:8},maxTicksLimit:8,callback:function(v,i){return i%Math.ceil(n/8)===0?this.getLabelForValue(v):'';}},grid:{display:false},border:{display:false}},
      y:{display:true,grid:{color:'rgba(52,61,90,.8)',lineWidth:.5},ticks:{color:'#8090b0',font:{family:'Courier New',size:9},maxTicksLimit:4,callback:v=>fmtV(v,true)},border:{display:false}}}}});
}
function drawDgBars(sid,nWeeks){
  const d=_dg[sid];if(!d)return;
  const n=Math.min(nWeeks,d.dates.length);const dates=d.dates.slice(-n);
  drawOneDgBar('dg_barcv_mm_'+sid,dates,d.mm_d.slice(-n),DG_MM);
  drawOneDgBar('dg_barcv_pm_'+sid,dates,d.pm_d.slice(-n),DG_PM);
  drawOneDgBar('dg_barcv_sd_'+sid,dates,d.sd_d.slice(-n),DG_SD);
}
// v59: DISAG бари всередині карток — ЗАВЖДИ 26 тижнів
function drawDgCardBars(sid){
  const d=_dg[sid];if(!d)return;
  const dates=d.dates.slice(-26);
  drawOneDgBar('mcbar_dg_mm_'+sid,dates,d.mm_d.slice(-26),DG_MM);
  drawOneDgBar('mcbar_dg_pm_'+sid,dates,d.pm_d.slice(-26),DG_PM);
  drawOneDgBar('mcbar_dg_sd_'+sid,dates,d.sd_d.slice(-26),DG_SD);
  // v61: OPEN INTEREST — сірий бар зростання, червоний спад
  if(d.oi_d)drawOneDgBar('mcbar_dg_oi_'+sid,dates,d.oi_d.slice(-26),CL_OI);
}
function drawOneDgBar(cvId,dates,data,color){
  const cv=document.getElementById(cvId);if(!cv)return;
  const key='dg_'+cvId;if(DgBarChts[key]){DgBarChts[key].destroy();delete DgBarChts[key];}
  const colors=data.map(v=>v>=0?color+'cc':'rgba(240,81,90,.75)');
  DgBarChts[key]=new Chart(cv.getContext('2d'),{type:'bar',data:{labels:dates,datasets:[{data:data,backgroundColor:colors,borderWidth:0,borderRadius:1}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      plugins:{legend:{display:false},tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',titleFont:{family:'Courier New',size:9},bodyFont:{family:'Courier New',size:9},callbacks:{label:ctx=>fmtFull(ctx.parsed.y)}}},
      scales:{x:{display:false},y:{display:true,grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},ticks:{color:'#8090b0',font:{family:'Courier New',size:8},maxTicksLimit:3,callback:v=>fmtV(v,true)},border:{display:false}}}}});
}

// ── PCT bar (Legacy) ──
function pctSel(btn,sid){btn.closest('.psel-group').querySelectorAll('.psel').forEach(b=>b.classList.remove('active'));btn.classList.add('active');updatePctBar(sid);}
function pperSel(btn,sid){btn.closest('.psel-group').querySelectorAll('.pper').forEach(b=>b.classList.remove('active'));btn.classList.add('active');updatePctBar(sid);}
function updatePctBar(sid){
  const view=document.getElementById('iv_'+sid);if(!view)return;
  const p=view.querySelector('#rpt_legacy_'+sid+' .psel.active')?.dataset.p||'ls';
  const per=view.querySelector('#rpt_legacy_'+sid+' .pper.active')?.dataset.per||'all';
  const val=_ci[sid]?.[p]?.[per]??50;
  setPctBar(sid,'',val);
}
function setPctBar(sid,pfx,val){
  const pos=Math.min(Math.max(val,0),100);
  const col=val<15?'#f0515a':val>85?'#20d483':'#dde2ee';
  const lbl=val<15?'— екстрем. шорт':val>85?'— екстрем. лонг':'— нейтральна зона';
  const mk=document.getElementById(pfx+'pctmk_'+sid);
  const valEl=document.getElementById(pfx+'pctval_'+sid);
  const lblEl=document.getElementById(pfx+'pctlbl_'+sid);
  const curEl=document.getElementById(pfx+'pctcur_'+sid);
  if(mk)mk.style.left=pos+'%';
  if(valEl){valEl.style.color=col;valEl.textContent=val.toFixed(1)+'%';}
  if(lblEl)lblEl.textContent=lbl;
  if(curEl){curEl.style.left=pos+'%';curEl.textContent=val.toFixed(1)+'%';}
}

// ── PCT bar (COT INDEX Ranked M) ──
function pctMSel(btn,sid){btn.closest('.psel-group').querySelectorAll('.psm').forEach(b=>b.classList.remove('active'));btn.classList.add('active');updateMPctBar(sid);}
function pperMSel(btn,sid){btn.closest('.psel-group').querySelectorAll('.ppm').forEach(b=>b.classList.remove('active'));btn.classList.add('active');updateMPctBar(sid);}
function updateMPctBar(sid){
  const mk=document.getElementById('pctmmk_'+sid);if(!mk)return;
  const panel=mk.closest('.pct-panel');if(!panel)return;
  const p=panel.querySelector('.psm.active')?.dataset.p||'ls';
  const per=panel.querySelector('.ppm.active')?.dataset.per||'all';
  let val=_ci_m[sid]?.[p]?.[per];
  // ЗАХИСТ: null/undefined/нескінченність/поза діапазоном -> показуємо '—'
  if(val==null||!isFinite(val)||val<0||val>100){
    const vEl=document.getElementById('pctmval_'+sid);
    const lEl=document.getElementById('pctmcls_'+sid);
    const cEl=document.getElementById('pctmcur_'+sid);
    if(mk)mk.style.left='50%';
    if(vEl){vEl.style.color='#8090b0';vEl.textContent='—';}
    if(lEl)lEl.textContent='— немає даних';
    if(cEl){cEl.style.left='50%';cEl.textContent='—';}
    return;
  }
  const pos=Math.min(Math.max(val,0),100);
  const col=val<15?'#f0515a':val>85?'#20d483':'#dde2ee';
  const lbl=val<15?'— екстрем. шорт':val>85?'— екстрем. лонг':'— нейтральна зона';
  const vEl=document.getElementById('pctmval_'+sid);
  const lEl=document.getElementById('pctmcls_'+sid);
  const cEl=document.getElementById('pctmcur_'+sid);
  if(mk)mk.style.left=pos+'%';
  if(vEl){vEl.style.color=col;vEl.textContent=val!=null?val.toFixed(1)+'%':'—';}
  if(lEl)lEl.textContent=lbl;
  if(cEl){cEl.style.left=pos+'%';cEl.textContent=val!=null?val.toFixed(1)+'%':'—';}
}

// ── PCT bar (TFF) ──
function pctTffSel(btn,sid){btn.closest('.psel-group').querySelectorAll('.psel').forEach(b=>b.classList.remove('active'));btn.classList.add('active');updateTffPctBar(sid);}
function pperTffSel(btn,sid){btn.closest('.psel-group').querySelectorAll('.pper').forEach(b=>b.classList.remove('active'));btn.classList.add('active');updateTffPctBar(sid);}
function updateTffPctBar(sid){
  const sec=document.getElementById('rpt_tff_'+sid);if(!sec)return;
  const p=sec.querySelector('.psel.active')?.dataset.p||'am';
  const per=sec.querySelector('.pper.active')?.dataset.per||'all';
  const val=_ci_tff[sid]?.[p]?.[per]??50;
  setPctBar(sid,'tff_',val);
}

// ── PCT bar (DISAG) ──
function pctDgSel(btn,sid){btn.closest('.psel-group').querySelectorAll('.psel').forEach(b=>b.classList.remove('active'));btn.classList.add('active');updateDgPctBar(sid);}
function pperDgSel(btn,sid){btn.closest('.psel-group').querySelectorAll('.pper').forEach(b=>b.classList.remove('active'));btn.classList.add('active');updateDgPctBar(sid);}
function updateDgPctBar(sid){
  const sec=document.getElementById('rpt_dg_'+sid);if(!sec)return;
  const p=sec.querySelector('.psel.active')?.dataset.p||'mm';
  const per=sec.querySelector('.pper.active')?.dataset.per||'all';
  const val=_ci_dg[sid]?.[p]?.[per]??50;
  setPctBar(sid,'dg_',val);
}

// ── Таблиці ──
function setHist(btn,sid){
  const n=parseInt(btn.dataset.n);
  btn.closest('.htable-hdr').querySelectorAll('.hbtn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');filterRows(sid,n);
}
function filterRows(sid,n){
  const view=document.getElementById('iv_'+sid);if(!view)return;
  view.querySelectorAll('#rpt_legacy_'+sid+' tbody.data-tbody tr').forEach(tr=>{
    tr.style.display=parseInt(tr.dataset.row)<n?'':'none';
  });
}
function setTffHist(btn,sid){
  const n=parseInt(btn.dataset.n);
  btn.closest('.htable-hdr').querySelectorAll('.hbtn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');filterTffRows(sid,n);
}
function filterTffRows(sid,n){
  const sec=document.getElementById('rpt_tff_'+sid);if(!sec)return;
  sec.querySelectorAll('tbody.data-tbody tr').forEach(tr=>{
    tr.style.display=parseInt(tr.dataset.row)<n?'':'none';
  });
}
function setDgHist(btn,sid){
  const n=parseInt(btn.dataset.n);
  btn.closest('.htable-hdr').querySelectorAll('.hbtn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');filterDgRows(sid,n);
}
function filterDgRows(sid,n){
  const sec=document.getElementById('rpt_dg_'+sid);if(!sec)return;
  sec.querySelectorAll('tbody.data-tbody tr').forEach(tr=>{
    tr.style.display=parseInt(tr.dataset.row)<n?'':'none';
  });
}

// ── Reports stances ──
function setRpt(sid,rptId,stance){
  const rows=document.querySelectorAll('#rpt_'+sid+'_'+rptId);
  rows.forEach(row=>{
    row.querySelectorAll('.rb').forEach(b=>b.classList.remove('active'));
    const cls=stance==='long'?'.rb-l':stance==='short'?'.rb-s':'.rb-n';
    row.querySelector(cls)?.classList.add('active');
  });
  localStorage.setItem('rpt_'+sid+'_'+rptId,stance);
  if(_loggedIn){fetch('/api/rpt_stance',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({instrument:sid,report:rptId,stance})}).catch(()=>{});}
}
function loadRptStances(sid){
  const rptIds=['usda_crop','eia_petrol','usda_exp','cot_cftc','usda_wasde','usda_oil'];
  rptIds.forEach(rptId=>{
    const saved=localStorage.getItem('rpt_'+sid+'_'+rptId)||'neutral';
    applyRptStance(sid,rptId,saved);
  });
  if(_loggedIn){
    fetch('/api/rpt_stances?instrument='+sid).then(r=>r.json()).then(data=>{
      if(data&&data.stances){Object.entries(data.stances).forEach(([rptId,stance])=>{localStorage.setItem('rpt_'+sid+'_'+rptId,stance);applyRptStance(sid,rptId,stance);});}
    }).catch(()=>{});
  }
}
function applyRptStance(sid,rptId,stance){
  document.querySelectorAll('#rpt_'+sid+'_'+rptId).forEach(row=>{
    row.querySelectorAll('.rb').forEach(b=>b.classList.remove('active'));
    const cls=stance==='long'?'.rb-l':stance==='short'?'.rb-s':'.rb-n';
    row.querySelector(cls)?.classList.add('active');
  });
}
function loadAllStancesFromServer(){
  fetch('/api/rpt_stances/all').then(r=>r.json()).then(all=>{
    if(!all) return;
    Object.entries(all).forEach(([inst,reports])=>{Object.entries(reports).forEach(([rptId,stance])=>{localStorage.setItem('rpt_'+inst+'_'+rptId,stance);applyRptStance(inst,rptId,stance);});});
  }).catch(()=>{});
}

// ── Форматування ──
function fmtV(n,short=false){
  if(n===null||isNaN(n))return'—';n=Math.round(n);
  if(short){if(Math.abs(n)>=1e6)return(n/1e6).toFixed(1)+'M';if(Math.abs(n)>=1e3)return(n/1e3).toFixed(0)+'K';return''+n;}
  const sign=n>0?'+':n<0?'-':'';return sign+Math.abs(n).toLocaleString('uk-UA');
}
function fmtFull(n){
  if(n===null||isNaN(n))return'—';n=Math.round(n);
  const sign=n>0?'+':n<0?'-':'';return sign+Math.abs(n).toLocaleString('uk-UA');
}

// ── Auth ──
function openAuth(){document.getElementById('authOverlay').classList.add('open');}
function closeAuth(){document.getElementById('authOverlay').classList.remove('open');}
function authTab(t){
  document.querySelectorAll('.auth-tab').forEach((el,i)=>el.classList.toggle('active',(i===0&&t==='login')||(i===1&&t==='reg')));
  document.getElementById('at-login').style.display=t==='login'?'':'none';
  document.getElementById('at-reg').style.display=t==='reg'?'':'none';
  ['al-msg','ar-msg'].forEach(id=>{const el=document.getElementById(id);if(el)el.style.display='none';});
}
function showAuthMsg(id,text,isErr){const el=document.getElementById(id);el.textContent=text;el.className='auth-msg '+(isErr?'err':'ok');el.style.display='block';}
async function doAuthLogin(){
  const email=document.getElementById('al-email').value.trim();const pass=document.getElementById('al-pass').value;
  if(!email||!pass){showAuthMsg('al-msg','Заповніть всі поля',true);return;}
  try{const res=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email,password:pass})});const data=await res.json();
    if(data.ok){setLoggedIn(data.email);closeAuth();}else showAuthMsg('al-msg',data.error||'Помилка входу',true);
  }catch(e){showAuthMsg('al-msg','Сервер недоступний',true);}
}
async function doAuthReg(){
  const email=document.getElementById('ar-email').value.trim();const pass=document.getElementById('ar-pass').value;const pass2=document.getElementById('ar-pass2').value;
  if(!email||!pass){showAuthMsg('ar-msg','Заповніть всі поля',true);return;}
  if(pass!==pass2){showAuthMsg('ar-msg','Паролі не збігаються',true);return;}
  try{const res=await fetch('/api/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email,password:pass})});const data=await res.json();
    if(data.ok){showAuthMsg('ar-msg','Успішно! Тепер увійдіть.',false);setTimeout(()=>authTab('login'),1500);}
    else showAuthMsg('ar-msg',data.error||'Помилка реєстрації',true);
  }catch(e){showAuthMsg('ar-msg','Сервер недоступний',true);}
}
async function doLogout(){await fetch('/api/logout',{method:'POST'});setLoggedOut();closeAuth();}
function setLoggedIn(email){
  _loggedIn=true;_userEmail=email;
  const btn=document.getElementById('authBtn');
  if(btn){btn.textContent='● '+email.split('@')[0].toUpperCase().substring(0,8);btn.classList.add('logged');}
  document.getElementById('auth-logged').style.display='';
  document.getElementById('auth-loggedout').style.display='none';
  document.getElementById('auth-email-display').textContent=email;
  loadAllStancesFromServer();
}
function setLoggedOut(){
  _loggedIn=false;_userEmail='';
  const btn=document.getElementById('authBtn');
  if(btn){btn.textContent='УВІЙТИ';btn.classList.remove('logged');}
  document.getElementById('auth-logged').style.display='none';
  document.getElementById('auth-loggedout').style.display='';
}
fetch('/api/me').then(r=>r.json()).then(d=>{if(d.logged_in)setLoggedIn(d.email);}).catch(()=>{});

// ── Breadcrumb (по центру хедера) ──────────────────────────────
let _bcCat='',_bcInst='',_bcRpt='Legacy Report';
function updateBreadcrumb(){
  const el=document.getElementById('hdrBreadcrumb');if(!el)return;
  const parts=[
    _bcCat?`<span>${_bcCat}</span>`:'',
    _bcInst?`<span class="bc-sep">›</span><span>${_bcInst}</span>`:'',
    _bcRpt?`<span class="bc-sep">›</span><span>${_bcRpt}</span>`:'',
  ].filter(Boolean).join('');
  el.innerHTML=parts?`<span class="hdr-bc-pill">${parts}</span>`:'';
}

// ── Sync stances ──
function exportStances(){
  const data={};for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i);if(k&&k.startsWith('rpt_'))data[k]=localStorage.getItem(k);}
  return btoa(unescape(encodeURIComponent(JSON.stringify(data))));
}
function importStances(code){
  try{const data=JSON.parse(decodeURIComponent(escape(atob(code))));Object.entries(data).forEach(([k,v])=>{if(k.startsWith('rpt_'))localStorage.setItem(k,v);});return true;}
  catch(e){return false;}
}
function openSyncModal(){
  const code=exportStances();
  const modal=document.createElement('div');modal.id='syncModal';
  modal.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.8);z-index:2000;display:flex;align-items:center;justify-content:center;';
  modal.innerHTML=`<div style="background:#21263a;border:1px solid #343d5a;border-radius:8px;padding:28px;width:420px;max-width:95vw;font-family:Courier New,monospace;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;"><div style="font-size:11px;color:#fff;letter-spacing:1px;">СИНХРОНІЗАЦІЯ НАЛАШТУВАНЬ</div><button onclick="document.getElementById('syncModal').remove()" style="background:none;border:none;color:#8090b0;cursor:pointer;font-size:16px;">✕</button></div>
    <div style="font-size:9px;color:#8090b0;margin-bottom:6px;">КОД НАЛАШТУВАНЬ:</div>
    <textarea id="syncCode" style="width:100%;height:80px;background:#13162a;border:1px solid #343d5a;color:#20d483;font-family:Courier New,monospace;font-size:9px;padding:8px;border-radius:4px;resize:none;">${code}</textarea>
    <div style="display:flex;gap:8px;margin-top:12px;">
      <button onclick="navigator.clipboard.writeText(document.getElementById('syncCode').value).then(()=>{this.textContent='✓ Скопійовано!';setTimeout(()=>this.textContent='Копіювати',1500)})" style="flex:1;padding:8px;background:rgba(32,212,131,.15);border:1px solid #20d483;color:#20d483;font-family:Courier New,monospace;font-size:10px;border-radius:4px;cursor:pointer;">Копіювати</button>
      <button onclick="const c=document.getElementById('syncCode').value;if(importStances(c)){alert('✅ Завантажено!');document.getElementById('syncModal').remove();}else{alert('❌ Невірний код');}" style="flex:1;padding:8px;background:rgba(74,158,255,.15);border:1px solid #4a9eff;color:#4a9eff;font-family:Courier New,monospace;font-size:10px;border-radius:4px;cursor:pointer;">Завантажити</button>
    </div>
  </div>`;
  document.body.appendChild(modal);modal.addEventListener('click',e=>{if(e.target===modal)modal.remove();});
}
const firstCat=document.querySelector('.ctab');
if(firstCat)selCat(firstCat.dataset.c);

// ── v27: Overview сортування по колонках ──
let _ovSortCol=null,_ovSortDir=-1,_ovOrigRows=null;
function ovParseNum(txt){
  if(txt==null) return null;
  const s=String(txt);
  let out='';
  for(let i=0;i<s.length;i++){
    const c=s.charCodeAt(i);
    if(c===0x202f||c===0xa0||c===32||c===9||c===37) continue;
    if(c===0x2212){out+='-';continue;}
    if(c===44){out+='.';continue;}
    out+=s[i];
  }
  if(out===''||out==='-') return null;
  const f=out.charCodeAt(0);
  if(f===0x2014||f===0x2013) return null;
  const v=parseFloat(out);
  return isNaN(v)?null:v;
}
function ovCellVal(tr,col,type){
  const tds=tr.children;
  if(col>=tds.length) return null;
  const td=tds[col];
  if(type==='crowd'){
    const s=(td.textContent||'').toLowerCase();
    if(s.indexOf('very')>=0) return 2;
    if(s.indexOf('crowd')>=0) return 1;
    if(s.indexOf('yes')>=0) return 1;   // CM LEAD
    return 0;
  }
  if(type==='cot'){
    const el=td.querySelector('.ov-cot-val');
    return ovParseNum(el?el.textContent:td.textContent);
  }
  return ovParseNum(td.textContent);
}
function ovSort(th){
  const table=th.closest('table');
  const tbody=table?table.querySelector('tbody'):null;
  if(!tbody) return;
  if(!table._ovOrig) table._ovOrig=Array.from(tbody.children);
  const col=parseInt(th.dataset.col);
  const type=th.dataset.stype||'num';
  table.querySelectorAll('th.ov-sortable').forEach(h=>h.classList.remove('ov-sort-asc','ov-sort-desc'));
  if(type==='reset'){
    _ovSortCol=null;
    tbody.innerHTML='';
    table._ovOrig.forEach(r=>tbody.appendChild(r));
    if(window.ovLoadFavs)ovLoadFavs();
    if(window.ovRenumber)ovRenumber();
    return;
  }
  if(_ovSortCol===col){_ovSortDir=-_ovSortDir;} else {_ovSortCol=col;_ovSortDir=-1;}
  th.classList.add(_ovSortDir===-1?'ov-sort-desc':'ov-sort-asc');
  // групові рядки (ВАЛЮТИ/МЕТАЛИ/...) при сортуванні приховуємо
  const rows=table._ovOrig.filter(r=>!r.classList.contains('ov-group'));
  const dir=_ovSortDir;
  rows.sort((a,b)=>{
    const va=ovCellVal(a,col,type),vb=ovCellVal(b,col,type);
    if(va==null&&vb==null) return 0;
    if(va==null) return 1;
    if(vb==null) return -1;
    if(va===vb) return 0;
    return dir===-1?(vb-va):(va-vb);
  });
  tbody.innerHTML='';
  rows.forEach(r=>tbody.appendChild(r));
  if(window.ovLoadFavs)ovLoadFavs();
  if(window.ovRenumber)ovRenumber();
}

// ── v44: Overview -> вкладка TABLE цього активу ──
function ovGoTable(sid){
  if(typeof _tbl!=='undefined'&&_tbl[sid]&&typeof tblSel==='function'){
    selMain('tbl');
    tblSel(sid,null);
    window.scrollTo({top:0,behavior:'smooth'});
    return;
  }
  ovGoInstrument(sid);   // фолбек, якщо активу немає в таблиці
}

// ── v23: Overview -> перехід на вкладку інструмента ──
function ovGoInstrument(sid){
  // знаходимо кнопку інструмента серед усіх категорій
  let btn=null, cat=null;
  document.querySelectorAll('.itab[data-i]').forEach(b=>{
    if(btn) return;
    const k=b.dataset.i;
    const s=k.replaceAll(' ','_').replaceAll('&','n').replaceAll('/','_');
    if(s===sid){btn=b;cat=b.dataset.cat;}
  });
  if(!btn){console.warn('Інструмент не знайдено:',sid);return;}
  selMain('cot');           // перемикаємось на COT Dashboard
  selCat(cat);              // відкриваємо потрібну категорію
  selInst(cat,btn.dataset.i); // відкриваємо інструмент (Legacy за замовчуванням)
  window.scrollTo({top:0,behavior:'smooth'});
}

// ── v22: Overview favorites (зірочки) ──
function ovToggleFav(el){
  const sid=el.dataset.fav;
  const key='ovfav_'+sid;
  const on=localStorage.getItem(key)==='1';
  if(on){localStorage.removeItem(key);}else{localStorage.setItem(key,'1');}
  ovApplyFav(el,!on);
}
function ovApplyFav(el,on){
  el.classList.toggle('on',on);
  el.textContent=on?'★':'☆';
}
function ovLoadFavs(){
  document.querySelectorAll('.ov-fav').forEach(el=>{
    const on=localStorage.getItem('ovfav_'+el.dataset.fav)==='1';
    ovApplyFav(el,on);
  });
}
setTimeout(ovLoadFavs,60);

// ── v50: наскрізна нумерація рядків (не залежить від сортування) ──
function ovRenumber(){
  document.querySelectorAll('.ov-table').forEach(function(tb){
    let n=1;
    tb.querySelectorAll('tbody tr').forEach(function(tr){
      if(tr.classList.contains('ov-group'))return;
      const td=tr.querySelector('.ov-idx');
      if(td)td.textContent=n++;
    });
  });
}
setTimeout(ovRenumber,70);

// ── v48: масштаб таблиці Overview ──
let _ovZoom=null;
try{const _oz=localStorage.getItem('ov_zoom');if(_oz)_ovZoom=parseFloat(_oz);}catch(e){}
const OV_ZOOM_BASE=17, OV_ZOOM_MIN=9, OV_ZOOM_MAX=32;
function ovApplyZoom(){
  const fs=Math.round((_ovZoom!=null&&isFinite(_ovZoom))?_ovZoom:OV_ZOOM_BASE);
  document.querySelectorAll('.ov-table').forEach(function(t){t.style.fontSize=fs+'px';});
  const l=document.getElementById('ovZoomLbl');
  if(l)l.textContent=fs+'px';
}
function ovZoom(step){
  const cur=(_ovZoom!=null&&isFinite(_ovZoom))?_ovZoom:OV_ZOOM_BASE;
  _ovZoom=Math.max(OV_ZOOM_MIN,Math.min(OV_ZOOM_MAX,cur+step));
  try{localStorage.setItem('ov_zoom',_ovZoom);}catch(e){}
  ovApplyZoom();
}
function ovZoomReset(){
  _ovZoom=null;
  try{localStorage.removeItem('ov_zoom');}catch(e){}
  ovApplyZoom();
}
setTimeout(ovApplyZoom,80);

// ── v21: Overview COT period switcher ──
let _ovCotPer='all';
function ovSetPer(btn){
  document.querySelectorAll('.ov-per').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  _ovCotPer=btn.dataset.per;
  document.querySelectorAll('.ov-cot-cell-td').forEach(td=>{
    const raw=td.dataset[_ovCotPer];
    const cell=td.querySelector('.ov-cot-cell');
    if(!cell) return;
    if(raw===''||raw==null){
      cell.innerHTML='<span class="ov-cot-val d">—</span>';
      return;
    }
    const v=parseFloat(raw);
    const color=v<20?'#20d483':v>80?'#f0515a':'#4a9eff';
    const cls=v<20?'ov-cot-hi':v>80?'ov-cot-lo':'';
    const pct=Math.min(Math.max(v,0),100);
    cell.innerHTML='<div class="ov-bar-bg"><div class="ov-bar-fill" style="width:'+pct.toFixed(1)+'%;background:'+color+'"></div></div>'
      +'<span class="ov-cot-val '+cls+'">'+v.toFixed(0)+'%</span>';
  });
}

// ── v61: Сезонність — перемикач Таблиця/Графік + Chart.js line ──
// Перші дні місяців у невисокосному році (індекси 0..364) — підписи осі X.
const SN_MSTART=[0,31,59,90,120,151,181,212,243,273,304,334];
const SN_MNAME=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const SN_LBL=(()=>{const a=new Array(365).fill('');SN_MSTART.forEach((d,i)=>{a[d]=SN_MNAME[i];});return a;})();
function snToday(){
  const t=new Date();
  return Math.floor((Date.UTC(t.getFullYear(),t.getMonth(),t.getDate())-Date.UTC(t.getFullYear(),0,1))/864e5);
}
// ── v63: період середнього + розгортання таблиці ──
const SnPer={};   // key → активний пресет ('10y'); SnOpen — чи розгорнута таблиця
const SnOpen={};
function snPerN(pk){return parseInt(pk,10)||0;}
// Клітинки Probability/Average return — та сама розмітка, що й у Python (_sn_prob/_sn_ret)
function snProbCell(v){
  if(v==null)return '<td class="sn-na">--</td>';
  const c=v>60?'sn-p':(v<40?'sn-n':'sn-z');
  return '<td class="'+c+'">'+Math.round(v)+'%</td>';
}
function snRetCell(v){
  if(v==null)return '<td class="sn-na">--</td>';
  if(v>0)return '<td class="sn-p">+'+v.toFixed(2)+'</td>';
  if(v<0)return '<td class="sn-n">'+v.toFixed(2)+'</td>';
  return '<td class="sn-z">0.00</td>';
}
function snSetPeriod(btn,key,sid){
  const grp=document.getElementById('snper_'+key);
  if(grp)grp.querySelectorAll('.psel').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  SnPer[key]=btn.dataset.p;
  snApply(key,sid);
}
function snToggleAll(btn,key){
  SnOpen[key]=!SnOpen[key];
  btn.textContent=SnOpen[key]?'Згорнути ▲':'Показати всі роки ▼';
  const box=btn.closest('.season-box');
  if(box)box.classList.toggle('sn-open',SnOpen[key]);
  snApply(key,null);
}
// Єдина точка перемальовування таблиці: рядки статистики, заголовок, мітка межі,
// приглушення та видимість рядків. Графік чіпаємо тільки якщо він уже побудований.
function snApply(key,sid){
  const sd=_season[key.replace(/^(tff_|dg_)/,'')];if(!sd)return;
  const pk=SnPer[key]||sd.dflt;
  const p=sd.per[pk];if(!p)return;
  const open=!!SnOpen[key];
  const pr=document.getElementById('snprobr_'+key);
  const ar=document.getElementById('snavgr_'+key);
  if(pr)pr.innerHTML='<td class="sn-y">Probability %</td>'+p.prob.map(snProbCell).join('');
  if(ar)ar.innerHTML='<td class="sn-y">Average return%</td>'+p.avg.map(snRetCell).join('');
  const ttl=document.getElementById('sntitle_'+key);
  if(ttl)ttl.textContent='СЕЗОННІСТЬ ('+snPerN(pk)+' РОКІВ)';
  const tbl=document.getElementById('sntbl_'+key);
  if(tbl){
    tbl.querySelectorAll('tbody tr[data-y]').forEach(tr=>{
      const y=+tr.dataset.y;
      if(y===0){tr.style.display='';tr.classList.remove('sn-out','sn-brd');return;}  // поточний рік
      const inPre=y>=p.from;
      tr.style.display=(open||inPre)?'':'none';
      tr.classList.toggle('sn-out',!inPre);
      tr.classList.toggle('sn-brd',y===p.from);
    });
  }
  // лінія «Сер. NY» у графіку — якщо він уже намальований
  const ch=SeasonCharts[key];
  if(ch&&ch.data.datasets.length&&sd.pre&&sd.pre[pk]){
    const ds=ch.data.datasets[0];
    ds.data=sd.pre[pk];
    ds.label='Сер. '+snPerN(pk)+'Y';
    ch.update('none');
    const row=document.getElementById('snlines_'+key);
    const b0=row&&row.querySelector('.snl[data-di="0"]');
    if(b0)b0.innerHTML='<span class="snl-dot" style="background:'+b0.dataset.col+'"></span>'+ds.label;
  }
}
function seasonView(btn,key,sid,mode){
  const grp=btn.closest('.sn-sel');
  if(grp)grp.querySelectorAll('.psel').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const tp=document.getElementById('snpane_tbl_'+key);
  const cp=document.getElementById('snpane_chart_'+key);
  if(tp)tp.style.display=(mode==='tbl')?'':'none';
  if(cp){
    cp.style.display=(mode==='chart')?'':'none';
    // лінива ініціалізація: графік будується тільки при першому показі
    if(mode==='chart'&&!cp.dataset.loaded){
      cp.dataset.loaded='1';
      setTimeout(()=>drawSeasonChart(key,sid),30);
    }
  }
}
// Палітра тонких ліній окремих років — по колу; у лінію йде з суфіксом '99' (напівпрозоро)
const SN_YCOL=['#20d483','#f0b429','#a78bfa','#22d3ee','#f59420','#e8a838','#8090b0','#4a9eff','#f0515a','#20d483'];
// Кнопки-тумблери ліній: порядок збігається з порядком datasets (обидва будуються тут).
// Активна кнопка підсвічена кольором своєї лінії (без прозорості).
function snBuildBtns(key,ds){
  const row=document.getElementById('snlines_'+key);if(!row)return;
  row.innerHTML='';
  ds.forEach((s,i)=>{
    const col=s._snCol||s.borderColor;
    const b=document.createElement('button');
    b.className='snl'+(s.hidden?'':' active');
    b.dataset.di=i;
    b.dataset.col=col;
    b.onclick=function(){seasonLine(this,key);};
    b.innerHTML='<span class="snl-dot" style="background:'+col+'"></span>'+s.label;
    snBtnPaint(b,!s.hidden);
    row.appendChild(b);
  });
}
function snBtnPaint(btn,on){
  const col=btn.dataset.col||'#8090b0';
  btn.style.borderColor=on?col:'';
  btn.style.color=on?col:'';
}
// Перемикання однієї лінії — без повного перемальовування ('none' = без анімації)
function seasonLine(btn,key){
  const ch=SeasonCharts[key];if(!ch)return;
  const s=ch.data.datasets[+btn.dataset.di];if(!s)return;
  s.hidden=!s.hidden;
  btn.classList.toggle('active',!s.hidden);
  snBtnPaint(btn,!s.hidden);
  ch.update('none');
}
function drawSeasonChart(key,sid){
  const cv=document.getElementById('sncv_'+key);if(!cv)return;
  const d=_season[sid];if(!d)return;
  if(SeasonCharts[key]){SeasonCharts[key].destroy();delete SeasonCharts[key];}
  const pk=SnPer[key]||d.dflt;
  const p10=(d.pre&&d.pre[pk])||null;   // крива середнього за активний період
  const cur=d.cur||null;
  if(!p10&&!cur)return;
  const today=snToday();
  // Порядок: Сер. 10Y → поточний рік → решта років (новіші зверху).
  // Увімкнені за замовчуванням лише перші дві лінії.
  const ds=[];
  if(p10)ds.push({label:'Сер. '+snPerN(pk)+'Y',data:p10,borderColor:'#4a9eff',backgroundColor:'transparent',
                  borderWidth:2.5,pointRadius:0,tension:.2,spanGaps:true,hidden:false,_snCol:'#4a9eff'});
  if(cur)ds.push({label:(d.cur_year||'')+'',data:cur,borderColor:'#f0515a',backgroundColor:'transparent',
                  borderWidth:2,pointRadius:0,tension:.2,spanGaps:false,hidden:false,_snCol:'#f0515a'});
  // _snCol — «чистий» колір для кнопки; у лінії він же, але напівпрозорий ('99')
  const yrs=Object.keys(d.years||{}).filter(y=>y!==String(d.cur_year)).sort((a,b)=>b-a);
  yrs.forEach((y,i)=>{
    const col=SN_YCOL[i%SN_YCOL.length];
    ds.push({label:y,data:d.years[y],borderColor:col+'99',backgroundColor:'transparent',
             borderWidth:1,pointRadius:0,tension:.2,spanGaps:true,hidden:true,_snCol:col});
  });
  // нульова горизонталь + пунктирна вертикаль на сьогоднішньому дні (патерн smThresholdLines)
  const snGuides={
    id:'snGuides',
    beforeDatasetsDraw(chart){
      const {ctx,chartArea,scales}=chart;
      if(!chartArea||!scales.y||!scales.x)return;
      ctx.save();
      ctx.lineWidth=1;
      ctx.setLineDash([]);
      ctx.strokeStyle='rgba(221,226,238,.28)';
      const y0=scales.y.getPixelForValue(0);
      ctx.beginPath();ctx.moveTo(chartArea.left,y0);ctx.lineTo(chartArea.right,y0);ctx.stroke();
      if(today>=0&&today<365){
        ctx.setLineDash([4,4]);
        ctx.strokeStyle='rgba(221,226,238,.35)';
        const x=scales.x.getPixelForValue(today);
        ctx.beginPath();ctx.moveTo(x,chartArea.top);ctx.lineTo(x,chartArea.bottom);ctx.stroke();
      }
      ctx.restore();
    }
  };
  SeasonCharts[key]=new Chart(cv.getContext('2d'),{
    type:'line',
    data:{labels:SN_LBL.map((_,i)=>i),datasets:ds},
    plugins:[snGuides],
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      interaction:{mode:'index',intersect:false},
      plugins:{legend:{display:false},
        tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',
          titleFont:{family:'Courier New',size:10},bodyFont:{family:'Courier New',size:10},
          // тільки видимі лінії: вимкнені кнопкою у тултип не потрапляють
          filter:item=>item.chart.isDatasetVisible(item.datasetIndex),
          callbacks:{title:items=>{const i=items[0].dataIndex;
                       let m=0;for(let k=0;k<12;k++){if(i>=SN_MSTART[k])m=k;}
                       return SN_MNAME[m]+' '+(i-SN_MSTART[m]+1);},
                     label:ctx=>' '+ctx.dataset.label+': '+(ctx.parsed.y==null?'--':ctx.parsed.y.toFixed(2)+'%')}}},
      scales:{
        x:{ticks:{color:'#8090b0',font:{family:'Courier New',size:8},autoSkip:false,maxRotation:0,
                  callback:(v,i)=>SN_LBL[i]||''},
           grid:{display:false},border:{display:false}},
        y:{grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},
           ticks:{color:'#8090b0',font:{family:'Courier New',size:9},callback:v=>v.toFixed(1)+'%'},
           border:{display:false}}}}});
  snBuildBtns(key,ds);
}

// ── Overview SM DIV bar chart ──
let _ovSmChart=null;let _ovSmKey='div';
function selSmTab(btn,key){
  document.querySelectorAll('.ov-sm-tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  _ovSmKey=key;
  initOvSmChart(key);
}
function initOvSmChart(key){
  const cv=document.getElementById('ovSmChart');if(!cv)return;
  if(_ovSmChart){_ovSmChart.destroy();_ovSmChart=null;}
  const src=window._ovSmInit||[];const labels=src.map(d=>d.label);
  const vals=src.map(d=>d[key]||0);
  const colors=vals.map(v=>v>=0?'rgba(32,212,131,.75)':'rgba(240,81,90,.75)');
  const bdrColors=vals.map(v=>v>=0?'#20d483':'#f0515a');
  const titles={'div':'SM DIV — All Time','div_6m':'SM DIV — 6 Months','div_3m':'SM DIV — 3 Months'};
  const _SM_THRESHOLD=0.8;
  // підпис активу: зелений якщо >0.8, червоний якщо <-0.8
  const tickColors=vals.map(v=>v>_SM_THRESHOLD?'#20d483':(v<-_SM_THRESHOLD?'#f0515a':'#8090b0'));
  // ледь помітні горизонтальні лінії на ±0.8
  const thresholdLines={
    id:'smThresholdLines',
    beforeDatasetsDraw(chart){
      const {ctx,chartArea,scales}=chart;
      if(!chartArea||!scales.y) return;
      ctx.save();
      ctx.strokeStyle='rgba(221,226,238,.13)';
      ctx.lineWidth=1;
      ctx.setLineDash([4,4]);
      [_SM_THRESHOLD,-_SM_THRESHOLD].forEach(lv=>{
        const y=scales.y.getPixelForValue(lv);
        ctx.beginPath();
        ctx.moveTo(chartArea.left,y);
        ctx.lineTo(chartArea.right,y);
        ctx.stroke();
      });
      ctx.restore();
    }
  };
  _ovSmChart=new Chart(cv.getContext('2d'),{
    type:'bar',
    data:{labels,datasets:[{data:vals,backgroundColor:colors,borderColor:bdrColors,borderWidth:1,borderRadius:2,label:titles[key]||'SM DIV'}]},
    plugins:[thresholdLines],
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      plugins:{legend:{display:false},
        title:{display:true,text:titles[key]||'SM DIV',color:'#dde2ee',font:{family:'Courier New',size:11},padding:{bottom:8}},
        tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',
          titleFont:{family:'Courier New',size:10},bodyFont:{family:'Courier New',size:10},
          callbacks:{label:ctx=>{const v=ctx.parsed.y;return ' '+v.toFixed(2);}}}},
      scales:{
        x:{ticks:{color:tickColors,font:{family:'Courier New',size:9},maxRotation:45,minRotation:30},
           grid:{color:'rgba(52,61,90,.5)',lineWidth:.5},border:{display:false}},
        y:{display:true,grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},
           ticks:{color:'#8090b0',font:{family:'Courier New',size:9},callback:v=>v.toFixed(2)},
           border:{display:false},
           min:-1,max:1}}}});
}
</script>
</body>
</html>
"""


def load_crop_data():
    """Читає ALL_Crops_Dashboard.xlsx і повертає dict з даними."""
    if not CROP_FILE.exists():
        print(f"  ⚠  Crop файл не знайдено: {CROP_FILE}"); return {}
    print(f"\n📂  {CROP_FILE}")
    xl = pd.ExcelFile(CROP_FILE)

    # ── 1. All Crops summary ──
    summary = []
    try:
        df = xl.parse('All Crops', header=None)
        for i in range(8, 38):
            row = df.iloc[i]
            crop  = str(row.iloc[0]).strip()
            stage = str(row.iloc[1]).strip().replace('\n',' ')
            cur   = str(row.iloc[2]).strip()
            avg   = str(row.iloc[3]).strip()
            vs    = str(row.iloc[4]).strip()
            sym   = str(row.iloc[5]).strip() if df.shape[1]>5 else ''
            lw    = str(row.iloc[6]).strip() if df.shape[1]>6 else ''
            ly    = str(row.iloc[7]).strip() if df.shape[1]>7 else ''
            if not crop or crop=='nan': continue
            try: float(cur)
            except:
                if cur not in('—',): continue
            summary.append({'crop':crop,'stage':stage,'cur':cur,'avg':avg,'vs':vs,'sym':sym,'lw':lw,'ly':ly})
    except Exception as e: print(f"  ⚠  All Crops: {e}")

    # ── 2. Historical per-crop ──
    hist = {}
    for sheet, meta in CROP_META.items():
        if sheet not in xl.sheet_names: continue
        try:
            cdf = xl.parse(sheet, header=None)
            rows = []
            for ri in range(5, min(40, len(cdf))):
                week = str(cdf.iloc[ri, 0]).strip()
                if not week or week=='nan': continue
                entry = {'week': week}
                for stage_name, c26, cavg, c25 in meta['stages']:
                    try:
                        v26  = cdf.iloc[ri, c26];  vavg = cdf.iloc[ri, cavg]; v25 = cdf.iloc[ri, c25]
                        to_f = lambda x: (float(x) if str(x).strip() not in('nan','—','') else -1)
                        v26f  = to_f(v26); vavgf = to_f(vavg); v25f = to_f(v25)
                        if v26f >= 0 or vavgf >= 0:
                            entry[stage_name] = {'cur': v26f if v26f>=0 else None,
                                                 'avg': vavgf if vavgf>=0 else None,
                                                 'y25': v25f if v25f>=0 else None}
                    except: pass
                rows.append(entry)
            hist[sheet] = rows
        except Exception as e: print(f"  ⚠  {sheet}: {e}")

    result = {'summary': summary, 'hist': hist}
    crops_loaded = len([r for r in summary if r['cur'] not in('0','—','',)])
    print(f"  ✓  Crop Progress: {len(summary)} рядків, {crops_loaded} з даними\n")
    return result



# ================================================================
# CROP PROGRESS UPDATED — make_crop_tab v2

def make_crop_gauge(cur, avg, color, size=64):
    """Gauge з двома дугами: основна (cur) + прозора (avg 5yr)."""
    import math as _m
    cur  = max(0.0, min(100.0, float(cur  or 0)))
    avg  = max(0.0, min(100.0, float(avg  or 0)))
    cx = cy = size / 2
    r_outer = size * 0.40
    r_inner = size * 0.28
    START_DEG = 140.0
    SWEEP     = 240.0
    def pt(r, deg):
        a = _m.radians(deg)
        return round(cx + r * _m.cos(a), 2), round(cy + r * _m.sin(a), 2)
    # outer arc (track)
    s_o  = pt(r_outer, START_DEG)
    e_o  = pt(r_outer, START_DEG + SWEEP)
    # outer fill (2026)
    vs   = cur / 100.0 * SWEEP
    v_o  = pt(r_outer, START_DEG + vs)
    fg_o = f"M{s_o[0]},{s_o[1]} A{r_outer:.1f},{r_outer:.1f} 0 {1 if vs>180 else 0},1 {v_o[0]},{v_o[1]}" if cur>0 else None
    # inner arc (5yr avg) — lighter
    s_i  = pt(r_inner, START_DEG)
    e_i  = pt(r_inner, START_DEG + SWEEP)
    va   = avg / 100.0 * SWEEP
    v_i  = pt(r_inner, START_DEG + va)
    fg_i = f"M{s_i[0]},{s_i[1]} A{r_inner:.1f},{r_inner:.1f} 0 {1 if va>180 else 0},1 {v_i[0]},{v_i[1]}" if avg>0 else None
    tx, ty = round(cx,1), round(cy + r_outer * 0.22, 1)
    return (
        f'<svg viewBox="0 0 {size} {size}" width="{size}" height="{size}" '
        f'xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;display:block">'
        # track outer
        f'<path d="M{s_o[0]},{s_o[1]} A{r_outer:.1f},{r_outer:.1f} 0 1,1 {e_o[0]},{e_o[1]}" '
        f'stroke="#252d48" stroke-width="3" fill="none" stroke-linecap="round"/>'
        # track inner
        f'<path d="M{s_i[0]},{s_i[1]} A{r_inner:.1f},{r_inner:.1f} 0 1,1 {e_i[0]},{e_i[1]}" '
        f'stroke="#1e2538" stroke-width="2.5" fill="none" stroke-linecap="round"/>'
        # avg fill (inner, transparent color)
        + (f'<path d="{fg_i}" stroke="{color}55" stroke-width="2.5" fill="none" stroke-linecap="round"/>' if fg_i else '')
        # cur fill (outer, solid)
        + (f'<path d="{fg_o}" stroke="{color}" stroke-width="3" fill="none" stroke-linecap="round"/>' if fg_o else '')
        # dot at current position
        + f'<circle cx="{v_o[0]}" cy="{v_o[1]}" r="3.5" fill="{color}"/>'
        # label
        f'<text x="{tx}" y="{ty-5}" text-anchor="middle" font-family="Courier New,monospace" '
        f'font-size="5.5" fill="{color}" opacity="0.7">%</text>'
        f'<text x="{tx}" y="{ty+6}" text-anchor="middle" font-family="Courier New,monospace" '
        f'font-size="12" font-weight="bold" fill="{color}">{cur:.0f}</text>'
        f'</svg>'
    )

# ================================================================

# Точний маппінг колонок для кожного аркуша (0-based)
CROP_SHEET_COLS = {
    'Corn': {
        'week': 0,
        'stages': [
            ('Planted',    1, 2, 3),
            ('Emerged',    4, 5, 6),
            ('Silked',     7, 8, 9),
            ('Dough',     10,11,12),
            ('Dent',      13,14,15),
            ('Mature',    16,17,18),
            ('Harvested', 19,20,21),
        ]
    },
    'Soybeans': {
        'week': 0,
        'stages': [
            ('Planted',         1, 2, 3),
            ('Emerged',         4, 5, 6),
            ('Blooming',        7, 8, 9),
            ('Setting Pods',   10,11,12),
            ('Dropping Leaves',13,14,15),
            ('Harvested',      16,17,18),
        ]
    },
    'Winter_Wheat': {
        'week': 0,
        'stages': [
            ('Planted (Fall)',   1, 2,14),
            ('Emerged (Fall)',   4, 5,15),
            ('Headed (Spring)', 7, 8,16),
            ('Harvested',      10,11,17),
        ]
    },
    'Spring_Wheat': {
        'week': 0,
        'stages': [
            ('Planted',   1, 2, 3),
            ('Emerged',   4, 5, 6),
            ('Headed',    7, 8, 9),
            ('Harvested',10,11,12),
        ]
    },
    'Cotton': {
        'week': 0,
        'stages': [
            ('Planted',    1, 2, 3),
            ('Squaring',   4, 5, 6),
            ('Blooming',   7, 8, 9),
            ('Bolls Open',10,11,12),
            ('Harvested', 13,14,15),
        ]
    },
    'Rice': {
        'week': 0,
        'stages': [
            ('Planted',   1, 2, 3),
            ('Emerged',   4, 5, 6),
            ('Headed',    7, 8, 9),
            ('Harvested',10,11,12),
        ]
    },
}

# Переклади стадій — з заголовків ALL_Crops_Dashboard.xlsx (рядок "Planted\nПосіяно" тощо)
STAGE_UA = {
    'Planted':         'Посіяно',
    'Emerged':         'Зійшло',
    'Silked':          'Шовкування',
    'Dough':           'Молочна стиглість',
    'Dent':            'Воскова стиглість',
    'Mature':          'Повна стиглість',
    'Harvested':       'Зібрано',
    'Blooming':        'Цвітіння',
    'Setting Pods':    'Формування бобів',
    'Dropping Leaves': 'Опадання листків',
    'Headed':          'Колосіння',
    'Squaring':        'Бутонізація',
    'Bolls Open':      'Розкриття коробочок',
}
def stage_ua(name):
    """UA-переклад стадії; для 'Planted (Fall)' бере базову назву 'Planted'"""
    base = name.split('(')[0].strip()
    return STAGE_UA.get(base, '')

def _parse_crop_sheet(xl, sheet_name):
    """Читає аркуш культури, повертає list рядків {week, stage→{cur,avg,y25}}"""
    cfg = CROP_SHEET_COLS.get(sheet_name)
    if not cfg: return []
    df = xl.parse(sheet_name, header=None)
    rows = []
    for ri in range(5, len(df)):
        week = str(df.iloc[ri, cfg['week']]).strip()
        if not week or week=='nan' or week.startswith('Source'): continue
        entry = {'week': week}
        for stage_name, c26, cavg, c25 in cfg['stages']:
            def gv(ci):
                if ci >= df.shape[1]: return None
                v = df.iloc[ri, ci]
                try:
                    f = float(v)
                    return None if f < 0 else f   # -1 = future data
                except: return None
            entry[stage_name] = {'cur': gv(c26), 'avg': gv(cavg), 'y25': gv(c25)}
        rows.append(entry)
    return rows


def make_crop_tab(crop_data):
    """Crop Progress вкладка v2 — оновлена"""
    if not crop_data:
        return '<p style="padding:24px;color:#8090b0">Файл ALL_Crops_Dashboard.xlsx не знайдено</p>'

    import json as _json

    summary  = crop_data.get('summary', [])
    hist_raw = crop_data.get('hist', {})

    # Re-parse sheets for accurate data
    try:
        xl = pd.ExcelFile(CROP_FILE)
        sheet_rows = {s: _parse_crop_sheet(xl, s) for s in CROP_SHEET_COLS if s in xl.sheet_names}
    except Exception as e:
        sheet_rows = hist_raw
        print(f"  ⚠ make_crop_tab re-parse: {e}")

    # Group All Crops summary by crop
    by_crop = {}
    for row in summary:
        c = row['crop']
        if c not in by_crop: by_crop[c] = []
        by_crop[c].append(row)

    CROP_ORDER = [
        ('🌽', 'Corn',         'Corn',         'ZC1!',  '#f59420'),
        ('🫘', 'Soybeans',     'Soybeans',     'ZS1!',  '#20d483'),
        ('❄️', 'Winter Wheat', 'Winter_Wheat', 'ZW1!',  '#a78bfa'),
        ('🌾', 'Spring Wheat', 'Spring_Wheat', 'MWE1!', '#4a9eff'),
        ('🌿', 'Cotton',       'Cotton',       'CT1!',  '#22d3ee'),
        ('🍚', 'Rice',         'Rice',         'ZR1!',  '#f0b429'),
    ]

    def fmt_pct(v):
        if v is None: return '—'
        return f'{int(round(v))}%'

    def vs_badge(v_str):
        if not v_str or v_str=='—': return ''
        try:
            v = float(v_str.replace('+','').replace('pts','').replace(' ',''))
        except: return f'<span style="color:var(--d);font-size:9px">{v_str}</span>'
        col = '#20d483' if v>0 else '#f0515a' if v<0 else '#8090b0'
        bg  = 'rgba(32,212,131,.13)' if v>0 else 'rgba(240,81,90,.13)'
        bdr = '#20d483' if v>0 else '#f0515a'
        # v_str вже може мати знак (+5 pts) — використовуємо як є
        display = v_str.strip()
        if not display.startswith(('+','-')) and v>0:
            display = '+' + display
        return f'<span style="font-size:9px;padding:1px 7px;border-radius:10px;background:{bg};border:1px solid {bdr};color:{col};font-weight:bold">{display}</span>'

    # ── Stage cards ──
    def stage_cards_html(sheet_key, color, summary_rows):
        rows_by_stage = {r['stage'].split('(')[0].strip(): r for r in summary_rows}
        cfg = CROP_SHEET_COLS.get(sheet_key, {})
        stages = [s[0] for s in cfg.get('stages', [])]

        cards = ''
        for stage_name in stages:
            # Find matching summary row
            def safe_f(x, default=0):
                try:
                    v=float(x)
                    import math
                    return default if math.isnan(v) else v
                except: return default
            srow = None
            for k, r in rows_by_stage.items():
                if stage_name.lower().startswith(k.lower()[:6]) or k.lower().startswith(stage_name.lower()[:6]):
                    srow = r; break
            if not srow:
                # Stage with no data — still show
                cards += (f'<div class="cp-sc">'
                          f'<div class="cp-sc-name">{stage_name.upper()}<span class="cp-sc-ua">{stage_ua(stage_name)}</span></div>'
                          f'<div class="cp-sc-cur" style="color:var(--d)">—</div>'
                          f'<div class="cp-sc-row"><span class="cp-sc-sub" style="color:var(--d)">Avg —</span></div>'
                          f'<div class="cp-sc-row"><span class="cp-sc-sub" style="color:var(--d)">2025: —</span></div>'
                          f'</div>')
                continue

            cur_f = safe_f(srow.get('cur',0))
            avg_f = safe_f(srow.get('avg',0))
            lw_f  = safe_f(srow.get('lw',0))
            ly_f  = safe_f(srow.get('ly',0))

            bar_cur = min(max(cur_f,0),100)
            bar_avg = min(max(avg_f,0),100)

            cards += (
                f'<div class="cp-sc">'
                f'<div class="cp-sc-name">{stage_name.upper()}<span class="cp-sc-ua">{stage_ua(stage_name)}</span></div>'
                f'<div class="cp-sc-top" style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">'
                f'<div style="display:flex;flex-direction:column">'
                f'<div class="cp-sc-cur" style="color:{color}">{int(cur_f)}%</div>'
                + f'<div style="margin-top:2px">' + vs_badge(srow['vs']) + f'</div>' +
                f'</div>'
                + make_crop_gauge(cur_f, avg_f, color, size=58) +
                f'</div>'
                # bars
                # bar-рядок 2026 (поточний)
                f'<div class="cp-sc-bar-row">'
                f'<span class="cp-sc-bl">2026</span>'
                f'<div class="cp-sc-bg"><div class="cp-sc-fill" style="width:{bar_cur:.0f}%;background:{color}"></div></div>'
                f'<span class="cp-sc-bv" style="color:{color}">{int(cur_f)}%</span>'
                f'</div>'
                # bar-рядок Avg
                f'<div class="cp-sc-bar-row">'
                f'<span class="cp-sc-bl">Avg</span>'
                f'<div class="cp-sc-bg"><div class="cp-sc-fill" style="width:{bar_avg:.0f}%;background:#4a5580"></div></div>'
                f'<span class="cp-sc-bv" style="color:var(--d)">{int(avg_f)}%</span>'
                f'</div>'
                # bar-рядок 2025
                f'<div class="cp-sc-bar-row">'
                f'<span class="cp-sc-bl">2025</span>'
                f'<div class="cp-sc-bg"><div class="cp-sc-fill" style="width:{min(max(ly_f,0),100):.0f}%;background:#343d5a"></div></div>'
                f'<span class="cp-sc-bv" style="color:var(--d)">{int(ly_f)}%</span>'
                f'</div>'
                # bar-рядок LW (мин.тиждень)
                f'<div class="cp-sc-bar-row">'
                f'<span class="cp-sc-bl">LW</span>'
                f'<div class="cp-sc-bg"><div class="cp-sc-fill" style="width:{min(max(lw_f,0),100):.0f}%;background:{color}55"></div></div>'
                f'<span class="cp-sc-bv" style="color:var(--d)">{int(lw_f)}%</span>'
                f'</div>'
                f'<div style="border-top:1px solid var(--bd);margin:5px 0 2px"></div>'
                f'</div>'
            )
        return cards

    # ── Chart JSON per crop (all stages × weeks) ──
    def chart_json(sheet_key, color):
        rows = sheet_rows.get(sheet_key, [])
        cfg  = CROP_SHEET_COLS.get(sheet_key, {})
        stages = [s[0] for s in cfg.get('stages', [])]
        # Collect stage colors (hue shift from base color)
        stage_colors = {
            0: color,
            1: color+'cc',
            2: color+'99',
            3: color+'77',
            4: color+'55',
            5: color+'44',
            6: color+'33',
        }
        # Build datasets
        weeks   = [r['week'] for r in rows]
        datasets = []
        for si, stage_name in enumerate(stages):
            cur_vals = [r.get(stage_name,{}).get('cur') for r in rows]
            avg_vals = [r.get(stage_name,{}).get('avg') for r in rows]
            # 2026 actual
            datasets.append({'label': stage_name, 'data': cur_vals,
                              'type':'cur', 'color': stage_colors.get(si, color+'55')})
            # 5yr avg dashed
            datasets.append({'label': stage_name+' Avg', 'data': avg_vals,
                              'type':'avg', 'color': stage_colors.get(si, color+'55')})
        return _json.dumps({'weeks': weeks, 'stages': stages, 'datasets': datasets}, ensure_ascii=False)

    # ── Weekly table (like Excel) ──
    def weekly_table_html(sheet_key, color):
        rows = sheet_rows.get(sheet_key, [])
        if not rows: return '<p style="color:var(--d);padding:12px">Немає даних</p>'
        cfg    = CROP_SHEET_COLS.get(sheet_key, {})
        stages = [s[0] for s in cfg.get('stages', [])]
        ns = len(stages)

        # Header
        thead_r1 = '<th class="cp-wt-date">Week</th>'
        for sn in stages:
            r,g,b = int(color[1:3],16), int(color[3:5],16), int(color[5:7],16)
            bg = f'background:rgba({r},{g},{b},.18);border-left:2px solid {color}88'
            thead_r1 += f'<th colspan="3" class="cp-wt-th" style="{bg}">{sn}<br><span style="font-size:7px;opacity:.75;font-weight:normal">{stage_ua(sn)}</span></th>'
        thead_r2 = '<th></th>'
        for _ in stages:
            sp = 'background:var(--bg3);color:var(--d)'
            thead_r2 += f'<th class="cp-wt-sub" style="{sp}">2026</th><th class="cp-wt-sub" style="{sp}">Avg</th><th class="cp-wt-sub" style="{sp}">2025</th>'

        # Body rows (newest first)
        body = ''
        # Фільтруємо рядки де є хоч якийсь Avg (щоб не показувати порожні)
        valid_rows = [r for r in rows if any(
            r.get(s,{}).get('avg') is not None for s in stages)]
        for row in valid_rows:
            tds = f'<td class="cp-wt-date">{row["week"]}</td>'
            for sn in stages:
                sv = row.get(sn, {})
                cur = sv.get('cur'); avg = sv.get('avg'); y25 = sv.get('y25')
                def td(v, is_cur=False):
                    try:
                        vf = float(v)
                    except (TypeError, ValueError):
                        return '<td class="cp-wt-val d">—</td>'
                    import math
                    if v is None or math.isnan(vf) or vf == 0:
                        return '<td class="cp-wt-val d">—</td>'
                    c = color if is_cur else 'var(--d)'
                    fw = 'font-weight:bold;' if is_cur else ''
                    return f'<td class="cp-wt-val" style="color:{c};{fw}">{int(round(vf))}%</td>'
                tds += td(cur, True) + td(avg) + td(y25)
            body += f'<tr>{tds}</tr>'

        cols = f'<colgroup><col style="width:72px">' + f'<col><col><col>' * ns + '</colgroup>'
        return (f'<div style="overflow-x:auto">'
                f'<table class="cp-wt">{cols}'
                f'<thead><tr>{thead_r1}</tr><tr>{thead_r2}</tr></thead>'
                f'<tbody>{body}</tbody></table></div>')

    # ── Build tabs & panels ──
    tabs_html = ''
    panels_html = ''
    all_chart_data = {}
    first = True

    for emoji, label, sheet_key, tv, color in CROP_ORDER:
        cid = sheet_key.replace('_','').lower()
        act = ' active' if first else ''

        # Summary rows for this crop
        sum_rows = []
        for crop_name, rows in by_crop.items():
            short = label.split()[0].lower()
            if short in crop_name.lower() or sheet_key.lower().replace('_',' ') in crop_name.lower():
                sum_rows = rows; break

        # Stage cards
        sc_html = stage_cards_html(sheet_key, color, sum_rows)

        # Chart data
        cj = chart_json(sheet_key, color)
        all_chart_data[cid] = cj

        # Weekly table
        wt_html = weekly_table_html(sheet_key, color)

        # Latest date from data
        rows_data = sheet_rows.get(sheet_key, [])
        actual = [r for r in rows_data if any(r.get(s,{}).get('cur') is not None for s in CROP_SHEET_COLS.get(sheet_key,{}).get('stages',[{}])) ]
        latest_date = actual[-1]['week'] if actual else '—'

        tabs_html += (
            f'<button class="cp-tab{act}" data-crop="{cid}" onclick="selCrop(\'{cid}\')"'
            f' style="{"border-bottom-color:"+color if first else ""}">'
            f'{emoji} {label}</button>'
        )

        panels_html += (
            f'<div class="cp-panel{"" + act}" id="cp_{cid}">'
            # ── Header ──
            f'<div class="cp-phdr">'
            f'<div style="display:flex;align-items:center;gap:12px">'
            f'<span style="font-size:26px">{emoji}</span>'
            f'<div><div style="font-size:13px;font-weight:bold;color:#fff;letter-spacing:.8px">{label}</div>'
            f'<div style="font-size:9px;color:var(--d)">USDA NASS &nbsp;|&nbsp; {tv} &nbsp;|&nbsp; Останні дані: {latest_date}</div>'
            f'</div></div>'
            f'<div style="font-size:9px;color:var(--d)">Crop Progress Report</div>'
            f'</div>'
            # ── Stage cards ──
            f'<div class="cp-sg">{sc_html}</div>'
            # ── Charts section ──
            f'<div class="cp-charts-wrap">'
            # Перемикач графіків (як фото3 — tabs для LS/CM/ST)
            f'<div class="cp-chart-top">'
            f'<div class="cp-chart-lbl">ДИНАМІКА ПОСІВУ</div>'
            f'<div class="cp-stage-btns" id="cpStBtns_{cid}">'
        )
        # Stage selector buttons
        cfg = CROP_SHEET_COLS.get(sheet_key, {})
        stg_list = [s[0] for s in cfg.get('stages', [])]
        for si, sn in enumerate(stg_list):
            act2 = ' active' if si==0 else ''
            panels_html += (f'<button class="cp-stbtn{act2}" data-stage="{sn}" '
                            f'onclick="selCropStage(\'{cid}\',this)" title="{stage_ua(sn)}">{sn}</button>')

        panels_html += (
            f'</div></div>'
            # Лінійний графік вибраної стадії (2026 vs Avg vs 2025)
            f'<div class="cp-line-wrap"><canvas id="cpLine_{cid}"></canvas></div>'
            # S-curve графік (всі стадії разом)
            f'<div class="cp-scurve-hdr">КРИВІ РОЗВИТКУ СЕЗОНУ (всі стадії)</div>'
            f'<div class="cp-scurve-wrap"><canvas id="cpScurve_{cid}"></canvas></div>'
            f'</div>'
            # ── Weekly table ──
            f'<div class="cp-wt-hdr">ТИЖНЕВА СТАТИСТИКА</div>'
            + wt_html +
            f'<script>_cpData["{cid}"]={cj};</script>'
            f'</div>'
        )
        first = False

    # ── CSS ──
    css = """
<style>
/* Crop Progress v2 */
.cp-wrap{padding:16px 24px;}
.cp-tabs{display:flex;gap:0;border-bottom:2px solid var(--bd);margin-bottom:14px;}
.cp-tab{padding:8px 18px;border:none;background:transparent;color:#b0bcd4;font-family:var(--f);font-size:11px;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-2px;}
.cp-tab:hover{color:#fff;}.cp-tab.active{color:#fff;font-weight:bold;}
.cp-panel{display:none;}.cp-panel.active{display:block;}
.cp-phdr{display:flex;justify-content:space-between;align-items:center;background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:12px 16px;margin-bottom:12px;}
/* Stage cards */
.cp-sg{display:grid;grid-template-columns:repeat(auto-fill,minmax(155px,1fr));gap:8px;margin-bottom:12px;}
.cp-sc{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:11px 13px;}
.cp-sc-name{font-size:8px;color:var(--d);letter-spacing:.6px;text-transform:uppercase;margin-bottom:4px;}
.cp-sc-ua{display:block;font-size:8px;color:#aeb9d6;letter-spacing:.2px;text-transform:none;margin-top:1px;}
.cp-sc-cur{font-size:28px;font-weight:bold;line-height:1.1;margin-bottom:7px;}
.cp-sc-bar-row{display:flex;align-items:center;gap:5px;margin-bottom:3px;}
.cp-sc-bl{font-size:8px;color:var(--d);width:24px;flex-shrink:0;}
.cp-sc-bg{flex:1;height:5px;background:var(--bg3);border-radius:2px;overflow:hidden;}
.cp-sc-fill{height:100%;border-radius:2px;}
.cp-sc-bv{font-size:8px;width:28px;text-align:right;flex-shrink:0;}
.cp-sc-meta{display:flex;align-items:center;gap:5px;flex-wrap:wrap;margin-top:6px;padding-top:5px;border-top:1px solid var(--bd);}
.cp-sc-meta2{margin-top:3px;}
.cp-sc-sub{font-size:9px;color:var(--d);}
/* Charts */
.cp-charts-wrap{background:var(--bg2);border:1px solid var(--bd);border-radius:5px;padding:14px 16px;margin-bottom:10px;}
.cp-chart-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:8px;}
.cp-chart-lbl{font-size:9px;color:#fff;letter-spacing:.5px;}
.cp-stage-btns{display:flex;gap:3px;flex-wrap:wrap;}
.cp-stbtn{padding:2px 8px;border:1px solid var(--bd);border-radius:3px;cursor:pointer;color:#b0bcd4;font-family:var(--f);font-size:9px;background:transparent;}
.cp-stbtn:hover{border-color:var(--accent);color:#fff;}
.cp-stbtn.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);}
.cp-line-wrap{height:280px;position:relative;margin-bottom:12px;}
.cp-scurve-hdr{font-size:9px;color:#fff;letter-spacing:.5px;margin-bottom:8px;padding-top:8px;border-top:1px solid var(--bd);}
.cp-scurve-wrap{height:320px;position:relative;}
/* Weekly table */
.cp-wt-hdr{font-size:9px;color:#fff;letter-spacing:.5px;padding:8px 0 6px;margin-top:10px;}
table.cp-wt{width:100%;border-collapse:collapse;font-size:10px;white-space:nowrap;}
table.cp-wt th{padding:4px 8px;background:var(--bg3);font-weight:normal;font-size:8px;letter-spacing:.3px;text-align:right;border-bottom:1px solid var(--bd);}
table.cp-wt .cp-wt-date{text-align:left;color:var(--d);}
table.cp-wt .cp-wt-th{text-align:center;}
table.cp-wt .cp-wt-sub{text-align:right;color:var(--d);}
table.cp-wt td{padding:4px 8px;border-bottom:1px solid var(--bg3);text-align:right;}
table.cp-wt .cp-wt-val{text-align:right;}
table.cp-wt .cp-wt-val.d{color:var(--d);}

</style>
"""

    # ── JS ──
    js = """
<script>
(function(){
const _CROP_CHARTS={};

function drawCropLine(cid, stageName){
  const cv=document.getElementById('cpLine_'+cid); if(!cv)return;
  if(_CROP_CHARTS['line_'+cid]){_CROP_CHARTS['line_'+cid].destroy();}
  const d=_cpData[cid]; if(!d)return;
  const di=d.datasets.find(ds=>ds.label===stageName&&ds.type==='cur');
  const da=d.datasets.find(ds=>ds.label===stageName+' Avg'&&ds.type==='avg');
  const d25=d.datasets.find(ds=>ds.label===stageName+' 2025'||ds.type==='y25');
  const weeks=d.weeks;
  // Build 2025 data from sheet_rows if available (passed as separate key)
  const cur25=(d.y25||{})[stageName]||null;
  const color=di?di.color:'#f59420';
  const datasets=[
    {label:'2026',data:di?di.data:[],borderColor:color,backgroundColor:color+'22',borderWidth:2,pointRadius:3,tension:.3,fill:true,spanGaps:false},
    {label:'5yr Avg',data:da?da.data:[],borderColor:'#4a5580',backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3,spanGaps:false},
  ];
  if(cur25){datasets.push({label:'2025',data:cur25,borderColor:'#8090b0',backgroundColor:'transparent',borderWidth:1,pointRadius:0,tension:.3,borderDash:[3,3],spanGaps:false});}
  _CROP_CHARTS['line_'+cid]=new Chart(cv.getContext('2d'),{
    type:'line',data:{labels:weeks,datasets},
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      interaction:{mode:'index',intersect:false},
      plugins:{legend:{display:false},tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',titleFont:{family:'Courier New',size:9},bodyFont:{family:'Courier New',size:9},callbacks:{label:ctx=>ctx.dataset.label+': '+(ctx.parsed.y!=null?ctx.parsed.y+'%':'—')}}},
      scales:{x:{ticks:{color:'#8090b0',font:{family:'Courier New',size:8}},grid:{display:false},border:{display:false}},
        y:{min:0,max:100,grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},ticks:{color:'#8090b0',font:{family:'Courier New',size:8},callback:v=>v+'%'},border:{display:false}}}}
  });
}

function drawCropScurve(cid){
  const cv=document.getElementById('cpScurve_'+cid); if(!cv)return;
  if(_CROP_CHARTS['sc_'+cid]){_CROP_CHARTS['sc_'+cid].destroy();}
  const d=_cpData[cid]; if(!d)return;
  const weeks=d.weeks;
  // Only 2026 actual per stage
  const datasets=d.datasets.filter(ds=>ds.type==='cur').map(ds=>({
    label:ds.label,data:ds.data,borderColor:ds.color,
    backgroundColor:ds.color.length>7?ds.color:ds.color+'33',
    borderWidth:1.5,pointRadius:2,tension:.3,fill:false,spanGaps:false
  }));
  _CROP_CHARTS['sc_'+cid]=new Chart(cv.getContext('2d'),{
    type:'line',data:{labels:weeks,datasets},
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      interaction:{mode:'index',intersect:false},
      plugins:{legend:{display:true,labels:{color:'#8090b0',font:{family:'Courier New',size:8},boxWidth:12}},
        tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',titleFont:{family:'Courier New',size:9},bodyFont:{family:'Courier New',size:9},callbacks:{label:ctx=>ctx.dataset.label+': '+(ctx.parsed.y!=null?ctx.parsed.y+'%':'—')}}},
      scales:{x:{ticks:{color:'#8090b0',font:{family:'Courier New',size:8}},grid:{display:false},border:{display:false}},
        y:{min:0,max:100,grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},ticks:{color:'#8090b0',font:{family:'Courier New',size:8},callback:v=>v+'%'},border:{display:false}}}}
  });
}

window.selCropStage=function(cid,btn){
  document.querySelectorAll('#cpStBtns_'+cid+' .cp-stbtn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  drawCropLine(cid, btn.dataset.stage);
};

window.selCrop=function(cid){
  document.querySelectorAll('.cp-tab').forEach(b=>{b.classList.remove('active');b.style.borderBottomColor='';});
  document.querySelectorAll('.cp-panel').forEach(p=>p.classList.remove('active'));
  const btn=document.querySelector('[data-crop="'+cid+'"]');
  if(btn){btn.classList.add('active');
    const d=_cpData[cid];const color=(d&&d.datasets&&d.datasets[0])?d.datasets[0].color:'#f59420';
    btn.style.borderBottomColor=color;}
  const panel=document.getElementById('cp_'+cid);
  if(panel){panel.classList.add('active');}
  const firstBtn=document.querySelector('#cpStBtns_'+cid+' .cp-stbtn.active')||document.querySelector('#cpStBtns_'+cid+' .cp-stbtn');
  if(firstBtn){firstBtn.classList.add('active');drawCropLine(cid,firstBtn.dataset.stage);}
  drawCropScurve(cid);
};

// Init first
setTimeout(()=>{
  const f=document.querySelector('.cp-tab.active');
  if(f){const c=f.dataset.crop;
    const d=_cpData[c];const color=(d&&d.datasets&&d.datasets[0])?d.datasets[0].color:'#f59420';
    f.style.borderBottomColor=color;
    const fb=document.querySelector('#cpStBtns_'+c+' .cp-stbtn.active')||document.querySelector('#cpStBtns_'+c+' .cp-stbtn');
    if(fb){fb.classList.add('active');drawCropLine(c,fb.dataset.stage);}
    drawCropScurve(c);
  }
},120);
})();
</script>
"""

    return f'{css}<div class="cp-wrap"><div class="cp-tabs">{tabs_html}</div><div class="cp-panels">{panels_html}</div></div>{js}'


def main():
    print()
    print("="*55)
    print("   COT Dashboard Generator v15 — Legacy + TFF + DISAG")
    print("="*55)
    print()
    OUTPUT_FILE.parent.mkdir(parents=True,exist_ok=True)
    try: data=load_all()
    except FileNotFoundError as e: print(e); return
    if not data: print("❌  Дані порожні."); return
    tff_data=load_tff_data()
    disag_data=load_disag_data()
    crop_data=load_crop_data()
    print("🔧  Генеруємо HTML...")
    html=generate_html(data, tff_data, disag_data, crop_data)
    OUTPUT_FILE.write_text(html, encoding='utf-8')
    kb=OUTPUT_FILE.stat().st_size/1024
    print(f"✅  Збережено: {OUTPUT_FILE}  ({kb:.0f} KB)")
    FLASK_URL="http://localhost:5000"
    print(f"🌐  Відкриваємо: {FLASK_URL}")
    import os
    if not webbrowser.open(FLASK_URL):
        try: os.startfile(str(OUTPUT_FILE))
        except: pass
    print(f"   Файл: {OUTPUT_FILE}")
    print("\n✨  Готово!\n")

if __name__=='__main__':
    main()

# ================================================================
# CROP PROGRESS — читання та UI
# ================================================================