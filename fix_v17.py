#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v21.py — Overview tab: компактність + кольорові фони + перемикач COT періодів
             + стовпець CROWDED + вищий графік SM DIV.

Запускати з папки проекту (де лежить generate.py):
    python fix_v21.py

Ідемпотентний: якщо вже застосовано — пропустить. Робить бекап generate.py.
"""
import re, sys, shutil
from pathlib import Path
from datetime import datetime

GEN = Path("generate.py")
if not GEN.exists():
    print("❌  generate.py не знайдено. Запусти скрипт з папки проекту.")
    sys.exit(1)

src = GEN.read_text(encoding="utf-8")
orig = src
applied = []
skipped = []

# ────────────────────────────────────────────────────────────────
# КРОК 1. read_overview: додаємо crowded + cot_idx (за періодами) у рядок таблиці
# ────────────────────────────────────────────────────────────────
if "'crowded':" in src and "OVERVIEW_TABLE.append({'asset'" in src:
    skipped.append("1) crowded/cot_idx у read_overview (вже є)")
else:
    # 1a. зчитати crowded (кол U -> 0-based 20) і cot_idx (net серії з інстр-аркушів немає,
    #     тому period-COT братимемо пізніше з data у generate_html). Тут лише crowded + net.
    anchor1 = "sm_div=safe(8);sm_div_3m=safe(18);sm_div_6m=safe(19)"
    repl1 = "sm_div=safe(8);sm_div_3m=safe(18);sm_div_6m=safe(19)\n            crowded=str(row.iloc[20]).strip() if pd.notna(row.iloc[20]) else '—'"
    if anchor1 in src:
        src = src.replace(anchor1, repl1, 1)
    else:
        print("❌ 1a anchor не знайдено"); sys.exit(1)

    # 1b. додати 'crowded' у dict рядка OVERVIEW_TABLE
    anchor2 = "'sm_div':sm_div,'sm_div_3m':sm_div_3m,'sm_div_6m':sm_div_6m})"
    repl2 = "'sm_div':sm_div,'sm_div_3m':sm_div_3m,'sm_div_6m':sm_div_6m,'crowded':crowded})"
    if anchor2 in src:
        src = src.replace(anchor2, repl2, 1)
    else:
        print("❌ 1b anchor не знайдено"); sys.exit(1)
    applied.append("1) crowded зчитується з overview (кол U) + кладеться у рядок")

# ────────────────────────────────────────────────────────────────
# КРОК 2. generate_html: пробросити COT-періоди з data у OVERVIEW_TABLE
#          (data вже має d['cot_idx']['ls'|'cm'|'st']['all'|'3y'|'1y'|'6m'|'3m'])
# ────────────────────────────────────────────────────────────────
if "# v21: enrich overview з cot_idx періодів" in src:
    skipped.append("2) enrich cot_idx періодів (вже є)")
else:
    anchor_gh = "def generate_html(data, tff_data=None, disag_data=None, crop_data=None):\n    if tff_data is None: tff_data={}"
    repl_gh = ("def generate_html(data, tff_data=None, disag_data=None, crop_data=None):\n"
               "    if tff_data is None: tff_data={}\n"
               "    # v21: enrich overview з cot_idx періодів (беремо з data по sid)\n"
               "    for _it in OVERVIEW_TABLE:\n"
               "        if isinstance(_it, dict):\n"
               "            _d = data.get(_it.get('sid'))\n"
               "            if _d and _d.get('cot_idx'):\n"
               "                _it['cot_idx'] = _d['cot_idx']")
    if anchor_gh in src:
        src = src.replace(anchor_gh, repl_gh, 1)
        applied.append("2) COT-періоди пробрасуються з data у рядок overview")
    else:
        print("❌ 2 anchor (generate_html) не знайдено"); sys.exit(1)

# ────────────────────────────────────────────────────────────────
# КРОК 3. Повна заміна make_overview_tab (нова таблиця з фонами, перемикачем, CROWDED)
# ────────────────────────────────────────────────────────────────
if "# v21 overview" in src:
    skipped.append("3) make_overview_tab (вже v21)")
else:
    # знайти межі функції: від 'def make_overview_tab():' до наступного 'AUTH_MODAL_HTML ='
    m_start = src.find("def make_overview_tab():")
    m_end = src.find("AUTH_MODAL_HTML =")
    if m_start == -1 or m_end == -1 or m_end < m_start:
        print("❌ 3 не знайдено межі make_overview_tab"); sys.exit(1)

    new_fn = r'''def make_overview_tab():
    # v21 overview — компактна, кольорові фони, перемикач COT-періодів, CROWDED
    rows_html=[];rep_date='—';today_date='—'
    # збираємо cot-періоди для JS-перемикача: { sid: {ls:{all,3y,1y,6m,3m}, cm:{...}, st:{...}} }
    cot_periods={}
    for item in OVERVIEW_TABLE:
        if isinstance(item,tuple) and item[0]=='_meta': rep_date=item[1];today_date=item[2];continue
        if isinstance(item,tuple) and item[0]=='_group': rows_html.append(f'<tr class="ov-group"><td colspan="14">{item[1]}</td></tr>');continue
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
            pct=min(max(v/100,0),1);color='#f0515a'if v<lo else('#20d483'if v>hi else'#4a9eff')
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
            return (f'<td class="ov-cot-cell-td" data-cotgrp="{grp}" '
                    f'data-all="{_f(da)}" data-3y="{_f(d3y)}" data-1y="{_f(d1y)}" '
                    f'data-6m="{_f(d6)}" data-3m="{_f(d3)}">'
                    f'<div class="ov-cot-cell">{pct_bar(init)}<span class="ov-cot-val">{init:.0f}%</span></div></td>')
        def sm_fmt(v):
            if v is None: return '<span class="d">—</span>'
            cls='g'if float(v)>0 else('r'if float(v)<0 else'd'); return f'<span class="{cls}">{float(v):+.2f}</span>'
        def crowded_fmt(cw):
            if not cw or cw in('—','nan','None'): return '<span class="d">—</span>'
            low=cw.lower()
            if 'very' in low: return '<span class="ov-crowd ov-crowd-vc">Very Crowded</span>'
            if 'crowd' in low: return '<span class="ov-crowd ov-crowd-c">Crowded</span>'
            return f'<span class="d">{cw}</span>'
        rows_html.append(f'<tr class="ov-row">'
                         f'<td class="ov-asset">{d["asset"]}</td>'
                         f'{fnum_td(d["net_ls"])}'
                         f'{cell_bg(d["chg_ls"])}'
                         f'{pctcell_bg(d.get("chg_pct_ls"))}'
                         f'{fnum_td(d["net_cm"])}'
                         f'{cell_bg(d["chg_cm"])}'
                         f'{pctcell_bg(d.get("chg_pct_cm"))}'
                         f'{cot_td(d["sid"],"ls")}'
                         f'{cot_td(d["sid"],"cm")}'
                         f'{cot_td(d["sid"],"st")}'
                         f'<td>{crowded_fmt(d.get("crowded"))}</td>'
                         f'<td>{sm_fmt(d["sm_div"])}</td><td>{sm_fmt(d["sm_div_6m"])}</td><td>{sm_fmt(d["sm_div_3m"])}</td></tr>')
    thead=(f'<thead><tr><th class="ov-asset">ASSET</th><th>NET LS</th><th>CHG LS</th><th>%OIΔ</th>'
           f'<th>NET CM</th><th>CHG CM</th><th>%OIΔ</th>'
           f'<th style="color:{COLOR_LS}">COT LS</th><th style="color:{COLOR_CM}">COT CM</th><th style="color:{COLOR_ST}">COT ST</th>'
           f'<th>CROWDED</th>'
           f'<th>SM DIV</th><th>SM 6M</th><th>SM 3M</th></tr></thead>')
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

def fnum_td(v):
    # NET-клітинка без фону (звичайна кольорова цифра)
    if v is None: return '<td class="ov-num d">—</td>'
    try:
        nv=int(round(float(v)));body=f"{abs(nv):,}".replace(',','\u202f')
        s2='+' if nv>0 else('-' if nv<0 else'');cls='g'if nv>0 else('r'if nv<0 else'd')
        return f'<td class="ov-num {cls}">{s2}{body}</td>'
    except: return '<td class="ov-num d">—</td>'

'''
    src = src[:m_start] + new_fn + "\n" + src[m_end:]
    applied.append("3) make_overview_tab повністю замінено (v21)")

# ────────────────────────────────────────────────────────────────
# КРОК 4. CSS: компактний overview + кольорові фони + crowded + вищий SM graph
# ────────────────────────────────────────────────────────────────
if "/* v21 overview css */" in src:
    skipped.append("4) CSS overview (вже v21)")
else:
    css_anchor = ".ov-cot-cell{display:flex;align-items:center;gap:6px;justify-content:flex-end;}"
    css_add = css_anchor + """
/* v21 overview css */
#ms_ov>div{max-width:1180px;margin:0 auto;}
.ov-table{font-size:11px;}
.ov-table th{padding:7px 9px;font-size:9px;}
.ov-table td{padding:6px 9px;}
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
.ov-crowd-c{background:rgba(240,180,41,.15);border:1px solid #f0b429;color:#f0b429;}
.ov-crowd-vc{background:rgba(240,81,90,.18);border:1px solid #f0515a;color:#f0515a;}
.ov-sm-cv-wrap{height:340px!important;}
"""
    if css_anchor in src:
        src = src.replace(css_anchor, css_add, 1)
        applied.append("4) CSS overview (компактність, фони, crowded, вищий графік)")
    else:
        print("❌ 4 css anchor не знайдено"); sys.exit(1)

# ────────────────────────────────────────────────────────────────
# КРОК 5. JS: перемикач COT-періодів (ovSetPer)
# ────────────────────────────────────────────────────────────────
if "function ovSetPer(" in src:
    skipped.append("5) JS ovSetPer (вже є)")
else:
    js_anchor = "// ── Overview SM DIV bar chart ──"
    js_add = """// ── v21: Overview COT period switcher ──
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
    const color=v<15?'#f0515a':v>85?'#20d483':'#4a9eff';
    const pct=Math.min(Math.max(v,0),100);
    cell.innerHTML='<div class="ov-bar-bg"><div class="ov-bar-fill" style="width:'+pct.toFixed(1)+'%;background:'+color+'"></div></div>'
      +'<span class="ov-cot-val">'+v.toFixed(0)+'%</span>';
  });
}

// ── Overview SM DIV bar chart ──"""
    if js_anchor in src:
        src = src.replace(js_anchor, js_add, 1)
        applied.append("5) JS ovSetPer — перемикач COT-періодів")
    else:
        print("❌ 5 js anchor не знайдено"); sys.exit(1)

# ────────────────────────────────────────────────────────────────
# Збереження
# ────────────────────────────────────────────────────────────────
if src == orig:
    print("\nℹ️  Нічого не змінено — всі кроки вже застосовані.")
    for s in skipped: print("   ⏭  ", s)
    sys.exit(0)

bak = GEN.with_suffix(f".py.bak_v21_{datetime.now():%Y%m%d_%H%M%S}")
shutil.copy2(GEN, bak)
GEN.write_text(src, encoding="utf-8")

print("\n✅  Патч v21 застосовано.")
print(f"   Бекап: {bak.name}")
for s in applied: print("   ✓ ", s)
for s in skipped: print("   ⏭  ", s)

# синтаксична перевірка
import py_compile
try:
    py_compile.compile(str(GEN), doraise=True)
    print("\n✅  generate.py компілюється без помилок.")
except py_compile.PyCompileError as e:
    print("\n❌  Синтаксична помилка після патчу:\n", e)
    print("   Відновлюю з бекапу...")
    shutil.copy2(bak, GEN)
    print("   Відновлено. Напиши мені текст помилки.")
    sys.exit(1)

print("\n▶  Далі: python generate.py  (перегенерувати HTML)")