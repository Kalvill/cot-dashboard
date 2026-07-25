#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v26.py — 2 зміни:
  1) TFF/DISAG: gauge COT INDEX починаються з однієї вертикальної лінії
     (фіксована ширина лівого блоку метрик, щоб gauges не "плавали")
  2) Legacy: блок analysis_row -> той самий вигляд, що й TFF
     (метрики зліва: CHG LONG / CHG SHORT / Δ NET / NET POS,
      gauges COT INDEX ALL + 1Y справа)

Запускати з папки проекту:
    python fix_v26.py

Ідемпотентний, робить бекап, відкочується при синтаксичній помилці.
"""
import sys, shutil
from pathlib import Path
from datetime import datetime

GEN = Path("generate.py")
if not GEN.exists():
    print("❌  generate.py не знайдено. Запусти з папки проекту.")
    sys.exit(1)

src = GEN.read_text(encoding="utf-8")
orig = src
applied, skipped = [], []

def need(anchor, step):
    if anchor not in src:
        print(f"❌  Крок {step}: anchor не знайдено:\n    {anchor[:90]}")
        sys.exit(1)

# ════════════════════════════════════════════════════════════════
# КРОК 1 — Фіксована ширина метрик, gauges з однієї лінії
# Причина: .tff-a-left має flex:0 1 auto -> ширина залежить від вмісту,
# тому gauges стартують по-різному. Робимо left фіксованої ширини.
# ════════════════════════════════════════════════════════════════
if "/* v26 gauge align */" in src:
    skipped.append("1) вирівнювання gauges (вже v26)")
else:
    old_css = """/* v25 gauges center */
.tff-a-row{display:flex;align-items:center;justify-content:flex-start;padding:12px 16px;border-bottom:1px solid var(--bd);gap:16px;}
.tff-a-row:last-child{border-bottom:none;}
.tff-a-left{flex:0 1 auto;min-width:0;}"""

    new_css = """/* v26 gauge align */
.tff-a-row{display:flex;align-items:center;justify-content:flex-start;padding:12px 16px;border-bottom:1px solid var(--bd);gap:16px;}
.tff-a-row:last-child{border-bottom:none;}
.tff-a-left{flex:0 0 620px;max-width:620px;min-width:0;}"""
    need(old_css, 1)
    src = src.replace(old_css, new_css, 1)

    # gauges: тепер відступ не потрібен (left має фіксовану ширину) — прибираємо margin
    old_g = """.tff-a-gauges{display:flex;gap:10px;flex-shrink:0;align-items:center;border-left:1px solid var(--bd);padding-left:24px;margin-left:40px;}"""
    new_g = """.tff-a-gauges{display:flex;gap:10px;flex-shrink:0;align-items:center;border-left:1px solid var(--bd);padding-left:24px;}"""
    need(old_g, 1)
    src = src.replace(old_g, new_g, 1)

    # метрики: фіксуємо ширину кожного стовпчика, щоб NET POS теж був на місці
    old_m = """.tff-a-metrics{display:flex;align-items:flex-end;gap:22px;flex-wrap:wrap;}
.tff-ag-item{display:flex;flex-direction:column;gap:3px;}"""
    new_m = """.tff-a-metrics{display:flex;align-items:flex-end;gap:22px;flex-wrap:nowrap;}
.tff-ag-item{display:flex;flex-direction:column;gap:3px;min-width:110px;}"""
    need(old_m, 1)
    src = src.replace(old_m, new_m, 1)
    applied.append("1) gauges COT INDEX стартують з однієї вертикальної лінії")

# ════════════════════════════════════════════════════════════════
# КРОК 2 — Legacy analysis_row у стилі TFF
# ════════════════════════════════════════════════════════════════
if "# v26 legacy analysis" in src:
    skipped.append("2) analysis_row (вже v26)")
else:
    old_fn = """def analysis_row(group_label,group_color,net,cl,cs,chg,chg_pct):
    dc='g'if net>0 else'r'
    return(f'<div class="arow"><div class="arow-body"><div class="arow-left">'
           f'<div class="arow-grid2">'
           f'<div class="ag-item"><span class="ag-lbl">CHG LONG</span><span class="{cc(cl)} ag-val">{fv_full(cl,sign=True)}</span></div>'
           f'<div class="ag-item"><span class="ag-lbl">CHG SHORT</span><span class="{cc(cs)} ag-val">{fv_full(cs,sign=True)}</span></div>'
           f'</div>'
           f'<div class="arow-dnet"><span class="ag-lbl">Δ NET</span>'
           f'<span class="{cc(chg)} ag-val-net">{fv_full(chg,sign=True)}<span class="ag-pct"> ({chg_pct})</span></span></div>'
           f'</div><div class="arow-right">'
           f'<div class="ar-glbl" style="color:{group_color}">{group_label}</div>'
           f'<div class="ar-net {dc}">{fv_full(net,sign=True)}</div>'
           f'</div></div></div>')"""

    new_fn = """def analysis_row(group_label,group_color,net,cl,cs,chg,chg_pct,idx=None):
    # v26 legacy analysis — вигляд як TFF: метрики зліва, gauges COT INDEX справа
    dc='g'if net>0 else'r'
    idx=idx or {}
    g_all=idx.get('all',50.0);g_1y=idx.get('1y',50.0)
    col_all=gauge_color(g_all);col_1y=gauge_color(g_1y)
    gauge_all=make_gauge_svg(g_all,col_all,size=62,label='COT ALL')
    gauge_1y =make_gauge_svg(g_1y, col_1y, size=62,label='COT 1Y')
    return(f'<div class="tff-a-row">'
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
           f'<div class="tff-a-gwrap">{gauge_1y}</div>'
           f'</div></div>')"""
    need(old_fn, 2)
    src = src.replace(old_fn, new_fn, 1)
    applied.append("2) Legacy analysis_row переверстано у стиль TFF")

# ════════════════════════════════════════════════════════════════
# КРОК 3 — передати cot_idx у виклики analysis_row + обгорнути у tff-панель
# ════════════════════════════════════════════════════════════════
if "# v26 analysis calls" in src:
    skipped.append("3) виклики analysis_row (вже v26)")
else:
    old_call = """    analysis_panel=(f'<div class="panel">'
                    +analysis_row('LARGE SPEC',COLOR_LS,c['ls_net'],c['ls_cl'],c['ls_cs'],c['ls_chg'],c['ls_chg_pct'])
                    +analysis_row('COMMERCIALS',COLOR_CM,c['cm_net'],c['cm_cl'],c['cm_cs'],c['cm_chg'],c['cm_chg_pct'])
                    +f'</div>')"""
    new_call = """    # v26 analysis calls — передаємо cot_idx для gauges, обгортка як у TFF
    _ci=d['cot_idx']
    analysis_panel=(f'<div class="panel tff-analysis-panel">'
                    +analysis_row('LARGE SPEC',COLOR_LS,c['ls_net'],c['ls_cl'],c['ls_cs'],c['ls_chg'],c['ls_chg_pct'],_ci['ls'])
                    +analysis_row('COMMERCIALS',COLOR_CM,c['cm_net'],c['cm_cl'],c['cm_cs'],c['cm_chg'],c['cm_chg_pct'],_ci['cm'])
                    +f'</div>')"""
    need(old_call, 3)
    src = src.replace(old_call, new_call, 1)
    applied.append("3) LARGE SPEC/COMMERCIALS отримують gauges COT INDEX")

# ════════════════════════════════════════════════════════════════
# КРОК 4 — Винести analysis_panel окремим повним рядком над mid
# (широкий TFF-стиль не влазить у grid-колонку 1fr)
# ════════════════════════════════════════════════════════════════
if "# v26 mid layout" in src:
    skipped.append("4) mid layout (вже v26)")
else:
    old_mid = "    mid=f'<div class=\"mid\">{analysis_panel}{sm_panel}{pct_combined}</div>'"
    new_mid = ("    # v26 mid layout — analysis повним рядком зверху, далі sm+pct\n"
               "    mid=f'{analysis_panel}<div class=\"mid mid-nopanel\">{sm_panel}{pct_combined}</div>'")
    need(old_mid, 4)
    src = src.replace(old_mid, new_mid, 1)

    # grid тепер без першої 1fr-колонки: sm(180px) + pct(решта)
    grid_anchor = ".mid{display:grid;grid-template-columns:1fr 180px 1fr;gap:8px;margin-bottom:12px;}"
    grid_add = (".mid{display:grid;grid-template-columns:1fr 180px 1fr;gap:8px;margin-bottom:12px;}\n"
                "/* v26 mid layout */\n"
                ".mid-nopanel{grid-template-columns:180px 1fr;}")
    need(grid_anchor, 4)
    src = src.replace(grid_anchor, grid_add, 1)
    applied.append("4) analysis_panel винесено повним рядком над sm/pct")

# ════════════════════════════════════════════════════════════════
if src == orig:
    print("\nℹ️  Нічого не змінено — все вже застосовано.")
    for s in skipped: print("   ⏭ ", s)
    sys.exit(0)

bak = GEN.with_suffix(f".py.bak_v26_{datetime.now():%Y%m%d_%H%M%S}")
shutil.copy2(GEN, bak)
GEN.write_text(src, encoding="utf-8")

print("\n✅  Патч v26 застосовано.")
print(f"   Бекап: {bak.name}\n")
for s in applied: print("   ✓", s)
for s in skipped: print("   ⏭", s)

import py_compile
try:
    py_compile.compile(str(GEN), doraise=True)
    print("\n✅  generate.py компілюється без помилок.")
except py_compile.PyCompileError as e:
    print("\n❌  Синтаксична помилка:\n", e)
    shutil.copy2(bak, GEN)
    print("   Відкочено з бекапу. Надішли текст помилки.")
    sys.exit(1)

print("\n▶  Далі: python generate.py")