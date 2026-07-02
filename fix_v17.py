#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v20.py — патч для generate.py (після v19), вкладка Overview:
  1. Групи (ВАЛЮТИ / МЕТАЛИ / ІНДЕКСИ / ...) читаються з колонки A
     аркуша overview і показуються рядками-розділювачами.
     Рядки без даних (напр. ETH CASH) більше не ламають групування.
  2. Новий порядок колонок:
     NET LS | CHG LS | %OIΔ | NET CM | CHG CM | %OIΔ | COT LS/CM/ST | SM...
     (%OIΔ — це Chg % LS (кол.J) та Chg % CM (кол.L) з overview,
      ті самі значення що %OIΔ у тижневій таблиці)
     Стара єдина колонка %OI CHG прибрана.

Запуск:  python fix_v20.py   (у папці проєкту, поруч з generate.py)
"""

import shutil, sys
from pathlib import Path

GEN = Path(__file__).parent / "generate.py"
BAK = Path(__file__).parent / "generate_v19_backup.py"

if not GEN.exists():
    print(f"❌  Не знайдено {GEN}"); sys.exit(1)

src = GEN.read_text(encoding='utf-8')
shutil.copy(GEN, BAK)
print(f"💾  Бекап: {BAK.name}")

ok = True
def patch(name, old, new, count=1, skip_if=None):
    global src, ok
    if skip_if and skip_if in src:
        print(f"✓   [{name}] вже застосовано — пропускаю"); return
    found = src.count(old)
    if found != count:
        print(f"❌  [{name}] знайдено {found} входжень (очікувалось {count}) — ПРОПУЩЕНО")
        ok = False
        return
    src = src.replace(old, new)
    print(f"✓   [{name}]")

# ================================================================
# 1) ГРУПИ З КОЛОНКИ A + пропуск порожніх рядків
# ================================================================
patch("1: групи з колонки A",
    """        cur_group=''
        for i in range(4,len(raw)):
            row=raw.iloc[i]
            asset=str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            if not asset or asset=='nan': continue
            def safe(c):
                v=pd.to_numeric(row.iloc[c],errors='coerce'); return float(v) if pd.notna(v) else None
            cot_ls=safe(4)
            if cot_ls is None:
                cur_group=asset;OVERVIEW_TABLE.append(('_group',asset));continue
""",
    """        OV_GROUP_UA={'CURRENCIES':'ВАЛЮТИ','METALS':'МЕТАЛИ','METALAS':'МЕТАЛИ',
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
""",
    skip_if="OV_GROUP_UA")

# ================================================================
# 2) НОВІ ДАНІ У РЯДКУ ТАБЛИЦІ: Chg % LS (кол.9), Chg % CM (кол.11)
# ================================================================
patch("2a: chg_pct_ls/cm у OVERVIEW_TABLE",
    "'net_ls':safe(2),'net_cm':safe(3),",
    "'net_ls':safe(2),'net_cm':safe(3),'chg_pct_ls':safe(9),'chg_pct_cm':safe(11),")

patch("2b: helper fpct",
    """        def sm_fmt(v):
            if v is None: return '<span class="d">—</span>'
            cls='g'if float(v)>0 else('r'if float(v)<0 else'd'); return f'<span class="{cls}">{float(v):+.2f}</span>'""",
    """        def sm_fmt(v):
            if v is None: return '<span class="d">—</span>'
            cls='g'if float(v)>0 else('r'if float(v)<0 else'd'); return f'<span class="{cls}">{float(v):+.2f}</span>'
        def fpct(v):
            # частка 0..1 -> %; аномалії ховаємо
            if v is None: return '<span class="d">—</span>'
            v2=float(v)*100
            if abs(v2)>999: return '<span class="d">—</span>'
            cls='g'if v2>0 else('r'if v2<0 else'd')
            return f'<span class="{cls}">{v2:+.1f}%</span>'""",
    skip_if="def fpct(v):")

patch("2c: заголовки таблиці",
    """    thead=(f'<thead><tr><th class="ov-asset">ASSET</th><th>NET LS</th><th>CHG LS</th>'
           f'<th>NET CM</th><th>CHG CM</th><th>%OI CHG</th>'""",
    """    thead=(f'<thead><tr><th class="ov-asset">ASSET</th><th>NET LS</th><th>CHG LS</th><th>%OIΔ</th>'
           f'<th>NET CM</th><th>CHG CM</th><th>%OIΔ</th>'""")

patch("2d: клітинки рядка — новий порядок",
    """                         f'<td>{fnum(d["net_ls"],sign=True)}</td><td>{fnum(d["chg_ls"],sign=True)}</td>'
                         f'<td>{fnum(d["net_cm"],sign=True)}</td><td>{fnum(d["chg_cm"],sign=True)}</td>'
                         f'<td>{fnum(d["oi_chg_pct"],pct=True)}</td>'""",
    """                         f'<td>{fnum(d["net_ls"],sign=True)}</td><td>{fnum(d["chg_ls"],sign=True)}</td>'
                         f'<td>{fpct(d.get("chg_pct_ls"))}</td>'
                         f'<td>{fnum(d["net_cm"],sign=True)}</td><td>{fnum(d["chg_cm"],sign=True)}</td>'
                         f'<td>{fpct(d.get("chg_pct_cm"))}</td>'""")

patch("2e: colspan групи 12→13",
    '<tr class="ov-group"><td colspan="12">',
    '<tr class="ov-group"><td colspan="13">')

# ================================================================
if not ok:
    print("\n⚠️  Не всі патчі застосовано — generate.py НЕ змінено, бекап лишається.")
    sys.exit(1)

GEN.write_text(src, encoding='utf-8')
print(f"\n✅  generate.py оновлено (v20). Запусти: python generate.py")