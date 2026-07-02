#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v17.py (v2) — патч для generate.py (після v16):
  1. SMALL TRADERS: MAX/MIN для CHG L / CHG S показували 0 —
     тепер st_cl/st_cs зчитуються і передаються у stats()
  2. Вертикальні кольорові лінії (синя/зелена/червона) у тижневій
     таблиці тепер ідуть через ВСЮ таблицю, а не лише заголовок.
     Для TFF і Disaggregated — аналогічно з їхніми кольорами.

Запуск:  python fix_v17.py   (у папці проєкту, поруч з generate.py)
"""

import shutil, sys, re
from pathlib import Path

GEN = Path(__file__).parent / "generate.py"
BAK = Path(__file__).parent / "generate_v16_backup.py"

if not GEN.exists():
    print(f"❌  Не знайдено {GEN}"); sys.exit(1)

src = GEN.read_text(encoding='utf-8')
shutil.copy(GEN, BAK)
print(f"💾  Бекап: {BAK.name}")

ok = True
def patch(name, old, new, count=1):
    global src, ok
    found = src.count(old)
    if found != count:
        print(f"❌  [{name}] знайдено {found} входжень (очікувалось {count}) — ПРОПУЩЕНО")
        ok = False
        return
    src = src.replace(old, new)
    print(f"✓   [{name}]")

def patch_re(name, pattern, new, count=1):
    """Заміна за регулярним виразом (стійка до пробілів)."""
    global src, ok
    matches = re.findall(pattern, src)
    if len(matches) != count:
        print(f"❌  [{name}] regex знайшов {len(matches)} входжень (очікувалось {count}) — ПРОПУЩЕНО")
        # діагностика: показати рядки з ключовим словом
        for ln in src.splitlines():
            if 'stats_st' in ln:
                print(f"    ↳ рядок у файлі: {ln.strip()[:120]}")
        ok = False
        return
    src = re.sub(pattern, new, src, count=count)
    print(f"✓   [{name}]")

# ================================================================
# 1) SMALL TRADERS — зчитуємо st_cl/st_cs і рахуємо MAX/MIN
# ================================================================

# 1a: якщо вже застосовано попереднім запуском — пропускаємо тихо
if "st_cl=gc(COL['st_cl'])" in src:
    print("✓   [1a] вже застосовано — пропускаю")
else:
    patch("1a: зчитування st_cl/st_cs у read_sheet",
        "        cm_cl=gc(COL['cm_cl']);cm_cs=gc(COL['cm_cs'])",
        "        cm_cl=gc(COL['cm_cl']);cm_cs=gc(COL['cm_cs'])\n"
        "        st_cl=gc(COL['st_cl']);st_cs=gc(COL['st_cs'])")

# 1b: гнучкий пошук 'stats_st': stats(st_net) з будь-якими пробілами
if re.search(r"'stats_st'\s*:\s*stats\(\s*st_net\s*,\s*st_cl", src):
    print("✓   [1b] вже застосовано — пропускаю")
else:
    patch_re("1b: stats_st тепер з CHG L / CHG S",
        r"'stats_st'\s*:\s*stats\(\s*st_net\s*\)",
        "'stats_st':stats(st_net,st_cl,st_cs)")

# ================================================================
# 2) ВЕРТИКАЛЬНІ ЛІНІЇ ЧЕРЕЗ ВСЮ ТАБЛИЦЮ
#    Колонки CHG L кожної групи: 2 (LS/AM/MM), 5 (CM/LEV/PM), 8 (ST/DL/SD)
# ================================================================

if 'table.ht tbody td:nth-child(2)' in src:
    print("✓   [2a] вже застосовано — пропускаю")
else:
    patch("2a: CSS вертикальні лінії у tbody",
        "table.ht .sep-r{border-right:1px solid var(--bd);}.sm-td{text-align:center;font-size:10px;padding:4px 6px;}",
        "table.ht .sep-r{border-right:1px solid var(--bd);}.sm-td{text-align:center;font-size:10px;padding:4px 6px;}\n"
        "table.ht tbody td:nth-child(2){border-left:2px solid rgba(74,158,255,.45);}\n"
        "table.ht tbody td:nth-child(5){border-left:2px solid rgba(32,212,131,.45);}\n"
        "table.ht tbody td:nth-child(8){border-left:2px solid rgba(240,81,90,.45);}\n"
        "table.ht[id^=\"tff_tbl_\"] tbody td:nth-child(5){border-left-color:rgba(240,180,41,.45);}\n"
        "table.ht[id^=\"tff_tbl_\"] tbody td:nth-child(8){border-left-color:rgba(32,212,131,.45);}\n"
        "table.ht[id^=\"dg_tbl_\"] tbody td:nth-child(2){border-left-color:rgba(167,139,250,.45);}\n"
        "table.ht[id^=\"dg_tbl_\"] tbody td:nth-child(5){border-left-color:rgba(32,212,131,.45);}\n"
        "table.ht[id^=\"dg_tbl_\"] tbody td:nth-child(8){border-left-color:rgba(240,180,41,.45);}")

# ================================================================
if not ok:
    print("\n⚠️  Не всі патчі застосовано — generate.py НЕ змінено, бекап лишається.")
    sys.exit(1)

GEN.write_text(src, encoding='utf-8')
print(f"\n✅  generate.py оновлено (v17). Запусти: python generate.py")