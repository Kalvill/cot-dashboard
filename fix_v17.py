#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v28.py — COT INDEX gauge:
  • число більше (з 18px до ~21px при size=62), але ГАРАНТОВАНО в межах дуги
    — розмір шрифту підбирається автоматично під довжину числа
  • значення показується цілим (100 замість 100.0) — саме це дозволяє
    збільшити шрифт: "100.0" це 5 символів = 54px при fs=18, а корисна
    ширина всередині дуги лише ~39px
  • підпис (COT ALL / COT 3Y / COT 1Y) збільшено з 7.1px до ~9.6px

Запускати з папки проекту:
    python fix_v28.py

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
        print(f"❌  Крок {step}: anchor не знайдено:\n    {anchor[:100]}")
        print("    (спочатку має бути застосований fix_v27.py)")
        sys.exit(1)

# ════════════════════════════════════════════════════════════════
# КРОК 1 — Адаптивний розмір числа + більший підпис
# ════════════════════════════════════════════════════════════════
if "val_txt=" in src:
    skipped.append("1) Розміри тексту gauge (вже v28)")
else:
    old_sizes = """    val_fs=round(size*0.29,1)    # велике число в центрі
    lbl_fs=round(size*0.115,1)   # маленький підпис під дугою
    val_y =round(cy+val_fs*0.35,1)
    lbl_y =round(cy+r+lbl_fs*0.95,1)"""

    new_sizes = """    # v28: розмір числа підбирається так, щоб воно завжди вміщалось у дугу.
    # Ціле число (макс. 3 символи "100") дозволяє більший шрифт, ніж "100.0".
    val_txt=f"{value:.0f}"
    inner_w=2*(r-1.5-2.0)                        # діаметр мінус обведення і відступ
    fs_fit =inner_w/(max(len(val_txt),1)*0.60)   # Courier: ширина символу ~0.6em
    val_fs =round(min(size*0.34,fs_fit),1)       # більше ніж було, але без виходу за дугу
    lbl_fs =round(size*0.155,1)                  # більший підпис
    val_y  =round(cy+val_fs*0.35,1)
    lbl_y  =round(cy+r+lbl_fs*0.80,1)"""
    need(old_sizes, 1)
    src = src.replace(old_sizes, new_sizes, 1)
    applied.append("1) Число gauge — адаптивний розмір, завжди в межах дуги")
    applied.append("2) Підпис (COT ALL / 3Y / 1Y) збільшено")

# ════════════════════════════════════════════════════════════════
# КРОК 2 — Виводимо val_txt замість {value:.1f}
# ════════════════════════════════════════════════════════════════
if ">{val_txt}</text>" in src:
    skipped.append("3) Вивід цілого числа (вже v28)")
else:
    old_txt = 'font-weight="bold" fill="{color}">{value:.1f}</text>'
    new_txt = 'font-weight="bold" fill="{color}">{val_txt}</text>'
    need(old_txt, 2)
    src = src.replace(old_txt, new_txt, 1)
    applied.append("3) Значення виводиться цілим (100 замість 100.0)")

# ════════════════════════════════════════════════════════════════
if src == orig:
    print("\nℹ️  Нічого не змінено — все вже застосовано.")
    for s in skipped: print("   ⏭ ", s)
    sys.exit(0)

bak = GEN.with_suffix(f".py.bak_v28_{datetime.now():%Y%m%d_%H%M%S}")
shutil.copy2(GEN, bak)
GEN.write_text(src, encoding="utf-8")

print("\n✅  Патч v28 застосовано.")
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