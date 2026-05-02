#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ДІАГНОСТИКА COT_OVERVIEW.xlsx
Запусти цей скрипт і пришли мені весь вивід з консолі.
"""

import pandas as pd
from pathlib import Path

DATA_FILE   = Path(__file__).parent / "data" / "COT_OVERVIEW.xlsx"
SKIP_SHEETS = {'COT', 'info', 'lkupAUD'}

# ────────────────────────────────────
print("=" * 60)
print("  COT ДІАГНОСТИКА")
print("=" * 60)
print(f"\nФайл: {DATA_FILE}")
print(f"Існує: {DATA_FILE.exists()}")

if not DATA_FILE.exists():
    print("\n❌  Файл НЕ знайдено!")
    print("    Переклади COT_OVERVIEW.xlsx у папку data/")
    exit()

xl = pd.ExcelFile(DATA_FILE)

# ── 1. Список всіх вкладок ──────────────────────────────────
print(f"\n📋  ВКЛАДКИ у файлі ({len(xl.sheet_names)} шт.):")
for s in xl.sheet_names:
    skip_mark = " ← ПРОПУСКАЄМО" if s in SKIP_SHEETS else ""
    print(f"  {repr(s)}{skip_mark}")

# ── 2. Аналіз першої НЕ-пропущеної вкладки ─────────────────
data_sheets = [s for s in xl.sheet_names if s not in SKIP_SHEETS]
if not data_sheets:
    print("\n❌  Немає вкладок для читання!")
    exit()

test_name = data_sheets[0]          # Беремо першу (напр. "BRENT" або "EUR")
print(f"\n📊  АНАЛІЗ ВКЛАДКИ: '{test_name}'")
print("-" * 60)

raw = xl.parse(test_name, header=None)
print(f"Розмір: {raw.shape[0]} рядків × {raw.shape[1]} колонок\n")

# Виводимо перші 20 рядків (всі колонки)
print("Перші 20 рядків:")
for i in range(min(20, len(raw))):
    row_vals = [str(raw.iloc[i, j])[:18] for j in range(min(20, raw.shape[1]))]
    print(f"  Рядок {i+1:2d} | " + " | ".join(f"{v:18s}" for v in row_vals))

# ── 3. Пошук рядка з датами ────────────────────────────────
print(f"\n🔍  ПОШУК РЯДКІВ З ДАТАМИ в колонці A (col 0):")
for i in range(min(30, len(raw))):
    val = raw.iloc[i, 0]
    parsed = pd.to_datetime(val, errors='coerce')
    if pd.notna(parsed) and parsed.year > 2000:
        print(f"  → Рядок {i+1} (індекс {i}): '{val}' = {parsed.date()}")

# ── 4. Перевірка ще однієї вкладки (EUR або GBP) ───────────
second = next((s for s in data_sheets if s in ('EUR', 'GBP', 'GOLD')), None)
if second and second != test_name:
    print(f"\n📊  АНАЛІЗ ВКЛАДКИ: '{second}'")
    print("-" * 60)
    raw2 = xl.parse(second, header=None)
    print(f"Розмір: {raw2.shape}")
    for i in range(min(5, len(raw2))):
        row_vals = [str(raw2.iloc[i, j])[:18] for j in range(min(12, raw2.shape[1]))]
        print(f"  Рядок {i+1:2d} | " + " | ".join(f"{v:18s}" for v in row_vals))

print("\n" + "=" * 60)
print("  Скопіюй ВЕСЬ цей вивід і пришли мені!")
print("=" * 60)
