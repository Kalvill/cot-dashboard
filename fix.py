#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_ranked.py — виправляє показ COT INDEX RANKED (M)

ПРОБЛЕМА:
  Дашборд показував '178217280000000000.0%' у блоці RANKED INDEX.
  Причина: функція gcm() у read_sheet() не мала захисту від
  ситуації, коли в колонці ranked (87-101) опинялось значення-дата.
  pd.to_numeric(дата) -> наносекунди (~1.78e15), * 100 -> ~1.78e17.
  У старому файлі колонок ranked ще не було, тож на цих позиціях
  стояли дати/інші дані -> сміттєве число потрапляло у _ci_m.

ВИПРАВЛЕННЯ:
  1. gcm() тепер приймає лише float у діапазоні 0..1 (частка) АБО 0..100
     (відсоток). Усе інше -> None. Дати/timestamp відсікаються.
  2. _make_ranked_section(): якщо всі значення None -> панель не
     показується (замість '50.0%'-заглушки).
  3. JS updateMPctBar(): додано перевірку val на скінченність і
     діапазон, щоб навіть аномалія не рендерилась.

Запуск:  python fix_ranked.py
(скрипт правит generate.py на місці, робить .bak резервну копію)
"""

from pathlib import Path
import shutil

GEN = Path(__file__).parent / "generate.py"

def main():
    if not GEN.exists():
        print(f"❌  Не знайдено {GEN}")
        return

    src = GEN.read_text(encoding="utf-8")
    shutil.copy(GEN, GEN.with_suffix(".py.bak"))
    print(f"💾  Резервна копія: {GEN.with_suffix('.py.bak')}")

    n = 0

    # ── ПАТЧ 1: захищена gcm() ──────────────────────────────────
    old_gcm = """        def gcm(ci):
            if ci>=df.shape[1]: return None
            try:
                v=pd.to_numeric(df.iloc[-1, ci], errors='coerce')  # df вже відсортований по даті
                return round(float(v)*100,1) if pd.notna(v) and float(v)>0 else None
            except: return None"""

    new_gcm = """        def gcm(ci):
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
            except: return None"""

    if old_gcm in src:
        src = src.replace(old_gcm, new_gcm)
        n += 1
        print("✓  Патч 1: gcm() захищено від дат/аномалій")
    else:
        print("⚠  Патч 1: gcm() не знайдено (можливо вже виправлено)")

    # ── ПАТЧ 2: _make_ranked_section приховати якщо порожньо ─────
    old_sect = '''def _make_ranked_section(s, cot_m):
    """Вбудована секція COT INDEX RANKED (M) всередині pct-panel"""
    if not cot_m: return \'\''''

    new_sect = '''def _make_ranked_section(s, cot_m):
    """Вбудована секція COT INDEX RANKED (M) всередині pct-panel"""
    if not cot_m: return \'\'
    # Якщо жодного валідного значення немає у всіх групах — не показуємо панель
    _has_any = any(
        (cot_m.get(g,{}) or {}).get(p) is not None
        for g in ('ls','cm','st')
        for p in ('all','3y','1y','6m','3m')
    )
    if not _has_any: return \'\''''

    if old_sect in src:
        src = src.replace(old_sect, new_sect)
        n += 1
        print("✓  Патч 2: _make_ranked_section приховується коли даних немає")
    else:
        print("⚠  Патч 2: _make_ranked_section не знайдено")

    # ── ПАТЧ 3: JS updateMPctBar — валідація val ────────────────
    old_js = """function updateMPctBar(sid){
  const mk=document.getElementById('pctmmk_'+sid);if(!mk)return;
  const panel=mk.closest('.pct-panel');if(!panel)return;
  const p=panel.querySelector('.psm.active')?.dataset.p||'ls';
  const per=panel.querySelector('.ppm.active')?.dataset.per||'all';
  const val=(_ci_m[sid]?.[p]?.[per])??50;
  const pos=Math.min(Math.max(val,0),100);"""

    new_js = """function updateMPctBar(sid){
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
  const pos=Math.min(Math.max(val,0),100);"""

    if old_js in src:
        src = src.replace(old_js, new_js)
        n += 1
        print("✓  Патч 3: JS updateMPctBar валідує значення")
    else:
        print("⚠  Патч 3: JS updateMPctBar не знайдено")

    GEN.write_text(src, encoding="utf-8")
    print(f"\n✅  Застосовано патчів: {n}/3  →  {GEN}")
    print("   Далі: запусти generate.py щоб перегенерувати output/cot_dashboard.html")

if __name__ == "__main__":
    main()