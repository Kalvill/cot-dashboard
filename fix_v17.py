#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v50.py — вкладка Overview: нумерація, фон OI/ACC, інверсія COT
(потрібні fix_v29 … fix_v49)

Зміни:
  1. Нумерація рядків 1, 2, 3 … у крайній лівій колонці.
     Номери проставляються після рендеру і після кожного сортування,
     тому вони не «їдуть» разом із рядками, а завжди йдуть по порядку
     зверху вниз. Групові заголовки (ВАЛЮТИ, МЕТАЛИ …) пропускаються.
     Індекси data-col усіх колонок зсунуто на 1, щоб сортування
     працювало як раніше.

  2. %OI CHG, ACC LS, ACC CM — фон прибрано повністю. Лишився
     звичайний колір за знаком. Тепер жодна колонка Overview не має
     суцільної заливки, крім слабкого тінту SM DIV.

  3. COT INDEX перевернуто, як у таблиці:
       < 20%  зелений  (екстремальний шорт — потенціал росту)
       > 80%  червоний (перегрів)
       решта  синій
     Змінено і в HTML, і в JS-перемикачі періодів, щоб при
     перемиканні Весь час / 3Y / 1Y / 6M / 3M логіка не збивалась.

Скрипт ідемпотентний. Запускати з папки проєкту:
    python fix_v50.py
"""
import shutil, sys
from pathlib import Path
from datetime import datetime

SRC = Path(__file__).parent / "generate.py"

# ── 1. нумерація: клітинка ─────────────────────────────────────
OLD_ROWSTART = """        rows_html.append(f'<tr class="ov-row">'
                         f'<td class="ov-asset">"""
NEW_ROWSTART = """        rows_html.append(f'<tr class="ov-row">'
                         f'<td class="ov-idx"></td>'
                         f'<td class="ov-asset">"""

OLD_GROUP = """rows_html.append(f'<tr class="ov-group"><td colspan="18">{item[1]}</td></tr>');continue"""
NEW_GROUP = """rows_html.append(f'<tr class="ov-group"><td colspan="19">{item[1]}</td></tr>');continue"""

# ── 2. шапка з нумерацією і зсунутими data-col ─────────────────
OLD_THEAD = """    thead=(f'<thead><tr>'
           f'<th class="ov-asset ov-sortable" data-col="0" data-stype="reset" onclick="ovSort(this)" title="Скинути сортування">ASSET</th>'
           f'<th class="ov-sortable" data-col="1" onclick="ovSort(this)">NET LS</th>'
           f'<th class="ov-sortable" data-col="2" onclick="ovSort(this)">NET CM</th>'
           f'<th class="ov-sortable" data-col="3" data-stype="cot" onclick="ovSort(this)" style="color:{COLOR_LS}">COT LS</th>'
           f'<th class="ov-sortable" data-col="4" data-stype="cot" onclick="ovSort(this)" style="color:{COLOR_CM}">COT CM</th>'
           f'<th class="ov-sortable" data-col="5" data-stype="cot" onclick="ovSort(this)" style="color:{COLOR_ST}">COT ST</th>'
           f'<th class="ov-sortable" data-col="6" onclick="ovSort(this)">SM DIV</th>'
           f'<th class="ov-sortable" data-col="7" onclick="ovSort(this)">CHG %LS</th>'
           f'<th class="ov-sortable" data-col="8" onclick="ovSort(this)">CHG LS</th>'
           f'<th class="ov-sortable" data-col="9" onclick="ovSort(this)">CHG %CM</th>'
           f'<th class="ov-sortable" data-col="10" onclick="ovSort(this)">CHG CM</th>'
           f'<th class="ov-sortable" data-col="11" data-stype="crowd" onclick="ovSort(this)">CM LEAD</th>'
           f'<th class="ov-sortable" data-col="12" onclick="ovSort(this)">%OI CHG</th>'
           f'<th class="ov-sortable" data-col="13" onclick="ovSort(this)">ACC LS</th>'
           f'<th class="ov-sortable" data-col="14" onclick="ovSort(this)">ACC CM</th>'
           f'<th class="ov-sortable" data-col="15" onclick="ovSort(this)">SM DIV 6M</th>'
           f'<th class="ov-sortable" data-col="16" onclick="ovSort(this)">SM DIV 3M</th>'
           f'<th class="ov-sortable" data-col="17" data-stype="crowd" onclick="ovSort(this)">CROWDED ATH</th>'
           f'</tr></thead>')"""

NEW_THEAD = """    thead=(f'<thead><tr>'
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
           f'</tr></thead>')"""

# ── 3. фон прибрано з OI / ACC ─────────────────────────────────
OLD_USE = """                         f'{pct_td_bg(d.get("oi_chg_pct"),_MX_OI)}'
                         f'{num_td_bg(d.get("acc_ls"),_MX_AL)}'
                         f'{num_td_bg(d.get("acc_cm"),_MX_AC)}'"""
NEW_USE = """                         f'{pct_td_plain(d.get("oi_chg_pct"))}'
                         f'{fnum_td(d.get("acc_ls"))}'
                         f'{fnum_td(d.get("acc_cm"))}'"""

OLD_FMT = """        def pct_td30(v):"""
NEW_FMT = """        def pct_td_plain(v):
            # відсоток без фону, колір за знаком
            if v is None: return '<td class="ov-num d">—</td>'
            try: f=float(v)*100
            except: return '<td class="ov-num d">—</td>'
            if abs(f)>9999: return '<td class="ov-num d">—</td>'
            cls='g' if f>0 else ('r' if f<0 else 'd')
            s2='+' if f>0 else ''
            return f'<td class="ov-num {cls}">{s2}{f:.1f}%</td>'
        def pct_td30(v):"""

# ── 4. інверсія COT ────────────────────────────────────────────
OLD_BAR = """            pct=min(max(v/100,0),1);color='#f0515a'if v<lo else('#20d483'if v>hi else'#4a9eff')"""
NEW_BAR = """            pct=min(max(v/100,0),1);color='#20d483'if v<20 else('#f0515a'if v>80 else'#4a9eff')"""

OLD_CLS = """            _cls = 'ov-cot-hi' if init>85 else ('ov-cot-lo' if init<15 else '')"""
NEW_CLS = """            _cls = 'ov-cot-lo' if init>80 else ('ov-cot-hi' if init<20 else '')"""

OLD_JSCOT = """    const color=v<15?'#f0515a':v>85?'#20d483':'#4a9eff';
    const cls=v>85?'ov-cot-hi':v<15?'ov-cot-lo':'';"""
NEW_JSCOT = """    const color=v<20?'#20d483':v>80?'#f0515a':'#4a9eff';
    const cls=v<20?'ov-cot-hi':v>80?'ov-cot-lo':'';"""

# ── 5. CSS + JS нумерації ──────────────────────────────────────
OLD_CSS = """.ov-zoom{display:flex;align-items:center;gap:3px;margin-left:18px;}"""
NEW_CSS = """.ov-idx{color:#5a6482;text-align:right;font-size:.72em;padding-right:.3em!important;}
.ov-idx-th{color:#5a6482;text-align:right;width:1px;white-space:nowrap;}
.ov-zoom{display:flex;align-items:center;gap:3px;margin-left:18px;}"""

OLD_JS = """// ── v48: масштаб таблиці Overview ──"""
NEW_JS = """// ── v50: наскрізна нумерація рядків (не залежить від сортування) ──
function ovRenumber(){
  let n=1;
  document.querySelectorAll('.ov-table tbody tr').forEach(function(tr){
    if(tr.classList.contains('ov-group'))return;
    const td=tr.querySelector('.ov-idx');
    if(td)td.textContent=n++;
  });
}
setTimeout(ovRenumber,70);

// ── v48: масштаб таблиці Overview ──"""

OLD_SORT1 = """    if(window.ovLoadFavs)ovLoadFavs();
    return;
  }"""
NEW_SORT1 = """    if(window.ovLoadFavs)ovLoadFavs();
    if(window.ovRenumber)ovRenumber();
    return;
  }"""

OLD_SORT2 = """  tbody.innerHTML='';
  rows.forEach(r=>tbody.appendChild(r));
  if(window.ovLoadFavs)ovLoadFavs();
}"""
NEW_SORT2 = """  tbody.innerHTML='';
  rows.forEach(r=>tbody.appendChild(r));
  if(window.ovLoadFavs)ovLoadFavs();
  if(window.ovRenumber)ovRenumber();
}"""

EDITS = [
    ("клітинка номера", 'f\'<td class="ov-idx"></td>\'', [(OLD_ROWSTART, NEW_ROWSTART)]),
    ("групи: 19 колонок", 'colspan="19"', [(OLD_GROUP, NEW_GROUP)]),
    ("шапка з нумерацією", '<th class="ov-idx-th">#</th>', [(OLD_THEAD, NEW_THEAD)]),
    ("форматер %OI CHG без фону", "def pct_td_plain(v):", [(OLD_FMT, NEW_FMT)]),
    ("прибрати фон з OI / ACC", "pct_td_plain(d.get(\"oi_chg_pct\"))", [(OLD_USE, NEW_USE)]),
    ("інверсія COT: смуга", "'#20d483'if v<20 else", [(OLD_BAR, NEW_BAR)]),
    ("інверсія COT: клас", "'ov-cot-lo' if init>80", [(OLD_CLS, NEW_CLS)]),
    ("інверсія COT: JS-перемикач", "v<20?'#20d483':v>80?'#f0515a'", [(OLD_JSCOT, NEW_JSCOT)]),
    ("CSS нумерації", ".ov-idx{color:#5a6482;", [(OLD_CSS, NEW_CSS)]),
    ("JS нумерації", "function ovRenumber(){", [(OLD_JS, NEW_JS)]),
    ("нумерація після скидання", "ovRenumber();\n    return;", [(OLD_SORT1, NEW_SORT1)]),
    ("нумерація після сортування", "if(window.ovRenumber)ovRenumber();\n}", [(OLD_SORT2, NEW_SORT2)]),
]


def main():
    if not SRC.exists():
        print(f"❌  Не знайдено {SRC}. Поклади fix_v50.py поруч із generate.py."); sys.exit(1)
    src = SRC.read_text(encoding='utf-8')
    print(f"📄  {SRC}  ({len(src)} символів)\n")

    if 'def make_overview_tab(' not in src:
        print("❌  У файлі немає make_overview_tab."); sys.exit(1)

    changed = False
    for name, guard, variants in EDITS:
        if guard in src:
            print(f"⏭  {name} — вже застосовано"); continue
        done = False
        for old, new in variants:
            if src.count(old) == 1:
                src = src.replace(old, new, 1); print(f"✓  {name}")
                done = True; changed = True; break
        if not done:
            print(f"❌  {name} — якір не знайдено"); sys.exit(1)

    if not changed:
        print("\n✅  Усе вже пропатчено, змін немає."); return

    try:
        compile(src, 'generate.py', 'exec')
    except SyntaxError as ex:
        print(f"\n❌  Синтаксична помилка: рядок {ex.lineno}: {ex.msg}\n    Файл НЕ змінено."); sys.exit(1)

    bak = SRC.with_suffix(f".py.bak_v50_{datetime.now():%Y%m%d_%H%M%S}")
    shutil.copy2(SRC, bak)
    SRC.write_text(src, encoding='utf-8')
    print(f"\n💾  Бекап: {bak.name}")
    print("✅  generate.py оновлено\n")
    print("Далі:  python generate.py")


if __name__ == '__main__':
    main()