#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v54.py — вкладка Overview: перемикач LEGACY / TFF / DISAGGREGATED
(потрібні fix_v29 … fix_v53)

У COT_TFF_REPORTS.xlsx з'явився аркуш «Overview», у
COT_DISAGRAGATE_REPORTS.xlsx — «OVERVIEW». Структура в обох однакова
(заголовки в рядку 4, дані з рядка 5):

    A  #            H  COT (група 3)
    B  ASSET        I  Chg % (група 1)
    C  NET (гр. 1)  J  Chg   (група 1)
    D  NET (гр. 2)  K  Chg % (група 2)
    E  NET (гр. 3)  L  Chg   (група 2)
    F  COT (гр. 1)  M  Chg % (група 3)
    G  COT (гр. 2)  N  Chg   (група 3)
                    O  % OI Chg

Групи: TFF — LevMon / A.MGR / DLRS, Disaggregated — MM / PROD / DLRS.
Підписи колонок читаються прямо з рядка 4, тож якщо ти їх переназвеш
в Excel — вони підтягнуться самі.

Форматування те саме, що в Legacy Overview:
  NET, Chg     — колір за знаком, без фону
  COT          — смуга + значення, <20% зелений, >80% червоний
  Chg %        — >+30% зелений, <-30% червоний, решта сірий
  % OI Chg     — колір за знаком
  нумерація, сортування по колонках і масштаб — спільні

Скрипт ідемпотентний. Запускати з папки проєкту:
    python fix_v54.py
"""
import shutil, sys
from pathlib import Path
from datetime import datetime

SRC = Path(__file__).parent / "generate.py"

OV2_BLOCK = '''
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
    body = f"{abs(nv):,}".replace(',', '\\u202f')
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
               f'onclick="ovGoTable(\\'{sid(r["asset"])}\\')">{r["asset"]}</span></td>']
        for ix, kind, _gi in order:
            v = r['v'][ix] if ix < len(r['v']) else None
            tds.append(_ov2_cot(v) if kind == 'cot' else
                       _ov2_pct30(v) if kind == 'p30' else
                       _ov2_pct(v) if kind == 'pct' else _ov2_num(v))
        body.append('<tr class="ov-row">' + ''.join(tds) + '</tr>')

    return ('<div class="ov-scroll"><table class="ov-table">' + thead
            + '<tbody>' + ''.join(body) + '</tbody></table></div>')


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
        f'onclick="ovSrcSet(\\'{k}\\',this)">{t}</button>'
        for i, (k, t, _h) in enumerate(secs))
    panes = ''.join(
        f'<div class="ov-sec{" active" if i == 0 else ""}" id="ovsec_{k}">{h}</div>'
        for i, (k, _t, h) in enumerate(secs))
    return f'<div class="ov-src">{btns}</div>{panes}'

'''

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

EDITS = [
    ("блок Overview 2 (TFF / DISAG)", "def make_overview_all(", [
        ("\ndef fnum_td(v):", OV2_BLOCK + "\ndef fnum_td(v):")]),

    ("виклик make_overview_all", "ov_html=make_overview_all()", [
        ("    ov_html=make_overview_tab()", "    ov_html=make_overview_all()")]),

    ("CSS + JS перемикача", "function ovSrcSet(src,btn){", [
        ("+AUTH_MODAL_HTML+HTML_FOOT)",
         "+AUTH_MODAL_HTML+OV2_CSS+OV2_JS+HTML_FOOT)")]),

    # ── спільні JS-функції мають працювати з кількома таблицями ──
    ("масштаб для всіх таблиць Overview", "querySelectorAll('.ov-table').forEach(function(t)", [
        ("""  const t=document.querySelector('.ov-table');
  const fs=Math.round((_ovZoom!=null&&isFinite(_ovZoom))?_ovZoom:OV_ZOOM_BASE);
  if(t)t.style.fontSize=fs+'px';""",
         """  const fs=Math.round((_ovZoom!=null&&isFinite(_ovZoom))?_ovZoom:OV_ZOOM_BASE);
  document.querySelectorAll('.ov-table').forEach(function(t){t.style.fontSize=fs+'px';});"""),
    ]),
    ("нумерація в межах кожної таблиці", "document.querySelectorAll('.ov-table').forEach(function(tb){", [
        ("""function ovRenumber(){
  let n=1;
  document.querySelectorAll('.ov-table tbody tr').forEach(function(tr){
    if(tr.classList.contains('ov-group'))return;
    const td=tr.querySelector('.ov-idx');
    if(td)td.textContent=n++;
  });
}""",
         """function ovRenumber(){
  document.querySelectorAll('.ov-table').forEach(function(tb){
    let n=1;
    tb.querySelectorAll('tbody tr').forEach(function(tr){
      if(tr.classList.contains('ov-group'))return;
      const td=tr.querySelector('.ov-idx');
      if(td)td.textContent=n++;
    });
  });
}"""),
    ]),
    ("сортування: свій порядок для кожної таблиці", "table._ovOrig", [
        ("  if(!_ovOrigRows) _ovOrigRows=Array.from(tbody.children);",
         "  if(!table._ovOrig) table._ovOrig=Array.from(tbody.children);"),
    ]),
    ("скидання сортування", "table._ovOrig.forEach(r=>tbody.appendChild(r));", [
        ("    _ovOrigRows.forEach(r=>tbody.appendChild(r));",
         "    table._ovOrig.forEach(r=>tbody.appendChild(r));"),
    ]),
    ("вибірка рядків для сортування", "const rows=table._ovOrig.filter(", [
        ("  const rows=_ovOrigRows.filter(r=>!r.classList.contains('ov-group'));",
         "  const rows=table._ovOrig.filter(r=>!r.classList.contains('ov-group'));"),
    ]),
]


def main():
    if not SRC.exists():
        print(f"❌  Не знайдено {SRC}. Поклади fix_v54.py поруч із generate.py."); sys.exit(1)
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
            print(f"❌  {name} — якір не знайдено (входжень: {src.count(variants[0][0])})"); sys.exit(1)

    # оголошення OV2_CSS / OV2_JS
    if 'OV2_CSS = """' not in src:
        src = src.replace("\ndef make_overview_all(",
                          '\nOV2_CSS = """' + OV2_CSS + '"""\n\nOV2_JS = """' + OV2_JS +
                          '"""\n\n\ndef make_overview_all(', 1)
        print("✓  стилі та скрипт перемикача")
        changed = True

    if not changed:
        print("\n✅  Усе вже пропатчено, змін немає."); return

    try:
        compile(src, 'generate.py', 'exec')
    except SyntaxError as ex:
        print(f"\n❌  Синтаксична помилка: рядок {ex.lineno}: {ex.msg}\n    Файл НЕ змінено."); sys.exit(1)

    bak = SRC.with_suffix(f".py.bak_v54_{datetime.now():%Y%m%d_%H%M%S}")
    shutil.copy2(SRC, bak)
    SRC.write_text(src, encoding='utf-8')
    print(f"\n💾  Бекап: {bak.name}")
    print("✅  generate.py оновлено\n")
    print("Далі:  python generate.py")


if __name__ == '__main__':
    main()