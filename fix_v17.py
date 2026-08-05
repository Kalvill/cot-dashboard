#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v52.py — TFF: порядок груп, назва вкладки, шапка, міні-таблиця
(потрібні fix_v29 … fix_v51)

Зміни:
  1. Кнопка OVERVIEW перейменована на LEGACY.
  2. У режимі TFF з шапки прибрано «SM DIV —» і бейдж CROWDED —
     цих колонок у TFF-файлі немає, тож прочерки лише заважали.
  3. Порядок груп TFF: LEV MONEY -> ASSET MGR -> DEALER.
     NET-колонки підписані NET LEV / NET AM / NET DL.
  4. Таблиця «ТИЖНЕВА СТАТИСТИКА ПОЗИЦІЙ» на вкладці TFF Report
     інструмента замінена на верстку вкладки TABLE, обрізану по
     OPEN INTEREST: LEV MONEY, ASSET MGR, DEALER, OPEN INTEREST.

Про EXTRA (BB…BH): це не помилка рендеру. У COT_TFF_REPORTS.xlsx
ці сім стовпців справді не мають заголовків — рядки 3, 4 і 5 аркуша
для них порожні, і в аркуші info їх теж немає. Тому підписані літерами
Excel. Щоб дати назви, відредагуй TFF_EXTRA_LABELS у generate.py.

Скрипт ідемпотентний. Запускати з папки проєкту:
    python fix_v52.py
"""
import shutil, sys
from pathlib import Path
from datetime import datetime

SRC = Path(__file__).parent / "generate.py"

# ── 1. групи TFF: порядок і підписи NET ────────────────────────
OLD_GRP = """def _tff_grp(base, name, color):
    \"\"\"7 стандартних колонок групи TFF, base — індекс колонки Long\"\"\"
    return (name, color, [
        (base + 0, 'LONG', 'grad'), (base + 1, 'SHORT', 'grad'),
        (base + 2, 'CHG LONG', 'chg'), (base + 3, 'CHG SHORT', 'chg'),
        (base + 4, '%NET/OI', 'pct'), (base + 5, '%OI CHG', 'pct'),
        (base + 6, 'NET ' + name.split()[0][:3], 'net')])

TFF_TBL_GROUPS = [
    _tff_grp(2,  'DEALER',     TFF_COLOR_DL),
    _tff_grp(9,  'ASSET MGR',  TFF_COLOR_AM),
    _tff_grp(16, 'LEV MONEY',  TFF_COLOR_LEV),"""

NEW_GRP = """def _tff_grp(base, name, color, net_lbl):
    \"\"\"7 стандартних колонок групи TFF, base — індекс колонки Long\"\"\"
    return (name, color, [
        (base + 0, 'LONG', 'grad'), (base + 1, 'SHORT', 'grad'),
        (base + 2, 'CHG LONG', 'chg'), (base + 3, 'CHG SHORT', 'chg'),
        (base + 4, '%NET/OI', 'pct'), (base + 5, '%OI CHG', 'pct'),
        (base + 6, net_lbl, 'net')])

TFF_TBL_GROUPS = [
    _tff_grp(16, 'LEV MONEY',  TFF_COLOR_LEV, 'NET LEV'),
    _tff_grp(9,  'ASSET MGR',  TFF_COLOR_AM,  'NET AM'),
    _tff_grp(2,  'DEALER',     TFF_COLOR_DL,  'NET DL'),"""

# ── 2. назва кнопки ────────────────────────────────────────────
OLD_BTN = """'onclick=\"tblSrcSet(\\'leg\\',this)\">OVERVIEW</button>'"""
NEW_BTN = """'onclick=\"tblSrcSet(\\'leg\\',this)\">LEGACY</button>'"""

# ── 3. експорт міні-шапки TFF ──────────────────────────────────
OLD_HEADS = """    theadT, _mt,        _mc,       specT = _tbl_heads(TFF_TBL_GROUPS)"""
NEW_HEADS = """    theadT, mini_theadT, mini_colsT, specT = _tbl_heads(TFF_TBL_GROUPS)"""

OLD_EXPORT = """               'window._TBL_MINI_COLS=' + str(mini_cols) + ';'"""
NEW_EXPORT = """               'window._TBL_MINI_COLS=' + str(mini_cols) + ';'
               'window._TBL_MINI_THEAD_T=' + json.dumps(mini_theadT, ensure_ascii=False) + ';'
               'window._TBL_MINI_COLS_T=' + str(mini_colsT) + ';'"""

# ── 4. міні-таблиця TFF у вкладці інструмента ──────────────────
OLD_TFFVIEW = """    table_block=(f'<div class=\"htable-wrap\"><div class=\"htable-hdr\"><span>ТИЖНЕВА СТАТИСТИКА ПОЗИЦІЙ</span>'
                 f'<div class=\"hsel\">'
                 f'<button class=\"hbtn active\" data-n=\"10\" onclick=\"setTffHist(this,\\'{s}\\')\">10</button>'
                 f'<button class=\"hbtn\" data-n=\"26\" onclick=\"setTffHist(this,\\'{s}\\')\">26</button>'
                 f'<button class=\"hbtn\" data-n=\"52\" onclick=\"setTffHist(this,\\'{s}\\')\">52</button>'
                 f'</div></div><div class=\"htable-scroll\">{tbl}</div></div>')"""

NEW_TFFVIEW = """    # v52: верстка вкладки TABLE, обрізана по OPEN INTEREST
    table_block=(f'<div class=\"htable-wrap\"><div class=\"htable-hdr\"><span>ТИЖНЕВА СТАТИСТИКА ПОЗИЦІЙ</span>'
                 f'<div class=\"hsel\">'
                 f'<button class=\"hbtn active\" data-n=\"10\" onclick=\"setMiniHistT(this,\\'{s}\\')\">10</button>'
                 f'<button class=\"hbtn\" data-n=\"26\" onclick=\"setMiniHistT(this,\\'{s}\\')\">26</button>'
                 f'<button class=\"hbtn\" data-n=\"52\" onclick=\"setMiniHistT(this,\\'{s}\\')\">52</button>'
                 f'</div></div><div class=\"tb-scroll tb-mini\" id=\"mini_tff_{s}\"></div></div>')"""

# ── 5. JS ──────────────────────────────────────────────────────
OLD_JS_INFO = """function tblHeadInfo(d){
  const btn=document.getElementById('tblDashBtn');
  if(btn)btn.textContent=(tblNm()[_tblCur]||_tblCur)+' Dashboard';
  const sm=document.getElementById('tblSmDiv'),cw=document.getElementById('tblCrowd');"""

NEW_JS_INFO = """function tblHeadInfo(d){
  const btn=document.getElementById('tblDashBtn');
  if(btn)btn.textContent=(tblNm()[_tblCur]||_tblCur)+' Dashboard';
  // у TFF немає колонок SM DIV / CROWDED ATH — ховаємо блок цілком
  const info=document.querySelector('.tb-info');
  if(info)info.style.display=(_tblSrc==='tff'?'none':'');
  const sm=document.getElementById('tblSmDiv'),cw=document.getElementById('tblCrowd');"""

OLD_JS_MINI = """function setMiniHist(btn,sid){"""
NEW_JS_MINI = """// ── Міні-таблиця TFF (LEV MONEY / ASSET MGR / DEALER / OPEN INTEREST) ──
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
function setMiniHistT(btn,sid){
  btn.parentNode.querySelectorAll('.hbtn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  tblMiniRenderT(sid,parseInt(btn.dataset.n));
}
function setMiniHist(btn,sid){"""

OLD_SWITCH = """  } else if(type==='tff'){
    filterTffRows(sid,10);"""
NEW_SWITCH = """  } else if(type==='tff'){
    filterTffRows(sid,10);
    if(typeof tblMiniRenderT==='function')
      tblMiniRenderT(sid,(typeof _miniNT!=='undefined'&&_miniNT[sid])||10);"""

EDITS = [
    ("порядок груп TFF", "TFF_COLOR_LEV, 'NET LEV'", [(OLD_GRP, NEW_GRP)]),
    ("кнопка LEGACY", ">LEGACY</button>", [(OLD_BTN, NEW_BTN)]),
    ("міні-шапка TFF", "mini_theadT, mini_colsT", [(OLD_HEADS, NEW_HEADS)]),
    ("експорт міні-шапки TFF", "_TBL_MINI_THEAD_T=", [(OLD_EXPORT, NEW_EXPORT)]),
    ("міні-таблиця у TFF Report", 'id="mini_tff_{s}"', [(OLD_TFFVIEW, NEW_TFFVIEW)]),
    ("шапка без SM DIV у TFF", "if(info)info.style.display=", [(OLD_JS_INFO, NEW_JS_INFO)]),
    ("JS міні-таблиці TFF", "function tblMiniRenderT(", [(OLD_JS_MINI, NEW_JS_MINI)]),
    ("рендер при відкритті TFF", "typeof tblMiniRenderT==='function'", [(OLD_SWITCH, NEW_SWITCH)]),
]


def main():
    if not SRC.exists():
        print(f"❌  Не знайдено {SRC}. Поклади fix_v52.py поруч із generate.py."); sys.exit(1)
    src = SRC.read_text(encoding='utf-8')
    print(f"📄  {SRC}  ({len(src)} символів)\n")

    if 'TFF_TBL_GROUPS' not in src:
        print("❌  Спочатку запусти fix_v51.py."); sys.exit(1)

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

    if not changed:
        print("\n✅  Усе вже пропатчено, змін немає."); return

    try:
        compile(src, 'generate.py', 'exec')
    except SyntaxError as ex:
        print(f"\n❌  Синтаксична помилка: рядок {ex.lineno}: {ex.msg}\n    Файл НЕ змінено."); sys.exit(1)

    bak = SRC.with_suffix(f".py.bak_v52_{datetime.now():%Y%m%d_%H%M%S}")
    shutil.copy2(SRC, bak)
    SRC.write_text(src, encoding='utf-8')
    print(f"\n💾  Бекап: {bak.name}")
    print("✅  generate.py оновлено\n")
    print("Далі:  python generate.py")


if __name__ == '__main__':
    main()