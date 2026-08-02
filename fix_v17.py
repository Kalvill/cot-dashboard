#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v38.py — вкладка TABLE: неоновий шрифт, автомасштаб під екран,
              фон для %-колонок, перероблені підсумкові рядки
(потрібні fix_v29 … fix_v37)

Зміни:
  1. Шрифт жирний + легке неонове світіння (text-shadow власним кольором).
  2. АВТОМАСШТАБ: після рендеру JS міряє сумарну ширину колонок до
     OPEN INTEREST включно і підбирає розмір шрифту так, щоб вони точно
     вміщались у видиму область. Працює на будь-якій ширині екрана
     (обмеження 9–26px), перераховується при зміні розміру вікна.
     Усі відступи переведено в em, тож масштабуються разом зі шрифтом.
  3. %NET/OI та %OI CHG (у т.ч. в групі OPEN INTEREST) отримали фон.
     Шкала — частка від «робастного максимуму» (90-й перцентиль |v| за
     всю історію), бо звичайний максимум у %OI CHG буває 18 873% і
     занулив би всі інші комірки.
  4. Фон NET POSITION зроблено менш вираженим: стеля 0.45 -> 0.32.
     Кольори не змінені.
  5. Підсумкові рядки:
       AVG 13W прибрано;
       MAX (ALL) — насичений зелений, MAX (5Y) — приглушений;
       MIN (ALL) — насичений червоний, MIN (5Y) — приглушений;
       усі клітинки рядка одного кольору, порожніх більше немає
       (ліміт форматування % піднято з 999% до 99 999%).
  6. Заголовки колонок (ДАТА, LONG, SHORT, NET POS …) стали білими.

Скрипт ідемпотентний. Запускати з папки проєкту:
    python fix_v38.py
"""
import shutil, sys
from pathlib import Path
from datetime import datetime

SRC = Path(__file__).parent / "generate.py"

OLD_CSS = """table.dt{border-collapse:separate;border-spacing:0;font-size:18px;white-space:nowrap;}
table.dt th{padding:7px 15px;font-weight:normal;font-size:14px;letter-spacing:.6px;
  text-align:right;background:var(--bg3);position:sticky;z-index:4;}
table.dt tr.tb-r1 th{top:0;height:38px;text-align:center;}
table.dt tr.tb-r2 th{top:38px;color:var(--d);border-bottom:1px solid var(--bd);}
table.dt th.tb-corner{left:0;z-index:6;top:0;text-align:left;color:var(--d);
  border-right:1px solid var(--bd);border-bottom:1px solid var(--bd);}
table.dt td{padding:6px 15px;text-align:right;border-bottom:1px solid rgba(52,61,90,.45);
  border-right:1px solid rgba(128,144,176,.10);color:#9fabc7;}
table.dt td.tb-date{position:sticky;left:0;z-index:3;background:var(--bg3);color:var(--d);
  text-align:left;border-right:1px solid var(--bd);}"""

NEW_CSS = """/* v38: розмір шрифту виставляє JS (tblFit), відступи в em — масштабуються разом */
table.dt{border-collapse:separate;border-spacing:0;font-size:18px;white-space:nowrap;}
table.dt th{padding:.42em .85em;font-weight:bold;font-size:.78em;letter-spacing:.6px;
  text-align:right;background:var(--bg3);position:sticky;z-index:4;
  text-shadow:0 0 4px currentColor;}
table.dt tr.tb-r1 th{top:0;text-align:center;}
table.dt tr.tb-r2 th{top:38px;color:#fff;border-bottom:1px solid var(--bd);}
table.dt th.tb-corner{left:0;z-index:6;top:0;text-align:left;color:#fff;
  border-right:1px solid var(--bd);border-bottom:1px solid var(--bd);}
table.dt td{padding:.34em .84em;text-align:right;border-bottom:1px solid rgba(52,61,90,.45);
  border-right:1px solid rgba(128,144,176,.10);color:#9fabc7;
  font-weight:bold;text-shadow:0 0 4px currentColor;}
table.dt td.tb-date{position:sticky;left:0;z-index:3;background:var(--bg3);color:var(--d);
  text-align:left;border-right:1px solid var(--bd);}"""

OLD_STATS_CSS = """table.dt tbody.tb-stats td{background:var(--bg3);font-size:16px;
  border-bottom:1px solid rgba(52,61,90,.9);}
table.dt tbody.tb-stats tr:last-child td{border-bottom:3px solid var(--bd);}
table.dt tbody.tb-stats td.tb-date{font-size:14px;letter-spacing:.5px;font-weight:bold;}
table.dt tbody.tb-stats tr.tb-dim td{opacity:.72;}"""

NEW_STATS_CSS = """table.dt tbody.tb-stats td{background:var(--bg3);font-size:.9em;
  border-bottom:1px solid rgba(52,61,90,.9);}
table.dt tbody.tb-stats tr:last-child td{border-bottom:3px solid var(--bd);}
table.dt tbody.tb-stats td.tb-date{font-size:.8em;letter-spacing:.5px;font-weight:bold;}"""

OLD_BADGE = ".tb-badge{font-size:13px;padding:2px 10px;border-radius:10px;font-weight:bold;}"
NEW_BADGE = ".tb-badge{font-size:.72em;padding:.12em .55em;border-radius:10px;font-weight:bold;}"

OLD_STATS_JS = """function tblStatsRows(d){
  const S=window._TBL_SPEC,N=d.d.length;
  const defs=[['MAX (ALL)',N,'max',''],['MIN (ALL)',N,'min',''],
              ['MAX (5Y)',Math.min(260,N),'max',' tb-dim'],
              ['MIN (5Y)',Math.min(260,N),'min',' tb-dim'],
              ['AVG 13W',Math.min(13,N),'avg',' tb-dim']];
  let html='';
  for(const[lbl,lim,mode,cls]of defs){
    let tds='<td class="tb-date">'+lbl+'</td>';
    for(let ci=0;ci<S.length;ci++){
      const k=S[ci].k,sep=S[ci].s?' tb-sep':'';
      if(k==='txt'){tds+='<td class="d'+sep+'">—</td>';continue;}
      const col=d.c[ci];let acc=null,sum=0,cnt=0;
      for(let i=0;i<lim;i++){
        const v=col[i];if(v==null)continue;
        if(mode==='max')acc=(acc==null||v>acc)?v:acc;
        else if(mode==='min')acc=(acc==null||v<acc)?v:acc;
        else{sum+=v;cnt++;}
      }
      if(mode==='avg')acc=cnt?Math.round(sum/cnt):null;
      let txt='—',c='d';
      if(acc!=null){
        if(k==='int'||k==='oi'||k==='grad'){txt=tblFmtInt(acc);c='t';}
        else if(k==='pct'){txt=tblFmtPct(acc);c=tblPctCls(acc);}
        else if(k==='cot'){txt=tblFmtCot(acc);c=acc>850?'g':(acc<150?'r':'t');}
        else if(k==='ratio'){txt=tblFmtRatio(acc);c=tblCls(acc);}
        else{txt=tblFmtSign(acc);c=tblCls(acc);}
      }
      const st=(k==='net')?' style="border-left:2px solid '+S[ci].c
                +';border-right:2px solid '+S[ci].c+'"':'';
      tds+='<td class="'+c+sep+'"'+st+'>'+txt+'</td>';
    }
    html+='<tr class="'+cls.trim()+'">'+tds+'</tr>';
  }
  return html;
}"""

NEW_STATS_JS = """// Підсумкові рядки. Кожен рядок цілком одного кольору:
// MAX(ALL) насичений зелений, MAX(5Y) приглушений; MIN — дзеркально червоним.
const TB_STAT_ROWS=[['MAX (ALL)',0,'max','#20d483'],
                    ['MIN (ALL)',0,'min','#f0515a'],
                    ['MAX (5Y)',260,'max','#7abba6'],
                    ['MIN (5Y)',260,'min','#c47884']];
function tblStatsRows(d){
  const S=window._TBL_SPEC,N=d.d.length;
  let html='';
  for(const[lbl,win,mode,rcol]of TB_STAT_ROWS){
    const lim=win?Math.min(win,N):N;
    let tds='<td class="tb-date" style="color:'+rcol+'">'+lbl+'</td>';
    for(let ci=0;ci<S.length;ci++){
      const k=S[ci].k,sep=S[ci].s?' tb-sep':'';
      const brd=(k==='net')?'border-left:2px solid '+S[ci].c
                +';border-right:2px solid '+S[ci].c+';':'';
      if(k==='txt'){
        tds+='<td class="'+sep.trim()+'" style="'+brd+'color:'+rcol+'">—</td>';continue;
      }
      const col=d.c[ci];let acc=null;
      for(let i=0;i<lim;i++){
        const v=col[i];if(v==null)continue;
        if(mode==='max')acc=(acc==null||v>acc)?v:acc;
        else acc=(acc==null||v<acc)?v:acc;
      }
      let txt='—';
      if(acc!=null){
        if(k==='int'||k==='oi'||k==='grad')txt=tblFmtInt(acc);
        else if(k==='pct')txt=tblFmtPct(acc);
        else if(k==='cot')txt=tblFmtCot(acc);
        else if(k==='ratio')txt=tblFmtRatio(acc);
        else txt=tblFmtSign(acc);
      }
      tds+='<td class="'+sep.trim()+'" style="'+brd+'color:'+rcol+'">'+txt+'</td>';
    }
    html+='<tr>'+tds+'</tr>';
  }
  return html;
}

// ── Автомасштаб: підбирає розмір шрифту так, щоб колонки до OPEN INTEREST
// включно вміщались у видиму ширину на будь-якому екрані ──
const TB_FIT_COLS=24, TB_FIT_BASE=18, TB_FIT_MIN=9, TB_FIT_MAX=26;
function tblFit(){
  const tbl=document.getElementById('dtTable'),sc=document.querySelector('.tb-scroll');
  if(!tbl||!sc)return;
  tbl.style.fontSize=TB_FIT_BASE+'px';
  const row=tbl.querySelector('tbody#dtBody tr')||tbl.querySelector('tbody tr');
  if(row){
    let w=0;const n=Math.min(TB_FIT_COLS,row.children.length);
    for(let i=0;i<n;i++)w+=row.children[i].getBoundingClientRect().width;
    const avail=sc.clientWidth-2;
    if(w>0&&avail>0){
      let fs=TB_FIT_BASE*avail/w;
      fs=Math.max(TB_FIT_MIN,Math.min(TB_FIT_MAX,fs));
      tbl.style.fontSize=fs.toFixed(2)+'px';
    }
  }
  // другий рядок шапки має липнути рівно під першим
  const h1=tbl.querySelector('tr.tb-r1 th');
  if(h1){
    const h=h1.getBoundingClientRect().height;
    tbl.querySelectorAll('tr.tb-r2 th').forEach(function(th){th.style.top=h+'px';});
  }
}
window.addEventListener('resize',function(){
  clearTimeout(window._tbFitT);
  window._tbFitT=setTimeout(tblFit,150);
});"""

EDITS = [
    ("CSS: неон + em-відступи + білі заголовки", "text-shadow:0 0 4px currentColor;", [
        (OLD_CSS, NEW_CSS)]),
    ("CSS: підсумкові рядки в em", "tbody.tb-stats td{background:var(--bg3);font-size:.9em;", [
        (OLD_STATS_CSS, NEW_STATS_CSS)]),
    ("CSS: бейджі в em", ".tb-badge{font-size:.72em;", [
        (OLD_BADGE, NEW_BADGE)]),

    ("ліміт форматування %", "Math.abs(f)>99999", [
        ("""function tblFmtPct(v){if(v==null)return'—';const f=v/10;
  if(!isFinite(f)||Math.abs(f)>999)return'—';return(f>0?'+':'')+f.toFixed(1)+'%';}""",
         """function tblFmtPct(v){if(v==null)return'—';const f=v/10;
  if(!isFinite(f)||Math.abs(f)>99999)return'—';return(f>0?'+':'')+f.toFixed(1)+'%';}""")]),

    ("фон NET менш виражений", "TB_POS_BG_MAX=0.32", [
        ("const TB_POS_BG_MAX=0.45, TB_POS_BG_GAMMA=1.3;",
         "const TB_POS_BG_MAX=0.32, TB_POS_BG_GAMMA=1.3;")]),

    ("фон для %-колонок", "function tblBgPct(", [
        ("""// %NET/OI та %OI CHG: >+30% зелений, <-30% червоний, решта — дефолт""",
         """// Фон для %NET/OI та %OI CHG. Шкала — частка від «робастного максимуму»
// (90-й перцентиль |v| за всю історію): звичайний максимум у %OI CHG
// буває 18873% і занулив би решту комірок.
const TB_PCT_BG_MAX=0.32;
function tblBgPct(v,rg){
  if(!rg||!rg.p90)return'';
  const o=TB_PCT_BG_MAX*Math.min(Math.abs(v)/rg.p90,1);
  if(o<0.012)return'';
  return'background:rgba('+(v>0?'32,212,131':'240,81,90')+','+o.toFixed(3)+');';
}
// %NET/OI та %OI CHG: >+30% зелений, <-30% червоний, решта — дефолт""")]),

    ("клітинка % з фоном", """+sp+'" style="'+tblBgPct(""", [
        ("""  if(kind==='pct')  return'<td class="'+tblPctCls(v)+sp+'">'+tblFmtPct(v)+'</td>';""",
         """  if(kind==='pct')  return'<td class="'+tblPctCls(v)+sp+'" style="'+tblBgPct(v,rg)
                          +'">'+tblFmtPct(v)+'</td>';""")]),

    ("діапазони для %-колонок", "p90:p90", [
        ("""    if(sp.k!=='net'&&sp.k!=='oi'&&sp.k!=='grad'&&sp.k!=='chg')return null;""",
         """    if(sp.k!=='net'&&sp.k!=='oi'&&sp.k!=='grad'&&sp.k!=='chg'&&sp.k!=='pct')return null;"""),
    ]),
    ("розрахунок p90", "asrt=null,p90=0", [
        ("""    let asrt=null;
    if(sp.k==='chg'){asrt=arr.map(function(x){return Math.abs(x);}).sort(function(a,b){return a-b;});}
    arr.sort((a,b)=>a-b);
    return{mn:arr[0],mx:arr[arr.length-1],srt:arr,mxa:mxa||1,asrt:asrt};""",
         """    let asrt=null,p90=0;
    if(sp.k==='chg'||sp.k==='pct'){
      asrt=arr.map(function(x){return Math.abs(x);}).sort(function(a,b){return a-b;});
      p90=asrt[Math.floor(0.9*(asrt.length-1))]||mxa;
    }
    arr.sort((a,b)=>a-b);
    return{mn:arr[0],mx:arr[arr.length-1],srt:arr,mxa:mxa||1,asrt:asrt,p90:p90};""")]),

    ("підсумкові рядки + автомасштаб", "function tblFit(){", [
        (OLD_STATS_JS, NEW_STATS_JS)]),

    ("виклик автомасштабу після рендеру", "requestAnimationFrame(tblFit)", [
        ("""  if(mt)mt.textContent=N+' тижнів  ·  '+(d.d[N-1]||'—')+' → '+(d.d[0]||'—')
        +'  ·  показано '+n;""",
         """  if(mt)mt.textContent=N+' тижнів  ·  '+(d.d[N-1]||'—')+' → '+(d.d[0]||'—')
        +'  ·  показано '+n;
  requestAnimationFrame(tblFit);""")]),
]


def main():
    if not SRC.exists():
        print(f"❌  Не знайдено {SRC}. Поклади fix_v38.py поруч із generate.py."); sys.exit(1)
    src = SRC.read_text(encoding='utf-8')
    print(f"📄  {SRC}  ({len(src)} символів)\n")

    if 'function tblStatsRows(' not in src:
        print("❌  Спочатку запусти fix_v29 … fix_v37."); sys.exit(1)

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

    bak = SRC.with_suffix(f".py.bak_v38_{datetime.now():%Y%m%d_%H%M%S}")
    shutil.copy2(SRC, bak)
    SRC.write_text(src, encoding='utf-8')
    print(f"\n💾  Бекап: {bak.name}")
    print("✅  generate.py оновлено\n")
    print("Далі:  python generate.py")


if __name__ == '__main__':
    main()