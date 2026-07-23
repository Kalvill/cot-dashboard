#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_v23.py — 3 зміни:
  1) Зірочка: підсвічується тільки сама зірка, рядок не змінює колір
  2) SM DIV графік: назва активу зелена при >0.8 / червона при <-0.8
     + ледь помітні горизонтальні лінії на рівнях +0.8 та -0.8
  3) Overview: клік по назві активу -> перехід на вкладку цього активу (Legacy)

Запускати з папки проекту:
    python fix_v23.py

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
# КРОК 1 — Зірочка підсвічує тільки себе
# ════════════════════════════════════════════════════════════════
if "/* v23 fav */" in src:
    skipped.append("1) Зірочка (вже v23)")
else:
    # 1a. прибираємо CSS підсвітки рядка
    old_css = """.ov-table tr.ov-faved td{background:rgba(240,180,41,.06);}
.ov-table tr.ov-faved .ov-asset{color:#f0b429;}"""
    new_css = """/* v23 fav — підсвічується тільки зірка, рядок без змін */"""
    need(old_css, 1)
    src = src.replace(old_css, new_css, 1)

    # 1b. прибираємо додавання класу рядку в JS
    old_js = """function ovApplyFav(el,on){
  el.classList.toggle('on',on);
  el.textContent=on?'★':'☆';
  const tr=el.closest('tr');
  if(tr) tr.classList.toggle('ov-faved',on);
}"""
    new_js = """function ovApplyFav(el,on){
  el.classList.toggle('on',on);
  el.textContent=on?'★':'☆';
}"""
    need(old_js, 1)
    src = src.replace(old_js, new_js, 1)
    applied.append("1) Зірочка — підсвічується тільки вона, рядок без підсвітки")

# ════════════════════════════════════════════════════════════════
# КРОК 2 — Клік по активу -> вкладка активу (Legacy)
# ════════════════════════════════════════════════════════════════
if "ovGoInstrument" in src:
    skipped.append("2) Клік по активу (вже є)")
else:
    # 2a. робимо назву активу клікабельною (зірка окремо, щоб не конфліктувала)
    old_asset = """                         f'<td class="ov-asset"><span class="ov-fav" data-fav="{d["sid"]}" onclick="ovToggleFav(this)">☆</span>{d["asset"]}</td>'"""
    new_asset = """                         f'<td class="ov-asset"><span class="ov-fav" data-fav="{d["sid"]}" onclick="event.stopPropagation();ovToggleFav(this)">☆</span>'
                         f'<span class="ov-asset-link" onclick="ovGoInstrument(\\'{d["sid"]}\\')">{d["asset"]}</span></td>'"""
    need(old_asset, 2)
    src = src.replace(old_asset, new_asset, 1)

    # 2b. CSS для клікабельної назви
    css_anchor2 = """/* v23 fav — підсвічується тільки зірка, рядок без змін */"""
    css_add2 = css_anchor2 + """
.ov-asset-link{cursor:pointer;transition:color .15s;}
.ov-asset-link:hover{color:var(--accent);text-decoration:underline;}"""
    need(css_anchor2, 2)
    src = src.replace(css_anchor2, css_add2, 1)

    # 2c. JS-функція переходу
    js_anchor2 = "// ── v22: Overview favorites (зірочки) ──"
    js_add2 = """// ── v23: Overview -> перехід на вкладку інструмента ──
function ovGoInstrument(sid){
  // знаходимо кнопку інструмента серед усіх категорій
  let btn=null, cat=null;
  document.querySelectorAll('.itab[data-i]').forEach(b=>{
    if(btn) return;
    const k=b.dataset.i;
    const s=k.replaceAll(' ','_').replaceAll('&','n').replaceAll('/','_');
    if(s===sid){btn=b;cat=b.dataset.cat;}
  });
  if(!btn){console.warn('Інструмент не знайдено:',sid);return;}
  selMain('cot');           // перемикаємось на COT Dashboard
  selCat(cat);              // відкриваємо потрібну категорію
  selInst(cat,btn.dataset.i); // відкриваємо інструмент (Legacy за замовчуванням)
  window.scrollTo({top:0,behavior:'smooth'});
}

// ── v22: Overview favorites (зірочки) ──"""
    need(js_anchor2, 2)
    src = src.replace(js_anchor2, js_add2, 1)
    applied.append("2) Клік по назві активу -> вкладка інструмента (Legacy)")

# ════════════════════════════════════════════════════════════════
# КРОК 3 — SM DIV графік: кольорові підписи + лінії ±0.8
# ════════════════════════════════════════════════════════════════
if "_SM_THRESHOLD" in src:
    skipped.append("3) SM DIV пороги (вже є)")
else:
    old_chart = """  _ovSmChart=new Chart(cv.getContext('2d'),{
    type:'bar',
    data:{labels,datasets:[{data:vals,backgroundColor:colors,borderColor:bdrColors,borderWidth:1,borderRadius:2,label:titles[key]||'SM DIV'}]},
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      plugins:{legend:{display:false},
        title:{display:true,text:titles[key]||'SM DIV',color:'#dde2ee',font:{family:'Courier New',size:11},padding:{bottom:8}},
        tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',
          titleFont:{family:'Courier New',size:10},bodyFont:{family:'Courier New',size:10},
          callbacks:{label:ctx=>{const v=ctx.parsed.y;return ' '+v.toFixed(2);}}}},
      scales:{
        x:{ticks:{color:'#8090b0',font:{family:'Courier New',size:9},maxRotation:45,minRotation:30},
           grid:{color:'rgba(52,61,90,.5)',lineWidth:.5},border:{display:false}},
        y:{display:true,grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},
           ticks:{color:'#8090b0',font:{family:'Courier New',size:9},callback:v=>v.toFixed(2)},
           border:{display:false},
           min:-1,max:1}}}});"""

    new_chart = """  const _SM_THRESHOLD=0.8;
  // підпис активу: зелений якщо >0.8, червоний якщо <-0.8
  const tickColors=vals.map(v=>v>_SM_THRESHOLD?'#20d483':(v<-_SM_THRESHOLD?'#f0515a':'#8090b0'));
  // ледь помітні горизонтальні лінії на ±0.8
  const thresholdLines={
    id:'smThresholdLines',
    beforeDatasetsDraw(chart){
      const {ctx,chartArea,scales}=chart;
      if(!chartArea||!scales.y) return;
      ctx.save();
      ctx.strokeStyle='rgba(221,226,238,.13)';
      ctx.lineWidth=1;
      ctx.setLineDash([4,4]);
      [_SM_THRESHOLD,-_SM_THRESHOLD].forEach(lv=>{
        const y=scales.y.getPixelForValue(lv);
        ctx.beginPath();
        ctx.moveTo(chartArea.left,y);
        ctx.lineTo(chartArea.right,y);
        ctx.stroke();
      });
      ctx.restore();
    }
  };
  _ovSmChart=new Chart(cv.getContext('2d'),{
    type:'bar',
    data:{labels,datasets:[{data:vals,backgroundColor:colors,borderColor:bdrColors,borderWidth:1,borderRadius:2,label:titles[key]||'SM DIV'}]},
    plugins:[thresholdLines],
    options:{responsive:true,maintainAspectRatio:false,animation:false,
      plugins:{legend:{display:false},
        title:{display:true,text:titles[key]||'SM DIV',color:'#dde2ee',font:{family:'Courier New',size:11},padding:{bottom:8}},
        tooltip:{backgroundColor:'#21263a',borderColor:'#343d5a',borderWidth:1,titleColor:'#dde2ee',bodyColor:'#dde2ee',
          titleFont:{family:'Courier New',size:10},bodyFont:{family:'Courier New',size:10},
          callbacks:{label:ctx=>{const v=ctx.parsed.y;return ' '+v.toFixed(2);}}}},
      scales:{
        x:{ticks:{color:tickColors,font:{family:'Courier New',size:9},maxRotation:45,minRotation:30},
           grid:{color:'rgba(52,61,90,.5)',lineWidth:.5},border:{display:false}},
        y:{display:true,grid:{color:'rgba(52,61,90,.6)',lineWidth:.5},
           ticks:{color:'#8090b0',font:{family:'Courier New',size:9},callback:v=>v.toFixed(2)},
           border:{display:false},
           min:-1,max:1}}}});"""
    need(old_chart, 3)
    src = src.replace(old_chart, new_chart, 1)
    applied.append("3) SM DIV — кольорові підписи (±0.8) + пунктирні лінії порогів")

# ════════════════════════════════════════════════════════════════
if src == orig:
    print("\nℹ️  Нічого не змінено — все вже застосовано.")
    for s in skipped: print("   ⏭ ", s)
    sys.exit(0)

bak = GEN.with_suffix(f".py.bak_v23_{datetime.now():%Y%m%d_%H%M%S}")
shutil.copy2(GEN, bak)
GEN.write_text(src, encoding="utf-8")

print("\n✅  Патч v23 застосовано.")
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