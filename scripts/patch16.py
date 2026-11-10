# -*- coding: utf-8 -*-
# Этап 16 → v16.0: портфельная экспертная матрица «I_А × I_Ц» на вкладке «Графики БД».
# Мастер-патч: перезапускаемый, все замены с assert count==N. НЕ менять порядок блоков.
import io, sys

P = "output/Аудит_НИОКР_интерактивная_модель.html"
s = io.open(P, encoding="utf-8").read()

def rep(old, new, n=1):
    global s
    c = s.count(old)
    assert c == n, "MISS(%d!=%d): %r" % (c, n, old[:80])
    s = s.replace(old, new)

# ---------- 1. версии в шапке/подвалах ----------
rep("интерактивная модель оценки, v15.0</title>", "интерактивная модель оценки, v16.0</title>")
rep('v15.0 · этап 15', 'v16.0 · этап 16')
rep("аудита НИОКР v15.0", "аудита НИОКР v16.0", 8)  # подвалы актов/отчётов/дека/печати + KB чек-лист
rep("НИОКР_аудит_v15.json", "НИОКР_аудит_v16.json", 2)

# ---------- 2. welcome-модалка: h2 + абзац «Новое в v16.0» ----------
rep("Добро пожаловать в модель аудита НИОКР (v15.0)</h2>", "Добро пожаловать в модель аудита НИОКР (v16.0)</h2>")
V16P = ("<p><b>Новое в v16.0 — портфельная экспертная матрица «Актуальность × Целесообразность»:</b> "
        "на вкладке «📉 Графики БД» — вся база НИОКР как точки в матрице решений 5×5 "
        "(X — индекс целесообразности I_Ц, Y — индекс актуальности I_А; оценки не складываются — позиция определяется парой индексов). "
        "Цвет ячейки — группа рекомендации (продолжение · корректировка · пересмотр · приостановка · прекращение), "
        "наведение на ячейку — полный текст рекомендации, <b>красное кольцо — сработавший стоп-фактор</b> "
        "(рекомендация заблокирована → внеочередное рассмотрение руководством), клик по точке открывает НИОКР, "
        "легенда — со счётчиками по группам. Общие фильтры выборки действуют на матрицу; она включается в «🖨 Печать / PDF» дашборда. "
        "Внешний инструмент: в Excel-файле «7_Оценка_актуальности_целесообразности.xlsx» добавлен лист «Сравнение НИОКР» — "
        "до 10 разработок: индексы I_А/I_Ц, авто-уровни, рекомендации по матрице решений, перспективы, группы и сводные счётчики.</p>")
rep("<p><b>Новое в v15.0 — экспертная методика", V16P + "<p><b>Новое в v15.0 — экспертная методика")

# ---------- 3. KB: описание вкладки «Графики БД» и строка хоткея v ----------
rep("тепловая карта «НИОКР × блоки». Общие фильтры",
    "тепловая карта «НИОКР × блоки», экспертная матрица портфеля «I_А × I_Ц» (v16.0): все НИОКР базы — точки в матрице решений 5×5 "
    "(цвет ячейки — группа рекомендации, красное кольцо — стоп-фактор, клик открывает НИОКР). Общие фильтры")
rep("тепловая карта, фильтры, печать)", "тепловая карта, матрица А×Ц, фильтры, печать)")

# ---------- 4. новые функции: recGroup + expertMatrixCard ----------
FN = '''/* ============ ЭКСПЕРТНАЯ МАТРИЦА ПОРТФЕЛЯ «I_А × I_Ц» (этап 16, v16.0) ============ */
const RECGROUPS=[["Продолжение","#2e7d32","#e8f5e9"],["Корректировка","#7cb342","#f1f8e9"],["Пересмотр","#f9a825","#fffde7"],["Приостановка","#ef6c00","#fff3e0"],["Прекращение","#c62828","#ffebee"]];
function recGroup(txt){const t=String(txt||"");
 if(/прекрат|прекращ/i.test(t))return"Прекращение";
 if(/приостанов/i.test(t))return"Приостановка";
 if(/пересмотр|проверить/i.test(t))return"Пересмотр";
 if(/корректировкой|устранения|уточнения/i.test(t))return"Корректировка";
 if(/продолжить/i.test(t))return"Продолжение";
 return"Пересмотр";}
function expertMatrixCard(vis){
 const pts=[],nod=[];
 vis.forEach(idx=>{const r=state.list[idx],c=calc(r);
  if(c.iA==null||c.iC==null){nod.push(r.shifr);return;}
  pts.push({idx,r,c,st:stopsOn(r).length>0,rec:expertRec(r)});});
 let h='<div class="card" style="margin-top:12px"><h3>Экспертная матрица портфеля «Актуальность × Целесообразность» <span class="mini">(v16.0)</span></h3>'+
 '<p class="mini">Каждая НИОКР — точка: X = индекс целесообразности I_Ц, Y = индекс актуальности I_А. Оценки А и Ц <b>не складываются</b>: ячейку определяет пара индексов; цвет ячейки — группа рекомендации матрицы решений (наведите на ячейку — полный текст). Красное кольцо — сработал стоп-фактор (рекомендация заблокирована → внеочередное рассмотрение руководством). Клик по точке открывает НИОКР. Действуют общие фильтры выборки.</p>';
 if(!vis.length){return h+'<p class="mini">Нет НИОКР под выбранными фильтрами.</p></div>';}
 const W=660,H=560,pl=56,pr=14,ptop=26,pb=58;
 const X=v=>pl+(v/100)*(W-pl-pr),Y=v=>ptop+(1-(v/100))*(H-ptop-pb);
 const TH=[0,30,50,70,85,100];
 let s='<svg viewBox="0 0 '+W+' '+H+'" style="width:100%" role="img" aria-label="Экспертная матрица портфеля: актуальность × целесообразность">';
 for(let a=0;a<5;a++)for(let k=0;k<5;k++){
  const rec=RECMX[a][4-k],g=recGroup(rec),pal=RECGROUPS.find(x=>x[0]===g)||RECGROUPS[2];
  const x1=X(TH[k]),x2=X(TH[k+1]),y1=Y(TH[5-a]),y2=Y(TH[4-a]);
  s+='<rect x="'+x1.toFixed(1)+'" y="'+y1.toFixed(1)+'" width="'+(x2-x1).toFixed(1)+'" height="'+(y2-y1).toFixed(1)+'" fill="'+pal[2]+'" stroke="#fff" stroke-width="1.5"><title>А: '+LVL5N[a]+' · Ц: '+LVL5N[4-k]+' → '+rec+'</title></rect>';}
 const LVL5SHORT=["Высокая","Хорошая","Удовл.","Низкая","Критич."];
 for(let a=0;a<5;a++){const ym=(Y(TH[4-a])+Y(TH[5-a]))/2;
  s+='<text x="'+(pl-8)+'" y="'+ym.toFixed(1)+'" font-size="9.5" text-anchor="end" fill="#666" dominant-baseline="middle">'+LVL5SHORT[a]+'</text>';}
 for(let k=0;k<5;k++){const xm=(X(TH[k])+X(TH[k+1]))/2;
  s+='<text x="'+xm.toFixed(1)+'" y="'+(H-pb+18)+'" font-size="9.5" text-anchor="middle" fill="#666">'+LVL5N[4-k]+'</text>';}
 [0,25,50,75,100].forEach(t=>{s+='<text x="'+X(t).toFixed(1)+'" y="'+(H-pb+34)+'" font-size="8.5" text-anchor="middle" fill="#aaa">'+t+'</text>';});
 s+='<text x="'+((pl+W-pr)/2).toFixed(0)+'" y="'+(H-6)+'" font-size="11" text-anchor="middle" fill="#333">Индекс целесообразности I_Ц, % →</text>';
 s+='<text x="14" y="'+((ptop+H-pb)/2).toFixed(0)+'" font-size="11" text-anchor="middle" fill="#333" transform="rotate(-90 14 '+((ptop+H-pb)/2).toFixed(0)+')">← Индекс актуальности I_А, %</text>';
 const seen={},OFF=[[0,0],[8,-6],[-8,-6],[8,7],[-8,7],[0,10],[0,-10],[11,0],[-11,0],[14,8],[-14,8]];
 pts.forEach(p=>{const key=Math.round(p.c.iA)+"_"+Math.round(p.c.iC);const n=seen[key]=(seen[key]||0)+1;
  const off=OFF[Math.min(n-1,OFF.length-1)];
  const cx=X(p.c.iC)+off[0],cy=Y(p.c.iA)+off[1],col=COLCAT[p.c.cat]||"#9e9e9e";
  s+='<g data-vopen="'+p.idx+'" style="cursor:pointer"><title>'+p.r.shifr+" — "+p.r.name+'\\nI_А='+p.c.iA.toFixed(1)+' ('+p.c.lA+') · I_Ц='+p.c.iC.toFixed(1)+' ('+p.c.lC+')\\nI='+(p.c.I==null?"Н/П":p.c.I.toFixed(1))+' · категория: '+p.c.cat+'\\nРекомендация: '+p.rec.text+(p.st?'\\n⛔ СТОП-ФАКТОР: '+stopsOn(p.r).join("; "):"")+'\\nКлик — открыть НИОКР</title>';
  if(p.st)s+='<circle cx="'+cx.toFixed(1)+'" cy="'+cy.toFixed(1)+'" r="9" fill="none" stroke="#c62828" stroke-width="2.5"/>';
  s+='<circle cx="'+cx.toFixed(1)+'" cy="'+cy.toFixed(1)+'" r="5.5" fill="'+col+'" stroke="#fff" stroke-width="1.6"/></g>';});
 s+='</svg>';
 const cnt={};pts.forEach(p=>{const g=p.st?"Стоп-фактор":recGroup(p.rec.text);cnt[g]=(cnt[g]||0)+1;});
 let leg='<div style="display:flex;gap:12px;flex-wrap:wrap;margin-top:6px;font-size:12px;align-items:center">';
 RECGROUPS.forEach(gr=>{leg+='<span><span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:'+gr[2]+';border:2px solid '+gr[1]+';vertical-align:-2px"></span> '+gr[0]+(cnt[gr[0]]?' <b>('+cnt[gr[0]]+')</b>':'')+'</span>';});
 leg+='<span><span style="display:inline-block;width:12px;height:12px;border-radius:50%;border:2.5px solid #c62828;vertical-align:-2px"></span> стоп-фактор'+(cnt["Стоп-фактор"]?' <b>('+cnt["Стоп-фактор"]+')</b>':'')+'</span>';
 leg+='<span class="mini">точек на матрице: '+pts.length+' из '+vis.length+'</span></div>';
 if(!pts.length)h+='<p class="mini">В выборке нет НИОКР с обеими оценками I_А и I_Ц — заполните экспертные блоки чек-листа (вопросы 1–10 и 20–31).</p>';
 else h+=s+leg;
 if(nod.length)h+='<p class="mini">Без данных А/Ц ('+nod.length+'): '+nod.join(", ")+'</p>';
 return h+'</div>';}
function buildVizCharts(vis){'''
rep("function buildVizCharts(vis){", FN)

# ---------- 4б. niokrAPI: recGroup + expertMatrixCard ----------
rep("vizTimelineSVG,heatHTML,buildVizCharts,renderViz,printViz",
    "vizTimelineSVG,heatHTML,recGroup,expertMatrixCard,buildVizCharts,renderViz,printViz")

# ---------- 5. подключение карточки в buildVizCharts ----------
rep("""h+='<div class="card" style="margin-top:12px"><h3>Интегральный показатель I по НИОКР</h3><p class="mini">Клик по строке открывает НИОКР на дашборде. Пунктир — пороги шкалы (50 / 70 / 85).</p>'+barsSVG(vis)+'</div>';""",
    """h+='<div class="card" style="margin-top:12px"><h3>Интегральный показатель I по НИОКР</h3><p class="mini">Клик по строке открывает НИОКР на дашборде. Пунктир — пороги шкалы (50 / 70 / 85).</p>'+barsSVG(vis)+'</div>';
 h+=expertMatrixCard(vis);""")

io.open(P, "w", encoding="utf-8").write(s)
print("PATCH16 OK, size:", len(s))
