# -*- coding: utf-8 -*-
"""Этап 22 (v22.0): математический аппарат обоснованности решения.
Атомарный патч model_v21_template.html -> model_v22_template.html + сборка HTML.
Базовое правило решения (матрица 3x3 + оверрайд СФ) НЕ изменяется;
новый статистический слой — совещательный. Формат данных v:21 без изменений.
"""
import json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'model_v21_template.html')
DST = os.path.join(HERE, 'model_v22_template.html')
OUT = os.path.join(HERE, '..', 'output')

t = open(SRC, encoding='utf-8').read()
orig_len = len(t)

def rep(old, new, cnt=1):
    global t
    n = t.count(old)
    assert n == cnt, 'COUNT %d != %d for: %r' % (n, cnt, old[:90])
    t = t.replace(old, new)

# ---------- 1. Версионные строки ----------
rep('method:"ЭО-НИОКР v2.0 (портфель ≤30, веса, I_R, калибровка, бланк)"',
    'method:"ЭО-НИОКР v2.1 (портфель ≤30, веса, I_R, калибровка, бланк, обоснованность)"')
rep('<p class="note">Модель v21.1 (этап 21Б: корректирующий выпуск — робастность API) реализует методику ЭО-НИОКР v2.0 — расширение v1.0 (чат «Методика аудита НИОКР», скан Приложения №2) по итогам апробации: портфель ≤30, веса, I_R, калибровка, печатный бланк. Базовый состав 22/5/9 не изменялся. История версий — VERSIONS.md.</p>',
    '<p class="note">Модель v22.0 (этап 22: математический аппарат обоснованности решения) реализует методику ЭО-НИОКР v2.1 — расширение v1.0 (чат «Методика аудита НИОКР», скан Приложения №2): портфель ≤30, веса, I_R, калибровка, печатный бланк; v22.0 добавляет совещательный статистический слой (95% ДИ Стьюдента, W Кендалла, χ², jackknife, ворота G1–G5). Базовый состав 22/5/9 и базовое правило решения не изменялись. Формат данных v21 совместим. История версий — VERSIONS.md.</p>')
assert t.count('ЭО-НИОКР v2.0') == 6, t.count('ЭО-НИОКР v2.0')
t = t.replace('ЭО-НИОКР v2.0', 'ЭО-НИОКР v2.1')
assert t.count('v21.1') == 6, t.count('v21.1')
t = t.replace('v21.1', 'v22.0')
rep('version:"21.1"', 'version:"22.0"')
rep('11_Методика_ЭО_НИОКР_22_вопроса_v21', '11_Методика_ЭО_НИОКР_22_вопроса_v22')
rep('(веса равные)', '(веса равные)')

# ---------- 2. Математический модуль ----------
MATH = '''/* ---------- математический аппарат обоснованности (v22.0) ---------- */
var TCRIT95=[12.706,4.303,3.182,2.776,2.571,2.447,2.365,2.306,2.262]; /* t(0,975; df), df=1..9 */
var CHI2CRIT={9:16.919,11:19.675,21:32.671}; /* χ²(0,95; df), α=0,05 */
function median(a){if(!a.length)return null;var s=a.slice().sort(function(x,y){return x-y;});var m=Math.floor(s.length/2);return s.length%2?s[m]:(s[m-1]+s[m])/2;}
function stdevS(a){if(a.length<2)return null;var m=mean(a);return Math.sqrt(a.reduce(function(s,x){return s+(x-m)*(x-m);},0)/(a.length-1));}
function ciStat(vals,weights){
  var n=vals.length;if(n<2)return {m:n?mean(vals):null,n:n,se:null,df:null,t:null,lo:null,hi:null,neff:n};
  var i,m,sw=0,sw2=0;
  if(weights){
    for(i=0;i<n;i++){sw+=weights[i];sw2+=weights[i]*weights[i];}
    if(!(sw>0)||!(sw2>0))return {m:null,n:n,se:null,df:null,t:null,lo:null,hi:null,neff:null};
    m=0;for(i=0;i<n;i++)m+=weights[i]*vals[i];m/=sw;
    var v=0;for(i=0;i<n;i++)v+=weights[i]*weights[i]*(vals[i]-m)*(vals[i]-m);
    var se=Math.sqrt(v)/sw,neff=sw*sw/sw2,df=Math.max(1,Math.min(9,Math.floor(neff)-1)),tc=TCRIT95[df-1];
    return {m:m,n:n,se:se,df:df,t:tc,lo:m-tc*se,hi:m+tc*se,neff:neff};}
  m=mean(vals);var s=stdevS(vals),se2=s/Math.sqrt(n),df2=Math.min(9,n-1),t2=TCRIT95[df2-1];
  return {m:m,n:n,se:se2,df:n-1,t:t2,lo:m-t2*se2,hi:m+t2*se2,neff:n};
}
function wScale(W){return W==null?"—":W<0.3?"слабая":W<0.5?"умеренная":W<0.7?"хорошая":"высокая";}
function kendallW(rows){
  var m=rows.length;if(m<2)return null;var k=rows[0].length;if(k<2)return null;
  var e,j;
  for(e=0;e<m;e++){if(rows[e].length!==k)return null;for(j=0;j<k;j++){if(typeof rows[e][j]!=="number")return null;}}
  var Rj=[],Tsum=0;for(j=0;j<k;j++)Rj.push(0);
  for(e=0;e<m;e++){
    var r=rows[e],ranks=[];
    for(j=0;j<k;j++){var less=0,eq=0;for(var u=0;u<k;u++){if(r[u]<r[j])less++;else if(r[u]===r[j])eq++;}ranks.push(less+(eq+1)/2);}
    for(j=0;j<k;j++)Rj[j]+=ranks[j];
    var cnt={};for(j=0;j<k;j++){var v=r[j];cnt[v]=(cnt[v]||0)+1;}
    for(var key in cnt){var c=cnt[key];Tsum+=c*c*c-c;}
  }
  var Rbar=m*(k+1)/2,S=0;for(j=0;j<k;j++)S+=(Rj[j]-Rbar)*(Rj[j]-Rbar);
  var den=m*m*(k*k*k-k)-m*Tsum,df=k-1,crit=CHI2CRIT[df]||null;
  if(!(den>0))return {W:null,chi2:null,df:df,crit:crit,sig:false,m:m,k:k,scale:"—",note:"все оценки совпадают (связки)"};
  var W=12*S/den,chi2=m*(k-1)*W;
  return {W:W,chi2:chi2,df:df,crit:crit,sig:crit!=null?chi2>=crit:null,m:m,k:k,scale:wScale(W)};
}
'''
rep('/* ---------- расчётное ядро ---------- */', MATH + '/* ---------- расчётное ядро ---------- */')

# ---------- 3. justOf + рендер протокола ----------
JUST = '''/* ---------- обоснованность решения (v22.0, совещательный слой) ---------- */
function justOf(P0){
  var proj=P0||S,R=calc(proj);
  var J={R:R,n:R.nExp,perQ:[],ciA:null,ciC:null,ciR:null,W:{all:null,A:null,C:null},jack:null,gates:[],verdict:"",short:"н/д",okCount:0,noTotal:0,complete:true};
  var act=[];proj.exp.forEach(function(e,i){if(CODES.filter(function(c){return e.q[c]!=null;}).length>0)act.push(i);});
  act.forEach(function(i){var e=proj.exp[i];
    CODES.forEach(function(c){if(e.q[c]==="НО")J.noTotal++;});
    if(CODES.filter(function(c){return typeof e.q[c]==="number";}).length<22)J.complete=false;});
  if(!act.length)J.complete=false;
  R.q.forEach(function(r){J.perQ.push({code:r.code,n:r.vals.length,mean:r.mean,median:median(r.vals),s:stdevS(r.vals),sd:r.sd,ci:ciStat(r.vals)});});
  function idxCI(key){var vals=[],ws=[];R.exp.forEach(function(x){if(x[key]!=null){vals.push(x[key]);ws.push(DB.wmode?(x.w||1):1);}});
    if(vals.length<2)return null;var c=ciStat(vals,DB.wmode?ws:null);
    c.zLo=zone(c.lo);c.zHi=zone(c.hi);c.stable=(c.zLo!=null&&c.zLo===c.zHi);c.zone=zone(c.m);return c;}
  J.ciA=idxCI("IA");J.ciC=idxCI("IC");J.ciR=DB.rmode?idxCI("IR"):null;
  function wBlock(codes){var rows=[];act.forEach(function(i){var e=proj.exp[i];var vs=codes.map(function(c){return e.q[c];});
      if(vs.every(function(v){return typeof v==="number";}))rows.push(vs);});
    var res=kendallW(rows);
    if(!res)res={W:null,chi2:null,df:codes.length-1,crit:CHI2CRIT[codes.length-1]||null,sig:false,m:0,k:codes.length,scale:"—",note:"менее 2 полных наборов оценок"};
    res.mUsed=rows.length;return res;}
  J.W.all=wBlock(CODES);
  J.W.A=wBlock(QEO22.q.filter(function(q){return q.dir=="A";}).map(function(q){return q.code;}));
  J.W.C=wBlock(QEO22.q.filter(function(q){return q.dir=="C";}).map(function(q){return q.code;}));
  var runs=[];
  if(act.length>=2){act.forEach(function(i){var clone={exp:proj.exp.filter(function(_,k){return k!==i;})};var Rk=calc(clone);
    runs.push({e:i+1,IA:Rk.IA,IC:Rk.IC,zA:Rk.zA,zC:Rk.zC,rec:Rk.rec,changed:Rk.rec!==R.rec,
      dIA:(Rk.IA!=null&&R.IA!=null)?Math.abs(Rk.IA-R.IA):null,dIC:(Rk.IC!=null&&R.IC!=null)?Math.abs(Rk.IC-R.IC):null});});}
  var changes=runs.filter(function(r){return r.changed;}).length;
  J.jack={n:runs.length,runs:runs,changes:changes,stable:runs.length?changes===0:null,
    maxDIA:runs.length?Math.max.apply(null,runs.map(function(r){return r.dIA||0;})):null,
    maxDIC:runs.length?Math.max.apply(null,runs.map(function(r){return r.dIC||0;})):null};
  var G=[];
  G.push({id:"G1",name:"Полнота оценок",ok:J.noTotal===0&&J.complete,
    note:J.noTotal===0&&J.complete?"все 22 вопроса оценены всеми действующими экспертами, Н/О отсутствуют":"Н/О: "+J.noTotal+"; неполные наборы: "+(J.complete?"нет":"есть")});
  var wOk=J.W.all.W!=null&&J.W.all.W>=0.5&&J.W.all.sig===true;
  G.push({id:"G2",name:"Согласованность экспертов (W≥0,5; χ² значим)",ok:wOk,
    note:J.W.all.W==null?"W не рассчитан (менее 2 полных наборов)":"W="+f2(J.W.all.W)+" ("+J.W.all.scale+"), χ²="+f2(J.W.all.chi2)+" при крит. "+f2(J.W.all.crit)});
  var zOk=!!(J.ciA&&J.ciA.stable&&J.ciC&&J.ciC.stable);
  G.push({id:"G3",name:"Устойчивость зон I_A и I_C (95% ДИ)",ok:zOk,
    note:zOk?"доверительные интервалы обоих индексов не пересекают границ зон":"ДИ пересекает границу зоны либо не рассчитан — зона неустойчива"});
  var jOk=J.jack.stable===true&&J.jack.n>=2;
  G.push({id:"G4",name:"Устойчивость рекомендации (jackknife)",ok:jOk,
    note:J.jack.n<2?"недостаточно экспертов для проверки":"изменений рекомендации при исключении одного эксперта: "+J.jack.changes+" из "+J.jack.n});
  var dOk=R.flags.length===0&&R.kvGroup.length===0;
  G.push({id:"G5",name:"Отсутствие неразрешённых разногласий",ok:dOk,
    note:dOk?"σ вопросов в пределах порога; групповых «красных» сигналов нет":"разногласия: "+(R.flags.join(", ")||"нет")+"; КВ: "+(R.kvGroup.join(", ")||"нет")});
  J.gates=G;J.okCount=G.filter(function(g){return g.ok;}).length;
  if(J.n<3)J.verdict="Недостаточно экспертов (менее 3) — статистическая проверка ограничена";
  else if(J.okCount===5)J.verdict="Решение статистически обосновано";
  else if(J.okCount>=3)J.verdict="Обосновано с оговорками";
  else J.verdict="Требуется доработка экспертной оценки";
  J.short=J.n<3?"н/д":J.okCount===5?"обосновано":J.okCount>=3?"с оговорками":"доработка";
  return J;
}
function jbadge(J){
  if(!J)return "—";
  var c=J.short=="обосновано"?"#2e7d32":J.short=="с оговорками"?"#ef6c00":J.short=="доработка"?"#c62828":"#9e9e9e";
  return '<span class="pill" style="background:'+c+'" title="'+esc(J.verdict)+'">'+J.short+' '+J.okCount+'/5</span>';
}
function justTablesHTML(J){
  var R=J.R,h="";
  h+='<table><tr><th>Ворота</th><th>Статус</th><th>Пояснение</th></tr>';
  J.gates.forEach(function(g){h+='<tr><td><b>'+g.id+'</b> '+esc(g.name)+'</td><td>'+(g.ok?'<span class="pill" style="background:#2e7d32">ВЫПОЛНЕНО</span>':'<span class="pill" style="background:#c62828">НЕ ВЫПОЛНЕНО</span>')+'</td><td>'+esc(g.note)+'</td></tr>';});
  h+='</table>';
  h+='<h4 style="margin:10px 0 4px">Согласованность экспертов — коэффициент конкордации Кендалла W (с поправкой на связки)</h4><table><tr><th>Блок</th><th class="num">Полных наборов</th><th class="num">W</th><th>Интерпретация</th><th class="num">χ²</th><th class="num">χ² крит. (α=0,05)</th><th>Значимость</th></tr>';
  [["Все 22 вопроса",J.W.all],["Актуальность (10)",J.W.A],["Целесообразность (12)",J.W.C]].forEach(function(p){var w=p[1];
    h+='<tr><td>'+p[0]+'</td><td class="num">'+w.mUsed+'</td><td class="num"><b>'+(w.W==null?"н/д":f2(w.W))+'</b></td><td>'+w.scale+'</td><td class="num">'+(w.chi2==null?"—":f2(w.chi2))+'</td><td class="num">'+(w.crit==null?"—":f2(w.crit))+'</td><td>'+(w.sig===true?"значима":w.sig===false?"не значима":"—")+'</td></tr>';});
  h+='</table><p class="note">Шкала интерпретации W: &lt;0,3 — слабая; 0,3–0,5 — умеренная; 0,5–0,7 — хорошая; ≥0,7 — высокая (Schmidt, 1997). χ²=m·(k−1)·W, df=k−1; значимость при χ²≥критического. W рассчитывается по экспертам с полным числовым набором оценок блока (complete cases); ранги — средние при связках.</p>';
  h+='<h4 style="margin:10px 0 4px">Доверительные интервалы групповых индексов (95%, критерий Стьюдента; эксперт — единица наблюдения)</h4><table><tr><th>Индекс</th><th class="num">Значение</th><th class="num">n</th><th class="num">SE</th><th class="num">t (df)</th><th class="num">ДИ нижн.</th><th class="num">ДИ верх.</th><th>Зона</th><th>Устойчивость зоны</th></tr>';
  var rowsCI=[["I_A",J.ciA,R.IA,R.zA],["I_C",J.ciC,R.IC,R.zC]];
  if(DB.rmode)rowsCI.push(["I_R",J.ciR,R.IR,R.zR]);
  rowsCI.forEach(function(p){var c=p[1];
    if(!c){h+='<tr><td>'+p[0]+'</td><td class="num">'+f2(p[2])+'</td><td colspan="7">недостаточно данных (менее 2 действующих индивидуальных индексов)</td></tr>';return;}
    h+='<tr><td>'+p[0]+'</td><td class="num"><b>'+f2(c.m)+'</b></td><td class="num">'+c.n+(DB.wmode?' (n_eff '+f2(c.neff)+')':'')+'</td><td class="num">'+f2(c.se)+'</td><td class="num">'+f2(c.t)+' (df '+c.df+')</td><td class="num">'+f2(c.lo)+'</td><td class="num">'+f2(c.hi)+'</td><td>'+zbadge(p[3])+'</td><td>'+(c.stable?'<span class="pill" style="background:#2e7d32">УСТОЙЧИВА</span>':'<span class="pill" style="background:#ef6c00">НЕУСТОЙЧИВА</span>')+'</td></tr>';});
  h+='</table><p class="note">ДИ = среднее ± t(0,975; df)·s/√n, df=n−1 (выборочное s). В режиме весов: SE=√(Σw²·(I−m_w)²)/Σw, эффективное число экспертов n_eff=(Σw)²/Σw² (Kish), df=⌊n_eff⌋−1. Зона устойчива, если ДИ целиком лежит в одной зоне; пересечение границы — повод для обсуждения или привлечения дополнительных экспертов.</p>';
  if(J.jack.n){
    h+='<h4 style="margin:10px 0 4px">Устойчивость рекомендации — jackknife (поочерёдное исключение одного эксперта)</h4><table><tr><th>Исключён</th><th class="num">I_A</th><th class="num">I_C</th><th>Зоны (А/Ц)</th><th>Рекомендация</th><th>Изменение</th></tr>';
    J.jack.runs.forEach(function(r){h+='<tr><td>Э'+r.e+'</td><td class="num">'+f2(r.IA)+'</td><td class="num">'+f2(r.IC)+'</td><td>'+(r.zA||"—")+'/'+(r.zC||"—")+'</td><td style="font-size:11.5px">'+esc(r.rec||"не формируется")+'</td><td>'+(r.changed?'<span class="pill" style="background:#c62828">ДА</span>':"нет")+'</td></tr>';});
    h+='</table><p class="note">Изменений рекомендации: '+J.jack.changes+' из '+J.jack.n+'. Максимальное влияние эксперта: |ΔI_A|='+f2(J.jack.maxDIA)+', |ΔI_C|='+f2(J.jack.maxDIC)+'. Пересчёт выполняется по всем правилам базового ядра (включая оверрайд стоп-фактора по оставшимся экспертам).</p>';}
  h+='<h4 style="margin:10px 0 4px">Повопросная статистика: медиана, выборочное s, 95% ДИ среднего</h4><table><tr><th>Код</th><th class="num">n</th><th class="num">Среднее</th><th class="num">Медиана</th><th class="num">s (выб.)</th><th class="num">σ (попул.)</th><th class="num">ДИ нижн.</th><th class="num">ДИ верх.</th></tr>';
  J.perQ.forEach(function(r){h+='<tr><td class="num"><b>'+r.code+'</b></td><td class="num">'+r.n+'</td><td class="num">'+f2(r.mean)+'</td><td class="num">'+(r.median==null?"—":r.median)+'</td><td class="num">'+f2(r.s)+'</td><td class="num">'+f2(r.sd)+'</td><td class="num">'+(r.ci.lo==null?"—":f2(r.ci.lo))+'</td><td class="num">'+(r.ci.hi==null?"—":f2(r.ci.hi))+'</td></tr>';});
  h+='</table>';
  return h;
}
function justProtocolHTML(J){
  var R=J.R;
  return '<div class="card pform"><p class="noprint"><button class="btn sm gray" id="bBackJust">← К сводной оценке</button> <button class="btn sm" id="bPrintJust">🖨 Печать протокола</button></p>'+
  '<h2>ПРИЛОЖЕНИЕ №13 — ПРОТОКОЛ СТАТИСТИЧЕСКОЙ ОБОСНОВАННОСТИ РЕШЕНИЯ<br>(ЭО-НИОКР v2.1 · модель v22.0)</h2>'+
  '<p><b>НИОКР:</b> '+esc(S.obj.name||"__")+' (шифр '+esc(S.obj.shifr||"__")+'). <b>Дата оценки:</b> '+esc(S.obj.date||"__")+'. <b>Экспертов (действующих):</b> '+J.n+'.</p>'+
  '<p><b>Базовая рекомендация:</b> '+esc(R.rec||"не формируется")+'.<br><b>Статистический вердикт (совещательный):</b> '+esc(J.verdict)+' — ворот выполнено: '+J.okCount+' из 5.</p>'+
  justTablesHTML(J)+
  '<p class="note"><b>Ограничения.</b> Шкала 1–5 порядковая: ДИ по t-критерию — приближение для малых выборок (n≤10); W чувствителен к составу группы и числу вопросов; jackknife оценивает влияние отдельных экспертов, но не заменяет обсуждение. Статистический слой не изменяет базовую рекомендацию; окончательное решение принимает комиссия. Пороги зон: В≥'+f2(CAL.zoneV)+', С≥'+f2(CAL.zoneS)+'; σ '+f2(CAL.s1)+'/'+f2(CAL.s2)+'/'+f2(CAL.s3)+'.</p>'+
  '<div class="sig">Председатель комиссии: ____________________ &nbsp; Члены комиссии: ____________________<br>Ответственный секретарь: ____________________</div></div>';
}
'''
rep('/* ---------- рендер: дашборд ---------- */', JUST + '/* ---------- рендер: дашборд ---------- */')

# ---------- 4. JFORM + renderSvod ----------
rep('var PFORM=false;', 'var PFORM=false,JFORM=false;')
rep('function renderSvod(){\n  var R=calc(),h="";',
    'function renderSvod(){\n  var R=calc(),h="";\n  if(JFORM){$("#tab-svod").innerHTML=justProtocolHTML(justOf());\n    $("#bBackJust").addEventListener("click",function(){JFORM=false;renderSvod();});\n    $("#bPrintJust").addEventListener("click",function(){printTab("svod");});\n    return;}')
rep("  h+='<div class=\"card\" style=\"margin-top:14px\"><h3>Повестка обсуждения комиссии</h3>'+agendaHTML(R)+'</div>';",
    """  h+='<div class="card" style="margin-top:14px"><h3>Повестка обсуждения комиссии</h3>'+agendaHTML(R)+'</div>';
  var J=justOf();
  h+='<div class="card" style="margin-top:14px"><h3>Обоснованность решения (v22.0) '+jbadge(J)+'</h3>'+
     '<p class="note">Совещательный статистический слой: базовая рекомендация (матрица 3×3 + оверрайд СФ) не изменяется, решение принимает комиссия. ДИ — 95% (Стьюдент, df=n−1); W — конкордация Кендалла с поправкой на связки; χ² — значимость W; jackknife — исключение одного эксперта. Ворота G1–G5 — см. методику ЭО-НИОКР v2.1, раздел 21.</p>'+
     justTablesHTML(J)+
     '<p class="noprint" style="margin-top:10px"><button class="btn" id="bJustProt">🖨 Протокол обоснованности (Прил. №13)</button></p></div>';""")
rep('  $("#tab-svod").innerHTML=h;\n  $("#bCsv").addEventListener',
    '  $("#tab-svod").innerHTML=h;\n  if($("#bJustProt"))$("#bJustProt").addEventListener("click",function(){JFORM=true;renderSvod();printTab("svod");});\n  $("#bCsv").addEventListener')
rep('CUR=0;PFORM=false;', 'CUR=0;PFORM=false;JFORM=false;')

# ---------- 5. Портфель: реестр, CSV, MD ----------
rep('function portRows(){return DB.projects.map(function(p,i){return {i:i,p:p,R:calc(p)};});}',
    'function portRows(){return DB.projects.map(function(p,i){return {i:i,p:p,R:calc(p),J:justOf(p)};});}')
rep('<th>Сигналы</th><th class="noprint">Действия</th>',
    '<th>Сигналы</th><th>Обосн.</th><th class="noprint">Действия</th>')
rep("      '<td>'+(sig||\"—\")+'</td>'+",
    "      '<td>'+(sig||\"—\")+'</td>'+\n      '<td>'+jbadge(r.J)+'</td>'+")
rep('"КВ групповые","Решение комиссии"];',
    '"КВ групповые","Решение комиссии","Обоснованность"];')
rep('r.R.kvGroup.join(" ")||"нет",r.p.com.rec||""]);});',
    'r.R.kvGroup.join(" ")||"нет",r.p.com.rec||"",r.J?(r.J.short+" ("+r.J.okCount+"/5)"):"н/д"]);});')
rep('" Зоны | Рекомендация | СФ | КВ |","| --- | --- | --- | --- | --- | --- |"+(DB.rmode?" --- |":"")+" --- | --- | --- | --- |"];',
    '" Зоны | Рекомендация | СФ | КВ | Обосн. |","| --- | --- | --- | --- | --- | --- |"+(DB.rmode?" --- |":"")+" --- | --- | --- | --- | --- |"];')
rep('(r.R.kvGroup.join(", ")||"нет")+" |");});',
    '(r.R.kvGroup.join(", ")||"нет")+" | "+(r.J?r.J.short:"н/д")+" |");});')
rep('  projects:function(){return DB.projects.map(function(p,i){return {i:i,obj:p.obj,R:calc(p)};});},',
    '  projects:function(){return DB.projects.map(function(p,i){return {i:i,obj:p.obj,R:calc(p),J:justOf(p)};});},')

# ---------- 6. Справка и методика (вкладка) ----------
rep('  <p><b>Горячие клавиши:</b>',
    '  <p><b>Обоснованность решения (v22.0).</b> Вкладка «Свод» → карточка «Обоснованность решения»: 95% доверительные интервалы индексов (Стьюдент), устойчивость зон, коэффициент конкордации Кендалла W с χ²-значимостью, jackknife-устойчивость рекомендации, повопросные медианы и ДИ; ворота G1–G5 и совещательный вердикт. Кнопка «🖨 Протокол обоснованности (Прил. №13)» — печатная форма для приложения к протоколу комиссии. Базовая рекомендация не изменяется; в реестре портфеля — колонка «Обосн.».</p>\n  <p><b>Горячие клавиши:</b>')
rep("     '<table class=\"mtx\" style=\"margin-top:6px\">'+mtxStatic()+'</table></div></div>';",
    """     '<table class="mtx" style="margin-top:6px">'+mtxStatic()+'</table></div></div>';
  h+='<div class="card mblock" style="margin-top:14px"><h3>2б. Математический аппарат обоснованности решения (v22.0, ЭО-НИОКР v2.1 §21)</h3>'+
     '<p><b>Назначение:</b> повысить обоснованность решения комиссии после экспертной оценки, не изменяя базовое правило (матрица 3×3 + оверрайд СФ). Слой совещательный.</p>'+
     '<p><b>Доверительный интервал (95%):</b> m ± t(0,975; n−1)·s/√n, где s — выборочное стандартное отклонение индивидуальных индексов (эксперт — единица наблюдения). В режиме весов: SE=√(Σw²(I−m_w)²)/Σw, n_eff=(Σw)²/Σw² (Kish), df=⌊n_eff⌋−1. Зона устойчива, если ДИ целиком в одной зоне.</p>'+
     '<p><b>Конкордация Кендалла:</b> W=12S/(m²(k³−k)−m·ΣT), S=Σ(R_j−R̄)², R̄=m(k+1)/2, T=Σ(t³−t) по группам связок каждого эксперта; ранги средние при связках. χ²=m(k−1)W, df=k−1; значимость при χ²≥χ²крит(0,05): df 9→16,919; 11→19,675; 21→32,671. Шкала W: &lt;0,3 слабая; 0,3–0,5 умеренная; 0,5–0,7 хорошая; ≥0,7 высокая. Complete cases: эксперты с полным числовым набором блока.</p>'+
     '<p><b>Jackknife:</b> поочерёдное исключение каждого эксперта с полным пересчётом ядра (зоны, матрица, оверрайд СФ по оставшимся); фиксируются изменения рекомендации и максимальное влияние |ΔI|.</p>'+
     '<p><b>Ворота:</b> G1 полнота (нет Н/О, все 22 оценены); G2 W≥0,5 и χ² значим; G3 ДИ обоих индексов не пересекают границ зон; G4 рекомендация не меняется ни при одном исключении; G5 нет вопросов с σ&gt;порога и групповых КВ. Вердикт: 5/5 — «Решение статистически обосновано»; 3–4/5 — «Обосновано с оговорками»; ≤2/5 — «Требуется доработка экспертной оценки»; при n&lt;3 — «проверка ограничена». Решение принимает комиссия.</p>'+
     '<p class="note">Ограничения: порядковая шкала 1–5 (t-ДИ — приближение), малые выборки n≤10, чувствительность W к составу группы. Печатная форма — Протокол обоснованности (Прил. №13), кнопка в карточке «Обоснованность решения».</p></div>';""")

# ---------- 7. API ----------
rep('  goTab:goTab\n};',
    '''  goTab:goTab,
  just:function(i){return justOf(i==null?S:DB.projects[i]);},
  kendallW:kendallW,ciStat:ciStat,median:median,stdevS:stdevS,TCRIT95:TCRIT95,CHI2CRIT:CHI2CRIT,
  printJust:function(){JFORM=true;goTab("svod");renderSvod();return true;}
};''')

# ---------- запись шаблона и сборка ----------
open(DST, 'w', encoding='utf-8').write(t)
qeo = json.load(open(os.path.join(HERE, 'data', 'qeo_v21.json'), encoding='utf-8'))
assert t.count('/*__QEO21__*/null') == 1
html = t.replace('/*__QEO21__*/null', json.dumps(qeo, ensure_ascii=False))
vdir = os.path.join(OUT, 'versions')
os.makedirs(vdir, exist_ok=True)
p1 = os.path.join(vdir, 'Аудит_НИОКР_модель_v22.0.html')
p2 = os.path.join(OUT, 'Аудит_НИОКР_интерактивная_модель.html')
open(p1, 'w', encoding='utf-8').write(html)
open(p2, 'w', encoding='utf-8').write(html)
print('template: %d -> %d chars' % (orig_len, len(t)))
print('html: %d chars / %d bytes' % (len(html), len(html.encode('utf-8'))))
print('OK: v22.0 built')