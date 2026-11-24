#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Этап 24: единая форма организации по скану Приложения №2.
Сохраняет конверт данных v21; расширяет домен СФ до 1..5 с миграцией старых строк.
Печатная форма — ровно две страницы, без дополнительных полей и блока Р."""
import io, os, json, re
H=os.path.dirname(os.path.abspath(__file__)); R=os.path.join(H,'..')
org=json.load(io.open(os.path.join(H,'data','org_form_v23.json'),encoding='utf-8'))
org['interpretation']['note']=org['interpretation']['note'].replace('absence риска','отсутствие риска')
json.dump(org,io.open(os.path.join(H,'data','org_form_v23.json'),'w',encoding='utf-8'),ensure_ascii=False,indent=2)
q=json.load(io.open(os.path.join(H,'data','qeo_v21.json'),encoding='utf-8'))
for x in q['base']['q']:
    o=next(y for y in org['main'] if y['code']==x['code']); x['crit']=o['criterion']; x['q']=o['question']
q['base']['sf']=org['sf'][:]
for i,k in enumerate(q['base']['kv']):
    k['q']=org['kv'][i]; k['legacyDir']=k['dir']; k['dir']='risk_yes'
q['version']='23.0'; q['method']='ЭО-НИОКР v2.2'; q['orgForm']=org
json.dump(q,io.open(os.path.join(H,'data','qeo_v23.json'),'w',encoding='utf-8'),ensure_ascii=False,indent=2)
t=io.open(os.path.join(H,'model_v22_template.html'),encoding='utf-8').read()
def rep(a,b,n=1):
    global t
    c=t.count(a); assert c==n,(c,n,a[:90]); t=t.replace(a,b)
def block(start,end,new):
    global t
    i=t.index(start); j=t.index(end,i); t=t[:i]+new+t[j:]
rep('var QEO21 = /*__QEO21__*/null;','var QEO21 = /*__QEO21__*/null;\nvar ORG = QEO21.orgForm;')
rep('version:"22.0"','version:"23.0"')
rep('ЭО-НИОКР v2.1','ЭО-НИОКР v2.2',t.count('ЭО-НИОКР v2.1'))
rep('v22.0','v23.0',t.count('v22.0'))
rep('11_Методика_ЭО_НИОКР_22_вопроса_v22','11_Методика_ЭО_НИОКР_22_вопроса_v23')
rep('Прил. №1)</button>','Прил. №2)</button>')
rep('Бланк анкеты (Прил. №1)','Бланк анкеты (Прил. №2)',t.count('Бланк анкеты (Прил. №1)'))
rep('Документальные подтверждения не требуются.','Форма организации — Приложение №2. Исходные формулировки сохранены; пояснения к шкале и правила интерпретации приведены в методике ЭО-НИОКР v2.2.')
# Совместимость хранения: конверт v21, числовой домен СФ.
rep('function kvSig(k,v){if(typeof v!="number")return false;return k.dir=="neg"?v<=2:v>=4;}', '''function normSF(v){
  if(v==null||v==="")return null;
  if(typeof v==="number")return v>=1&&v<=5&&v===Math.floor(v)?v:null;
  return {"Да":1,"Нет":5,"Неопределённо":3}[v]||null;
}
function normAnswers(e){if(!e.sf)e.sf={};if(!e.kv)e.kv={};for(var i=1;i<=5;i++)e.sf[i]=normSF(e.sf[i]);for(var j=1;j<=9;j++){var v=e.kv[j];e.kv[j]=(typeof v==="number"&&v>=1&&v<=5&&v===Math.floor(v))?v:null;}return e;}
function kvSig(k,v){return typeof v==="number"&&v<=2;} /* 1=да, 2=возможно да — для всех 9 вопросов */
function kvClarify(v){return v===3;}''')
rep('var ans=proj.exp.map(function(e){return e.sf[i+1]||null;}).filter(Boolean);','var ans=proj.exp.map(function(e){return normSF(e.sf[i+1]);}).filter(Boolean);')
rep('var yes=ans.filter(function(a){return a=="Да";}).length;\n    var unc=ans.filter(function(a){return a=="Неопределённо";}).length;\n    var no=ans.filter(function(a){return a=="Нет";}).length;',
    'var yes=ans.filter(function(a){return a<=2;}).length;\n    var unc=ans.filter(function(a){return a==3;}).length;\n    var no=ans.filter(function(a){return a>=4;}).length;')
rep('R.sf.push({n:i+1,yes:yes,unc:unc,no:no,st:st});','R.sf.push({n:i+1,yes:yes,unc:unc,no:no,st:st,counts:[1,2,3,4,5].map(function(v){return ans.filter(function(a){return a==v;}).length;})});')
rep('var sig=0,single=false;\n    proj.exp.forEach(function(e){if(kvSig(k,e.kv[i+1]))sig++;});\n    var g=sig>=2;single=sig==1;\n    R.kv.push({n:k.n,sig:sig,group:g,single:single});',
    'var sig=0,clar=0,single=false;\n    proj.exp.forEach(function(e){if(kvSig(k,e.kv[i+1]))sig++;if(kvClarify(e.kv[i+1]))clar++;});\n    var g=sig>=2;single=sig==1;\n    R.kv.push({n:k.n,sig:sig,clarify:clar,group:g,single:single});\n    if(clar)R.agenda.push("КВ"+k.n+": "+clar+" ответ(ов) «неопределённо» — требуется уточнение комиссии");')
rep('p.exp.forEach(function(e){if(e.w==null)e.w=1;if(!e.qr)e.qr={};if(!e.qrc)e.qrc={};});','p.exp.forEach(function(e){if(e.w==null)e.w=1;if(!e.qr)e.qr={};if(!e.qrc)e.qrc={};normAnswers(e);});',t.count('p.exp.forEach(function(e){if(e.w==null)e.w=1;if(!e.qr)e.qr={};if(!e.qrc)e.qrc={};});'))
rep('e.sf[n]=(e.sf[n]==v?null:v);','var nv=normSF(v);e.sf[n]=(e.sf[n]==nv?null:nv);')
rep('e.sf[n]=v;save();renderAll();return true;','e.sf[n]=normSF(v);save();renderAll();return true;')
rep('S.exp[sl-1]=d.expert;save();renderAll();return true;','S.exp[sl-1]=normAnswers(d.expert);save();renderAll();return true;')
rep('S.exp[slot]=d.expert;save();renderAll();toast','S.exp[slot]=normAnswers(d.expert);save();renderAll();toast')
rep('for(var sfi=1;sfi<=5;sfi++)ex.sf[sfi]="Нет";','for(var sfi=1;sfi<=5;sfi++)ex.sf[sfi]=5;')
rep('for(var kvi=1;kvi<=9;kvi++)ex.kv[kvi]=(kvi==3&&e2<2)?1:3;','for(var kvi=1;kvi<=9;kvi++)ex.kv[kvi]=(kvi==3&&e2<2)?1:5;')
rep('var dir=QEO22.kv[kvi-1].dir;var base=dir=="neg"?5:1;','var base=5;')
rep('(sp.kv&&sp.kv.indexOf(kvi-1)>=0&&e<2)?(dir=="neg"?1:5):base;','(sp.kv&&sp.kv.indexOf(kvi-1)>=0&&e<2)?1:base;')
rep('sf:["Нет","Нет","Нет","Нет","Нет"],kv:[5,5,5,5,2,5,1,1,5]','sf:[5,5,5,5,5],kv:[5,5,5,5,2,5,5,5,5]')
rep('sf:["Да","Нет","Нет","Нет","Нет"]','sf:[1,5,5,5,5]')
rep('sf:["Нет","Нет","Нет","Нет","Нет"],kv:[5,5,5,5,5,5,4,4,5]','sf:[5,5,5,5,5],kv:[5,5,5,5,5,5,1,1,5]')
rep('ex.sf[sfi]=(sp.sf&&sp.sf.indexOf(sfi-1)>=0&&e==0)?"Да":"Нет";','ex.sf[sfi]=(sp.sf&&sp.sf.indexOf(sfi-1)>=0&&e==0)?1:5;')
# Электронные блоки СФ/КВ: исходная шкала, без смены текстов.
SFNEW='''function sfBlockHTML(e){
  var lab=ORG.sf_labels;
  var h='<div class="card" style="margin-top:14px"><h3>СТОП-ФАКТОРЫ</h3><table><tr><th style="width:5%">№</th><th>Стоп-фактор</th><th style="width:30%">Ответ 1–5</th><th style="width:18%">Комментарий</th></tr>';
  QEO22.sf.forEach(function(t,i){var v=normSF(e.sf[i+1]);
    h+='<tr><td class="num">'+(i+1)+'</td><td>'+esc(t)+'</td><td><span class="sb">'+[1,2,3,4,5].map(function(x){return '<button class="k'+x+(v==x?" on":"")+'" data-sf="'+(i+1)+'" data-v="'+x+'" title="'+lab[x]+'">'+x+'</button>';}).join('')+'</span></td><td><input type="text" style="width:100%" data-sfc="'+(i+1)+'" value="'+esc(e.sfc[i+1]||"")+'"></td></tr>';});
  return h+'</table><p class="note">1 — да; 2 — возможно да; 3 — неопределённо; 4 — возможно нет; 5 — нет. СФ выявлен при 1–2 хотя бы у одного эксперта; 3 — уточнить комиссией; 4–5 — не выявлен. Пустой ответ не означает «нет». В индексы не входит.</p></div>';}
'''
block('function sfBlockHTML(e){','function kvBlockHTML(e){',SFNEW)
KVNEW='''function kvBlockHTML(e){
  var lab=ORG.response_labels;
  var h='<div class="card" style="margin-top:14px"><h3>«КРАСНЫЕ» ВОПРОСЫ</h3><table><tr><th style="width:5%">№</th><th>Вопрос</th><th style="width:30%">Ответ 1–5</th><th style="width:18%">Комментарий</th></tr>';
  QEO22.kv.forEach(function(k,i){var v=e.kv[i+1];
    h+='<tr><td class="num">'+k.n+'</td><td>'+esc(k.q)+'</td><td><span class="sb">'+[1,2,3,4,5].map(function(x){return '<button class="k'+x+(v==x?" on":"")+'" data-kv="'+(i+1)+'" data-v="'+x+'" title="'+lab[x]+'">'+x+'</button>';}).join('')+'</span></td><td><input type="text" style="width:100%" data-kvc="'+(i+1)+'" value="'+esc(e.kvc[i+1]||"")+'"></td></tr>';});
  return h+'</table><p class="note">Единая шкала скана: 1 — да … 5 — нет. Сигнал риска при 1–2 по всем 9 вопросам, включая КВ7–8 (нет перспективы); 3 — уточнить; 4–5 — нет сигнала. Групповой сигнал — ≥2 экспертов. В индексы не входит.</p></div>';}
'''
block('function kvBlockHTML(e){','/* ---------- рендер: свод ---------- */',KVNEW)
FORMNEW='''function ankFormHTML(e){
  function marks(v){return [1,2,3,4,5].map(function(k){return '<td class="circ'+(v==k?' mark':'')+'">'+(v==k?'✓':'')+'</td>';}).join('');}
  function labels(low,high){return '<tr><td colspan="3"></td><td class="ol">'+low+'</td><td></td><td></td><td></td><td class="ol">'+high+'</td></tr><tr class="onum"><td colspan="3"></td>'+[1,2,3,4,5].map(function(k){return '<td>'+k+'</td>';}).join('')+'</tr>';}
  function mainRows(qs,sec){var h='<tr><td colspan="8" class="osec">'+esc(sec.title)+'</td></tr><tr><th>№</th><th>Критерий</th><th>Вопрос для экспертизы</th><th colspan="5">Ответ</th></tr>'+labels(sec.low,sec.high);
    qs.forEach(function(q){h+='<tr><td class="num">'+q.n+'</td><td><b>'+esc(q.criterion)+'</b></td><td>'+esc(q.question)+'</td>'+marks(e.q[q.code])+'</tr>';});return h;}
  var h='<div class="orgform"><p class="noprint"><button class="btn sm gray" id="bBackForm">← К электронной анкете</button> <button class="btn sm" onclick="window.print()">🖨 Печать единой формы</button></p>';
  h+='<div class="orgpage"><div class="opr">Приложение №2</div><div class="ocipher">ОКР шифр "'+esc(S.obj.shifr||'________________________________')+'"</div><table class="ot">';
  h+=mainRows(ORG.main.slice(0,10),ORG.sections[0])+mainRows(ORG.main.slice(10,16),ORG.sections[1])+'</table></div>';
  h+='<div class="orgpage"><table class="ot">'+mainRows(ORG.main.slice(16),ORG.sections[1]);
  h+='<tr><td colspan="8" class="osec">СТОП-ФАКТОРЫ</td></tr><tr><th>№</th><th colspan="2">Стоп-фактор</th><th colspan="5">Ответ</th></tr>'+labels('да','нет');
  ORG.sf.forEach(function(t,i){h+='<tr><td class="num">'+(i+1)+'</td><td colspan="2">'+esc(t)+'</td>'+marks(normSF(e.sf[i+1]))+'</tr>';});
  h+='<tr><td colspan="8" class="osec">«КРАСНЫЕ» ВОПРОСЫ</td></tr><tr><th>№</th><th colspan="2">Вопрос</th><th colspan="5">Ответ</th></tr>'+labels('да','нет');
  ORG.kv.forEach(function(t,i){h+='<tr><td class="num">'+(i+1)+'</td><td colspan="2">'+esc(t)+'</td>'+marks(e.kv[i+1])+'</tr>';});
  h+='</table></div></div>';return h;}
'''
block('function ankFormHTML(e){','/* ---------- навигация, сервис ---------- */',FORMNEW)
rep('.drawer .box{max-width:860px}', '''.orgform{background:#fff;color:#000;font-family:"Times New Roman",Times,serif}
.orgpage{width:190mm;min-height:277mm;margin:0 auto;padding:8mm 0;box-sizing:border-box}
.opr{text-align:right;font-weight:bold;font-size:12pt}.ocipher{text-align:center;font-weight:bold;font-size:12pt;margin:5mm 0}
.ot{width:100%;border-collapse:collapse;table-layout:fixed;font-size:8.5pt;line-height:1.12}
.ot td,.ot th{border:1px solid #000;padding:2px 3px;vertical-align:middle;word-wrap:break-word}
.ot th{text-align:center;font-weight:bold}.ot .osec{text-align:center;font-weight:bold;font-size:9pt}
.ot col,.ot td.circ{width:8mm;text-align:center}.ot td.num{width:7mm;text-align:center}.ot td.ol{font-size:7pt;text-align:center;font-weight:bold}.ot tr.onum td{text-align:center;font-weight:bold}
.ot td:nth-child(1){width:7mm}.ot td:nth-child(2){width:34mm}.ot td:nth-child(3){width:auto}.ot td:nth-child(n+4){width:13mm;text-align:center}
.orgform td.mark{font-weight:bold;background:#eee}
.drawer .box{max-width:860px}''')
rep('@media print{','@media print{\n  @page{size:A4 portrait;margin:10mm}\n  .orgpage{width:190mm;min-height:0;padding:0;margin:0;break-after:page;page-break-after:always}\n  .orgpage:last-child{break-after:auto;page-break-after:auto}\n  .orgform .noprint{display:none!important}\n  .ot{font-size:8pt}.ot td,.ot th{padding:1.5px 2px}')
rep('  calc:calc,','  calc:calc,\n  orgForm:ORG,normSF:normSF,kvSig:kvSig,')
rep('sfDet:false,sfUnc:false,kvGroup:[],flags:[],lows:[],wmode:DB.wmode,rmode:DB.rmode,irMod:false,calibOn:false};','sfDet:false,sfUnc:false,kvGroup:[],flags:[],lows:[],agenda:[],wmode:DB.wmode,rmode:DB.rmode,irMod:false,calibOn:false};')
rep('  R.agenda=[];\n','')
assert not re.search(r'[\u4e00-\u9fff]',t)
io.open(os.path.join(H,'model_v23_template.html'),'w',encoding='utf-8').write(t)
data=json.dumps(q,ensure_ascii=False,separators=(',',':'))
html=t.replace('/*__QEO21__*/null',data)
for p in ['output/versions/Аудит_НИОКР_модель_v23.0.html','output/Аудит_НИОКР_интерактивная_модель.html']:
    io.open(os.path.join(R,p),'w',encoding='utf-8').write(html)
print('OK model v23.0:',len(html.encode('utf-8')),'bytes; template',len(t))
