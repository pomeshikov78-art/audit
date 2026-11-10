# -*- coding: utf-8 -*-
# Этап 18Б, шаг 2: UI — опросник 15/30/60, стоп-факторы, зрелость, экономика, AHP(CR), версии v18.0.
# Запуск: /tmp/venv/bin/python3 scripts/patch18_ui.py
import re, json

HTML = 'output/Аудит_НИОКР_интерактивная_модель.html'
s = open(HTML, encoding='utf-8').read()
orig = len(s)

def rep(old, new, cnt=1):
    global s
    n = s.count(old)
    assert n == cnt, f'count({old[:70]!r}) = {n}, ждал {cnt}'
    s = s.replace(old, new)

# ============================================================ 1. renderCheck → опросник 15/30/60
i0 = s.find('function renderCheck(){')
assert i0 > 0
i1 = s.find('\n}\n', i0)
assert i1 > 0
OLD_RC = s[i0:i1 + 3]
assert 'expertCard(r)' in OLD_RC and len(OLD_RC) < 20000, len(OLD_RC)

NEW_RC = r'''function renderCheck(){
  const r=cur();if(!r.q)r.q={K:{},S:{},P:{}};
  const lvl=r.lvl||'S';const L=Q18.levels[lvl];
  const x=idx18(r,lvl),cc=concl18(r,lvl),st=cc.st,mt=mat18(r),ec=econ18(r);
  const qq=r.q[lvl]||(r.q[lvl]={});
  let h='';
  h+='<div class="card"><h3>🎚 Уровень аудита — опросник «'+L.name+'» ('+L.n+' вопросов)</h3>'+
   '<div style="display:flex;gap:8px;flex-wrap:wrap;margin:6px 0">'+LVLS.map(k=>'<button class="btn sm'+(k===lvl?'':' gray')+'" data-lvl="'+k+'">'+Q18.levels[k].name+' — '+Q18.levels[k].n+' вопр. · Coverage ≥ '+POROG[k]+'%</button>').join('')+'</div>'+
   '<div class="mini">Цель уровня: '+LVLGOAL[lvl]+'. Шкала 0–5, «НД» (нет данных — снижает Coverage) и «НП» (неприменимо — исключается из расчёта и знаменателя Coverage, обязательно письменное обоснование). Баллы 4–5 — только при указании доказательства; баллы 0–2 — только с обоснованием (иначе контрольное замечание).</div></div>';
  const idxChip=v=>v==null?'<span class="mini">н/д</span>':'<b>'+v.toFixed(1)+'%</b> <span style="color:'+(PLVLC[plvl(v)]||'#8c8c8c')+';font-weight:700">'+plvl(v)+'</span>';
  const concCol=(cc.kind==='nodata'||cc.kind==='cov')?'#8c8c8c':((cc.kind==='sf2'||cc.kind==='s4'||cc.kind==='p7a'||cc.kind==='k0')?'#c62828':((cc.kind==='k80'||cc.kind==='s1'||cc.kind==='p1'||cc.kind==='p2')?'#2e7d32':'#ef6c00'));
  h+='<div class="blk-sum" style="flex-wrap:wrap;gap:14px">'+
   '<span>I_А актуальность: '+idxChip(x.IA)+'</span><span>I_Ц целесообразность: '+idxChip(x.IC)+'</span><span>I_Р результативность: '+idxChip(x.IR)+'</span>'+
   '<span>Комплексный (∛): '+(x.kompl==null?'н/д':'<b>'+x.kompl.toFixed(1)+'%</b>')+'</span>'+
   '<span>Coverage: '+(x.cov==null?'н/д':'<b>'+x.cov.toFixed(1)+'%</b> '+(x.covOk?'<span style="color:#2e7d32">✔ ≥ '+x.thr+'%</span>':'<span style="color:#c62828">✘ &lt; '+x.thr+'% — положительное заключение не выносится</span>'))+'</span>'+
   '<span>Отвечено: <b>'+x.nAns+'</b> / '+L.n+'</span><span>Контрольных замечаний: <b>'+x.ctrl+'</b></span><span>Балл ≤ 2: <b>'+x.low2+'</b></span></div>';
  h+='<div class="note" style="border-left:5px solid '+concCol+';background:rgba(127,127,127,.06)"><b>ЗАКЛЮЧЕНИЕ ' + L.name.toUpperCase() + 'ОГО ОПРОСНИКА (авто):</b> <span style="color:'+concCol+';font-weight:700">'+cc.txt+'</span></div>';
  const recs=autoRecs18(r);
  if(recs.length)h+='<div class="note" style="background:#fdf3e7;border-color:#ef6c00"><b>🎯 Авторекомендации (балл ≤ 2, уровень «Средний»):</b><ul class="mini" style="margin:6px 0 0 18px">'+recs.map(a=>'<li><b>'+a.code+'</b> = '+a.val+' → '+a.rec+'</li>').join('')+'</ul></div>';
  const gen=genRules18(r,lvl,x);
  if(gen.length)h+='<div class="note mini" style="background:#fbeeee;border-color:#c62828"><b>Сработали общие правила заключения:</b> '+gen.map(g=>g.rule).join('; ')+'</div>';
  h+='<div style="display:flex;gap:8px;margin:8px 0;flex-wrap:wrap"><button class="btn sm" id="q18Concl">📋 Экспертное заключение (Ф-06)</button><button class="btn sm gray" id="q18Md">⬇ Опросник (MD)</button></div>';
  for(let dd=0;dd<3;dd++){
    const dn=['А. АКТУАЛЬНОСТЬ','Б. ЦЕЛЕСООБРАЗНОСТЬ','В. РЕЗУЛЬТАТИВНОСТЬ'][dd];
    h+='<div class="blk-h">'+dn+'<span style="font-weight:400;font-size:12px"> · индекс направления: '+idxChip(x.dirs[dd])+'</span></div>';
    if(lvl==='P'){
      L.blocks.filter(b=>b.dir===dd).forEach(b=>{
        const xb=x.blocks.filter(t=>t.code===b.code)[0]||{E:b.w,dk:null};
        h+='<div class="blk-sum"><span>Блок <b>'+b.code+'</b> «'+b.name+'» (вопросы '+b.from+'–'+b.to+')</span><span>вес: <b>'+((r.ahp18&&r.ahp18.useBlocks)?xb.E.toFixed(3)+' (AHP)':String(b.w))+'</b></span><span>заполненность D_k: <b>'+(xb.dk==null?'н/д':xb.dk.toFixed(1)+'%')+'</b></span></div>';
        L.q.filter(q=>q.blk===b.code).forEach(q=>{h+=qRow18(r,lvl,q,qq[q.code]);});});
    }else{
      L.q.filter(q=>q.dir===dd).forEach(q=>{h+=qRow18(r,lvl,q,qq[q.code]);});
    }
  }
  h+='<div class="blk-h">⛔ СТОП-ФАКТОРЫ<span style="font-weight:400;font-size:12px"> · '+st.res+'</span></div>';
  h+='<div class="card">'+Q18.sf2.map((t,i)=>sfRow18(r,i,'<b>Проверка SF-2 №'+(i+1)+':</b> '+t)).join('')+
    [0,1,2,3,4].map(j=>sfRow18(r,5+j,'<b>Иное обстоятельство '+(j+1)+'</b>')).join('')+
    '<div class="mini" style="margin-top:6px">Статусы: «не выявлено» → 0 · «выявлена неопределённость» → SF-1 (обязательная дополнительная проверка) · «подтверждено» → SF-2 (эскалация руководству; безусловно положительное заключение невозможно до рассмотрения критического обстоятельства).</div></div>';
  h+='<div class="blk-h">🧬 ЗРЕЛОСТЬ РЕЗУЛЬТАТА M0–M9</div>';
  h+='<div class="card"><div style="display:flex;gap:16px;flex-wrap:wrap">'+
   '<label>Текущий уровень: <select data-mat="cur"><option value="">—</option>'+Q18.m09.map((m2,i)=>'<option value="'+i+'"'+(mt.cur===i?' selected':'')+'>'+m2[0]+' «'+m2[1]+'»</option>').join('')+'</select></label>'+
   '<label>Планируемый (по этапу): <select data-mat="plan"><option value="">—</option>'+Q18.m09.map((m2,i)=>'<option value="'+i+'"'+(mt.plan===i?' selected':'')+'>'+m2[0]+' «'+m2[1]+'»</option>').join('')+'</select></label>'+
   '<label>Подтверждён испытаниями: <select data-mat="tests"><option value="Нет"'+(mt.tests!=='Да'?' selected':'')+'>Нет</option><option value="Да"'+(mt.tests==='Да'?' selected':'')+'>Да</option></select></label></div>'+
   '<div class="mini" style="margin-top:6px">Отклонение (факт − план): <b>'+(mt.dev==null?'—':mt.dev)+'</b> · '+(mt.note||'—')+' · Автозамечание: <b style="color:'+(mt.warn==='OK'?'#2e7d32':'#c62828')+'">'+mt.warn+'</b></div></div>';
  h+='<div class="blk-h">💰 ЭКОНОМИКА: NPV / BCR / LCC / АЛЬТЕРНАТИВЫ</div>'+econCard18(r,ec);
  const ea=early18(r,x,st);
  h+='<div class="blk-h">🚨 РАННИЕ СИГНАЛЫ</div><div class="card">'+(ea.length?ea.map(sg=>'<div class="mini" style="margin:3px 0"><b style="color:#c62828">'+sg.sig+'</b> — '+sg.rule+' → '+sg.act+'</div>').join(''):'<div class="mini">Сигналов нет. Правила: падение I_А/I_Ц на ≥10 п. и I_Р на ≥10 п. (против предыдущих значений — задаются в паспорте/снимках), Coverage &lt; 85%, появление SF-2, превышение срока этапа.</div>')+'</div>';
  if(r.legacy31)h+='<div class="note mini">ℹ Эта НИОКР мигрировала с чек-листа v17 (31 вопрос): прежние ответы сохранены в архиве legacy31, семантический перенос выполнен в опросник «Средний» (30).</div>';
  document.getElementById('tab-check').innerHTML=h;
  document.querySelectorAll('#tab-check [data-lvl]').forEach(b=>b.onclick=()=>{const rr=cur();rr.lvl=b.dataset.lvl;save();renderCheck();renderDash();
    notify('info','Уровень аудита: '+LVLN[rr.lvl]+' ('+Q18.levels[rr.lvl].n+' вопросов)','порог Coverage ≥ '+POROG[rr.lvl]+'% · '+LVLGOAL[rr.lvl]);});
  document.querySelectorAll('#tab-check select[data-qa]').forEach(el=>el.onchange=e=>setQA18(el.dataset.qa,'a',e.target.value));
  document.querySelectorAll('#tab-check input[data-qj]').forEach(el=>el.onchange=e=>setQA18(el.dataset.qj,'j',e.target.value));
  document.querySelectorAll('#tab-check input[data-qe]').forEach(el=>el.onchange=e=>setQA18(el.dataset.qe,'e',e.target.value));
  document.querySelectorAll('#tab-check select[data-sf]').forEach(el=>el.onchange=e=>{const rr=cur();rr.sf[+el.dataset.sf]=e.target.value;save();renderCheck();renderDash();
    if(sfVal(e.target.value)===2)notify('err','⛔ SF-2 подтверждён','эскалация руководству; безусловно положительное заключение невозможно; требуется внеплановое заседание комиссии');
    else if(sfVal(e.target.value)===1)notify('warn','⚠ SF-1: выявлена неопределённость','требуется отдельный план дополнительной проверки');});
  document.querySelectorAll('#tab-check input[data-sfc]').forEach(el=>el.onchange=e=>{const rr=cur();if(!rr.sfc)rr.sfc=['','','','',''];rr.sfc[+el.dataset.sfc]=e.target.value;save();});
  document.querySelectorAll('#tab-check select[data-mat]').forEach(el=>el.onchange=e=>{const rr=cur();const k=el.dataset.mat;
    rr.mat[k]=(k==='tests')?e.target.value:(e.target.value===''?null:+e.target.value);save();renderCheck();});
  wireEcon18();
  const bc=document.getElementById('q18Concl');if(bc)bc.onclick=()=>concl18Modal(cur());
  const bm=document.getElementById('q18Md');if(bm)bm.onclick=()=>exportMD18dl(cur());
}
function qRow18(r,lvl,q,cell){
  cell=cell||{a:'',j:'',e:''};const fl=qFlags(cell);const n=qNum(cell.a);
  const SC={'5':'#2e7d32','4':'#7cb342','3':'#fbc02d','2':'#ef6c00','1':'#e65100','0':'#c62828','НД':'#8c8c8c','НП':'#9e9e9e'};
  let hint='';
  if(lvl==='K')hint='Якоря: 5 — '+q.a5+' · 3 — '+q.a3+' · 0 — '+q.a0;
  else if(lvl==='S')hint=(q.check?'Проверить: '+q.check+'. ':'')+(q.ev?'Доказательства: '+q.ev+'. ':'')+(q.anch?'Якоря: 5 — '+q.anch['5']+' 3 — '+q.anch['3']+' 0 — '+q.anch['0']:'');
  else hint=(q.est?'Оценить: '+q.est+'. ':'')+(q.ev?'Доказательства: '+q.ev:'');
  const code=(lvl==='P')?q.n:q.code;
  const av=(cell.a===''||cell.a==null)?'—':String(cell.a);
  return '<div class="qrow"><div class="qh"><div class="num" style="font-size:11px;min-width:36px">'+code+'</div><div class="qt"><b>'+q.q+'</b>'+(lvl==='S'&&q.w!==1?' <span class="mini" title="Вес вопроса (Σ весов раздела = 10)">w='+q.w+'</span>':'')+(n!=null&&n<=2?' <span title="Балл ≤ 2 — требуется обоснование и корректирующие меры" style="color:#c62828;font-weight:800">≤2</span>':'')+(fl?' <span title="'+fl.replace(/"/g,'')+'" style="color:#ef6c00;font-weight:800">⚠</span>':'')+'<span class="est">'+hint+'</span></div>'+
   '<span class="score-chip" style="background:'+((cell.a===''||cell.a==null)?'#dfe6f2;color:#8c8c8c':((SC[String(cell.a)]||'#dfe6f2')+';color:#fff'))+'">'+av+'</span></div>'+
   '<div class="qb"><label>Ответ: <select data-qa="'+q.code+'"><option value="">— не выбран —</option>'+Q18.scale.map(o=>'<option value="'+o[0]+'"'+(String(cell.a)===o[0]?' selected':'')+'>'+o[0]+' · '+o[1]+'</option>').join('')+'</select></label>'+
   '<input type="text" data-qj="'+q.code+'" placeholder="Обоснование (обязательно для 0–2 и «НП»)" value="'+String(cell.j||'').replace(/"/g,'&quot;')+'" style="flex:1;min-width:180px">'+
   '<input type="text" data-qe="'+q.code+'" placeholder="Доказательство / документ (обязательно для 3–5)" value="'+String(cell.e||'').replace(/"/g,'&quot;')+'" style="flex:1;min-width:180px"></div></div>';
}
function sfRow18(r,i,label,custom){
  const v=(r.sf||[])[i]||'не выявлено';
  return '<div style="display:flex;gap:10px;align-items:center;margin:4px 0;flex-wrap:wrap"><span class="mini" style="flex:1;min-width:280px">'+label+'</span>'+
  (custom?'<input type="text" data-sfc="'+(i-5)+'" placeholder="Вписать обстоятельство" value="'+String((r.sfc||[])[i-5]||'').replace(/"/g,'&quot;')+'" style="flex:1;min-width:180px">':'')+
  '<select data-sf="'+i+'">'+['не выявлено','выявлена неопределённость','подтверждено'].map(o=>'<option'+(v===o?' selected':'')+'>'+o+'</option>').join('')+'</select></div>';
}
function setQA18(code,field,val){
  const r=cur();const lvl=r.lvl||'S';if(!r.q[lvl])r.q[lvl]={};
  if(!r.q[lvl][code])r.q[lvl][code]={a:'',j:'',e:''};
  r.q[lvl][code][field]=val;
  save();
  if(field!=='a')return;
  renderCheck();renderDash();
  const cell=r.q[lvl][code];
  if(val==='НП'&&!String(cell.j||'').trim()){
    showModal('<h2>«НП» — требуется письменное обоснование</h2><p class="mini">Ответ «неприменимо» исключает вопрос из индекса и из знаменателя Coverage. Методика требует письменного обоснования — иначе фиксируется контрольное замечание.</p><div class="field"><label>Обоснование «НП»</label><input type="text" id="npJust18" style="width:100%" placeholder="Почему вопрос неприменим на данной стадии"></div>',
     [{label:'Сохранить обоснование',fn:()=>{const t=document.getElementById('npJust18').value;r.q[lvl][code].j=t;save();renderCheck();notify(t?'ok':'warn',t?'Обоснование «НП» сохранено':'«НП» без обоснования',t?'вопрос исключён из расчёта корректно':'зафиксировано контрольное замечание в сводке опросника');}},{label:'Позже',cls:'gray'}]);}
  else{const n2=qNum(val);
    if(n2!=null&&n2>=3&&!String(cell.e||'').trim())notify('warn','Балл '+n2+' без доказательства','укажите документ/подтверждение в поле «Доказательство» — иначе контрольное замечание «3–5 без доказательства»');
    else if(n2!=null&&n2<=2&&!String(cell.j||'').trim())notify('warn','Балл '+n2+' без обоснования','баллы 0–2 требуют обоснования; сработала авторекомендация (уровень «Средний»)');}
}
function econCard18(r,ec){
  if(!r.econ)r.econ={rate:0.1,T:5,lcc:Array.from({length:8},()=>['','']),flows:Array.from({length:11},()=>['','']),altW:Array(9).fill(1),alt:{c:Array(9).fill(''),d:Array(9).fill(''),e:Array(9).fill(''),f:Array(9).fill('')}};
  const e=r.econ;
  while((e.flows||[]).length<11)e.flows.push(['','']);
  while((e.lcc||[]).length<8)e.lcc.push(['','']);
  let h='<div class="card"><div style="display:flex;gap:16px;flex-wrap:wrap;align-items:center">'+
   '<label>Ставка дисконтирования r: <input type="number" step="0.01" data-econ="rate" value="'+(e.rate==null?0.1:e.rate)+'" style="width:80px"></label>'+
   '<span class="mini">NPV = <b>'+(ec.npv==null?'—':ec.npv)+'</b> млн ₽ · PV(выгоды) = '+ec.pvB+' · PV(затраты) = '+ec.pvC+'</span>'+
   '<span class="mini">BCR = <b>'+(ec.bcr==null?'—':ec.bcr)+'</b> → '+ec.bcrTxt+'</span>'+
   '<span class="mini">LCC (PV) = <b>'+(ec.lccPv==null?'—':ec.lccPv)+'</b> млн ₽</span></div>';
  h+='<div class="mini" style="margin-top:8px"><b>Денежные потоки по годам (NPV/BCR):</b></div><table><tr><th>t</th><th>Эффект B_t, млн ₽</th><th>Затраты C_t, млн ₽</th></tr>'+
   e.flows.map((f,t)=>'<tr><td>'+t+'</td><td><input data-econ="fb'+t+'" value="'+(f[0]==null?'':f[0])+'" style="width:110px"></td><td><input data-econ="fc'+t+'" value="'+(f[1]==null?'':f[1])+'" style="width:110px"></td></tr>').join('')+'</table>';
  h+='<div class="mini" style="margin-top:8px"><b>LCC — полная стоимость жизненного цикла (8 стадий; в затратах учитываются не только контракты, но и эксплуатация, ремонт, вывод):</b></div><table><tr><th>Стадия</th><th>Сумма, млн ₽</th><th>Год от старта</th></tr>'+
   Q18.lcc.map((L2,i)=>'<tr><td class="mini">'+L2[0]+' <span style="color:#8c8c8c">('+L2[1]+')</span></td><td><input data-econ="lv'+i+'" value="'+((e.lcc[i]||[])[0]==null?'':(e.lcc[i]||[])[0])+'" style="width:100px"></td><td><input data-econ="lt'+i+'" value="'+((e.lcc[i]||[])[1]==null?'':(e.lcc[i]||[])[1])+'" style="width:60px"></td></tr>').join('')+'</table>';
  h+='<div class="mini" style="margin-top:8px"><b>Сравнение альтернатив (функция полезности U_s = Σ w_k · x_ks; оценки 0–5):</b> '+(ec.Us?('наилучший сценарий — <b style="color:#2e7d32">'+ec.best+'</b> (преимущество над «Продолжить разработку»: '+ec.margin+')'):'заполните оценки — модель выберет наилучший сценарий')+'</div>'+
   '<table><tr><th>Критерий</th><th>Вес</th>'+ALTNAMES4.map((n3,i)=>'<th'+(i===0?' style="color:#2e74b5"':'')+'>'+n3+(ec.Us?'<br><span class="mini">U = '+ec.Us[i]+'</span>':'')+'</th>').join('')+'</tr>'+
   Q18.altcrit.map((cr2,i)=>'<tr><td class="mini">'+cr2+'</td><td><input data-econ="aw'+i+'" value="'+(e.altW[i]==null?1:e.altW[i])+'" style="width:52px"></td>'+['c','d','e','f'].map(k=>'<td><input data-econ="a'+k+i+'" value="'+((e.alt[k]||[])[i]==null?'':(e.alt[k]||[])[i])+'" style="width:52px"></td>').join('')+'</tr>').join('')+'</table></div>';
  return h;}
function wireEcon18(){
  document.querySelectorAll('#tab-check [data-econ]').forEach(el=>el.onchange=ev=>{
    const r=cur();const k=el.dataset.econ;const v=el.value;const e=r.econ;
    if(k==='rate')e.rate=(v===''?null:Number(v));
    else if(k.slice(0,2)==='fb')e.flows[+k.slice(2)][0]=v;
    else if(k.slice(0,2)==='fc')e.flows[+k.slice(2)][1]=v;
    else if(k.slice(0,2)==='lv')e.lcc[+k.slice(2)][0]=v;
    else if(k.slice(0,2)==='lt')e.lcc[+k.slice(2)][1]=v;
    else if(k.slice(0,2)==='aw')e.altW[+k.slice(2)]=(v===''?1:Number(v));
    else if(k[0]==='a'&&'cdef'.indexOf(k[1])>=0)e.alt[k[1]][+k.slice(2)]=v;
    save();renderCheck();});}
function concl18Modal(r){
  const lvl=r.lvl||'S',x=idx18(r,lvl),cc=concl18(r,lvl),st=cc.st,mt=mat18(r),ec=econ18(r),a=ahp18cr(r);
  const f=v=>v==null?'недостаточно данных':(Math.round(v*10)/10).toFixed(1);
  const recs=autoRecs18(r),gen=genRules18(r,lvl,x);
  const lines=['ЭКСПЕРТНОЕ ЗАКЛЮЧЕНИЕ АУДИТА НИОКР (форма Ф-06)',
  'НИОКР: «'+r.name+'» (шифр '+(r.shifr||'—')+')',
  'Уровень аудита: опросник «'+LVLN[lvl]+'» ('+Q18.levels[lvl].n+' вопросов) — '+LVLGOAL[lvl]+'.',
  '',
  '1. Индекс актуальности: I_А = '+f(x.IA)+'% — уровень «'+(plvl(x.IA)||'—')+'».',
  '2. Индекс целесообразности: I_Ц = '+f(x.IC)+'% — уровень «'+(plvl(x.IC)||'—')+'».',
  '3. Индекс результативности: I_Р = '+f(x.IR)+'% — уровень «'+(plvl(x.IR)||'—')+'».',
  '4. Комплексный индекс (геометрическое среднее, справочно): '+f(x.kompl)+'%.',
  '5. Полнота доказательств (Coverage): '+f(x.cov)+'% при пороге ≥ '+x.thr+'% — '+(x.cov==null?'нет данных':(x.covOk?'порог пройден':'НИЖЕ ПОРОГА — положительное заключение не выносится'))+'.',
  '6. Стоп-факторы: '+st.res+(st.list.length?'. Сработали SF-2: '+st.list.join('; '):'')+'.',
  '7. Зрелость результата: '+(mt.name||'не указана')+'; план M'+(mt.plan==null?'—':mt.plan)+'; '+(mt.note||'отклонение не рассчитано')+'; контроль доказательств: '+mt.warn+'.',
  '8. AHP (согласованность экспертных матриц): CR max = '+a.crMax.toFixed(3)+(a.ok?' — согласовано (≤ 0,10)':' — НЕ согласовано (> 0,10), пересмотрите сравнения')+'.',
  '9. Экономика: '+(ec.npv==null?'не заполнена':'NPV = '+ec.npv+' млн ₽; BCR = '+(ec.bcr==null?'—':ec.bcr)+' ('+ec.bcrTxt+')'+(ec.best?'; наилучший сценарий: '+ec.best+' (отрыв '+ec.margin+')':''))+'.',
  '10. ЗАКЛЮЧЕНИЕ: '+cc.txt];
  if(recs.length)lines.push('11. Авторекомендации (баллы ≤ 2, уровень «Средний»): '+recs.map(t=>t.code+' → '+t.rec).join('; ')+'.');
  if(gen.length)lines.push((recs.length?'12':'11')+'. Сработали общие правила: '+gen.map(g=>g.rule).join('; ')+'.');
  lines.push('','Председатель комиссии _______________    Секретарь _______________    Дата: '+new Date().toLocaleDateString('ru-RU'));
  const txt=lines.join('\n');
  showModal('<h2>📋 Экспертное заключение — уровень «'+LVLN[lvl]+'» (Ф-06)</h2><pre style="white-space:pre-wrap;font-family:inherit;background:rgba(127,127,127,.08);padding:12px;border-radius:8px">'+txt.replace(/</g,'&lt;')+'</pre>',
   [{label:'📋 Копировать',fn:()=>{if(navigator.clipboard)navigator.clipboard.writeText(txt);notify('ok','Заключение скопировано','Ф-06 — для вставки в протокол комиссии / Word');}},{label:'Закрыть',cls:'gray'}]);}
function exportMD18(r){
  const lvl=r.lvl||'S',L=Q18.levels[lvl],x=idx18(r,lvl),cc=concl18(r,lvl),st=cc.st,mt=mat18(r),ec=econ18(r),recs=autoRecs18(r),gen=genRules18(r,lvl,x),qq=(r.q&&r.q[lvl])||{};
  const f=v=>v==null?'н/д':(Math.round(v*10)/10).toFixed(1);
  const M=[];
  M.push('# ОПРОСНИК АУДИТА НИОКР — УРОВЕНЬ «'+L.name.toUpperCase()+'» ('+L.n+' ВОПРОСОВ, заполненный)');M.push('');
  M.push('**НИОКР:** «'+r.name+'» (шифр: '+(r.shifr||'—')+')  ');
  M.push('**Уровень аудита:** '+LVLN[lvl]+' — '+LVLGOAL[lvl]+'  ');
  M.push('**Дата формирования:** '+new Date().toLocaleDateString('ru-RU')+'  ');M.push('');
  ['Актуальность (I_А)','Целесообразность (I_Ц)','Результативность (I_Р)'].forEach((nm,dd)=>{
    M.push('## Направление «'+nm+'» — '+f(x.dirs[dd])+' % ('+(x.dirs[dd]==null?'нет данных':plvl(x.dirs[dd]))+')');M.push('');
    M.push('| Код | Вопрос | Ответ | Вес | Обоснование | Доказательство | Замечание |');
    M.push('| --- | --- | --- | --- | --- | --- | --- |');
    L.q.filter(q=>q.dir===dd).forEach(q=>{const c2=qq[q.code]||{};const fl=qFlags(c2);
      M.push('| '+((lvl==='P')?(q.n+' ('+q.blk+')'):q.code)+' | '+q.q.replace(/\|/g,'/')+' | '+((c2.a===''||c2.a==null)?'—':c2.a)+' | '+(q.w||1)+' | '+String(c2.j||'').replace(/\|/g,'/')+' | '+String(c2.e||'').replace(/\|/g,'/')+' | '+fl.replace(/\|/g,'/')+' |');});
    M.push('');});
  M.push('## СВОДНЫЕ ПОКАЗАТЕЛИ');M.push('');
  M.push('| Показатель | Значение |');M.push('| --- | --- |');
  M.push('| I_А актуальность | '+f(x.IA)+' % — '+(plvl(x.IA)||'—')+' |');
  M.push('| I_Ц целесообразность | '+f(x.IC)+' % — '+(plvl(x.IC)||'—')+' |');
  M.push('| I_Р результативность | '+f(x.IR)+' % — '+(plvl(x.IR)||'—')+' |');
  M.push('| Комплексный индекс (∛ I_А·I_Ц·I_Р) | '+f(x.kompl)+' % |');
  M.push('| Полнота доказательств Coverage | '+f(x.cov)+' % (порог ≥ '+x.thr+' %) '+(x.cov==null?'':(x.covOk?'✔':'✘'))+' |');
  M.push('| Отвечено вопросов | '+x.nAns+' из '+L.n+' |');
  M.push('| Контрольных замечаний | '+x.ctrl+' |');
  M.push('| Вопросов с баллом ≤ 2 | '+x.low2+' |');M.push('');
  M.push('## ЗАКЛЮЧЕНИЕ (уровень «'+LVLN[lvl]+'»)');M.push('');M.push('> '+cc.txt);M.push('');
  M.push('**Стоп-факторы:** '+st.res+(st.list.length?'. Сработали SF-2: '+st.list.join('; '):'')+'  ');
  M.push('**Зрелость:** '+(mt.name||'не указана')+'; '+(mt.note||'—')+'; '+mt.warn+'  ');
  M.push('**Экономика:** '+(ec.npv==null?'не заполнена':'NPV '+ec.npv+' млн ₽; BCR '+(ec.bcr==null?'—':ec.bcr)+' ('+ec.bcrTxt+')'+(ec.best?'; наилучший сценарий: '+ec.best:''))+'  ');
  if(recs.length){M.push('');M.push('### Авторекомендации (баллы ≤ 2)');recs.forEach(t=>M.push('- **'+t.code+'** = '+t.val+' → '+t.rec));}
  if(gen.length){M.push('');M.push('### Сработавшие общие правила');gen.forEach(g=>M.push('- '+g.rule));}
  M.push('');M.push('---');M.push('');M.push('Председатель комиссии _______________ · Секретарь _______________');M.push('');
  return M.join('\n');}
function exportMD18dl(r){
  const nm='Вопросник_НИОКР_'+String(r.shifr||'шифр').replace(/[\\/:*?"<>|]+/g,'_')+'_'+LVLN[r.lvl||'S']+'_'+Q18.levels[r.lvl||'S'].n+'.md';
  download(nm,exportMD18(r),'text/markdown');
  notify('ok','Опросник выгружен в Markdown',nm+' — формат для вставки в Word/протокол комиссии');}
function dashCard18(r){
  const lvl=r.lvl||'S',x=idx18(r,lvl),cc=concl18(r,lvl);
  const chip=(nm,v)=>'<span style="margin-right:16px">'+nm+': '+(v==null?'<b>н/д</b>':'<b>'+v.toFixed(1)+'%</b> <span style="color:'+(PLVLC[plvl(v)]||'#8c8c8c')+';font-weight:700">'+plvl(v)+'</span>')+'</span>';
  const concCol=(cc.kind==='nodata'||cc.kind==='cov')?'#8c8c8c':((cc.kind==='sf2'||cc.kind==='s4'||cc.kind==='p7a'||cc.kind==='k0')?'#c62828':((cc.kind==='k80'||cc.kind==='s1'||cc.kind==='p1'||cc.kind==='p2')?'#2e7d32':'#ef6c00'));
  return '<div class="card"><h3>🧭 Экспертная оценка — опросник «'+LVLN[lvl]+'» ('+Q18.levels[lvl].n+' вопросов)</h3>'+
   '<div class="mini" style="line-height:2">'+chip('I_А',x.IA)+chip('I_Ц',x.IC)+chip('I_Р',x.IR)+'<br>Комплексный индекс: <b>'+(x.kompl==null?'н/д':x.kompl.toFixed(1)+'%')+'</b> · Coverage: <b>'+(x.cov==null?'н/д':x.cov.toFixed(1)+'%')+'</b> '+(x.cov==null?'':((x.covOk?'✔ ≥ ':'✘ &lt; ')+x.thr+'%'))+' · отвечено '+x.nAns+'/'+Q18.levels[lvl].n+' · замечаний '+x.ctrl+'</div>'+
   '<div class="note" style="margin-top:8px;border-left:5px solid '+concCol+'"><b>Заключение:</b> <span style="color:'+concCol+';font-weight:700">'+cc.txt+'</span></div>'+
   '<div class="mini">Перспективы завершения: <b>'+expertProspect18(r)+'</b>'+((r.sf||[]).some(v=>sfVal(v)===2)?' · <b style="color:#c62828">⛔ выявлен SF-2 — эскалация руководству</b>':'')+'</div></div>';}
function renderAhp18(){
  const r=cur();
  if(!r.ahp18)r.ahp18={dir:{a12:'1',a13:'1',a23:'1'},A:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},B:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},V:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},useBlocks:false};
  const st=ahp18cr(r);
  const SAATY=['1','2','3','4','5','6','7','8','9','1/2','1/3','1/4','1/5','1/6','1/7','1/8','1/9'];
  const pairs=n=>{const p=[];for(let i=1;i<=n;i++)for(let j=i+1;j<=n;j++)p.push([i,j]);return p;};
  const card=(key,title,names,one,ri)=>{const o=r.ahp18[key]||{};
    return '<div class="card"><h3>'+title+'</h3><table><tr><th>Парное сравнение</th><th>Во сколько раз строка важнее столбца (шкала Сати 1–9 и обратные 1/2…1/9)</th></tr>'+
    pairs(names.length).map(p=>'<tr><td class="mini">'+names[p[0]-1]+' ↔ '+names[p[1]-1]+'</td><td><select data-ahp="'+key+'" data-k="a'+p[0]+p[1]+'">'+SAATY.map(v=>'<option'+(String(o['a'+p[0]+p[1]]||'1')===v?' selected':'')+'>'+v+'</option>').join('')+'</select></td></tr>').join('')+'</table>'+
    '<div class="mini" style="margin-top:6px;line-height:1.9">Веса (нормированное геометрическое среднее): '+one.w.map((wv,i)=>names[i]+' — <b>'+(wv*100).toFixed(1)+'%</b>').join(' · ')+'<br>'+
    'λmax = '+one.lmax.toFixed(4)+' · CI = (λmax − n)/(n − 1) = '+one.ci.toFixed(4)+' · RI('+names.length+') = '+ri+' · <b style="color:'+(one.ok?'#2e7d32':'#c62828')+'">CR = '+one.cr.toFixed(3)+(one.ok?' ✔ согласовано (≤ 0,10)':' ✘ не согласовано (> 0,10) — пересмотрите сравнения')+'</b></div></div>';};
  let h='<div class="note">Метод анализа иерархий (Т. Сати) — <b>реальный расчёт согласованности</b>: λmax → CI = (λmax − n)/(n − 1) → CR = CI / RI (таблица Сати: n=3 → 0,58; n=4 → 0,90). CR ≤ 0,10 — матрица согласована. Веса блоков А/Б/В подставляются в «Полный» опросник (60) при включённом переключателе; веса направлений могут использоваться комиссией при взвешивании I_А/I_Ц/I_Р.</div>';
  h+=card('dir','Матрица 1 — направления оценки НИОКР (3×3)',['Актуальность','Целесообразность','Результативность'],st.dir,Q18.ri[2]);
  ['A','B','V'].forEach((k,ki)=>{const sec=['А','Б','В'][ki];
    h+=card(k,'Матрица блоков «'+sec+'» (4×4) — полный опросник',[1,2,3,4].map(i=>sec+i),st[k],Q18.ri[3]);});
  h+='<div class="card"><h3>Сводный контроль согласованности</h3><div class="mini">CR max по четырём матрицам: <b style="color:'+(st.ok?'#2e7d32':'#c62828')+'">'+st.crMax.toFixed(3)+'</b> — '+(st.ok?'✔ результаты AHP применимы':'✘ хотя бы одна матрица не согласована — веса не применяются до исправления')+'</div>'+
   '<label style="display:block;margin-top:8px"><input type="checkbox" id="ahpUse18"'+(r.ahp18.useBlocks?' checked':'')+'> Использовать веса блоков AHP в «Полном» опроснике (60 вопросов)</label></div>';
  document.getElementById('tab-wgt').innerHTML=h;
  document.querySelectorAll('#tab-wgt select[data-ahp]').forEach(el=>el.onchange=e=>{const rr=cur();
    rr.ahp18[el.dataset.ahp][el.dataset.k]=e.target.value;save();renderAhp18();renderCheck();renderDash();});
  const ub=document.getElementById('ahpUse18');
  if(ub)ub.onchange=e=>{const rr=cur();rr.ahp18.useBlocks=e.target.checked;save();renderAhp18();renderCheck();renderDash();
    notify('info','Веса блоков AHP для «Полного» '+(e.target.checked?'ВКЛЮЧЕНЫ':'ВЫКЛЮЧЕНЫ'),e.target.checked?'D_k блоков взвешиваются по матрицам AHP (4×4)':'базовые веса: А 0,20/0,20/0,20/0,40 · Б 0,25×4 · В 0,30/0,25/0,30/0,15');};
}
'''
assert "'''" not in NEW_RC
s = s[:i0] + NEW_RC + s[i1 + 3:]

# ============================================================ 2. переключения рендеров
rep('renderEvid();renderAnal();renderWeights();renderPass();', 'renderEvid();renderAnal();renderAhp18();renderPass();')
rep('else if(t==="wgt")renderWeights();', 'else if(t==="wgt")renderAhp18();')
rep('  expertDashCard(r)+', '  dashCard18(r)+')
rep('data-t="check">✅ Чек-лист (31)<', 'data-t="check">✅ Опросник 15/30/60<')

# ============================================================ 3. KPI-карточка дашборда
i = s.find('ИНТЕГРАЛЬНЫЙ ПОКАЗАТЕЛЬ I')
assert i > 0
ls = s.rfind("'<div class=\"card kpi\">", 0, i)
le = s.find("</span></div>'+", i) + len("</span></div>'+")
assert ls > 0 and le > i
NEW_KPI = ("'<div class=\"card kpi\"><span class=\"l\">КОМПЛЕКСНЫЙ ИНДЕКС I (∛ I_А·I_Ц·I_Р) — уровень «'+c.lvlName+'»</span>"
           "<span class=\"v\">'+(c.I==null?\"н/д\":c.I.toFixed(1))+' <span style=\"font-size:14px;color:#8c8c8c\">/ 100</span></span>"
           "<span class=\"l\">'+(c.cat||\"недостаточно данных\")+' · I_А '+(c.iA==null?\"н/д\":c.iA.toFixed(1)+\"%\")+' · I_Ц '+(c.iC==null?\"н/д\":c.iC.toFixed(1)+\"%\")+' · I_Р '+(c.iR==null?\"н/д\":c.iR.toFixed(1)+\"%\")+"
           "' · Coverage '+(c.cov==null?\"н/д\":c.cov.toFixed(1)+\"% \"+(c.covOk?\"✔ ≥\":\"✘ <\")+c.thr+\"%\")+'</span></div>'+")
s = s[:ls] + NEW_KPI + s[le:]

# ============================================================ 4. версионные строки
rep('<title>Аудит НИОКР — интерактивная модель оценки, v17.0</title>',
    '<title>Аудит НИОКР — интерактивная модель оценки, v18.0</title>')
rep('class="ver">v17.0 · этап 17</span>', 'class="ver">v18.0 · этап 18</span>')
n = s.count('аудита НИОКР v17.0')
assert n >= 8, n
s = s.replace('аудита НИОКР v17.0', 'аудита НИОКР v18.0')
n2 = s.count('НИОКР_аудит_v17.json')
assert n2 == 2, n2
s = s.replace('НИОКР_аудит_v17.json', 'НИОКР_аудит_v18.json')
rep('Добро пожаловать в модель аудита НИОКР (v17.0)', 'Добро пожаловать в модель аудита НИОКР (v18.0)')
i = s.find('Новое в v17.0')
assert i > 0
ps = s.rfind('<p', 0, i); pe = s.find('</p>', i) + 4
NEW_P = ('<p><b>Новое в v18.0 (этап 18Б):</b> вместо чек-листа 31 вопрос — <b>трёхуровневая система аудита «15 / 30 / 60»</b> '
         'по методике «Научно.docx»: Краткий опросник (15 вопросов, экспресс-оценка, Coverage ≥ 80 %), Средний (30 вопросов с весами, углублённая оценка, ≥ 85 %), '
         'Полный (60 вопросов в 12 блоках, комплексный аудит с AHP и экономикой, ≥ 90 %). Каждый уровень оценивает ТРИ направления — '
         'актуальность I_А, целесообразность I_Ц, результативность I_Р — и даёт собственную, неповторяющуюся формулировку заключения '
         '(4 формулировки / категории I–IV / решения R1–R7). Новое: шкала 0–5 с «НД» и «НП», стоп-факторы SF-0/1/2 с блокировкой положительного заключения, '
         'зрелость результата M0–M9, AHP с реальным расчётом CR (λmax, CI, RI Сати), экономика NPV/BCR/LCC и сравнение 4 альтернатив, '
         'авторекомендации по баллам ≤ 2, ранние сигналы, семантическая миграция старых ответов 31 → 30 (история сохранена) '
         'и импорт ответов из Excel-калькулятора № 9 (офлайн-заполнение).</p>')
s = s[:ps] + NEW_P + s[pe:]

# ============================================================ 5. хоткей заключения и niokrAPI
n3 = s.count('conclModal(cur())')
assert n3 >= 1, n3
s = s.replace('conclModal(cur())', 'concl18Modal(cur())')
rep('sfVal,ahpNum,LVLN,POROG,ALTNAMES4};',
    'sfVal,ahpNum,LVLN,POROG,ALTNAMES4,concl18Modal,exportMD18,renderAhp18,dashCard18,setQA18,qRow18,econCard18,expertProspect18};')

open(HTML, 'w', encoding='utf-8').write(s)
print(f'PATCH18-UI OK: {orig} → {len(s)} (+{len(s)-orig})')
