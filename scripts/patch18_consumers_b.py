# -*- coding: utf-8 -*-
"""
ЭТАП 18 (шаг 4, часть B): потребители v18 — respStats, setQA18/qRow18, renderResp,
discrepancies, scChip/avgScore/qList, fixAuditModal, MC/торнадо, сценарии,
renderEvid, propModal, дэш/дека/акт, buildSlides, KB, TOUR, лингвистическая замена.
Гард повторного запуска: 'function scenTargets('.
"""
import io, re, sys

MODEL = "output/Аудит_НИОКР_интерактивная_модель.html"
data = io.open(MODEL, encoding="utf-8").read()
if "function scenTargets(" in data:
    sys.exit("ЧАСТЬ B уже применена — выход без изменений.")
assert "const NQ30=" in data, "ЧАСТЬ A не применена"

mstart = data.find("<script>\n")
s = data[mstart + 9: data.rfind("</script>")]
nrep = [0]

def repl(old, new, name, cnt=1):
    global s
    c = s.count(old)
    assert c == cnt, "ANCHOR FAIL [%s]: expected %d, got %d :: %r" % (name, cnt, c, old[:80])
    s = s.replace(old, new)
    nrep[0] += 1

def cut(a, b):
    i = s.find(a)
    assert i >= 0, "cut-a FAIL " + a[:70]
    j = s.find(b, i + len(a))
    assert j >= 0, "cut-b FAIL " + b[:70]
    return s[i:j]

def replfn(start, name, newcode):
    global s
    i = s.find(start)
    assert i >= 0 and s.find(start, i + 1) < 0, "FN ANCHOR FAIL [%s]: %r" % (name, start[:70])
    j = s.find("\nfunction ", i + 20)
    for t in ("\nasync function ", "\n/* ="):
        k = s.find(t, i + 20)
        if 0 <= k and (j < 0 or k < j):
            j = k
    assert j > i, "FN END FAIL [%s]" % name
    s = s[:i] + newcode.rstrip("\n") + s[j:]
    nrep[0] += 1

# =====================================================================
# 1) calcCore — защита от устаревшего/отсутствующего r.ans (латентные пути)
# =====================================================================
repl("function calcCore(ans,Wv){", "function calcCore(ans,Wv){if(!ans||!ans.length)return{I:null,cat:null};", "calcCore-guard")

# =====================================================================
# 2) respStats — контракт {n,rows,cI,mode,dis}
# =====================================================================
replfn("function respStats(r){", "respStats", r"""function respStats(r){
 const P=(r.resp?r.resp.people:[]).filter(p=>r.resp.data[p.id]);
 const rows=P.map(p=>{const dd=r.resp.data[p.id];const c=calc(idxFake18(r,consensusTargets30(dd)))||{I:null};
  const f=personAns30(dd).filter(x=>x!=="").length;
  return{p,dd,nFilled:f,filled:f,I:c.I};});
 return{n:P.length,rows,cI:(calc(r)||{I:null}).I,mode:state.respMode||"median",dis:disagreementOf(r)};}""")

# =====================================================================
# 3) setQA18 — маршрутизация к активной персоне; qRow18 — id и 💬/📎
# =====================================================================
repl("function setQA18(code,field,val){\n  const r=cur();const lvl=r.lvl||'S';if(!r.q[lvl])r.q[lvl]={};\n  if(!r.q[lvl][code])r.q[lvl][code]={a:'',j:'',e:''};\n  r.q[lvl][code][field]=val;\n  save();",
     "function setQA18(code,field,val){\n  const r=cur();const lvl=r.lvl||'S';if(!r.q[lvl])r.q[lvl]={};\n  const ps=actResp(r);\n  if(ps&&lvl==='S'){if(!ps.qS)ps.qS={};if(!ps.qS[code])ps.qS[code]={a:'',j:'',e:''};ps.qS[code][field]=val;\n   syncConsensus(r);if(!r.q[lvl][code])r.q[lvl][code]={a:'',j:'',e:''};}\n  else{if(!r.q[lvl][code])r.q[lvl][code]={a:'',j:'',e:''};r.q[lvl][code][field]=val;}\n  save();",
     "setQA18-route")
repl("  return '<div class=\"qrow\"><div class=\"qh\"><div class=\"num\" style=\"font-size:11px;min-width:36px\">'+code+'</div>",
     "  const tpos=(lvl==='S')?q18pos(code):-1;\n  const tbtn=(lvl==='S'&&tpos>=0)?'<button class=\"btn sm gray\" data-thq18=\"'+tpos+'\" title=\"Обсуждение вопроса\">💬</button> <button class=\"btn sm gray\" data-doc18=\"'+tpos+'\" title=\"Реестр доказательств по вопросу\">📎</button>':'';\n  return '<div class=\"qrow\" id=\"qr'+lvl+'-'+code+'\"><div class=\"qh\"><div class=\"num\" style=\"font-size:11px;min-width:36px\">'+code+'</div>",
     "qrow-id")
repl("   '<div class=\"qb\"><label>Ответ: ", "   '<div class=\"qb\">'+tbtn+'<label>Ответ: ", "qrow-btn")
repl("  document.querySelectorAll('#tab-check input[data-qe]').forEach(el=>el.onchange=e=>setQA18(el.dataset.qe,'e',e.target.value));",
     "  document.querySelectorAll('#tab-check input[data-qe]').forEach(el=>el.onchange=e=>setQA18(el.dataset.qe,'e',e.target.value));\n"
     "  document.querySelectorAll('#tab-check [data-thq18]').forEach(el=>el.onclick=e=>threadsModal(+el.dataset.thq18));\n"
     "  document.querySelectorAll('#tab-check [data-doc18]').forEach(el=>el.onclick=e=>docsModal(+el.dataset.doc18));",
     "qrow-wire")

# =====================================================================
# 4) renderResp
# =====================================================================
repl("разногласий: <b>'+st.dis.length+'</b> из 31", "разногласий: <b>'+st.dis.length+'</b> из 30", "rr-dis31")
repl("'<td>'+x.filled+' / 31</td>", "'<td>'+x.filled+' / 30</td>", "rr-filled31")
repl("R.data[id1].ans=r.ans.slice();R.data[id1].conf=r.conf.slice();R.data[id1].com=r.com.slice();",
     "R.data[id1].qS=JSON.parse(JSON.stringify((r.q&&r.q.S)||{}));", "rr-seed")
repl("чек-лист показывает их ответы", "сетка «Средний» показывает их ответы", "rr-rsw")
repl("((r.resp.data[id]||{ans:[]}).ans||[]).filter(a=>a!==\"\").length",
     "personAns30(r.resp.data[id]||{}).filter(a=>a!==\"\").length", "rr-del")

# =====================================================================
# 5) discrepancies — кодовая сетка + v18extra + флаги
# =====================================================================
replfn("function discrepancies(r){", "discrepancies", r"""function discrepancies(r){
 const c=calc(r),ev=evidStats(r),out=[],A=ansS30(r);
 const ex=r.v18extra||{budget:"Б9",trl:"В1",imp:"Б10"};
 const a=code=>{const p=q18pos(code);return p>=0?(A[p]||""):"";};
 const dev=parseFloat(devPct(r));
 if((a("А1")==="0"||a("А1")==="1")&&r.status!=="Завершена")
  out.push(["err","Потребность (код А1) оценена в "+a("А1")+" балла — статус «"+r.status+"» под вопросом"]);
 if(r.status==="Выполняется"&&c.I!=null&&c.I<50)
  out.push(["err","Подразделение заявляет статус «Выполняется», при этом комплексный индекс I = "+c.I.toFixed(1)+" % (< 50)"]);
 if(r.decision&&/Продолжить без изменения/.test(r.decision)&&c.I!=null&&c.I<50)
  out.push(["warn","Предложение «"+r.decision+"» расходится с оценкой I = "+c.I.toFixed(1)+" %"]);
 if(isFinite(dev)&&dev>15&&["4","5"].indexOf(a(ex.budget))>=0)
  out.push(["err","Прогнозное превышение бюджета δб = "+dev.toFixed(1)+" % при высокой оценке по коду "+ex.budget]);
 else if(isFinite(dev)&&dev>15&&a(ex.budget)==="")
  out.push(["warn","Превышение бюджета δб = "+dev.toFixed(1)+" %, код "+ex.budget+" не оценён"]);
 const ready=parseFloat(r.ready);
 if(isFinite(ready)&&ready>=80&&["0","1","3"].indexOf(a(ex.imp))>=0)
  out.push(["warn","Заявлена готовность "+ready+" %, но возможности внедрения (код "+ex.imp+") оценены в "+a(ex.imp)]);
 if(ev.viol.length)
  out.push(["warn","Доказательный пробел: "+ev.viol.length+" высоких баллов (4–5) без документов — коды "+ev.viol.map(p=>CODES30[p]).join(", ")]);
 const fq=flagsOn(r);
 if(fq.length)out.push(["err","Контрольные замечания v18 ("+fq.length+"): "+fq.slice(0,3).join("; ")+(fq.length>3?" …":"")]);
 if(r.decision&&!(r.just||"").trim())out.push(["warn","Решение выбрано, но обоснование не заполнено"]);
 return out;}""")

# =====================================================================
# 6) avgScore / qList / scChip — LEG-совместимые обёртки над кодами
# =====================================================================
replfn("function avgScore(r,idxs){", "avgScore", r"""function avgScore(r,idxs){const A=ansS30(r);
 const v=idxs.map(i=>LEG2S31[i]>=0?A[LEG2S31[i]]:"").filter(a=>a&&a!=="НП"&&a!=="НД").map(Number).filter(isFinite);
 return v.length?v.reduce((s2,x)=>s2+x,0)/v.length:null;}""")
replfn("function qList(idxs){", "qList", r"""function qList(idxs){return idxs.map(i=>LEG2S31[i]>=0?CODES30[LEG2S31[i]]:String(i+1)).join(", ");}""")
replfn("function scChip(r,i){", "scChip", r"""function scChip(r,i){const p=LEG2S31[i];const a=p>=0?ansS30(r)[p]:"";const v=a===""?"—":a;
 const SC18={'5':'#2e7d32','4':'#7cb342','3':'#fbc02d','2':'#ef6c00','1':'#e65100','0':'#c62828','НД':'#8c8c8c','НП':'#9e9e9e'};
 return'<span class="pill" style="background:'+(SC18[String(v)]||"#9e9e9e")+'">'+(p>=0&&a!==""?CODES30[p]+" · ":"")+v+'</span>';}""")

# =====================================================================
# 7) fixAuditModal — восстановление из q/qS/легаси-снимка
# =====================================================================
repl("      const A2=ansArr(r),C2=confArr(r);\n      for(let j=0;j<NQ;j++){A2[j]=a.ans&&a.ans[j]!=null?a.ans[j]:\"\";C2[j]=a.conf&&a.conf[j]?a.conf[j]:\"\";}\n      if(respOn(r))syncConsensus(r);",
     "      if(a.q){r.q=JSON.parse(JSON.stringify(a.q));}\n"
     "      else if(a.qS){r.q=r.q||{K:{},S:{},P:{}};r.q.S=JSON.parse(JSON.stringify(a.qS));}\n"
     "      else if(a.ans&&a.ans.length===31){const S={};for(let j=0;j<31;j++){const p=LEG2S31[j];if(p<0)continue;const v=a.ans[j];if(v===\"\"||v==null)continue;let t=String(v);if(t===\"Н/П\")t=\"НП\";\n"
     "       if(!S[CODES30[p]]||qNum(t)>qNum(S[CODES30[p]].a))S[CODES30[p]]={a:t,j:\"\",e:\"\"};}\n"
     "       r.q=r.q||{K:{},S:{},P:{}};r.q.S=S;}\n      if(respOn(r))syncConsensus(r);",
     "hrest")

# =====================================================================
# 8) Монте-Карло / торнадо / применимость наборов весов
# =====================================================================
repl("function applyWeightSet(i){", "function iAtWeights(r,Wv){const old=state.weights;state.weights=Wv;try{const cc=calc(r);return cc?cc.I:null;}finally{state.weights=old;}}\nfunction applyWeightSet(i){", "iAtW-helper")
repl("const r=cur(),bI=calcCore(r.ans,W).I,nI=calc(r).I;", "const r=cur(),bI=iAtWeights(r,W),nI=calc(r).I;", "aws-bI")
repl("const baseI=calcCore(r.ans,W).I;", "const baseI=iAtWeights(r,W);", "aws-baseI")
replfn("function mcRun(r,N,seed){", "mcRun", r"""function mcRun(r,N,seed){
 const rnd=mulberry(seed);const A=ansS30(r);const idx=[];
 for(let i=0;i<NQ30;i++){const a=A[i];if(a!==""&&a!=="НП"&&a!=="НД"&&qNum(a)!=null)idx.push(i);}
 const out=[];
 for(let k=0;k<N;k++){const tg={};
  for(const i of idx){const u=rnd();let v=qNum(A[i]);
   if(u<0.15)v=Math.max(0,v-1);else if(u<0.30)v=Math.min(5,v+1);
   tg[CODES30[i]]=String(v);}
  const c2=calc(idxFake18(r,tg));out.push(c2?c2.I:null);}
 return out;}""")
oldpr=cut("const pr=[[85,","let rows")
repl(oldpr,'''const pr=[[85,"P(I ≥ 85) «Высокий»"],[70,"P(70–84) «Приемлемый»"],[55,"P(55–69) «Проблемный»"],[0,"P(I < 55) «Критический»"]];
   ''',"mc-pr")
repl('let rows="";const cnt=[0,0,0,0,0];','let rows="";const cnt=[0,0,0,0];',"mc-cnt0")
repl("res.forEach(v=>{cnt[v>=85?0:v>=70?1:v>=50?2:v>=30?3:4]++;});","res.forEach(v=>{cnt[v>=85?0:v>=70?1:v>=55?2:3]++;});","mc-cnt")
replfn("function tornado(r){", "tornado", r"""function tornado(r){
 const c=calc(r);if(c.I==null)return null;const x=idx18(r,"S");const A=ansS30(r);
 return[0,1,2].map(b=>{const N=x.nByDir[b]||0;
  const dPoint=N?100/(5*N):0;
  const tg={};let touched=0;
  CODES30.forEach((code,p)=>{const qd=Q18.levels.S.q[p];if(+qd.dir!==b)return;const v=qNum(A[p]);if(v==null||v>=5)return;
   tg[code]=String(v+1);touched++;});
  const c2=touched?(calc(idxFake18(r,tg))||c):c;
  return{b,dPoint,dW:(c2.I==null?0:c2.I-c.I)};});}""")
repl('const BNAME=["I. Актуальность","II. Эффективность","III. Целесообразность"];',
     'const BNAME=["А. Актуальность (I_А)","Б. Целесообразность (I_Ц)","В. Результативность (I_Р)"];', "BNAME")
repl("const BRNG=[[1,10],[11,19],[20,31]]; // номера отображения вопросов по блокам",
     "const BRNG=[[1,10],[11,20],[21,30]]; // коды А1–А10 · Б1–Б10 · В1–В10 уровня «Средний»", "BRNG")
repl("t.map(x=>'<tr><td>'+BNAME[x.b]+'</td><td><b>'+(x.dPoint>=0?\"+\":\"\")+x.dPoint.toFixed(2)+'</b></td><td><b>'+(x.dW>=0?\"+\":\"\")+x.dW.toFixed(2)+'</b></td></tr>')",
     "t.map(x=>'<tr><td>'+BNAME[x.b]+'</td><td><b>'+(x.dPoint>=0?\"+\":\"\")+x.dPoint.toFixed(2)+'</b> (1 балл одного вопроса)</td><td><b>'+(x.dW>=0?\"+\":\"\")+x.dW.toFixed(2)+'</b> (весь блок +1 балл)</td></tr>')",
     "tor-tab")

# =====================================================================
# 9) Сценарии — снимок qS + helper scenTargets + рендер таблицы
# =====================================================================
repl('r.scenarios.push({name:document.getElementById("scName").value||("Сценарий "+(r.scenarios.length+1)),date:new Date().toLocaleDateString("ru-RU"),ans:r.ans.slice(),conf:r.conf.slice(),ws:state.weights.slice(),wsName:state.wsName});',
     'r.scenarios.push({name:document.getElementById("scName").value||("Сценарий "+(r.scenarios.length+1)),date:new Date().toLocaleDateString("ru-RU"),lvl:r.lvl||"S",qS:JSON.parse(JSON.stringify((r.q&&r.q.S)||{})),ws:state.weights.slice(),wsName:state.wsName});',
     "sc-push")
i=s.find("function renderAnal(){")
assert i > 0, "renderAnal"
s = s[:i] + '''function scenTargets(s2){if(s2.qS&&Object.keys(s2.qS).length)return s2.qS;
 const tg={};if(s2.ans)for(let j=0;j<Math.min(s2.ans.length,31);j++){const p=LEG2S31[j];if(p<0)continue;let v=String(s2.ans[j]==null?"":s2.ans[j]);if(v==="")continue;if(v==="Н/П")v="НП";
  if(!tg[CODES30[p]]||qNum(v)>qNum(tg[CODES30[p]]))tg[CODES30[p]]=v;}
 return tg;}
''' + s[i:]
nrep[0] += 1
repl('    [{label:"Отмена",cls:"gray"},{label:"Применить",fn:()=>{r.ans=s2.ans.slice();r.conf=s2.conf.slice();\n',
     '    [{label:"Отмена",cls:"gray"},{label:"Применить",fn:()=>{const tg=scenTargets(s2);r.lvl="S";r.q=r.q||{K:{},S:{},P:{}};r.q.S=r.q.S||{};\n'
     '      Object.keys(tg).forEach(cd=>{const cl=r.q.S[cd]||{a:"",j:"",e:""};cl.a=String(tg[cd]);r.q.S[cd]=cl;});\n',
     "sc-apply")
repl("Текущие ответы чек-листа будут заменены", "Текущие ответы сетки «Средний» будут заменены", "sc-modal-txt")
old_rows=cut("    r.scenarios.map((s2,i)=>{const cc=calcCore(s2.ans,s2.ws||state.weights);","  }else sc+='<div class=\"mini\"")
repl(old_rows,"""    r.scenarios.map((s2,i)=>{const tg=scenTargets(s2);const cc=calc(idxFake18(r,tg))||{I:null,cat:""};const d=(cc.I==null?0:cc.I)-(c.I==null?0:c.I);
      return'<tr><td><b>'+s2.name+'</b><br><span class="mini">'+s2.date+'</span></td><td class="mini">'+Object.keys(tg).length+' ответов'+(s2.ws?(sameW(s2.ws,state.weights)?'':' · ⚠ прежняя схема весов'):'')+'</td><td>'+(cc.I==null?"Н/П":cc.I.toFixed(1))+'</td><td style="color:'+(d>=0?"#2e7d32":"#c62828")+'"><b>'+(d>=0?"+":"")+d.toFixed(1)+'</b></td><td><span class="pill" style="background:'+(COLCAT[cc.cat]||"#9e9e9e")+'">'+(cc.cat||"—")+'</span></td><td class="noprint"><button class="btn sm" data-scapp="'+i+'">⇐ Применить</button> <button class="btn sm warn" data-scdel="'+i+'">✕</button></td></tr>';}).join("")+'</table>';
""","sc-rows")
repl('notify("ok","Сценарий сохранён","ответов в сценарии: "+r.ans.filter(a=>a!=="").length);',
     'notify("ok","Сценарий сохранён","ответов сетки «Средний» в сценарии: "+Object.keys(scenTargets(r.scenarios[r.scenarios.length-1])).length);', "sc-notify")

# =====================================================================
# 10) renderEvid — кодовая сетка, переходы, новый документ по кодам
# =====================================================================
replfn("function renderEvid(){", "renderEvid", r"""function renderEvid(){
 const r=cur(),st=evidStats(r);
 if(!state.evidF)state.evidF={blk:"all",type:"",status:"",search:""};
 const f=state.evidF;
 const dirOf=p=>+((Q18.levels.S.q[p]||{}).dir||0);
 const idxDocs=(r.docs||[]).map((dc,gi)=>({dc,gi})).filter(x=>{const dc=x.dc;const p=docPos18(dc);
  if(f.blk==="gen"&&p!==-1)return false;
  if(f.blk==="0"||f.blk==="1"||f.blk==="2"){if(p<0||dirOf(p)!==+f.blk)return false;}
  if(f.type&&dc.type!==f.type)return false;
  if(f.status&&dc.status!==f.status)return false;
  if(f.search){const s2=f.search.toLowerCase();
   const hay=((p>=0?qLabel18(CODES30[p])+" ":"")+(dc.name||"")+" "+dc.type+" "+(dc.note||"")).toLowerCase();
   if(hay.indexOf(s2)<0)return false;}
  return true;}).sort((a,b)=>(docPos18(a.dc)-docPos18(b.dc))||((a.dc.q||0)-(b.dc.q||0)));
 const dCol=st.D==null?"#9e9e9e":st.D>=80?"#2e7d32":st.D>=50?"#fbc02d":"#c62828";
 const kCol=st.K==null?"#9e9e9e":st.K>=99.5?"#2e7d32":st.K>=70?"#fbc02d":"#c62828";
 const scoredProv=st.scored.filter(st.hasProv).length,highProv=st.high.filter(st.hasProv).length;
 let h='<div class="grid g4">'+
 '<div class="card kpi"><span class="l">ДОКАЗАТЕЛЬНАЯ ЗРЕЛОСТЬ D</span><span class="v" style="color:'+dCol+'">'+(st.D==null?"Н/П":st.D.toFixed(1)+" %")+'</span><span class="l">кодов с ответом и доказательством: '+scoredProv+' / '+st.scored.length+'</span></div>'+
 '<div class="card kpi"><span class="l">ПРАВИЛО «4–5 С ДОКУМЕНТОМ» K</span><span class="v" style="color:'+kCol+'">'+(st.K==null?"Н/П":st.K.toFixed(1)+" %")+'</span><span class="l">высоких баллов обеспечено: '+highProv+' / '+st.high.length+'</span></div>'+
 '<div class="card kpi"><span class="l">ДОКУМЕНТОВ В РЕЕСТРЕ</span><span class="v">'+st.total+'</span><span class="l">предоставлено '+st.prov+' · заявлено '+st.decl+' · запрошено '+st.req+' · отсутствует '+st.abs+'</span></div>'+
 '<div class="card kpi"><span class="l">НАРУШЕНИЯ (4–5 БЕЗ ДОКУМЕНТА)</span><span class="v" style="color:'+(st.viol.length?"#c62828":"#2e7d32")+'">'+st.viol.length+'</span><span class="l">по кодам сетки «Средний»</span></div></div>';
 const NM3=["А · Актуальность (I_А)","Б · Целесообразность (I_Ц)","В · Результативность (I_Р)"];
 const blkRows=[0,1,2].map(b=>{const sb=st.scored.filter(i=>dirOf(i)===b),hb=st.high.filter(i=>dirOf(i)===b);
  const sp=sb.filter(st.hasProv).length,hp=hb.filter(st.hasProv).length;
  return'<tr><td>'+NM3[b]+'</td><td>'+sb.length+'</td><td>'+sp+'</td><td>'+(sb.length?(100*sp/sb.length).toFixed(0)+" %":"—")+'</td><td>'+hb.length+'</td><td>'+hp+'</td><td>'+(hb.length?(100*hp/hb.length).toFixed(0)+" %":"—")+'</td></tr>';}).join("");
 h+='<div class="grid g2" style="margin-top:14px"><div class="card"><h3>Зрелость доказательств по направлениям (Средний, 30 кодов)</h3><table><tr><th>Направление</th><th>Оценено</th><th>С док.</th><th>D</th><th>Баллы 4–5</th><th>С док.</th><th>K</th></tr>'+blkRows+'</table><div class="mini">Доказательством считается предоставленный документ в реестре или заполненное поле «Доказательство» вопроса.</div></div>';
 h+='<div class="card"><h3>⚠️ Высокие баллы без доказательств</h3>'+(st.viol.length?'<div class="feed">'+st.viol.map(i=>'<div class="fi warn">Код <b>'+CODES30[i]+'</b> — балл '+ansS30(r)[i]+': '+String(Q18.levels.S.q[i].q).slice(0,64)+' <button class="btn sm gray" data-goq="'+i+'">→ к вопросу</button> <button class="btn sm" data-dm="'+i+'">📎 документы</button></div>').join("")+'</div>':'<div class="mini">Нарушений правила «4–5 только с документом» нет.</div>')+'</div></div>';
 h+='<div class="card" style="margin-top:14px"><h3>Реестр документов</h3>'+
 '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:8px" class="noprint">'+
 '<label style="font-size:12px">Направление: <select id="evBlk"><option value="all"'+(f.blk==="all"?" selected":"")+'>все</option><option value="0"'+(f.blk==="0"?" selected":"")+'>А · Актуальность</option><option value="1"'+(f.blk==="1"?" selected":"")+'>Б · Целесообразность</option><option value="2"'+(f.blk==="2"?" selected":"")+'>В · Результативность</option><option value="gen"'+(f.blk==="gen"?" selected":"")+'>общие</option></select></label>'+
 '<label style="font-size:12px">Тип: <select id="evType"><option value="">все</option>'+DOCTYPES.map(t=>'<option'+(f.type===t?" selected":"")+'>'+t+'</option>').join("")+'</select></label>'+
 '<label style="font-size:12px">Статус: <select id="evStatus"><option value="">все</option>'+DOCSTATUS.map(s2=>'<option value="'+s2[0]+'"'+(f.status===s2[0]?" selected":"")+'>'+s2[1]+'</option>').join("")+'</select></label>'+
 '<input type="text" id="evSearch" placeholder="🔎 поиск по реестру…" value="'+(f.search||"").replace(/"/g,"&quot;")+'" style="flex:1;min-width:150px"></div>'+
 (idxDocs.length?'<table><tr><th>Код · вопрос</th><th>Наименование документа</th><th>Тип</th><th>Дата</th><th>Статус</th><th>Примечание</th><th></th></tr>'+
  idxDocs.map(x=>{const dc=x.dc;const p=docPos18(dc);
   return'<tr><td>'+(p>=0?'<b>'+CODES30[p]+'</b> <button class="btn sm gray" data-goq="'+p+'" title="К вопросу опросника">→</button><br><span class="mini">'+String(Q18.levels.S.q[p].q).slice(0,54)+'</span>':'<span class="mini">общий</span>')+'</td>'+
   '<td>'+(dc.name||'<span class="mini">— не заполнено —</span>')+'</td><td style="font-size:12px">'+dc.type+'</td><td>'+fmtDate(dc.date)+'</td>'+
   '<td><span class="pill" style="background:'+DSCOL[dc.status]+'">'+docStatusLbl(dc.status)+'</span></td><td class="mini">'+(dc.note||"")+'</td>'+
   '<td class="noprint"><button class="btn sm warn" data-edel="'+x.gi+'">✕</button></td></tr>';}).join("")+'</table>'
  :'<div class="mini">Документов по текущему фильтру нет.</div>')+
 '<div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap" class="noprint"><select id="evNewQ" style="min-width:280px"><option value="-1">Общий документ (вся НИОКР)</option>'+CODES30.map((cd,j)=>'<option value="'+j+'">'+cd+": "+String(Q18.levels.S.q[j].q).slice(0,50)+'</option>').join("")+'</select><button class="btn sm" id="evAdd">＋ Добавить документ</button> <button class="btn gray sm" id="evCsv">Экспорт CSV</button> <button class="btn gray sm" id="evPrint">🖨 Печать реестра</button></div></div>';
 document.getElementById("tab-evid").innerHTML=h;
 document.getElementById("evBlk").onchange=e=>{f.blk=e.target.value;renderEvid();};
 document.getElementById("evType").onchange=e=>{f.type=e.target.value;renderEvid();};
 document.getElementById("evStatus").onchange=e=>{f.status=e.target.value;renderEvid();};
 document.getElementById("evSearch").oninput=e=>{f.search=e.target.value;state.evidFocus=true;renderEvid();};
 if(state.evidFocus){const s3=document.getElementById("evSearch");s3.focus();s3.setSelectionRange(s3.value.length,s3.value.length);state.evidFocus=false;}
 document.querySelectorAll("#tab-evid [data-goq]").forEach(b=>b.onclick=e=>gotoQ(+e.target.closest("[data-goq]").dataset.goq));
 document.querySelectorAll("#tab-evid [data-dm]").forEach(b=>b.onclick=e=>docsModal(+e.target.closest("[data-dm]").dataset.dm));
 document.querySelectorAll("#tab-evid [data-edel]").forEach(b=>b.onclick=e=>{const gi=+e.target.dataset.edel;const dc=r.docs[gi];const p=docPos18(dc);
  showModal('<h2>Удалить документ из реестра?</h2><p><b>'+(dc.name||"(без наименования)")+'</b> — '+(p>=0?"код "+CODES30[p]:"общий документ")+'. Ответы вопросов останутся.</p>',
  [{label:"Отмена",cls:"gray"},{label:"Удалить",cls:"warn",fn:()=>{r.docs.splice(gi,1);save();renderEvid();renderCheck();
   notify("warn","Документ удалён из реестра",p>=0?"код "+CODES30[p]:"общий");}}]);});
 document.getElementById("evAdd").onclick=()=>{const qv=parseInt(document.getElementById("evNewQ").value);
  r.docs.push(qv>=0?{q:qv,q18:CODES30[qv],name:"",type:DOCTYPES[0],date:"",status:"declared",note:""}:{q:-1,q18:null,name:"",type:DOCTYPES[0],date:"",status:"declared",note:""});
  save();renderEvid();renderCheck();
  notify("ok","Документ добавлен в реестр",(qv>=0?"код "+CODES30[qv]:"общий документ")+" — заполните наименование и статус");};
 document.getElementById("evCsv").onclick=()=>{download("Реестр_доказательств_"+r.shifr+".csv",buildDocsCsv(r),"text/csv");
  notify("ok","Реестр доказательств экспортирован","CSV: "+st.total+" док., D = "+(st.D==null?"Н/П":st.D.toFixed(1)+"%"));};
 document.getElementById("evPrint").onclick=printEvidAkt;
 if(!canPass()){
  document.querySelectorAll("#tab-evid select,#tab-evid input,#tab-evid button[data-edel],#tab-evid button[data-dm],#tab-evid #evAdd").forEach(el=>{el.disabled=true;});
  document.getElementById("tab-evid").insertAdjacentHTML("afterbegin",'<div class="note">🔒 Роль «Наблюдатель»: реестр доказательств доступен только для чтения.</div>');
 }
 updateEvidBadge();
}""")

# =====================================================================
# 11) propModal — цели по кодам
# =====================================================================
repl("'<select id=\"ptQ\" style=\"max-width:340px\">'+QS.map((qq,j)=>'<option value=\"'+j+'\">№'+(j+1)+\". \"+qq[0].slice(0,48)+(qq[0].length>48?\"…\":\"\")+'</option>').join(\"\")+'</select>'+",
     "'<select id=\"ptQ\" style=\"max-width:340px\">'+CODES30.map((cd,j)=>'<option value=\"'+cd+'\">'+cd+\" · \"+String(Q18.levels.S.q[j].q).slice(0,44)+(String(Q18.levels.S.q[j].q).length>44?\"…\":\"\")+'</option>').join(\"\")+'</select>'+",
     "pm-ptQ")
repl("g(\"ptWeak\").onclick=()=>{const a=ansArr(r);let n=0;for(let j=0;j<NQ;j++){if(a[j]!==\"\"&&a[j]!==\"НП\"&&Number(a[j])<=2){p.targets[String(j)]=\"4\";n++;}}",
     "g(\"ptWeak\").onclick=()=>{const a=ansS30(r);let n=0;for(let j=0;j<NQ30;j++){if(a[j]!==\"\"&&a[j]!==\"НП\"&&a[j]!==\"НД\"&&Number(a[j])<=2){p.targets[CODES30[j]]=\"4\";n++;}}",
     "pm-ptWeak")
repl("Object.keys(p.targets).map(q=>'<tr><td><b>'+(Number(q)+1)+'</b> <span class=\"mini\">'+QS[q][0].slice(0,52)+(QS[q][0].length>52?\"…\":\"\")+'</span></td><td style=\"text-align:center\">'+(ansArr(r)[q]===\"\"?\"—\":ansArr(r)[q]===\"НП\"?\"Н/П\":ansArr(r)[q])+'</td>",
     "Object.keys(p.targets).map(q=>'<tr><td><b>'+q+'</b> <span class=\"mini\">'+qLabel18(q).slice(0,60)+'</span></td><td style=\"text-align:center\">'+(ansArr(r)[q18pos(q)]===\"\"?\"—\":ansArr(r)[q18pos(q)])+'</td>",
     "pm-ptList")
repl("Целевые баллы по вопросам (прогноз I)","Целевые баллы по кодам сетки «Средний» (прогноз I)","pm-hint")

# =====================================================================
# 12) База/дэш/дека/акт — строки и счётчики
# =====================================================================
repl('+["Высокая","Хорошая","Удовлетворительная","Неудовлетворительная","Критическая","Не оценена"].map(s=>\'<option\'',
     "+PLVLO.concat(\"Не оценена\").map(s=>'<option'", "db-cat")
repl("const ansN=r.ans.filter(a=>a!==\"\").length;","const ansN=ansS30(r).filter(a=>a!==\"\").length;","deck-ansN")
repl("Ответов чек-листа: <b>'+ansN+' / 31</b>","Ответов опросника (Средний): <b>'+ansN+' / 30</b>","deck-ans31")
repl("разногласий: '+st.dis.length+' из 31 · личные I","разногласий: '+st.dis.length+' из 30 · личные I","dash-dis")
repl("+x.filled+\"/31, личный I = \"", "+x.filled+\"/30, личный I = \"", "act-filled31")
repl("Вопросов с разногласиями: '+st.dis.length+' из 31 (разобраны","Вопросов с разногласиями: '+st.dis.length+' из 30 (разобраны","act-dis31")

# =====================================================================
# 12b) легаси-потребители c.bl / c.nCtrl / BN (дэш совета, отчёты)
# =====================================================================
repl("c.early=early18(r,x,st);c.blocks=x.blocks;", "c.early=early18(r,x,st);c.blocks=x.blocks;c.nCtrl=(r.ctrl||[]).map(String).filter(v=>v===\"Выявлено\").length;", "nCtrl")
repl('BN=["Актуальность","Эффективность","Целесообразность"];', 'BN=["Актуальность","Целесообразность","Результативность"];', "da-BN")
old_da=cut("[0,1,2].forEach(b=>{const P=c.bl[b].P;","if(c.I!=null&&c.I<50)")
repl(old_da,"[0,1,2].forEach(b=>{const dv=(c.idx&&c.idx.dirs[b])!=null?c.idx.dirs[b]:null;if(dv!=null&&dv<50)out.push(\"Низкий индекс направления «\"+BN[b]+\"» (\"+dv.toFixed(0)+\" %) — подготовить мероприятия по кодам направления с баллами 0–2.\");});\n","da-bl")
repl('ev.viol.length?\" по вопросам \"+ev.viol.map(i=>i+1).join(", ")','ev.viol.length?\" по кодам \"+ev.viol.map(i=>CODES30[i]).join(", ")',"da-viol")
old_bd=cut("Выполнение блоков: P₁ =",".</p>'+")
repl(old_bd,"Индексы направлений: I_А = '+((c.idx&&c.idx.dirs[0])!=null?c.idx.dirs[0].toFixed(1)+\" %\":\"Н/П\")+', I_Ц = '+((c.idx&&c.idx.dirs[1])!=null?c.idx.dirs[1].toFixed(1)+\" %\":\"Н/П\")+', I_Р = '+((c.idx&&c.idx.dirs[2])!=null?c.idx.dirs[2].toFixed(1)+\" %\":\"Н/П\")+'","bd-p")
repl("'<tr><td>Актуальность (B₁)</td><td>'+c.bl[0].B.toFixed(2)+' / '+fmtW1(state.weights[0])+' (P₁='+(c.bl[0].P==null?\"Н/П\":(c.bl[0].P*100).toFixed(1)+\"%\")+')</td></tr>'+",
     "'<tr><td>Актуальность (I_А)</td><td>'+((c.idx&&c.idx.dirs[0])!=null?c.idx.dirs[0].toFixed(1)+\" %\":\"Н/П\")+' · коды А1–А10</td></tr>'+", "rep-r1")
repl("'<tr><td>Эффективность (B₂)</td><td>'+c.bl[1].B.toFixed(2)+' / '+fmtW1(state.weights[1])+' (P₂='+(c.bl[1].P==null?\"Н/П\":(c.bl[1].P*100).toFixed(1)+\"%\")+')</td></tr>'+",
     "'<tr><td>Целесообразность (I_Ц)</td><td>'+((c.idx&&c.idx.dirs[1])!=null?c.idx.dirs[1].toFixed(1)+\" %\":\"Н/П\")+' · коды Б1–Б10</td></tr>'+", "rep-r2")
repl("'<tr><td>Целесообразность (B₃)</td><td>'+c.bl[2].B.toFixed(2)+' / '+fmtW1(state.weights[2])+' (P₃='+(c.bl[2].P==null?\"Н/П\":(c.bl[2].P*100).toFixed(1)+\"%\")+')</td></tr>'+",
     "'<tr><td>Результативность (I_Р)</td><td>'+((c.idx&&c.idx.dirs[2])!=null?c.idx.dirs[2].toFixed(1)+\" %\":\"Н/П\")+' · коды В1–В10</td></tr>'+", "rep-r3")
repl('(c.renorm?" (применена ренормировка весов)":"")','', "rep-renorm")

# =====================================================================
# 13) buildSlides
# =====================================================================
repl("const ctrlHit=CTRL.map((cc,i)=>r.ctrl[i]===\"Выявлено\"?cc:null).filter(Boolean);",
     "const ctrlHit=flagsOn(r);","bs-ctrl")
repl('ctrlHit.map(x=>"• "+x).join("<br>")','ctrlHit.map(x=>"• "+String(x).slice(0,120)).join("<br>")',"bs-ctrlmap")
old_s4=cut("'<div><h3>Баллы чек-листа (блок II, вес '+fmtW1(w[1])+' %)</h3><div class=\"kv\">№13 бюджет: ", "'<div class=\"band q\"><b>Вопросы комиссии:")
repl(old_s4,"""'<div><h3>Экономика аудита v18.0 + финансовые коды (Средний)</h3><div class="kv">NPV: <b>'+(((c.econ||{}).npv)==null?"—":(c.econ||{}).npv+" млн ₽")+'</b> · BCR: <b>'+(((c.econ||{}).bcr)==null?"—":(c.econ||{}).bcr)+'</b>'+(((c.econ||{}).bcrTxt)?' <span class="mini">'+(c.econ||{}).bcrTxt+'</span>':'')+
 '<br>LCC (PV): '+(((c.econ||{}).lccPv)==null?"—":(c.econ||{}).lccPv+" млн ₽")+
 '<br>Затраты и ресурсы: '+scChip(r,12)+' '+scChip(r,13)+'<br>Прозрачность затрат и остатки закупок: '+scChip(r,16)+' '+scChip(r,17)+'<br>Потребность до завершения: '+scChip(r,18)+'</div>'+
  ""","bs-slide4")
old9=cut("S.push('<h2>9. Интегральная оценка состояния НИОКР (по чек-листу)</h2>'"," S.push('<h2>10. Предложения")
repl(old9,"""S.push('<h2>9. Интегральная оценка состояния НИОКР (модель v18.0)</h2>'+(respOn(r)?'<div class="kv" style="margin-bottom:8px">👥 Итоговые ответы — консенсус '+r.resp.people.length+' респондентов ('+(state.respMode==="mean"?"среднее":"медиана")+'); разногласий: '+disagreementOf(r).length+' из 30 кодов</div>':'')+'<table><tr><th>Направление (сетка «Средний»)</th><th>Индекс, %</th><th>Уровень</th><th>Coverage / порог</th></tr>'+
  [0,1,2].map(b=>{const dv=c.idx?c.idx.dirs[b]:null;const nm=dv==null?null:plvl(dv);
   return'<tr><td>'+BNAME[b]+'</td><td><b>'+(dv==null?"н/д":dv.toFixed(1))+'</b></td><td><span class="pill" style="background:'+(PLVLC[nm]||"#9e9e9e")+'">'+(nm||"—")+'</span></td><td>'+(b===0?(c.cov==null?"—":c.cov.toFixed(1)+" % / ≥ "+c.thr+"%"):"")+'</td></tr>';}).join("")+
  '<tr><td><b>Комплексный I = ∛(I_А·I_Ц·I_Р)</b></td><td><b>'+I+'</b></td><td><span class="pill" style="background:'+(COLCAT[c.cat]||"#9e9e9e")+'">'+(c.cat||"—")+'</span></td><td class="mini">'+(c.covOk===true?"порог пройден":(c.covOk===false?"ПОРОГ НЕ ПРОЙДЕН":"н/д"))+'</td></tr></table>'+
  '<div style="text-align:center;margin:10px 0"><span class="badge">'+I+' / 100 — КОМПЛЕКСНЫЙ ИНДЕКС</span> <span class="pill" style="background:'+(COLCAT[c.cat]||"#9e9e9e")+'">'+(c.cat||"недостаточно данных")+'</span></div>'+
  (disc.length?'<div class="band disc"><b>Лента расхождений (подразделение vs аудитор):</b><br>'+disc.map(x2=>"• "+x2[1]).join("<br>")+'</div>':"")+
  '<div class="band rule">Индексы направлений и комплексный I — инструмент поддержки решения; итоговое решение принимает комиссия с учётом стоп-факторов SF-0/1/2 и Coverage. '+
  'Рекомендация: '+String(c.rec||(expertRec(r)||{}).text||"—").slice(0,160)+'.</div>');
""","bs-slide9")
repl("/ 100 — интегральная оценка (по чек-листу аудита)","/ 100 — комплексная оценка (опросник v18.0)","bs-title")
repl("<th>Вопросы чек-листа</th>","<th>Коды опросника</th>","bs-crit-th")

# =====================================================================
# 14) expertDashCard — цвета PLVLC
# =====================================================================
repl("(c.lA?LVL5COL[c.lA]","(c.lA?PLVLC[c.lA]","edc-la")
repl("(c.lC?LVL5COL[c.lC]","(c.lC?PLVLC[c.lC]","edc-lc")

# =====================================================================
# 15) KB: таблица якорей и источники
# =====================================================================
old_anch=cut("<h3>Якорные описания баллов по всем 31 вопросу</h3>","+'</table>'+")
repl(old_anch,"""<h3>Якорные описания баллов (опросник «Средний», 30 кодов)</h3><table><tr><th>Код</th><th>Вопрос</th><th>5 баллов</th><th>3 балла</th><th>0 баллов</th></tr>'+
   CODES30.map((cd,j)=>{const qq=Q18.levels.S.q[j];const an=qq.anch||{};return'<tr><td><b>'+cd+'</b></td><td class="mini">'+String(qq.q).slice(0,74)+'</td><td class="mini">'+String(an["5"]||"—")+'</td><td class="mini">'+String(an["3"]||an["4"]||"—")+'</td><td class="mini">'+String(an["0"]||an["1"]||"—")+'</td></tr>';}).join("")""","kb-anch")
old_src=cut("'8. Внутренние документы методики: чек-лист", ".</div>'+")
repl(old_src,"'8. Внутренние документы методики: опросник аудита НИОКР v18.0 (уровни 15/30/60), регламент аудита, форма «Экспертное заключение Ф-06», Excel-калькулятор «9_Калькулятор_аудита_НИОКР_v2.xlsx».</div>","kb-src")

# =====================================================================
# 16) TOUR — полное переписывание блока
# =====================================================================
i=s.find("const TOUR=[")
assert i >= 0, "TOUR anchor"
j = s.find("];\n", i)
assert j > i, "TOUR end"
s = s[:i] + '''const TOUR=[
["dash","Дашборд","Комплексный индекс I","I = ∛(I_А·I_Ц·I_Р) по сетке «Средний» (30 кодов: А1–А10 · Б1–Б10 · В1–В10). Карточка перспектив M0–M9, экономика NPV/BCR/LCC (Excel-калькулятор №9) и контрольные замечания qFlags."],
["check","Опросник 15/30/60","Три уровня опросника","Переключатель уровней К(15) / С(30) / П(60). Балл 4–5 — только с доказательством; «НП» — с письменным обоснованием (уменьшает Coverage), «НД» — только из числителя Coverage. Пороги Coverage: 80 / 85 / 90 %."],
["check","Опросник 15/30/60","💬 и 📎 у каждого кода","У вопросов уровня «Средний» — кнопки 💬 (обсуждение) и 📎 (реестр доказательств). При включённых респондентах сетка показывает ответы активной персоны, консенсус формируется автоматически."],
["evid","Реестр доказательств","D и K по кодам","D — доля оцененных кодов с доказательствами; K — доля баллов 4–5, обеспеченных документом. Пробелы подписаны кодами (А/Б/В) с переходом к вопросу."],
["anal","Аналитика","Сценарии и вероятность","Сценарии сохраняют снимок ответов сетки «Средний» (qS); Монте-Карло и торнадо считаются по индексам v18. Полосы вероятности: ≥85 «Высокий», 70–84 «Приемлемый», 55–69 «Проблемный», <55 «Критический»."],
["wgt","Веса (AHP)","AHP по направлениям","Парные сравнения Саати для трёх направлений (I_А / I_Ц / I_Р) с контролем согласованности CR. 21 пара и веса импортируются из Excel-калькулятора №9."]
];
''' + s[j+2:]
nrep[0] += 1

# =====================================================================
# 17) niokrAPI
# =====================================================================
repl("window.niokrAPI={get state(){return state;},cur,calcCore,",
     "window.niokrAPI={get state(){return state;},cur,calcCore,calc,plvl,ansS30,CODES30,scenTargets,iAtWeights,", "api")

# =====================================================================
# 18) exportChecklistMD — мост легаси-экспорта на кодовую сетку
# =====================================================================
repl("function exportChecklistMD(r){",
     "function exportChecklistMD(r){if(!r.ans||r.ans.length!==31){const _A=new Array(31).fill(\"\");const _S=ansS30(r);for(let j=0;j<31;j++){const p=LEG2S31[j];if(p>=0&&_S[p]!==\"\")_A[j]=String(_S[p]);}r=Object.assign({},r,{ans:_A});}",
     "md-bridge")

# =====================================================================
# 19) Лингвистическая замена чек-лист → опросник (с защитой исторических сносок)
# =====================================================================
PROTECT = ["чек-листа 31", "с чек-листа v17"]
ph = {}
for k2, pt in enumerate(PROTECT):
    tok = "@@PROT%d@@" % k2
    ph[tok] = pt
    s = s.replace(pt, tok)

def word_sub(text, a, b):
    return re.sub("(?<![\\w-])" + re.escape(a) + "(?![\\w-])", b, text)

PAIRS = [("чек-листов", "опросников"), ("чек-листа", "опросника"), ("чек-листу", "опроснику"),
         ("чек-листом", "опросником"), ("чек-лист", "опросник"),
         ("Чек-листов", "Опросников"), ("Чек-листа", "Опросника"), ("Чек-листу", "Опроснику"),
         ("Чек-листом", "Опросником"), ("Чек-листы", "Опросники"), ("Чек-лист", "Опросник")]
for a, b in PAIRS:
    s = word_sub(s, a, b)
for tok, pt in ph.items():
    s = s.replace(tok, pt)

# =====================================================================
# Финальные проверки и запись
# =====================================================================
assert "function scenTargets(" in s, "scenTargets lost"
assert "function iAtWeights(" in s, "iAtWeights lost"
assert s.count("QDOC[") == 0, "QDOC consumers remain"
for tok in ("c.bl[", "c.renorm", "c.Wsum"):
    assert s.count(tok) == 0, "legacy token remains: %s (%d)" % (tok, s.count(tok))
assert "@@PROT" not in s, "protect tokens leaked"
out = data[:mstart + 9] + s + data[data.rfind("</script>"):]
io.open(MODEL, "w", encoding="utf-8").write(out)
io.open("/tmp/model18_partB.js", "w", encoding="utf-8").write(s)
print("Замен в PART B:", nrep[0], "| размер модели:", len(out.encode("utf-8")))
print("Остаточные 'чек-лист':", len(re.findall("[Чч]ек-лист", s)))
for mm in re.finditer("[Чч]ек-лист", s):
    print("  ·", s[max(0, mm.start()-40):mm.end()+30].replace("\n", " "))
