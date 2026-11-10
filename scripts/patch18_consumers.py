# -*- coding: utf-8 -*-
"""
ЭТАП 18Б, ШАГ 4 (часть A): переключение legacy-потребителей модели на
каноническую сетку уровня S (30 вопросов, коды А1..А10, Б1..Б10, В1..В10).
Идемпотентность: повторный запуск прерывается по якорю `const NQ30=`.
"""
import io, re, sys

MODEL = "output/Аудит_НИОКР_интерактивная_модель.html"
data = io.open(MODEL, encoding="utf-8").read()
if "const NQ30=" in data:
    sys.exit("ШАГ 4 (часть A) уже применён — повторный запуск запрещён.")

m = re.search(r"<script>\n(.*)</script>", data, re.S)
assert m, "script block not found"
s = m.group(1)
nrep = [0]

def repl(old, new, name, cnt=1):
    global s
    c = s.count(old)
    assert c == cnt, ("ANCHOR FAIL [%s]: expected %d, got %d :: %r" % (name, cnt, c, old[:80]))
    s = s.replace(old, new)
    nrep[0] += 1

def replfn(start, name, newcode):
    """Полная замена функции: от `start` до следующей строки 'function '/'async function '/`/* ==`."""
    global s
    i = s.find(start)
    assert i >= 0 and s.find(start, i + 1) < 0, ("FN ANCHOR FAIL [%s]: %r" % (name, start[:70]))
    j = s.find("\nfunction ", i + 20)
    for t in ("\nasync function ", "\n/* ="):
        k = s.find(t, i + 20)
        if 0 <= k and (j < 0 or k < j):
            j = k
    assert j > i, ("FN END FAIL [%s]" % name)
    old = s[i:j]
    if old.count("{") != old.count("}"):
        print("WARN: brace skew in old [%s] (likely string literals) — proceeding" % name)
    s = s[:i] + newcode.rstrip("\n") + s[j:]
    nrep[0] += 1

# =====================================================================
# 0) ХЕЛПЕРЫ (вставляются перед function evidStats(r){)
# =====================================================================
HELP = r"""
/* ===== v18.0 шаг 4: каноническая сетка потребителей = уровень S (30 вопросов) ===== */
const NQ30=30;
const CODES30=(Q18.levels.S.q||[]).map(function(q){return String(q.code);});
const Q30IDX={};CODES30.forEach(function(c,p){Q30IDX[c]=p;});
function q18pos(code){var p=Q30IDX[String(code)];return p==null?-1:p;}
const LEG2S31=(function(){var v=Array(31).fill(-1);(Q18.mig31||[]).forEach(function(pr){var p=q18pos(pr[1]);if(p>=0)v[pr[0]]=p;});return v;})();
const PLVLO=["Высокий","Приемлемый","Проблемный","Критический"];
function plvlIdx(v){var x=plvl(v);return x?PLVLO.indexOf(x):-1;}
Object.assign(COLCAT,{"Высокий":"#2e7d32","Приемлемый":"#7cb342","Проблемный":"#ef6c00","Критический":"#c62828"});
function lvOf(src){return (src&&src.lvl)||"S";}
function vecFrom(src,field){var v=new Array(NQ30).fill("");
 var g=(src.q&&src.q.S)||null;
 if(g)CODES30.forEach(function(c,pp){var cl=g[c];if(cl&&cl[field]!=null&&String(cl[field])!=="")v[pp]=String(cl[field]);});
 if(field==="a"&&src.ans&&src.ans.length&&!g){for(var j2=0;j2<Math.min(src.ans.length,31);j2++){var t=src.ans[j2];if(t===""||t==null)continue;var q=LEG2S31[j2];if(q<0||v[q]!=="")continue;var t2=String(t);if(t2==="Н/П")t2="НП";v[q]=t2;}}
 return v;}
function ansS30(r){return vecFrom(r,"a");}
function jS30(r){return vecFrom(r,"j");}
function eS30(r){return vecFrom(r,"e");}
function personVec(dd,field){if(!dd)return new Array(NQ30).fill("");
 var g=(dd.q&&dd.q.S)||dd.qS||null;
 if(g){var v=new Array(NQ30).fill("");CODES30.forEach(function(c,p){var cl=g[c];if(cl&&cl[field]!=null&&String(cl[field])!=="")v[p]=String(cl[field]);});return v;}
 return vecFrom({ans:dd.ans||[],conf:dd.conf||[],com:dd.com||[]},field);}
function personAns30(dd){return personVec(dd,"a");}
function personJ30(dd){return personVec(dd,"j");}
function personE30(dd){return personVec(dd,"e");}
function personCellMap30(dd){var v=new Array(NQ30).fill(null);if(!dd)return v;
 var g=(dd.q&&dd.q.S)||dd.qS||null;
 if(g){CODES30.forEach(function(c,p){var cl=g[c];if(cl)v[p]={a:String(cl.a==null?"":cl.a),j:String(cl.j==null?"":cl.j),e:String(cl.e==null?"":cl.e)};});return v;}
 if(dd.ans&&dd.ans.length){for(var k=0;k<Math.min(dd.ans.length,31);k++){var p=LEG2S31[k];if(p<0||v[p])continue;var a3=String(dd.ans[k]==null?"":dd.ans[k]);
  if(a3==="Н/П")a3="НП";
  v[p]={a:a3,j:"",e:String((dd.com&&dd.com[k])||"")};}
  return v;}
 return v;}
function consensusTargets30(dd){var t={};personAns30(dd).forEach(function(v,p){if(v!=="")t[CODES30[p]]=v;});return t;}
function qLabel18(code){var p=q18pos(code);var qd=p>=0?(Q18.levels.S.q[p]||{}):{};
 return String(code)+(qd.q?" — "+qd.q:"");}
function docPos18(dc){if(!dc)return -1;
 if(dc.q18!=null){var p=(typeof dc.q18==="number"&&dc.q18>=0&&dc.q18<NQ30)?dc.q18:q18pos(dc.q18);if(p>=0)return p;}
 if(typeof dc.q==="number"){var p2=q18pos(dc.q);if(p2>=0)return p2;
  if(dc.q>=0&&dc.q<31&&LEG2S31[dc.q]>=0)return LEG2S31[dc.q];}
 return -1;}
function idxFake18(r,targets){var rr=JSON.parse(JSON.stringify(r||{}));rr.lvl="S";rr.q=rr.q||{K:{},S:{},P:{}};rr.q.S=rr.q.S||{};
 Object.keys(targets||{}).forEach(function(code){var p=q18pos(code);if(p<0)return;var cl=rr.q.S[code]||{a:"",j:"",e:""};cl.a=String(targets[code]);rr.q.S[code]=cl;});
 return rr;}
function fakeR18(r,targets){return idxFake18(r,targets);}
const RECMX4=[
 [["R1 · Продолжить без существенных изменений — работа ведётся корректно, требования к доказательности соблюдайте",0],
  ["R2 · Продолжить с корректировкой по результативности — уточните ТЗ, план этапов и подтверждающие испытания",1],
  ["R3 · Пересмотреть модель реализации — актуальность высока, результативность под вопросом; сравните альтернативы",2],
  ["R6 · Приостановить до восстановления доказательной базы по направлению «Результативность»",3]],
 [["R2 · Продолжить с корректировкой по направлению «Актуальность» — обновите подтверждение потребности и анализ аналогов",1],
  ["R3 · Корректировка: доработка ТЗ, уточнение планов этапов и механизмов внедрения, усиление доказательности",1],
  ["R4 · Пересмотр НИОКР — сравнение альтернатив и оценка целесообразности продолжения",2],
  ["R6 · Приостановить: требуется отдельный план восстановления направления «Результативность»",3]],
 [["R2 · Продолжить с корректировкой — акцент на укрепление целесообразности реализации",1],
  ["R4 · Пересмотр: результативность подтверждена, обоснование актуальности и целесообразности недостаточно",2],
  ["R5 · Пересмотреть цели и границы НИОКР; при отсутствии подтверждения — готовить прекращение",3],
  ["R7 · Готовить прекращение — индексы актуальности и результативности в критической зоне",4]],
 [["R4 · Пересмотреть актуальность — без подтверждения потребности дальнейшие затраты нецелесообразны",2],
  ["R5 · Приостановить с внеочередным рассмотрением руководством (актуальность критическая)",3],
  ["R6 · Приостановить; при неустранении факторов — прекращение",3],
  ["R7 · Прекратить работы — индексы актуальности и целесообразности критические, SF-факторы доминируют",4]]];
const RECGCOL=["#2e7d32","#7cb342","#f9a825","#ef6c00","#c62828"];
const RECGNAME=["Продолжение","Корректировка","Пересмотр","Приостановка","Прекращение"];
function recGroup4(g){return RECGCOL[g]==null?"#9e9e9e":RECGCOL[g];}
function migrate18extra(r){var done=false;
 if(!r.v18extra){var T=Q18.levels.S.q,L="";
  var BUD=L,TRL=L,IMP=L;T.forEach(function(q){var t=(String(q.q)+" "+String(q.check||"")+" "+String(q.ev||"")).toLowerCase();
   if(!BUD&&/бюджет|финансир|затрат|смет/.test(t))BUD=String(q.code);
   if(!TRL&&/trl|готовност|зрелост/.test(t))TRL=String(q.code);
   if(!IMP&&/внедрен|освоени|производствен|эксплуатац/.test(t))IMP=String(q.code);});
  (r.docs||[]).forEach(function(d){if(typeof d.q==="number"&&d.q>=0&&d.q<31&&d.q18==null){var h=(Q18.mig31||[]).filter(function(pr){return pr[0]===d.q;})[0];if(h)d.q18=h[1];}});
  var nt={};Object.keys(r.threads||{}).forEach(function(k){var p=+k;
   if(isFinite(p)&&p>=0&&p<31&&LEG2S31[p]>=0){var nk=String(LEG2S31[p]);if(nt[nk]&&nt[nk].concat)nt[nk]=nt[nk].concat(r.threads[k]);else nt[nk]=r.threads[k];}
   else nt[String(k)]=r.threads[k];});
  r.threads=nt;
  (r.scenarios||[]).forEach(function(sc){var qd=null;
   if(sc.q18!=null){var p=q18pos(sc.q18);if(p>=0)qd=T[p];}
   if(!qd&&typeof sc.q==="number"&&sc.q>=0&&sc.q<30)qd=T[sc.q];
   if(!qd&&typeof sc.q==="number"&&sc.q>=30&&sc.q<31){var j=LEG2S31[sc.q];if(j>=0)qd=T[j];}
   if(!qd&&sc.code!=null){var p2=q18pos(sc.code);if(p2>=0)qd=T[p2];}
   if(!qd)qd=T[0];
   sc.qS=String(qd.code);sc.q18=String(qd.code);sc.q=Q30IDX[String(qd.code)];});
  r.v18extra={ok:true,budget:BUD||"Б9",trl:TRL||"В1",imp:IMP||"Б10"};done=true;}
 return done;}
"""
i0 = s.find("\nfunction evidStats(r){")
assert i0 > 0, "evidStats anchor missing"
s = s[:i0] + "\n" + HELP + s[i0:]
nrep[0] += 1

# =====================================================================
# 1) accessors → сетка S
# =====================================================================
repl('function ansArr(r){const a=actResp(r);return a?a.ans:r.ans;}\n'
     'function confArr(r){const a=actResp(r);return a?a.conf:r.conf;}\n'
     'function comArr(r){const a=actResp(r);return a?a.com:r.com;}',
     'function ansArr(r){const a=actResp(r);return a?personAns30(a):ansS30(r);}\n'
     'function confArr(r){const a=actResp(r);return a?personVec(a,"j"):jS30(r);}\n'
     'function comArr(r){const a=actResp(r);return a?personVec(a,"e"):eS30(r);}',
     "accessors")

# =====================================================================
# 2) syncConsensus — по кодам S из экспертных сеток (dd.qS)
# =====================================================================
replfn("function syncConsensus(r){", "syncConsensus", r"""function syncConsensus(r){if(!respOn(r))return;
 const P=r.resp.people.filter(p=>r.resp.data[p.id]);if(!P.length)return;
 const wts=P.map(p=>{const w=r.resp.data[p.id].w;return (w==null||!isFinite(+w))?1:+w;});
 r.q=r.q||{K:{},S:{},P:{}};if(!r.q.S)r.q.S={};const g=r.q.S;
 CODES30.forEach(code=>{const vals=[],js=[],es=[];
  P.forEach(p=>{const dd=r.resp.data[p.id];const cl=(dd.qS&&dd.qS[code])||null;
   if(cl){if(String(cl.a||"")!=="")vals.push(String(cl.a));if(String(cl.j||"").trim())js.push(String(cl.j).trim());if(String(cl.e||"").trim())es.push(String(cl.e).trim());}
   else{const aj=personAns30(dd)[q18pos(code)];if(aj!==""){vals.push(aj);const ej=personE30(dd)[q18pos(code)];if(ej)es.push(ej);}}});
  if(!vals.length)return;
  const old=g[code]||{a:"",j:"",e:""};const cv=deriveConsensus18(vals,wts);
  g[code]={a:cv,j:js.length?js[0]:old.j,e:es.length?es.join("; "):old.e};});}""")

# =====================================================================
# 3) disagreementOf / respStats / buildRespCsv
# =====================================================================
replfn("function disagreementOf(r,all){", "disagreementOf", r"""function disagreementOf(r,all){if(!respOn(r))return[];
 const P=r.resp.people.filter(p=>r.resp.data[p.id]),out=[];
 if(!P.length)return out;
 CODES30.forEach((code,p)=>{const V=P.map(pt=>{const A=personAns30(r.resp.data[pt.id]);return{name:pt.name,color:pt.color,a:A[p]===""||A[p]==null?"":A[p]};});
  if(!all&&!V.some(v=>v.a!==""))return;
  const ns=V.map(v=>v.a).filter(a=>a!==""&&a!=="НП"&&a!=="НД").map(Number).filter(isFinite);
  const spread=ns.length?Math.max.apply(null,ns)-Math.min.apply(null,ns):0;
  if(!all&&spread<2)return;
  const md=ns.length?med(ns):null,m2=ns.length?ns.reduce((a,b)=>a+b,0)/ns.length:null;
  const g=(r.q&&r.q.S&&r.q.S[code])||null;
  out.push({code,i:p,spread:spread,vals:V,min:ns.length?Math.min.apply(null,ns):null,max:ns.length?Math.max.apply(null,ns):null,med:md,mean:m2,cons:g?String(g.a||""):""});});
 if(!all)out.sort((x,y)=>y.spread-x.spread||q18pos(x.code)-q18pos(y.code));return out;}""")

replfn("function respStats(r){", "respStats", r"""function respStats(r){
 const P=(r.resp?r.resp.people:[]).filter(p=>r.resp.data[p.id]);
 return P.map(p=>{const dd=r.resp.data[p.id];const fake=idxFake18(r,consensusTargets30(dd));const c=calc(fake);
  return{p,dd,nFilled:personAns30(dd).filter(x=>x!=="").length,I:c?c.I:null};});}""")

replfn("function buildRespCsv(r){", "buildRespCsv", r"""function buildRespCsv(r){
 const P=(r.resp?r.resp.people:[]).filter(p=>r.resp.data[p.id]);
 const head=["Шифр","Код","Вопрос"].concat(P.map(p=>p.name)).join(";");
 let csv="Респонденты — сетка уровня «Средний» (шкала v18)\n"+head+"\n";
 CODES30.forEach((code,p)=>{const qd=Q18.levels.S.q[p];
  csv+=[r.shifr,code,String(qd.q).replace(/;/g,",")].concat(P.map(pt=>{const A=personAns30(r.resp.data[pt.id]);return A[p]===""?"—":A[p];})).join(";")+"\n";});
 csv+="\nРеспондент;Заполнено из 30;I (по сетке «Средний»)\n";
 respStats(r).forEach(x=>{csv+=[x.p.name,x.nFilled,x.I==null?"Н/П":x.I.toFixed(1)].join(";")+"\n";});
 return csv;}""")

# =====================================================================
# 4) makeAudit — снимок v18
# =====================================================================
replfn("function makeAudit(r,note){", "makeAudit", r"""function makeAudit(r,note){
 const c=calc(r),ev=evidStats(r),x=idx18(r,lvOf(r)),sf=stops18(r);
 return{id:"a"+Date.now().toString(36)+Math.floor(Math.random()*1000),date:new Date().toISOString(),dateStr:new Date().toLocaleString("ru-RU"),
  I:c.I,cat:c.cat||"",lvl:r.lvl||"S",IA:x.IA,IC:x.IC,IR:x.IR,cov:x.cov==null?null:Math.round(x.cov*10)/10,thr:x.thr,
  min:x.min==null?null:Math.round(x.min*10)/10,sf2:sf.n2>0,
  ws:fmtW(state.weights),wsName:state.wsName,form:r.form||"",decision:r.decision||"",status:r.status||"",
  by:whoAmI(),note:note||"",
  q:JSON.parse(JSON.stringify(r.q||{K:{},S:{},P:{}})),
  qS:JSON.parse(JSON.stringify((r.q&&r.q.S)||{})),
  ans:ansArr(r).slice(),conf:r.conf?r.conf.slice():[],
  D:ev.D,K:ev.K,docsN:ev.total,
  resp:respOn(r)?{mode:state.respMode||"median",people:r.resp.people.map(p=>{const dd=r.resp.data[p.id]||{};const fake=idxFake18(r,consensusTargets30(dd));const cc=calc(fake);
   return{name:p.name,title:p.title||"",filled:personAns30(dd).filter(a=>a!=="").length,I:cc?cc.I:null};})}:null,
  flow:flowOf(r)?{stage:r.flow.stage,done:!!r.flow.done}:null};}""")

# =====================================================================
# 5) cmpModal — коды S, старые снимки через LEG2S31
# =====================================================================
replfn("function cmpModal(r,k1,k2){", "cmpModal", r"""function cmpModal(r,k1,k2){
 const a=r.audits[Math.min(k1,k2)],b=r.audits[Math.max(k1,k2)];
 function auditVec(a2){const v=new Array(NQ30).fill("");
  if(a2.qS)CODES30.forEach((c,p)=>{const cl=a2.qS[c];if(cl&&cl.a!=null)v[p]=String(cl.a);});
  else if(a2.ans&&a2.ans.length===31)for(let j=0;j<31;j++){const p=LEG2S31[j];if(p<0||v[p]!=="")continue;let t=String(a2.ans[j]==null?"":a2.ans[j]);if(t==="")continue;if(t==="Н/П")t="НП";v[p]=t;}
  return v;}
 const av=auditVec(a),bv=auditVec(b),rows=[];
 for(let p=0;p<NQ30;p++){const x=av[p],y=bv[p];if(x!==y)rows.push('<tr><td><b>'+CODES30[p]+'</b> <span class="mini">'+String(Q18.levels.S.q[p].q).slice(0,58)+'</span></td><td>'+(x||"—")+'</td><td>'+(y||"—")+'</td><td>'+((qNum(x)!=null&&qNum(y)!=null)?(((qNum(x)-qNum(y))>0?"+":"")+(qNum(x)-qNum(y))):"—")+'</td></tr>');}
 showModal('<h2>⇄ Сравнение аудитов</h2><p class="mini"><b>A:</b> '+a.dateStr+' · '+a.by+' · I='+(a.I==null?"Н/П":a.I.toFixed(1))+' ('+(a.cat||"—")+')'+(a.lvl?" · "+LVLN[a.lvl]:"")+(a.cov!=null?" · Coverage "+a.cov+"%":"")+'<br><b>B:</b> '+b.dateStr+' · '+b.by+' · I='+(b.I==null?"Н/П":b.I.toFixed(1))+' ('+(b.cat||"—")+')'+(b.lvl?" · "+LVLN[b.lvl]:"")+(b.cov!=null?" · Coverage "+b.cov+"%":"")+
  '<br>ΔI = '+((a.I==null||b.I==null)?"н/д":((b.I-a.I>0?"+":"")+(b.I-a.I).toFixed(1)))+' п.п.</p>'+
  (rows.length?'<table><tr><th>Код · вопрос</th><th>A</th><th>B</th><th>ΔA−B</th></tr>'+rows.join("")+'</table>':'<p class="mini">Ответы полностью совпадают.</p>'),
  [{label:"Закрыть",cls:"gray"}]);}""")

# =====================================================================
# 6) stopsOn / flagsOn
# =====================================================================
repl('function stopsOn(r){return(r.stops||[]).map((v,i)=>v==="ДА"?STOPS[i]:null).filter(Boolean);}',
     'function stopsOn(r){try{const st18=stops18(r);return (st18&&st18.list)||[];}catch(e){return[];}}',
     "stopsOn")
replfn("function flagsOn(r){", "flagsOn", r"""function flagsOn(r){const out=[];const g=(r.q&&r.q.S)||{};
 CODES30.forEach(code=>{const cl=g[code];if(!cl)return;const fl=qFlags(cl);if(fl)out.push(code+": "+fl);});
 return out;}""")

# =====================================================================
# 7) expertRec / expertProspect / expertConclusion / conclModal
# =====================================================================
replfn("function expertRec(r){", "expertRec", r"""function expertRec(r){const st=stopsOn(r);const x=idx18(r,"S");
 if(st.length)return{text:"Внеочередное рассмотрение руководством (сработал стоп-фактор SF-2) — вплоть до приостановки работ",g:3,stop:true,st,mx:null};
 if(x.IA==null||x.IC==null)return{text:"Недостаточно данных для экспертной оценки (индексы I_А и/или I_Ц по сетке «Средний» не рассчитаны)",stop:false,st,mx:null,g:null};
 const la=plvlIdx(x.IA),lc=plvlIdx(x.IC);
 return{text:RECMX4[la][lc][0],g:RECMX4[la][lc][1],stop:false,st,mx:[la,lc]};}""")
replfn("function expertProspect(r){", "expertProspect", r"""function expertProspect(r){return expertProspect18(r);}""")
replfn("function expertConclusion(r){", "expertConclusion", r"""function expertConclusion(r){const x=idx18(r,"S"),rec=expertRec(r),ev=evidStats(r),dc=discrepancies(r);
 const problems=[];if(ev.viol.length)problems.push("оценки ≥4 без документов: "+ev.viol.map(p=>CODES30[p]).join(", "));
 if(dc&&dc.length)problems.push("расхождения с данными паспорта/финансов: "+dc.length);
 if(x.ctrl)problems.push("контрольных замечаний: "+x.ctrl+(x.low2?" · балл ≤2: "+x.low2:""));
 if(!problems.length&&!rec.stop)problems.push("критичных замечаний не выявлено");
 return{iA:x.IA,lA:plvl(x.IA)||"—",iC:x.IC,lC:plvl(x.IC)||"—",iR:x.IR,lR:plvl(x.IR)||"—",cov:x.cov,
  problems,prospect:expertProspect(r),rec:rec.text,stop:rec.stop,stops:rec.st,flags:flagsOn(r)};}""")
replfn("function conclModal(r){", "conclModal", r"""function conclModal(r){const ec=expertConclusion(r);
 showModal('<h2>📋 Экспертное заключение (5 обязательных позиций)</h2><div style="font-size:14px;line-height:1.6">'+
  "<p><b>1. Уровень актуальности:</b> "+(ec.iA==null?"не оценён (недостаточно данных)":ec.iA.toFixed(1)+" % — «"+ec.lA+"»")+".</p>"+
  "<p><b>2. Уровень целесообразности:</b> "+(ec.iC==null?"не оценён (недостаточно данных)":ec.iC.toFixed(1)+" % — «"+ec.lC+"»")+".</p>"+
  "<p><b>3. Проблемы и материалы, требующие уточнения:</b> "+ec.problems.join("; ")+".</p>"+
  "<p><b>4. Перспективы завершения работ:</b> "+ec.prospect+".</p>"+
  "<p><b>5. Рекомендация:</b> "+(ec.stop?"⛔ ":"")+ec.rec+".</p>"+
  "<p class=\"mini\">Итоги сетки «Средний»: I_Р = "+(ec.iR==null?"н/д":ec.iR.toFixed(1)+" %")+" · Coverage = "+(ec.cov==null?"—":ec.cov.toFixed(1)+" %")+".</p>"+
  "</div>",[{label:"Закрыть",cls:"gray"}]);}""")

# =====================================================================
# 8) exportChecklistMD → exportMD18
# =====================================================================
replfn("function exportChecklistMD(r){", "exportChecklistMD", r"""function exportChecklistMD(r){return exportMD18(r);}""")

# =====================================================================
# 9) evidStats — сетка S + документы по docPos18
# =====================================================================
replfn("function evidStats(r){", "evidStats", r"""function evidStats(r){
 const docs=r.docs||[];
 const provByPos={};docs.forEach(d=>{if(d.status==="provided"){const p=docPos18(d);if(p>=0)provByPos[p]=true;}});
 const A=ansS30(r),E=eS30(r);
 const scored=[],high=[],viol=[];
 for(let i=0;i<NQ30;i++){const a=A[i];if(a===""||a==null||a==="НП")continue;scored.push(i);
  const n=qNum(a);if(n!=null&&n>=4){high.push(i);if(!provByPos[i]&&!String(E[i]||"").trim())viol.push(i);}}
 const has=i=>!!provByPos[i]||String(E[i]||"").trim()!=="";
 const D=scored.length?100*scored.filter(has).length/scored.length:null;
 const K=high.length?100*high.filter(has).length/high.length:null;
 return{docs,hasProv:i=>!!provByPos[i],scored,high,viol,D,K,total:docs.length,
  prov:docs.filter(d=>d.status==="provided").length,
  decl:docs.filter(d=>d.status==="declared").length,
  req:docs.filter(d=>d.status==="requested").length,
  abs:docs.filter(d=>d.status==="absent").length};}""")

# =====================================================================
# 10) syncConfFromDocs / docsModal / gotoQ / buildDocsCsv / printEvidAkt
# =====================================================================
replfn("function syncConfFromDocs(r,i){", "syncConfFromDocs", r"""function syncConfFromDocs(r,i){
 const hasProv=(r.docs||[]).some(d=>docPos18(d)===i&&d.status==="provided");
 if(!hasProv)return;
 r.q=r.q||{K:{},S:{},P:{}};if(!r.q.S)r.q.S={};
 const cl=r.q.S[CODES30[i]]||{a:"",j:"",e:""};
 if(!String(cl.e||"").trim()){const dp=(r.docs||[]).filter(d=>docPos18(d)===i&&d.status==="provided")[0];
  cl.e="Документ из реестра: "+((dp&&dp.name)||"без наименования");r.q.S[CODES30[i]]=cl;
  if(respOn(r))syncConsensus(r);notify("info","Доказательство обновлено автоматически",CODES30[i]+": в реестре появился предоставленный документ");}}""")

replfn("function docsModal(i){", "docsModal", r"""function docsModal(i){const r=cur();if(!r.docs)r.docs=[];
 const pos=(typeof i==="string"?q18pos(i):i);const code=CODES30[pos]||CODES30[0];
 const canEdit=canPass();
 const docs=r.docs.filter(d=>docPos18(d)===pos);
 const av=ansS30(r)[pos],highScore=(qNum(av)!=null&&qNum(av)>=4),hasProv=docs.some(d=>d.status==="provided")||String(eS30(r)[pos]||"").trim()!=="";
 let h='<h2>📎 Реестр доказательств — вопрос '+code+'</h2><p class="mini">'+qLabel18(code)+'</p>';
 if(highScore&&!hasProv)h+='<div class="note">⚠️ Высокий балл ('+av+') без предоставленного документа — нарушение правила «оценки 4–5 только с документом».</div>';
 if(highScore&&hasProv)h+='<div class="note" style="background:#f2f8ee;border-color:#7cb342">✓ Правило «высокая оценка — с доказательством» соблюдено.</div>';
 h+=docs.length?'<table><tr><th>Наименование и реквизиты</th><th>Тип</th><th>Дата</th><th>Статус</th><th>Примечание</th><th></th></tr>'+
  docs.map(dc=>{const gi=r.docs.indexOf(dc);
   return'<tr><td><input type="text" data-df="name" data-dg="'+gi+'" value="'+(dc.name||"").replace(/"/g,"&quot;")+'" style="width:100%" placeholder="наименование, №, дата"></td>'+
   '<td><select data-df="type" data-dg="'+gi+'">'+DOCTYPES.map(t=>'<option'+(dc.type===t?" selected":"")+'>'+t+'</option>').join("")+'</select></td>'+
   '<td><input type="date" data-df="date" data-dg="'+gi+'" value="'+(dc.date||"")+'"></td>'+
   '<td><select data-df="status" data-dg="'+gi+'">'+DOCSTATUS.map(s=>'<option value="'+s[0]+'"'+(dc.status===s[0]?" selected":"")+'>'+s[1]+'</option>').join("")+'</select></td>'+
   '<td><input type="text" data-df="note" data-dg="'+gi+'" value="'+(dc.note||"").replace(/"/g,"&quot;")+'" style="width:100%" placeholder="комментарий"></td>'+
   '<td><button class="btn sm warn" data-ddel="'+gi+'">✕</button></td></tr>';}).join("")+'</table>'
  :'<div class="mini">Документов по вопросу пока нет — добавьте первый.</div>';
 h+='<div style="margin-top:10px" class="noprint"><button class="btn sm" id="docAdd">＋ Добавить документ</button></div>';
 showModal(h,
  [{label:"Закрыть",cls:"gray",fn:()=>{syncConfFromDocs(r,pos);save();renderCheck();renderEvid();}},
   {label:"Готово",fn:()=>{syncConfFromDocs(r,pos);save();renderCheck();renderEvid();renderDash();
    notify("ok","Реестр доказательств обновлён",code+": документов "+r.docs.filter(d=>docPos18(d)===pos).length);}}]);
 if(!canEdit)document.querySelectorAll("#mbox input,#mbox select,#mbox button[data-ddel],#mbox #docAdd").forEach(el=>el.disabled=true);
 const add=document.getElementById("docAdd");
 if(add)add.onclick=()=>{r.docs.push({q:pos,q18:code,name:"",type:DOCTYPES[0],date:"",status:"declared",note:""});save();docsModal(pos);};
 document.querySelectorAll("#mbox [data-df]").forEach(el=>el.onchange=e=>{const gi=+e.target.dataset.dg,f=e.target.dataset.df;
  r.docs[gi][f]=e.target.value;
  logEvent(r,"реестр: документ «"+(r.docs[gi].name||"без наименования")+"» (вопрос "+code+") "+(f==="status"?"статус → "+docStatusLbl(e.target.value):f));
  save();
  if(f==="status"&&e.target.value==="provided"){syncConfFromDocs(r,docPos18(r.docs[gi]));save();
   notify("ok","Документ предоставлен",code+": "+(r.docs[gi].name||"без наименования"));docsModal(pos);}});
 document.querySelectorAll("#mbox [data-ddel]").forEach(el=>el.onclick=e=>{const gi=+e.target.dataset.ddel;
  r.docs.splice(gi,1);save();notify("warn","Документ удалён из реестра",code);docsModal(pos);});}""")

replfn("function gotoQ(i){", "gotoQ", r"""function gotoQ(i){goTab("check");const r=cur();const p=(typeof i==="string"?q18pos(i):i);
 setTimeout(()=>{if(r.lvl!=="S"&&!(r.q&&r.q.S&&Object.keys(r.q.S).length)){notify("warn","Откройте уровень «Средний»","вопрос "+CODES30[p]+" из сетки «Средний»");return;}
  const el=document.getElementById("qrS-"+CODES30[p]);
  if(!el){notify("warn","Вопрос недоступен на текущем уровне","код "+CODES30[p]);return;}
  el.scrollIntoView({block:"center"});el.style.outline="3px solid #fbc02d";setTimeout(()=>{el.style.outline="";},1800);},60);}""")

repl("(r.docs||[]).slice().sort((a,b)=>a.q-b.q)","(r.docs||[]).slice().sort((a,b)=>(docPos18(a)-docPos18(b))||((a.q||0)-(b.q||0)))","docs-sort",2)
repl('dc.q>=0?dc.q+1:""','docPos18(dc)>=0?CODES30[docPos18(dc)]:""',"csvQ")
repl('dc.q>=0?QS[dc.q][0]:"Общий по НИОКР"','docPos18(dc)>=0?qLabel18(CODES30[docPos18(dc)]):"Общий по НИОКР"',"csvLbl")
repl('Нарушения (вопросы);"+(st.viol.map(i=>i+1)','Нарушения (коды);"+(st.viol.map(i=>CODES30[i]',"csv-viol")

replfn("function printEvidAkt(){", "printEvidAkt", r"""function printEvidAkt(){
 const r=cur(),st=evidStats(r);
 const rows=(r.docs||[]).slice().sort((a,b)=>(docPos18(a)-docPos18(b))||((a.q||0)-(b.q||0))).map(dc=>'<tr><td>'+(docPos18(dc)>=0?CODES30[docPos18(dc)]:"общий")+'</td><td>'+(docPos18(dc)>=0?String(qLabel18(CODES30[docPos18(dc)])).slice(0,90):"—")+'</td><td>'+(dc.name||"—")+'</td><td>'+dc.type+'</td><td>'+fmtDate(dc.date)+'</td><td>'+docStatusLbl(dc.status)+'</td><td class="mini">'+(dc.note||"")+'</td></tr>').join("");
 const h='<div class="card"><h3>РЕЕСТР ДОКАЗАТЕЛЬСТВ — НИОКР «'+r.name+'» ('+r.shifr+')</h3>'+
 '<p>Документов в реестре: <b>'+st.total+'</b> (предоставлено: '+st.prov+'). Показатель доказательной зрелости <b>D = '+(st.D==null?"Н/П":st.D.toFixed(1)+" %")+'</b>, соблюдения правила «высокая оценка — с документом» <b>K = '+(st.K==null?"Н/П":st.K.toFixed(1)+" %")+'</b>.</p>'+
 '<table><tr><th>Код</th><th>Вопрос</th><th>Документ</th><th>Тип</th><th>Дата</th><th>Статус</th><th>Примечание</th></tr>'+rows+'</table>'+
 '<p class="mini">Сформировано интерактивной моделью аудита НИОКР v18.0. Аудитор: ____________ Представитель подразделения: ____________</p></div>';
 document.getElementById("printarea").innerHTML=h;
 window.print();}""")

# =====================================================================
# 11) expertCard
# =====================================================================
replfn("function expertCard(r){", "expertCard", r"""function expertCard(r){const x=idx18(r,"S"),rec=expertRec(r),pr=expertProspect(r);
 const chip=v=>v==null?'<span class="mini">н/д</span>':'<b>'+v.toFixed(1)+'%</b> <span style="font-weight:700;color:'+(PLVLC[plvl(v)]||"#8c8c8c")+'">'+plvl(v)+'</span>';
 let h='<div class="card"><h3>🧠 Экспертное заключение по методике (сетка «Средний», 30 вопросов)</h3>'+
 '<div class="grid g3"><div><span class="l">I_А актуальность</span>'+chip(x.IA)+'</div>'+
 '<div><span class="l">I_Ц целесообразность</span>'+chip(x.IC)+'</div>'+
 '<div><span class="l">I_Р результативность</span>'+chip(x.IR)+'</div></div>';
 if(rec.mx!=null)h+='<div class="mini" style="margin:6px 0">Матрица решений: строка I_А «'+PLVLO[rec.mx[0]]+'» × столбец I_Ц «'+PLVLO[rec.mx[1]]+'».</div>';
 h+='<div class="note" style="border-left:5px solid '+(rec.stop?"#c62828":recGroup4(rec.g))+'"><b>'+(rec.stop?"⛔ ":"")+rec.rec+'</b></div>';
 const ev=evidStats(r),sf=stops18(r);
 h+='<div class="mini" style="margin-top:6px">Перспективы завершения: <b>'+pr+'</b> · оценок ≥4 без документов: <b>'+ev.viol.length+'</b>'+(ev.viol.length?" ("+ev.viol.map(p=>CODES30[p]).join(", ")+")":"")+' · Coverage: <b>'+(x.cov==null?"—":x.cov.toFixed(1)+" %")+'</b> (порог ≥ '+x.thr+'%).</div>';
 h+='<div class="mini" style="margin-top:4px">Стоп-факторы SF-2'+(sf.list.length?": <b style=\"color:#c62828\">сработали — "+sf.list.join("; ")+"</b>":" — не сработали")+'.</div></div>';
 return h;}""")

# =====================================================================
# 12) threadsModal / disTable
# =====================================================================
replfn("function threadsModal(i){", "threadsModal", r"""function threadsModal(i){const r=cur();const pos=(typeof i==="string"?q18pos(i):i);
 const key=String(pos),lbl=(pos>=0&&pos<NQ30)?"код "+CODES30[pos]:"вопрос "+key;
 if(!r.threads)r.threads={};const th=r.threads[key]||(r.threads[key]=[]);
 const canW=state.role!=="viewer";
 const QD=(pos>=0&&pos<NQ30)?Q18.levels.S.q[pos]:null;
 let h='<h2>💬 Обсуждение — '+lbl+'</h2>'+(QD?'<p class="mini">'+String(QD.q)+'</p>':'');
 h+=th.length?'<div style="max-height:320px;overflow:auto">'+th.map(m=>'<div style="border-bottom:1px solid #eee;padding:5px 0"><b>'+m.by+'</b> <span class="mini">'+new Date(m.t).toLocaleString("ru-RU")+'</span><div>'+(m.txt||"").replace(/</g,"&lt;")+'</div></div>').join("")+'</div>':'<p class="mini">Сообщений пока нет.</p>';
 if(canW)h+='<textarea id="thTxt" placeholder="Ваш комментарий…" style="width:100%;min-height:70px;margin-top:8px"></textarea>';
 showModal(h,canW?[{label:"Добавить сообщение",fn:()=>{const v=(document.getElementById("thTxt")||{}).value||"";
  if(!v.trim())return notify("warn","Пустое сообщение","введите текст");
  th.push({by:whoAmI(),t:new Date().toISOString(),txt:v.trim()});save();
  logEvent(r,"тред "+lbl+": добавлено сообщение");notify("ok","Комментарий добавлен",lbl);}}]:[{label:"Закрыть",cls:"gray"}]);}""")

replfn("function disTable(r,dis){", "disTable", r"""function disTable(r,dis){
 return '<table><tr><th>Код · вопрос</th>'+r.resp.people.map(p=>'<th style="color:'+p.color+'">'+p.name+'</th>').join("")+'<th>Медиана</th><th>Среднее</th><th>Разброс</th><th>Консенсус</th></tr>'+
 (dis.length?dis.map(x=>'<tr><td><b>'+x.code+'</b> <span class="mini">'+String(Q18.levels.S.q[x.i].q).slice(0,58)+(String(Q18.levels.S.q[x.i].q).length>58?"…":"")+'</span></td>'+
  x.vals.map(v=>'<td style="text-align:center"><b>'+v.a+'</b></td>').join("")+
  '<td style="text-align:center">'+(x.med==null?"—":x.med)+'</td><td style="text-align:center">'+(x.mean==null?"—":x.mean.toFixed(2))+'</td>'+
  '<td style="text-align:center"><b>'+x.spread+'</b></td><td style="text-align:center">'+(x.cons===""?'<span class="mini">—</span>':'<b>'+x.cons+'</b>')+'</td></tr>').join("")
  :'<tr><td colspan="8" class="mini">Значимых разногласий (разброс ≥2 баллов) нет.</td></tr>')+'</table>';}""")

# =====================================================================
# 13) propAns / propForecast / propStats / buildPropsCsv
# =====================================================================
replfn("function propAns(r,p){", "propAns", r"""function propAns(r,p){const a=ansS30(r).slice();Object.keys(p.targets||{}).forEach(k=>{const j=q18pos(k);
 if(j>=0&&j<NQ30)a[j]=String(p.targets[k]);});return a;}""")
replfn("function propForecast(r,p){", "propForecast", r"""function propForecast(r,p){
 const base=calc(r)||{I:null,cat:""};
 const fake=idxFake18(r,p.targets||{});const c2=calc(fake)||{I:null,cat:""};
 const dI=(base.I!=null&&c2.I!=null)?c2.I-base.I:null;
 const extra=parseFloat(String(p.extra||"").replace(",","."));const ex=isFinite(extra)&&String(p.extra||"").trim()!==""?extra:null;
 const months=parseFloat(String(p.months||"").replace(",","."));
 return{I0:base.I,cat0:base.cat||"",col0:COLCAT[base.cat]||"#9e9e9e",I1:c2.I,cat1:c2.cat||"",col1:COLCAT[c2.cat]||"#9e9e9e",dI,
  extra:ex,months:isFinite(months)&&String(p.months||"").trim()!==""?months:null,
  cpp:(ex!=null&&dI!=null&&dI>0)?ex/dI:null,nTargets:Object.keys(p.targets||{}).length};}""")
replfn("function propStats(){", "propStats", r"""function propStats(){const acc=[];state.list.forEach(r=>(r.props||[]).forEach(p=>{if(p.status!=="accepted")return;
 const f=propForecast(r,p);acc.push({r,p,f});}));
 if(!acc.length)return null;
 const sumExtra=acc.reduce((s,x)=>s+(x.f.extra||0),0);
 const dis=acc.filter(x=>x.f.dI!=null);
 return{n:acc.length,sumExtra,avgDI:dis.length?dis.reduce((s,x)=>s+x.f.dI,0)/dis.length:null,
  nAct:acc.reduce((s,x)=>s+((x.p.actions||[]).length),0),
  items:acc.map(x=>({shifr:x.r.shifr,name:x.r.name,p:x.p,f:x.f}))};}""")
replfn("function buildPropsCsv(){", "buildPropsCsv", r"""function buildPropsCsv(){
 let csv="Шифр;Наименование;Предложение;Статус;Дата;Доп. финансирование млн;Срок мес;Целевые коды (С);I до;I прогноз;ΔI\n";
 state.list.forEach(r=>(r.props||[]).forEach(p=>{const f=propForecast(r,p);
  const tg=Object.keys(p.targets||{}).map(k=>k+"="+p.targets[k]).join(" ");
  csv+=[r.shifr,String(r.name).replace(/;/g,","),String(p.name||"").replace(/;/g,","),p.status,(p.dateStr||""),p.extra||"",p.months||"",tg.replace(/;/g,","),
   f.I0==null?"":f.I0.toFixed(1),f.I1==null?"":f.I1.toFixed(1),f.dI==null?"":f.dI.toFixed(1)].join(";")+"\n";}));
 return csv;}""")

# =====================================================================
# 14) визуализация
# =====================================================================
replfn("function radarSVG(vis){", "radarSVG", r"""function radarSVG(vis){
 const cr=cur(),cx=idx18(cr,"S");const cur3=[cx.iA==null?cx.IA:cx.iA,cx.IC,cx.IR];
 const avg=[0,1,2].map(d=>{const V=vis.map(i=>idx18(state.list[i],"S").dirs[d]).filter(v=>v!=null);
  return V.length?V.reduce((a,b)=>a+b,0)/V.length:0;});
 const CX=190,CY=140,R=98,LB=["А. Актуальность","Б. Целесообразность","В. Результативность"];
 const ang=i=>-Math.PI/2+i*2*Math.PI/3;
 const pt=(v,i)=>{const rr=R*Math.max(0,Math.min(100,v||0))/100;return[(CX+rr*Math.cos(ang(i))).toFixed(1),(CY+rr*Math.sin(ang(i))).toFixed(1)];};
 let s='<svg viewBox="0 0 380 260" width="100%">';
 [25,50,75,100].forEach(g2=>{s+='<polygon points="'+[0,1,2].map(i=>pt(g2,i).join(",")).join(" ")+'" fill="none" stroke="#dfe6f2"/>';});
 [0,1,2].forEach(i=>{const e=pt(100,i);s+='<line x1="'+CX+'" y1="'+CY+'" x2="'+e[0]+'" y2="'+e[1]+'" stroke="#c5d0e6"/>';
  const l=pt(120,i);s+='<text x="'+l[0]+'" y="'+l[1]+'" font-size="11" text-anchor="middle" fill="#1f3864">'+LB[i]+'</text>';});
 s+='<polygon points="'+[0,1,2].map(i=>pt(avg[i],i).join(",")).join(" ")+'" fill="rgba(31,56,100,.15)" stroke="#1f3864" stroke-width="2"/>';
 s+='<polygon points="'+[0,1,2].map(i=>pt(cur3[i]||0,i).join(",")).join(" ")+'" fill="rgba(226,78,28,.15)" stroke="#e24e1c" stroke-width="2" stroke-dasharray="4 3"/>';
 s+='<text x="10" y="252" font-size="11" fill="#667">сплошная — среднее по выборке · пунктир — текущая НИОКР (I_А/I_Ц/I_Р)</text></svg>';
 return s;}""")
i_bar=s.find('const cur3=(function(){const r=cur();const x=idx18(r,"S");return[x.iA,x.iC,x.iR];})();')
if i_bar<0:
    # radar уже заменён целиком — старая строка cur3 не нужна
    pass

replfn("function barsSVG(vis){", "barsSVG", r"""function barsSVG(vis){if(!vis.length)return'<p class="mini">Нет НИОКР в текущей выборке.</p>';
 const rows=vis.map(i=>{const r=state.list[i],c=calc(r);return{r,I:c?c.I:null,cat:c?c.cat:""};});
 const Wd=640,H=250,pl=42,pb=26;const bw=Math.min(46,(Wd-pl-12)/rows.length);
 const Y=v=>pb+(H-pb-16)*(1-v/100);
 let s='<svg viewBox="0 0 '+Wd+' '+H+'" width="100%">';
 [0,25,50,75,100].forEach(g2=>{s+='<line x1="'+pl+'" y1="'+Y(g2)+'" x2="'+(Wd-6)+'" y2="'+Y(g2)+'" stroke="'+(g2===0?"#9aa7bd":"#e8edf6")+'"/>'+
  '<text x="'+(pl-6)+'" y="'+(Y(g2)+4)+'" font-size="10" text-anchor="end" fill="#667">'+g2+'</text>';});
 [[85,"Высокий","#2e7d32"],[70,"Приемлемый","#7cb342"],[55,"Проблемный","#ef6c00"]].forEach(t2=>{s+='<line x1="'+pl+'" y1="'+Y(t2[0])+'" x2="'+(Wd-6)+'" y2="'+Y(t2[0])+'" stroke="'+t2[2]+'" stroke-dasharray="5 4" opacity=".55"><title>порог уровня «'+t2[1]+'» — '+t2[0]+'%</title></line>';});
 rows.forEach((x,i)=>{const X=pl+4+i*((Wd-pl-12)/rows.length);const v=x.I==null?0:x.I;const col=x.I==null?"#9e9e9e":(PLVLC[plvl(x.I)]||"#9e9e9e");
  s+='<rect x="'+X.toFixed(1)+'" y="'+Y(v).toFixed(1)+'" width="'+(bw-6).toFixed(1)+'" height="'+Math.max(2,(H-pb-16)*v/100).toFixed(1)+'" rx="3" fill="'+col+'"><title>'+x.r.shifr+": "+(x.I==null?"нет данных":x.I.toFixed(1)+"% · "+(plvl(x.I)||""))+'</title></rect>';
  if(i%2===0||rows.length<=14)s+='<text x="'+(X+(bw-6)/2).toFixed(1)+'" y="'+(H-4)+'" font-size="8" fill="#667" text-anchor="middle">'+String(x.r.shifr).slice(0,10)+'</text>';});
 return s+'</svg>';}""")

replfn("function vizTimelineSVG(){", "vizTimelineSVG", r"""function vizTimelineSVG(){
 const pts=[];state.list.forEach((r,idx)=>(r.audits||[]).forEach(a=>{if(a.I!=null)pts.push({idx,r,a,t:new Date(a.date).getTime()});}));
 if(!pts.length)return'<p class="mini">Нет зафиксированных аудитов с числовым I — зафиксируйте аудит (📸 на акте оценки или в строке реестра).</p>';
 pts.sort((x,y)=>x.t-y.t);
 const Wd=660,H=220,pl=36,pr=10,pt2=12,pb=26;
 const t0=pts[0].t,t1=(pts[pts.length-1].t===t0)?t0+1:pts[pts.length-1].t;
 const X=t=>pl+(Wd-pl-pr)*(t-t0)/(t1-t0);
 const Y=v=>pt2+(H-pb-pt2)*(1-v/100);
 let s='<svg viewBox="0 0 '+Wd+' '+H+'" width="100%">';
 [0,50,100].forEach(g2=>{s+='<line x1="'+pl+'" y1="'+Y(g2)+'" x2="'+(Wd-pr)+'" y2="'+Y(g2)+'" stroke="#e8edf6"/><text x="'+(pl-5)+'" y="'+(Y(g2)+4)+'" font-size="10" text-anchor="end" fill="#667">'+g2+'</text>';});
 const byR={};pts.forEach(p=>{(byR[p.idx]=byR[p.idx]||[]).push(p);});
 const pal=["#1f3864","#e24e1c","#2e7d32","#7b1fa2","#00838f","#ef6c00","#455a64","#c62828"];
 Object.keys(byR).forEach((k,n)=>{const arr=byR[k];const col=pal[n%pal.length];
  s+='<polyline fill="none" stroke="'+col+'" stroke-width="1.8" points="'+arr.map(p=>X(p.t).toFixed(1)+","+Y(p.a.I).toFixed(1)).join(" ")+'"/>';
  arr.forEach(p=>{s+='<circle cx="'+X(p.t).toFixed(1)+'" cy="'+Y(p.a.I).toFixed(1)+'" r="3.6" fill="'+(PLVLC[p.a.cat]||COLCAT[p.a.cat]||col)+'" stroke="'+col+'" stroke-width="1.4"><title>'+p.r.shifr+" · "+p.a.dateStr+": I="+p.a.I.toFixed(1)+"% ("+(p.a.cat||"—")+")"+(p.a.lvl?" · "+LVLN[p.a.lvl]:"")+(p.a.cov!=null?" · Coverage "+p.a.cov+"%":"")+'</title></circle>';});});
 s+='<text x="'+pl+'" y="'+(H-6)+'" font-size="10" fill="#667">'+new Date(t0).toLocaleDateString("ru-RU")+'</text>'+
    '<text x="'+(Wd-pr)+'" y="'+(H-6)+'" font-size="10" fill="#667" text-anchor="end">'+new Date(t1).toLocaleDateString("ru-RU")+'</text></svg>';
 return s;}""")

replfn("function heatHTML(vis){", "heatHTML", r"""function heatHTML(vis){if(!vis.length)return'<p class="mini">Нет НИОКР в выборке.</p>';
 const LB3=["А","Ц","Р"];const NM3=["Актуальность","Целесообразность","Результативность"];
 let h='<table><tr><th style="text-align:left">НИОКР</th>'+LB3.map((t2,i)=>'<th>'+t2+' — '+NM3[i]+'</th>').join("")+'<th>I (∛)</th></tr>';
 vis.forEach(i=>{const r=state.list[i],c=calc(r),x=idx18(r,"S");
  h+='<tr><td style="text-align:left"><b>'+r.shifr+'</b></td>'+
   x.dirs.map(v=>{const col=v==null?"#eef2f9":(PLVLC[plvl(v)]||"#eef2f9");const fg=v==null?"#8c8c8c":"#fff";
    return'<td style="text-align:center;background:'+col+';color:'+fg+';font-weight:700">'+(v==null?"н/д":v.toFixed(0))+'</td>';}).join("")+
   '<td style="text-align:center;background:'+(c&&c.I!=null?(PLVLC[plvl(c.I)]||"#eef2f9"):"#eef2f9")+';color:#fff;font-weight:700">'+(c&&c.I!=null?c.I.toFixed(0):"—")+'</td></tr>';});
 return h+'</table><p class="mini">Ячейки — индексы направлений сетки «Средний» (I_А · I_Ц · I_Р), %.</p>';}""")

replfn("function expertMatrixCard(vis){", "expertMatrixCard", r"""function expertMatrixCard(vis){
 let h='<div class="card"><h3>🧭 Матрица портфеля «Актуальность × Целесообразность» (модель v18.0)</h3>'+
 '<p class="mini">X — I_Ц (%), Y — I_А (%) по сетке «Средний»; полосы 4×4: ≥85 · 70–84 · 55–69 · &lt;55. Цвет ячейки — группа рекомендации; красное кольцо — SF-2. Наведение — полный текст рекомендации.</p>';
 const Wd=560,Hh=400,pl=52,pb=40,pt2=14,pr=14;
 const X=v=>pl+(Wd-pl-pr)*v/100, Yc=v=>pt2+(Hh-pb-pt2)*(1-v/100);
 const bands=[0,55,70,85,100];
 h+='<svg viewBox="0 0 '+Wd+' '+Hh+'" width="100%" style="max-width:'+Wd+'px">';
 for(let a=0;a<4;a++)for(let c=0;c<4;c++){const rx1=X(bands[c]),rx2=X(bands[c+1]);
  const ry1=Yc(bands[a+1]),ry2=Yc(bands[a]);
  const rec=RECMX4[a][3-c];
  h+='<rect x="'+rx1.toFixed(1)+'" y="'+ry1.toFixed(1)+'" width="'+(rx2-rx1).toFixed(1)+'" height="'+(ry2-ry1).toFixed(1)+'" fill="'+recGroup4(rec[1])+'" opacity="0.12" stroke="#dfe6f2"><title>I_А «'+PLVLO[a]+'» × I_Ц «'+PLVLO[3-c]+'»: '+rec[0]+'</title></rect>';}
 bands.forEach(v=>{h+='<line x1="'+X(v)+'" y1="'+Yc(0)+'" x2="'+X(v)+'" y2="'+Yc(100)+'" stroke="#c5d0e6" stroke-width="'+(v===0?1.4:0.7)+'"/>'+
  '<line x1="'+X(0)+'" y1="'+Yc(v)+'" x2="'+X(100)+'" y2="'+Yc(v)+'" stroke="#c5d0e6" stroke-width="'+(v===0?1.4:0.7)+'"/>';});
 h+='<text x="'+(pl+(Wd-pl-pr)/2)+'" y="'+(Hh-4)+'" font-size="11" fill="#1f3864" text-anchor="middle">Индекс целесообразности I_Ц, %</text>'+
    '<text transform="rotate(-90 14 '+(pt2+(Hh-pb-pt2)/2)+')" x="14" y="'+(pt2+(Hh-pb-pt2)/2)+'" font-size="11" fill="#1f3864" text-anchor="middle">Индекс актуальности I_А, %</text>';
 const nod=[];const cnt={};
 vis.forEach(i=>{const r=state.list[i],x=idx18(r,"S"),sf=stops18(r);
  if(x.IA==null||x.IC==null){nod.push(r.shifr);return;}
  const la=plvlIdx(x.IA),lc=plvlIdx(x.IC),rec=RECMX4[la][lc];
  cnt[rec[1]]=(cnt[rec[1]]||0)+1;
  const col=recGroup4(rec[1]);
  h+='<circle cx="'+X(x.IC).toFixed(1)+'" cy="'+Yc(x.IA).toFixed(1)+'" r="5.5" fill="'+col+'" fill-opacity=".85" stroke="'+(sf.n2>0?"#c62828":"#37474f")+'" stroke-width="'+(sf.n2>0?3:1.2)+'" data-mxniokr="'+i+'"><title>'+r.shifr+" «"+String(r.name).slice(0,40)+"»\nI_А "+x.IA.toFixed(1)+"% («"+PLVLO[la]+"») · I_Ц "+x.IC.toFixed(1)+"% («"+PLVLO[lc]+"»)\n"+rec[0]+(sf.n2>0?"\n⛔ SF-2: "+sf.list.join("; "):"")+'</title></circle>';});
 leg='';["Продолжение","Корректировка","Пересмотр","Приостановка","Прекращение"].forEach((n2,g2)=>{leg+='<span style="display:inline-block;margin:2px 10px 2px 0"><span style="display:inline-block;width:10px;height:10px;border-radius:5px;background:'+RECGCOL[g2]+'"></span> '+n2+": "+(cnt[g2]||0)+"</span>";});
 h+='</svg>';
 if(nod.length)h+='<p class="mini">Без данных I_А/I_Ц ('+nod.length+'): '+nod.join(", ")+'</p>';
 h+='<div class="mini" style="margin-top:4px">'+leg+'</div>';
 return h;}""")

repl('const catEntries=["Высокая","Хорошая","Удовлетворительная","Неудовлетворительная","Критическая","Не оценена","Недостаточно данных"]',
     'const catEntries=PLVLO.concat("Не оценена","Недостаточно данных")',
     "catEntries1",2)

# =====================================================================
# 15) хук migrate18extra (load + btnImp — по 1 вхождению на строку)
# =====================================================================
repl("if(migrate31to18(r))mig18N++;","if(migrate31to18(r))mig18N++;if(migrate18extra(r))mig18N++;","extra-hook",2)

# =====================================================================
# 16) niokrAPI
# =====================================================================
assert s.count("expertProspect18};")==1,"niokrAPI tail"
repl("expertProspect18};",
     "expertProspect18,NQ30,CODES30,Q30IDX,LEG2S31,PLVLO,ansS30,jS30,eS30,personAns30,personJ30,personE30,personCellMap30,personVec,consensusTargets30,docPos18,q18pos,qLabel18,idxFake18,fakeR18,migrate18extra,RECMX4,RECGCOL,RECGNAME,plvlIdx,recGroup4};",
     "niokrAPI")

# ---------- проверка и запись ----------
node_s = s
data2 = data.replace("<script>\n" + m.group(1) + "</script>", "<script>\n" + s + "</script>")
assert data2 != data, "no change produced"
io.open(MODEL, "w", encoding="utf-8").write(data2)
io.open("/tmp/model18_partA.js", "w", encoding="utf-8").write(node_s)
print("PART A OK — replacements:", nrep[0], "; model size:", len(data2.encode("utf-8")), "B")
