# -*- coding: utf-8 -*-
# Этап 18Б, шаг 1: слой данных Q18 + ядро расчёта v18 + миграция 31→30 + компат-обёртка calc().
# Запуск: /tmp/venv/bin/python3 scripts/patch18_data.py   (из корня workspace)
import json, re, sys

HTML = 'output/Аудит_НИОКР_интерактивная_модель.html'
s = open(HTML, encoding='utf-8').read()
orig_len = len(s)
q18 = json.load(open('scripts/data/q18.json', encoding='utf-8'))
q18_js = json.dumps(q18, ensure_ascii=False, separators=(',', ':'))

def rep(old, new, cnt=1):
    global s
    n = s.count(old)
    assert n == cnt, f'count({old[:60]!r}) = {n}, ждал {cnt}'
    s = s.replace(old, new, cnt)

# ============================================================ 1. ЯДРО v18 (вставка после FLAGBAD)
ANCHOR = 'const FLAGBAD=[1,1,1,1,1,1,0,0,1,0]; // 1 — тревожный ответ «ДА», 0 — тревожный ответ «НЕТ»'
assert s.count(ANCHOR) == 1

ENGINE = r'''
/* ================= ЭТАП 18Б (v18.0): ТРЁХУРОВНЕВАЯ СИСТЕМА АУДИТА 15 / 30 / 60 =================
   Методика «Научно.docx»: три самостоятельных опросника (Краткий 15 — экспресс-оценка, Средний 30 —
   углублённая оценка, Полный 60 — комплексный аудит), каждый по ТРЁМ направлениям:
   Актуальность (I_А), Целесообразность (I_Ц), Результативность (I_Р). Шкала 0–5 + «НД» + «НП».
   I_k = Σw·x / (5·Σw) · 100 %; «НП» исключается из индекса И знаменателя Coverage (требует письменного
   обоснования), «НД» исключается из числителя, но остаётся в знаменателе Coverage.
   Coverage ≥ 80/85/90 % (по уровню), иначе положительное заключение не выносится.
   Полный уровень: индекс направления = взвешенное среднее заполненности блоков D_k (веса блоков
   базовые А .20/.20/.20/.40, Б .25×4, В .30/.25/.30/.15 либо AHP при включённом переключателе).
   Заключения трёх уровней НЕ одинаковы: Краткий — 4 формулировки о необходимости дальнейшего изучения;
   Средний — категории I–IV о необходимости корректировки; Полный — решения R1–R7 о судьбе НИОКР.
   Расчёт зеркалит формулы Excel-калькулятора № 9 (верифицирован: 125 проверок, 0 ошибок). */
const Q18=__Q18JSON__;
const LVLS=['K','S','P'];
const LVLN={K:'Краткий',S:'Средний',P:'Полный'};
const LVLGOAL={K:'экспресс-оценка: имеет ли смысл дальнейшее изучение проекта',S:'углублённая оценка: причины проблем, обоснованность разработки, рекомендации',P:'комплексный аудит: всесторонняя научно-техническая, экономическая и управленческая экспертиза с математическим обоснованием решения'};
const POROG={K:80,S:85,P:90};
const PLVL=[[85,'Высокий','#2e7d32'],[70,'Приемлемый','#7cb342'],[55,'Проблемный','#ef6c00'],[0,'Критический','#c62828']];
const PLVLC={'Высокий':'#2e7d32','Приемлемый':'#7cb342','Проблемный':'#ef6c00','Критический':'#c62828'};
function plvl(x){if(x==null||!isFinite(x))return null;for(const p of PLVL)if(x>=p[0])return p[1];return'Критический';}
const rnd1=x=>Math.round(x*10)/10, rnd2=x=>Math.round(x*100)/100;
function qNum(a){if(a===''||a==null||a==='НП'||a==='НД')return null;const n=Number(a);return isFinite(n)?n:null;}
/* контрольные замечания вопроса (зеркало колонки Q опросников Excel № 9) */
function qFlags(cell){const a=cell?cell.a:'',n=qNum(a);let out='';
 if(n!=null&&n<=2&&!(cell.j&&String(cell.j).trim()))out+='⚠ балл 0–2: нет обоснования; ';
 if(n!=null&&n>=3&&!(cell.e&&String(cell.e).trim()))out+='⚠ балл 3–5: нет доказательства; ';
 if(a==='НП'&&!(cell.j&&String(cell.j).trim()))out+='⚠ НП: нет письменного обоснования; ';
 return out;}
function q18cell(r,lvl,code){const L=Q18.levels[lvl];const cell=((r.q&&r.q[lvl])||{})[code];return cell||{a:'',j:'',e:''};}
function idx18(r,lvl){
 lvl=lvl||r.lvl||'S';const L=Q18.levels[lvl],qq=(r.q&&r.q[lvl])||{};
 const res={lvl,IA:null,IC:null,IR:null,kompl:null,cov:null,thr:POROG[lvl],min:null,flags:0,ctrl:0,low2:0,nAns:0,dirs:[null,null,null],blocks:null,nByDir:[0,0,0]};
 const D=[{R:0,S:0,T:0},{R:0,S:0,T:0},{R:0,S:0,T:0}];
 L.q.forEach(q=>{const cell=qq[q.code]||{};const a=cell.a,n=qNum(a);
  if(qFlags(cell))res.ctrl++;
  if(n!=null&&n<=2)res.low2++;
  if(a!==''&&a!=null){res.nAns++;res.nByDir[q.dir]++;}
  if(a==='НП')return;
  D[q.dir].T+=q.w;
  if(n!=null){D[q.dir].R+=q.w;D[q.dir].S+=q.w*n;}});
 const tR=D[0].R+D[1].R+D[2].R,tT=D[0].T+D[1].T+D[2].T;
 res.cov=tT?rnd1(tR/tT*100):null;
 if(lvl==='P'){
  const BM={};Q18.levels.P.blocks.forEach(b=>{BM[b.code]={R:0,S:0};});
  L.q.forEach(q=>{const cell=qq[q.code]||{};const n=qNum(cell.a);
   if(cell.a==='НП'||cell.a===''||cell.a==null||n==null)return;
   BM[q.blk].R+=1;BM[q.blk].S+=n;});
  const ahpOn=!!(r.ahp18&&r.ahp18.useBlocks),AW=ahp18Weights(r);
  res.blocks=Q18.levels.P.blocks.map(b=>{const m=BM[b.code];
   const dk=m.R?rnd1(m.S/(5*m.R)*100):null;
   const E=ahpOn&&AW[b.code]!=null?AW[b.code]:b.w;
   return{code:b.code,name:b.name,dir:b.dir,base:b.w,E:E,dk:dk,G:dk!=null?E:0,H:dk!=null?E*dk:0};});
  [0,1,2].forEach(dd=>{const g=res.blocks.filter(b=>b.dir===dd);
   const G=g.reduce((s2,b)=>s2+b.G,0),H=g.reduce((s2,b)=>s2+b.H,0);
   res.dirs[dd]=G?rnd1(H/G):null;});
 }else{
  [0,1,2].forEach(dd=>{res.dirs[dd]=D[dd].R?rnd1(D[dd].S/(5*D[dd].R)*100):null;});
 }
 res.IA=res.dirs[0];res.IC=res.dirs[1];res.IR=res.dirs[2];
 res.kompl=(res.IA!=null&&res.IC!=null&&res.IR!=null)?rnd1(Math.pow(res.IA*res.IC*res.IR,1/3)):null;
 res.min=(res.IA==null||res.IC==null||res.IR==null)?null:Math.min(res.IA,res.IC,res.IR);
 res.covOk=res.cov!=null&&res.cov>=res.thr;
 return res;}
/* -------- стоп-факторы: 5 проверок SF-2 + 5 «иных обстоятельств» (зеркало листа «Стоп_факторы») -------- */
function sfVal(v){if(!v||v==='не выявлено')return 0;if(v==='выявлена неопределённость')return 1;return 2;}
function stops18(r){const sf=r.sf||[];const vals=sf.map(sfVal);
 const max=vals.length?Math.max.apply(null,vals):0;
 const n1=vals.filter(v=>v===1).length,n2=vals.filter(v=>v===2).length;
 const list=(r.sf||[]).map((v,i)=>sfVal(v)===2?(Q18.sf2[i]||('Иное обстоятельство '+(i-4))):null).filter(Boolean);
 const res=n2>0?'SF-2: эскалация руководству; безусловно положительное заключение невозможно до рассмотрения критического обстоятельства'
  :(n1>0?'SF-1: обязательная дополнительная проверка (отдельный план проверки; не равен автоматически критическому)'
  :'SF-0: обычный порядок оценки');
 return{max:max,n1:n1,n2:n2,res:res,list:list};}
/* -------- зрелость M0–M9 (зеркало листа «Зрелость_M0_M9») -------- */
function mat18(r){const m=r.mat||{};const cur=(m.cur==null||m.cur==='')?null:Number(m.cur),plan=(m.plan==null||m.plan==='')?null:Number(m.plan);
 const dev=(cur!=null&&plan!=null&&isFinite(cur)&&isFinite(plan))?cur-plan:null;
 const note=dev==null?'':(dev>0?'опережение планового этапа':(dev<0?'⚠ отставание от планового этапа':'соответствует плановому этапу'));
 const warn=(cur!=null&&isFinite(cur)&&cur>=5&&m.tests!=='Да')?('⚠ Зрелость M'+cur+' заявлена без подтверждения испытаниями — недостаточно доказательств'):'OK';
 const nm=cur!=null&&Q18.m09[cur]?Q18.m09[cur][0]+' «'+Q18.m09[cur][1]+'»':null;
 return{cur:cur,plan:plan,dev:dev,note:note,warn:warn,name:nm,tests:m.tests||'Нет'};}
/* -------- AHP: 4 матрицы (направления 3×3, блоки А/Б/В 4×4) — λmax, CI, CR по таблице Сати -------- */
function ahpNum(v){if(v==null||v==='')return 1;if(typeof v==='number')return v;
 const t=String(v).trim();const m=t.match(/^1\s*\/\s*(\d+(?:\.\d+)?)$/);if(m)return 1/parseFloat(m[1]);
 const n=parseFloat(t.replace(',','.'));return isFinite(n)&&n>0?n:1;}
function ahp18matrix(o,n){const M=[];for(let i=0;i<n;i++){M.push([]);for(let j=0;j<n;j++)M[i].push(i===j?1:1);}
 for(let i=0;i<n;i++)for(let j=i+1;j<n;j++){const v=ahpNum(o['a'+(i+1)+(j+1)]);M[i][j]=v;M[j][i]=1/v;}
 return M;}
function ahp18one(o,n,ri){const M=ahp18matrix(o,n);
 const w=[];for(let i=0;i<n;i++){let p=1;for(let j=0;j<n;j++)p*=M[i][j];w.push(Math.pow(p,1/n));}
 const sw=w.reduce((a,b)=>a+b,0)||1;const wn=w.map(x=>x/sw);
 let lmax=0;for(let i=0;i<n;i++){let s2=0;for(let j=0;j<n;j++)s2+=M[i][j]*wn[j];lmax+=s2/(n*wn[i]);}
 const ci=(lmax-n)/(n-1),cr=ri>0?ci/ri:0;
 return{w:wn,wPct:wn.map(x=>rnd1(x*100)),lmax:lmax,ci:ci,cr:cr,ok:cr<=0.10+1e-9,M:M};}
function ahp18cr(r){const a=r.ahp18||{};
 const dir=ahp18one(a.dir||{},3,Q18.ri[2]);
 const A=ahp18one(a.A||{},4,Q18.ri[3]),B=ahp18one(a.B||{},4,Q18.ri[3]),V=ahp18one(a.V||{},4,Q18.ri[3]);
 const crMax=Math.max(dir.cr,A.cr,B.cr,V.cr);
 return{dir:dir,A:A,B:B,V:V,crMax:crMax,ok:crMax<=0.10+1e-9};}
function ahp18Weights(r){const st=ahp18cr(r),out={};
 ['A','B','V'].forEach((k,ki)=>{const sec=['А','Б','В'][ki];
  st[k].w.forEach((wv,i)=>{out[sec+(i+1)]=wv;});});
 return out;}
/* -------- экономика: LCC (8 стадий), NPV, BCR, сравнение 4 альтернатив (зеркало «Экономика_LCC») -------- */
const ALTNAMES4=['Продолжить разработку'].concat((Q18.altnames||[]).slice(1,4));
function econ18(r){const e=r.econ||{};const rate=(e.rate==null||e.rate==='')?0.1:Number(e.rate);
 const num=v=>(v===''||v==null||!isFinite(Number(v)))?null:Number(v);
 let pvB=0,pvC=0,npv=null,bcr=null,lccPv=null;
 if(isFinite(rate)&&rate>-1){
  (e.flows||[]).forEach((f,t)=>{const b=num(f&&f[0]),c=num(f&&f[1]);
   if(b!=null)pvB+=rnd2(b/Math.pow(1+rate,t));
   if(c!=null)pvC+=rnd2(c/Math.pow(1+rate,t));});
  pvB=rnd2(pvB);pvC=rnd2(pvC);
  if((e.flows||[]).some(f=>num(f&&f[0])!=null||num(f&&f[1])!=null))npv=rnd2(pvB-pvC);
  let lsum=0,any=false;
  (e.lcc||[]).forEach(st=>{const v=num(st&&st[0]),t=num(st&&st[1]);
   if(v!=null&&t!=null){lsum+=rnd2(v/Math.pow(1+rate,t));any=true;}});
  if(any)lccPv=rnd2(lsum);}
 if(npv!=null&&pvC)bcr=rnd2(pvB/pvC);
 const bcrTxt=bcr==null?'—':(bcr>1.5?'Высокая экономическая привлекательность':(bcr>=1?'Приемлемая':(bcr>=0.7?'Требуется дополнительное обоснование':'Низкая экономическая привлекательность')));
 const altW=(e.altW||[]).map(w2=>num(w2)==null?1:num(w2));while(altW.length<9)altW.push(1);
 const alt=e.alt||{};const Us=['c','d','e','f'].map(k=>{const arr=alt[k]||[];let s2=0;
  for(let i=0;i<9;i++){const v=num(arr[i]);if(v!=null)s2+=altW[i]*v;}
  return rnd2(s2);});
 let bi=0;Us.forEach((u,i)=>{if(u>Us[bi])bi=i;});
 const anyAlt=(e.alt&&['c','d','e','f'].some(k=>(e.alt[k]||[]).some(v=>num(v)!=null)));
 return{rate:rate,pvB:pvB,pvC:pvC,npv:npv,bcr:bcr,bcrTxt:bcrTxt,lccPv:lccPv,
  Us:anyAlt?Us:null,best:anyAlt?ALTNAMES4[bi]:null,margin:anyAlt?rnd2(Us[bi]-Us[0]):null,altNames:ALTNAMES4};}
/* -------- авторекомендации (15 правил, балл ≤ 2; зеркало «Заключение» 28–42: срабатывают на уровне «Средний») -------- */
function autoRecs18(r){const out=[];if((r.lvl||'S')!=='S')return out;
 const qs=(r.q&&r.q.S)||{};
 Q18.auto_rec.forEach(rule=>{const cell=qs[rule.code]||{};const n=qNum(cell.a);
  if(n!=null&&n<=rule.le)out.push({code:rule.code,rec:rule.rec,val:n});});
 return out;}
/* -------- общие правила заключения (зеркало «Заключение» 47–50) -------- */
function genRules18(r,lvl,x){lvl=lvl||r.lvl||'S';x=x||idx18(r,lvl);const out=[];
 if(x.IA!=null&&x.IA<55)out.push({rule:'I_А < 55%'});
 if(x.IC!=null&&x.IC<55)out.push({rule:'I_Ц < 55%'});
 if(x.IR!=null&&x.IR<55)out.push({rule:'I_Р < 55%'});
 if(x.cov!=null&&x.cov<x.thr)out.push({rule:'Coverage ниже порога'});
 return out;}
/* -------- заключения трёх уровней — формулировки НЕ одинаковы (зеркало Краткий!C35, Средний!C56, Полный!C102) -------- */
function concl18(r,lvl){lvl=lvl||r.lvl||'S';
 const x=idx18(r,lvl),st=stops18(r),ec=econ18(r);
 let txt,kind='nodata';
 if(x.min==null){txt='Недостаточно данных для достоверной оценки';}
 else if(st.n2>0){kind='sf2';txt=lvl==='K'?'Выявлен SF-2: безусловно положительное заключение невозможно; эскалация руководству'
  :(lvl==='S'?'IV. Отрицательное / эскалация — выявлен SF-2: передать на рассмотрение комиссии, оценить изменение концепции и альтернативы'
  :'R7 — передать на рассмотрение комиссии по прекращению или изменению проекта (выявлен SF-2)');}
 else if(x.cov!=null&&x.cov<x.thr){kind='cov';txt=lvl==='K'?'Оценка предварительная. Требуется получение доказательств (Coverage < 80%)'
  :(lvl==='S'?'Недостаточно данных для достоверной оценки (Coverage ниже порога 85%)'
  :'Недостаточно данных для достоверной оценки (Coverage < 90%)');}
 else if(lvl==='K'){kind=x.min>=80?'k80':(x.min>=60?'k60':(x.min>=40?'k40':'k0'));
  txt=x.min>=80?'Предварительно обоснованная НИОКР. Оснований для внепланового углублённого аудита не выявлено'
  :(x.min>=60?'НИОКР имеет основания для продолжения, необходим мониторинг выявленных недостатков'
  :(x.min>=40?'Выявлены признаки недостаточной актуальности или целесообразности. Рекомендован средний опросник'
  :'Выявлены существенные сомнения в обоснованности НИОКР. Необходим внеплановый углублённый аудит'));}
 else if(lvl==='S'){kind=x.min<55?'s4':(x.min>=85?'s1':(x.min>=70?'s2':'s3'));
  txt=x.min<55?'IV. Отрицательное / эскалация — передать на рассмотрение комиссии, оценить изменение концепции и альтернативы'
  :(x.min>=85?'I. Положительное — продолжить НИОКР'
  :(x.min>=70?'II. Положительное с замечаниями — продолжить с корректирующими мероприятиями'
  :'III. Условно отрицательное — утвердить план восстановления, провести повторную оценку'));}
 else{
  const dkB3=x.blocks?(x.blocks.filter(b=>b.code==='Б3')[0]||{}).dk:null;
  if(x.IA<55&&x.IC<55&&x.IR<55){kind='p7a';txt='R7 — передать на рассмотрение комиссии по прекращению или изменению проекта (все три индекса < 55%)';}
  else if(x.IA>=85&&x.IC>=85&&x.IR>=85){kind='p1';txt='R1 — продолжить без существенных изменений';}
  else if(x.min>=70){kind='p2';txt='R2 — продолжить с корректировкой';}
  else if(ec.best&&ec.best!=='Продолжить разработку'){kind='p6';txt='R6 — рассмотреть альтернативное решение (по сравнению альтернатив предпочтительнее: '+ec.best+')';}
  else if(dkB3!=null&&dkB3<70){kind='p5';txt='R5 — пересмотреть сроки и ресурсы (блок Б3 «Ресурсы и ресурсы и реализуемость» < 70%)';}
  else if(x.IA<70&&x.IC>=70){kind='p4';txt='R4 — изменить состав и объём работ (актуальность снижена при обоснованной целесообразности)';}
  else if(x.IC<70||x.IR<70){kind='p3';txt='R3 — изменить техническую концепцию (целесообразность или результативность < 70%)';}
  else{kind='p2';txt='R2 — продолжить с корректировкой';}}
 return{txt:txt,kind:kind,lvl:lvl,x:x,st:st,ec:ec};}
/* -------- перспективы завершения (по I_Ц активного уровня; ручной ввод приоритетен) -------- */
function expertProspect18(r){if(r.prospect)return r.prospect;const st=stops18(r);if(st.n2>0)return'Маловероятная';
 const x=idx18(r);if(x.IC==null)return'—';if(x.IC>=70)return'Реалистичная';if(x.IC>=50)return'Условно реалистичная';return'Маловероятная';}
/* -------- ранние сигналы (зеркало «Панель_руководителя» 24–30; EARLY-правила) -------- */
function early18(r,x,st){x=x||idx18(r);st=st||stops18(r);const out=[];const pv=r.prev||{};
 const dl=(k)=>(x[k]!=null&&pv[k]!=null&&isFinite(pv[k]))?rnd1(x[k]-pv[k]):null;
 [['IA','I_А','Предупреждение','Повторная проверка потребности'],['IC','I_Ц','Предупреждение','Пересмотр концепции и ресурсов'],['IR','I_Р','Критический сигнал','Проверка технических результатов']].forEach(sp=>{
  const d=dl(sp[0]);if(d!=null&&d<=-10)out.push({sig:'⚠ '+sp[2],rule:sp[1]+' снизился на '+Math.abs(d).toFixed(1)+' п. и более',act:sp[3]});});
 if(x.cov!=null&&x.cov<85)out.push({sig:'⚠ Недостаточно данных',rule:'Coverage = '+x.cov+'% (< 85%)',act:'Запрос доказательств'});
 if(st.n2>0)out.push({sig:'⚠ Эскалация',rule:'Появился SF-2 ('+st.n2+')',act:'Внеплановое заседание комиссии'});
 if(r.endDate){const days=Math.round((new Date(r.endDate)-new Date())/86400000);
  if(isFinite(days)&&days<0)out.push({sig:'⚠ Отклонение',rule:'Срок этапа превышен на '+(-days)+' дн.',act:'Анализ критического пути'});}
 return out;}
/* -------- миграция v17 (31 вопрос) → v18 (Средний 30): семантическая карта Q18.mig31 -------- */
function migAggregate(r,codes31){const nums=[],nps=[],nds=[];let j='',e='';
 codes31.forEach(oi=>{const a=(r.ans||[])[oi];const cm=(r.com||[])[oi];const cf=(r.conf||[])[oi];
  if(a==='НП'){nps.push(oi);if(cm&&!j)j=cm;}
  else{const n=qNum(a);
   if(n!=null){nums.push(n);if(cm&&!j)j=cm;if(cf==='Да'&&!e)e='Подтверждено (миграция данных v17)';}
   else if(a==='НД')nds.push(oi);}});
 let a='';
 if(nums.length)a=String(Math.round((nums.reduce((s2,x)=>s2+x,0)/nums.length)*2)/2);
 else if(nps.length)a='НП';
 else if(nds.length)a='НД';
 return{a:a,j:j,e:e};}
function migrate31to18(r){
 if(!r.ans||r.ans.length!==31)return false;
 if(r.q&&r.q.S&&Object.keys(r.q.S).length)return false;
 const tgt={};Q18.mig31.forEach(pr=>{const oi=pr[0],code=pr[1];(tgt[code]=tgt[code]||[]).push(oi);});
 r.q=r.q||{K:{},S:{},P:{}};if(!r.q.K)r.q.K={};if(!r.q.S)r.q.S={};if(!r.q.P)r.q.P={};
 Object.keys(tgt).forEach(code=>{const c=migAggregate(r,tgt[code]);if(c.a!==''||c.j)r.q.S[code]=c;});
 if(!r.lvl)r.lvl='S';
 if(!r.legacy31)r.legacy31={ans:(r.ans||[]).slice(),conf:(r.conf||[]).slice(),com:(r.com||[]).slice(),stops:(r.stops||[]).slice(),flags:(r.flags||[]).slice()};
 if(!r.sf)r.sf=Array(10).fill('не выявлено');
 Q18.migstops.forEach((ni,oi)=>{if(ni!=null&&(r.stops||[])[oi]==='ДА')r.sf[ni]='подтверждено';});
 if(!r.mat)r.mat={cur:null,plan:null,tests:'Нет'};
 if(!r.econ)r.econ={rate:0.1,T:5,lcc:Array.from({length:8},()=>['','']),flows:Array.from({length:11},()=>['','']),altW:Array(9).fill(1),alt:{c:Array(9).fill(''),d:Array(9).fill(''),e:Array(9).fill(''),f:Array(9).fill('')}};
 if(!r.ahp18)r.ahp18={dir:{a12:'1',a13:'1',a23:'1'},A:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},B:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},V:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},useBlocks:false};
 if(!r.prev)r.prev={IA:null,IC:null,IR:null};
 (r.props||[]).forEach(p=>{if(p.targets&&!p.targetsLegacy){const nt={};
  Object.keys(p.targets).forEach(k=>{const hit=Q18.mig31.filter(pr=>String(pr[0])===String(k))[0];
   if(hit){const cur=nt[hit[1]],v=Number(p.targets[k]);if(cur==null||v>Number(cur))nt[hit[1]]=String(v);}});
  p.targetsLegacy=p.targets;p.targets=nt;}});
 if(r.resp&&r.resp.data)Object.keys(r.resp.data).forEach(pid=>{const dd=r.resp.data[pid];
  if(dd.ans&&dd.ans.length===31&&!dd.qS){const fake={ans:dd.ans,com:dd.com||[],conf:dd.conf||[]};const qs={};
   Object.keys(tgt).forEach(code=>{const c=migAggregate(fake,tgt[code]);if(c.a!==''||c.j)qs[code]=c;});
   dd.qS=qs;dd.ansLegacy=dd.ans;}});
 return true;}
'''
ENGINE = ENGINE.replace('__Q18JSON__', q18_js)
# опечатка-страховка: исправляем задвоенное слово в R5 (должно совпадать с текстом Excel)
ENGINE = ENGINE.replace('«Ресурсы и ресурсы и реализуемость»', '«Ресурсы и реализуемость»')
rep(ANCHOR, ANCHOR + '\n' + ENGINE)

# ============================================================ 2. calc(r) — компат-обёртка v18
OLD_CALC = '''function calc(r){
  const c=calcCore(r.ans,state.weights);
  c.nCtrl=r.ctrl.filter(x=>x==="Выявлено").length;
  const dA=dirIdx(r,0),dC=dirIdx(r,2);
  c.iA=dA.I;c.iC=dC.I;c.nA=dA.n;c.nC=dC.n;c.lA=lvl5(dA.I);c.lC=lvl5(dC.I);
  const er=expertRec(r);c.rec=er.text;c.recStop=er.stop;c.stopsHit=er.st;
  c.prospect=expertProspect(r);
  return c;
}'''
NEW_CALC = '''function calc(r){
  /* v18.0: ядро — трёхуровневая система 15/30/60 (idx18/concl18); форма ответа совместима
     с потребителями v5–v17: bl[] = ТРИ НАПРАВЛЕНИЯ (P = I_d/100), I = комплексный индекс
     (геометрическое среднее), cat = уровень 85/70/55, rec = заключение активного уровня. */
  const lvl=r.lvl||'S',x=idx18(r,lvl),cc=concl18(r,lvl),st=cc.st,ec=cc.ec,mt=mat18(r);
  const c={bl:[0,1,2].map(d=>{const I=x.dirs[d];const P=I==null?null:I/100;
    return{S:I==null?0:I,N:x.nByDir[d],P:P,B:P==null?0:P*100/3,Wapp:P==null?0:100/3};}),
   I:x.kompl,cat:plvl(x.kompl),col:PLVLC[plvl(x.kompl)]||'#9e9e9e',rec:cc.txt,renorm:false,
   Wsum:x.min==null?0:100,nCtrl:(r.ctrl||[]).filter(t=>t==="Выявлено").length};
  c.lvl=lvl;c.lvlName=LVLN[lvl];c.idx=x;c.iA=x.IA;c.iC=x.IC;c.iR=x.IR;c.nA=x.nByDir[0];c.nC=x.nByDir[1];
  c.lA=plvl(x.IA);c.lC=plvl(x.IC);c.lR=plvl(x.IR);
  c.cov=x.cov;c.thr=x.thr;c.covOk=x.covOk;c.min=x.min;c.kompl=x.kompl;c.conc=cc.txt;c.concKind=cc.kind;
  c.ctrlQ=x.ctrl;c.low2=x.low2;c.flagsQ=x.flags;
  c.recs=autoRecs18(r);c.gen=genRules18(r,lvl,x);
  c.sf=st;c.sf2=st.n2;c.recStop=st.n2>0;c.stopsHit=st.list;
  c.rec=cc.txt;c.prospect=expertProspect18(r);c.mat=mt;c.econ=ec;
  c.early=early18(r,x,st);c.blocks=x.blocks;
  return c;
}'''
rep(OLD_CALC, NEW_CALC)

# ============================================================ 3. newRnd — новые поля v18
OLD_NR = 'ans:Array(NQ).fill(""),conf:Array(NQ).fill(""),com:Array(NQ).fill(""),stops:Array(5).fill(""),flags:Array(10).fill(""),prospect:"",cat:"",docs:[],scenarios:[],snaps:[],flow:null,threads:{},resp:null,audits:[],props:[],created:new Date().toISOString()}}'
NEW_NR = '''lvl:"S",q:{K:{},S:{},P:{}},sf:Array(10).fill("не выявлено"),mat:{cur:null,plan:null,tests:"Нет"},
econ:{rate:0.1,T:5,lcc:Array.from({length:8},()=>["",""]),flows:Array.from({length:11},()=>["",""]),altW:Array(9).fill(1),alt:{c:Array(9).fill(""),d:Array(9).fill(""),e:Array(9).fill(""),f:Array(9).fill("")}},
ahp18:{dir:{a12:"1",a13:"1",a23:"1"},A:{a12:"1",a13:"1",a14:"1",a23:"1",a24:"1",a34:"1"},B:{a12:"1",a13:"1",a14:"1",a23:"1",a24:"1",a34:"1"},V:{a12:"1",a13:"1",a14:"1",a23:"1",a24:"1",a34:"1"},useBlocks:false},
prev:{IA:null,IC:null,IR:null},endDate:"",
ans:Array(NQ).fill(""),conf:Array(NQ).fill(""),com:Array(NQ).fill(""),stops:Array(5).fill(""),flags:Array(10).fill(""),prospect:"",cat:"",docs:[],scenarios:[],snaps:[],flow:null,threads:{},resp:null,audits:[],props:[],created:new Date().toISOString()}}'''
rep(OLD_NR, NEW_NR)

# ============================================================ 4. load() и btnImp: миграция 31→18
OLD_L1 = 'let migN=0;state.list=d.list;state.cur=d.cur||0;'
NEW_L1 = 'let migN=0,mig18N=0;state.list=d.list;state.cur=d.cur||0;'
rep(OLD_L1, NEW_L1)
OLD_L2 = 'if(r.form===undefined)r.form="";if(r.flow===undefined)r.flow=null;if(!r.threads)r.threads={};if(r.resp===undefined)r.resp=null;if(!r.audits)r.audits=[];if(!r.props)r.props=[];if(!r.created)r.created=new Date().toISOString();if(migrate27to31(r))migN++;});'
NEW_L2 = 'if(r.form===undefined)r.form="";if(r.flow===undefined)r.flow=null;if(!r.threads)r.threads={};if(r.resp===undefined)r.resp=null;if(!r.audits)r.audits=[];if(!r.props)r.props=[];if(!r.created)r.created=new Date().toISOString();if(migrate27to31(r))migN++;if(migrate31to18(r))mig18N++;});'
rep(OLD_L2, NEW_L2)
OLD_L2B = 'let migN=0;state.list.forEach(r=>{if(r.flow===undefined)r.flow=null;if(!r.threads)r.threads={};if(r.resp===undefined)r.resp=null;if(!r.audits)r.audits=[];if(!r.props)r.props=[];if(!r.created)r.created=new Date().toISOString();if(migrate27to31(r))migN++;});'
NEW_L2B = 'let migN=0,mig18N=0;state.list.forEach(r=>{if(r.flow===undefined)r.flow=null;if(!r.threads)r.threads={};if(r.resp===undefined)r.resp=null;if(!r.audits)r.audits=[];if(!r.props)r.props=[];if(!r.created)r.created=new Date().toISOString();if(migrate27to31(r))migN++;if(migrate31to18(r))mig18N++;});'
rep(OLD_L2B, NEW_L2B)
OLD_L2C = '(migN?" · миграция 27→31: "+migN+"":"")'
NEW_L2C = '(migN?" · миграция 27→31: "+migN+"":"")+(mig18N?" · миграция 31→30 (v18.0): "+mig18N:"")'
rep(OLD_L2C, NEW_L2C)
OLD_L3 = '''  if(migN)setTimeout(()=>notify("info","Выполнена миграция данных (v15.0)",migN+" НИОКР перенесено со структуры 27 → 31 вопрос: блоки I и III пересобраны по экспертной методике «Актуальность + Целесообразность»; снимки истории аудитов сохранены как есть"),900);
  return true;}}catch(e){}return false;}'''
NEW_L3 = '''  if(migN)setTimeout(()=>notify("info","Выполнена миграция данных (v15.0)",migN+" НИОКР перенесено со структуры 27 → 31 вопрос: блоки I и III пересобраны по экспертной методике «Актуальность + Целесообразность»; снимки истории аудитов сохранены как есть"),900);
  if(mig18N)setTimeout(()=>notify("info","Выполнена миграция данных (v18.0)",mig18N+" НИОКР перенесено с чек-листа 31 вопрос на систему 15/30/60: ответы семантически отображены в Средний опросник (30 вопросов, коды А1–В10); прежние ответы сохранены в архиве legacy31; снимки истории аудитов не изменялись"),1400);
  return true;}}catch(e){}return false;}'''
rep(OLD_L3, NEW_L3)

# ============================================================ 5. niokrAPI — новые функции
OLD_API = 'exportChecklistMD,btnChecklistMD};'
NEW_API = 'exportChecklistMD,btnChecklistMD,get Q18(){return Q18;},idx18,concl18,autoRecs18,genRules18,ahp18cr,ahp18Weights,econ18,mat18,stops18,qFlags,plvl,early18,migrate31to18,sfVal,ahpNum,LVLN,POROG,ALTNAMES4};'
rep(OLD_API, NEW_API)

open(HTML, 'w', encoding='utf-8').write(s)
print(f'PATCH18-DATA OK: {orig_len} → {len(s)} (+{len(s)-orig_len})')
