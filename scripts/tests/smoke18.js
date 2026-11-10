const fs=require("fs");
const {JSDOM}=require("jsdom");
const P="/tmp/arena/tenants/71ebd7cc-b365-41c5-863c-79f587b1df1f/sessions/76eaa8bd-b9c6-43e4-8bfe-0b8d30b3272e/workspace/output/Аудит_НИОКР_интерактивная_модель.html";
const dom=new JSDOM(fs.readFileSync(P,"utf8"),{runScripts:"dangerously",pretendToBeVisual:true,url:"http://localhost/"});
const w=dom.window,d=w.document;
w.HTMLCanvasElement.prototype.getContext=()=>({clearRect(){},fillRect(){},beginPath(){},moveTo(){},lineTo(){},stroke(){},fill(){},arc(){},save(){},restore(){},translate(){},rotate(){},fillText(){},setLineDash(){},closePath(){},quadraticCurveTo(){}});
w.URL.createObjectURL=()=>"blob:x";w.URL.revokeObjectURL=()=>{};
let winErr=0;
w.addEventListener("error",e=>{if(/Not implemented/.test(e.message||""))return;winErr++;console.log("WINDOW-ERR:",e.message);});
let fails=0;
const ok=(n,c)=>{console.log((c?"PASS":"FAIL")+" | "+n);if(!c)fails++;};
const close=(n,a,b,t=0.051)=>ok(n+" (="+a+")",a!=null&&Math.abs(a-b)<=t);
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state;
 ok("нет ошибок окна при загрузке",winErr===0);
 ok("niokrAPI.Q18 доступен",!!API.Q18);
 // ---- 1. структура
 ok("Q18: уровни K15/S30/P60",API.Q18.levels.K.q.length===15&&API.Q18.levels.S.q.length===30&&API.Q18.levels.P.q.length===60);
 ok("Q18: Σw Средний = 10/10/10",[0,1,2].every(dd=>Math.abs(API.Q18.levels.S.q.filter(q=>q.dir===dd).reduce((s,q)=>s+q.w,0)-10)<1e-9));
 ok("Q18: Σw блоков Полный = 1/1/1",[0,1,2].every(dd=>Math.abs(API.Q18.levels.P.blocks.filter(b=>b.dir===dd).reduce((s,b)=>s+b.w,0)-1)<1e-9));
 ok("Q18: 12 блоков, 15 авторекомендаций, 5 SF-2, 10 M",API.Q18.levels.P.blocks.length===12&&API.Q18.auto_rec.length===15&&API.Q18.sf2.length===5&&API.Q18.m09.length===10);
 const r0=S.list[0];
 ok("newRnd: lvl='S', q={K,S,P}, sf×10, mat, econ, ahp18, prev",r0.lvl==="S"&&r0.q&&r0.q.K&&r0.q.S&&r0.q.P&&r0.sf.length===10&&r0.mat&&r0.econ&&r0.ahp18&&r0.prev);
 const setAll=(r,lvl,v)=>{API.Q18.levels[lvl].q.forEach(q=>{r.q[lvl][q.code]={a:String(v),j:"",e:""};});};
 const setAllJE=(r,lvl,v)=>{API.Q18.levels[lvl].q.forEach(q=>{r.q[lvl][q.code]={a:String(v),j:"обоснование",e:"доказательство"};});};
 // ---- 2. контрольный пример 1: все 5 → 100/100/100
 setAllJE(r0,"S",5);let x=w.idx18(r0,"S");
 close("S все 5: I_А=100",x.IA,100);close("S все 5: I_Ц=100",x.IC,100);close("S все 5: I_Р=100",x.IR,100);
 close("S все 5: Coverage=100",x.cov,100);close("S все 5: комплексный=100",x.kompl,100);
 ok("S все 5: заключение «I. Положительное — продолжить НИОКР»",w.concl18(r0,"S").txt==="I. Положительное — продолжить НИОКР");
 setAllJE(r0,"K",5);x=w.idx18(r0,"K");
 ok("K все 5: «Предварительно обоснованная НИОКР…»",w.concl18(r0,"K").txt.startsWith("Предварительно обоснованная НИОКР"));
 setAllJE(r0,"P",5);x=w.idx18(r0,"P");
 close("P все 5: I_А=100 (через блоки)",x.IA,100);
 ok("P все 5: «R1 — продолжить без существенных изменений»",w.concl18(r0,"P").txt==="R1 — продолжить без существенных изменений");
 // ---- 3. контрольный пример 2: все 0 → 0
 ["K","S","P"].forEach(lv=>{API.Q18.levels[lv].q.forEach(q=>{r0.q[lv][q.code]={a:"0",j:"",e:""};});});
 x=w.idx18(r0,"S");
 close("S все 0: I_А=0",x.IA,0);close("S все 0: I_Р=0",x.IR,0);
 ok("S все 0: ctrl-замечаний=30",x.ctrl===30);ok("S все 0: low2=30",x.low2===30);
 ok("S все 0: «IV. Отрицательное / эскалация — передать…»",w.concl18(r0,"S").txt.startsWith("IV. Отрицательное / эскалация — передать"));
 x=w.idx18(r0,"K");ok("K все 0: ctrl=15",x.ctrl===15);
 ok("K все 0: «существенные сомнения… внеплановый углублённый аудит»",w.concl18(r0,"K").txt.includes("внеплановый углублённый аудит"));
 x=w.idx18(r0,"P");ok("P все 0: ctrl=60",x.ctrl===60);
 ok("P все 0: «R7 … (все три индекса < 55%)»",w.concl18(r0,"P").txt.includes("R7")&&w.concl18(r0,"P").txt.includes("все три индекса < 55%"));
 // ---- 4. контрольный пример 3: НД → пересчёт + падение Coverage (Средний все 4, А1)
 API.Q18.levels.S.q.forEach(q=>{r0.q.S[q.code]={a:"4",j:"",e:""};});
 r0.q.S["А1"]={a:"5",j:"",e:""};x=w.idx18(r0,"S");
 close("S: все 4 + А1=5 → I_А=82.4",x.IA,82.4);close("S: Coverage=100",x.cov,100);
 r0.q.S["А1"]={a:"НД",j:"",e:""};x=w.idx18(r0,"S");
 close("S: А1=«НД» → I_А=80.0",x.IA,80.0);close("S: Coverage=96.0 (28.8/30)",x.cov,96.0);
 ok("S: min=80 → «II. Положительное с замечаниями…»",w.concl18(r0,"S").txt.startsWith("II. Положительное с замечаниями"));
 ok("S: 29 замечаний «3–5 без доказательства»",x.ctrl===29);
 // ---- 5. контрольный пример 4: НП с обоснованием и без
 API.Q18.levels.S.q.forEach(q=>{r0.q.S[q.code]={a:"5",j:"",e:""};});
 r0.q.S["А1"]={a:"НП",j:"Не применимо на данной стадии",e:""};x=w.idx18(r0,"S");
 close("S: А1=«НП»+обоснование → I_А=100",x.IA,100);close("S: Coverage=100 (НП вне знаменателя)",x.cov,100);
 ok("S: 29 замечаний (НП с обоснованием — без замечания)",x.ctrl===29);
 r0.q.S["А1"]={a:"НП",j:"",e:""};x=w.idx18(r0,"S");
 ok("S: НП без обоснования → 30 замечаний",x.ctrl===30);
 ok("S: qFlags НП без j",w.qFlags({a:"НП",j:"",e:""}).includes("НП: нет письменного обоснования"));
 // ---- 6. контрольные примеры 5–6: AHP
 let a=w.ahp18cr(r0);
 close("AHP: все 1 → w1=1/3",a.dir.w[0],1/3,0.001);close("AHP: CR=0",a.dir.cr,0,1e-9);
 close("AHP: блок А w=0.25",a.A.w[0],0.25,0.001);close("AHP: crMax=0",a.crMax,0,1e-9);
 r0.ahp18.dir={a12:"2",a13:"4",a23:"2"};a=w.ahp18cr(r0);
 close("AHP 2/4/2: w1=4/7",a.dir.w[0],4/7,0.001);close("AHP 2/4/2: w2=2/7",a.dir.w[1],2/7,0.001);
 close("AHP 2/4/2: w3=1/7",a.dir.w[2],1/7,0.001);close("AHP 2/4/2: λmax=3",a.dir.lmax,3,1e-6);
 close("AHP 2/4/2: CR=0",a.dir.cr,0,1e-6);
 r0.ahp18.dir={a12:"1",a13:"1",a23:"1"};
 // ---- 7. контрольный пример 7: SF-2 блокирует положительное заключение
 setAllJE(r0,"K",5);setAllJE(r0,"S",5);setAllJE(r0,"P",5);
 r0.sf[0]="подтверждено";let st=w.stops18(r0);
 ok("SF: n2=1",st.n2===1);ok("SF: max=2",st.max===2);
 ok("SF: res — эскалация + запрет",st.res.includes("эскалация")&&st.res.includes("безусловно положительное заключение невозможно"));
 ok("S при SF-2 и всех 100: «IV. Отрицательное / эскалация — выявлен SF-2…»",w.concl18(r0,"S").txt.includes("выявлен SF-2"));
 ok("P при SF-2 и всех 100: «R7 … (выявлен SF-2)»",w.concl18(r0,"P").txt.includes("R7")&&w.concl18(r0,"P").txt.includes("выявлен SF-2"));
 ok("K при SF-2: эскалация руководству",w.concl18(r0,"K").txt.includes("эскалация руководству"));
 r0.sf[1]="выявлена неопределённость";st=w.stops18(r0);ok("SF: n1=1 (неопределённость)",st.n1===1);
 r0.sf[0]="не выявлено";r0.sf[1]="не выявлено";st=w.stops18(r0);
 ok("SF: все «не выявлено» → SF-0",st.res==="SF-0: обычный порядок оценки");
 // ---- 8. контрольный пример 8: зрелость
 r0.mat={cur:5,plan:7,tests:"Нет"};let m=w.mat18(r0);
 ok("M: warn «без подтверждения испытаниями»",m.warn.includes("без подтверждения испытаниями"));
 ok("M: dev=-2",m.dev===-2);ok("M: note «⚠ отставание от планового этапа»",m.note==="⚠ отставание от планового этапа");
 ok("M: name = «M5 …»",m.name&&m.name.startsWith("M5"));
 r0.mat={cur:5,plan:5,tests:"Да"};m=w.mat18(r0);
 ok("M: испытания подтверждены → OK",m.warn==="OK"&&m.note==="соответствует плановому этапу");
 r0.mat={cur:null,plan:null,tests:"Нет"};
 // ---- 9. экономика
 r0.econ.rate=0;r0.econ.flows=Array.from({length:11},(_,i)=>i===0?["200","100"]:["",""]);
 let e=w.econ18(r0);
 close("Экон: NPV=100",e.npv,100,0.011);close("Экон: BCR=2",e.bcr,2,0.011);
 ok("Экон: BCR>1.5 → «Высокая экономическая привлекательность»",e.bcrTxt==="Высокая экономическая привлекательность");
 r0.econ.alt={c:Array(9).fill("5"),d:Array(9).fill(""),e:Array(9).fill(""),f:Array(9).fill("")};
 e=w.econ18(r0);
 close("Экон: Us «Продолжить разработку»=45",e.Us[0],45,0.011);
 ok("Экон: лучший = «Продолжить разработку»",e.best==="Продолжить разработку");close("Экон: отрыв=0",e.margin,0,0.011);
 r0.econ.alt.d=Array(9).fill("5");r0.econ.alt.c=Array(9).fill("1");e=w.econ18(r0);
 ok("Экон: альтернатива побеждает → best≠«Продолжить…»",e.best!=="Продолжить разработку"&&e.margin>0);
 r0.econ.rate=0.1;r0.econ.flows=Array.from({length:11},()=>["",""]);r0.econ.alt={c:Array(9).fill(""),d:Array(9).fill(""),e:Array(9).fill(""),f:Array(9).fill("")};
 e=w.econ18(r0);ok("Экон: пусто → NPV=null, BCR «—»",e.npv===null&&e.bcrTxt==="—"&&e.best===null);
 // ---- 10. авторекомендации (гейт: только уровень «Средний»)
 r0.lvl="S";API.Q18.levels.S.q.forEach(q=>{r0.q.S[q.code]={a:"5",j:"",e:"ev"};});
 r0.q.S["А6"]={a:"2",j:"поверхностный анализ",e:""};
 let rc=w.autoRecs18(r0);
 ok("Автореки: А6=2 → «Провести сравнительный анализ аналогов»",rc.length===1&&rc[0].code==="А6"&&rc[0].rec==="Провести сравнительный анализ аналогов");
 r0.lvl="K";rc=w.autoRecs18(r0);
 ok("Автореки: на уровне «Краткий» не срабатывают (зеркало Excel C28)",rc.length===0);
 r0.lvl="S";r0.q.S["А6"]={a:"5",j:"",e:"ev"};
 // ---- 11. Coverage-порог Краткого (S8b Excel)
 API.Q18.levels.K.q.forEach((q,i)=>{r0.q.K[q.code]=i<11?{a:"5",j:"",e:"ev"}:{a:"НД",j:"",e:""};});
 x=w.idx18(r0,"K");
 close("K: 11×5 + 4×НД → Coverage=73.3",x.cov,73.3);
 close("K: I_Р=100 (НД исключены из числителя)",x.IR,100);
 ok("K: cov<80 → «Оценка предварительная…» (несмотря на min=100)",w.concl18(r0,"K").txt.startsWith("Оценка предварительная"));
 // ---- 12. полосы заключений Краткого 40–59 / 60–79
 [ [[2,3,3,2,2],"Рекомендован средний опросник"], [[4,3,3,3,3],"основания для продолжения"] ].forEach(t=>{
  const vals=t[0];API.Q18.levels.K.q.forEach((q,i)=>{const v=vals[i%5];r0.q.K[q.code]={a:String(v),j:v<=2?"обосн":"",e:"ev"};});
  ok("K: ответы "+vals.join(",")+" → «"+t[1]+"»",w.concl18(r0,"K").txt.includes(t[1]));});
 // ---- 13. Полный: AHP-переключатель весов блоков
 setAllJE(r0,"P",5);r0.ahp18.useBlocks=true;r0.ahp18.A={a12:"1",a13:"1",a14:"1",a23:"1",a24:"1",a34:"1"};
 x=w.idx18(r0,"P");close("P: AHP вкл, все 5 → I_А=100 (веса не искажают)",x.IA,100);
 ok("P: блоки E = AHP-веса (0.25)",Math.abs(x.blocks[0].E-0.25)<0.001);
 r0.ahp18.useBlocks=false;x=w.idx18(r0,"P");
 ok("P: AHP выкл → базовый вес А1 = 0.2",Math.abs(x.blocks[0].E-0.2)<0.001);
 // ---- 14. R6/R5/R4/R3 ветви Полного
 r0.lvl="P";API.Q18.levels.P.q.forEach(q=>{r0.q.P[q.code]={a:"4",j:"",e:"ev"};});
 // все 4 → индексы 80: min≥70 → R2
 ok("P: все 4 → R2 (min=80≥70)",w.concl18(r0,"P").txt.startsWith("R2"));
 // А-блок 60 (3), Б/В 85+ (4.5→90): IA<70, IC≥70 → R4
 API.Q18.levels.P.q.forEach(q=>{r0.q.P[q.code]={a:q.dir===0?"3":"5",j:"",e:"ev"};});
 x=w.idx18(r0,"P");let cc=w.concl18(r0,"P");
 ok("P: I_А=60 <70, I_Ц=100 → R4 (min=60<70)",cc.txt.startsWith("R4")&&Math.abs(x.IA-60)<0.051);
 // Б 65 (Б3=80 ≥70 — иначе раньше сработает R5, как в Excel), А/В 100 → IC<70 → R3
 API.Q18.levels.P.q.forEach(q=>{r0.q.P[q.code]={a:(q.dir===1&&q.blk!=="Б3")?"3":(q.dir===1?"4":"5"),j:"",e:"ev"};});
 x=w.idx18(r0,"P");cc=w.concl18(r0,"P");
 ok("P: I_Ц=65 <70, Б3=80 → R3",cc.txt.startsWith("R3")&&Math.abs(x.IC-65)<0.051);
 // R6: индексы <85, min<70, альтернатива лучше
 API.Q18.levels.P.q.forEach(q=>{r0.q.P[q.code]={a:"3",j:"",e:"ev"};});
 r0.econ.alt={c:Array(9).fill("1"),d:Array(9).fill("5"),e:Array(9).fill(""),f:Array(9).fill("")};
 cc=w.concl18(r0,"P");
 ok("P: альтернатива предпочтительнее → R6 (минует R2 при min=60)",cc.txt.startsWith("R6")&&cc.txt.includes("предпочтительнее"));
 // R5: Б3 <70, min≥… состав: все 4 кроме блока Б3 → 2
 API.Q18.levels.P.q.forEach(q=>{r0.q.P[q.code]={a:"3",j:"",e:"ev"};});
 API.Q18.levels.P.q.filter(q=>q.blk==="Б3").forEach(q=>{r0.q.P[q.code]={a:"2",j:"обосн",e:"ev"};});
 r0.econ.alt={c:Array(9).fill(""),d:Array(9).fill(""),e:Array(9).fill(""),f:Array(9).fill("")};
 x=w.idx18(r0,"P");cc=w.concl18(r0,"P");
 ok("P: Б3 dk=40<70 → R5 (min=56.x<70, R6 не сработал)",cc.txt.startsWith("R5")&&x.blocks.filter(b=>b.code==="Б3")[0].dk===40);
 r0.lvl="S";
 // ---- 15. миграция 31→30
 const A31=["5","4","4","3","4","4","5","НП","4","3","4","3","3","4","3","3","4","3","1","4","4","3","4","4","4","4","5","4","3","3","4"];
 const r2={ans:A31.slice(),conf:Array(31).fill(""),com:Array(31).fill(""),stops:["","ДА","","",""],flags:Array(10).fill(""),props:[{targets:{4:"2",18:"4"}}],resp:null,audits:[]};
 ok("миграция: возвращает true",w.migrate31to18(r2));
 ok("миграция: legacy31 сохранён",r2.legacy31&&r2.legacy31.ans.length===31&&r2.legacy31.ans[0]==="5");
 ok("миграция: q.S заполнен",Object.keys(r2.q.S).length>0);
 ok("миграция: А1 = «5» (old0=5)",r2.q.S["А1"].a==="5");
 ok("миграция: А6 = «4» (old2=4, old5=4 → среднее 4)",r2.q.S["А6"].a==="4");
 ok("миграция: А3 = «НП» (old7=НП)",r2.q.S["А3"].a==="НП");
 ok("миграция: А10 = «3.5» (old9=3, old25=4 → среднее)",r2.q.S["А10"].a==="3.5");
 ok("миграция: Б3 = «4» (old22=4)",r2.q.S["Б3"].a==="4");
 ok("миграция: В1 = «4» (old10=4)",r2.q.S["В1"].a==="4");
 ok("миграция: Б6 = среднее(4,3,1)=2.67 → «2.5» (old13,17,18)",r2.q.S["Б6"].a==="2.5");
 ok("миграция: sf[2]=«подтверждено» (old stops1=ДА)",r2.sf[2]==="подтверждено"&&r2.sf[0]==="не выявлено");
 ok("миграция: props targets → коды (4→А7:2, 18→Б6:4)",r2.props[0].targets["А7"]==="2"&&r2.props[0].targets["Б6"]==="4"&&r2.props[0].targetsLegacy);
 ok("миграция: lvl='S'",r2.lvl==="S");
 ok("миграция: повторный вызов = false (идемпотентность)",w.migrate31to18(r2)===false);
 const xm=w.idx18(r2,"S");
 ok("миграция: индексы конечны (IA/IC/IR не null)",xm.IA!=null&&xm.IC!=null&&xm.IR!=null);
 // ---- 16. calc(r)-обёртка и ранние сигналы
 setAllJE(r0,"S",4);r0.lvl="S";
 let c=w.calc(r0);
 close("calc: I = комплексный (геом.) = 80",c.I,80);
 ok("calc: cat = «Приемлемый» (80 ∈ 70–84)",c.cat==="Приемлемый");
 ok("calc: bl = 3 направления, P = I_d/100",c.bl.length===3&&Math.abs(c.bl[0].P-0.8)<0.001);
 ok("calc: conc = заключение уровня S",c.conc.startsWith("II. Положительное с замечаниями"));
 ok("calc: covOk = true (100 ≥ 85)",c.covOk===true);
 r0.prev={IA:95,IC:80,IR:80};
 let ea=w.early18(r0);
 ok("early: ΔI_А = −15 → «Предупреждение»",ea.some(s2=>s2.sig.includes("Предупреждение")&&s2.rule.includes("I_А")));
 r0.prev={IA:null,IC:null,IR:null};ea=w.early18(r0);
 ok("early: без истории → только Coverage/SF сигналы (cov=100 → пусто)",ea.length===0);
 // ---- 17. общие правила заключения
 API.Q18.levels.S.q.forEach(q=>{r0.q.S[q.code]={a:"2",j:"обосн",e:"ev"};});
 x=w.idx18(r0,"S");let g=w.genRules18(r0,"S",x);
 ok("genRules: все индексы 40 <55 → 3 правила (cov=100)",g.length===3&&g.map(t=>t.rule).includes("I_Р < 55%"));
 ok("genRules: 40<55 → «IV.» (min=40)",w.concl18(r0,"S").txt.startsWith("IV."));
 // ---- 18. UI не сломан: вкладки рендерятся
 ok("вкладки на месте",d.querySelectorAll(".tabs button").length>=15);
 ok("версия в топбаре (пока v17 — обновится в UI-патче)",!!d.querySelector(".ver"));
}catch(e){console.log("EXCEPTION:",e&&e.stack||e);fails++;}
console.log("\nИТОГ SMOKE18: fails="+fails);
process.exit(fails?1:0);
},1500);
