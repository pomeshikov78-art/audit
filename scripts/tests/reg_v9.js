/* Сводная регрессия v1–v9: математические якоря, формы (v8), workflow-нейтральность (v9), ключевые UI-структуры.
   Запуск: node reg_v9.js <путь к html> */
const fs=require("fs");const {JSDOM}=require("jsdom");
const P=process.argv[2];
const dom=new JSDOM(fs.readFileSync(P,"utf8"),{runScripts:"dangerously",pretendToBeVisual:true,url:"http://localhost/"});
const w=dom.window,d=w.document;
w.HTMLCanvasElement.prototype.getContext=()=>({clearRect(){},fillRect(){},beginPath(){},moveTo(){},lineTo(){},stroke(){},fill(){},arc(){},save(){},restore(){},translate(){},rotate(){},fillText(){},setLineDash(){},closePath(){},quadraticCurveTo(){}});
w.URL.createObjectURL=()=>"blob:x";w.URL.revokeObjectURL=()=>{};
w.addEventListener("error",e=>{if(/Not implemented/.test(e.message||""))return;console.log("WINDOW-ERR:",e.message);});
let fails=0;
const ok=(n,c)=>{console.log((c?"PASS":"FAIL")+" | "+n);if(!c)fails++;};
const close=(a,b,e)=>Math.abs(a-b)<(e||0.05);
const A=["5","4","4","3","5","4","НП","3","4","4","3","3","4","3","3","4","3","1","5","4","4","3","4","4","3","4","НП"];
const blk3=A.slice();for(let i=18;i<27;i++)blk3[i]="НП";
const allNP=Array(27).fill("НП");
setTimeout(()=>{
try{
 const cc=w.calcCore;
 // ---- математические якоря (v1–v5) ----
 ok("якорь: база 30/35/35 → I=72.9",close(cc(A,[30,35,35]).I,72.9));
 ok("якорь: блок III все «Н/П» → 70.4 (ренорм)",close(cc(blk3,[30,35,35]).I,70.4));
 ok("якорь: веса 25/40/35 → 72.0",close(cc(A,[25,40,35]).I,72.0));
 ok("якорь: 25/40/35 + блок III «Н/П» → 69.1",close(cc(blk3,[25,40,35]).I,69.1));
 ok("якорь: все 27 «Н/П» → I=null («Недостаточно данных»)",cc(allNP,[30,35,35]).I===null);
 const ah=w.ahpStats({a12:"3",a13:"9",a23:"3"});
 ok("AHP 3/9/3: веса 69.2/23.1/7.7",close(ah.wPct[0],69.2,0.1)&&close(ah.wPct[1],23.1,0.1)&&close(ah.wPct[2],7.7,0.1));
 ok("AHP 3/9/3: CR=0",close(ah.CR,0,1e-9));
 ok("AHP 3/9/3: I=75.7",close(cc(A,ah.wPct).I,75.7));
 const ahBad=w.ahpStats({a12:"9",a13:"1",a23:"9"});
 ok("AHP несогласованная: CR≥0.1",ahBad.CR>=0.1);
 // ---- структуры v8 (формы) ----
 ok("FORMS: 6 форм",w.eval("FORMS.length")===6);
 ok("avp: 13 фокус / 7 на",w.eval("FORMS[0].focus.length")===13&&w.eval("FORMS[0].na.length")===7);
 const S=w.niokrAPI.state;
 ok("демо: form=okr, I не зависит от формы",S.list[0].form==="okr"&&(S.list[0].ans=A.slice(),close(w.calc(S.list[0]).I,72.9)));
 ok("чек-лист: панель профиля ОКР",d.getElementById("tab-check").innerHTML.includes("Профиль формы «ОКР»"));
 // ---- v9: workflow выключен → нейтральность ----
 ok("flow выключен по умолчанию (v8-поведение)",S.list[0].flow===null||S.list[0].flow.enabled===false);
 ok("15 вкладок + «Маршрут»/«БД»/«Графики»/«Предложения»",d.querySelectorAll(".tabs button").length===15&&!!d.getElementById("tab-flow")&&!!d.getElementById("tab-resp")&&!!d.getElementById("tab-db")&&!!d.getElementById("tab-prop")&&!!d.getElementById("tab-viz"));
 // ---- портфель: обе колонки ----
 const hp=d.getElementById("tab-port").innerHTML;
 ok("портфель: колонки Форма и Маршрут",hp.includes("<th>Форма</th>")&&hp.includes("<th>Маршрут</th>"));
 // ---- evidStats (v6) ----
 const r2=w.newRnd("Регресс","Р-01");r2.ans=Array(27).fill("5");
 let ev=w.evidStats(r2);
 ok("evidStats: 27 высоких без документов → K=0, viol=27",ev.K===0&&ev.viol.length===27);
 r2.docs.push({q:0,name:"Протокол",type:w.eval("DOCTYPES[0]"),date:"01.01.2026",status:"provided",note:""});
 ev=w.evidStats(r2);
 ok("evidStats: документ «Предоставлен» → K>0, viol=26",ev.K>0&&ev.viol.length===26);
 // ---- презентация (v7) ----
 const sl=w.buildSlides(S.list[0]);
 ok("презентация: 12 слайдов",sl.length===12);
 ok("слайд 9: расчёт I",sl[9].h.includes("72.9")||sl[9].h.includes("Расчёт"));
 ok("дек: автономный HTML",w.buildDeckHTML().includes("<!DOCTYPE html>"));
 ok("discrepancies: массив",Array.isArray(w.discrepancies(S.list[0])));
}catch(e){console.log("ERR:",e.stack||e.message);fails++;}
console.log(fails?(">>> ПРОВАЛЕНО: "+fails):">>> РЕГРЕССИЯ v1–v9 ЗЕЛЁНАЯ");
process.exit(fails?1:0);
},400);
