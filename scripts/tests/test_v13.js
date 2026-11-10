const fs=require("fs");
const {JSDOM}=require("jsdom");
const P="/tmp/arena/tenants/71ebd7cc-b365-41c5-863c-79f587b1df1f/sessions/76eaa8bd-b9c6-43e4-8bfe-0b8d30b3272e/workspace/output/Аудит_НИОКР_интерактивная_модель.html";
const dom=new JSDOM(fs.readFileSync(P,"utf8"),{runScripts:"dangerously",pretendToBeVisual:true,url:"http://localhost/"});
const w=dom.window,d=w.document;
w.HTMLCanvasElement.prototype.getContext=()=>({clearRect(){},fillRect(){},beginPath(){},moveTo(){},lineTo(){},stroke(){},fill(){},arc(){},save(){},restore(){},translate(){},rotate(){},fillText(){},setLineDash(){},closePath(){},quadraticCurveTo(){}});
w.URL.createObjectURL=()=>"blob:x";w.URL.revokeObjectURL=()=>{};
w.addEventListener("error",e=>{if(/Not implemented/.test(e.message||""))return;console.log("WINDOW-ERR:",e.message);});
let fails=0;
const ok=(n,c)=>{console.log((c?"PASS":"FAIL")+" | "+n);if(!c)fails++;};
const q=s=>d.querySelector(s),qa=s=>[...d.querySelectorAll(s)];
const key=k=>d.dispatchEvent(new w.KeyboardEvent("keydown",{key:k,bubbles:true}));
const A=["5","4","4","3","5","4","НП","3","4","4","3","3","4","3","3","4","3","1","5","4","4","3","4","4","3","4","НП"];
const chg=(id,v)=>{const e=d.getElementById(id);e.value=v;e.dispatchEvent(new w.Event("change",{bubbles:true}));};
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state;
 // ---- присутствие ----
 ok("v13.0: topbar",q(".ver").textContent.includes("v14.0"));
 ok("15 вкладок",qa(".tabs button").length===15);
 ok("вкладка Графики БД",!!q('.tabs button[data-t="viz"]')&&!!d.getElementById("tab-viz"));
 ok("TOUR: 15 шагов",w.eval("TOUR.length")===15);
 ok("TOUR содержит шаг viz",w.eval("TOUR.some(t=>t[0]==='viz')"));
 ok("API: 9 функций визуализации",["vizFiltered","donutSVG","barsSVG","radarSVG","vizTimelineSVG","heatHTML","buildVizCharts","renderViz","printViz"].every(k=>typeof API[k]==="function"));
 // ---- хоткеи ----
 key("v");ok("хоткей v → Графики БД",q("section.tab.on").id==="tab-viz");
 key("м");ok("хоткей м → Графики БД",q("section.tab.on").id==="tab-viz");
 key("m");ok("хоткей m → Предложения (не сдвинут)",q("section.tab.on").id==="tab-prop");
 key("b");ok("хоткей b → База (не сдвинут)",q("section.tab.on").id==="tab-db");
 key("9");ok("хоткей 9 → Справочник (не сдвинут)",q("section.tab.on").id==="tab-ref");
 key("0");ok("хоткей 0 → Акт (не сдвинут)",q("section.tab.on").id==="tab-rep");
 key("p");ok("хоткей p → Презентация (не сдвинут)",q("section.tab.on").id==="tab-pres");
 // ---- данные: 3 НИОКР + аудиты ----
 const r0=S.list[0];
 r0.ans=A.slice();r0.conf=Array(27).fill("");r0.com=Array(27).fill("");
 r0.shifr="НИОКР-01";r0.name="Разработка А";r0.podrazd="КБ-1";r0.form="okr";r0.status="Выполняется";
 r0.fin.plan="100";r0.fin.fact="50";
 ok("якорь I=72.9 (r0)",Math.abs(w.calc(r0).I-72.9)<0.05);
 const r1=w.eval("newRnd('Разработка Б','НИОКР-02')");
 r1.ans=Array(27).fill("1");r1.conf=Array(27).fill("");r1.com=Array(27).fill("");
 r1.podrazd="КБ-2";r1.form="nir";r1.fin.plan="200";
 const r2=w.eval("newRnd('Разработка В','НИОКР-03')");
 r2.ans=Array(27).fill("5");r2.conf=Array(27).fill("");r2.com=Array(27).fill("");
 r2.podrazd="КБ-1";r2.form="avp";r2.status="Завершена";
 S.list.push(r1,r2);
 S.me="Тестов Т.Т.";
 w.saveAudit(r0,"первый");            // I=72.9
 r0.ans=Array(27).fill("5");w.saveAudit(r0,"после мероприятий"); // I=100
 w.saveAudit(r1,"проблемная");        // I=20
 r0.ans=A.slice();                    // вернуть 72.9
 w.renderAll();
 ok("r1 Критическая (I=20)",Math.abs(w.calc(r1).I-20)<0.05&&w.calc(r1).cat==="Критическая");
 ok("r2 Высокая (I=100)",Math.abs(w.calc(r2).I-100)<0.05&&w.calc(r2).cat==="Высокая");
 // ---- рендер вкладки ----
 w.goTab("viz");
 let h=d.getElementById("tab-viz").innerHTML;
 ok("KPI: выборка 3",h.includes("НИОКР В ВЫБОРКЕ")&&h.includes("оценено: 3 из 3"));
 ok("KPI: средний I = 64.3",h.includes("64.3")); // (72.9+20+100)/3
 ok("KPI: Σ план 300.0 / факт 50.0",h.includes("300.0")&&h.includes("50.0 млн ₽ факт"));
 ok("KPI: аудитов 3",h.includes("АУДИТОВ / ПРЕДЛОЖЕНИЙ ПРИНЯТО")&&h.includes("3 / 0"));
 // ---- диаграмма I ----
 ok("бары: 3 строки data-vopen",qa('#tab-viz [data-vopen]').length===3);
 ok("бары: пороги 85/70/50",h.includes("85 — высокая")&&h.includes("70 — хорошая")&&h.includes("50 — удовл."));
 ok("бары: значения I",h.includes("72.9")&&h.includes("100.0")&&h.includes("20.0"));
 // ---- донаты ----
 ok("донат категорий",h.includes("Категории состояния")&&h.includes("Хорошая")&&h.includes("Критическая")&&h.includes("Высокая"));
 ok("донат форм",h.includes("Формы проведения НИОКР")&&h.includes("ОКР")&&h.includes("НИР")&&h.includes("Аванпроект"));
 ok("donutSVG: пустой набор → заглушка",API.donutSVG([],"тест").includes("Нет данных"));
 // ---- радар ----
 ok("радар: заголовок и оси",h.includes("Профили блоков (радар Pⱼ)")&&h.includes("Актуальность P₁")&&h.includes("Эффективность P₂")&&h.includes("Целесообразность P₃"));
 ok("радар: 2 полигона (выборка + текущая)",(h.match(/<polygon points=/g)||[]).length>=2+4);
 ok("радар: подпись текущей НИОКР",h.includes("синий — НИОКР-01"));
 // ---- динамика ----
 ok("динамика: 3 точки аудитов",h.includes("Динамика I по зафиксированным аудитам"));
 const tl=API.vizTimelineSVG();
 ok("динамика: 3 circle",((tl.match(/<circle/g)||[]).length)===3);
 ok("динамика: polyline для НИОКР с 2 аудитами",(tl.match(/<polyline/g)||[]).length===1);
 ok("динамика: легенда с числом записей",tl.includes("НИОКР-01 (2)")&&tl.includes("НИОКР-02 (1)"));
 // ---- тепловая карта ----
 ok("тепловая карта: заголовок",h.includes("Тепловая карта «НИОКР × блоки»"));
 const hh=API.heatHTML([0,1,2]);
 ok("тепловая: 3 строки и заголовки блоков",hh.includes("I · Актуальность")&&hh.includes("II · Эффективность")&&(hh.match(/НИОКР-0/g)||[]).length>=3);
 ok("тепловая: r1 — 20% во всех блоках",(hh.match(/>20%</g)||[]).length===3);
 ok("тепловая: r2 — 100%",(hh.match(/>100%</g)||[]).length===3);
 S.list.push(w.eval("newRnd('Пустая','НИОКР-04')"));
 ok("тепловая: Н/П при пустом блоке",API.heatHTML([S.list.length-1]).includes("Н/П"));
 S.list.pop();
 // ---- клик по бару открывает НИОКР ----
 q('#tab-viz [data-vopen="2"]').dispatchEvent(new w.MouseEvent("click",{bubbles:true}));
 ok("клик по бару: cur=2 и переход на дашборд",S.cur===2&&q("section.tab.on").id==="tab-dash");
 S.cur=0;w.renderAll();w.goTab("viz");
 // ---- фильтры ----
 ok("фильтры: контролы на месте",["vizForm","vizStatus","vizPodr","vizCat","vizReset","vizPrint"].every(id=>!!d.getElementById(id)));
 chg("vizForm","nir");
 h=d.getElementById("tab-viz").innerHTML;
 ok("фильтр формы nir: 1 из 3",h.includes("показано 1 из 3")&&qa('#tab-viz [data-vopen]').length===1);
 ok("фильтр формы: в барах только НИОКР-02",h.includes("НИОКР-02")&&qa('#tab-viz [data-vopen]')[0].dataset.vopen==="1");
 ok("радар: среднее по выборке (1)",d.getElementById("tab-viz").innerHTML.includes("среднее по выборке (1)"));
 d.getElementById("vizReset").click();
 h=d.getElementById("tab-viz").innerHTML;
 ok("сброс фильтров: снова 3",h.includes("показано 3 из 3"));
 chg("vizCat","Критическая");
 ok("фильтр категории: 1 (НИОКР-02)",d.getElementById("tab-viz").innerHTML.includes("показано 1 из 3")&&qa('#tab-viz [data-vopen]').length===1);
 d.getElementById("vizReset").click();
 chg("vizPodr","КБ-1");
 ok("фильтр подразделения КБ-1: 2",d.getElementById("tab-viz").innerHTML.includes("показано 2 из 3"));
 d.getElementById("vizReset").click();
 chg("vizStatus","Завершена");
 ok("фильтр статуса: 1 (НИОКР-03)",d.getElementById("tab-viz").innerHTML.includes("показано 1 из 3"));
 d.getElementById("vizReset").click();
 ok("фильтры не влияют на выбор данных БД",API.vizFiltered().length===3);
 // ---- печать ----
 d.getElementById("vizPrint").click();
 const pa=d.getElementById("printarea").innerHTML;
 ok("печать: заголовок и выборка",pa.includes("ГРАФИЧЕСКИЙ ДАШБОРД БАЗЫ НИОКР")&&pa.includes("Выборка: 3 из 3")&&pa.includes("без фильтров"));
 ok("печать: диаграммы без контролов",pa.includes("<svg")&&pa.includes("Тепловая карта")&&!pa.includes('id="vizForm"'));
 ok("печать: футер v14.0",pa.includes("v14.0"));
 // ---- сессия: фильтры не сохраняются ----
 ok("save() не включает viz",!w.eval("localStorage.getItem(LSKEY)").includes('"viz"'));
 // ---- пустая база ----
 const keep=S.list.slice();S.list=[w.eval("newRnd('Единственная','НИОКР-05')")];S.cur=0;
 API.renderViz();
 h=d.getElementById("tab-viz").innerHTML;
 ok("пустая база: KPI Н/П и заглушки",h.includes("Н/П — недостаточно данных")&&h.includes("оценено: 0 из 1"));
 ok("пустая база: таймлайн-заглушка",h.includes("Нет зафиксированных аудитов"));
 S.list=keep;S.cur=0;w.renderAll();
}catch(e){fails++;console.log("EXCEPTION:",e.stack||e);}
console.log(fails?(">>> ПРОВАЛЕНО: "+fails):">>> ВСЕ ТЕСТЫ v13 ПРОЙДЕНЫ");
process.exit(fails?1:0);
},400);
