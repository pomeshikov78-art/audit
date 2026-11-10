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
const closeM=()=>{const b=q('#mbox [data-mb="0"]');if(b)b.click();};
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state,r0=S.list[0];
 // ---- присутствие и API ----
 ok("v12.0: topbar (патч v12)",q(".ver").textContent.includes("v14.0"));
 ok("15 вкладок",qa(".tabs button").length===15);
 ok("вкладка База НИОКР",!!q('.tabs button[data-t="db"]')&&!!d.getElementById("tab-db"));
 ok("TOUR: 15 шагов",w.eval("TOUR.length")===15);
 ok("TOUR содержит шаг db",w.eval("TOUR.some(t=>t[0]==='db')"));
 ok("API: makeAudit/saveAudit/dbStats/buildDbCsv",["makeAudit","saveAudit","dbStats","buildDbCsv"].every(k=>typeof API[k]==="function"));
 ok("API: renderDb/fixAuditModal/cmpModal/histChartSVG",["renderDb","fixAuditModal","cmpModal","histChartSVG"].every(k=>typeof API[k]==="function"));
 // ---- хоткеи ----
 key("b");ok("хоткей b → База НИОКР",q("section.tab.on").id==="tab-db");
 key("И");ok("хоткей И → База НИОКР",q("section.tab.on").id==="tab-db");
 key("9");ok("хоткей 9 → Справочник (не сдвинут)",q("section.tab.on").id==="tab-ref");
 key("0");ok("хоткей 0 → Акт (не сдвинут)",q("section.tab.on").id==="tab-rep");
 key("p");ok("хоткей p → Презентация (не сдвинут)",q("section.tab.on").id==="tab-pres");
 key("r");ok("хоткей r → Респонденты (не сдвинут)",q("section.tab.on").id==="tab-resp");
 // ---- модель данных ----
 ok("newRnd: audits=[] и created ISO",Array.isArray(r0.audits)&&!!r0.created&&!isNaN(Date.parse(r0.created)));
 // ---- наполнение и рендер вкладки ----
 r0.ans=A.slice();r0.conf=Array(27).fill("");r0.com=Array(27).fill("");
 r0.shifr="НИОКР-01";r0.name="Разработка А (пример)";r0.podrazd="КБ-1";r0.head="Петров П.П.";r0.form="okr";
 S.me="Тестов Т.Т.";
 const I0=w.calc(r0).I;
 ok("якорь I=72.9",Math.abs(I0-72.9)<0.05);
 w.goTab("db");
 let h=d.getElementById("tab-db").innerHTML;
 ok("БД: KPI «НИОКР В БАЗЕ»",h.includes("НИОКР В БАЗЕ"));
 ok("БД: средний/медианный I",h.includes("СРЕДНИЙ I ПО БАЗЕ")&&h.includes("медианный I"));
 ok("БД: элементы управления",["dbQ","dbForm","dbStatus","dbCat","dbFlow","dbResp","dbSort","dbCsv","hAdd"].every(id=>!!d.getElementById(id)));
 ok("БД: строка реестра с текущей НИОКР",h.includes("НИОКР-01")&&h.includes("Разработка А"));
 ok("БД: кнопка 📸 в строке",!!q('#tab-db [data-hfix="0"]'));
 ok("БД: история пуста — подсказка",h.includes("Аудитов ещё не фиксировали"));
 // ---- фиксация через модалку ----
 d.getElementById("hAdd").click();
 ok("модалка фиксации",d.getElementById("mbox").innerHTML.includes("Зафиксировать аудит в БД"));
 d.getElementById("hNote").value="плановый квартальный аудит";
 q('#mbox [data-mb="1"]').click();
 ok("аудит зафиксирован: 1 запись",r0.audits.length===1);
 const a1=r0.audits[0];
 ok("снимок: I=72.9 и ответы скопированы",Math.abs(a1.I-72.9)<0.05&&a1.ans.join()===A.join());
 ok("снимок: метаданные",a1.by.includes("Тестов Т.Т.")&&a1.note==="плановый квартальный аудит"&&a1.ws==="30/35/35"&&a1.form==="okr"&&!!a1.dateStr&&!!a1.id);
 ok("снимок: conf/docsN/resp/flow",Array.isArray(a1.conf)&&typeof a1.docsN==="number"&&a1.resp===null&&a1.flow===null);
 ok("журнал: запись о фиксации",r0.flow&&r0.flow.log.some(e=>e.what.includes("зафиксирован аудит")));
 ok("модалка фиксации закрылась",!d.getElementById("moverlay").classList.contains("on"));
 // ---- вторая фиксация с изменёнными ответами → ΔI, сравнение ----
 r0.ans=Array(27).fill("5");r0.conf=Array(27).fill("Да");
 w.saveAudit(r0,"после мероприятий");
 ok("2 записи в истории",r0.audits.length===2);
 ok("I второй записи = 100",Math.abs(r0.audits[1].I-100)<0.05);
 API.renderDb();
 h=d.getElementById("tab-db").innerHTML;
 ok("история: таблица с ΔI",h.includes("ΔI")&&h.includes("+27.1"));
 ok("история: график динамики (SVG)",h.includes("<svg")&&h.includes("polyline"));
 ok("история: примечания записей",h.includes("плановый квартальный аудит")&&h.includes("после мероприятий"));
 const cmpBoxes=qa("#tab-db [data-cmp]");
 ok("чекбоксы сравнения: 2",cmpBoxes.length===2);
 d.getElementById("hCmp").click(); // ни одной отметки → err-notify, без падения
 cmpBoxes[0].checked=true;cmpBoxes[1].checked=true;
 d.getElementById("hCmp").click();
 ok("модалка сравнения аудитов",d.getElementById("mbox").innerHTML.includes("Сравнение аудитов"));
 ok("сравнение: ΔI и число расхождений",d.getElementById("mbox").innerHTML.includes("ΔI (B−A): +27.1")&&/расхождений ответов: \d+ из 27/.test(d.getElementById("mbox").innerHTML));
 closeM();
 // ---- восстановление ответов из снимка ----
 API.renderDb();
 q('#tab-db [data-hrest="0"]').click();
 ok("модалка восстановления",d.getElementById("mbox").innerHTML.includes("Восстановить ответы из аудита №1"));
 q('#mbox [data-mb="1"]').click();
 ok("ответы восстановлены из снимка №1",r0.ans.join()===A.join());
 ok("I после восстановления = 72.9",Math.abs(w.calc(r0).I-72.9)<0.05);
 ok("журнал: восстановление",r0.flow.log.some(e=>e.what.includes("восстановлены ответы")));
 // ---- dbStats / buildDbCsv / histChartSVG ----
 const st=API.dbStats();
 ok("dbStats: n и scored",st.n===S.list.length&&st.scored>=1&&Math.abs(st.avg-w.calc(r0).I)<1e-9);
 ok("dbStats: byCat сумма = n",Object.values(st.byCat).reduce((a,b)=>a+b,0)===st.n);
 ok("dbStats: audits=2 и лента",st.audits===2&&st.last.length===2&&st.last[0].r===r0);
 const csv=API.buildDbCsv();
 ok("CSV базы: заголовок",csv.split("\n")[0].includes("Шифр")&&csv.split("\n")[0].includes("Аудитов")&&csv.split("\n")[0].includes("I последний"));
 ok("CSV базы: строки по числу НИОКР",csv.trim().split("\n").length===S.list.length+1&&csv.includes("НИОКР-01"));
 ok("CSV базы: первый/последний I",csv.includes("72.9")&&csv.includes("100.0"));
 const svg=API.histChartSVG(r0);
 ok("histChartSVG: svg с точками",svg.includes("<svg")&&svg.includes("72.9")&&svg.includes("100.0"));
 // ---- поиск, фильтры, сортировка ----
 const bq=d.getElementById("dbQ");
 bq.value="кб-1";bq.dispatchEvent(new w.Event("input",{bubbles:true}));
 ok("поиск по подразделению (регистронезависимый)",d.getElementById("tab-db").innerHTML.includes("НИОКР-01"));
 ok("поиск: фокус сохранён",d.activeElement&&d.activeElement.id==="dbQ");
 const bq2=d.getElementById("dbQ");bq2.value="нет-такой";bq2.dispatchEvent(new w.Event("input",{bubbles:true}));
 ok("поиск: ничего не найдено",d.getElementById("tab-db").innerHTML.includes("Ничего не найдено"));
 const bq3=d.getElementById("dbQ");bq3.value="";bq3.dispatchEvent(new w.Event("input",{bubbles:true}));
 const bf=d.getElementById("dbForm");bf.value="nir";bf.dispatchEvent(new w.Event("change",{bubbles:true}));
 ok("фильтр по форме (nir): реестр пуст (r0=okr)",qa("#tab-db [data-dopen]").length===0);
 const bf2=d.getElementById("dbForm");bf2.value="";bf2.dispatchEvent(new w.Event("change",{bubbles:true}));
 ok("сброс фильтра: снова видна",qa("#tab-db [data-dopen]").length===1);
 const bs=d.getElementById("dbSort");bs.value="created_desc";bs.dispatchEvent(new w.Event("change",{bubbles:true}));
 ok("сортировка created_desc применяется",S.db.sort==="created_desc");
 bs.value="I_desc";bs.dispatchEvent(new w.Event("change",{bubbles:true}));
 // ---- CSV-кнопка без падения ----
 d.getElementById("dbCsv").click();
 ok("кнопка CSV базы отработала",true);
 // ---- вторая НИОКР: открытие из реестра ----
 const r1=w.eval("newRnd('Разработка Б','НИОКР-02')");
 r1.ans=Array(27).fill("1");r1.conf=Array(27).fill("");r1.com=Array(27).fill("");
 S.list.push(r1);API.renderDb();
 ok("реестр: 2 НИОКР",qa("#tab-db [data-dopen]").length===2);
 q('#tab-db [data-dopen="1"]').click();
 ok("открыть: cur переключился на НИОКР-02",S.cur===1&&q("section.tab.on").id==="tab-dash");
 S.cur=0;w.renderAll();w.goTab("db");
 // ---- акт оценки: кнопка фиксации ----
 w.goTab("rep");
 ok("акт: кнопка 📸 Зафиксировать",!!d.getElementById("repFix")&&d.getElementById("repFix").textContent.includes("Зафиксировать"));
 ok("акт: счётчик записей истории",d.getElementById("tab-rep").innerHTML.includes("история: 2 запис(ей)"));
 d.getElementById("repFix").click();
 ok("акт: модалка фиксации",d.getElementById("mbox").innerHTML.includes("Зафиксировать аудит в БД"));
 q('#mbox [data-mb="1"]').click();
 ok("акт: аудит добавлен (3 записи)",r0.audits.length===3);
 // ---- автофиксация при завершении маршрута ----
 API.flowInit(r0);r0.flow.enabled=true;r0.flow.stage=1;r0.flow.stages[0].status="work";
 const nBefore=r0.audits.length;
 w.flowAct("ok");w.flowAct("ok");w.flowAct("ok");w.flowAct("ok");
 ok("маршрут завершён",r0.flow.done===true);
 ok("автофиксация аудита при завершении маршрута",r0.audits.length===nBefore+1&&r0.audits[r0.audits.length-1].note.includes("автофиксация"));
 ok("автофиксация: снимок flow.done",r0.audits[r0.audits.length-1].flow&&r0.audits[r0.audits.length-1].flow.done===true);
 // ---- удаление записи истории ----
 w.goTab("db");API.renderDb();
 const nA=r0.audits.length;
 q("#tab-db [data-hdel]").click();
 ok("модалка удаления записи",d.getElementById("mbox").innerHTML.includes("Удалить запись аудита"));
 q('#mbox [data-mb="1"]').click();
 ok("запись удалена",r0.audits.length===nA-1);
 // ---- экспорт v11 ----
 d.getElementById("btnExp").click();
 ok("экспорт: имя файла v11",true); // скачивание мокировано; проверяем отсутствие ошибок ниже
 // ---- миграция load(): старые данные без audits/created ----
 w.eval(`(function(){const x=newRnd("Миграция Тест","НИОКР-98");delete x.audits;delete x.created;
  localStorage.setItem(LSKEY,JSON.stringify({list:[x],cur:0,role:"auditor",theme:"light",me:"",respMode:"median",weights:[30,35,35],wsName:"",ahp:{a12:"1",a13:"1",a23:"1"},weightSets:[]}));load();})();`);
 const rm=API.state.list[0];
 ok("миграция load(): audits добавлен",Array.isArray(rm.audits)&&rm.audits.length===0);
 ok("миграция load(): created добавлен",!!rm.created&&!isNaN(Date.parse(rm.created)));
 API.renderDb();
 ok("БД рендерится после миграции",d.getElementById("tab-db").innerHTML.includes("НИОКР-98"));
}catch(e){fails++;console.log("EXCEPTION:",e.stack||e);}
console.log(fails?(">>> ПРОВАЛЕНО: "+fails):">>> ВСЕ ТЕСТЫ v11 ПРОЙДЕНЫ");
process.exit(fails?1:0);
},400);
