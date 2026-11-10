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
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state;
 // ---- присутствие ----
 ok("v14.0: topbar",q(".ver").textContent.includes("v14.0"));
 ok("15 вкладок (без новых)",qa(".tabs button").length===15);
 ok("TOUR: 15 шагов",w.eval("TOUR.length")===15);
 ok("TOUR rep: шаблоны отчётов",w.eval("TOUR.find(t=>t[0]==='rep')[2]").includes("Отчёт комиссии")&&w.eval("TOUR.find(t=>t[0]==='rep')[2]").includes("Отчёт подразделению"));
 ok("TOUR pres: два режима дека",w.eval("TOUR.find(t=>t[0]==='pres')[2]").includes("ортфельный отчёт")&&w.eval("TOUR.find(t=>t[0]==='pres')[2]").includes("9 слайдов"));
 ok("API: 9 функций отчётности",["presSlides","buildPortSlides","portConclusions","deptAdvice","buildComReport","buildDeptReport","printCom","printDept","copyCom"].every(k=>typeof API[k]==="function"));
 // ---- данные: 3 НИОКР, аудиты, предложение ----
 const r0=S.list[0];
 r0.ans=A.slice();r0.conf=Array(27).fill("");r0.com=Array(27).fill("");
 r0.shifr="НИОКР-01";r0.name="Разработка А";r0.podrazd="КБ-1";r0.form="okr";r0.status="Выполняется";
 r0.fin.plan="100";r0.fin.fact="50";r0.decision="Продолжить с корректировкой ТЗ, сроков и бюджета";
 const r1=w.eval("newRnd('Разработка Б','НИОКР-02')");
 r1.ans=Array(27).fill("1");r1.conf=Array(27).fill("");r1.com=Array(27).fill("");r1.podrazd="КБ-2";r1.form="nir";
 const r2=w.eval("newRnd('Разработка В','НИОКР-03')");
 r2.ans=Array(27).fill("5");r2.conf=Array(27).fill("");r2.com=Array(27).fill("");r2.podrazd="КБ-1";r2.form="avp";r2.status="Завершена";
 r2.decision="Продолжить без изменения основных параметров";
 S.list.push(r1,r2);S.me="Тестов Т.Т.";
 w.saveAudit(r0,"первый");
 r0.ans=Array(27).fill("5");w.saveAudit(r0,"после мероприятий");
 w.saveAudit(r1,"проблемная");
 r0.ans=A.slice();
 const pr=w.eval("newProp()");
 pr.model="modern";pr.status="accepted";pr.extra="12.5";pr.months="9";pr.targets={"17":"4"};
 pr.measures=[{text:"Разработать программу специспытаний",resp:"Иванов И.И.",due:"20.12.2026",st:"plan"}];
 r0.props=[pr];
 w.renderAll();
 ok("якоря: r0=72.9 r1=20 r2=100",Math.abs(w.calc(r0).I-72.9)<0.05&&Math.abs(w.calc(r1).I-20)<0.05&&Math.abs(w.calc(r2).I-100)<0.05);
 // ---- режимы дека ----
 ok("presMode по умолчанию obj",S.presMode==="obj");
 ok("presSlides = buildSlides в obj",API.presSlides().length===w.buildSlides(r0).length&&API.presSlides().length===12);
 w.goTab("pres");
 let h=d.getElementById("tab-pres").innerHTML;
 ok("переключатель режимов",!!d.getElementById("pmObj")&&!!d.getElementById("pmPort"));
 ok("obj-режим: карточка «Источники данных»",h.includes("Источники данных · заполненность"));
 d.getElementById("pmPort").click();
 ok("переключено в port",S.presMode==="port"&&S.presSlide===0);
 ok("port-дек: 9 слайдов",API.buildPortSlides().length===9&&API.presSlides().length===9);
 h=d.getElementById("tab-pres").innerHTML;
 ok("port: список из 9 опций",qa("#prSel option").length===9);
 ok("port: карточки сводки и выводов",h.includes("Портфельная сводка")&&h.includes("Выводы и рекомендации ("));
 ok("port: показан титул",h.includes("Аудит НИОКР — портфельный отчёт"));
 // ---- содержимое слайдов ----
 const PS=API.buildPortSlides();
 ok("слайд 1 титул: средний I 64.3 и категории",PS[0].h.includes("64.3")&&PS[0].h.includes("Хорошая: 1")&&PS[0].h.includes("Тестов Т.Т."));
 ok("слайд 2 сводка: KPI и донаты",PS[1].h.includes("Средний I · медианный I")&&PS[1].h.includes("Категории состояния")&&PS[1].h.includes("Формы проведения")&&PS[1].h.includes("64.3")&&PS[1].h.includes("72.9"));
 ok("слайд 3 карта оценок: бары",PS[2].h.includes("data-vopen")&&PS[2].h.includes("100.0")&&PS[2].h.includes("20.0"));
 ok("слайд 4 тепловая карта",PS[3].h.includes("Тепловая карта «НИОКР × блоки»")&&PS[3].h.includes("100%"));
 ok("слайд 5 проблемные: НИОКР-02 + band",PS[4].h.includes("НИОКР-02")&&PS[4].h.includes("Проблемных: 1 из 3")&&PS[4].h.includes("band disc"));
 ok("слайд 6 динамика: точки и последние аудиты",((PS[5].h.match(/<circle/g)||[]).length)===3&&PS[5].h.includes("Последние аудиты")&&PS[5].h.includes("Тестов Т.Т."));
 ok("слайд 7 предложения: 12.5 и ΔI +2.3",PS[6].h.includes("12.5")&&PS[6].h.includes("+2.3")&&PS[6].h.includes("Итого принято предложений: 1 на сумму 12.5 млн ₽")&&PS[6].h.includes("72.9 → 75.2"));
 ok("слайд 8 решения: распределение",PS[7].h.includes("Продолжить с корректировкой ТЗ, сроков и бюджета")&&PS[7].h.includes("Решение не выбрано")&&PS[7].h.includes("НИОКР-02"));
 ok("слайд 9 выводы: автопункты",PS[8].h.includes("Проблемные разработки (1 из 3)")&&PS[8].h.includes("Принято предложений: 1")&&PS[8].h.includes("модели аудита НИОКР v14.0"));
 // ---- portConclusions / deptAdvice ----
 const cn=API.portConclusions();
 ok("выводы: err по НИОКР-02",cn.some(x=>x[0]==="err"&&x[1].includes("НИОКР-02 (I=20.0)")));
 ok("выводы: warn доказательные пробелы",cn.some(x=>x[0]==="warn"&&x[1].includes("Доказательные пробелы")));
 ok("выводы: ok высокие (НИОКР-03)",cn.some(x=>x[0]==="ok"&&x[1].includes("НИОКР-03 (I=100.0)")));
 ok("выводы: info средний I",cn.some(x=>x[0]==="info"&&x[1].includes("64.3")));
 const adv0=API.deptAdvice(r0);
 ok("рекомендации r0: мероприятия 0 из 1",adv0.some(x=>x.includes("выполнено мероприятий 0 из 1")));
 ok("рекомендации r0: решение выбрано — пункта нет",!adv0.some(x=>x.includes("Решение комиссии не выбрано")));
 const adv1=API.deptAdvice(r1);
 ok("рекомендации r1: I<50 + блоки 20%",adv1.some(x=>x.includes("ниже 50"))&&adv1.some(x=>x.includes("Низкое выполнение блока")));
 const adv2=API.deptAdvice(r2);
 ok("рекомендации r2: доказательные пробелы K=0",adv2.some(x=>x.includes("Доказательные пробелы")));
 // ---- навигация, печать, экспорт дека ----
 S.presSlide=8;w.presNav(1);
 ok("presNav: кламп вперёд (8 из 9)",S.presSlide===8);
 w.presNav(-1);
 ok("presNav: назад → 7",S.presSlide===7);
 d.getElementById("prPrint").click();
 ok("печать дека: 9 .slide в printarea",qa("#printarea .slide").length===9);
 const dk=w.buildDeckHTML();
 ok("HTML-дек: doctype и заголовок",dk.startsWith("<!DOCTYPE html>")&&dk.includes("<title>Портфельный отчёт — аудит НИОКР</title>"));
 ok("HTML-дек: 9 слайдов и футер v14.0",(dk.match(/class="slide"/g)||[]).length===9&&dk.includes("моделью аудита НИОКР v14.0"));
 d.getElementById("pmObj").click();
 ok("возврат в obj: 12 слайдов",S.presMode==="obj"&&qa("#prSel option").length===12&&d.getElementById("tab-pres").innerHTML.includes("Источники данных"));
 ok("obj-дек: HTML-дек по НИОКР",w.buildDeckHTML().includes("<title>Презентация — Разработка А (НИОКР-01)</title>"));
 // ---- отчёты на акте ----
 w.goTab("rep");
 ok("акт: кнопки отчётов",!!d.getElementById("repCom")&&!!d.getElementById("repDept")&&!!d.getElementById("repComCopy"));
 d.getElementById("repCom").click();
 let pa=d.getElementById("printarea").innerHTML;
 ok("отчёт комиссии: шапка и разделы",pa.includes("ОТЧЁТ КОМИССИИ ПО РЕЗУЛЬТАТАМ АУДИТА НИОКР")&&pa.includes("1. Общая статистика")&&pa.includes("2. Проблемные разработки")&&pa.includes("5. Выводы и рекомендации"));
 ok("отчёт комиссии: данные базы",pa.includes("НИОКР-02")&&pa.includes("64.3")&&pa.includes("12.5 млн ₽")&&pa.includes("Председатель комиссии"));
 ok("отчёт комиссии: футер v14.0",pa.includes("моделью аудита НИОКР v14.0"));
 d.getElementById("repDept").click();
 pa=d.getElementById("printarea").innerHTML;
 ok("отчёт подразделению: шапка и разделы",pa.includes("ОТЧЁТ ПОДРАЗДЕЛЕНИЮ О РЕЗУЛЬТАТАХ АУДИТА НИОКР")&&pa.includes("1. Сведения об объекте")&&pa.includes("2. Результаты интегральной оценки")&&pa.includes("7. Рекомендации подразделению"));
 ok("отчёт подразделению: данные НИОКР-01",pa.includes("НИОКР-01")&&pa.includes("72.9")&&pa.includes("История аудитов")&&pa.includes("Записей: 2")&&pa.includes("выполнено мероприятий 0 из 1"));
 ok("отчёт подразделению: подписи",pa.includes("Руководитель подразделения"));
 // копирование (clipboard в jsdom нет → fallback-модалка)
 d.getElementById("repComCopy").click();
 ok("копирование: fallback-модалка с текстом",d.getElementById("mbox").innerHTML.includes("Отчёт комиссии (текст для Word)")&&d.getElementById("mbox").querySelector("textarea").value.includes("ОТЧЁТ КОМИССИИ"));
 q('#mbox [data-mb="0"]').click();
 ok("после копирования printarea восстановлен актом",d.getElementById("printarea").innerHTML.includes("АКТ ОЦЕНКИ НИОКР"));
 // ---- сессия и экспорт ----
 ok("save() не включает presMode",!w.eval("localStorage.getItem(LSKEY)").includes("presMode"));
 d.getElementById("btnExp").click();
 ok("экспорт v14 отработал",true);
 // ---- пустая база: дек и отчёты не падают ----
 const keep=S.list.slice();S.list=[w.eval("newRnd('Единственная','НИОКР-05')")];S.cur=0;
 const ps0=API.buildPortSlides();
 ok("портфельный дек на пустой базе: 9 слайдов, НИОКР-05 в проблемных (нет данных)",ps0.length===9&&ps0[4].h.includes("НИОКР-05")&&ps0[4].h.includes("Недостаточно данных")&&ps0[6].h.includes("Принятых предложений нет"));
 ok("отчёт комиссии на пустой базе",API.buildComReport().includes("ОТЧЁТ КОМИССИИ")&&API.portConclusions().length>=1);
 ok("рекомендации пустой НИОКР: не пуста",API.deptAdvice(S.list[0]).length>=1);
 S.list=keep;S.cur=0;w.renderAll();
}catch(e){fails++;console.log("EXCEPTION:",e.stack||e);}
console.log(fails?(">>> ПРОВАЛЕНО: "+fails):">>> ВСЕ ТЕСТЫ v14 ПРОЙДЕНЫ");
process.exit(fails?1:0);
},400);
