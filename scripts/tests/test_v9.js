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
const onTab=id=>d.querySelector("section.tab.on")&&d.querySelector("section.tab.on").id===id;
const ansVals=["5","4","4","3","5","4","НП","3","4","4","3","3","4","3","3","4","3","1","5","4","4","3","4","4","3","4","НП"];
const setRole=v=>{const s=q("#selRole");s.value=v;s.dispatchEvent(new w.Event("change"));};
setTimeout(()=>{
try{
 const S=w.niokrAPI.state;
 // ---- базовое присутствие ----
 ok("topbar: версия v10+",/v10\.0/.test(q(".ver").textContent));
 ok("12 вкладок",qa(".tabs button").length===12);
 ok("вкладка Маршрут есть",!!q('.tabs button[data-t="flow"]')&&!!d.getElementById("tab-flow"));
 ok("TOUR: 12 шагов",w.eval("TOUR.length")===12);
 ok("роли: 5 (fin/exp добавлены)",qa("#selRole option").length===5&&!!q('#selRole option[value="fin"]')&&!!q('#selRole option[value="exp"]'));
 ok("FLOW: 4 этапа",w.eval("FLOW.length")===4);
 // ---- хоткеи ----
 key("8");ok("хоткей 8 → Маршрут",onTab("tab-flow"));
 key("9");ok("хоткей 9 → Справочник",onTab("tab-ref"));
 key("0");ok("хоткей 0 → Акт",onTab("tab-rep"));
 key("p");ok("хоткей p → Презентация",onTab("tab-pres"));
 key("1");ok("хоткей 1 → Дашборд (регрос)",onTab("tab-dash"));
 // ---- маршрут выключен по умолчанию ----
 ok("flow выключен по умолчанию",w.niokrAPI.flowOf(S.list[0])===null);
 w.goTab("flow");
 let h=d.getElementById("tab-flow").innerHTML;
 ok("flow-вкладка: текст «выключен»",h.includes("выключен"));
 ok("flow-вкладка: кнопка включения (аудитор)",!!d.getElementById("flowOn"));
 ok("дашборд: карточка маршрута",d.getElementById("tab-dash").innerHTML.includes("Маршрут согласования"));
 // ---- ответы + якорь I ----
 S.list[0].ans=ansVals.slice();
 ok("якорь I=72.9",Math.abs(w.calc(S.list[0]).I-72.9)<0.05);
 // ---- включение маршрута ----
 S.me="Иванова А.П.";
 d.getElementById("flowOn").click();
 let mb=q('#mbox [data-mb="1"]');
 ok("модалка включения",!!mb&&/Включить/.test(mb.textContent));
 mb.click();
 let f=w.niokrAPI.flowOf(S.list[0]);
 ok("маршрут включён: stage=1",!!f&&f.stage===1&&f.stages[0].status==="work");
 ok("журнал: запись включения",f.log.some(e=>e.what.includes("маршрут включён")));
 ok("журнал: «кто» с ФИО",f.log[0].who.includes("Иванова А.П."));
 h=d.getElementById("tab-flow").innerHTML;
 ok("flow UI: Этап 1 из 4 · Подразделение",h.includes("Этап 1 из 4")&&h.includes("Подразделение"));
 ok("flow UI: журнал «кто / когда / что»",h.includes("Журнал изменений"));
 // ---- журнал событий чек-листа ----
 w.goTab("check");
 const sel=qa('#tab-check select[data-a="3"]')[0];
 sel.value="4";sel.dispatchEvent(new w.Event("change"));
 ok("журнал: изменение ответа q4",f.log.some(e=>e.what.includes("чек-лист: вопрос 4 ответ")));
 // ---- треды ----
 const thBtn=q('#tab-check [data-th="0"]');
 ok("чек-лист: кнопка 💬 у вопроса",!!thBtn);
 thBtn.click();
 ok("тред-модалка открыта",!!d.getElementById("thTxt"));
 d.getElementById("thTxt").value="Уточните источник данных по вопросу 1";
 q('#mbox [data-mb="1"]').click();
 ok("тред: комментарий записан",(S.list[0].threads[0]||[]).length===1);
 ok("тред: автор = роль+ФИО",S.list[0].threads[0][0].who.includes("Иванова"));
 ok("журнал: запись треда",f.log.some(e=>e.what.includes("тред вопроса 1")));
 ok("чек-лист: бейдж 💬 1",q('#tab-check [data-th="0"]').textContent.includes("1"));
 w.goTab("flow");
 ok("flow UI: чип треда",!!q('#tab-flow [data-thq="0"]'));
 // ---- прохождение маршрута ----
 w.goTab("flow");
 qa('#tab-flow [data-fa="ok"]')[0].click(); // этап 1 → 2
 f=w.niokrAPI.flowOf(S.list[0]);
 ok("этап 1 согласован → этап 2 Финансы",f.stage===2&&f.stages[0].status==="ok"&&f.stages[1].status==="work");
 ok("журнал: СОГЛАСОВАН + передача",f.log.some(e=>e.what.includes("СОГЛАСОВАН"))&&f.log.some(e=>e.what.includes("передано на этап 2")));
 // ---- блокировки ролей на этапе 2 ----
 setRole("dept");
 w.goTab("pass");
 ok("dept на этапе 2: паспорт заблокирован",qa('#tab-pass [data-p="podrazd"]')[0].disabled===true);
 ok("dept: плашка блокировки",d.getElementById("tab-pass").innerHTML.includes("Маршрут согласования: этап 2"));
 w.goTab("flow");
 ok("dept на этапе 2: нет кнопок этапа",!q('#tab-flow [data-fa="ok"]'));
 setRole("fin");
 w.goTab("pass");
 ok("fin на этапе 2: fin.plan доступен",qa('#tab-pass [data-p="fin.plan"]')[0].disabled===false);
 ok("fin на этапе 2: podrazd заблокирован",qa('#tab-pass [data-p="podrazd"]')[0].disabled===true);
 w.goTab("flow");
 ok("fin на этапе 2: кнопки этапа есть",!!q('#tab-flow [data-fa="ok"]')&&!!q('#tab-flow [data-fa="back"]'));
 // ---- возврат с этапа 2 ----
 q('#tab-flow [data-fa="back"]').click();
 ok("модалка возврата",!!d.getElementById("backTxt"));
 d.getElementById("backTxt").value="";
 q('#mbox [data-mb="1"]').click();
 ok("возврат без комментария отклонён",w.niokrAPI.flowOf(S.list[0]).stage===2);
 // модалка закрылась (hideModal до fn) — открываем заново
 q('#tab-flow [data-fa="back"]')&&0;
 w.renderFlow();q('#tab-flow [data-fa="back"]').click();
 d.getElementById("backTxt").value="Уточнить факт 2025 — расхождение с δб";
 q('#mbox [data-mb="1"]').click();
 f=w.niokrAPI.flowOf(S.list[0]);
 ok("возврат: stage=1, этап2=back",f.stage===1&&f.stages[1].status==="back");
 ok("возврат: комментарий сохранён",f.stages[1].comment.includes("Уточнить факт"));
 ok("возврат: этап1 снова work",f.stages[0].status==="work");
 ok("журнал: ВОЗВРАТ",f.log.some(e=>e.what.includes("ВОЗВРАТ")));
 // ---- dept на этапе 1: паспорт доступен, решение — нет ----
 setRole("dept");
 w.goTab("pass");
 ok("dept на этапе 1: podrazd доступен",qa('#tab-pass [data-p="podrazd"]')[0].disabled===false);
 ok("dept на этапе 1: решение заблокировано",qa('#tab-pass [data-p="decision"]')[0].disabled===true);
 setRole("auditor");
 // ---- полный проход до завершения ----
 w.goTab("flow");
 for(let k=0;k<4;k++){const b=q('#tab-flow [data-fa="ok"]');if(b)b.click();}
 f=w.niokrAPI.flowOf(S.list[0]);
 ok("маршрут завершён (done)",f.done===true&&f.stages.every(s=>s.status==="ok"));
 ok("flow UI: «Маршрут завершён»",d.getElementById("tab-flow").innerHTML.includes("Маршрут завершён"));
 ok("журнал: ЗАВЕРШЁН",f.log.some(e=>e.what.includes("ЗАВЕРШЁН")));
 ok("после завершения кнопок нет",!q('#tab-flow [data-fa="ok"]'));
 // ---- портфель / CSV / акты ----
 S.list.push(w.newRnd("Вторая разработка","НИОКР-99"));
 w.goTab("port");
 h=d.getElementById("tab-port").innerHTML;
 ok("портфель: колонка Маршрут",h.includes("<th>Маршрут</th>"));
 ok("портфель: завершён у первой",h.includes("✔ завершен")||h.includes("завершён"));
 ok("портфель: — у второй",h.includes("—"));
 let ex=null;try{w.exportCsv();ex="ok";}catch(e){ex=e.message;}
 ok("exportCsv не падает",ex==="ok");
 w.goTab("rep");
 h=d.getElementById("tab-rep").innerHTML;
 ok("акт: пункт «Маршрут согласования (workflow)»",h.includes("Маршрут согласования (workflow)")&&h.includes("Завершён"));
 try{w.printPortAkt();}catch(e){}
 h=d.getElementById("printarea").innerHTML;
 ok("акт портфеля: th Маршрут + завершён",h.includes("<th>Маршрут</th>")&&h.includes("завершён"));
 // ---- справка ----
 w.eval("helpModal()");
 h=d.getElementById("mbox").innerHTML;
 ok("справка: клавиши 1…9, 0, p",h.includes("1…9, 0, p"));
 ok("справка: роли Финансы/Техэксперт",h.includes("Финансы")&&h.includes("Техэксперт"));
 w.eval("hideModal()");
 // ---- регрессия: flow выключен → поведение v8 ----
 S.cur=1;
 S.list[1].ans=ansVals.slice();
 ok("регрос: I=72.9 при выключенном маршруте",Math.abs(w.calc(S.list[1]).I-72.9)<0.05);
 setRole("dept");
 w.goTab("pass");
 ok("регрос: dept без маршрута — паспорт доступен",qa('#tab-pass [data-p="podrazd"]')[0].disabled===false);
 ok("регрос: dept без маршрута — решение доступно",qa('#tab-pass [data-p="decision"]')[0].disabled===false);
 setRole("view");
 ok("регрос: view — всё заблокировано",qa('#tab-pass [data-p="podrazd"]')[0].disabled===true);
 w.goTab("flow");
 ok("view: нет кнопки включения маршрута",!d.getElementById("flowOn")&&d.getElementById("tab-flow").innerHTML.includes("только роль «Аудитор»"));
 setRole("auditor");
 ok("регрос: discrepancies/evidStats",Array.isArray(w.discrepancies(S.list[0]))&&typeof w.evidStats(S.list[0])==="object");
}catch(e){console.log("ERR:",e.stack||e.message);fails++;}
console.log(fails?(">>> ПРОВАЛЕНО: "+fails):">>> ВСЕ ТЕСТЫ v9 ПРОЙДЕНЫ");
process.exit(fails?1:0);
},400);
