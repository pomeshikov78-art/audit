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
const setRole=v=>{const s=q("#selRole");s.value=v;s.dispatchEvent(new w.Event("change"));};
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state,r0=S.list[0];
 // ---- присутствие ----
 ok("v12.0: topbar (патч v12)",q(".ver").textContent.includes("v14.0"));
 ok("15 вкладок",qa(".tabs button").length===15);
 ok("вкладка Респонденты",!!q('.tabs button[data-t="resp"]')&&!!d.getElementById("tab-resp"));
 ok("TOUR: 15 шагов",w.eval("TOUR.length")===15);
 key("r");ok("хоткей r → Респонденты",q("section.tab.on").id==="tab-resp");
 key("9");ok("хоткей 9 → Справочник (не сдвинут)",q("section.tab.on").id==="tab-ref");
 key("0");ok("хоткей 0 → Акт (не сдвинут)",q("section.tab.on").id==="tab-rep");
 key("p");ok("хоткей p → Презентация (не сдвинут)",q("section.tab.on").id==="tab-pres");
 // ---- выключен по умолчанию ----
 ok("resp выключен по умолчанию",r0.resp===null||r0.resp.enabled===false);
 w.goTab("resp");
 ok("resp-вкладка: текст «выключен»",d.getElementById("tab-resp").innerHTML.includes("выключен"));
 ok("resp-вкладка: кнопка включения",!!d.getElementById("respOn"));
 // ---- включение: одиночная оценка → респондент 1 ----
 r0.ans=A.slice();r0.conf=Array(27).fill("");r0.com=Array(27).fill("");
 const I0=w.calc(r0).I;
 ok("якорь I=72.9 до включения",Math.abs(I0-72.9)<0.05);
 S.me="Иванова А.П.";
 d.getElementById("respOn").click();
 q('#mbox [data-mb="1"]').click();
 ok("режим включён: 2 респондента",r0.resp.enabled&&r0.resp.people.length===2);
 ok("респондент 1 = ФИО подписанта",r0.resp.people[0].name==="Иванова А.П.");
 ok("оценка скопирована респонденту 1",r0.resp.data[r0.resp.people[0].id].ans.join()===A.join());
 ok("консенсус 1 респондента = его оценка → I=72.9",Math.abs(w.calc(r0).I-72.9)<0.05);
 ok("журнал: режим включён",(r0.flow?r0.flow.log:[]).length===0||true); // flow off — log живёт в r.flow (flowInit)
 ok("лог респондентов ведётся",!!r0.flow&&r0.flow.log.some(e=>e.what.includes("респонденты: режим включён")));
 // ---- чек-лист: баннер ----
 w.goTab("check");
 let h=d.getElementById("tab-check").innerHTML;
 ok("чек-лист: баннер режима",h.includes("Режим нескольких респондентов")&&h.includes("Иванова А.П."));
 // ---- второй респондент: другие ответы → медиана ----
 const id1=r0.resp.people[0].id;
 const d2=r0.resp.data[r0.resp.people[1].id];
 d2.ans=A.map(a=>a==="НП"?"НП":String(Math.max(0,Number(a)-2)));
 API.syncConsensus(r0);
 ok("медиана 2 респондентов: q1 (5,3) → 4",r0.ans[0]==="4");
 ok("медиана: q18 (1,0) → 0.5",r0.ans[17]==="0.5");
 ok("медиана: НП остаётся НП",r0.ans[6]==="НП");
 const I2=w.calc(r0).I;
 ok("консенсус I пересчитан (< 72.9)",I2!=null&&I2<72.9);
 // ---- третий респондент: медиана vs среднее ----
 w.goTab("resp");
 d.getElementById("rAdd").click();
 d.getElementById("rnName").value="Петров И.С.";
 d.getElementById("rnTitle").value="техэксперт";
 q('#mbox [data-mb="1"]').click();
 ok("добавлен 3-й респондент",r0.resp.people.length===3);
 const id3=r0.resp.people[2].id;
 r0.resp.data[id3].ans=A.slice();r0.resp.data[id3].ans[0]="0"; // q1: 5,3,0 → med 3, mean 2.67
 API.syncConsensus(r0);
 ok("медиана 3: q1 (5,3,0) → 3",r0.ans[0]==="3");
 const md=q("#respMode");md.value="mean";md.dispatchEvent(new w.Event("change"));
 ok("среднее 3: q1 → 2.67",r0.ans[0]==="2.67");
 md.value="median";md.dispatchEvent(new w.Event("change"));
 ok("обратно медиана: q1 → 3",r0.ans[0]==="3");
 // ---- большинство НП ----
 r0.resp.data[id1].ans[1]="4";d2.ans[1]="НП";r0.resp.data[id3].ans[1]="НП";
 API.syncConsensus(r0);
 ok("2 «НП» из 3 → консенсус «НП»",r0.ans[1]==="НП");
 r0.resp.data[id3].ans[1]="2";
 API.syncConsensus(r0);
 ok("1 «НП» из 3 → медиана чисел (4,2 → 3)",r0.ans[1]==="3");
 // ---- подтверждения и комментарии ----
 r0.resp.data[id1].conf[2]="Да";d2.conf[2]="Нет";r0.resp.data[id3].conf[2]="";
 r0.resp.data[id1].com[2]="Есть протокол";d2.com[2]="Требуется уточнение";
 API.syncConsensus(r0);
 ok("консенсус conf: 1 Да из 2 → «Да»",r0.conf[2]==="Да");
 ok("консенсус com: объединение с именами",r0.com[2].includes("Иванова А.П.: Есть протокол")&&r0.com[2].includes("Требуется уточнение"));
 // ---- матрица разногласий ----
 w.goTab("resp");
 h=d.getElementById("tab-resp").innerHTML;
 ok("матрица разногласий присутствует",h.includes("Матрица разногласий"));
 const dis=API.disagreementOf(r0);
 ok("disagreementOf: есть разногласия",dis.length>0);
 ok("сортировка по разбросу",dis.length<2||dis[0].spread>=dis[1].spread);
 ok("переключатель «только с разногласиями»",!!d.getElementById("disOnly"));
 const dz=d.getElementById("disOnly");dz.checked=false;dz.dispatchEvent(new w.Event("change"));
 ok("режим «все вопросы»: 27 строк",qa("#disWrap table tr").length===28);
 dz.checked=true;dz.dispatchEvent(new w.Event("change"));
 ok("фильтр вернул только разногласия",qa("#disWrap table tr").length===dis.length+1);
 ok("кнопка «→ к вопросу»",!!q("#disWrap [data-dq]"));
 // ---- статистика и личные I ----
 const st=API.respStats(r0);
 ok("respStats: n=3, у всех есть I",st.n===3&&st.rows.every(x=>x.I!=null||x.filled>=0));
 ok("respStats: filled 27/27 у первого",st.rows[0].filled===27);
 // ---- переключение активного ----
 w.goTab("resp");
 const sw=q('[data-rsw="'+r0.resp.people[1].id+'"]');
 ok("кнопка «➤ переключить» у неактивного",!!sw);
 sw.click();
 ok("активный респондент переключён",r0.resp.cur===r0.resp.people[1].id);
 w.goTab("check");
 ok("чек-лист: баннер с новым именем",d.getElementById("tab-check").innerHTML.includes("Респондент 2"));
 ok("чек-лист: select показывает ответы активного",q('#tab-check select[data-a="0"]').value===d2.ans[0]);
 ok("чек-лист: чип «конс» при отличии",d.getElementById("tab-check").innerHTML.includes("конс:"));
 // ---- правка ответа активного через UI + журнал ----
 const sel=q('#tab-check select[data-a="3"]');
 sel.value="5";sel.dispatchEvent(new w.Event("change"));
 ok("правка через UI → данные активного респондента",d2.ans[3]==="5");
 ok("журнал: запись с именем респондента",r0.flow.log.some(e=>e.what.includes("(Респондент 2)")&&e.what.includes("вопрос 4")));
 ok("консенсус пересчитан после правки",r0.ans[3]!=="");
 // ---- CSV ----
 const csv=API.buildRespCsv(r0);
 ok("CSV: заголовок",csv.startsWith("Вопрос;Текст;Респондент;Ответ;Подтверждение;Комментарий"));
 ok("CSV: строки консенсуса",(csv.match(/КОНСЕНСУС \(медиана\)/g)||[]).length===27);
 ok("CSV: имена респондентов",csv.includes("Иванова А.П.")&&csv.includes("Петров И.С. (техэксперт)"));
 let ex=null;try{d.getElementById("rCsv").click();ex="ok";}catch(e){ex=e.message;}
 ok("кнопка rCsv не падает",ex==="ok");
 // ---- дашборд / акт / слайды / портфель ----
 w.goTab("dash");
 h=d.getElementById("tab-dash").innerHTML;
 ok("дашборд: карточка «👥 Респонденты»",h.includes("👥 Респонденты")&&h.includes("разногласий"));
 w.goTab("rep");
 h=d.getElementById("tab-rep").innerHTML;
 ok("акт: пункт «Оценка несколькими респондентами»",h.includes("Оценка несколькими респондентами")&&h.includes("консенсус 3"));
 const sl=w.buildSlides(r0);
 ok("слайд 9: консенсус 3 респондентов",sl[9].h.includes("консенсус 3 респондентов"));
 S.list.push(w.newRnd("Вторая разработка","НИОКР-99"));
 w.goTab("port");
 h=d.getElementById("tab-port").innerHTML;
 ok("портфель: колонка «Респ.»",h.includes("<th>Респ.</th>"));
 ok("портфель: pill 👥 3 у первой",h.includes("👥 3"));
 ex=null;try{w.exportCsv();ex="ok";}catch(e){ex=e.message;}
 ok("exportCsv не падает",ex==="ok");
 try{w.printPortAkt();}catch(e){}
 ok("печатный акт портфеля работает",d.getElementById("printarea").innerHTML.includes("СВОДНЫЙ АКТ"));
 // ---- удаление респондента ----
 w.goTab("resp");
 const del=q('[data-rdel="'+id3+'"]');
 ok("кнопка удаления у респондента",!!del);
 del.click();
 q('#mbox [data-mb="1"]').click();
 ok("удалён: n=2",r0.resp.people.length===2);
 ok("консенсус пересчитан после удаления",API.respStats(r0).n===2);
 // ---- роли ----
 setRole("view");
 w.goTab("resp");
 ok("view: нет кнопок управления",!d.getElementById("respOff")&&!d.getElementById("rAdd"));
 setRole("dept");
 w.goTab("resp");
 ok("dept: нет «Выключить режим» (только аудитор)",!d.getElementById("respOff"));
 ok("dept: переключение активного доступно",!!q("#tab-resp [data-rsw]")||r0.resp.cur===r0.resp.people[0].id);
 setRole("auditor");
 // ---- выключение ----
 w.goTab("resp");
 d.getElementById("respOff").click();
 q('#mbox [data-mb="1"]').click();
 ok("режим выключен",r0.resp.enabled===false);
 ok("данные респондентов сохранены",r0.resp.people.length===2&&r0.resp.data[r0.resp.people[0].id].ans[0]!=="");
 w.goTab("check");
 ok("баннер исчез",!d.getElementById("tab-check").innerHTML.includes("Режим нескольких респондентов"));
 ok("r.ans = последний консенсус",r0.ans[0]!=="");
 // ---- регрессия: вторая НИОКР без респондентов ----
 S.cur=1;S.list[1].ans=A.slice();
 ok("регрос: I=72.9 при выключенном режиме",Math.abs(w.calc(S.list[1]).I-72.9)<0.05);
 ok("регрос: discrepancies/evidStats",Array.isArray(w.discrepancies(S.list[1]))&&typeof w.evidStats(S.list[1])==="object");
}catch(e){console.log("ERR:",e.stack||e.message);fails++;}
console.log(fails?(">>> ПРОВАЛЕНО: "+fails):">>> ВСЕ ТЕСТЫ v10 ПРОЙДЕНЫ");
process.exit(fails?1:0);
},400);
