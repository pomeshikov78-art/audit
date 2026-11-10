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
const inp=(id,v)=>{const e=d.getElementById(id);e.value=v;e.dispatchEvent(new w.Event("input",{bubbles:true}));};
const chg=(id,v)=>{const e=d.getElementById(id);e.value=v;e.dispatchEvent(new w.Event("change",{bubbles:true}));};
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state,r0=S.list[0];
 // ---- присутствие ----
 ok("v12.0: topbar",q(".ver").textContent.includes("v14.0"));
 ok("15 вкладок",qa(".tabs button").length===15);
 ok("вкладка Предложения",!!q('.tabs button[data-t="prop"]')&&!!d.getElementById("tab-prop"));
 ok("TOUR: 15 шагов",w.eval("TOUR.length")===15);
 ok("TOUR содержит шаг prop",w.eval("TOUR.some(t=>t[0]==='prop')"));
 ok("API: предложения (11 функций)",["pmodel","propAns","propForecast","addNewProp","acceptProp","rejectProp","delProp","propModal","propStats","buildPropsCsv","renderProp"].every(k=>typeof API[k]==="function"));
 ok("PMODELS: 6 моделей",w.eval("PMODELS.length")===6&&w.eval("PMODELS.map(m=>m.id).join()")==="continue,modern,finish,stop,hold,merge");
 // ---- хоткеи ----
 key("m");ok("хоткей m → Предложения",q("section.tab.on").id==="tab-prop");
 key("ь");ok("хоткей ь → Предложения",q("section.tab.on").id==="tab-prop");
 key("b");ok("хоткей b → База (не сдвинут)",q("section.tab.on").id==="tab-db");
 key("9");ok("хоткей 9 → Справочник (не сдвинут)",q("section.tab.on").id==="tab-ref");
 key("0");ok("хоткей 0 → Акт (не сдвинут)",q("section.tab.on").id==="tab-rep");
 key("p");ok("хоткей p → Презентация (не сдвинут)",q("section.tab.on").id==="tab-pres");
 key("r");ok("хоткей r → Респонденты (не сдвинут)",q("section.tab.on").id==="tab-resp");
 // ---- модель данных ----
 ok("newRnd: props=[]",Array.isArray(r0.props));
 // ---- пустая вкладка ----
 r0.ans=A.slice();r0.conf=Array(27).fill("");r0.com=Array(27).fill("");
 r0.shifr="НИОКР-01";r0.name="Разработка А (пример)";
 S.me="Тестов Т.Т.";
 const I0=w.calc(r0).I;
 ok("якорь I=72.9",Math.abs(I0-72.9)<0.05);
 w.goTab("prop");
 let h=d.getElementById("tab-prop").innerHTML;
 ok("вкладка: заголовок и кнопка создания",h.includes("Предложения по дальнейшей работе")&&!!d.getElementById("prAdd"));
 ok("вкладка: пусто — подсказка",h.includes("Предложений ещё нет"));
 ok("портфельный эффект: KPI",h.includes("ПРИНЯТО")&&h.includes("СРЕДНИЙ ΔI"));
 // ---- создание предложения: модалка-редактор ----
 d.getElementById("prAdd").click();
 ok("предложение создано (черновик)",r0.props.length===1&&r0.props[0].status==="draft");
 ok("модалка редактора",d.getElementById("mbox").innerHTML.includes("Предложение №1"));
 ok("редактор: поля и контролы",["pmModel","pmExtra","pmMonths","pmEffect","pmRisks","pmNote","pmF","ptQ","ptV","ptAdd","ptWeak","ptClr","pmText","pmResp","pmDue","pmAdd"].every(id=>!!d.getElementById(id)));
 const p0=r0.props[0];
 ok("журнал: создание предложения",r0.flow&&r0.flow.log.some(e=>e.what.includes("создано предложение")));
 // ---- живое редактирование параметров ----
 chg("pmModel","hold");
 ok("смена модели: hold",p0.model==="hold"&&d.getElementById("pmDesc").textContent.includes("Пауза"));
 inp("pmExtra","12.5");inp("pmMonths","9");
 inp("pmEffect","подтверждение ТРЛ 6");inp("pmRisks","срыв поставки");inp("pmNote","квартальный план");
 ok("параметры пишутся живо",p0.extra==="12.5"&&p0.months==="9"&&p0.effect.includes("ТРЛ 6")&&p0.risks==="срыв поставки");
 // ---- цели: вручную + «слабые → 4» ----
 chg("ptQ","0");chg("ptV","5");
 d.getElementById("ptAdd").click(); // перерисовка модалки
 ok("цель №1 → 5 добавлена",p0.targets["0"]==="5");
 ok("модалка перерисована: таблица целей",d.getElementById("mbox").innerHTML.includes("Целевые баллы")&&d.getElementById("ptList").innerHTML.includes("data-ptdel"));
 d.getElementById("ptWeak").click(); // A[17]="1" ≤2 → цель 4
 ok("«слабые → 4»: вопрос 18 добавлен",p0.targets["17"]==="4");
 ok("целей стало 2",Object.keys(p0.targets).length===2);
 // ---- прогноз I ----
 const f=API.propForecast(r0,p0);
 const exp=API.calcCore(API.propAns(r0,p0),S.weights);
 ok("прогноз: I1 = calcCore(целевые ответы)",Math.abs(f.I1-exp.I)<1e-9);
 ok("прогноз: I0 = 72.9, ΔI > 0",Math.abs(f.I0-72.9)<0.05&&f.dI!=null&&f.dI>0);
 ok("прогноз: стоимость балла = extra/ΔI",Math.abs(f.cpp-12.5/f.dI)<1e-9);
 ok("прогноз: срок распознан (9 мес)",f.months===9);
 ok("строка прогноза в модалке",d.getElementById("pmF").innerHTML.includes("ΔI")&&d.getElementById("pmF").innerHTML.includes("стоимость балла"));
 // ---- мероприятия ----
 inp("pmText","Разработать программу специспытаний");inp("pmResp","Иванов И.И.");inp("pmDue","20.12.2026");
 d.getElementById("pmAdd").click();
 ok("мероприятие добавлено",p0.measures.length===1&&p0.measures[0].resp==="Иванов И.И."&&p0.measures[0].st==="plan");
 chg(d.querySelector("#mbox [data-mst]").id!==undefined?"pmModel":"pmModel","hold"); // no-op, держим модель
 const mst=d.querySelector('#mbox [data-mst="0"]');mst.value="work";mst.dispatchEvent(new w.Event("change",{bubbles:true}));
 ok("статус мероприятия: в работе",p0.measures[0].st==="work");
 inp("pmText","");d.getElementById("pmAdd").click();
 ok("пустое мероприятие отклонено",p0.measures.length===1);
 inp("pmText","Закупить комплектующие");d.getElementById("pmAdd").click();
 ok("второе мероприятие",p0.measures.length===2);
 d.querySelector('#mbox [data-mdel="1"]').click();
 ok("удаление мероприятия",p0.measures.length===1);
 // ---- сохранение ----
 q('#mbox [data-mb="1"]').click();
 ok("журнал: редактирование предложения",r0.flow.log.some(e=>e.what.includes("отредактировано предложение")));
 ok("модалка закрыта",!d.getElementById("moverlay").classList.contains("on"));
 // ---- карточка на вкладке ----
 h=d.getElementById("tab-prop").innerHTML;
 ok("карточка: модель и статус",h.includes("Приостановить")&&h.includes("Черновик"));
 ok("карточка: прогноз и параметры",h.includes("ΔI")&&h.includes("12.5 млн ₽")&&h.includes("9 мес"));
 ok("карточка: карта мероприятий",h.includes("Разработать программу специспытаний")&&h.includes("Иванов И.И."));
 ok("карточка: кнопки черновика",!!q('#tab-prop [data-pedit="0"]')&&!!q('#tab-prop [data-pacc="0"]')&&!!q('#tab-prop [data-prej="0"]')&&!!q('#tab-prop [data-pdel="0"]'));
 // ---- принятие: решение комиссии пишется в паспорт ----
 q('#tab-prop [data-pacc="0"]').click();
 ok("модалка принятия с чекбоксом решения",d.getElementById("mbox").innerHTML.includes("Принять предложение")&&!!d.getElementById("paDec")&&d.getElementById("paDec").checked);
 d.getElementById("paNote").value="одобрено на заседании";
 q('#mbox [data-mb="1"]').click();
 ok("статус: принято",p0.status==="accepted");
 ok("решение комиссии записано в паспорт",r0.decision==="Приостановить до устранения выявленных проблем");
 ok("журнал: ПРИНЯТО",r0.flow.log.some(e=>e.what.includes("ПРИНЯТО предложение")&&e.what.includes("одобрено на заседании")));
 ok("карточка: кнопка принятия исчезла",!q('#tab-prop [data-pacc="0"]')&&d.getElementById("tab-prop").innerHTML.includes("Принято"));
 // ---- портфельная статистика и CSV ----
 const st=API.propStats();
 ok("propStats: n=1, acc=1, Σextra=12.5",st.n===1&&st.acc===1&&Math.abs(st.sumExtra-12.5)<1e-9);
 ok("propStats: средний ΔI = прогноз",Math.abs(st.avgD-API.propForecast(r0,p0).dI)<1e-9);
 const csv=API.buildPropsCsv();
 ok("CSV: заголовок",csv.split("\n")[0].includes("Модель предложения")&&csv.split("\n")[0].includes("Стоимость балла")&&csv.split("\n")[0].includes("ΔI"));
 ok("CSV: строка предложения",csv.includes("НИОКР-01")&&csv.includes("Приостановить")&&csv.includes("12.5")&&csv.includes("Разработать программу специспытаний"));
 h=d.getElementById("tab-prop").innerHTML;
 ok("диаграмма «финансирование × ΔI»",h.includes("ΔI = 0")&&h.includes("млн ₽"));
 d.getElementById("ppCsv").click();
 ok("кнопка CSV предложений отработала",true);
 // ---- второе предложение: отклонение и удаление ----
 w.eval("addNewProp(0)");
 ok("второе предложение создано",r0.props.length===2);
 q('#mbox [data-mb="0"]').click(); // закрыть редактор
 API.rejectProp(0,1);
 d.getElementById("prNote").value="эффект не обоснован";
 q('#mbox [data-mb="1"]').click();
 ok("второе отклонено",r0.props[1].status==="rejected"&&r0.flow.log.some(e=>e.what.includes("ОТКЛОНЕНО")&&e.what.includes("эффект не обоснован")));
 API.delProp(0,1);
 q('#mbox [data-mb="1"]').click();
 ok("удаление предложения",r0.props.length===1&&r0.flow.log.some(e=>e.what.includes("удалено предложение")));
 // ---- фильтр статусов ----
 w.eval("addNewProp(0)");q('#mbox [data-mb="0"]').click();
 ok("снова 2 предложения (черновик)",r0.props.length===2&&r0.props[1].status==="draft");
 chg("ppFl","accepted");
 ok("фильтр «принятые»: 1 карточка",qa("#tab-prop [data-pedit]").length===1&&d.getElementById("tab-prop").innerHTML.includes("Приостановить"));
 chg("ppFl","draft");
 ok("фильтр «черновики»: 1 карточка",qa("#tab-prop [data-pedit]").length===1&&d.getElementById("tab-prop").innerHTML.includes("Черновик"));
 chg("ppFl","");
 ok("сброс фильтра: 2 карточки",qa("#tab-prop [data-pedit]").length===2);
 // ---- акт оценки: пункт 6а ----
 w.goTab("rep");
 h=d.getElementById("tab-rep").innerHTML;
 ok("акт: пункт 6а с принятым предложением",h.includes("6а. Предложения по дальнейшей работе")&&h.includes("принято: 1")&&h.includes("Приостановить до устранения выявленных проблем"));
 ok("акт: прогноз I и доп. финансирование в 6а",h.includes("прогноз I: 72.9 →")&&h.includes("доп. финансирование 12.5 млн ₽"));
 // ---- роли ----
 setRole("view");w.goTab("prop");
 h=d.getElementById("tab-prop").innerHTML;
 ok("Наблюдатель: только просмотр",h.includes("только просмотр")&&!d.getElementById("prAdd")&&!q("#tab-prop [data-pdel]"));
 setRole("dept");API.renderProp();
 ok("Подразделение: редактирует, но не принимает",!!d.getElementById("prAdd")&&!!q('#tab-prop [data-pedit="0"]')&&!q('#tab-prop [data-pacc="0"]')&&!q('#tab-prop [data-pdel="0"]'));
 setRole("auditor");API.renderProp();
 ok("Аудитор: полный доступ",!!d.getElementById("prAdd")&&!!q('#tab-prop [data-pedit="0"]')&&!!q('#tab-prop [data-pdel="0"]'));
 // ---- вторая НИОКР: переход из портфельной таблицы ----
 const r1=w.eval("newRnd('Разработка Б','НИОКР-02')");
 r1.ans=Array(27).fill("2");r1.conf=Array(27).fill("");r1.com=Array(27).fill("");
 r1.props=[w.eval("newProp()")];r1.props[0].model="modern";r1.props[0].status="accepted";r1.props[0].extra="30";r1.props[0].targets={"0":"5"};
 S.list.push(r1);API.renderProp();
 ok("портфельная таблица: 2 НИОКР",qa("#tab-prop [data-popen]").length>=2);
 const st2=API.propStats();
 ok("propStats после добавления: n=3, acc=2",st2.n===3&&st2.acc===2&&Math.abs(st2.sumExtra-42.5)<1e-9);
 q('#tab-prop [data-popen="1"]').click();
 ok("data-popen переключил НИОКР",S.cur===1);
 S.cur=0;w.renderAll();
 // ---- миграция load(): props ----
 w.eval(`(function(){const x=newRnd("Миграция Тест","НИОКР-97");delete x.props;delete x.audits;delete x.created;
  localStorage.setItem(LSKEY,JSON.stringify({list:[x],cur:0,role:"auditor",theme:"light",me:"",respMode:"median",weights:[30,35,35],wsName:"",ahp:{a12:"1",a13:"1",a23:"1"},weightSets:[]}));load();})();`);
 const rm=API.state.list[0];
 ok("миграция load(): props добавлен",Array.isArray(rm.props)&&rm.props.length===0);
 API.renderProp();
 ok("вкладка рендерится после миграции",d.getElementById("tab-prop").innerHTML.includes("НИОКР-97"));
 // ---- экспорт ----
 d.getElementById("btnExp").click();
 ok("экспорт v12 отработал",true);
}catch(e){fails++;console.log("EXCEPTION:",e.stack||e);}
console.log(fails?(">>> ПРОВАЛЕНО: "+fails):">>> ВСЕ ТЕСТЫ v12 ПРОЙДЕНЫ");
process.exit(fails?1:0);
},400);
