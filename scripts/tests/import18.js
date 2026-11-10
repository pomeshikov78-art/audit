/* Этап 18Б, шаг 3: встроенный читатель xlsx + импорт калькулятора № 9 (interop). */
const fs=require("fs");
const {JSDOM}=require("jsdom");
const WS="/tmp/arena/tenants/71ebd7cc-b365-41c5-863c-79f587b1df1f/sessions/76eaa8bd-b9c6-43e4-8bfe-0b8d30b3272e/workspace";
const P=WS+"/output/Аудит_НИОКР_интерактивная_модель.html";
const dom=new JSDOM(fs.readFileSync(P,"utf8"),{runScripts:"dangerously",pretendToBeVisual:true,url:"http://localhost/"});
const w=dom.window,d=w.document;
w.HTMLCanvasElement.prototype.getContext=()=>({clearRect(){},fillRect(){},beginPath(){},moveTo(){},lineTo(){},stroke(){},fill(){},arc(){},save(){},restore(){},translate(){},rotate(){},fillText(){},setLineDash(){},closePath(){},quadraticCurveTo(){}});
w.URL.createObjectURL=()=>"blob:x";w.URL.revokeObjectURL=()=>{};
let winErr=0;
w.addEventListener("error",e=>{if(/Not implemented/.test(e.message||""))return;winErr++;console.log("WINDOW-ERR:",e.message);});
let fails=0;
const ok=(n,c)=>{console.log((c?"PASS":"FAIL")+" | "+n);if(!c)fails++;};
const q=s=>d.querySelector(s),qa=s=>[...d.querySelectorAll(s)];
const click=el=>el.dispatchEvent(new w.MouseEvent("click",{bubbles:true}));
const near=(a,b,e)=>a!=null&&Math.abs(a-b)<=e;
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state;
 ok("нет ошибок окна (загрузка)",winErr===0);
 ok("API отдаёт xlsx18/importExcel9/importPreview18",!!(API.xlsx18&&API.importExcel9&&API.importPreview18&&API.importShow18&&API.deriveConsensus18));

 /* ===== A. unit: deriveConsensus18 (зеркало формулы M листа № 9) ===== */
 const dc=API.deriveConsensus18;
 ok("dc: [5,4,1] равные веса → 3.33",dc([5,4,1],[1,1,1])==="3.33");
 ok("dc: НП большинство → НП",dc(["НП","НП","НД"],[1,1,1])==="НП");
 ok("dc: 1 НП + 1 число → среднее (Excel: COUNT>=COUNTIF)",dc([2,"НП"],[1,1])==="2");
 ok("dc: пусто → ''",dc([null,"",undefined],[1,1,1])==="");
 ok("dc: НД без чисел → НД",dc(["НД","НД",null],[1,1,1])==="НД");
 ok("dc: веса 0.4/0.3/0.3 по одинаковым → тот же балл",dc([3,3,3],[0.4,0.3,0.3])==="3");
 ok("dc: [5,4] равные → 4.5",dc([5,4,null],[1,1,1])==="4.5");

 /* ===== B. xlsx18 против пустого файла № 9 ===== */
 const buf9=fs.readFileSync(WS+"/output/9_Калькулятор_аудита_НИОКР_v2.xlsx");
 const wbk9=API.xlsx18.parse(buf9);
 ok("№9: листов 14",wbk9.sheetNames.length===14);
 ok("№9: есть ключевые листы",["Паспорт_НИОКР","Краткий_15","Средний_30","Полный_60","Стоп_факторы","Зрелость_M0_M9","Экономика_LCC","AHP","Эксперты"].every(n=>!!wbk9.cells[n]));
 ok("№9: Паспорт C17 = «Средний» (shared strings декодирован)",wbk9.get("Паспорт_НИОКР","C17")==="Средний");
 ok("№9: M6 формула без кэша → null",wbk9.get("Средний_30","M6")===null);
 ok("№9: текст вопроса C6 прочитан",typeof wbk9.get("Средний_30","C6")==="string"&&wbk9.get("Средний_30","C6").length>10);
 ok("№9: AHP B18 — строка «1»",String(wbk9.get("AHP","B18"))==="1");
 ok("№9: toggle Полный C82 = «Нет»",wbk9.get("Полный_60","C82")==="Нет");
 const pv9=API.importPreview18(wbk9);
 ok("№9 пустой: предпросмотр 0 ответов",pv9.q.K===0&&pv9.q.S===0&&pv9.q.P===0&&pv9.total===0);
 ok("№9 пустой: все 9 обязательных листов найдены",pv9.found.length===9&&pv9.miss.length===0);
 ok("№9 пустой: lvl из паспорта = S",pv9.lvl==="S");
 const r2=S.list[0];const nm2=r2.name;
 const rep9=API.importExcel9(wbk9,r2);
 ok("№9 пустой: импорт не падает, ответов 0",rep9.answers.K===0&&rep9.answers.S===0&&rep9.answers.P===0);
 ok("№9 пустой: имя НИОКР не затерто",r2.name===nm2);
 ok("№9 пустой: стоп-факторы не активны",r2.sf.every(x=>x==="не выявлено"));

 /* ===== C. фикстура: заполненный «офлайн» калькулятор ===== */
 const buff=fs.readFileSync(WS+"/scripts/tests/fixture_import.xlsx");
 const fwbk=API.xlsx18.parse(buff);
 ok("фикстура: читается, листов 14",fwbk.sheetNames.length===14);
 ok("фикстура: Паспорт C6 прочитан",fwbk.get("Паспорт_НИОКР","C6")==="НИОКР «Квантовый гетеродин»");
 const pvf=API.importPreview18(fwbk);
 ok("pv: заполнено S=30, K=15, P=10",pvf.q.S===30&&pvf.q.K===15&&pvf.q.P===10);
 ok("pv: уровень Средний",pvf.lvl==="S");
 ok("pv: стоп-факторы 1 подтв + 1 неопределённость",pvf.sf2===1&&pvf.sf1===1);
 ok("pv: AHP 9 пар ≠ 1, веса блоков Да",pvf.ahp===9&&pvf.ub===true);
 ok("pv: экономика строка описана",/r=0\.08/.test(pvf.econ)&&/потоки 2\/11/.test(pvf.econ)&&/LCC 2\/8/.test(pvf.econ));
 const r=S.list[0];
 const rep=API.importExcel9(fwbk,r);
 ok("rep: счётчики ответов 30/15/10",rep.answers.S===30&&rep.answers.K===15&&rep.answers.P===10);
 /* паспорт */
 ok("импорт: имя/шифр/контракт",r.name==="НИОКР «Квантовый гетеродин»"&&r.shifr==="НТ-2026-ГЕТ"&&r.contract==="ГЗ-114/26");
 ok("импорт: заказчик/исполнитель/кооперация",r.customer==="Минобрнауки России"&&r.executor==="НПО «Спектр»"&&r.coop==="ИТЦ РАН (соисполнитель)");
 ok("импорт: даты из serial",r.dates==="01.03.2024 — 31.12.2026"&&r.endDate==="2026-12-31");
 ok("импорт: стоимость/стадия/председатель/секретарь",r.cost==="75,5"&&r.stage==="ОКР"&&r.head==="Иванов И. И."&&r.secretary==="Петрова С. С.");
 ok("импорт: аудит-дата",r.auditDate==="01.10.2025");
 ok("импорт: активный уровень → Средний",r.lvl==="S");
 /* ответы */
 ok("q.S А1 = 5 из H (M — формула без кэша → консенсус)",r.q.S["А1"].a==="5"&&r.q.S["А1"].e==="ПРОТОКОЛ-17");
 ok("q.S А7 = 4.5 (средневзвешенное 5 и 4)",r.q.S["А7"].a==="4.5");
 ok("q.S А5 — обоснование из N10",r.q.S["А5"].j==="Обоснование для А5: отчёт по этапу 2");
 ok("q.K: 11×«4» и 4×«НД»",Object.values(r.q.K).filter(x=>x.a==="4").length===11&&Object.values(r.q.K).filter(x=>x.a==="НД").length===4);
 ok("q.P: код 1 = «3», код 2 = «НП» с обоснованием",r.q.P["1"].a==="3"&&r.q.P["2"].a==="НП"&&/отказался/.test(r.q.P["2"].j));
 ok("q.P: ровно 10 записей",Object.keys(r.q.P).length===10);
 /* стоп-факторы, зрелость */
 ok("sf[2] подтверждён, sf[5] неопределённость, sf[0] не выявлено",r.sf[2]==="подтверждено"&&r.sf[5]==="выявлена неопределённость"&&r.sf[0]==="не выявлено");
 ok("sfc[0] — вписанное обстоятельство",r.sfc[0]==="Отсутствие финансирования по госпрограмме");
 ok("мат: cur=4 plan=5 tests=Нет",r.mat.cur===4&&r.mat.plan===5&&r.mat.tests==="Нет");
 /* экономика */
 ok("econ: rate 0.08, T 5",r.econ.rate===0.08&&r.econ.T===5);
 ok("econ: lcc[0]=[50,0], flows[1]=[200,'']",r.econ.lcc[0][0]==="50"&&r.econ.lcc[0][1]==="0"&&r.econ.flows[1][0]==="200"&&r.econ.flows[1][1]==="");
 ok("econ: alt.c[0]=5, altW[0]=1",r.econ.alt.c[0]==="5"&&String(r.econ.altW[0])==="1");
 const ec=API.econ18(r);
 ok("econ18: NPV≈85.19",near(ec.npv,85.19,0.02));
 ok("econ18: BCR≈1.85 высокая",near(ec.bcr,1.85,0.01)&&/Высокая/.test(ec.bcrTxt));
 ok("econ18: LCC PV≈67.15",near(ec.lccPv,67.15,0.02));
 ok("econ18: Us 45/36/27/18, лучшая — продолжить",ec.Us[0]===45&&ec.Us[1]===36&&ec.best==="Продолжить разработку");
 /* AHP */
 ok("ahp18.dir: 2/4/2 импортированы",r.ahp18.dir.a12==="2"&&r.ahp18.dir.a13==="4"&&r.ahp18.dir.a23==="2");
 ok("ahp18.A: все 6 пар из листа AHP",r.ahp18.A.a12==="2"&&r.ahp18.A.a13==="4"&&r.ahp18.A.a14==="2"&&r.ahp18.A.a23==="2"&&r.ahp18.A.a24==="2"&&r.ahp18.A.a34==="2");
 ok("ahp18.V остался по умолчанию",r.ahp18.V.a12==="1");
 ok("ahp18.useBlocks=true из Полный C82",r.ahp18.useBlocks===true);
 const ah=API.ahp18cr(r);
 ok("ahp18cr: веса направлений 4/7, 2/7, 1/7",near(ah.dir.w[0],4/7,0.001)&&near(ah.dir.w[1],2/7,0.001)&&near(ah.dir.w[2],1/7,0.001));
 ok("ahp18cr: CR направлений = 0, ок",near(ah.dir.cr,0,1e-6)&&ah.ok===true);
 /* индексы и заключение */
 const ix=API.idx18(r,"S");
 const qs=API.Q18.levels.S.q.filter(x=>String(x.code).startsWith("А"));
 let num=0,den=0;qs.forEach(x=>{const c=r.q.S[x.code];if(c&&c.a!==""&&c.a!=="НП"&&c.a!=="НД"){num+=Number(x.w)*Number(c.a);den+=Number(x.w)*5;}});
 const IAexp=Math.round(num/den*1000)/10;
 ok("idx18: I_А = независимый пересчёт "+IAexp+"%",near(ix.IA,IAexp,0.051));
 ok("idx18: Coverage=100, min = min(I_А;I_Ц;I_Р)=98.8",ix.cov===100&&near(ix.min,98.8,0.051)&&near(ix.min,Math.min(ix.IA,ix.IC,ix.IR),0.001));
 const cc=API.concl18(r,"S");
 ok("concl18: при SF-2 → IV (даже с оценками 5/4.5)",cc.kind==="sf2"&&/IV/.test(cc.txt));
 /* ===== D. UI: карточка и модалка предпросмотра ===== */
 API.goTab("check");
 const tc=q("#tab-check");
 ok("карточка импорта на вкладке опросника",tc.textContent.includes("Импорт из Excel-калькулятора № 9"));
 ok("input #impXlsx18 присутствует",!!q("#impXlsx18"));
 API.importShow18(fwbk,API.cur(),"fixture_import.xlsx");
 const mb=q("#mbox");
 ok("модалка предпросмотра открыта",d.getElementById("moverlay").classList.contains("on")&&/Импорт из Excel-калькулятора № 9/.test(mb.textContent));
 ok("модалка: счётчики 30 из 30",/30 из 30/.test(mb.textContent)&&/15 из 15/.test(mb.textContent));
 const btns=qa("#mbox [data-mb]");
 click(btns[btns.length-1]); /* ✅ Импортировать */
 ok("после применения: модалка закрыта",!d.getElementById("moverlay").classList.contains("on"));
 const feed=JSON.stringify(S.feed||[]);
 ok("уведомление «Импорт выполнен»",/Импорт выполнен/.test(feed));
 const fl=API.flowInit(S.list[0]);
 ok("журнал потока: запись об импорте Excel",fl.log.some(x=>/Импорт Excel № 9/.test(x.what||"")));
 ok("селект А7 показал внесмасштабное 4.5",qa('#tab-check select[data-qa]').some(sel=>[...sel.options].some(o=>o.value==="4.5"&&o.selected)));
 ok("после импорта данные сохраняются в localStorage",(()=>{try{const raw=w.localStorage.getItem("niokr_audit_v1");return !!raw&&JSON.parse(raw).list[0].name==="НИОКР «Квантовый гетеродин»";}catch(e){return false;}})());
 API.goTab("dash");API.goTab("check");
 ok("дашборд и опросник перерисовываются без ошибок",winErr===0&&q("#tab-check").textContent.includes("Coverage"));
 /* ===== E. роль «Наблюдатель» — импорт запрещён ===== */
 S.role="view";API.goTab("check");
 ok("наблюдатель: без input импорта, с запретом",!q("#impXlsx18")&&q("#tab-check").textContent.includes("Наблюдатель"));
 S.role="admin";API.goTab("check");
 ok("администратор: input снова есть",!!q("#impXlsx18"));
 /* ===== F. idempotence: повторный импорт того же файла ===== */
 const repB=API.importExcel9(fwbk,API.state.list[0]);
 ok("повторный импорт даёт те же счётчики",repB.answers.S===30&&API.state.list[0].q.S["А7"].a==="4.5");
 /* ===== G. чужой файл → понятная ошибка ===== */
 let errTxt="";
 try{API.importExcel9({sheetNames:[],cells:{},get(){return null;}},S.list[0]);}catch(e){errTxt=(e&&e.message)||String(e);}
 ok("чужая книга → «не похож на калькулятор № 9»",/не похож/i.test(errTxt));
 ok("итог: нет ошибок окна",winErr===0);
}catch(e){fails++;console.log("FATAL",e&&e.stack||e);}
console.log("\nИТОГ IMPORT18: fails="+fails);
process.exit(fails?1:0);
},900);
