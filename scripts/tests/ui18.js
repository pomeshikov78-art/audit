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
const q=s=>d.querySelector(s),qa=s=>[...d.querySelectorAll(s)];
const chg=el=>el.dispatchEvent(new w.Event("change",{bubbles:true}));
const click=el=>el.dispatchEvent(new w.MouseEvent("click",{bubbles:true}));
// заглушка скачивания
let dlName=null;
w.HTMLAnchorElement.prototype.click=function(){dlName=this.download;};
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state,r0=S.list[0];
 ok("нет ошибок окна",winErr===0);
 // ---- версии
 ok("title v18.0",d.title.includes("v18.0"));
 ok(".ver = v18.0 · этап 18",q(".ver").textContent.includes("v18.0")&&q(".ver").textContent.includes("этап 18"));
 ok("welcome v18.0",q("#tab-home")?q("#tab-home").textContent.includes("v18.0"):d.body.textContent.includes("Новое в v18.0"));
 ok("подвалы v18.0 (≥8)",(d.body.innerHTML.match(/аудита НИОКР v18\.0/g)||[]).length>=8);
 ok("метка вкладки «Опросник 15/30/60»",qa(".tabs button").some(b=>b.textContent.includes("Опросник 15/30/60")));
 // ---- дашборд
 ok("дашборд: KPI «КОМПЛЕКСНЫЙ ИНДЕКС»",q("#tab-dash").textContent.includes("КОМПЛЕКСНЫЙ ИНДЕКС"));
 ok("дашборд: карточка «Экспертная оценка — опросник»",q("#tab-dash").textContent.includes("Экспертная оценка"));
 // ---- вкладка опросника
 API.goTab("check");
 const tc=q("#tab-check");
 ok("опросник: переключатель уровней (3 кнопки)",qa("#tab-check [data-lvl]").length===3);
 ok("опросник:по умолчанию уровень «Средний» — 30 строк вопросов",qa("#tab-check select[data-qa]").length===30);
 ok("опросник: направления А/Б/В в заголовках",tc.textContent.includes("А. АКТУАЛЬНОСТЬ")&&tc.textContent.includes("Б. ЦЕЛЕСООБРАЗНОСТЬ")&&tc.textContent.includes("В. РЕЗУЛЬТАТИВНОСТЬ"));
 ok("опросник: сводка (I_А/I_Ц/I_Р, Coverage, counters)",tc.textContent.includes("Coverage")&&tc.textContent.includes("Контрольных замечаний")&&tc.textContent.includes("Балл ≤ 2"));
 ok("опросник: заключение «Недостаточно данных» (пусто)",tc.textContent.includes("Недостаточно данных для достоверной оценки"));
 ok("опросник: стоп-факторы (10 строк)",qa("#tab-check select[data-sf]").length===10);
 ok("опросник: зрелость M0–M9 (селекты cur/plan/tests)",qa("#tab-check select[data-mat]").length===3);
 ok("опросник: экономика (ставка, потоки, LCC, альтернативы)",qa("#tab-check [data-econ]").length>=1+22+16+9+36);
 ok("опросник: кнопки Ф-06 и MD",!!q("#q18Concl")&&!!q("#q18Md"));
 // ---- ответ на вопрос (Средний А1 = 5)
 const sel=q('#tab-check select[data-qa="А1"]');
 sel.value="5";chg(sel);
 ok("ответ записан: r.q.S['А1'].a==='5'",r0.q.S["А1"].a==="5");
 ok("сводка: отвечено 1 / 30",q("#tab-check").textContent.includes("1")&&API.idx18(r0,"S").nAns===1);
 ok("замечание «3–5 без доказательства» (ctrl=1)",API.idx18(r0,"S").ctrl===1);
 // ---- НП без обоснования → модалка
 const sel2=q('#tab-check select[data-qa="А2"]');
 sel2.value="НП";chg(sel2);
 ok("НП без обоснования → модальное окно",!!q("#npJust18"));
 const nj=q("#npJust18");nj.value="Не применимо: стадия НИР";
 // кнопка «Сохранить обоснование»
 const btns=[...d.querySelectorAll(".modal button, #modalRoot button, .modal-wrap button")];
 const saveBtn=btns.find(b=>b.textContent.includes("Сохранить обоснование"));
 ok("модалка: кнопка «Сохранить обоснование» найдена",!!saveBtn);
 if(saveBtn){click(saveBtn);}
 ok("обоснование «НП» сохранено",r0.q.S["А2"]&&r0.q.S["А2"].j==="Не применимо: стадия НИР");
 ok("НП исключён из Coverage-знаменателя",Math.abs(API.idx18(r0,"S").cov-Math.round(1.2/ (30-1) *1000)/10)<0.6);
 // ---- переключение уровня на «Полный»
 const pb=qa("#tab-check [data-lvl]").find(b=>b.dataset.lvl==="P");click(pb);
 ok("уровень переключён на «Полный»",r0.lvl==="P");
 ok("полный: 60 строк вопросов",qa("#tab-check select[data-qa]").length===60);
 ok("полный: 12 блоков с весами",q("#tab-check").textContent.includes("Блок")&&q("#tab-check").textContent.includes("«Эффект и управление результатами»"));
 // ---- возврат на Средний и заполнение всех 5 → заключение I
 qa("#tab-check [data-lvl]").find(b=>b.dataset.lvl==="S").click?click(qa("#tab-check [data-lvl]").find(b=>b.dataset.lvl==="S")):null;
 API.Q18.levels.S.q.forEach(qq=>{r0.q.S[qq.code]={a:"5",j:"",e:"Протокол"};});
 API.goTab("dash");API.goTab("check");
 ok("все 5: заключение «I. Положительное — продолжить НИОКР»",q("#tab-check").textContent.includes("I. Положительное — продолжить НИОКР"));
 // ---- стоп-фактор SF-2 → блокировка
 const sfSel=qa("#tab-check select[data-sf]")[0];
 sfSel.value="подтверждено";chg(sfSel);
 ok("SF-2: r.sf[0]='подтверждено'",r0.sf[0]==="подтверждено");
 ok("SF-2: заключение «IV. Отрицательное / эскалация» несмотря на все 100%",q("#tab-check").textContent.includes("IV. Отрицательное / эскалация"));
 ok("SF-2: дашборд-карточка с эскалацией",r0.sf.some(v=>API.sfVal(v)===2));
 const sfNew=qa("#tab-check select[data-sf]")[0];sfNew.value="не выявлено";chg(sfNew);
 ok("SF-2 снят → снова «I. Положительное»",API.concl18(r0,"S").txt.startsWith("I. Положительное"));
 // ---- AHP вкладка
 API.goTab("wgt");
 const tw=q("#tab-wgt");
 ok("AHP: 4 матрицы (направления + А/Б/В)",tw.textContent.includes("направления оценки НИОКР (3×3)")&&(tw.innerHTML.match(/4×4/g)||[]).length===3);
 ok("AHP: CR/λmax/CI отображаются",tw.textContent.includes("CR = 0.000")&&tw.textContent.includes("λmax"));
 ok("AHP: переключатель весов для «Полного»",!!q("#ahpUse18"));
 const s12=qa('#tab-wgt select[data-ahp="dir"]')[0];s12.value="2";chg(s12);
 const s13=qa('#tab-wgt select[data-ahp="dir"]')[1];s13.value="4";chg(s13);
 const s23=qa('#tab-wgt select[data-ahp="dir"]')[2];s23.value="2";chg(s23);
 const ad=API.ahp18cr(r0).dir;
 ok("AHP 2/4/2: веса 57.1/28.6/14.3, λmax=3, CR=0",Math.abs(ad.w[0]-4/7)<0.001&&Math.abs(ad.lmax-3)<1e-6&&Math.abs(ad.cr)<1e-6&&q("#tab-wgt").textContent.includes("57.1%"));
 const ub=q("#ahpUse18");ub.checked=true;chg(ub);
 ok("переключатель AHP → r.ahp18.useBlocks=true",r0.ahp18.useBlocks===true);
 r0.ahp18.dir={a12:"1",a13:"1",a23:"1"};
 // ---- экономика: ввод через UI
 API.goTab("check");
 const rateEl=q('#tab-check [data-econ="rate"]');rateEl.value="0";chg(rateEl);
 const b0=q('#tab-check [data-econ="fb0"]');b0.value="200";chg(b0);
 const c0=q('#tab-check [data-econ="fc0"]');c0.value="100";chg(c0);
 const e2=API.econ18(r0);
 ok("экономика UI: NPV=100, BCR=2 → «Высокая экономическая привлекательность»",Math.abs(e2.npv-100)<0.011&&Math.abs(e2.bcr-2)<0.011&&q("#tab-check").textContent.includes("Высокая экономическая привлекательность"));
 // ---- Ф-06 модалка
 const bc=q("#q18Concl");click(bc);
 const mt=d.body.textContent;
 ok("Ф-06: модалка с 10+ позициями",mt.includes("Ф-06")&&mt.includes("Индекс актуальности")&&mt.includes("ЗАКЛЮЧЕНИЕ:"));
 // закрыть модалку
 const closeB=[...d.querySelectorAll("button")].find(b=>b.textContent.trim()==="Закрыть"&&b.offsetParent!==null);
 if(closeB)click(closeB);
 // ---- MD экспорт
 API.goTab("check");
 const bm=q("#q18Md");dlName=null;click(bm);
 ok("MD: скачивание «Вопросник_НИОКР_…_Средний_30.md»",dlName&&dlName.includes("Средний_30")&&dlName.endsWith(".md"));
 const md=API.exportMD18(r0);
 ok("MD: структура (направления, сводка, заключение)",md.includes("# ОПРОСНИК АУДИТА НИОКР")&&md.includes("## Направление «Актуальность (I_А)»")&&md.includes("## СВОДНЫЕ ПОКАЗАТЕЛИ")&&md.includes("I. Положительное — продолжить НИОКР"));
 // ---- хоткей e → Ф-06
 d.dispatchEvent(new w.KeyboardEvent("keydown",{key:"e",bubbles:true}));
 ok("хоткей e → модалка Ф-06",d.body.textContent.includes("Экспертное заключение — уровень"));
}catch(e){console.log("EXCEPTION:",e&&e.stack||e);fails++;}
console.log("\nИТОГ UI18: fails="+fails);
process.exit(fails?1:0);
},1500);
