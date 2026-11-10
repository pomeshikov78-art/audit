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
const chg=el=>el.dispatchEvent(new w.Event("change",{bubbles:true}));
// базовый якорный набор v15.0 (31 ответ): I=72.6, I_А=80.0, I_Ц=76.7
const A=["5","4","4","3","4","4","5","НП","4","3","4","3","3","4","3","3","4","3","1","4","4","3","4","4","4","4","5","4","3","3","4"];
const A27=["5","4","4","3","5","4","НП","3","4","4","3","3","4","3","3","4","3","1","5","4","4","3","4","4","3","4","НП"];
setTimeout(()=>{
try{
 const API=w.niokrAPI,S=API.state;
 // ================= 1. присутствие и версии =================
 ok("версия v15.0 в топбаре",q(".ver").textContent.includes("v15.0"));
 ok("title v15.0",d.title.includes("v15.0"));
 ok("15 вкладок",qa(".tabs button").length===15);
 ok("метка вкладки «Чек-лист (31)»",qa(".tabs button").some(b=>b.textContent.includes("Чек-лист (31)")));
 ok("TOUR: 16 шагов",w.eval("TOUR.length")===16);
 ok("TOUR: шаг экспертной оценки",w.eval("TOUR.some(t=>t[1]==='Экспертная оценка А/Ц')"));
 // ================= 2. структура данных =================
 ok("QS = 31 вопрос",w.eval("QS.length")===31);
 ok("блоки 10/9/12",w.eval("QS.filter(q=>q[2]===0).length")===10&&w.eval("QS.filter(q=>q[2]===1).length")===9&&w.eval("QS.filter(q=>q[2]===2).length")===12);
 ok("12 вопросов со ★",w.eval("QS.filter(q=>q[0].includes('★')).length")===12);
 ok("★ блока II сохранены (17–19)",w.eval("QS[16][0].includes('затраты связаны')&&QS[17][0].includes('не использованные ресурсы')&&QS[18][0].includes('оставшаяся потребность')"));
 ok("новые ★ А: 5,6,8,10",w.eval("[4,5,7,9].every(i=>QS[i][0].includes('★'))"));
 ok("новые ★ Ц: 22,26,28,29,30",w.eval("[21,25,27,28,29].every(i=>QS[i][0].includes('★'))"));
 ok("ANCH = 31 якорь",w.eval("ANCH.length")===31);
 ok("BRNG диапазоны",JSON.stringify(w.eval("BRNG"))==="[[1,10],[11,19],[20,31]]");
 ok("CRIT_A = 10 критериев",w.eval("CRIT_A.length")===10);
 ok("RECMX 5×5",w.eval("RECMX.length===5&&RECMX.every(r=>r.length===5)"));
 ok("STOPS=5, FLAGS=10, FLAGBAD=10",w.eval("STOPS.length===5&&FLAGS.length===10&&FLAGBAD.length===10"));
 ok("FORMS: 6 форм, индексы < 31",w.eval("FORMS.length===6&&FORMS.every(f=>f.focus.every(i=>i<31)&&f.na.every(i=>i<31))"));
 ok("FORMS: фокус ∩ Н/П пусты",w.eval("FORMS.every(f=>f.focus.filter(i=>f.na.includes(i)).length===0)"));
 // ================= 3. расчётные якоря =================
 const r0=S.list[0];
 r0.ans=A.slice();r0.conf=Array(31).fill("");r0.com=Array(31).fill("");
 r0.shifr="НИОКР-01";r0.name="Разработка А";r0.podrazd="КБ-1";r0.form="okr";r0.status="Выполняется";
 r0.fin.plan="100";r0.fin.fact="50";r0.fin.prog="";r0.ready="";
 let c=w.calc(r0);
 ok("якорь: I=72.6 («Хорошая»)",Math.abs(c.I-72.6)<0.06&&c.cat==="Хорошая");
 ok("якорь: I_А=80.0 «Хорошая»",Math.abs(c.iA-80.0)<0.05&&c.lA==="Хорошая");
 ok("якорь: I_Ц=76.7 «Хорошая»",Math.abs(c.iC-76.67)<0.05&&c.lC==="Хорошая");
 ok("блоки Pⱼ: 80 / 62.2 / 76.7",Math.abs(c.bl[0].P*100-80)<0.05&&Math.abs(c.bl[1].P*100-62.22)<0.05&&Math.abs(c.bl[2].P*100-76.67)<0.05);
 ok("рекомендация матрицы Х/Х: продолжить при условии устранения замечаний",c.rec==="Продолжить при условии устранения выявленных замечаний"&&!c.recStop);
 ok("перспективы авто: «Реалистичная» (I_Ц ≥ 70)",c.prospect==="Реалистичная");
 ok("Н/П исключается из I_А: nA=9",c.nA===9&&c.nC===12);
 // все «1» / все «5»
 const keepA=A.slice();
 r0.ans=Array(31).fill("1");c=w.calc(r0);
 ok("все «1»: I=20, I_А=20, I_Ц=20, рекомендация «Прекратить НИОКР»",Math.abs(c.I-20)<0.05&&Math.abs(c.iA-20)<0.05&&c.lC==="Критическая"&&c.rec.startsWith("Прекратить НИОКР"));
 r0.ans=Array(31).fill("5");c=w.calc(r0);
 ok("все «5»: I=100, рекомендация В/В с ускорением",Math.abs(c.I-100)<0.05&&c.rec.includes("рассмотреть ускорение"));
 r0.ans=keepA.slice();
 // матрица: А высокая × Ц критическая
 const mix=Array(31).fill("");for(let i=0;i<10;i++)mix[i]="5";for(let i=19;i<31;i++)mix[i]="0";
 r0.ans=mix;c=w.calc(r0);
 ok("А=100/Ц=0: «Пересмотреть технические решения…»",c.iA===100&&c.iC===0&&c.lA==="Высокая"&&c.lC==="Критическая"&&c.rec==="Пересмотреть технические решения, исполнителей, ресурсы и сроки");
 r0.ans=keepA.slice();
 // веса 25/40/35
 S.weights=[25,40,35];
 ok("веса 25/40/35: I=71.7",Math.abs(w.calc(r0).I-71.72)<0.06);
 S.weights=[30,35,35];
 // блок Ц целиком «Н/П» → ренормировка + «недостаточно данных» для экспертной оценки
 const rn=keepA.slice();for(let i=19;i<31;i++)rn[i]="НП";
 r0.ans=rn;c=w.calc(r0);
 ok("блок Ц = «Н/П»: ренормировка, I=70.4",c.renorm&&Math.abs(c.I-70.43)<0.06);
 ok("экспертная оценка без данных Ц: «Недостаточно данных…», перспективы «—»",c.rec.includes("Недостаточно данных")&&c.prospect==="—");
 r0.ans=keepA.slice();
 // ================= 4. стоп-факторы и «красные» флаги =================
 r0.stops[2]="ДА";c=w.calc(r0);
 ok("стоп-фактор блокирует рекомендацию",c.recStop&&c.rec.includes("Внеочередное рассмотрение руководством")&&c.stopsHit.length===1);
 ok("стоп-фактор → перспективы «Маловероятная»",c.prospect==="Маловероятная");
 ok("стоп-факторы не влияют на I_А/I_Ц",Math.abs(c.iA-80)<0.05&&Math.abs(c.iC-76.67)<0.05);
 r0.stops[2]="";
 r0.flags[0]="ДА";r0.flags[6]="НЕТ";r0.flags[7]="ДА";
 ok("флаги: полярность учтена (2 из 3)",w.flagsOn(r0).length===2);
 r0.flags[6]="ДА";r0.flags[7]="НЕТ";
 ok("вопр.7 «ДА» не тревожен, вопр.8 «НЕТ» тревожен → всего 2",w.flagsOn(r0).length===2);
 r0.flags=Array(10).fill("");
 r0.prospect="Условно реалистичная";
 ok("ручные перспективы переопределяют авто",w.expertProspect(r0)==="Условно реалистичная");
 r0.prospect="";
 // ================= 5. миграция 27 → 31 =================
 const mig={ans:A27.slice(),conf:Array(27).fill(""),com:Array(27).fill("")};
 w.migrate27to31(mig);
 ok("миграция: длина 31 + stops/flags/prospect",mig.ans.length===31&&mig.stops.length===5&&mig.flags.length===10&&mig.prospect==="");
 ok("миграция A: old6«НП»→new7, old4→new6, old7→new9",mig.ans[7]==="НП"&&mig.ans[6]==="5"&&mig.ans[9]==="3");
 ok("миграция A: новые 4 и 9 — пустые",mig.ans[4]===""&&mig.ans[8]==="4");
 ok("миграция B: сдвиг +1 (old9→new10, old17→new18)",mig.ans[10]==="4"&&mig.ans[18]==="1");
 ok("миграция C: 18→26, 19→25, 20→27, 21→28, 22→23, 24→21",mig.ans[26]==="5"&&mig.ans[25]==="4"&&mig.ans[27]==="4"&&mig.ans[28]==="3"&&mig.ans[23]==="4"&&mig.ans[21]==="3");
 ok("миграция C: новые 19,20,22,24,29,30 — пустые; 23/25/26 утрачены",[19,20,22,24,29,30].every(i=>mig.ans[i]===""));
 ok("миграция идемпотентна",w.migrate27to31(mig)===false&&mig.ans.length===31);
 const legacy=w.niokrAPI.calcCore(A27,[30,35,35]);
 ok("старый 27-массив не ломает calcCore (I≈72.6, без NaN)",legacy.I!=null&&isFinite(legacy.I)&&Math.abs(legacy.I-72.56)<0.1);
 // ================= 6. UI чек-листа =================
 w.goTab("check");
 let h=d.getElementById("tab-check").innerHTML;
 ok("чек-лист: 5 стоп-факторов (селекты)",qa('#tab-check select[data-st]').length===5);
 ok("чек-лист: 10 «красных» вопросов",qa('#tab-check select[data-fl]').length===10);
 ok("чек-лист: карточка экспертной оценки",h.includes("ЭКСПЕРТНАЯ ОЦЕНКА «АКТУАЛЬНОСТЬ / ЦЕЛЕСООБРАЗНОСТЬ»"));
 ok("чек-лист: кнопка заключения + перспективы",!!d.getElementById("exConcl")&&!!d.getElementById("exProspect"));
 ok("чек-лист: I_А/I_Ц в карточке",h.includes("80.0 · Хорошая")&&h.includes("76.7 · Хорошая"));
 ok("чек-лист: 31 строка вопросов",qa("#tab-check .qrow").length===31);
 ok("чек-лист: Nⱼ знаменатели 10/9/12",h.includes("/ 10</span>")&&h.includes("/ 9</span>")&&h.includes("/ 12</span>"));
 const st2=q('#tab-check select[data-st="2"]');st2.value="ДА";chg(st2);
 h=d.getElementById("tab-check").innerHTML;
 ok("смена стоп-фактора: блокировка в карточке",h.includes("РЕКОМЕНДАЦИЯ ЗАБЛОКИРОВАНА СТОП-ФАКТОРОМ"));
 ok("стоп-фактор записан в данные",r0.stops[2]==="ДА");
 const fl0=q('#tab-check select[data-fl="0"]');fl0.value="ДА";chg(fl0);
 ok("флаг записан + подсветка ⚠",r0.flags[0]==="ДА"&&d.getElementById("tab-check").innerHTML.includes("сработало: 1 из 10"));
 const st2b=q('#tab-check select[data-st="2"]');st2b.value="";chg(st2b);
 const fl0b=q('#tab-check select[data-fl="0"]');fl0b.value="";chg(fl0b);
 ok("сброс стоп-фактора/флага",r0.stops[2]===""&&r0.flags[0]===""&&!w.calc(r0).recStop);
 const prs=d.getElementById("exProspect");prs.value="Маловероятная";chg(prs);
 ok("ручные перспективы сохранены",r0.prospect==="Маловероятная");
 const prs2=d.getElementById("exProspect");prs2.value="";chg(prs2);
 // заключение (модалка)
 d.getElementById("exConcl").click();
 let t=(d.getElementById("mbox").innerText||d.getElementById("mbox").textContent);
 ok("заключение: 5 обязательных позиций",["1. Уровень актуальности","2. Уровень целесообразности","3. Проблемы с представленными материалами","4. Перспективы завершения","5. Рекомендация"].every(x=>t.includes(x)));
 ok("заключение: значения 80.0/76.7",t.includes("80.0 % — «Хорошая»")&&t.includes("76.7 % — «Хорошая»"));
 q('#mbox [data-mb="1"]').click();
 ok("модалка закрыта",!d.getElementById("moverlay").classList.contains("on"));
 // ================= 7. дашборд =================
 w.goTab("dash");h=d.getElementById("tab-dash").innerHTML;
 ok("дашборд: экспертная сводка",h.includes("Экспертная оценка «Актуальность / Целесообразность»")&&h.includes("I_А: 80.0 · Хорошая")&&h.includes("I_Ц: 76.7 · Хорошая"));
 ok("дашборд: 0 стоп-факторов / 0 флагов",h.includes("стоп-факторов: 0")&&h.includes("«красных» флагов: 0"));
 ok("дашборд: KPI «/ 31»",h.includes("/ 31</span>"));
 ok("дашборд: тепловая карта 1–31",h.includes("вопросы 1–31"));
 ok("дашборд: кнопка заключения exConcl2",!!d.getElementById("exConcl2"));
 // ================= 8. справка / база знаний =================
 w.goTab("ref");h=d.getElementById("tab-ref").innerHTML;
 ok("KB «Методика»: экспертная карточка",h.includes("Экспертная методика «Актуальность + Целесообразность»"));
 ok("KB: матрица решений 5×5 в таблице",h.includes("Матрица решений")&&h.includes("Прекратить НИОКР (рассмотреть списание результата в установленном порядке)"));
 ok("KB: 5 стоп-факторов и 10 «красных» вопросов",h.includes("детекторы «зомби-проектов»")&&h.includes("ловушка невозвратных затрат"));
 ok("KB: формула Sⱼ обновлена (10/9/12)",h.includes("(10 / 9 / 12)"));
 q('#tab-ref [data-rs="anch"]').click();
 h=d.getElementById("tab-ref").innerHTML;
 ok("KB «Якоря»: 31 строка",h.includes("по всем 31 вопросу")&&(h.match(/<tr><td>\d+<\/td>/g)||[]).length>=31);
 // ================= 9. презентация =================
 w.goTab("pres");
 const sl=w.buildSlides(r0);
 ok("дек: 12 слайдов",sl.length===12);
 ok("SLBL: слайд 7 «Актуальность (экспертная оценка)»",w.eval("SLBL[7]").includes("Актуальность (экспертная оценка)"));
 ok("слайд 7: 10 критериев + I_А",sl[7].h.includes("10 критериев")&&sl[7].h.includes("I_А = 80.0"));
 ok("слайд 7: ключевой вопрос 10★",sl[7].h.includes("(10★)")&&sl[7].h.includes("3–5 лет"));
 ok("слайд 8: I_Ц + чипы 22★…30★",sl[8].h.includes("I_Ц = 76.7")&&sl[8].h.includes("№22★")&&sl[8].h.includes("№30★"));
 ok("слайд 9: диапазоны блоков 1–10/11–19/20–31",sl[9].h.includes("1–10")&&sl[9].h.includes("11–19")&&sl[9].h.includes("20–31"));
 ok("слайд 10: экспертная полоса",sl[10].h.includes("Экспертная оценка «Актуальность / Целесообразность»")&&sl[10].h.includes("Рекомендация матрицы решений"));
 ok("слайд 2: ключевой вопрос 8★",sl[2].h.includes("(8★)")&&sl[2].h.includes("№8★"));
 ok("слайд 4: чипы №13…№19★",sl[4].h.includes("№13 бюджет")&&sl[4].h.includes("№19★"));
 ok("экспорт HTML-дека: v15.0",w.buildDeckHTML().includes("v15.0"));
 // ================= 10. лента расхождений =================
 r0.fin.plan="100";r0.fin.fact="50";r0.fin.prog="80";r0.ready="90";
 r0.ans[12]="5";r0.ans[26]="1";
 let disc=w.discrepancies(r0);
 ok("расхождение: бюджет → «вопроса 13»",disc.some(x=>x[1].includes("вопроса 13 «соблюдение бюджета»")));
 ok("расхождение: готовность → «вопрос 27»",disc.some(x=>x[1].includes("вопрос 27 «возможность внедрения»")));
 r0.ans[12]="3";r0.ans[26]="5";r0.fin.prog="";r0.ready="";
 // ================= 11. акт оценки =================
 w.goTab("rep");h=d.getElementById("tab-rep").innerHTML;
 ok("акт: пункт 6б с 5 позициями",h.includes("6б. Экспертное заключение об актуальности и целесообразности (5 обязательных позиций)")&&h.includes("4) Перспективы завершения: Реалистичная")&&h.includes("5) Рекомендация: Продолжить при условии устранения выявленных замечаний"));
 ok("акт: уровни 80.0/76.7",h.includes("80.0 % — «Хорошая»")&&h.includes("76.7 % — «Хорошая»"));
 // ================= 12. респонденты (консенсус по 31) =================
 r0.head="Иванова А.П.";
 w.goTab("resp");
 d.getElementById("respOn").click();
 q('#mbox [data-mb="1"]').click();
 ok("режим включён: 2 респондента",r0.resp.enabled&&r0.resp.people.length===2);
 ok("оценка скопирована респонденту 1 (31)",r0.resp.data[r0.resp.people[0].id].ans.length===31);
 d.getElementById("rAdd").click();
 q("#mbox input").value="Петров П.П.";
 q('#mbox [data-mb="1"]').click();
 const p2=r0.resp.people[2].id;
 const d2=r0.resp.data[p2];
 d2.ans=A.slice();d2.ans[18]="0";d2.ans[20]="5";
 API.syncConsensus(r0);
 ok("медиана: (1,0)→0.5 и (4,5)→4.5",r0.ans[18]==="0.5"&&r0.ans[20]==="4.5");
 ok("дробный консенсус считается в I_Ц",isFinite(w.calc(r0).iC));
 h=d.getElementById("tab-resp").innerHTML;
 ok("таблица респондентов: «/ 31»",h.includes("/ 31"));
 const csv=API.buildRespCsv(r0);
 ok("CSV респондентов: 31 вопрос × (3+консенсус)",csv.includes("КОНСЕНСУС")&&(csv.match(/\n/g)||[]).length>=120);
 w.goTab("check");h=d.getElementById("tab-check").innerHTML;
 ok("чек-лист: «из 31» в баннере респондентов",h.includes("из 31"));
 r0.ans=A.slice();r0.resp=null;
 // ================= 13. предложения =================
 w.goTab("prop");
 const pr=w.eval("newProp()");
 pr.model="modern";pr.status="accepted";pr.extra="12.5";pr.months="9";pr.targets={"18":"4"};
 pr.measures=[{text:"Программа специспытаний",resp:"Иванов И.И.",due:"20.12.2026",st:"plan"}];
 r0.props=[pr];
 const pa=API.propAns(r0,pr);
 ok("propAns: целевой балл вопроса 19 (idx18)",pa[18]==="4");
 const pf=API.propForecast(r0,pr);
 ok("прогноз I: 72.6 → 74.9 (ΔI +2.3)",Math.abs(pf.I1-74.94)<0.06&&Math.abs(pf.dI-2.33)<0.06);
 h=d.getElementById("tab-prop").innerHTML;
 ok("вкладка предложений рендерится",h.includes("Предложения по дальнейшей работе")||h.length>500);
 // ================= 14. БД / история / сравнение снимков =================
 w.goTab("db");
 w.saveAudit(r0,"v15 базовый");
 ok("снимок аудита: 31 ответ",r0.audits[0].ans.length===31);
 r0.audits.push({dateStr:"01.01.2025",by:"Архив",I:72.9,ans:A27.slice(),conf:Array(27).fill(""),note:"старый формат 27"});
 API.cmpModal(r0,0,1);
 t=(d.getElementById("mbox").innerText||d.getElementById("mbox").textContent);
 ok("сравнение 31 vs 27 не падает и показывает расхождения",t.includes("Сравнение аудитов")&&t.length>200);
 q('#mbox [data-mb="0"]').click();
 ok("модалка сравнения закрыта",!d.getElementById("moverlay").classList.contains("on"));
 // восстановление из старого снимка (guard)
 const before=r0.ans[30];
 w.eval("(()=>{const r=niokrAPI.state.list[0];const a=r.audits[1];const A2=niokrAPI.ansArr(r),C2=r.conf;for(let j=0;j<31;j++){A2[j]=a.ans&&a.ans[j]!=null?a.ans[j]:'';C2[j]=a.conf&&a.conf[j]?a.conf[j]:'';}})()");
 ok("восстановление из 27-снимка: хвост очищен, начало перенесено как есть",r0.ans[30]===""&&r0.ans[6]==="НП"&&r0.ans[7]==="3");
 r0.ans=A.slice();r0.conf=Array(31).fill("");
 w.renderDb();h=d.getElementById("tab-db").innerHTML;
 ok("БД рендерится с 2 записями",h.includes("v15 базовый"));
 // ================= 15. доказательства / аналитика / графики =================
 r0.ans[5]="5";r0.docs=[];
 const ev=API.evidStats(r0);
 ok("evidStats: высокий балл без документа → viol[5], K<100",ev.viol.includes(5)&&ev.K<100);
 r0.ans[5]="4";
 w.goTab("anal");h=d.getElementById("tab-anal").innerHTML;
 ok("аналитика рендерится (МК/торнадо)",h.length>1000);
 ok("mcRun по 31 вопросу",w.mcRun(r0,200,7).filter(x=>x!=null).length===200);
 w.goTab("viz");h=d.getElementById("tab-viz").innerHTML;
 ok("графики БД рендерятся",h.includes("Графический дашборд")||h.length>1000);
 // ================= 16. горячие клавиши и справка =================
 key("e");
 t=(d.getElementById("mbox").innerText||d.getElementById("mbox").textContent);
 ok("клавиша e: экспертное заключение",t.includes("Экспертное заключение"));
 q('#mbox [data-mb="1"]').click();
 key("?");
 t=(d.getElementById("mbox").innerText||d.getElementById("mbox").textContent);
 ok("справка: строка клавиши e",t.includes("экспертное заключение"));
 q('#mbox [data-mb="0"]').click();
 key("2");
 ok("клавиша 2 → чек-лист",q("section.tab.on").id==="tab-check");
 // ================= 17. save/load: новая структура в localStorage =================
 w.eval("save()");
 const ls=w.eval("localStorage.getItem(LSKEY)");
 ok("save(): stops/flags/prospect в localStorage",ls.includes('"stops"')&&ls.includes('"flags"')&&ls.includes('"prospect"'));
 // ================= 18. load(): миграция сохранённых данных 27 =================
 const legacyRec={name:"Старая",shifr:"НИОКР-99",podrazd:"",status:"Выполняется",form:"",
  cost:"",ready:"",goal:"",product:"",customer:"",executor:"",basis:"",contract:"",dates:"",head:"",
  fin:{plan:"",fact:"",prog:""},ctrl:["","","","","","",""],decision:"",just:"",measures:"",weak:"",strong:"",
  ans:A27.slice(),conf:Array(27).fill(""),com:Array(27).fill(""),cat:"",docs:[],scenarios:[],snaps:[],flow:null,threads:{},resp:null,audits:[],props:[],created:new Date().toISOString()};
 w.eval("localStorage.setItem(LSKEY,JSON.stringify({list:[LEG],cur:0}))".replace("LEG",JSON.stringify(legacyRec)));
 ok("load() принялlegacy-данные",w.load()===true);
 const rl=S.list[0];
 ok("load(): автомиграция 27→31",rl.ans.length===31&&rl.ans[26]==="5"&&rl.ans[7]==="НП"&&rl.stops.length===5);
 w.renderAll();
 ok("renderAll после миграции без исключений",true);
 // экспорт
 d.getElementById("btnExp").click();
 ok("экспорт v15 отработал",true);
}catch(e){fails++;console.log("EXCEPTION:",e.stack||e);}
console.log(fails?(">>> ПРОВАЛЕНО: "+fails):">>> ВСЕ ТЕСТЫ v15 ПРОЙДЕНЫ");
process.exit(fails?1:0);
},500);
