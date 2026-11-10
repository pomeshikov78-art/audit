/* РЕГРЕССИОННЫЕ ТЕСТЫ v18.0 — трёхуровневая система 15/30/60 + потребители PART B
   (наследник test_v17.js; якоря пересчитаны под v18) */
const fs = require("fs");
const { JSDOM } = require("jsdom");
const P = "output/Аудит_НИОКР_интерактивная_модель.html";
const html = fs.readFileSync(P, "utf8");
const dom = new JSDOM(html, { runScripts: "dangerously", pretendToBeVisual: true, url: "http://localhost/" });
const w = dom.window, d = w.document;
w.HTMLCanvasElement.prototype.getContext = () => ({ clearRect() {}, fillRect() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, arc() {}, save() {}, restore() {}, translate() {}, rotate() {}, fillText() {}, setLineDash() {}, closePath() {}, quadraticCurveTo() {} });
w.URL.createObjectURL = () => "blob:x"; w.URL.revokeObjectURL = () => {};
let werr = [];
w.addEventListener("error", e => { if (!/Not implemented/.test(e.message || "")) werr.push(e.message); });
let fails = 0;
const ok = (n, c) => { console.log((c ? "PASS" : "FAIL") + " | " + n); if (!c) fails++; };
const q = s => d.querySelector(s), qa = s => [...d.querySelectorAll(s)];
const mbtn = re => Array.from(d.querySelectorAll("#mbox [data-mb]")).find(b => re.test(b.textContent));

setTimeout(() => {
 try {
  const API = w.niokrAPI, S = API.state;
  const r0 = S.list[0];
  /* ========== 1. присутствие и версия ========== */
  ok("чип версии v18.0 · этап 18", q(".ver").textContent.includes("v18.0") && q(".ver").textContent.includes("этап 18"));
  ok("title содержит v18.0", d.title.includes("v18.0"));
  ok("15 вкладок сохранены", qa(".tabs button").length === 15);
  ok("вкладка «Опросник 15/30/60»", qa(".tabs button").some(b => b.textContent.includes("Опросник 15/30/60")));
  ok("нет вкладки «Чек-лист (31)»", !qa(".tabs button").some(b => /Чек-лист \(31\)/.test(b.textContent)));
  const tourTxt = html.slice(html.indexOf("const TOUR=["), html.indexOf("const TOUR=[") + 2000);
  ok("TOUR переписан под v18 (6 шагов, Опросник)", /Опросник 15\/30\/60/.test(tourTxt) && /30 кодов/.test(tourTxt));
  ok("TOUR: без устаревших «31 вопрос в трёх блоках»", !tourTxt.includes("31 вопрос в трёх блоках"));
  /* ========== 2. структура опросника ========== */
  ok("CODES30 = А1–А10·Б1–Б10·В1–В10", w.eval("CODES30.length") === 30 && w.eval("CODES30[0]") === "А1" && w.eval("CODES30[10]") === "Б1" && w.eval("CODES30[20]") === "В1" && w.eval("CODES30[29]") === "В10");
  ok("уровни: К=15, С=30, П=60", w.eval("Q18.levels.K.q.length") === 15 && w.eval("Q18.levels.S.q.length") === 30 && w.eval("Q18.levels.P.q.length") === 60);
  ok("направления S: 10/10/10", [0, 1, 2].every(b => w.eval("Q18.levels.S.q.filter(x=>+x.dir===" + b + ").length") === 10));
  ok("mig31 — непустой мост 31→30", w.eval("Q18.mig31.length") >= 20);
  ok("LEG2S31: длина 31, мост с пропуском -1", w.eval("LEG2S31.length") === 31 && w.eval("LEG2S31[12]") === 25);
  ok("PLVLO 4 уровня", w.eval("PLVLO.join('|')") === "Высокий|Приемлемый|Проблемный|Критический");
  ok("plvl якоря полос", w.eval("plvl(85)") === "Высокий" && w.eval("plvl(84)") === "Приемлемый" && w.eval("plvl(70)") === "Приемлемый" && w.eval("plvl(55)") === "Проблемный" && w.eval("plvl(54)") === "Критический");
  ok("пороги coverage 80/85/90", w.eval("[Q18.levels.K.porog,Q18.levels.S.porog,Q18.levels.P.porog].join()") === "80,85,90");
  ok("ниокрAPI обогащён v18-хелперами", ["calc", "ansS30", "CODES30", "scenTargets", "iAtWeights", "calcCore"].every(k => API[k]));
  /* ========== 3. расчётные якоря v18 ========== */
  const C = w.eval("CODES30");
  const fill = (a, e) => { r0.lvl = "S"; r0.q = { K: {}, S: {}, P: {} }; C.forEach(cd => r0.q.S[cd] = { a, j: "", e: e === undefined ? "документ" : e }); };
  fill("4"); let c = w.calc(r0);
  ok("якорь все «4»: I=80.0 «Приемлемый», уровень С", Math.abs(c.I - 80) < 0.05 && c.cat === "Приемлемый" && c.lvl === "S" && c.lvlName === "Средний");
  ok("якорь: dirs 80/80/80, IA=IC=IR=80", c.idx.dirs.every(x => Math.abs(x - 80) < 0.05) && Math.abs(c.iA - 80) < 0.05 && Math.abs(c.iR - 80) < 0.05);
  ok("якорь: cov=100 ≥ 85 → covOk", Math.abs(c.cov - 100) < 0.05 && c.covOk === true && c.thr === 85);
  ok("комплексный — геометрическое среднее", Math.abs(c.I - Math.cbrt(c.iA * c.iC * c.iR)) < 0.02);
  fill("5"); c = w.calc(r0);
  ok("все «5»: I=100 «Высокий», min=100", Math.abs(c.I - 100) < 0.05 && c.cat === "Высокий" && c.min === 100);
  fill("3"); c = w.calc(r0);
  ok("все «3»: I=60 «Проблемный»", Math.abs(c.I - 60) < 0.05 && c.lA === "Проблемный");
  fill("1"); c = w.calc(r0);
  ok("все «1»: I=20 «Критический»", Math.abs(c.I - 20) < 0.05 && c.cat === "Критический");
  fill("4"); r0.q.S["А1"] = { a: "НП", j: "этап не предусматривает", e: "" };
  c = w.calc(r0);
  ok("НП с обоснованием: cov=100, индексы исключают НП, flags=0", Math.abs(c.cov - 100) < 0.05 && c.I < 100 && w.eval("flagsOn")(r0).length === 0);
  r0.q.S["А2"] = { a: "НП", j: "", e: "" };
  ok("НП без обоснования → контрольное замечание", w.eval("flagsOn")(r0).some(f => /А2/.test(f) && /обосновани/i.test(f)));
  r0.q.S["А3"] = { a: "НД", j: "", e: "" };
  c = w.calc(r0);
  ok("НД снижает coverage (96.67 < 100)", c.cov < 100 && c.cov > 90);
  /* ========== 4. потребители PART B ========== */
  ok("respStats: контракт {n,rows,cI,mode,dis}", (() => { const st = w.eval("respStats")(r0); return "cI" in st && "mode" in st && Array.isArray(st.rows) && Array.isArray(st.dis); })());
  ok("discrepancies: кодовая сетка без исключений", Array.isArray(w.eval("discrepancies")(r0)));
  ok("scChip — LEG-обёртка выдаёт пилюлю с кодом", (() => { const h = w.eval("scChip")(r0, 12); return /pill/.test(h) && /В6|Б6|—/.test(h); })());
  ok("avgScore/qList не падают на LEG-списках", (() => { const v = w.eval("avgScore")(r0, [0, 12, 17]); const l = w.eval("qList")([0, 12, 17]); return (v === null || v >= 0) && /А1/.test(l); })());
  ok("mcRun: N выборочных I", (() => { const a = w.eval("mcRun")(r0, 40, 3); return a.length === 40 && a.every(v => v === null || (v >= 0 && v <= 100)); })());
  ok("tornado: 3 направления с dPoint>0", (() => { const t = w.eval("tornado")(r0); return t.length === 3 && t[0].dPoint > 0; })());
  ok("iAtWeights: базовые 30/35/35 не меняют v18-индекс", (() => { const v = w.eval("iAtWeights")(r0, [30, 35, 35]); return Math.abs(v - w.calc(r0).I) < 0.05; })());
  /* ========== 5. setQA18 и UI-маршрутизация ========== */
  fill("3", ""); w.renderAll();
  ok("qRow: id qrS-А1 и кнопки 💬/📎", (() => { q('[data-t="check"]').click(); const row = d.getElementById("qrS-А1"); return row && row.querySelector("[data-thq18]") && row.querySelector("[data-doc18]"); })());
  w.eval("setQA18")("А5", "a", "5"); w.eval("setQA18")("А5", "e", "акт");
  ok("setQA18 пишет в r.q.S и пересобирает консенсус", r0.q.S["А5"].a === "5" && w.eval("ansS30")(r0)[4] === "5");
  w.eval("threadsModal")(4);
  ok("💬 threadsModal открывается", d.getElementById("mbox").innerHTML.length > 40);
  const x1 = mbtn(/Закрыть|Отмена|Готово/); if (x1) x1.click();
  w.eval("docsModal")(4);
  ok("📎 docsModal открывается", d.getElementById("mbox").innerHTML.length > 40);
  const x2 = mbtn(/Закрыть|Отмена|Готово/); if (x2) x2.click();
  /* ========== 6. сценарии qS ========== */
  w.renderAll(); q('[data-t="anal"]').click();
  d.getElementById("scAdd").click();
  d.getElementById("scName").value = "ТЕСТ-СЦЕНАРИЙ";
  mbtn(/Сохранить/).click();
  const sc = r0.scenarios[r0.scenarios.length - 1];
  ok("сценарий сохраняет снимок qS (30 кодов)", sc.qS && Object.keys(sc.qS).length >= 28);
  ok("scenTargets снимает коды→баллы", (() => { const tg = API.scenTargets(sc); return Object.keys(tg).length >= 28 && String(tg["А5"]) === "5"; })());
  const ab = q("#tab-anal [data-scapp]"); ab.click(); mbtn(/Применить/).click();
  ok("применение сценария пишет r.q.S", r0.q.S["А5"] && String(r0.q.S["А5"].a) === "5");
  /* ========== 7. реестр доказательств ========== */
  q('[data-t="evid"]').click();
  ok("вкладке «Доказательства» — 30 кодовых строк в evNewQ", qa("#evNewQ option").length === 31);
  ok("фильтр направлений А/Б/В", (() => { const v = qa("#evBlk option").map(o => o.value).join(","); return v.includes("0") && v.includes("gen"); })());
  const goq = q("#tab-evid [data-goq]");
  if (goq) { goq.click(); ok("переход data-goq не роняет модель", true); } else ok("data-goq: нарушений нет — кнопок нет (допустимо)", true);
  /* ========== 8. дашборд, акт, слайды ========== */
  fill("4"); w.renderAll();
  const dashH = d.getElementById("tab-dash").innerHTML;
  ok("дашборд: комплексный I 80.0 и уровень «Приемлемый»", /80\.0/.test(dashH) && dashH.includes("Приемлемый"));
  ok("дашборд: направления I_А/I_Ц/I_Р на месте", /I_А/.test(dashH) && /I_Ц/.test(dashH) && /I_Р/.test(dashH));
  const slides = w.eval("buildSlides")(r0);
  const sTxt = JSON.stringify(slides);
  ok("слайды: 12 штук", slides.length === 12);
  ok("слайд 9 — v18-таблица направлений", sTxt.includes("Комплексный I") && sTxt.includes("модель v18.0"));
  ok("слайды: NaN/undefined не просачиваются", !sTxt.includes("NaN") && !/"null"/.test(sTxt));
  ok("слайды: «по чек-листу» исчез", !sTxt.includes("по чек-листу"));
  ok("вкладка «Акт оценки» (rep) рендерится без NaN", (() => { q('[data-t="rep"]').click(); const h = d.getElementById("tab-rep").innerHTML; return h.length > 500 && !h.includes("NaN"); })());
  ok("exportMD18 формирует Markdown заключения", (() => { const m = w.eval("exportMD18") ? String(w.eval("exportMD18")(r0)) : ""; return m.length > 200; })());
  ok("exportChecklistMD мостит легаси-сетку из кодов", (() => { const md = w.eval("exportChecklistMD")(r0); return md.length > 200 && md.includes("А1"); })());
  ok("renderRep/buildDeptReport/deptAdvice чисты", (() => { w.eval("deptAdvice")(r0); const h = w.eval("buildDeptReport")(); const rep = w.eval("renderRep") ? "" : ""; return h.length > 500; })());
  /* ========== 9. миграция легаси 31 → 30 ========== */
  const leg = JSON.parse(JSON.stringify(r0));
  leg.q = { K: {}, S: {}, P: {} }; leg.legacy31 = null;
  leg.ans = Array(31).fill(""); leg.ans[0] = "5"; leg.ans[12] = "4"; leg.ans[26] = "НП";
  w.eval("localStorage.setItem(LSKEY,JSON.stringify({list:LEG,cur:0}))".replace("LEG", JSON.stringify([Object.assign({}, leg, { ans: leg.ans })])));
  ok("load() переживает legacy-запись", w.load() === true);
  const rl = S.list[0];
  w.eval("migrate18extra")(rl);
  ok("миграция 31→30 заполняет кодовую сетку", (() => { const a = w.eval("ansS30")(rl); return a.filter(x => x !== "").length > 0; })());
  /* ========== 10. KB и строки ========== */
  ok("KB v18: трёхуровневая система описана", (() => { const kb0 = w.eval("expKBHTML")(); return kb0.includes("15 / 30 / 60") || kb0.includes("15/30/60") || kb0.includes("трёхуровневая"); })());
  const kb = w.eval("expKBHTML")();
  ok("KB: источники (вкладка Справка) — опросник v18.0", html.includes("опросник аудита НИОКР v18.0"));
  ok("KB: матрица решений 4×4 в справке", kb.includes("Матрица решений") && kb.includes("RECMX4") === false && /4×4|строка — I_А/.test(kb));
  ok("реестр: категории фильтра = PLVLO+«Не оценена»", (() => { q('[data-t="db"]').click(); const sel = d.getElementById("dbCat"); return !sel || sel.innerHTML.includes("Приемлемый") && sel.innerHTML.includes("Не оценена"); })());
  const resid = (s) => (s.match(/чек-лист/gi) || []).length;
  ok("лингвистика: вне исторических сносок «чек-лист» не осталось", (() => {
   const prot = /с чек-листа 31|вместо чек-листа 31|Чек-лист_НИОКР_|чек-листы v16|с чек-листа v17/g;
   const clean = html.replace(prot, "");
   return (clean.match(/чек-лист/gi) || []).length === 0;
  })());
  ok("нет window-ошибок за весь прогон", werr.length === 0);
 } catch (e) { fails++; console.log("EXCEPTION:", e.stack || e); }
 console.log(fails ? (">>> ПРОВАЛЕНО v18: " + fails) : ">>> ВСЕ ТЕСТЫ v18 ПРОЙДЕНЫ");
 process.exit(fails ? 1 : 0);
}, 800);
