/* Пробут PART B: jsdom + клик-тест всех затронутых поверхностей */
const fs = require("fs"), { JSDOM } = require("jsdom");
const html = fs.readFileSync("output/Аудит_НИОКР_интерактивная_модель.html", "utf-8");
const errors = [];
const dom = new JSDOM(html, { runScripts: "dangerously", pretendToBeVisual: true, url: "http://localhost/" });
const w = dom.window;
w.addEventListener("error", e => errors.push("window: " + e.message));
const realErr = w.console.error; w.console.error = (...a) => { errors.push("console: " + a.join(" ")); };
const mbtn = re => Array.from(w.document.querySelectorAll("#mbox [data-mb]")).find(b => re.test(b.textContent));
function step(name, fn) {
  try { fn(); console.log("PASS |", name); }
  catch (e) { errors.push(name + ": " + e.message); console.log("FAIL |", name, "|", e.message); }
}
w.document.addEventListener("DOMContentLoaded", () => {});
setTimeout(() => {
  const doc = w.document, $ = q => doc.querySelector(q); const doc2 = { $$ : q => Array.from(doc.querySelectorAll(q)) };
  const api = w.niokrAPI || w.window.niokrAPI;
  step("boot: renderAll filled", () => { if (!doc.getElementById("tab-dash").innerHTML.length) throw new Error("empty dash"); });
  step("API surface", () => { ["calc", "ansS30", "CODES30", "scenTargets", "iAtWeights"].forEach(k => { if (typeof api[k] !== "function" && !api[k]) throw new Error("missing " + k); }); });
  step("tab dash", () => doc.querySelector('[data-t="dash"]').click());
  step("tab check", () => { doc.querySelector('[data-t="check"]').click(); const h = doc.getElementById("tab-check").innerHTML; if (!/qrow/.test(h)) throw new Error("no qrows"); if (!/А1/.test(h)) throw new Error("no codes"); });
  step("tab evid", () => doc.querySelector('[data-t="evid"]').click());
  step("tab anal", () => doc.querySelector('[data-t="anal"]').click());
  step("tab db", () => doc.querySelector('[data-t="db"]').click());
  step("tab pres", () => doc.querySelector('[data-t="pres"]').click());
  step("deptAdvice", () => { const r = api.cur(); const out = w.deptAdvice(r); if (!Array.isArray(out)) throw new Error("no array"); });
  step("buildDeptReport", () => w.buildDeptReport());
  step("renderRep", () => { w.renderRep ? w.renderRep() : (() => { throw new Error("no renderRep"); })(); });
  step("buildSlides", () => { const r = api.cur(); const s = w.buildSlides(r); if (!s || !s.length) throw new Error("empty slides"); const j = s.map(x => x.h || x).join(""); if (/undefined%|NaN/.test(j)) throw new Error("NaN/undefined in slides"); if (/c\.bl/.test(j)) throw new Error("raw c.bl leaked"); });
  step("exportChecklistMD", () => { const r = api.cur(); const md = w.exportChecklistMD(r); if (!md || md.length < 200) throw new Error("short md"); });
  step("exportActMD", () => { const r = api.cur(); const md = w.exportActMD ? w.exportActMD(r) : (w.exportActHTML ? w.exportActHTML(r) : ""); if (String(md).indexOf("NaN") >= 0) throw new Error("NaN in act"); });
  step("setQA18 route (no respondent)", () => { w.setQA18("А1", "a", "5"); const r = api.cur(); if (r.q.S["А1"].a !== "5") throw new Error("not written"); });
  step("renderCheck after setQA", () => w.renderCheck());
  step("respStats", () => { const st = w.respStats(api.cur()); if (typeof st.cI === "undefined" || !("dis" in st) || !Array.isArray(st.rows)) throw new Error("bad contract"); });
  step("tornado", () => { const t = w.tornado(api.cur()); if (t && t.length !== 3) throw new Error("bad len"); });
  step("mcRun", () => { const arr = w.mcRun(api.cur(), 50, 7); if (!Array.isArray(arr) || arr.length !== 50) throw new Error("bad MC"); });
  step("scenTargets+snapshot", () => { const r = api.cur(); doc.querySelector('[data-t="anal"]').click(); doc.getElementById("scAdd").click(); doc.getElementById("scName").value = "ПРОБА"; const ok = mbtn(/Сохранить/); if (!ok) throw new Error("no save btn"); ok.click(); const s = r.scenarios[r.scenarios.length - 1]; if (!s.qS) throw new Error("no qS snapshot"); const tg = api.scenTargets(s); if (typeof tg !== "object") throw new Error("scenTargets fail"); });
  step("scenario apply", () => { const b = doc.querySelector("#tab-anal [data-scapp]"); if (!b) throw new Error("no apply btn"); b.click(); const ok = mbtn(/Применить/); if (!ok) throw new Error("no apply mbtn"); ok.click(); const r = api.cur(); const tg = api.scenTargets(r.scenarios[r.scenarios.length - 1]); Object.keys(tg).forEach(cd => { if (String(r.q.S[cd].a) !== String(tg[cd])) throw new Error("apply mismatch " + cd); }); w.renderAll(); });
  step("person route enable+seed", () => { const r = api.cur(); doc.querySelector('[data-t="resp"]').click(); doc.getElementById("respOn").click(); const en = mbtn(/Включить/); if (!en) throw new Error("no enable mbtn"); en.click(); if (!r.resp || !r.resp.enabled) throw new Error("resp not enabled"); const p1 = w.eval("actResp")(r); if (!p1 || !p1.qS) throw new Error("active person qS missing"); });
  step("setQA18 person route", () => { const r = api.cur(); const p1 = w.eval("actResp")(r); const before = Object.keys(p1.qS).length; w.setQA18("Б5", "a", "4"); if (String(p1.qS["Б5"].a) !== "4") throw new Error("person qS not updated"); if (String(r.q.S["Б5"].a) !== "4") throw new Error("consensus not synced"); });
  step("qrow ids + 💬/📎 buttons", () => { doc.querySelector('[data-t="check"]').click(); const row = doc.getElementById("qrS-Б5"); if (!row) throw new Error("no qrS-Б5"); if (!row.querySelector("[data-thq18]") || !row.querySelector("[data-doc18]")) throw new Error("no thq/doc buttons"); });
  step("threads+docs modal", () => { const b = doc.querySelector("#tab-check [data-thq18]"); b.click(); const h1 = doc.getElementById("mbox").innerHTML; if (h1.length < 50) throw new Error("threads modal empty"); const x1 = mbtn(/Закрыть|Отмена|ОК/); if (x1) x1.click(); w.renderCheck(); const d = doc.querySelector("#tab-check [data-doc18]"); d.click(); const h2 = doc.getElementById("mbox").innerHTML; if (h2.length < 50) throw new Error("docs modal empty"); const x2 = mbtn(/Закрыть|Отмена|Готово/); if (x2) x2.click(); w.renderCheck(); });
  step("KB sub=anch", () => { const h = w.expKBHTML ? w.expKBHTML() : ""; if (/31 вопросу/.test(h)) throw new Error("old anchor header"); });
  step("TOUR rebuilt", () => { const t = html.slice(html.indexOf("const TOUR=["), html.indexOf("const TOUR=[") + 1800); if (!/Опросник 15\/30\/60/.test(t)) throw new Error("tour stale"); });
  step("disagreements", () => { const d = w.disagreementOf(api.cur(), false); if (!Array.isArray(d)) throw new Error("not array"); });
  step("renderAll final", () => { ["dash", "check", "evid", "anal", "db", "pres", "act", "ref", "wgt", "flow", "resp"].forEach(t => { const b = doc.querySelector('[data-t="' + t + '"]'); if (b) b.click(); }); w.renderAll(); });
  setTimeout(() => {
    step("no console errors accumulated", () => { if (errors.length) throw new Error("errors: " + errors.slice(0, 6).join(" || ")); });
    console.log("=== PROBE18B RESULT:", errors.length ? "ERRORS(" + errors.length + ")" : "CLEAN");
    if (errors.length) errors.slice(0, 12).forEach(e => console.log(" *", e));
    process.exit(errors.length ? 1 : 0);
  }, 400);
}, 700);
