/* Этап 22 (v22.0): QA-зонд математического аппарата обоснованности.
   Контрольные значения вычислены вручную/независимо. Запуск: node scripts/_qa22_probe.js */
const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

const html = fs.readFileSync(path.join(__dirname, '..', 'output', 'versions', 'Аудит_НИОКР_модель_v22.0.html'), 'utf-8');
const errs = [];
const vc = new VirtualConsole();
vc.on('jsdomError', e => errs.push(String((e && e.message) || e)));
vc.on('error', (...a) => errs.push(a.join(' ')));
const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true, virtualConsole: vc, url: 'http://localhost/' });
const w = dom.window, d = w.document;
const api = w.niokrAPI;

let pass = 0, fail = 0;
function ck(name, cond, extra) {
  if (cond) { pass++; console.log('  ✔', name); }
  else { fail++; console.log('  ✘', name, extra !== undefined ? JSON.stringify(extra) : ''); }
}
function near(a, b, tol) { return a != null && Math.abs(a - b) <= (tol || 1e-3); }

setTimeout(function () {
  console.log('=== Зонд v22.0: математический аппарат ===');
  ck('1. версия API 22.0', api.version === '22.0', api.version);
  ck('2. TCRIT95: df2=4.303, df9=2.262', api.TCRIT95[1] === 4.303 && api.TCRIT95[8] === 2.262, api.TCRIT95);
  ck('3. median нечёт/чёт', api.median([5, 1, 3]) === 3 && api.median([1, 2, 3, 4]) === 2.5 && api.median([]) === null);
  // классический набор: популяционная σ=2, выборочная s=sqrt(32/7)=2.13809
  ck('4. stdevS выборочное', near(api.stdevS([2, 4, 4, 4, 5, 5, 7, 9]), Math.sqrt(32 / 7), 1e-4), api.stdevS([2, 4, 4, 4, 5, 5, 7, 9]));
  // [3,4,5]: m=4, s=1, se=1/sqrt(3)=0.57735, df=2, t=4.303 -> lo=1.51614 hi=6.48386
  const c1 = api.ciStat([3, 4, 5]);
  ck('5. ciStat [3,4,5]', near(c1.m, 4) && c1.df === 2 && near(c1.t, 4.303) && near(c1.lo, 4 - 4.303 / Math.sqrt(3), 1e-3) && near(c1.hi, 4 + 4.303 / Math.sqrt(3), 1e-3), c1);
  const c2 = api.ciStat([4, 4, 4, 4, 4]);
  ck('6. ciStat константа: se=0, ДИ вырожден', c2.se === 0 && c2.lo === 4 && c2.hi === 4, c2);
  const c3 = api.ciStat([3]);
  ck('7. ciStat n=1 -> ДИ не рассчитан', c3.lo === null && c3.hi === null && c3.m === 3, c3);
  // взвешенный: vals [3,4,5], w [1,1,2]: sw=4, sw2=6, m=4.25, v=Σw(x-m)²=2.75, den=4-6/4=2.5, s_w²=1.1, se=sqrt(1.1*6)/4=0.64226, neff=16/6, df=1
  const c4 = api.ciStat([3, 4, 5], [1, 1, 2]);
  ck('8. ciStat взвешенный (Kish)', near(c4.m, 4.25) && near(c4.neff, 16 / 6, 1e-4) && c4.df === 1 && near(c4.se, Math.sqrt(6.6) / 4, 1e-4) && near(c4.t, 12.706), c4);
  // W: полное согласие без связок, k=5, m=2 -> W=1
  const w1 = api.kendallW([[1, 2, 3, 4, 5], [1, 2, 3, 4, 5]]);
  ck('9. kendallW полное согласие W=1', w1 && near(w1.W, 1) && near(w1.chi2, 8), w1);
  // W: E1=E2=[1,2,3,4], E3=[4,3,2,1]: Rj=[6,7,8,9], Rbar=7.5, S=5, den=540 -> W=1/9=0.1111, chi2=1
  const w2 = api.kendallW([[1, 2, 3, 4], [1, 2, 3, 4], [4, 3, 2, 1]]);
  ck('10. kendallW несогласие W=1/9', w2 && near(w2.W, 1 / 9, 1e-4) && near(w2.chi2, 1, 1e-4), w2);
  // W: связки, полное согласие [1,1,2,2]x2: T=12 на эксперта, den=192, S=16 -> W=1
  const w3 = api.kendallW([[1, 1, 2, 2], [1, 1, 2, 2]]);
  ck('11. kendallW связки W=1', w3 && near(w3.W, 1), w3);
  const w4 = api.kendallW([[3, 3, 3], [3, 3, 3]]);
  ck('12. kendallW все совпадают -> W=null', w4 && w4.W === null && w4.note, w4);
  ck('13. kendallW m<2 -> null', api.kendallW([[1, 2, 3]]) === null);

  // демо-портфель
  api.demo('port');
  const ps = api.projects();
  ck('14. демо-портфель: 5 НИОКР, J у каждого', ps.length === 5 && ps.every(p => p.J && p.J.gates.length === 5), ps.length);
  const J0 = api.just(0);
  ck('15. just(0): perQ=22, jack.n>0, W.all.mUsed>0', J0.perQ.length === 22 && J0.jack.n > 0 && J0.W.all.mUsed > 0, { n: J0.jack.n, m: J0.W.all.mUsed });
  ck('16. just(0): вердикт и short согласованы', typeof J0.verdict === 'string' && J0.verdict.length > 0 && ['обосновано', 'с оговорками', 'доработка', 'н/д'].indexOf(J0.short) >= 0, J0.verdict);
  ck('17. базовые рекомендации не изменились (v21)', ps[0].R.rec === 'Продолжить НИОКР' && ps[3].R.rec.indexOf('Пересмотреть необходимость') === 0 && ps[4].R.rec === 'Рассмотреть прекращение', ps.map(p => p.R.rec));
  // CI индексов: ручной пересчёт для проекта 0 по индивидуальным IA
  const iaVals = ps[0].R.exp.filter(x => x.IA != null).map(x => x.IA);
  const n = iaVals.length, m = iaVals.reduce((a, b) => a + b, 0) / n;
  const s = Math.sqrt(iaVals.reduce((a, b) => a + (b - m) * (b - m), 0) / (n - 1));
  const t = api.TCRIT95[Math.min(9, n - 1) - 1];
  ck('18. ciA проекта 0 = ручной пересчёт', near(J0.ciA.m, m, 1e-6) && near(J0.ciA.lo, m - t * s / Math.sqrt(n), 1e-6) && near(J0.ciA.hi, m + t * s / Math.sqrt(n), 1e-6), { got: J0.ciA, want: { m, lo: m - t * s / Math.sqrt(n) } });
  ck('19. jackknife: полные прогоны, changed булевы', J0.jack.runs.length === J0.jack.n && J0.jack.runs.every(r => typeof r.changed === 'boolean'), J0.jack.n);

  // CSV / реестр
  const csv = api.portCSV();
  const head = csv.replace(/^\uFEFF/, '').split('\n')[0].split(';');
  ck('20. CSV реестра: 18 колонок, есть «Обоснованность»', head.length === 18 && head[17] === 'Обоснованность', head);
  api.goTab('port');
  ck('21. реестр: колонка «Обосн.» и пилюли', d.querySelector('#tab-port .ptable').textContent.indexOf('Обосн.') >= 0 && d.querySelectorAll('#tab-port .ptable .pill').length > 0);

  // печатный протокол
  ck('22. printJust -> Протокол (Прил. №13)', api.printJust() === true && d.getElementById('tab-svod').textContent.indexOf('ПРИЛОЖЕНИЕ №13') >= 0);
  const back = d.getElementById('bBackJust');
  ck('23. возврат из протокола к своду', !!back);
  if (back) { back.click(); ck('23b. свод восстановлен', d.getElementById('tab-svod').textContent.indexOf('Обоснованность решения') >= 0); }

  // совместимость формата данных
  ck('24. формат данных v21 не изменился', api.db().v === 21 && JSON.parse(api.exportExpert(1)).v === 21);
  ck('25. экспорт/импорт портфеля v21 совместим', api.importPortfolio(api.exportPortfolio()) === true);
  // режим весов: n_eff в ДИ
  api.setModes(true, false);
  const Jw = api.just(0);
  ck('26. режим весов: ciA.neff рассчитан', Jw.ciA && Jw.ciA.neff != null && Jw.ciA.neff > 0, Jw.ciA && Jw.ciA.neff);
  api.setModes(false, false);
  ck('27. ошибки окна отсутствуют', errs.length === 0, errs);

  console.log('=== ИТОГ: ' + pass + ' пройдено, ' + fail + ' ошибок; window errors: ' + errs.length + ' ===');
  process.exit(fail ? 1 : 0);
}, 300);