/* ---------- 25. Математический аппарат обоснованности (этап 22, v22.0) ---------- */
console.log('25. Математический аппарат обоснованности (v22.0)');
api.setRole('chair');
api.demo('port');
ck('25: API just/kendallW/ciStat/median/stdevS/printJust доступны', ['just', 'kendallW', 'ciStat', 'median', 'stdevS', 'printJust'].every(k => typeof api[k] === 'function'), null, true);
ck('25: TCRIT95 — 9 значений Стьюдента (df1=12.706 … df9=2.262)', Array.isArray(api.TCRIT95) && api.TCRIT95.length === 9 && api.TCRIT95[0] === 12.706 && api.TCRIT95[8] === 2.262, null, true);
ck('25: CHI2CRIT df 9/11/21 = 16.919/19.675/32.671', api.CHI2CRIT[9] === 16.919 && api.CHI2CRIT[11] === 19.675 && api.CHI2CRIT[21] === 32.671, null, true);
eq('25: median нечётная выборка', api.median([5, 1, 3]), 3);
eq('25: median чётная выборка', api.median([4, 1, 3, 2]), 2.5);
eq('25: median пустая → null', api.median([]), null);
close('25: stdevS классический набор √(32/7)', api.stdevS([2, 4, 4, 4, 5, 5, 7, 9]), Math.sqrt(32 / 7), 1e-4);
eq('25: stdevS n<2 → null', api.stdevS([3]), null);
const m25c = api.ciStat([3, 4, 5]);
close('25: ciStat среднее', m25c.m, 4);
eq('25: ciStat df=n−1', m25c.df, 2);
close('25: ciStat t(0.975;2)=4.303', m25c.t, 4.303);
close('25: ciStat SE=s/√n', m25c.se, 1 / Math.sqrt(3), 1e-4);
close('25: ciStat нижняя граница', m25c.lo, 4 - 4.303 / Math.sqrt(3), 1e-3);
close('25: ciStat верхняя граница', m25c.hi, 4 + 4.303 / Math.sqrt(3), 1e-3);
const m25cc = api.ciStat([4, 4, 4, 4, 4]);
ck('25: ciStat константа → вырожденный ДИ', m25cc.se === 0 && m25cc.lo === 4 && m25cc.hi === 4, null, true);
ck('25: ciStat n=1 → ДИ null', api.ciStat([3]).lo === null && api.ciStat([3]).hi === null, null, true);
const m25cw = api.ciStat([3, 4, 5], [1, 1, 2]);
close('25: ciStat взвешенное среднее', m25cw.m, 4.25);
close('25: ciStat n_eff (Kish) = 16/6', m25cw.neff, 16 / 6, 1e-4);
eq('25: ciStat df по n_eff = 1', m25cw.df, 1);
close('25: ciStat взвешенный SE = √6.6/4', m25cw.se, Math.sqrt(6.6) / 4, 1e-4);
const m25ce = api.ciStat([3, 4, 5], [1, 1, 1]), m25cu = api.ciStat([3, 4, 5]);
ck('25: равные веса = невзвешенный ДИ', Math.abs(m25ce.m - m25cu.m) < 1e-9 && Math.abs(m25ce.se - m25cu.se) < 1e-9 && m25ce.df === m25cu.df, null, true);
const m25kw1 = api.kendallW([[1, 2, 3, 4, 5], [1, 2, 3, 4, 5]]);
close('25: W полное согласие = 1', m25kw1.W, 1);
close('25: χ² = m(k−1)W = 8', m25kw1.chi2, 8);
const m25kw2 = api.kendallW([[1, 2, 3, 4], [1, 2, 3, 4], [4, 3, 2, 1]]);
close('25: W несогласие = 1/9', m25kw2.W, 1 / 9, 1e-4);
close('25: χ² несогласие = 1', m25kw2.chi2, 1, 1e-4);
const m25kw3 = api.kendallW([[1, 1, 2, 2], [1, 1, 2, 2]]);
close('25: W со связками (поправка T) = 1', m25kw3.W, 1);
const m25kw4 = api.kendallW([[3, 3, 3], [3, 3, 3]]);
ck('25: W все оценки совпадают → null + пояснение', m25kw4.W === null && !!m25kw4.note, null, true);
eq('25: W m<2 → null', api.kendallW([[1, 2, 3]]), null);
eq('25: W нечисловые оценки → null', api.kendallW([[1, 2], ['НО', 3]]), null);
const m25kw5 = api.kendallW([[1, 2, 3, 4, 5], [2, 1, 4, 3, 5], [1, 3, 2, 5, 4]]);
close('25: W контрольный пример = 816/1080', m25kw5.W, 816 / 1080, 1e-4);
close('25: χ² контрольный пример = 9.0667', m25kw5.chi2, 3 * 4 * 816 / 1080, 1e-3);
const m25J0 = api.just(0);
eq('25: perQ — 22 вопроса', m25J0.perQ.length, 22);
eq('25: gates — 5 ворот', m25J0.gates.length, 5);
ck('25: gate ids G1..G5', m25J0.gates.map(g => g.id).join(',') === 'G1,G2,G3,G4,G5', null, true);
ck('25: jackknife прогоны = действующие эксперты', m25J0.jack.n === m25J0.R.nExp && m25J0.jack.runs.length === m25J0.jack.n, null, true);
ck('25: W блоки 22/10/12', m25J0.W.all.k === 22 && m25J0.W.A.k === 10 && m25J0.W.C.k === 12, null, true);
ck('25: projects() содержит J', api.projects().every(p => p.J && p.J.gates.length === 5), null, true);
function m25manualCI(vals) { const n = vals.length, m = vals.reduce((a, b) => a + b, 0) / n; const s = Math.sqrt(vals.reduce((a, b) => a + (b - m) * (b - m), 0) / (n - 1)); const t = api.TCRIT95[Math.min(9, n - 1) - 1]; return { m, lo: m - t * s / Math.sqrt(n), hi: m + t * s / Math.sqrt(n) }; }
const m25iaV = m25J0.R.exp.filter(x => x.IA != null).map(x => x.IA);
const m25icV = m25J0.R.exp.filter(x => x.IC != null).map(x => x.IC);
const m25mA = m25manualCI(m25iaV), m25mC = m25manualCI(m25icV);
close('25: ciA = ручной пересчёт (центр)', m25J0.ciA.m, m25mA.m, 1e-6);
close('25: ciA = ручной пересчёт (нижн.)', m25J0.ciA.lo, m25mA.lo, 1e-6);
close('25: ciC = ручной пересчёт (верх.)', m25J0.ciC.hi, m25mC.hi, 1e-6);
const m25q0 = m25J0.R.q[0], m25pq0 = m25J0.perQ[0];
const m25srt = m25q0.vals.slice().sort((a, b) => a - b);
const m25med = m25srt.length % 2 ? m25srt[(m25srt.length - 1) / 2] : (m25srt[m25srt.length / 2 - 1] + m25srt[m25srt.length / 2]) / 2;
eq('25: perQ А1 медиана = ручная', m25pq0.median, m25med);
close('25: perQ А1 ДИ = ручной пересчёт', m25pq0.ci.lo, m25manualCI(m25q0.vals).lo, 1e-6);
const m25recs = api.projects().map(p => p.R.rec);
ck('25: базовые рекомендации демо не изменились', m25recs[0] === 'Продолжить НИОКР' && m25recs[3] === 'Пересмотреть необходимость и назначение работ' && m25recs[4] === 'Рассмотреть прекращение', m25recs, null);
// фикстура: полное согласие рангов → 5/5 ворот
api.reset();
const m25codes = api.QEO22.q.map(q => q.code);
for (let e = 0; e < 3; e++) m25codes.forEach((c, i) => { api.state().exp[e].q[c] = (i % 5) + 1; });
api.switchProject(0);
const m25Jp = api.just();
eq('25: фикстура согласия: okCount=5', m25Jp.okCount, 5);
eq('25: фикстура согласия: вердикт', m25Jp.verdict, 'Решение статистически обосновано');
eq('25: фикстура согласия: short', m25Jp.short, 'обосновано');
close('25: фикстура согласия: W=1', m25Jp.W.all.W, 1);
close('25: фикстура согласия: χ²=63', m25Jp.W.all.chi2, 63);
ck('25: фикстура согласия: χ² значим (63≥32.671)', m25Jp.W.all.sig === true, null, true);
ck('25: фикстура согласия: зоны устойчивы', m25Jp.ciA.stable && m25Jp.ciC.stable, null, true);
eq('25: фикстура согласия: jackknife без изменений', m25Jp.jack.changes, 0);
// фикстура: неустойчивая зона (I_A 3.5/4.0/4.5 → ДИ пересекает границы)
api.reset();
const m25patA = [[3, 3, 3, 3, 3, 4, 4, 4, 4, 4], [4, 4, 4, 4, 4, 4, 4, 4, 4, 4], [4, 4, 4, 4, 4, 5, 5, 5, 5, 5]];
for (let e = 0; e < 3; e++) {
  api.QEO22.q.filter(q => q.dir === 'A').forEach((q, i) => { api.state().exp[e].q[q.code] = m25patA[e][i]; });
  api.QEO22.q.filter(q => q.dir === 'C').forEach(q => { api.state().exp[e].q[q.code] = 4; });
}
api.switchProject(0);
const m25Ju = api.just();
close('25: неустойчивая зона: I_A=4.00', m25Ju.R.IA, 4, 1e-6);
ck('25: неустойчивая зона: ДИ пересекает границы зон', m25Ju.ciA.lo < 3 && m25Ju.ciA.hi > 4 && m25Ju.ciA.stable === false, { lo: m25Ju.ciA.lo, hi: m25Ju.ciA.hi }, null);
ck('25: неустойчивая зона: G3 не выполнено', m25Ju.gates[2].ok === false, null, true);
// фикстура: jackknife меняет рекомендацию (E1 I_A=2, E2/E3 I_A=4.5, I_C=4)
api.reset();
for (let e = 0; e < 3; e++) {
  api.QEO22.q.filter(q => q.dir === 'A').forEach((q, i) => { api.state().exp[e].q[q.code] = e === 0 ? 2 : (i < 5 ? 4 : 5); });
  api.QEO22.q.filter(q => q.dir === 'C').forEach(q => { api.state().exp[e].q[q.code] = 4; });
}
api.switchProject(0);
const m25Jj = api.just();
ck('25: jackknife: есть изменение рекомендации', m25Jj.jack.changes >= 1, m25Jj.jack.changes, '>=1');
ck('25: jackknife: G4 не выполнено', m25Jj.gates[3].ok === false, null, true);
ck('25: jackknife: вердикт не «обосновано»', m25Jj.short !== 'обосновано', m25Jj.short, null);
// фикстура: Н/О → G1 не выполнено, W по complete cases
api.reset();
for (let e = 0; e < 3; e++) m25codes.forEach(c => { api.state().exp[e].q[c] = 4; });
api.state().exp[0].q[m25codes[0]] = 'НО';
api.switchProject(0);
const m25Jn = api.just();
eq('25: Н/О: noTotal=1', m25Jn.noTotal, 1);
ck('25: Н/О: G1 не выполнено', m25Jn.gates[0].ok === false, null, true);
eq('25: Н/О: W complete cases mUsed=2', m25Jn.W.all.mUsed, 2);
// фикстура: разногласие σ>1 → G5 не выполнено
api.reset();
for (let e = 0; e < 3; e++) m25codes.forEach(c => { api.state().exp[e].q[c] = 4; });
api.state().exp[0].q[m25codes[0]] = 1;
api.state().exp[1].q[m25codes[0]] = 5;
api.switchProject(0);
const m25Jd = api.just();
ck('25: разногласие: флаг вопроса', m25Jd.R.flags.indexOf(m25codes[0]) >= 0, m25Jd.R.flags, null);
ck('25: разногласие: G5 не выполнено', m25Jd.gates[4].ok === false, null, true);
// n<3 → вердикт ограничен
api.reset();
m25codes.forEach(c => { api.state().exp[0].q[c] = 4; api.state().exp[1].q[c] = 4; });
api.switchProject(0);
has('25: n<3: вердикт ограничен', api.just().verdict, 'Недостаточно экспертов');
// режимы: веса (n_eff, W не зависит от весов), I_R (ciR)
api.demo('port');
const m25Jw0 = api.just(0);
api.setModes(true, false);
const m25Jw1 = api.just(0);
ck('25: веса: n_eff в ДИ', m25Jw1.ciA.neff != null && m25Jw1.ciA.neff > 0, null, true);
close('25: веса: W не зависит от весов (ранги)', m25Jw1.W.all.W, m25Jw0.W.all.W, 1e-9);
api.setModes(false, true);
ck('25: rmode: ciR рассчитан', api.just(0).ciR != null, null, true);
api.setModes(false, false);
eq('25: rmode выкл: ciR=null', api.just(0).ciR, null);
// UI: реестр, CSV, MD
api.demo('port');
api.goTab('port');
has('25: реестр: колонка «Обосн.»', w.document.querySelector('#tab-port table.ptable').textContent, 'Обосн.');
const m25csv = api.portCSV().replace(/^\uFEFF/, '').split('\n');
eq('25: CSV: 18 колонок', m25csv[0].split(';').length, 18);
has('25: CSV: заголовок «Обоснованность»', m25csv[0], 'Обоснованность');
has('25: CSV: значение обоснованности (n/5)', m25csv[1], '/5)');
has('25: MD портфеля: колонка «Обосн.»', api.portMD(), 'Обосн.');
// печатный протокол (Прил. №13)
ck('25: printJust возвращает true', api.printJust(), null, true);
const m25svod = w.document.getElementById('tab-svod').textContent;
has('25: протокол: ПРИЛОЖЕНИЕ №13', m25svod, 'ПРИЛОЖЕНИЕ №13');
has('25: протокол: статистический вердикт', m25svod, 'Статистический вердикт');
has('25: протокол: подписи комиссии', m25svod, 'Председатель комиссии');
has('25: протокол: ограничения (порядковая шкала)', m25svod, 'порядковая');
w.document.getElementById('bBackJust').click();
has('25: возврат из протокола к своду', w.document.getElementById('tab-svod').textContent, 'Обоснованность решения');
api.goTab('svod');
has('25: свод: карточка обоснованности', w.document.getElementById('tab-svod').textContent, 'Обоснованность решения (v22.0)');
has('25: свод: таблица W Кендалла', w.document.getElementById('tab-svod').textContent, 'коэффициент конкордации Кендалла');
api.goTab('metod');
has('25: методика: раздел 2б', w.document.getElementById('tab-metod').textContent, '2б. Математический аппарат');
// версионные строки и совместимость формата данных
has('25: title v22.0', w.document.title, 'v22.0');
has('25: бейдж v22.0', w.document.querySelector('.ver').textContent, 'v22.0');
has('25: справка: обоснованность (v22.0)', w.document.getElementById('mHelp').textContent, 'Обоснованность решения (v22.0)');
eq('25: формат данных v21 сохранён', api.db().v, 21);
eq('25: экспорт анкеты v21', JSON.parse(api.exportExpert(1)).v, 21);
ck('25: импорт портфеля v21 совместим', api.importPortfolio(api.exportPortfolio()), null, true);
api.demo('port');
