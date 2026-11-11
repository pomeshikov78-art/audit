// Этап 19 (v19.0): регрессионный комплект тестов модели ЭО-НИОКР (jsdom).
// Якоря: все «5» → В/В «Продолжить НИОКР»; все «1» → Н/Н «Рассмотреть прекращение»;
// оверрайд стоп-фактора; Н/О > 2 → индекс не рассчитан; σ 5/1/3 → существенные разногласия;
// групповые «красные» сигналы; демо С-1/С-2/С-3 → ожидаемые ячейки матрицы; MD-чек-лист 22 строки.
const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

const HTML = path.join(__dirname, '..', '..', 'output', 'versions', 'Аудит_НИОКР_модель_v19.0.html');
const src = fs.readFileSync(HTML, 'utf8');

let PASS = 0, FAIL = 0; const FAILED = [];
function ck(name, cond, got, want) {
  if (cond) { PASS++; console.log('  ✔', name); }
  else { FAIL++; FAILED.push(name); console.log('  ✘', name, '| got=', got, 'want=', want); }
}
function eq(n, g, w) { ck(n, g === w, g, w); }
function close(n, g, w, tol) { tol = tol || 0.005; ck(n, typeof g == 'number' && Math.abs(g - w) <= tol, g, w); }
function has(n, g, s) { ck(n, typeof g == 'string' && g.indexOf(s) >= 0, g, 'contains ' + s); }

const vc = new VirtualConsole();
let winErrors = [];
vc.on('jsdomError', e => winErrors.push(String(e && e.message || e)));
vc.on('error', (...a) => winErrors.push(a.join(' ')));

const dom = new JSDOM(src, { runScripts: 'dangerously', pretendToBeVisual: true, virtualConsole: vc, url: 'http://localhost/' });
const w = dom.window;
const api = w.niokrAPI;

console.log('1. Загрузка и структура данных');
ck('модель загрузилась без ошибок окна', winErrors.length === 0, winErrors.slice(0, 3), []);
ck('niokrAPI доступен', !!api, typeof api, 'object');
eq('версия 19.0', api && api.version, '19.0');
eq('22 вопроса', api.QEO22.q.length, 22);
eq('10 по актуальности', api.QEO22.q.filter(q => q.dir == 'A').length, 10);
eq('12 по целесообразности', api.QEO22.q.filter(q => q.dir == 'C').length, 12);
eq('5 стоп-факторов', api.QEO22.sf.length, 5);
eq('9 красных вопросов', api.QEO22.kv.length, 9);
ck('у каждого вопроса 5 якорей', api.QEO22.q.every(q => [1, 2, 3, 4, 5].every(k => (q.a[k] || '').length > 10)), null, 'все якоря');
ck('направления КВ: 7 негативных + 2 позитивных', api.QEO22.kv.filter(k => k.dir == 'neg').length === 7 && api.QEO22.kv.filter(k => k.dir == 'pos').length === 2, null, '7/2');
eq('6 вкладок', w.document.querySelectorAll('#tabs button').length, 6);

const QA = api.QEO22.q.filter(q => q.dir == 'A').map(q => q.code);
const QC = api.QEO22.q.filter(q => q.dir == 'C').map(q => q.code);
function fillAll(e, vA, vC) { QA.forEach(c => api.setScore(e, c, vA)); QC.forEach(c => api.setScore(e, c, vC)); }

console.log('2. Якорь: все «5» у трёх экспертов → В/В, «Продолжить НИОКР»');
api.reset();
[1, 2, 3].forEach(e => fillAll(e, 5, 5));
let R = api.calc();
close('I_A = 5', R.IA, 5); close('I_C = 5', R.IC, 5);
eq('зона А = В', R.zA, 'В'); eq('зона Ц = В', R.zC, 'В');
eq('рекомендация: Продолжить НИОКР', R.rec, 'Продолжить НИОКР');
eq('стоп-факторов нет', R.sfDet, false);
eq('σ = 0', R.sA, 0);
ck('дашборд: активная ячейка В/В', !!w.document.querySelector('#tab-dash .mtx td.act'), null, 'есть');

console.log('3. Якорь: все «1» → Н/Н, «Рассмотреть прекращение»');
api.reset();
[1, 2, 3].forEach(e => fillAll(e, 1, 1));
R = api.calc();
close('I_A = 1', R.IA, 1); close('I_C = 1', R.IC, 1);
eq('рекомендация: Рассмотреть прекращение', R.rec, 'Рассмотреть прекращение');
eq('низкие оценки: 22 кода', R.lows.length, 22);

console.log('4. Оверрайд стоп-фактора');
api.reset();
[1, 2, 3].forEach(e => fillAll(e, 5, 5));
api.setSF(1, 1, 'Да');
R = api.calc();
eq('стоп-фактор выявлен', R.sfDet, true);
eq('матрица без оверрайда', R.rec0, 'Продолжить НИОКР');
has('оверрайд понизил до Пересмотреть', R.rec, 'ПОНИЖЕНИЕ: Пересмотреть параметры НИОКР');
api.setSF(2, 1, 'Неопределённо');
R = api.calc();
eq('неопределённый учтён', R.sf.filter(s => s.n == 1)[0].unc, 1);

console.log('5. Н/О: >2 по направлению → индекс эксперта не рассчитан');
api.reset();
QA.forEach((c, i) => api.setScore(1, c, i < 3 ? 'НО' : 4));
QC.forEach(c => api.setScore(1, c, 4));
[2, 3].forEach(e => fillAll(e, 4, 4));
R = api.calc();
eq('I_A эксперта 1 = null', R.exp[0].IA, null);
close('I_C эксперта 1 = 4', R.exp[0].IC, 4);
close('групповой I_A = 4 (без Э1)', R.IA, 4);
eq('счётчик Н/О = 3', R.exp[0].noA, 3);
// ровно 2 Н/О — индекс считается
api.reset();
QA.forEach((c, i) => api.setScore(1, c, i < 2 ? 'НО' : 4));
QC.forEach(c => api.setScore(1, c, 4));
R = api.calc();
close('2 Н/О → I_A считается = 4', R.exp[0].IA, 4);

console.log('6. σ-согласованность: 5/1/3 → существенные разногласия, 10 флагов');
api.reset();
QA.forEach(c => api.setScore(1, c, 5));
QA.forEach(c => api.setScore(2, c, 1));
QA.forEach(c => api.setScore(3, c, 3));
[1, 2, 3].forEach(e => QC.forEach(c => api.setScore(e, c, 4)));
R = api.calc();
close('I_A = 3', R.IA, 3);
close('σ(I_A) = 1.633', R.sA, 1.632993, 0.001);
eq('интерпретация: существенные', R.cA, 'Существенные разногласия');
eq('флагов разногласия = 10', R.flags.length, 10);
close('σ(I_C) = 0', R.sC, 0);

console.log('7. «Красные» вопросы: групповой и единичный сигналы, направление');
api.reset();
[1, 2, 3].forEach(e => fillAll(e, 4, 4));
api.setKV(1, 1, 1); api.setKV(2, 1, 2); api.setKV(3, 1, 5);   // КВ1 негативный: 2 сигнала
api.setKV(1, 7, 5); api.setKV(2, 7, 4); api.setKV(3, 7, 1);   // КВ7 позитивный: 2 сигнала
api.setKV(1, 3, 1);                                           // КВ3: единичный
R = api.calc();
ck('КВ1 групповой', R.kv.filter(k => k.n == 1)[0].group, true, true);
ck('КВ7 групповой (позитивный, 4–5)', R.kv.filter(k => k.n == 7)[0].group, true, true);
ck('КВ3 единичный', R.kv.filter(k => k.n == 3)[0].single, true, true);
eq('групповых сигналов = 2', R.kvGroup.length, 2);

console.log('8. Демо-наборы апробации (С-1/С-2/С-3)');
api.demo('s1'); R = api.calc();
close('С-1: I_A = 4.5', R.IA, 4.5); close('С-1: I_C = 2.5', R.IC, 2.5);
eq('С-1: матрица В/Н', R.rec, 'Пересмотреть содержание и способ реализации');
ck('С-1: красный КВ5 групповой', R.kv.filter(k => k.n == 5)[0].group, true, true);
api.demo('s2'); R = api.calc();
close('С-2: I_A = 2.5', R.IA, 2.5); close('С-2: I_C = 4.5', R.IC, 4.5);
eq('С-2: матрица Н/В', R.rec, 'Пересмотреть необходимость и назначение работ');
eq('С-2: стоп-фактор СФ1 выявлен', R.sfDet, true);
api.demo('s3'); R = api.calc();
close('С-3: I_A = 3.9', R.IA, 3.9); close('С-3: I_C = 3.5', R.IC, 3.5);
eq('С-3: матрица С/С', R.rec, 'Пересмотреть параметры НИОКР');
ck('С-3: КВ7 и КВ8 групповые', R.kv.filter(k => k.n == 7)[0].group && R.kv.filter(k => k.n == 8)[0].group, true, true);

console.log('9. Документы: MD-чек-лист и протокол');
const md = api.exportMD();
ck('MD: 22 строки вопросов', api.QEO22.q.every(q => md.indexOf('| ' + q.code + ' |') >= 0), null, 'все коды');
has('MD: раздел индексов', md, '## Индексы');
has('MD: стоп-факторы', md, '## Стоп-факторы');
has('MD: красные вопросы', md, '## «Красные» вопросы');
const pt = api.protText();
has('протокол: природа оценки', pt, 'ПРИРОДА ОЦЕНКИ');
has('протокол: решение комиссии', pt, 'РЕШЕНИЕ КОМИССИИ');

console.log('10. UI: вкладки, анкеты, свод, заключение');
api.demo('s1');
api.goTab('ank');
ck('анкета: 22 строки вопросов', w.document.querySelectorAll('#tab-ank [data-score="1"]').length === 22, w.document.querySelectorAll('#tab-ank [data-score="1"]').length, 22);
ck('анкета: 5 строк СФ', w.document.querySelectorAll('#tab-ank [data-sf]').length === 15, w.document.querySelectorAll('#tab-ank [data-sf]').length, 15);
ck('анкета: 45 кнопок КВ', w.document.querySelectorAll('#tab-ank [data-kv]').length === 45, w.document.querySelectorAll('#tab-ank [data-kv]').length, 45);
api.goTab('svod');
ck('свод: таблица 22 вопросов отрисована', w.document.querySelectorAll('#tab-svod table')[1].querySelectorAll('tr').length === 23, w.document.querySelectorAll('#tab-svod table')[1].querySelectorAll('tr').length, 23);
api.goTab('zakl');
has('заключение: текст раздела 3', w.document.querySelector('#tab-zakl').textContent, 'Индекс актуальности составляет 4.50');
api.goTab('metod');
has('методика: якоря А1', w.document.querySelector('#tab-metod').textContent, 'потребность является очевидной');
api.goTab('panel');
ck('панель: SVG профиля', !!w.document.querySelector('#tab-panel svg'), null, 'есть');
ck('нет ошибок окна после всех рендеров', winErrors.length === 0, winErrors.slice(0, 3), []);

console.log();
console.log('ИТОГ: ' + (PASS + FAIL) + ' проверок, ' + FAIL + ' ошибок');
if (FAILED.length) { console.log('ПРОВАЛЕНО:', FAILED); process.exit(1); }
console.log('ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ');
