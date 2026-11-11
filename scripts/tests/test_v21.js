// Этап 21 (v21.0): регрессионный комплект тестов модели ЭО-НИОКР v2.0 (jsdom).
// База этапа 19 + группа до 10 экспертов (демо x10: I_A=2.99, Н/В), роли (эксперт/секретарь/председатель),
// раунды (фиксация раунда 1, правки = раунд 2), поанкетный экспорт/импорт, повестка обсуждения,
// предупреждение при составе <3, MD/CSV с 10 колонками экспертов.
const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

const HTML = path.join(__dirname, '..', '..', 'output', 'versions', 'Аудит_НИОКР_модель_v21.0.html');
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
eq('версия 21.0', api && api.version, '21.0');
eq('22 вопроса', api.QEO22.q.length, 22);
eq('10 по актуальности', api.QEO22.q.filter(q => q.dir == 'A').length, 10);
eq('12 по целесообразности', api.QEO22.q.filter(q => q.dir == 'C').length, 12);
eq('5 стоп-факторов', api.QEO22.sf.length, 5);
eq('9 красных вопросов', api.QEO22.kv.length, 9);
ck('у каждого вопроса 5 якорей', api.QEO22.q.every(q => [1, 2, 3, 4, 5].every(k => (q.a[k] || '').length > 10)), null, 'все якоря');
ck('направления КВ: 7 негативных + 2 позитивных', api.QEO22.kv.filter(k => k.dir == 'neg').length === 7 && api.QEO22.kv.filter(k => k.dir == 'pos').length === 2, null, '7/2');
eq('8 вкладок', w.document.querySelectorAll('#tabs button').length, 8);
eq('первый запуск: демо-портфель из 5 НИОКР', api.db().projects.length, 5);
api.clearPortfolio();
eq('после очистки: 1 пустая НИОКР', api.db().projects.length, 1);

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


console.log('11. Группа до 10 экспертов: демо x10');
api.demo('x10'); R = api.calc();
eq('x10: заполнено 10 экспертов', R.nExp, 10);
eq('x10: действующих I_A — 10', R.nIdxA, 10);
close('x10: I_A = 2.99', R.IA, 2.99);
close('x10: I_C = 4.00', R.IC, 4);
eq('x10: зона А = Н', R.zA, 'Н');
eq('x10: зона Ц = В', R.zC, 'В');
eq('x10: матрица Н/В', R.rec, 'Пересмотреть необходимость и назначение работ');
close('x10: σ(I_A) ≈ 0.08', R.sA, 0.0831, 0.001);
eq('x10: согласованность I_A высокая', R.cA, 'Высокая');
ck('x10: нет вопросов с σ>1', R.flags.length === 0, R.flags, []);
ck('x10: групповой сигнал только КВ3', R.kvGroup.length === 1 && R.kvGroup[0] === 3, R.kvGroup, [3]);
ck('x10: повестка содержит КВ3', R.agenda.some(a => a.indexOf('КВ3: групповой сигнал (2 из 10') === 0), R.agenda.slice(0, 2), 'КВ3 agenda');
api.goTab('svod');
ck('x10: свод — ≥10 колонок экспертов', w.document.querySelectorAll('#tab-svod table')[1].querySelectorAll('tr th').length >= 10, w.document.querySelectorAll('#tab-svod table')[1].querySelectorAll('tr th').length, '>=10');

console.log('12. Роли: эксперт / секретарь / председатель');
api.demo('s1');
api.setRole('exp'); api.setLogin(2);
ck('роль эксперта: свод недоступен', w.document.querySelector('#tabs button[data-t="svod"]').disabled, null, true);
ck('роль эксперта: методика доступна', !w.document.querySelector('#tabs button[data-t="metod"]').disabled, null, true);
api.goTab('ank');
has('роль эксперта: предупреждение о независимости', w.document.querySelector('#tab-ank').textContent, 'Режим эксперта');
ck('роль эксперта: чипсы скрыты', !w.document.querySelector('#tab-ank .chips'), null, null);
eq('роль эксперта: логин Э3', api.state().expLogin, 2);
api.setRole('sec');
ck('секретарь: вкладки доступны', !w.document.querySelector('#tabs button[data-t="svod"]').disabled, null, true);
api.goTab('zakl');
ck('секретарь: решение комиссии заблокировано', w.document.querySelector('#cRec').disabled, null, true);
has('секретарь: пояснение', w.document.querySelector('#tab-zakl').textContent, 'Роль «Секретарь»');
api.setRole('chair');
api.goTab('zakl');
ck('председатель: решение комиссии доступно', !w.document.querySelector('#cRec').disabled, null, true);

console.log('13. Раунды: фиксация раунда 1 и раунд 2 после обсуждения');
api.reset();
[1, 2, 3].forEach(e => fillAll(e, 5, 5));
ck('фиксация раунда 1 у Э1', api.fixRound(1), null, true);
R = api.calc();
eq('раунды: один эксперт с фиксацией', R.rounds.length, 1);
close('раунды: I_A раунд 1 = 5', R.rounds[0].IA0, 5);
api.setScore(1, QA[0], 3);
R = api.calc();
close('раунды: I_A Э1 раунд 2 = 4.8', R.exp[0].IA, 4.8);
close('раунды: I_A раунд 1 сохранён', R.rounds[0].IA0, 5);
close('раунды: групповой I_A пересчитан', R.IA, (4.8 + 5 + 5) / 3);
has('протокол: повторная оценка отражена', api.protText(), 'Повторная оценка после обсуждения');
api.goTab('ank');
ck('анкеты: бейдж R2 у Э1', !!w.document.querySelector('#tab-ank .chips button .pill'), null, true);

console.log('14. Поанкетный экспорт/импорт (распределённое слепое заполнение)');
api.reset(); api.demo('s1');
const ej = api.exportExpert(2);
const ejd = JSON.parse(ej);
eq('экспорт анкеты: версия 21', ejd.v, 21);
eq('экспорт анкеты: kind', ejd.kind, 'expert');
eq('экспорт анкеты: слот', ejd.slot, 2);
ck('импорт анкеты в слот 8', api.importExpert(ej, 8), null, true);
R = api.calc();
eq('после импорта: 4 заполненные анкеты', R.nExp, 4);
ck('слот 8 = копия Э2', api.state().exp[7].fio === api.state().exp[1].fio, api.state().exp[7].fio, api.state().exp[1].fio);
ck('импорт мусора отклоняется', api.importExpert('{"v":20}', 9) === false, null, false);

console.log('15. Состав группы и документы на 10 экспертов');
api.reset(); fillAll(1, 5, 5);
api.goTab('dash');
has('предупреждение: меньше 3 экспертов', w.document.querySelector('#tab-dash').textContent, 'рекомендует состав от 3');
api.demo('x10');
const md10 = api.exportMD();
has('MD: шапка с Э10', md10, 'Э10');
has('MD: список экспертов до Э10', md10, 'Э10 Эксперт 10 (демо)');
ck('MD: 22 строки вопросов (10 экспертов)', api.QEO22.q.every(q => md10.indexOf('| ' + q.code + ' |') >= 0), null, 'все коды');
has('протокол x10: согласованность', api.protText(), 'σ(I_A)=0.08 (Высокая)');
ck('протокол x10: нет пустого «КВ.»', api.protText().indexOf('(особые мнения): КВ.') < 0, null, 'нет пустого списка');
ck('нет ошибок окна после всех рендеров (v20)', winErrors.length === 0, winErrors.slice(0, 3), []);


console.log('16. Портфель: демо 5 НИОКР, сравнительная матрица');
api.demo('port');
let PR = api.projects();
eq('портфель: 5 НИОКР', PR.length, 5);
close('БПЛА-СУ-24: I_A = 4.5', PR[0].R.IA, 4.5);
close('БПЛА-СУ-24: I_C = 4.25', PR[0].R.IC, 4.25);
eq('БПЛА-СУ-24: В/В Продолжить', PR[0].R.rec, 'Продолжить НИОКР');
close('МАТ-ЛИТ-25: I_A = 4.325', PR[1].R.IA, 4.325);
close('МАТ-ЛИТ-25: I_C = 3.2708', PR[1].R.IC, 3.2708);
eq('МАТ-ЛИТ-25: В/С корректировка', PR[1].R.rec, 'Продолжить с корректировкой');
close('ЦД-ГТД-23: I_A = 3.3', PR[2].R.IA, 3.3);
close('ЦД-ГТД-23: I_C = 3.3333', PR[2].R.IC, 3.3333);
eq('ЦД-ГТД-23: С/С пересмотр параметров', PR[2].R.rec, 'Пересмотреть параметры НИОКР');
ck('ЦД-ГТД-23: флаг разногласия А3', PR[2].R.flags.indexOf('А3') >= 0, PR[2].R.flags, ['А3']);
close('СВЗ-19: I_A = 2.2', PR[3].R.IA, 2.2);
close('СВЗ-19: I_C = 4.25', PR[3].R.IC, 4.25);
eq('СВЗ-19: стоп-фактор выявлен', PR[3].R.sfDet, true);
eq('СВЗ-19: Н/В пересмотр необходимости', PR[3].R.rec, 'Пересмотреть необходимость и назначение работ');
close('СЕНС-АГРО-26: I_A = 1.6333', PR[4].R.IA, 1.6333);
close('СЕНС-АГРО-26: I_C = 1.6667', PR[4].R.IC, 1.6667);
eq('СЕНС-АГРО-26: Н/Н прекращение', PR[4].R.rec, 'Рассмотреть прекращение');
ck('СЕНС-АГРО-26: КВ3 и КВ5 групповые', PR[4].R.kvGroup.join(',') === '3,5', PR[4].R.kvGroup, [3, 5]);
api.goTab('port');
ck('портфель: SVG-матрица отрисована', !!w.document.querySelector('#tab-port svg.scatter'), null, 'есть');
eq('портфель: 5 точек на матрице', w.document.querySelectorAll('#tab-port svg .pt').length, 5);
ck('портфель: таблица реестра 5 строк', w.document.querySelectorAll('#tab-port table.ptable tr').length === 6, w.document.querySelectorAll('#tab-port table.ptable tr').length, 6);
has('портфель: сводка квадрантов', w.document.querySelector('#tab-port').textContent, 'Сводка портфеля');

console.log('17. Портфель: операции (добавление, копия, удаление, возврат, порядок, лимит)');
ck('добавление НИОКР', api.addProject('Тестовая НИОКР'), null, true);
eq('после добавления: 6', api.db().projects.length, 6);
eq('новая НИОКР активна', api.db().cur, 5);
api.switchProject(0);
ck('дублирование', api.duplicateProject(0), null, true);
eq('после копии: 7', api.db().projects.length, 7);
has('копия: имя с «(копия)»', api.db().projects[1].obj.name, '(копия)');
ck('удаление копии', api.removeProject(1), null, true);
eq('после удаления: 6', api.db().projects.length, 6);
ck('возврат удалённой', api.undoTrash(), null, true);
eq('после возврата: 7', api.db().projects.length, 7);
const name0 = api.db().projects[0].obj.name;
ck('перемещение 0 → 2', api.reorderProject(0, 2), null, true);
eq('порядок изменён', api.db().projects[2].obj.name, name0);
let capOk = true;
while (api.db().projects.length < 30) { if (!api.addProject('Лимит ' + api.db().projects.length)) { capOk = false; break; } }
eq('портфель заполнен до 30', api.db().projects.length, 30);
ck('31-я НИОКР отклонена', api.addProject('Лишняя') === false, null, false);
api.clearPortfolio();

console.log('18. Веса экспертов (v2.0)');
api.reset();
fillAll(1, 5, 5); fillAll(2, 3, 3);
api.setWeight(1, 2);
api.setModes(true, false);
R = api.calc();
close('взвешенный I_A = 4.3333', R.IA, 13 / 3);
close('взвешенный σ(I_A) = 0.9428', R.sA, 0.942809);
eq('взвешенная согласованность: умеренные', R.cA, 'Умеренные разногласия');
api.setModes(false, false);
R = api.calc();
close('равные веса: I_A = 4', R.IA, 4);
close('равные веса: σ = 1', R.sA, 1);
api.goTab('svod');
has('свод: колонка w при включённых весах', (api.setModes(true, false), w.document.querySelector('#tab-svod').textContent), 'взвешенное среднее');
api.setModes(false, false);

console.log('19. Блок результативности I_R (v2.0)');
api.reset();
[1, 2, 3].forEach(e => fillAll(e, 5, 5));
api.setModes(false, true);
const RC = api.QEO21.rblock.q.map(q => q.code);
eq('блок Р: 6 вопросов', RC.length, 6);
[1, 2, 3].forEach(e => RC.forEach(c => api.setR(e, c, 4)));
R = api.calc();
close('I_R = 4', R.IR, 4);
eq('зона I_R = В', R.zR, 'В');
api.setR(1, RC[0], 'НО');
R = api.calc();
close('I_R Э1 с 1 Н/О = 4', R.exp[0].IR, 4);
api.setR(2, RC[0], 'НО'); api.setR(2, RC[1], 'НО');
R = api.calc();
ck('I_R Э2 с 2 Н/О не рассчитан', R.exp[1].IR === null, R.exp[1].IR, null);
close('групповой I_R = 4 (без Э2)', R.IR, 4);
eq('модификатор не применяется при I_R=В', R.rec, 'Продолжить НИОКР');
[1, 2, 3].forEach(e => RC.forEach(c => api.setR(e, c, 2)));
R = api.calc();
close('I_R = 2', R.IR, 2);
eq('модификатор: план повышения результативности', R.rec, 'Продолжить НИОКР с планом повышения результативности (I_R ниже зоны С)');
ck('флаг irMod', R.irMod, null, true);
api.setSF(1, 1, 'Да');
R = api.calc();
has('стоп-фактор важнее модификатора', R.rec, 'ПОНИЖЕНИЕ');
api.setSF(1, 1, 'Нет');
has('заключение: раздел 4а', api.protText(), '4А. ОЦЕНКА РЕЗУЛЬТАТИВНОСТИ');
api.goTab('ank');
has('анкета: блок Р отрисован', w.document.querySelector('#tab-ank').textContent, 'Блок результативности Р1–Р6');
api.setModes(false, false);

console.log('20. Калибровка по апробации');
api.demo('port');
const st = api.calibStats();
eq('статистика: 5 НИОКР', st.nProj, 5);
eq('статистика: 16 анкет', st.nForms, 16);
eq('статистика: 352 числовые оценки', st.total, 352);
ck('статистика: сумма распределения = 352', [1, 2, 3, 4, 5].reduce((a, k) => a + st.counts[k], 0) === 352, null, 352);
const prop = api.proposeCalib();
ck('предложение порогов по данным', prop && prop.zoneV > prop.zoneS && prop.s1 < prop.s2 && prop.s2 < prop.s3, prop, 'корректные пороги');
ck('калибровка применена', api.setCalib({ zoneV: 4.5, zoneS: 3.2, s1: 0.4, s2: 0.7, s3: 0.9 }, { why: 'тест', prot: 'К-1' }), null, true);
PR = api.projects();
eq('БПЛА при В≥4.5: зона А = В', PR[0].R.zA, 'В');
eq('БПЛА при В≥4.5: зона Ц = С', PR[0].R.zC, 'С');
eq('БПЛА: рекомендация В/С', PR[0].R.rec, 'Продолжить с корректировкой');
eq('история калибровок: 1 запись', api.db().calibHist.length, 1);
ck('некорректные пороги отклонены', api.setCalib({ zoneV: 2, zoneS: 3, s1: 0.4, s2: 0.7, s3: 0.9 }, {}) === false, null, false);
ck('сброс калибровки', api.resetCalib(), null, true);
PR = api.projects();
eq('после сброса: БПЛА В/В', PR[0].R.rec, 'Продолжить НИОКР');
eq('история: 2 записи', api.db().calibHist.length, 2);
api.goTab('calib');
has('калибровка: гистограмма', w.document.querySelector('#tab-calib').textContent, 'распределение баллов');
api.goTab('zakl');
has('заключение: отметка о калибровке', (api.setCalib({ zoneV: 4.5, zoneS: 3.2, s1: 0.4, s2: 0.7, s3: 0.9 }, { why: 'тест2', prot: 'К-2' }), api.protText()), 'Пороги калиброваны');
api.resetCalib();

console.log('21. Печатный бланк анкеты (Приложение №1)');
api.demo('s1');
api.setModes(false, true);
ck('бланк открыт', api.printForm(), null, true);
const pf = w.document.querySelector('#tab-ank .pform');
ck('бланк: форма отрисована', !!pf, null, 'есть');
has('бланк: заголовок приложения', pf.textContent, 'ПРИЛОЖЕНИЕ №1');
eq('бланк: 22 строки вопросов + заголовок + 2 секции', pf.querySelectorAll('table.frm')[0].querySelectorAll('tr').length, 25);
has('бланк: стоп-факторы', pf.textContent, 'СТОП-ФАКТОРЫ');
has('бланк: красные вопросы', pf.textContent, '«КРАСНЫЕ» ВОПРОСЫ');
has('бланк: блок Р при включённом режиме', pf.textContent, 'БЛОК РЕЗУЛЬТАТИВНОСТИ');
has('бланк: подпись эксперта', pf.textContent, 'Подпись эксперта');
ck('бланк: отмечен текущий ответ А1=5', !!pf.querySelector('td.mark'), null, 'есть');
api.setModes(false, false);
api.goTab('ank');

console.log('22. Представления, отчёты, совместимость');
api.demo('port');
ck('представление сохранено', api.saveView('Тестовое'), null, true);
eq('представлений: 1', api.db().views.length, 1);
ck('представление применено', api.applyView('Тестовое'), null, true);
const csvp = api.portCSV();
has('CSV портфеля: заголовок I_A', csvp, 'I_A');
ck('CSV портфеля: 5 строк НИОКР', csvp.split('\n').length === 6, csvp.split('\n').length, 6);
const mdp = api.portMD();
has('MD портфеля: заголовок', mdp, 'Отчёт по портфелю НИОКР');
has('MD портфеля: шифр БПЛА', mdp, 'БПЛА-СУ-24');
const ej21 = JSON.parse(api.exportExpert(1));
eq('экспорт анкеты: v21', ej21.v, 21);
ej21.v = 20;
ck('импорт анкеты v20 совместим', api.importExpert(JSON.stringify(ej21), 4), null, true);
const pj = api.exportProject(0);
ck('экспорт НИОКР: kind project', JSON.parse(pj).kind === 'project', null, true);
const full = api.exportPortfolio();
ck('экспорт портфеля: v21', JSON.parse(full).v === 21, null, true);
ck('импорт портфеля', api.importPortfolio(full), null, true);
eq('после импорта: 5 НИОКР', api.db().projects.length, 5);
/* ---------- 23. Регрессия: демо не сбрасывает роль ---------- */
api.setRole('exp'); api.setLogin(2);
api.demo('s1');
eq('демо s1 в режиме эксперта: роль сохранена', api.state().role, 'exp');
eq('демо s1 в режиме эксперта: логин сохранён', api.state().expLogin, 2);
ck('демо s1 в режиме эксперта: свод заблокирован', w.document.querySelector('#tabs button[data-t="svod"]').disabled, null, true);
api.demo('x10');
eq('демо x10 в режиме эксперта: роль сохранена', api.state().role, 'exp');
api.reset();
eq('reset в режиме эксперта: роль сохранена', api.state().role, 'exp');
api.setRole('chair');
api.demo('port');
eq('возврат к председателю: роль chair', api.state().role, 'chair');

ck('нет ошибок окна после всех рендеров (v21)', winErrors.length === 0, winErrors.slice(0, 3), []);

console.log();
console.log('ИТОГ: ' + (PASS + FAIL) + ' проверок, ' + FAIL + ' ошибок');
if (FAILED.length) { console.log('ПРОВАЛЕНО:', FAILED); process.exit(1); }
console.log('ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ');
