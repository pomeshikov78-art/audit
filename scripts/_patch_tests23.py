# -*- coding: utf-8 -*-
import io
p='scripts/tests/test_v22.js';s=io.open(p,encoding='utf-8').read()
def rep(a,b,n=1):
 global s
 c=s.count(a);assert c==n,(c,n,a[:80]);s=s.replace(a,b)
rep("Аудит_НИОКР_модель_v22.0.html","Аудит_НИОКР_модель_v23.0.html")
rep("Этап 22 (v22.0)","Этап 23 (v23.0)")
rep("eq('версия 22.0', api && api.version, '22.0')","eq('версия 23.0', api && api.version, '23.0')")
rep("ck('направления КВ: 7 негативных + 2 позитивных', api.QEO22.kv.filter(k => k.dir == 'neg').length === 7 && api.QEO22.kv.filter(k => k.dir == 'pos').length === 2, null, '7/2')","ck('КВ: единая шкала risk_yes у всех 9', api.QEO22.kv.filter(k => k.dir === 'risk_yes').length === 9, null, '9')")
rep("ck('анкета: 5 строк СФ', w.document.querySelectorAll('#tab-ank [data-sf]').length === 15, w.document.querySelectorAll('#tab-ank [data-sf]').length, 15)","ck('анкета: 5 СФ × 5 числовых кнопок', w.document.querySelectorAll('#tab-ank [data-sf]').length === 25, w.document.querySelectorAll('#tab-ank [data-sf]').length, 25)")
rep("api.setKV(1, 7, 5); api.setKV(2, 7, 4); api.setKV(3, 7, 1);   // КВ7 позитивный: 2 сигнала","api.setKV(1, 7, 1); api.setKV(2, 7, 2); api.setKV(3, 7, 5);   // КВ7: единая шкала, 1–2 = риск")
rep("КВ7 групповой (позитивный, 4–5)","КВ7 групповой (единая шкала, 1–2)")
rep("С-3: КВ7 и КВ8 групповые", "С-3: КВ7 и КВ8 без группового сигнала после исправления полярности")
rep("R.kv.filter(k => k.n == 7)[0].group && R.kv.filter(k => k.n == 8)[0].group, true, true", "R.kv.filter(k => k.n == 7)[0].group || R.kv.filter(k => k.n == 8)[0].group, false, false")
rep("Обоснованность решения (v22.0)","Обоснованность решения (v23.0)",2)
rep("has('25: title v22.0'","has('25: title v23.0'");rep("w.document.title, 'v22.0'","w.document.title, 'v23.0'")
rep("has('25: бейдж v22.0'","has('25: бейдж v23.0'");rep(".ver').textContent, 'v22.0'",".ver').textContent, 'v23.0'")
rep("has('25: справка: обоснованность (v22.0)'","has('25: справка: обоснованность (v23.0)'")
rep("нет ошибок окна после всех рендеров (v22.0)","нет ошибок окна после всех рендеров (v23.0)")
a=s.index("console.log('21. Печатный бланк анкеты");b=s.index("console.log('22. Представления",a)
s=s[:a]+"""console.log('21. Единая форма организации (Приложение №2)');
api.demo('s1');
api.setModes(false, true);
ck('бланк открыт', api.printForm(), null, true);
const pf = w.document.querySelector('#tab-ank .orgform');
ck('бланк: форма отрисована', !!pf, null, 'есть');
has('бланк: заголовок приложения', pf.textContent, 'Приложение №2');
eq('бланк: ровно 2 страницы', pf.querySelectorAll('.orgpage').length, 2);
eq('бланк: 22 основных + 5 СФ + 9 КВ = 36 строк ответов', pf.querySelectorAll('table.ot td.num').length, 36);
has('бланк: стоп-факторы', pf.textContent, 'СТОП-ФАКТОРЫ');
has('бланк: красные вопросы', pf.textContent, '«КРАСНЫЕ» ВОПРОСЫ');
ck('бланк: без блока Р', pf.textContent.indexOf('БЛОК РЕЗУЛЬТАТИВНОСТИ') < 0, null, 'нет');
ck('бланк: без подписи эксперта', pf.textContent.indexOf('Подпись эксперта') < 0, null, 'нет');
ck('бланк: без Н/О в шкале', pf.textContent.indexOf('Н/О') < 0, null, 'нет');
ck('бланк: отмечен текущий ответ А1=5', !!pf.querySelector('td.circ.mark'), null, 'есть');
api.setModes(false, false);
api.goTab('ank');

""" +s[b:]
fin="console.log();\nconsole.log('ИТОГ:"
add="""console.log('26. Единая форма v23: числовые СФ/КВ, миграция и исправление КВ7–8');
api.reset();
[1,2,3].forEach(e => fillAll(e,5,5));
ck('normSF Да→1', api.normSF('Да'), 1, 1);
ck('normSF Нет→5', api.normSF('Нет'), 5, 5);
ck('normSF Неопределённо→3', api.normSF('Неопределённо'), 3, 3);
ck('normSF 2 сохраняется', api.normSF(2), 2, 2);
eq('normSF пусто≠Нет', api.normSF(''), null);
api.setSF(1,1,2); api.setSF(2,1,3); api.setSF(3,1,4);
R = api.calc();
eq('СФ1: выявлен у 1 (ответ2)', R.sf[0].yes, 1);
eq('СФ1: неопределённо у 1 (ответ3)', R.sf[0].unc, 1);
eq('СФ1: не выявлен у 1 (ответ4)', R.sf[0].no, 1);
eq('СФ оверрайд сработал', R.sfDet, true);
api.setSF(1,1,5); api.setSF(2,1,5); api.setSF(3,1,5);
[1,2,3].forEach(e => { api.setKV(e,7,1); api.setKV(e,8,2); });
R = api.calc();
ck('КВ7 риск при 1', api.kvSig(7,1), true, true);
ck('КВ8 риск при 2', api.kvSig(8,2), true, true);
eq('КВ7 нет риска при 4', api.kvSig(7,4), false);
eq('КВ7 групповой', R.kv.filter(k=>k.n===7)[0].group, true);
eq('КВ8 групповой', R.kv.filter(k=>k.n===8)[0].group, true);
api.setKV(1,7,3); api.setKV(2,7,3); api.setKV(3,7,3);
R = api.calc();
eq('КВ7 ответ3 не сигнал', R.kv.filter(k=>k.n===7)[0].sig, 0);
has('КВ7 ответ3 в повестке', R.agenda.join(' '), 'КВ7');
const legacy = JSON.parse(api.exportExpert(1)); legacy.expert.sf['1']='Да'; legacy.expert.sf['2']='Неопределённо'; legacy.expert.sf['3']='Нет';
ck('импорт legacy СФ', api.importExpert(JSON.stringify(legacy)), null, true);
eq('legacy Да→1', api.state().exp[0].sf[1], 1);
eq('legacy Неопределённо→3', api.state().exp[0].sf[2], 3);
eq('legacy Нет→5', api.state().exp[0].sf[3], 5);
eq('orgForm основных А=10', api.orgForm.main.filter(x=>x.code[0]==='А').length, 10);
eq('orgForm основных Ц=12', api.orgForm.main.filter(x=>x.code[0]==='Ц').length, 12);

"""
assert s.count(fin)==1;s=s.replace(fin,add+fin)
io.open('scripts/tests/test_v23.js','w',encoding='utf-8').write(s)
print('test_v23.js',len(s))
