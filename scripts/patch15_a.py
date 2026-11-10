# -*- coding: utf-8 -*-
p='output/Аудит_НИОКР_интерактивная_модель.html'
s=open(p,encoding='utf-8').read()
def rep1(old,new):
    global s
    assert s.count(old)==1, "NOT UNIQUE/FOUND: "+old[:70]
    s=s.replace(old,new)

# (a) ANCH 31 rows
i=s.index('const ANCH=['); j=s.index('];',i)+2
oldB=s[i:j].split('\n')[10:19]
assert len(oldB)==9
anch='''const ANCH=[
["спрос подтверждён документами заказчика","потребность частична, нужна модернизация","потребность отсутствует"],
["включена в стратегию и программы развития","связь со стратегией слабая","вне стратегии организации"],
["более новых решений не выявлено","есть отдельные новые решения","результат морально устарел"],
["уникальное решение и компетенция","новизна частично утрачена","повторяет известные решения"],
["аналогичных разработок нет — поиск проведён","частичное пересечение со смежными работами","дублирует другую разработку"],
["технологическое преимущество подтверждено","преимущество постепенно теряется","аналоги превосходят результат"],
["соответствие актуальным и вводимым нормам проверено","отдельные несоответствия нормам","не соответствует нормативам"],
["к завершению работ результат не устареет — прогноз подтверждён","умеренный риск устаревания, нужна адаптация","к моменту завершения результат устареет"],
["соответствует технологическому прогнозу отрасли; свежий внешний анализ","направление перспективно частично, анализ устарел","направление бесперспективно либо только мнение исполнителя"],
["потребность на 3–5 лет подтверждена планами заказчика","потребность вероятна, но не документирована","потребность за горизонтом завершения отсутствует"],
'''+ '\n'.join(oldB) +''',
["цель достижима — подтверждена расчётами и экспериментами","достижимость подтверждена частично, есть критические операции","цель недостижима выбранным способом"],
["все операции отработаны или подтверждены исследованиями","часть операций требует дооснащения или НИР","ключевые технологии не подтвернены"],
["стоимость и финансирование до завершения подтверждены; результат стоит остатка затрат","финансирование с разрывами, стоимость растёт","продолжение дороже ценности результата — ловушка невозвратных затрат"],
["штат укомплектован, компетенции сохраняются","недостаёт части специалистов, есть риск ухода","ключевые исполнители утрачены"],
["производственная, опытная и испытательная база готовы","база частична, требуется дооснащение","база для работ и внедрения отсутствует"],
["доступ ко всем материалам и комплектующим подтверждён","часть комплектующих под риском, план замены в работе","санкционные компоненты без плана импортозамещения"],
["потребитель и сроки внедрения подтверждены документально","потребитель определяется, интерес устный","подтверждённый потребитель отсутствует"],
["организационно-технические условия внедрения созданы","внедрение возможно после дополнительной подготовки","внедрение в производство практически невозможно"],
["эффект подтверждён расчётом, «стоимость — результат» положительна","оценка эффекта качественная","затраты заведомо превышают эффект"],
["альтернативы сопоставлены, своя разработка выгоднее","сравнение частичное","альтернативы не рассматривались либо покупка/модернизация выгоднее"],
["риски идентифицированы, меры и ответственные назначены","реестр рисков неполный","риски незавершения не оценивались"],
["результат — основа линейки продуктов, РИД оформлены","потенциал развития частичный, РИД в процессе","результат разовый, перспектив развития нет"]];'''
s=s[:i]+anch+s[j:]

# (b) calcCore guard
rep1('    for(let i=0;i<27;i++){if(QS[i][2]!==b)continue;const a=ans[i];if(a===""||a==="НП")continue;S+=Number(a);N++;}',
     '    for(let i=0;i<NQ;i++){if(QS[i][2]!==b)continue;const a=ans[i];if(a===""||a==null||a==="НП")continue;const v=Number(a);if(!isFinite(v))continue;S+=v;N++;}')

# (c) calc extension + expert engine + migration
ENGINE=open('scripts/engine15.js',encoding='utf-8').read()
rep1('''function calc(r){
  const c=calcCore(r.ans,state.weights);
  c.nCtrl=r.ctrl.filter(x=>x==="Выявлено").length;
  return c;
}''','''function calc(r){
  const c=calcCore(r.ans,state.weights);
  c.nCtrl=r.ctrl.filter(x=>x==="Выявлено").length;
  const dA=dirIdx(r,0),dC=dirIdx(r,2);
  c.iA=dA.I;c.iC=dC.I;c.nA=dA.n;c.nC=dC.n;c.lA=lvl5(dA.I);c.lC=lvl5(dC.I);
  const er=expertRec(r);c.rec=er.text;c.recStop=er.stop;c.stopsHit=er.st;
  c.prospect=expertProspect(r);
  return c;
}
'''+ENGINE)

# (e) newRnd
rep1('ans:Array(27).fill(""),conf:Array(27).fill(""),com:Array(27).fill(""),cat:"",docs:[]',
     'ans:Array(NQ).fill(""),conf:Array(NQ).fill(""),com:Array(NQ).fill(""),stops:Array(5).fill(""),flags:Array(10).fill(""),prospect:"",cat:"",docs:[]')

# (f) load + btnImp
rep1('if(!r.props)r.props=[];if(!r.created)r.created=new Date().toISOString();});return true;}}catch(e){}return false;}',
     'if(!r.props)r.props=[];if(!r.created)r.created=new Date().toISOString();if(migrate27to31(r))migN++;});\n  if(migN)setTimeout(()=>notify("info","Выполнена миграция данных (v15.0)",migN+" НИОКР перенесено со структуры 27 → 31 вопрос: блоки I и III пересобраны по экспертной методике «Актуальность + Целесообразность»; снимки истории аудитов сохранены как есть"),900);\n  return true;}}catch(e){}return false;}')
rep1('function load(){try{const d=JSON.parse(localStorage.getItem(LSKEY));if(d&&d.list&&d.list.length){state.list=d.list;state.cur=d.cur||0;',
     'function load(){try{const d=JSON.parse(localStorage.getItem(LSKEY));if(d&&d.list&&d.list.length){let migN=0;state.list=d.list;state.cur=d.cur||0;')
rep1('state.list.forEach(r=>{if(r.flow===undefined)r.flow=null;if(!r.threads)r.threads={};if(r.resp===undefined)r.resp=null;if(!r.audits)r.audits=[];if(!r.props)r.props=[];if(!r.created)r.created=new Date().toISOString();});\n    save();renderAll();notify("ok","Импорт выполнен","загружено НИОКР: "+state.list.length+" · веса "+fmtW(state.weights));',
     'let migN=0;state.list.forEach(r=>{if(r.flow===undefined)r.flow=null;if(!r.threads)r.threads={};if(r.resp===undefined)r.resp=null;if(!r.audits)r.audits=[];if(!r.props)r.props=[];if(!r.created)r.created=new Date().toISOString();if(migrate27to31(r))migN++;});\n    save();renderAll();notify("ok","Импорт выполнен","загружено НИОКР: "+state.list.length+" · веса "+fmtW(state.weights)+(migN?" · миграция 27→31: "+migN+"":""));')

# (g) guards
rep1('  for(let i=0;i<27;i++){const x=a.ans[i],y=b.ans[i];if(x!==y)rows.push(\'<tr><td><b>\'+(i+1)+\'</b> <span class="',
     '  for(let i=0;i<Math.max((a.ans||[]).length,(b.ans||[]).length);i++){const x=(a.ans||[])[i]||"",y=(b.ans||[])[i]||"";if(x!==y)rows.push(\'<tr><td><b>\'+(i+1)+\'</b> <span class="')
rep1('      for(let j=0;j<27;j++){A2[j]=a.ans[j];C2[j]=a.conf&&a.conf[j]?a.conf[j]:"";}',
     '      for(let j=0;j<NQ;j++){A2[j]=a.ans&&a.ans[j]!=null?a.ans[j]:"";C2[j]=a.conf&&a.conf[j]?a.conf[j]:"";}')

# (h) sweep loops
n1=s.count('i<27'); n2=s.count('j<27'); n3=s.count('Array(27)')
s=s.replace('i<27','i<NQ').replace('j<27','j<NQ').replace('Array(27)','Array(NQ)')
print("swept:",n1,n2,n3)

# (i) text
pairs=[
 ('✅ Чек-лист (27)','✅ Чек-лист (31)'),
 ('Тепловая карта ответов (вопросы 1–27)','Тепловая карта ответов (вопросы 1–31)'),
 (' из 27',' из 31'),
 (' / 27<',' / 31<'),
 ('"/27, личный','"/31, личный'),
 ('по всем 27 вопросам','по всем 31 вопросу'),

 ('чек-лист аудита НИОКР v2.0 (27 вопросов)','чек-лист аудита НИОКР v15.0 (31 вопрос: экспертные блоки «Актуальность» — 10 и «Целесообразность» — 12, блок «Эффективность» — 9)'),
 ('9 вопросов × 5 баллов = 45 «сырых» баллов блока','блоки 10 / 9 / 12 вопросов × 5 баллов = 50 / 45 / 60 «сырых» баллов'),
 ("</b> / 9</span>","</b> / '+[NA,NB,NC][b]+'+</span>"),
]
for old,new in pairs:
    c=s.count(old)
    assert c>=1,"NOT FOUND: "+old
    s=s.replace(old,new)
    print(repr(old),"x",c)
open(p,'w',encoding='utf-8').write(s)
print("SAVED OK")
