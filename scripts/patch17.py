# -*- coding: utf-8 -*-
# Этап 17 → v17.0: выгрузка заполненного чек-листа (31 вопрос) в Markdown для Word.
# Мастер-патч: перезапускаемый, все замены с assert count==N.
import io

P = "output/Аудит_НИОКР_интерактивная_модель.html"
s = io.open(P, encoding="utf-8").read()

def rep(old, new, n=1):
    global s
    c = s.count(old)
    assert c == n, "MISS(%d!=%d): %r" % (c, n, old[:90])
    s = s.replace(old, new)

# ---------- 1. версии ----------
rep("интерактивная модель оценки, v16.0</title>", "интерактивная модель оценки, v17.0</title>")
rep("v16.0 · этап 16", "v17.0 · этап 17")
rep("аудита НИОКР v16.0", "аудита НИОКР v17.0", 8)
rep("НИОКР_аудит_v16.json", "НИОКР_аудит_v17.json", 2)

# ---------- 2. welcome ----------
rep("Добро пожаловать в модель аудита НИОКР (v16.0)</h2>", "Добро пожаловать в модель аудита НИОКР (v17.0)</h2>")
V17P = ("<p><b>Новое в v17.0 — выгрузка заполненного чек-листа в Markdown (для Word):</b> кнопки «⬇ Чек-лист (MD)» "
        "в экспертной карточке чек-листа и на вкладке «🖨 Акт оценки» сохраняют документ «Чек-лист_НИОКР_&lt;шифр&gt;_31.md»: "
        "все 31 вопрос с ответами, подтверждениями и комментариями по блокам (Pⱼ), интегральный показатель I и категория, "
        "веса, доказательность D/K, экспертная оценка (I_А и I_Ц с уровнями, стоп-факторы, «красные» флаги), "
        "экспертное заключение из 5 обязательных позиций и решение комиссии — готовое приложение к акту. "
        "Пакет документов синхронизирован с моделью: чек-лист 31 вопрос (файл 2, редакция v16), стандарт презентации "
        "(файлы 4 и 5, редакция v16 — 12 слайдов, включая матрицу «А × Ц»), Excel-инструмент методики (4 листа, "
        "включая «Сравнение НИОКР»), полный текст методики (файл 8).</p>")
rep("<p><b>Новое в v16.0 —", V17P + "<p><b>Новое в v16.0 —")

# ---------- 3. KB: карточка вкладки «Акт оценки» ----------
rep("и «📋 Копировать отчёт комиссии» для Word.\"],",
    "и «📋 Копировать отчёт комиссии» для Word, «⬇ Чек-лист (MD)» — выгрузка заполненного чек-листа (31 вопрос с ответами, "
    "индексы I_А/I_Ц, экспертное заключение из 5 позиций) в Markdown для Word (новое в v17.0).\"],")

# ---------- 4. функции экспорта (после conclModal, перед блоком миграции) ----------
FN = '''/* -------- выгрузка заполненного чек-листа в Markdown (этап 17, v17.0) -------- */
function exportChecklistMD(r){
 const c=calc(r),ec=expertConclusion(r),ev=evidStats(r),F=formOf(r);
 const L=[];
 L.push("# ЧЕК-ЛИСТ АУДИТА НИОКР — 31 ВОПРОС (заполненный)");
 L.push("");
 L.push("**НИОКР:** «"+r.name+"» (шифр: "+(r.shifr||"—")+")  ");
 L.push("**Подразделение:** "+(r.podrazd||"—")+" · **Форма проведения:** "+(F?F.name:"—")+" · **Статус:** "+(r.status||"—")+"  ");
 L.push("**Дата формирования:** "+new Date().toLocaleString("ru-RU")+" · **Аудитор:** "+((state.me&&state.me.name)?state.me.name:"—")+" · **Модель:** v17.0");
 L.push("");
 L.push("## 1. Ответы по вопросам (31)");
 const BN=["I. Актуальность — экспертная методика","II. Эффективность","III. Целесообразность — экспертная методика"];
 [0,1,2].forEach(b=>{
  L.push("");
  L.push("### Блок "+BN[b]+" — P"+(b+1)+" = "+(c.bl[b].P==null?"Н/П":(100*c.bl[b].P).toFixed(1)+" %"));
  L.push("");
  L.push("| № | Вопрос | Балл | Подтв. | Комментарий |");
  L.push("| --- | --- | --- | --- | --- |");
  for(let i=0;i<NQ;i++){if(QS[i][2]!==b)continue;
   const a=(r.ans||[])[i],cf=(r.conf||[])[i],cm=(r.com||[])[i];
   const av=(a===""||a==null)?"—":(a==="НП"?"Н/П":String(a));
   L.push("| "+(i+1)+(QS[i][0].includes("★")?" ★":"")+" | "+QS[i][0].replace(/ ★$/,"")+" | "+av+" | "+(cf||"—")+" | "+((cm||"—").replace(/\\|/g,"/").replace(/\\n/g," "))+" |");}
 });
 L.push("");
 L.push("## 2. Интегральный показатель");
 L.push("");
 L.push("| Показатель | Значение |");
 L.push("| --- | --- |");
 L.push("| I (0–100) | "+(c.I==null?"Недостаточно данных":c.I.toFixed(1))+" |");
 L.push("| Категория | "+(c.cat||"—")+" |");
 L.push("| Веса блоков Wⱼ | "+state.weights.join(" / ")+" %"+(c.renorm?" (ренормированы — блок «Н/П»)":"")+" |");
 [0,1,2].forEach(b=>L.push("| Выполнение P"+(b+1)+" | "+(c.bl[b].P==null?"Н/П":(100*c.bl[b].P).toFixed(1)+" %")+" |"));
 L.push("| Доказательность D / K | "+(ev.D==null?"Н/П":ev.D.toFixed(0)+" %")+" / "+(ev.K==null?"—":ev.K.toFixed(0)+" %")+" |");
 L.push("");
 L.push("## 3. Экспертная оценка «Актуальность + Целесообразность»");
 L.push("");
 L.push("| Параметр | Значение |");
 L.push("| --- | --- |");
 L.push("| I_А (оценено "+c.nA+" из 10 критериев) | "+(ec.iA==null?"не оценён (недостаточно данных)":ec.iA.toFixed(1)+" % — «"+ec.lA+"»")+" |");
 L.push("| I_Ц (оценено "+c.nC+" из 12 критериев) | "+(ec.iC==null?"не оценён (недостаточно данных)":ec.iC.toFixed(1)+" % — «"+ec.lC+"»")+" |");
 L.push("| Стоп-факторы (5) | "+(ec.stops.length?"СРАБОТАЛИ: "+ec.stops.join("; "):"не сработали (0 из 5)")+" |");
 L.push("| «Красные» флаги (10) | "+(ec.flags.length?"сработало "+ec.flags.length+" из 10: "+ec.flags.join("; "):"не сработали (0 из 10)")+" |");
 L.push("");
 L.push("## 4. Экспертное заключение (5 обязательных позиций)");
 L.push("");
 L.push("1. Уровень актуальности: "+(ec.iA==null?"не оценён (недостаточно данных)":ec.iA.toFixed(1)+" % — «"+ec.lA+"»")+".");
 L.push("2. Уровень целесообразности: "+(ec.iC==null?"не оценён (недостаточно данных)":ec.iC.toFixed(1)+" % — «"+ec.lC+"»")+".");
 L.push("3. Проблемы с представленными материалами: "+ec.problems+".");
 L.push("4. Перспективы завершения: "+ec.prospect+".");
 L.push("5. Рекомендация: "+ec.rec+".");
 L.push("");
 L.push("## 5. Решение комиссии и контрольные условия");
 L.push("");
 L.push("| Параметр | Значение |");
 L.push("| --- | --- |");
 L.push("| Решение комиссии | "+(r.decision||"—")+" |");
 L.push("| Обоснование решения | "+((r.just||"—").replace(/\\|/g,"/"))+" |");
 L.push("| Меры | "+((r.measures||"—").replace(/\\|/g,"/"))+" |");
 const ct=(r.ctrl||[]).map((v,i)=>(v&&String(v).trim()&&String(v).trim()!=="—"?(i+1)+") "+String(v).trim():null)).filter(Boolean);
 L.push("| Контрольные условия | "+(ct.length?ct.join("; "):"—")+" |");
 L.push("");
 L.push("---");
 L.push("*Сформировано интерактивной моделью аудита НИОКР v17.0 · "+new Date().toLocaleDateString("ru-RU")+" · чек-лист из 31 вопроса (редакция v16). Документ копируется в Word целиком. Подписи: Аудитор ____________ / Представитель подразделения ____________ / Дата ____________*");
 return L.join("\\n");}
function btnChecklistMD(r){
 const safe=(r.shifr||"без_шифра").replace(/[\\\\/:*?"<>|\\s]+/g,"_");
 download("Чек-лист_НИОКР_"+safe+"_31.md",exportChecklistMD(r),"text/markdown;charset=utf-8");
 notify("ok","Чек-лист (MD) сохранён","31 вопрос с ответами, I_А/I_Ц и экспертное заключение — для Word");}
/* -------- миграция 27 → 31 (этап 15): старые блоки I/III переносятся в новую структуру -------- */'''
rep("/* -------- миграция 27 → 31 (этап 15): старые блоки I/III переносятся в новую структуру -------- */", FN)

# ---------- 5. кнопки: экспертная карточка + акт ----------
rep("'<button class=\"btn sm primary\" id=\"exConcl\">📋 Экспертное заключение (5 позиций)</button></div>';",
    "'<button class=\"btn sm primary\" id=\"exConcl\">📋 Экспертное заключение (5 позиций)</button>'+'<button class=\"btn sm\" id=\"exMd\" title=\"Заполненный чек-лист (31 вопрос, ответы, индексы I_А/I_Ц, заключение) в Markdown — для Word\">⬇ Чек-лист (MD)</button></div>';")
rep('const exC=document.getElementById("exConcl");if(exC)exC.onclick=()=>conclModal(cur());',
    'const exC=document.getElementById("exConcl");if(exC)exC.onclick=()=>conclModal(cur());\n'+
    '  const exM=document.getElementById("exMd");if(exM)exM.onclick=()=>btnChecklistMD(cur());')
rep('📋 Копировать отчёт комиссии</button> <span class="mini noprint">история:',
    '📋 Копировать отчёт комиссии</button> <button class="btn noprint" id="repMd" title="Заполненный чек-лист (31 вопрос, ответы, индексы I_А/I_Ц, заключение) в Markdown — для Word">⬇ Чек-лист (MD)</button> <span class="mini noprint">история:')
rep('const rx2=document.getElementById("repComCopy");if(rx2)rx2.onclick=copyCom;',
    'const rx2=document.getElementById("repComCopy");if(rx2)rx2.onclick=copyCom;\n'+
    '  const rm2=document.getElementById("repMd");if(rm2)rm2.onclick=()=>btnChecklistMD(cur());')

# ---------- 6. niokrAPI ----------
rep("expertCard,expertDashCard,expKBHTML};",
    "expertCard,expertDashCard,expKBHTML,exportChecklistMD,btnChecklistMD};")

io.open(P, "w", encoding="utf-8").write(s)
print("PATCH17 OK, chars:", len(s))
