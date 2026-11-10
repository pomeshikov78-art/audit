# -*- coding: utf-8 -*-
"""Генератор: 7_Оценка_актуальности_целесообразности_НИОКР.xlsx (этап 15; этап 16 — лист «Сравнение НИОКР»)
4 листа: 1 «Оценка» (ввод ответов, автобаллы, индексы I_А/I_Ц, стоп-факторы, «красные» вопросы),
2 «Матрица решений» (сопоставление уровней А×Ц, блокировка стоп-факторами, перспективы),
3 «Заключение» (генератор текста из 5 обязательных позиций для Word/протокола),
4 «Сравнение НИОКР» (портфельный обзор до 10 разработок: индексы → авто-уровни, рекомендации матрицы, перспективы, группы решений, сводные счётчики)."""
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule,FormulaRule

CRIT_A=[
 ("Устойчивость потребности в результатах","Сохраняется ли потребность заказчика в результатах НИОКР в настоящее время? Подтверждена ли документально?"),
 ("Соответствие задачам организации","Решает ли НИОКР актуальные производственно-технические и продуктовые задачи организации?"),
 ("Научно-технический уровень результата","Соответствует ли результат современному уровню науки и техники, конкурентоспособен ли он?"),
 ("Технологическая новизна","Содержит ли НИОКР новые технические решения по сравнению с известными аналогами?"),
 ("Отсутствие дублирования","Нет ли аналогичных разработок внутри организации или в рамках кооперации?"),
 ("Наличие аналогов на рынке","Существуют ли готовые решения, уже удовлетворяющие потребность? Не превосходят ли они разработку? (стоп-фактор 1)"),
 ("Нормативная актуальность","Отвечает ли результат действующим и вводимым нормативным требованиям, стандартам, требованиям заказчика?"),
 ("Актуальность к моменту завершения","Не устареет ли результат к моменту завершения НИОКР?"),
 ("Перспективность технологии","Соответствует ли технология перспективным направлениям отрасли и технической политики организации?"),
 ("Устойчивость потребности (3–5 лет)","Подтверждена ли потребность в результате на горизонте 3–5 лет, в т.ч. после завершения НИОКР? (стоп-фактор 2)"),
]
CRIT_C=[
 ("Достижимость результата","Достижима ли цель выбранным способом в установленные сроки?"),
 ("Техническая реализуемость","Подтверждена ли расчётами/экспериментами выполнимость всех технологических операций и требований?"),
 ("Финансовая обоснованность","Подтверждены ли стоимость и схема финансирования, обеспечены ли средства до завершения работ? (стоп-факторы 3, 5)"),
 ("Кадровая обеспеченность","Достаточно ли специалистов требуемой квалификации? Есть ли риск потери ключевых исполнителей?"),
 ("Производственная база","Имеются ли производственная, опытная и испытательная база для работ и внедрения?"),
 ("Обеспеченность ресурсами","Есть ли доступ к материалам, комплектующим и оборудованию? Учтены ли санкционные риски и план импортозамещения? (стоп-фактор 4)"),
 ("Наличие реального потребителя","Есть ли подтверждённый потребитель, готовый использовать результат? (стоп-фактор 2)"),
 ("Возможность внедрения","Созданы ли организационно-технические условия для внедрения результата в производство?"),
 ("Экономический эффект","Обоснован ли расчётом ожидаемый эффект? Подтверждена ли связка «стоимость — результат»?"),
 ("Альтернативные решения","Сопоставлены ли альтернативы: покупка, лицензия, собственная разработка, модернизация?"),
 ("Риски завершения","Идентифицированы ли основные риски недостижения результата? Есть ли меры снижения?"),
 ("Дальнейшее развитие","Создаёт ли результат основу для последующих разработок и новых продуктов?"),
]
STOPS=[
 "Появился рыночный аналог, превосходящий разрабатываемый результат по характеристикам и стоимости",
 "Ключевой заказчик официально отказался от потребности в результате",
 "Продолжение НИОКР обусловлено исключительно «освоением» уже выделенного бюджета (ловушка невозвратных затрат)",
 "Технология или ключевые комплектующие под санкционными ограничениями, утверждённый план импортозамещения отсутствует",
 "Выявлены хищения или нецелевое использование средств",
]
FLAGS=[
 ("Не утратила ли НИОКР первоначальную актуальность в процессе выполнения?","Да"),
 ("Не изменились ли потребности заказчика после утверждения ТЗ?","Да"),
 ("Не появились ли более современные и дешёвые технические решения?","Да"),
 ("Не ведётся ли аналогичная НИОКР другим подразделением или предприятием?","Да"),
 ("Не превышают ли оставшиеся затраты потенциальную ценность результата?","Да"),
 ("Можно ли получить требуемый результат модернизацией существующего изделия?","Да"),
 ("Есть ли реальные перспективы запуска результата в производство?","Нет"),
 ("Будет ли результат востребован через 3–5 лет?","Нет"),
 ("Не вызвано ли продолжение НИОКР исключительно уже понесёнными затратами?","Да"),
 ("Сохранит ли организация необходимые компетенции после завершения работ?","Нет"),
]
RECMX=[
 ["Продолжить работы в обычном режиме; рассмотреть ускорение НИОКР","Продолжить при условии устранения выявленных замечаний","Продолжить с корректировкой ресурсов и исполнителей, усилением контроля","Пересмотреть технические решения, исполнителей, ресурсы и сроки","Пересмотреть технические решения, исполнителей, ресурсы и сроки"],
 ["Продолжить работы в обычном режиме","Продолжить при условии устранения выявленных замечаний","Продолжить с корректировкой ресурсов и исполнителей, усилением контроля","Пересмотреть технические решения, исполнителей, ресурсы и сроки","Пересмотреть решения; при подтверждении нереализуемости — вопрос о прекращении"],
 ["Продолжить после уточнения потребности и потребителя","Продолжить при условии устранения выявленных замечаний","Пересмотреть цели, требования и параметры реализации НИОКР","Рассмотреть приостановку для дополнительного изучения","Рассмотреть приостановку для дополнительного изучения"],
 ["Проверить возможность изменения назначения или требований (поиск нового потребителя)","Проверить возможность изменения назначения или требований; при невозможности — приостановить","Пересмотреть цели, требования и параметры реализации НИОКР","Рассмотреть прекращение НИОКР","Прекратить НИОКР"],
 ["Проверить возможность изменения назначения или требований","Рассмотреть приостановку для дополнительного изучения","Рассмотреть прекращение НИОКР","Прекратить НИОКР","Прекратить НИОКР (рассмотреть списание результата в установленном порядке)"],
]
LV=["Высокая","Хорошая","Удовлетворительная","Низкая","Критическая"]

NAVY="1F3864";BLUE="2E74B5";LBLUE="EEF4FB";GOLD="FBC02D";RED="C62828";LRED="FDECEA";GREEN="2E7D32";LGREEN="EAF4EA";ORANGE="EF6C00";LGRAY="F2F2F2"
thin=Side(style="thin",color="B8C4D9")
BORD=Border(left=thin,right=thin,top=thin,bottom=thin)
def st(c,*,b=False,sz=10,color="20242C",fill=None,wrap=True,ha="left",va="center",border=True):
    c.font=Font(bold=b,size=sz,color=color,name="Calibri")
    if fill:c.fill=PatternFill("solid",fgColor=fill)
    c.alignment=Alignment(wrap_text=wrap,horizontal=ha,vertical=va)
    if border:c.border=BORD

wb=openpyxl.Workbook()
# ================= ЛИСТ 1: ОЦЕНКА =================
ws=wb.active;ws.title="Оценка"
widths={"A":5,"B":30,"C":46,"D":16,"E":9,"F":34,"G":14,"H":14}
for k,v in widths.items():ws.column_dimensions[k].width=v
ws.merge_cells("A1:H1");st(ws["A1"],b=True,sz=14,color="FFFFFF",fill=NAVY,ha="center");ws["A1"]="ОЦЕНКА АКТУАЛЬНОСТИ И ЦЕЛЕСООБРАЗНОСТИ ПРОДОЛЖЕНИЯ НИОКР"
ws.row_dimensions[1].height=26
ws.merge_cells("A2:H2");st(ws["A2"],sz=9,fill=LBLUE);ws["A2"]=("Самостоятельный экспертный инструмент. Оценки актуальности (А) и целесообразности (Ц) НЕ складываются — они сопоставляются через матрицу решений (лист «Матрица решений»). "
 "Шкала ответов: Да = 5 · В основном да = 4 · Частично = 3 · Скорее нет = 1 · Нет = 0 · Н/П — исключается из расчёта (балл «2» не используется намеренно). "
 "Индекс направления: I = ΣB / (5·n) · 100 %. Уровни: ≥ 85 Высокая · 70–84 Хорошая · 50–69 Удовлетворительная · 30–49 Низкая · < 30 Критическая.")
ws.row_dimensions[2].height=52
ws["A3"]="НИОКР:";st(ws["A3"],b=True,border=False);ws.merge_cells("B3:D3");st(ws["B3"],fill="FFFDE7");ws["B3"]=""
ws["E3"]="Шифр:";st(ws["E3"],b=True,border=False);st(ws["F3"],fill="FFFDE7")
ws["G3"]="Дата:";st(ws["G3"],b=True,border=False);st(ws["H3"],fill="FFFDE7")
ws["A4"]="Эксперт:";st(ws["A4"],b=True,border=False);ws.merge_cells("B4:D4");st(ws["B4"],fill="FFFDE7")
ws["E4"]="Должность:";st(ws["E4"],b=True,border=False);ws.merge_cells("F4:H4");st(ws["F4"],fill="FFFDE7")

ANS_LIST='"Да,В основном да,Частично,Скорее нет,Нет,Н/П"'
YESNO='"Да,Нет"'
dv_ans=DataValidation(type="list",formula1=ANS_LIST,allow_blank=True,showDropDown=False)
dv_yn=DataValidation(type="list",formula1=YESNO,allow_blank=True,showDropDown=False)
ws.add_data_validation(dv_ans);ws.add_data_validation(dv_yn)

def crit_header(row):
    ws.merge_cells(f"A{row}:A{row+0}")
    hdr=[("A","№"),("B","Критерий"),("C","Вопрос для экспертизы"),("D","Ответ"),("E","Балл"),("F","Обоснование / подтверждающие документы")]
    ws.merge_cells(f"F{row}:H{row}")
    for col,txt in hdr:
        c=ws[f"{col}{row}"];c.value=txt;st(c,b=True,color="FFFFFF",fill=BLUE,ha="center")
    ws.row_dimensions[row].height=16

def lvl_formula(idx_cell):
    return (f'=IF({idx_cell}="—","—",IF({idx_cell}>=85,"Высокая",IF({idx_cell}>=70,"Хорошая",'
            f'IF({idx_cell}>=50,"Удовлетворительная",IF({idx_cell}>=30,"Низкая","Критическая")))))')

def crit_block(start_row,crits,letter):
    crit_header(start_row)
    r=start_row+1
    for i,(name,quest) in enumerate(crits,1):
        ws[f"A{r}"]=i;st(ws[f"A{r}"],ha="center")
        ws[f"B{r}"]=name;st(ws[f"B{r}"],b=True)
        ws[f"C{r}"]=quest;st(ws[f"C{r}"],sz=9)
        st(ws[f"D{r}"],fill="FFFDE7",ha="center");dv_ans.add(ws[f"D{r}"])
        ws[f"E{r}"]=f'=IF(D{r}="","",IF(D{r}="Н/П","",VLOOKUP(D{r},$B$66:$C$71,2,0)))'
        st(ws[f"E{r}"],ha="center",b=True)
        ws.merge_cells(f"F{r}:H{r}");st(ws[f"F{r}"],fill="FFFDE7",sz=9)
        ws.row_dimensions[r].height=30
        r+=1
    last=r-1
    # индексная строка
    ws.merge_cells(f"A{r}:C{r}")
    ws[f"A{r}"]=f"ИНДЕКС {'АКТУАЛЬНОСТИ I_А' if letter=='A' else 'ЦЕЛЕСООБРАЗНОСТИ I_Ц'} = ΣB / (5·n) · 100 %,  где n — число оценённых критериев (без «Н/П»)"
    st(ws[f"A{r}"],b=True,fill=LBLUE)
    ws[f"D{r}"]=f"=COUNT(E{start_row+1}:E{last})&\" из {len(crits)}\"";st(ws[f"D{r}"],ha="center",fill=LBLUE)
    idx=f"E{r}"
    ws[idx]=f'=IFERROR(ROUND(SUM(E{start_row+1}:E{last})/(5*COUNT(E{start_row+1}:E{last}))*100,1),"—")'
    st(ws[idx],b=True,sz=12,ha="center",fill="FFFDE7")
    ws[f"F{r}"]="Уровень:";st(ws[f"F{r}"],b=True,ha="right",fill=LBLUE)
    ws.merge_cells(f"G{r}:H{r}")
    ws[f"G{r}"]=lvl_formula(idx);st(ws[f"G{r}"],b=True,ha="center",fill="FFFDE7")
    ws.row_dimensions[r].height=24
    # условное форматирование баллов
    rng=f"E{start_row+1}:E{last}"
    ws.conditional_formatting.add(rng,CellIsRule(operator="greaterThanOrEqual",formula=["4"],fill=PatternFill("solid",fgColor="C8E6C9")))
    ws.conditional_formatting.add(rng,CellIsRule(operator="equal",formula=["3"],fill=PatternFill("solid",fgColor="FFF9C4")))
    ws.conditional_formatting.add(rng,CellIsRule(operator="lessThanOrEqual",formula=["1"],fill=PatternFill("solid",fgColor="FFCCBC")))
    # цвет уровня
    lrng=f"G{r}:H{r}"
    for txt,col in [("Высокая","C8E6C9"),("Хорошая","DCEDC8"),("Удовлетворительная","FFF9C4"),("Низкая","FFE0B2"),("Критическая","FFCDD2")]:
        ws.conditional_formatting.add(lrng,CellIsRule(operator="equal",formula=[f'"{txt}"'],fill=PatternFill("solid",fgColor=col)))
    return r  # row of index

ws.merge_cells("A6:H6");st(ws["A6"],b=True,sz=11,color="FFFFFF",fill=NAVY);ws["A6"]="НАПРАВЛЕНИЕ 1. АКТУАЛЬНОСТЬ — сохраняется ли потребность в результате и его ценность (10 критериев)"
rowA=crit_block(7,CRIT_A,"A")           # index row = 18
rS=rowA+2                                # 20
ws.merge_cells(f"A{rS}:H{rS}");st(ws[f"A{rS}"],b=True,sz=11,color="FFFFFF",fill=NAVY)
ws[f"A{rS}"]="НАПРАВЛЕНИЕ 2. ЦЕЛЕСООБРАЗНОСТЬ — стоит ли продолжать НИОКР (12 критериев)"
rowC=crit_block(rS+1,CRIT_C,"C")        # header rS+1=21, crit 22..33, index 34

# стоп-факторы
rF=rowC+2  # 36
ws.merge_cells(f"A{rF}:H{rF}");st(ws[f"A{rF}"],b=True,sz=11,color="FFFFFF",fill=RED)
ws[f"A{rF}"]="СТОП-ФАКТОРЫ — ответ «Да» БЛОКИРУЕТ рекомендацию матрицы: внеочередное рассмотрение руководством, вплоть до приостановки работ"
hdr=["A","№"];ws[f"A{rF+1}"]="№";ws[f"B{rF+1}"]="Стоп-фактор";ws.merge_cells(f"B{rF+1}:C{rF+1}");ws[f"D{rF+1}"]="Ответ"
ws.merge_cells(f"E{rF+1}:H{rF+1}");ws[f"E{rF+1}"]="Последствие при ответе «Да»"
for col in ["A","B","D","E"]:st(ws[f"{col}{rF+1}"],b=True,color="FFFFFF",fill=BLUE,ha="center")
CONS=["пересмотр решения о продолжении; сравнение с аналогом в заключении","фиксация отказа потребителя; пересмотр назначения результата","проверка финансовой службой; оценка остатка работ по существу","план импортозамещения либо прекращение работ","служебная проверка; приостановка финансирования"]
for i,(sf,cons) in enumerate(zip(STOPS,CONS),1):
    r=rF+1+i
    ws[f"A{r}"]=i;st(ws[f"A{r}"],ha="center")
    ws.merge_cells(f"B{r}:C{r}");ws[f"B{r}"]=sf;st(ws[f"B{r}"],sz=9)
    st(ws[f"D{r}"],fill="FFFDE7",ha="center",b=True);dv_yn.add(ws[f"D{r}"])
    ws.merge_cells(f"E{r}:H{r}");ws[f"E{r}"]=cons;st(ws[f"E{r}"],sz=9,fill=LGRAY)
    ws.row_dimensions[r].height=26
rStopTot=rF+7  # 43
ws.merge_cells(f"A{rStopTot}:C{rStopTot}");ws[f"A{rStopTot}"]="СРАБОТАЛО СТОП-ФАКТОРОВ:";st(ws[f"A{rStopTot}"],b=True,fill=LRED,ha="right")
ws[f"D{rStopTot}"]=f'=COUNTIF(D{rF+2}:D{rF+6},"Да")';st(ws[f"D{rStopTot}"],b=True,sz=12,ha="center",fill="FFFDE7")
ws.merge_cells(f"E{rStopTot}:H{rStopTot}");ws[f"E{rStopTot}"]=f'=IF(D{rStopTot}>0,"РЕКОМЕНДАЦИЯ ЗАБЛОКИРОВАНА — внеочередное рассмотрение руководством","Стоп-факторов не выявлено — действует матрица решений")';st(ws[f"E{rStopTot}"],b=True,fill=LRED)
ws.conditional_formatting.add(f"A{rF+2}:H{rF+6}",CellIsRule(operator="equal",formula=['$D'+str(rF+2)+'="Да"'],fill=PatternFill("solid",fgColor="FFCDD2")))

# красные вопросы
rQ=rStopTot+2  # 45
ws.merge_cells(f"A{rQ}:H{rQ}");st(ws[f"A{rQ}"],b=True,sz=11,color="FFFFFF",fill=ORANGE)
ws[f"A{rQ}"]="«КРАСНЫЕ» ВОПРОСЫ — детекторы «зомби-проектов»: НИОКР продолжается по инерции, хотя цель утрачена"
ws[f"A{rQ+1}"]="№";ws.merge_cells(f"B{rQ+1}:C{rQ+1}");ws[f"B{rQ+1}"]="Вопрос";ws[f"D{rQ+1}"]="Ответ";ws[f"E{rQ+1}"]="Тревожный ответ"
ws.merge_cells(f"F{rQ+1}:H{rQ+1}");ws[f"F{rQ+1}"]="Сработал?"
for col in ["A","B","D","E","F"]:st(ws[f"{col}{rQ+1}"],b=True,color="FFFFFF",fill=BLUE,ha="center")
for i,(fl,bad) in enumerate(FLAGS,1):
    r=rQ+1+i
    ws[f"A{r}"]=i;st(ws[f"A{r}"],ha="center")
    ws.merge_cells(f"B{r}:C{r}");ws[f"B{r}"]=fl;st(ws[f"B{r}"],sz=9)
    st(ws[f"D{r}"],fill="FFFDE7",ha="center");dv_yn.add(ws[f"D{r}"])
    ws[f"E{r}"]=bad;st(ws[f"E{r}"],ha="center",sz=9)
    ws.merge_cells(f"F{r}:H{r}");ws[f"F{r}"]=f'=IF(D{r}="","",IF(D{r}=E{r},"СРАБОТАЛ","—"))';st(ws[f"F{r}"],ha="center",b=True,sz=9)
    ws.row_dimensions[r].height=24
rQTot=rQ+12  # 57
ws.merge_cells(f"A{rQTot}:C{rQTot}");ws[f"A{rQTot}"]="СРАБОТАЛО «КРАСНЫХ» ФЛАГОВ:";st(ws[f"A{rQTot}"],b=True,ha="right",fill="FFF3E0")
ws[f"D{rQTot}"]=f'=COUNTIF(F{rQ+2}:F{rQ+11},"СРАБОТАЛ")';st(ws[f"D{rQTot}"],b=True,ha="center",fill="FFFDE7")
ws.merge_cells(f"E{rQTot}:H{rQTot}");ws[f"E{rQTot}"]="из 10 «красных» вопросов. Флаги не меняют баллы — это сигналы для углублённой проверки и обсуждения рекомендации";st(ws[f"E{rQTot}"],sz=9,fill=LGRAY)
ws.conditional_formatting.add(f"F{rQ+2}:F{rQ+11}",CellIsRule(operator="equal",formula=['"СРАБОТАЛ"'],fill=PatternFill("solid",fgColor="FFE0B2")))

# служебная таблица соответствия
rM=65
ws[f"A{rM}"]="Служебная таблица (не заполнять):";st(ws[f"A{rM}"],sz=8,border=False,color="8C8C8C")
ws[f"B{rM}"]="Ответ";ws[f"C{rM}"]="Балл"
st(ws[f"B{rM}"],b=True,sz=9,fill=LGRAY);st(ws[f"C{rM}"],b=True,sz=9,fill=LGRAY)
for i,(a,ball) in enumerate([("Да",5),("В основном да",4),("Частично",3),("Скорее нет",1),("Нет",0),("Н/П","")],1):
    ws[f"B{rM+i}"]=a;ws[f"C{rM+i}"]=ball
    st(ws[f"B{rM+i}"],sz=9);st(ws[f"C{rM+i}"],sz=9,ha="center")
ws.freeze_panes="A5"
ws.sheet_view.zoomScale=90
ws.page_setup.orientation="portrait";ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0
ws.print_area=f"A1:H{rQTot}"

# ================= ЛИСТ 2: МАТРИЦА РЕШЕНИЙ =================
m=wb.create_sheet("Матрица решений")
for k,v in {"A":22,"B":30,"C":30,"D":30,"E":30,"F":30,"G":16,"H":16}.items():m.column_dimensions[k].width=v
m.merge_cells("A1:F1");st(m["A1"],b=True,sz=13,color="FFFFFF",fill=NAVY,ha="center");m["A1"]="МАТРИЦА РЕШЕНИЙ «АКТУАЛЬНОСТЬ × ЦЕЛЕСООБРАЗНОСТЬ»"
m.row_dimensions[1].height=24
m.merge_cells("A2:F2");st(m["A2"],sz=9,fill=LBLUE)
m["A2"]="Оценки направлений не суммируются: рекомендация определяется ПАРОЙ уровней. Строка — уровень актуальности (I_А), столбец — уровень целесообразности (I_Ц)."
m["A3"]="Уровень актуальности:";st(m["A3"],b=True,border=False)
m["B3"]="=Оценка!G18";st(m["B3"],b=True,sz=12,ha="center",fill="FFFDE7")
m["C3"]="I_А =";st(m["C3"],b=True,ha="right",border=False)
m["D3"]="=Оценка!E18";st(m["D3"],b=True,ha="center",fill="FFFDE7")
m["A4"]="Уровень целесообразности:";st(m["A4"],b=True,border=False)
m["B4"]="=Оценка!G34";st(m["B4"],b=True,sz=12,ha="center",fill="FFFDE7")
m["C4"]="I_Ц =";st(m["C4"],b=True,ha="right",border=False)
m["D4"]="=Оценка!E34";st(m["D4"],b=True,ha="center",fill="FFFDE7")
m["E3"]="Стоп-факторов:";st(m["E3"],b=True,ha="right",border=False)
m["F3"]="=Оценка!D43";st(m["F3"],b=True,sz=12,ha="center",fill=LRED)
m["E4"]="«Красных» флагов:";st(m["E4"],b=True,ha="right",border=False)
m["F4"]="=Оценка!D57";st(m["F4"],b=True,ha="center",fill="FFF3E0")
# матрица B6:F11? шапка строка6, строки 7-11
m["A6"]="А \\ Ц";st(m["A6"],b=True,color="FFFFFF",fill=BLUE,ha="center")
for j,lv in enumerate(LV):
    c=m.cell(row=6,column=2+j,value=lv);st(c,b=True,color="FFFFFF",fill=BLUE,ha="center")
for i,lv in enumerate(LV):
    c=m.cell(row=7+i,column=1,value=lv);st(c,b=True,color="FFFFFF",fill=BLUE)
    for j in range(5):
        c=m.cell(row=7+i,column=2+j,value=RECMX[i][j]);st(c,sz=8.5)
    m.row_dimensions[7+i].height=44
# подсветка активной ячейки
for i in range(5):
    for j in range(5):
        col=get_column_letter(2+j);row=7+i
        m.conditional_formatting.add(f"{col}{row}",
          FormulaRule(formula=[f'AND($B$3=$A{row},{col}$6=$B$4,$F$3=0)'],
            fill=PatternFill("solid",fgColor="BDD7EE"),font=Font(bold=True)))
r=13
m.merge_cells(f"A{r}:F{r}");m[f"A{r}"]="РЕКОМЕНДАЦИЯ (с учётом стоп-факторов):";st(m[f"A{r}"],b=True,sz=11,fill=LBLUE)
m.merge_cells(f"A{r+1}:F{r+2}")
m[f"A{r+1}"]=('=IF(Оценка!D43>0,"ВНЕОЧЕРЕДНОЕ РАССМОТРЕНИЕ РУКОВОДСТВОМ: сработал стоп-фактор — рекомендация матрицы заблокирована, вплоть до приостановки работ",'
 'IF(OR(Оценка!E18="—",Оценка!E34="—"),"Недостаточно данных: оцените оба направления на листе «Оценка»",'
 'INDEX($B$7:$F$11,MATCH($B$3,$A$7:$A$11,0),MATCH($B$4,$B$6:$F$6,0))))')
st(m[f"A{r+1}"],b=True,sz=12,fill="FFFDE7",ha="center")
m.row_dimensions[r+1].height=22;m.row_dimensions[r+2].height=22
r2=r+4
m.merge_cells(f"A{r2}:F{r2}");m[f"A{r2}"]="ПЕРСПЕКТИВЫ ЗАВЕРШЕНИЯ НИОКР (позиция 4 заключения):";st(m[f"A{r2}"],b=True,fill=LBLUE)
m.merge_cells(f"A{r2+1}:F{r2+1}")
m[f"A{r2+1}"]=('=IF(Оценка!D43>0,"Маловероятная",IF(Оценка!E34="—","—",'
 'IF(Оценка!E34>=70,"Реалистичная",IF(Оценка!E34>=50,"Условно реалистичная","Маловероятная"))))')
st(m[f"A{r2+1}"],b=True,sz=12,ha="center",fill="FFFDE7")
r3=r2+3
m.merge_cells(f"A{r3}:F{r3}");m[f"A{r3}"]="СРАБОТАВШИЕ СТОП-ФАКТОРЫ:";st(m[f"A{r3}"],b=True,fill=LRED)
for i in range(5):
    rr=r3+1+i
    m.merge_cells(f"A{rr}:F{rr}")
    m[f"A{rr}"]=(f'=IF(Оценка!D{38+i}="","{i+1}. {STOPS[i]} — не заполнено",'
                 f'IF(Оценка!D{38+i}="Да","{i+1}. СРАБОТАЛ: {STOPS[i]}","{i+1}. Не сработал: {STOPS[i]}"))')
    st(m[f"A{rr}"],sz=9,fill=LRED if i%2==0 else "FDE9E7")
r4=r3+7
m.merge_cells(f"A{r4}:F{r4+1}");st(m[f"A{r4}"],sz=9,fill=LGRAY)
m[f"A{r4}"]=("Правило матрицы: В/В — продолжить, возможно ускорение; В/Н — пересмотр решений, исполнителей, ресурсов и сроков; Н/В — изменить назначение или требования к результату; "
 "Н/Н — прекращение НИОКР; С/С — продолжить при условии устранения замечаний. «Красные» флаги (лист «Оценка») не меняют рекомендацию автоматически, но обязательны к разбору комиссией.")
m.freeze_panes="A5"
m.page_setup.orientation="landscape";m.page_setup.fitToWidth=1
m.print_area=f"A1:F{r4+1}"

# ================= ЛИСТ 3: ЗАКЛЮЧЕНИЕ =================
z=wb.create_sheet("Заключение")
for k,v in {"A":4,"B":26,"C":96}.items():z.column_dimensions[k].width=v
z.merge_cells("A1:C1");st(z["A1"],b=True,sz=13,color="FFFFFF",fill=NAVY,ha="center")
z["A1"]="ЭКСПЕРТНОЕ ЗАКЛЮЧЕНИЕ — 5 ОБЯЗАТЕЛЬНЫХ ПОЗИЦИЙ"
z.row_dimensions[1].height=24
z.merge_cells("A2:C2");st(z["A2"],sz=9,fill=LBLUE)
z["A2"]="Позиции 1, 2, 4, 5 формируются автоматически из листов «Оценка» и «Матрица решений». Позиция 3 заполняется экспертом вручную. Готовый текст для протокола/Word — в нижней ячейке (копируется одним блоком)."
z["A3"]="НИОКР:";st(z["A3"],b=True,border=False)
z.merge_cells("B3:C3");z["B3"]="=Оценка!B3";st(z["B3"],b=True,fill=LGRAY)
rows=[
 ("1. Уровень актуальности",'=IF(Оценка!E18="—","не оценён (недостаточно данных)",TEXT(Оценка!E18,"0.0")&" % — «"&Оценка!G18&"» ("&Оценка!D18&" критериев)")'),
 ("2. Уровень целесообразности",'=IF(Оценка!E34="—","не оценён (недостаточно данных)",TEXT(Оценка!E34,"0.0")&" % — «"&Оценка!G34&"» ("&Оценка!D34&" критериев)")'),
 ("3. Проблемы с представленными материалами",""),
 ("4. Перспективы завершения","='Матрица решений'!A18"),
 ("5. Рекомендация","='Матрица решений'!A14"),
]
r=5
for name,formula in rows:
    z[f"B{r}"]=name;st(z[f"B{r}"],b=True,fill=LBLUE)
    z[f"C{r}"]=formula if formula else None
    st(z[f"C{r}"],fill="FFFDE7" if not formula else LGRAY,sz=10)
    z.row_dimensions[r].height=30
    r+=1
z["C7"]="(заполняется экспертом: полнота и достоверность материалов, доказательные пробелы, расхождения данных подразделения и аудита)"
st(z["C7"],fill="FFFDE7",sz=10);z["C7"]=None
z["C7"].comment=None
z["B7"]="3. Проблемы с представленными материалами";st(z["B7"],b=True,fill=LBLUE)
z["C7"]=""
from openpyxl.comments import Comment
z["C7"].comment=Comment("Пример: материалы представлены не полностью — отсутствует расчёт экономического эффекта и план импортозамещения комплектующих; расхождений данных не выявлено.","Модель аудита НИОКР")
r+=1
z.merge_cells(f"A{r}:C{r}");z[f"A{r}"]="СТОП-ФАКТОРЫ И «КРАСНЫЕ» ФЛАГИ (приложения к заключению)";st(z[f"A{r}"],b=True,fill=LRED,color="FFFFFF")
r+=1
z.merge_cells(f"B{r}:C{r}")
z[f"B{r}"]='=IF(Оценка!D43>0,"СТОП-ФАКТОРЫ: сработало "&Оценка!D43&" из 5 — рекомендация заблокирована, внеочередное рассмотрение руководством.","Стоп-факторы не сработали (0 из 5).")'
st(z[f"B{r}"],b=True,fill=LRED)
r+=1
z.merge_cells(f"B{r}:C{r}")
z[f"B{r}"]='=IF(Оценка!D43=0,"Сработало «красных» флагов (детекторы «зомби-проекта»): "&Оценка!D57&" из 10 — разобрать комиссией.","«Красных» флагов: "&Оценка!D57&" из 10 — рассматриваются вместе со стоп-факторами.")'
st(z[f"B{r}"],fill="FFF3E0")
r+=2
z.merge_cells(f"A{r}:C{r}");z[f"A{r}"]="ГОТОВЫЙ ТЕКСТ ЗАКЛЮЧЕНИЯ (скопировать в Word / протокол комиссии)";st(z[f"A{r}"],b=True,fill=NAVY,color="FFFFFF")
r+=1
z.merge_cells(f"A{r}:C{r+9}")
z[f"A{r}"]=('="ЭКСПЕРТНОЕ ЗАКЛЮЧЕНИЕ"&CHAR(10)&"по НИОКР: "&Оценка!B3&"  (шифр: "&Оценка!F3&"; эксперт: "&Оценка!B4&"; дата: "&Оценка!H3&")"&CHAR(10)&CHAR(10)'
 '&"1. Уровень актуальности: "&C5&"."&CHAR(10)'
 '&"2. Уровень целесообразности: "&C6&"."&CHAR(10)'
 '&"3. Проблемы с представленными материалами: "&IF(C7="","(не заполнено — позиция обязательна)",C7)&"."&CHAR(10)'
 '&"4. Перспективы завершения: "&C8&"."&CHAR(10)'
 '&"5. Рекомендация: "&C9&"."&CHAR(10)&CHAR(10)'
 '&IF(Оценка!D43>0,"СТОП-ФАКТОРЫ: сработало "&Оценка!D43&" — внеочередное рассмотрение руководством."&CHAR(10),"")'
 '&IF(Оценка!D57>0,"«КРАСНЫЕ» ФЛАГИ («зомби-проект»): "&Оценка!D57&" из 10."&CHAR(10),"")'
 '&CHAR(10)&"Подпись эксперта: ________________  /  "&Оценка!B4&" /"')
st(z[f"A{r}"],sz=10.5,fill="FFFFFF",va="top")
for rr in range(r,r+10):z.row_dimensions[rr].height=18
z.freeze_panes="A4"
z.page_setup.orientation="portrait";z.page_setup.fitToWidth=1
z.print_area=f"A1:C{r+9}"

# ================= ЛИСТ 4: СРАВНЕНИЕ НИОКР (этап 16) =================
from openpyxl.formatting.rule import ColorScaleRule
q=wb.create_sheet("Сравнение НИОКР")
for k,v in {"A":4,"B":30,"C":9,"D":9,"E":12,"F":9,"G":17,"H":17,"I":52,"J":17,"K":15}.items():q.column_dimensions[k].width=v
q.merge_cells("A1:K1");st(q["A1"],b=True,sz=13,color="FFFFFF",fill=NAVY,ha="center");q["A1"]="СРАВНЕНИЕ НИОКР: ЭКСПЕРТНЫЕ ИНДЕКСЫ И РЕКОМЕНДАЦИИ МАТРИЦЫ РЕШЕНИЙ"
q.row_dimensions[1].height=24
q.merge_cells("A2:K2");st(q["A2"],sz=9,fill=LBLUE)
q["A2"]=("Портфельный обзор до 10 разработок. Введите наименование, индексы I_А и I_Ц (0–100 % — из листа «Оценка» каждой НИОКР: E18/E34, или из модели v16.0), "
 "признак стоп-фактора («Да», если сработал хотя бы один из 5) и число «красных» флагов (0–10). Уровни, рекомендации матрицы, перспективы и группы решений рассчитываются автоматически. "
 "Оценки А и Ц НЕ складываются: рекомендацию определяет пара уровней; стоп-фактор блокирует рекомендацию.")
q.row_dimensions[2].height=44
HDR=["№","НИОКР (шифр / наименование)","I_А, %","I_Ц, %","Стоп-фактор","«Красных» флагов","Уровень А","Уровень Ц","Рекомендация матрицы решений","Перспективы завершения","Группа решения"]
for j,hname in enumerate(HDR):
    c=q.cell(row=4,column=1+j,value=hname);st(c,b=True,sz=9.5,color="FFFFFF",fill=BLUE,ha="center")
q.row_dimensions[4].height=30
dv_yn2=DataValidation(type="list",formula1=YESNO,allow_blank=True,showDropDown=False)
dv_pct=DataValidation(type="decimal",operator="between",formula1=0,formula2=100,allow_blank=True,
                      errorTitle="Индекс 0–100",error="Введите значение индекса от 0 до 100 (%)",showErrorMessage=True)
dv_fl=DataValidation(type="whole",operator="between",formula1=0,formula2=10,allow_blank=True,
                     errorTitle="0–10",error="Число «красных» флагов: от 0 до 10",showErrorMessage=True)
q.add_data_validation(dv_yn2);q.add_data_validation(dv_pct);q.add_data_validation(dv_fl)
R0,R1=5,14
for i in range(R0,R1+1):
    q[f"A{i}"]=i-R0+1;st(q[f"A{i}"],ha="center",sz=9,color="8C8C8C")
    st(q[f"B{i}"],fill="FFFDE7");st(q[f"C{i}"],fill="FFFDE7",ha="center");st(q[f"D{i}"],fill="FFFDE7",ha="center")
    st(q[f"E{i}"],fill="FFFDE7",ha="center");st(q[f"F{i}"],fill="FFFDE7",ha="center")
    q[f"G{i}"]=(f'=IF($B{i}="","",IF($C{i}="","—",IF($C{i}>=85,"Высокая",IF($C{i}>=70,"Хорошая",'
                f'IF($C{i}>=50,"Удовлетворительная",IF($C{i}>=30,"Низкая","Критическая"))))))');st(q[f"G{i}"],sz=9,ha="center",fill=LGRAY)
    q[f"H{i}"]=(f'=IF($B{i}="","",IF($D{i}="","—",IF($D{i}>=85,"Высокая",IF($D{i}>=70,"Хорошая",'
                f'IF($D{i}>=50,"Удовлетворительная",IF($D{i}>=30,"Низкая","Критическая"))))))');st(q[f"H{i}"],sz=9,ha="center",fill=LGRAY)
    q[f"I{i}"]=(f'=IF($B{i}="","",IF($E{i}="Да","ВНЕОЧЕРЕДНОЕ РАССМОТРЕНИЕ РУКОВОДСТВОМ: сработал стоп-фактор — рекомендация матрицы заблокирована, вплоть до приостановки работ",'
                f'IF(OR($C{i}="",$D{i}="",$G{i}="—",$H{i}="—"),"Недостаточно данных для экспертной оценки (введите I_А и I_Ц)",'
                f'INDEX($B$26:$F$30,MATCH($G{i},$A$26:$A$30,0),MATCH($H{i},$B$25:$F$25,0)))))');st(q[f"I{i}"],sz=9,fill=LGRAY)
    q[f"J{i}"]=(f'=IF($B{i}="","",IF($E{i}="Да","Маловероятная",IF(OR($D{i}="",$H{i}="—"),"—",'
                f'IF($D{i}>=70,"Реалистичная",IF($D{i}>=50,"Условно реалистичная","Маловероятная")))))');st(q[f"J{i}"],sz=9,ha="center",fill=LGRAY)
    q[f"K{i}"]=(f'=IF($B{i}="","",IF($E{i}="Да","СТОП-ФАКТОР",IF(OR($C{i}="",$D{i}=""),"Нет данных",'
                f'IF(OR(ISNUMBER(SEARCH("прекрат",$I{i})),ISNUMBER(SEARCH("прекращ",$I{i}))),"Прекращение",'
                f'IF(ISNUMBER(SEARCH("приостанов",$I{i})),"Приостановка",'
                f'IF(OR(ISNUMBER(SEARCH("пересмотр",$I{i})),ISNUMBER(SEARCH("проверить",$I{i}))),"Пересмотр",'
                f'IF(OR(ISNUMBER(SEARCH("корректировкой",$I{i})),ISNUMBER(SEARCH("устранения",$I{i})),ISNUMBER(SEARCH("уточнения",$I{i}))),"Корректировка","Продолжение")))))))');st(q[f"K{i}"],sz=9,b=True,ha="center",fill="FFFFFF")
    q.row_dimensions[i].height=30
    dv_yn2.add(q[f"E{i}"]);dv_pct.add(q[f"C{i}"]);dv_pct.add(q[f"D{i}"]);dv_fl.add(q[f"F{i}"])
# цветовые шкалы индексов и подсветка групп/уровней
q.conditional_formatting.add(f"C{R0}:D{R1}",ColorScaleRule(start_type="num",start_value=0,start_color="F8696B",
    mid_type="num",mid_value=60,mid_color="FFEB84",end_type="num",end_value=100,end_color="63BE7B"))
GROUP_FILL={"СТОП-ФАКТОР":"FFCDD2","Прекращение":"F8CBCB","Приостановка":"FFE0B2","Пересмотр":"FFF9C4","Корректировка":"EAF4EA","Продолжение":"C8E6C9","Нет данных":"EEEEEE"}
for gname,gcol in GROUP_FILL.items():
    q.conditional_formatting.add(f"K{R0}:K{R1}",FormulaRule(formula=[f'EXACT($K{R0},"{gname}")'],fill=PatternFill("solid",fgColor=gcol),font=Font(bold=True,size=9)))
LV_FILL={"Критическая":"FFCDD2","Низкая":"FFE0B2","Удовлетворительная":"FFF9C4","Хорошая":"EAF4EA","Высокая":"C8E6C9"}
for lname,lcol in LV_FILL.items():
    q.conditional_formatting.add(f"G{R0}:H{R1}",FormulaRule(formula=[f'$G{R0}="{lname}"'],fill=PatternFill("solid",fgColor=lcol)))
# сводка
r=16
q.merge_cells(f"A{r}:B{r}");q[f"A{r}"]="ЗАПОЛНЕНО НИОКР:";st(q[f"A{r}"],b=True,ha="right",fill=LBLUE)
q[f"C{r}"]=f'=COUNTIF($B${R0}:$B${R1},"<>")';st(q[f"C{r}"],b=True,sz=12,ha="center",fill="FFFDE7")
q.merge_cells(f"D{r}:E{r}");q[f"D{r}"]="Средний I_А / I_Ц:";st(q[f"D{r}"],b=True,ha="right",fill=LBLUE)
q[f"F{r}"]=f'=IFERROR(ROUND(AVERAGE($C${R0}:$C${R1}),1),"—")';st(q[f"F{r}"],b=True,ha="center",fill="FFFDE7")
q[f"G{r}"]=f'=IFERROR(ROUND(AVERAGE($D${R0}:$D${R1}),1),"—")';st(q[f"G{r}"],b=True,ha="center",fill="FFFDE7")
q.merge_cells(f"H{r}:K{r}");q[f"H{r}"]='=IF(COUNTIF($K$5:$K$14,"СТОП-ФАКТОР")>0,"⚠ Есть НИОКР со стоп-фактором — внеочередное рассмотрение руководством","Стоп-факторов по портфелю не выявлено")';st(q[f"H{r}"],b=True,fill=LRED)
r=17
q.merge_cells(f"A{r}:B{r}");q[f"A{r}"]="ГРУППЫ РЕШЕНИЙ:";st(q[f"A{r}"],b=True,ha="right",fill=LBLUE)
cc=3
for gname in ["Продолжение","Корректировка","Пересмотр","Приостановка","Прекращение","СТОП-ФАКТОР","Нет данных"]:
    c=q.cell(row=r,column=cc,value=f'="{gname}: "&COUNTIF($K${R0}:$K${R1},"{gname}")');st(c,sz=9,b=True,ha="center",fill=GROUP_FILL[gname]);cc+=1
q.merge_cells(f"J{r}:K{r}");q[f"J{r}"]='="«Красных» флагов всего: "&SUM($F$5:$F$14)';st(q[f"J{r}"],sz=9,b=True,ha="center",fill="FFF3E0")
# служебная матрица для INDEX/MATCH
r=24
q.merge_cells(f"A{r}:K{r}");st(q[f"A{r}"],sz=8,border=False,color="8C8C8C")
q[f"A{r}"]="Служебная таблица (не заполнять): матрица решений для автоподстановки рекомендаций. Строка — уровень А, столбец — уровень Ц."
q["A25"]="А \\ Ц";st(q["A25"],b=True,sz=8,color="FFFFFF",fill=BLUE,ha="center")
for j,lv in enumerate(LV):
    c=q.cell(row=25,column=2+j,value=lv);st(c,b=True,sz=8,color="FFFFFF",fill=BLUE,ha="center")
for i,lv in enumerate(LV):
    c=q.cell(row=26+i,column=1,value=lv);st(c,b=True,sz=8,color="FFFFFF",fill=BLUE)
    for j in range(5):
        c=q.cell(row=26+i,column=2+j,value=RECMX[i][j]);st(c,sz=7.5)
    q.row_dimensions[26+i].height=26
q.freeze_panes="A5"
q.page_setup.orientation="landscape";q.page_setup.fitToWidth=1;q.page_setup.fitToHeight=0
q.print_area=f"A1:K{r+5}"

wb.save("output/7_Оценка_актуальности_целесообразности_НИОКР.xlsx")
print("XLSX saved (4 листа)")
