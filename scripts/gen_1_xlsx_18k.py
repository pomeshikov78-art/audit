# -*- coding: utf-8 -*-
"""Этап 18Г: генерация Excel-шаблона расчёта уровня «Краткий (15)» — переразметка файла 1 под v18.
Листы: Опросник_15 (15 кодов с якорями 5/3/0, ответы, примечания), Расчёт (I_А/I_Ц/I_Р, комплексный,
Coverage порог 80, маршрутизация по полосам К, матрица 4×4, SF, зрелость, предварительная формулировка),
Справочники (шкала, уровни, полосы К, M0–M9, R1–R7, SF2, формулировки К, действия ≤2)."""
import io, json, re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

q18 = json.load(io.open('scripts/data/q18.json', encoding='utf-8'))
d9 = json.load(io.open('scripts/data/data9.json', encoding='utf-8'))
model = io.open('output/Аудит_НИОКР_интерактивная_модель.html', encoding='utf-8').read()
RECMX4 = json.loads(re.search(r'RECMX4=(\[.*?\]);\n', model, re.S).group(1))
KQ = q18['levels']['K']['q']
LVL = [x[0] for x in d9['POROG_LVL']]  # Высокий..Критический
CONCK = q18['conc']['K']

ACT = {
 'А1': 'Повторно подтвердить потребность у заказчика',
 'А2': 'Проверить соответствие современным требованиям к средствам разработки и производственным возможностям',
 'А3': 'Подтвердить сохранение актуальности задачи на период завершения НИОКР',
 'А4': 'Выполнить дополнительную проверку учёта новых технологий и альтернативных решений',
 'А5': 'Подтвердить уникальность результата или обосновать необходимость разработки',
 'Б6': 'Дополнить обоснование выбора варианта разработки сравнением альтернатив',
 'Б7': 'Пересмотреть и обосновать сроки выполнения',
 'Б8': 'Проверить соответствие выделенных ресурсов задачам',
 'Б9': 'Провести проверку обоснованности стоимости НИОКР',
 'Б10': 'Проверить техническую реализуемость концепции',
 'В11': 'Проверить полноту и подтверждённость результатов завершённого этапа',
 'В12': 'Провести контрольные испытания и подтверждение характеристик',
 'В13': 'Оценить готовность результатов к переходу на следующий этап',
 'В14': 'Проверить условия и возможность внедрения результата',
 'В15': 'Провести анализ отклонений фактических затрат и сроков от плановых',
}

NAVY = "1F3864"; ACC = "2E74B5"; INp = "FFF2CC"
hdr_f = Font(bold=True, color="FFFFFF", size=10); hdr_fill = PatternFill("solid", fgColor=NAVY)
thin = Border(*[Side(style='thin', color='BBBBBB')] * 4)

wb = openpyxl.Workbook()

# ================= Опросник_15 =================
w1 = wb.active; w1.title = "Опросник_15"
w1['A1'] = "ОПРОСНИК АУДИТА НИОКР v18.0 — УРОВЕНЬ «КРАТКИЙ» (15 кодов)"; w1['A1'].font = Font(bold=True, size=14, color=NAVY)
w1['A2'] = "Кодовая сетка А1–А5 · Б6–Б10 · В11–В15 (screening). Ответы: 5 4 3 2 1 0 · «НД» (нет данных) · «НП» (неприменимо, обязательно обоснование). Порог Coverage — 80 %. Итог — предварительный."
w1['A3'] = "Ячейки ввода: H (Ответ), J (Обоснование), K (Доказательство). Расчётные колонки не редактировать."
w1['A3'].font = Font(italic=True, size=9)
heads = ["Код", "Направление", "Вопрос", "Якорь 5", "Якорь 3", "Якорь 0", "Вес wᵢ",
         "ОТВЕТ", "Балл", "Обоснование (0–2, НП)", "Доказательство (3–5)", "Контрольное примечание",
         "Применим", "Балл×вес", "Вес учтённый"]
for j, t in enumerate(heads, 1):
    c = w1.cell(5, j, t); c.font = hdr_f; c.fill = hdr_fill; c.border = thin; c.alignment = Alignment(wrap_text=True, vertical='center')
R0 = 6
for i, x in enumerate(KQ):
    r = R0 + i
    w1.cell(r, 1, x['code']).font = Font(bold=True)
    w1.cell(r, 2, x['sec'])
    w1.cell(r, 3, x['q'])
    w1.cell(r, 4, x.get('a5') or '')
    w1.cell(r, 5, x.get('a3') or '')
    w1.cell(r, 6, x.get('a0') or '')
    w1.cell(r, 7, float(x['w']))
    g = w1.cell(r, 8); g.fill = PatternFill("solid", fgColor=INp); g.alignment = Alignment(horizontal='center')
    w1.cell(r, 9, '=IF(ISNUMBER(H{0}),H{0},"—")'.format(r))
    w1.cell(r, 10).fill = PatternFill("solid", fgColor=INp)
    w1.cell(r, 11).fill = PatternFill("solid", fgColor=INp)
    w1.cell(r, 12, ('=IF(H{0}="НП",IF(TRIM(J{0})="","НП: нет письменного обоснования; ",""),'
                    'IF(ISNUMBER(H{0}),IF(AND(H{0}<=2,TRIM(J{0})=""),"Балл ≤2 без обоснования; ","")'
                    '&IF(AND(H{0}>=4,TRIM(K{0})=""),"Балл 4–5 без доказательства; ",""),""))').format(r))
    w1.cell(r, 13, '=IF(H{0}="НП",0,1)'.format(r))
    w1.cell(r, 14, '=IF(ISNUMBER(H{0}),H{0}*G{0},0)'.format(r))
    w1.cell(r, 15, '=IF(ISNUMBER(H{0}),G{0},0)'.format(r))
    for j in range(1, 16):
        w1.cell(r, j).border = thin
        w1.cell(r, j).alignment = Alignment(wrap_text=True, vertical='top')
    if x['sec'] == 'А': w1.cell(r, 1).fill = PatternFill("solid", fgColor="E8F5E9")
    if x['sec'] == 'Б': w1.cell(r, 1).fill = PatternFill("solid", fgColor="FFF8E1")
    if x['sec'] == 'В': w1.cell(r, 1).fill = PatternFill("solid", fgColor="E3F2FD")
R1 = R0 + 14
dv = DataValidation(type="list", formula1='"5,4,3,2,1,0,НД,НП"', allow_blank=True, showErrorMessage=True,
                    errorTitle="Ответ v18", error="Допустимы 0–5, НД, НП")
w1.add_data_validation(dv); dv.add("H{0}:H{1}".format(R0, R1))
w1.conditional_formatting.add("H{0}:H{1}".format(R0, R1), CellIsRule(operator="greaterThanOrEqual", formula=["4"], fill=PatternFill("solid", fgColor="C8E6C9")))
w1.conditional_formatting.add("H{0}:H{1}".format(R0, R1), CellIsRule(operator="between", formula=["0", "2"], fill=PatternFill("solid", fgColor="FFCDD2")))
w1.conditional_formatting.add("H{0}:H{1}".format(R0, R1), CellIsRule(operator="equal", formula=["3"], fill=PatternFill("solid", fgColor="FFF9C4")))
w1.conditional_formatting.add("H{0}:H{1}".format(R0, R1), FormulaRule(formula=['H{0}="НП"'.format(R0)], fill=PatternFill("solid", fgColor="E0E0E0")))
w1.conditional_formatting.add("H{0}:H{1}".format(R0, R1), FormulaRule(formula=['H{0}="НД"'.format(R0)], fill=PatternFill("solid", fgColor="FFE0B2")))
widths = [6, 11, 42, 26, 26, 26, 7, 8, 6, 24, 24, 28, 8, 8, 8]
for j, wd in enumerate(widths, 1):
    w1.column_dimensions[get_column_letter(j)].width = wd
w1.freeze_panes = "A6"

# ================= Расчёт =================
w2 = wb.create_sheet("Расчёт")
w2['A1'] = "РАСЧЁТ ИНДЕКСОВ v18.0 — УРОВЕНЬ «КРАТКИЙ» (итог предварительный)"; w2['A1'].font = Font(bold=True, size=14, color=NAVY)
SPEC = [
 ("B4", "I_А — Актуальность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_15!N{0}:N{1},Опросник_15!B{0}:B{1},"А")/(5*SUMIFS(Опросник_15!O{0}:O{1},Опросник_15!B{0}:B{1},"А")),1),"н/д")'.format(R0, R1)),
 ("B5", "I_Ц — Целесообразность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_15!N{0}:N{1},Опросник_15!B{0}:B{1},"Б")/(5*SUMIFS(Опросник_15!O{0}:O{1},Опросник_15!B{0}:B{1},"Б")),1),"н/д")'.format(R0, R1)),
 ("B6", "I_Р — Результативность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_15!N{0}:N{1},Опросник_15!B{0}:B{1},"В")/(5*SUMIFS(Опросник_15!O{0}:O{1},Опросник_15!B{0}:B{1},"В")),1),"н/д")'.format(R0, R1)),
 ("B7", "Комплексный I = ∛(I_А·I_Ц·I_Р)", '=IFERROR(ROUND((B4*B5*B6)^(1/3),1),"н/д")'),
 ("B8", "Минимальный индекс (I_А, I_Ц)", '=IF(OR(ISTEXT(B4),ISTEXT(B5)),"н/д",ROUND(MIN(B4,B5),1))'),
 ("B9", "Coverage, % (порог 80)", '=IFERROR(ROUND(100*COUNT(Опросник_15!H{0}:H{1})/(15-COUNTIF(Опросник_15!H{0}:H{1},"НП")),1),"н/д")'.format(R0, R1)),
 ("B10", "Coverage пройден?", '=IF(ISNUMBER(B9),IF(B9>=80,"да","НЕТ — положительная формулировка блокируется"),"н/д")'),
 ("B11", "Контрольных примечаний", '=SUMPRODUCT((LEN(Опросник_15!L{0}:L{1})>0)*1)'.format(R0, R1)),
]
for addr, lab, f in SPEC:
    r = int(addr[1:])
    w2.cell(r, 1, lab).font = Font(bold=True); w2.cell(r, 1).border = thin
    c = w2[addr]; c.value = f; c.font = Font(bold=True, size=12); c.alignment = Alignment(horizontal='center'); c.border = thin
# уровни (общая шкала методики — для матрицы)
w2['A13'] = "УРОВНИ НАПРАВЛЕНИЙ (общая шкала методики; применяются в матрице решений)"; w2['A13'].font = Font(bold=True, color=NAVY)
LVLF = 'IF({0}="н/д","—",IF({0}>=85,"Высокий",IF({0}>=70,"Приемлемый",IF({0}>=55,"Проблемный","Критический"))))'
for addr, src, lab in [("B14", "B4", "I_А"), ("B15", "B5", "I_Ц"), ("B16", "B6", "I_Р")]:
    w2.cell(int(addr[1:]), 1, "Уровень «" + lab + "»").font = Font(bold=True)
    w2[addr] = "=" + LVLF.format(src)
w2.cell(17, 1, "Маршрутизация (полосы уровня «Краткий», по min(I_А,I_Ц))").font = Font(bold=True)
w2['B17'] = ('=IF(OR(ISTEXT(B4),ISTEXT(B5)),"н/д",IF(MIN(B4,B5)>=80,"Высокий — продолжить выполнение в рамках текущего контроля",'
             'IF(MIN(B4,B5)>=60,"Приемлемый — продолжить, определить замечания и сроки их устранения",'
             'IF(MIN(B4,B5)>=40,"Низкий — направить на средний опросник",'
             '"Критически низкий — провести внеплановое углублённое обследование"))))')
w2['B17'].alignment = Alignment(wrap_text=True); w2['B17'].font = Font(bold=True)
# стоп-факторы
w2['A19'] = "СТОП-ФАКТОРЫ (ввод): SF-0 — фатальное; SF-1 — неопределённость; SF-2 — подтверждённый (блокирует)"; w2['A19'].font = Font(bold=True, color="C62828")
for i, lab in enumerate(["SF-0 (0/1 — отсутствует/подтверждён)", "SF-1", "SF-2"]):
    r = 20 + i
    w2.cell(r, 1, lab).font = Font(bold=True)
    c = w2.cell(r, 2, "0"); c.fill = PatternFill("solid", fgColor=INp); c.alignment = Alignment(horizontal='center')
dv2 = DataValidation(type="list", formula1='"0,1"', allow_blank=True); w2.add_data_validation(dv2); dv2.add("B20:B22")
w2['A24'] = "ЗРЕЛОСТЬ РЕЗУЛЬТАТА"; w2['A24'].font = Font(bold=True)
w2.cell(25, 1, "M текущий (M0…M9)").font = Font(bold=True); w2.cell(25, 2, "M0").fill = PatternFill("solid", fgColor=INp)
w2.cell(26, 1, "Подтверждение испытаниями (да/нет)").font = Font(bold=True); w2.cell(26, 2, "нет").fill = PatternFill("solid", fgColor=INp)
w2['B27'] = '=IF(AND(ISNUMBER(VALUE(SUBSTITUTE(B25,"M",""))),VALUE(SUBSTITUTE(B25,"M",""))>=5,B26<>"да"),"⚠ M≥5 без подтверждения испытаниями","")'
dv3 = DataValidation(type="list", formula1='"M0,M1,M2,M3,M4,M5,M6,M7,M8,M9"', allow_blank=True); w2.add_data_validation(dv3); dv3.add("B25")
# матрица решений (O..R / 10..13 — те же адреса, что в шаблоне «Средний»)
w2['N9'] = "МАТРИЦА РЕШЕНИЙ: строка — I_А, столбец — I_Ц"; w2['N9'].font = Font(bold=True, color=NAVY)
for j, lab in enumerate(LVL):
    c = w2.cell(10, 15 + j, lab); c.font = hdr_f; c.fill = PatternFill("solid", fgColor=ACC)
for i, lab in enumerate(LVL):
    c = w2.cell(11 + i, 14, lab); c.font = hdr_f; c.fill = hdr_fill
for i in range(4):
    for j in range(4):
        c = w2.cell(11 + i, 15 + j, RECMX4[i][j][0]); c.alignment = Alignment(wrap_text=True, vertical='top'); c.font = Font(size=8); c.border = thin
        w2.column_dimensions[get_column_letter(15 + j)].width = 26
w2.column_dimensions['N'].width = 14
w2['B29'] = '=IF(OR(B14="—",B15="—"),"н/д",INDEX(O11:R14,MATCH(B14,N11:N14,0),MATCH(B15,O10:R10,0)))'
w2['A29'] = "РЕКОМЕНДАЦИЯ (матрица, предварительная)"; w2['A29'].font = Font(bold=True, size=11)
w2['B29'].font = Font(bold=True, size=11, color=NAVY); w2['B29'].alignment = Alignment(wrap_text=True)
w2.row_dimensions[29].height = 56
w2['A31'] = "БЛОКИРОВКА"; w2['A31'].font = Font(bold=True, color="C62828")
w2['B31'] = '=IF(B22=1,"СТОП: SF-2 подтверждён — безусловно положительная формулировка запрещена; внеочередное рассмотрение руководством",IF(B20=1,"СТОП: SF-0 — аудирование прекращается, заключение не выносится","нет"))'
w2['B31'].alignment = Alignment(wrap_text=True); w2['B31'].font = Font(bold=True)
w2['A33'] = "ФОРМУЛИРОВКА ЗАКЛЮЧЕНИЯ (предварительная)"; w2['A33'].font = Font(bold=True, size=12)
w2['B33'] = ('=IF(B20=1,"Эскалация: SF-0 — аудирование прекращено",'
             'IF(B22=1,"Эскалация: SF-2 подтверждён — внеочередное рассмотрение руководством",'
             'IF(OR(ISTEXT(B4),ISTEXT(B5)),"Недостаточно данных: требуется сбор подтверждающих документов",'
             'IF(MIN(B4,B5)>=80,"%s",'
             'IF(MIN(B4,B5)>=60,"%s",'
             'IF(MIN(B4,B5)>=40,"%s","%s"))))))' % (CONCK[0][1], CONCK[1][1], CONCK[2][1], CONCK[3][1]))
w2['B33'].font = Font(bold=True, size=12); w2['B33'].alignment = Alignment(wrap_text=True)
w2.column_dimensions['A'].width = 40; w2.column_dimensions['B'].width = 52

# ================= Справочники =================
w3 = wb.create_sheet("Справочники")
r = 1
def block(title, headers, rows):
    global r
    w3.cell(r, 1, title).font = Font(bold=True, size=12, color=NAVY); r += 1
    for j, t in enumerate(headers, 1):
        c = w3.cell(r, j, t); c.font = hdr_f; c.fill = hdr_fill
    r += 1
    for row in rows:
        for j, x in enumerate(row, 1):
            w3.cell(r, j, x)
        r += 1
    r += 1
block("Шкала ответов", ["Балл", "Название", "Содержание"], [list(a) for a in d9['SCALE']])
block("Уровни индексов (общие)", ["Уровень", "Диапазон", "Действие"], [list(a) for a in d9['POROG_LVL']])
block("Полосы обоснованности «Краткий»", ["Диапазон", "Уровень", "Действия комиссии"], [list(a) for a in q18['porog15']])
block("Зрелость M0–M9", ["Уровень", "Название", "Признак"], [list(a) for a in d9['M09']])
block("Решения R1–R7", ["Код", "Решение", "Условия"], [list(a) for a in q18['r17']])
block("Обстоятельства SF-2", ["№", "Формулировка"], [[i, a] for i, a in enumerate(d9['SF2'], 1)])
block("Формулировки заключения «Краткий»", ["Условие", "Формулировка"], [list(a) for a in CONCK])
block("Обязательные действия при балле ≤ 2", ["Код", "Действие"], [[x['code'], ACT[x['code']]] for x in KQ])
block("Пороги Coverage", ["Уровень опросника", "Порог"], [["Краткий (15)", 80], ["Средний (30)", 85], ["Полный (60)", 90]])
for j, wd in enumerate([12, 30, 70, 30], 1):
    w3.column_dimensions[get_column_letter(j)].width = wd

OUT = "output/1_Шаблон_расчёта_аудита_НИОКР_v18_Краткий_15.xlsx"
wb.save(OUT)
import os
print("XLSX записан:", OUT, os.path.getsize(OUT), "B")
