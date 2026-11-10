# -*- coding: utf-8 -*-
"""Этап 18В: генерация Excel-шаблона расчёта уровня «Средний (30)» — переразметка файла 1 под v18.
Листы: Опросник_30 (30 кодов, ответы, примечания qFlags-подобные), Расчёт (I_А/I_Ц/I_Р, Coverage,
матрица 4×4, SF, зрелость, категория заключения), Справочники (шкала, уровни, M0–M9, R1–R7, SF2)."""
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
S = q18['levels']['S']['q']
LVL = [x[0] for x in d9['POROG_LVL']]  # Высокий..Критический

NAVY = "1F3864"; ACC = "2E74B5"; INp = "FFF2CC"
hdr_f = Font(bold=True, color="FFFFFF", size=10); hdr_fill = PatternFill("solid", fgColor=NAVY)
sec_fill = PatternFill("solid", fgColor="D9E2F3")
thin = Border(*[Side(style='thin', color='BBBBBB')] * 4)

wb = openpyxl.Workbook()

# ================= Опросник_30 =================
w1 = wb.active; w1.title = "Опросник_30"
w1['A1'] = "ОПРОСНИК АУДИТА НИОКР v18.0 — УРОВЕНЬ «СРЕДНИЙ» (30 кодов)"; w1['A1'].font = Font(bold=True, size=14, color=NAVY)
w1['A2'] = "Кодовая сетка А1–А10 · Б1–Б10 · В1–В10. Ответы: 5 4 3 2 1 0 · «НД» (нет данных) · «НП» (неприменимо, обязательно обоснование). Порог Coverage — 85 %."
w1['A3'] = "Ячейки ввода: G (Ответ), I (Обоснование), J (Доказательство). Расчётные колонки не редактировать."
w1['A3'].font = Font(italic=True, size=9)
heads = ["Код", "Направление", "Вопрос", "Что проверить", "Что считать доказательством", "Вес wᵢ",
         "ОТВЕТ", "Балл", "Обоснование (0–2, НП)", "Доказательство (3–5)", "Контрольное примечание",
         "Применим", "Балл×вес", "Вес учтённый"]
for j, t in enumerate(heads, 1):
    c = w1.cell(5, j, t); c.font = hdr_f; c.fill = hdr_fill; c.border = thin; c.alignment = Alignment(wrap_text=True, vertical='center')
R0 = 6
for i, x in enumerate(S):
    r = R0 + i
    w1.cell(r, 1, x['code']).font = Font(bold=True)
    w1.cell(r, 2, x['sec'])
    w1.cell(r, 3, x['q'])
    w1.cell(r, 4, x.get('check') or '')
    w1.cell(r, 5, x.get('ev') or '')
    w1.cell(r, 6, float(x['w']))
    g = w1.cell(r, 7); g.fill = PatternFill("solid", fgColor=INp); g.alignment = Alignment(horizontal='center')
    w1.cell(r, 8, '=IF(ISNUMBER(G{0}),G{0},"—")'.format(r))
    w1.cell(r, 9).fill = PatternFill("solid", fgColor=INp)
    w1.cell(r, 10).fill = PatternFill("solid", fgColor=INp)
    w1.cell(r, 11, ('=IF(G{0}="НП",IF(TRIM(I{0})="","НП: нет письменного обоснования; ",""),'
                    'IF(ISNUMBER(G{0}),IF(AND(G{0}<=2,TRIM(I{0})=""),"Балл ≤2 без обоснования; ","")'
                    '&IF(AND(G{0}>=4,TRIM(J{0})=""),"Балл 4–5 без доказательства; ",""),""))').format(r))
    w1.cell(r, 12, '=IF(G{0}="НП",0,1)'.format(r))
    w1.cell(r, 13, '=IF(ISNUMBER(G{0}),G{0}*F{0},0)'.format(r))
    w1.cell(r, 14, '=IF(ISNUMBER(G{0}),F{0},0)'.format(r))
    for j in range(1, 15):
        w1.cell(r, j).border = thin
        w1.cell(r, j).alignment = Alignment(wrap_text=True, vertical='top')
    if x['sec'] == 'А': w1.cell(r, 1).fill = PatternFill("solid", fgColor="E8F5E9")
    if x['sec'] == 'Б': w1.cell(r, 1).fill = PatternFill("solid", fgColor="FFF8E1")
    if x['sec'] == 'В': w1.cell(r, 1).fill = PatternFill("solid", fgColor="E3F2FD")
R1 = R0 + 29
dv = DataValidation(type="list", formula1='"5,4,3,2,1,0,НД,НП"', allow_blank=True, showErrorMessage=True,
                    errorTitle="Ответ v18", error="Допустимы 0–5, НД, НП")
w1.add_data_validation(dv); dv.add("G{0}:G{1}".format(R0, R1))
w1.conditional_formatting.add("G{0}:G{1}".format(R0, R1), CellIsRule(operator="greaterThanOrEqual", formula=["4"], fill=PatternFill("solid", fgColor="C8E6C9")))
w1.conditional_formatting.add("G{0}:G{1}".format(R0, R1), CellIsRule(operator="between", formula=["0", "2"], fill=PatternFill("solid", fgColor="FFCDD2")))
w1.conditional_formatting.add("G{0}:G{1}".format(R0, R1), CellIsRule(operator="equal", formula=["3"], fill=PatternFill("solid", fgColor="FFF9C4")))
w1.conditional_formatting.add("G{0}:G{1}".format(R0, R1), FormulaRule(formula=['G{0}="НП"'.format(R0)], fill=PatternFill("solid", fgColor="E0E0E0")))
w1.conditional_formatting.add("G{0}:G{1}".format(R0, R1), FormulaRule(formula=['G{0}="НД"'.format(R0)], fill=PatternFill("solid", fgColor="FFE0B2")))
widths = [6, 11, 46, 26, 30, 7, 8, 6, 26, 26, 30, 8, 8, 8]
for j, wd in enumerate(widths, 1):
    w1.column_dimensions[get_column_letter(j)].width = wd
w1.freeze_panes = "A6"

# ================= Расчёт =================
w2 = wb.create_sheet("Расчёт")
w2['A1'] = "РАСЧЁТ ИНДЕКСОВ v18.0 — УРОВЕНЬ «СРЕДНИЙ»"; w2['A1'].font = Font(bold=True, size=14, color=NAVY)
rows = [
 ("I_А — Актуальность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_30!M{0}:M{1},Опросник_30!B{0}:B{1},"А")/(5*SUMIFS(Опросник_30!N{0}:N{1},Опросник_30!B{0}:B{1},"А")),1),"н/д")'.format(R0, R1)),
 ("I_Ц — Целесообразность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_30!M{0}:M{1},Опросник_30!B{0}:B{1},"Б")/(5*SUMIFS(Опросник_30!N{0}:N{1},Опросник_30!B{0}:B{1},"Б")),1),"н/д")'.format(R0, R1)),
 ("I_Р — Результативность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_30!M{0}:M{1},Опросник_30!B{0}:B{1},"В")/(5*SUMIFS(Опросник_30!N{0}:N{1},Опросник_30!B{0}:B{1},"В")),1),"н/д")'.format(R0, R1)),
 ("Комплексный I = ∛(I_А·I_Ц·I_Р)", '=IFERROR(ROUND((B4*B5*B6)^(1/3),1),"н/д")'),
 ("Минимальный индекс направлений", '=IF(OR(ISTEXT(B4),ISTEXT(B5),ISTEXT(B6)),"н/д",ROUND(MIN(B4:B6),1))'),
 ("Coverage, % (порог 85)", '=IFERROR(ROUND(100*COUNT(Опросник_30!G{0}:G{1})/(30-COUNTIF(Опросник_30!G{0}:G{1},"НП")),1),"н/д")'.format(R0, R1)),
 ("Coverage пройден?", '=IF(ISNUMBER(B9),IF(B9>=85,"да","НЕТ — положительное заключение запрещено"),"н/д")'),
 ("Контрольных примечаний по вопросам", '=SUMPRODUCT((LEN(Опросник_30!K{0}:K{1})>0)*1)'.format(R0, R1)),
]
for i, (lab, f) in enumerate(rows):
    r = 4 + i * 1 if False else 3 + i
    w2.cell(r, 1, lab).font = Font(bold=True)
    c = w2.cell(r, 2, f); c.font = Font(bold=True, size=12); c.alignment = Alignment(horizontal='center')
    w2.cell(r, 1).border = thin; c.border = thin
# B4=I_A ... строки: 3..? пересчёт: r = 3+i → I_А в B4? r=3 → B3! Нужно B4.. — поправим ниже явными адресами.
wb.remove(w2)
w2 = wb.create_sheet("Расчёт")
w2['A1'] = "РАСЧЁТ ИНДЕКСОВ v18.0 — УРОВЕНЬ «СРЕДНИЙ»"; w2['A1'].font = Font(bold=True, size=14, color=NAVY)
SPEC = [
 ("B4", "I_А — Актуальность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_30!M{0}:M{1},Опросник_30!B{0}:B{1},"А")/(5*SUMIFS(Опросник_30!N{0}:N{1},Опросник_30!B{0}:B{1},"А")),1),"н/д")'.format(R0, R1)),
 ("B5", "I_Ц — Целесообразность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_30!M{0}:M{1},Опросник_30!B{0}:B{1},"Б")/(5*SUMIFS(Опросник_30!N{0}:N{1},Опросник_30!B{0}:B{1},"Б")),1),"н/д")'.format(R0, R1)),
 ("B6", "I_Р — Результативность, %", '=IFERROR(ROUND(100*SUMIFS(Опросник_30!M{0}:M{1},Опросник_30!B{0}:B{1},"В")/(5*SUMIFS(Опросник_30!N{0}:N{1},Опросник_30!B{0}:B{1},"В")),1),"н/д")'.format(R0, R1)),
 ("B7", "Комплексный I = ∛(I_А·I_Ц·I_Р)", '=IFERROR(ROUND((B4*B5*B6)^(1/3),1),"н/д")'),
 ("B8", "Минимальный индекс", '=IF(OR(ISTEXT(B4),ISTEXT(B5),ISTEXT(B6)),"н/д",ROUND(MIN(B4:B6),1))'),
 ("B9", "Coverage, % (порог 85)", '=IFERROR(ROUND(100*COUNT(Опросник_30!G{0}:G{1})/(30-COUNTIF(Опросник_30!G{0}:G{1},"НП")),1),"н/д")'.format(R0, R1)),
 ("B10", "Coverage пройден?", '=IF(ISNUMBER(B9),IF(B9>=85,"да","НЕТ — заключение блокируется"),"н/д")'),
 ("B11", "Контрольных примечаний", '=SUMPRODUCT((LEN(Опросник_30!K{0}:K{1})>0)*1)'.format(R0, R1)),
]
for addr, lab, f in SPEC:
    r = int(addr[1:])
    w2.cell(r, 1, lab).font = Font(bold=True); w2.cell(r, 1).border = thin
    c = w2[addr]; c.value = f; c.font = Font(bold=True, size=12); c.alignment = Alignment(horizontal='center'); c.border = thin
# уровни
w2['A13'] = "УРОВНИ НАПРАВЛЕНИЙ"; w2['A13'].font = Font(bold=True, color=NAVY)
LVLF = 'IF({0}="н/д","—",IF({0}>=85,"Высокий",IF({0}>=70,"Приемлемый",IF({0}>=55,"Проблемный","Критический"))))'
for addr, src, lab in [("B14", "B4", "I_А"), ("B15", "B5", "I_Ц"), ("B16", "B6", "I_Р")]:
    w2.cell(int(addr[1:]), 1, "Уровень «" + lab + "»").font = Font(bold=True)
    w2[addr] = "=" + LVLF.format(src)
# стоп-факторы
w2['A18'] = "СТОП-ФАКТОРЫ (ввод): SF-0 — фатальное; SF-1 — неопределённость; SF-2 — подтверждённый (блокирует)"; w2['A18'].font = Font(bold=True, color="C62828")
for i, lab in enumerate(["SF-0 (0/1 — отсутствует/подтверждён)", "SF-1", "SF-2"]):
    r = 19 + i
    w2.cell(r, 1, lab).font = Font(bold=True)
    c = w2.cell(r, 2, "0"); c.fill = PatternFill("solid", fgColor=INp); c.alignment = Alignment(horizontal='center')
dv2 = DataValidation(type="list", formula1='"0,1"', allow_blank=True); w2.add_data_validation(dv2); dv2.add("B19:B21")
w2['A23'] = "ЗРЕЛОСТЬ РЕЗУЛЬТАТА"; w2['A23'].font = Font(bold=True)
w2.cell(24, 1, "M текущий (M0…M9)").font = Font(bold=True); w2.cell(24, 2, "M0").fill = PatternFill("solid", fgColor=INp)
w2.cell(25, 1, "Подтверждение испытаниями (да/нет)").font = Font(bold=True); w2.cell(25, 2, "нет").fill = PatternFill("solid", fgColor=INp)
w2['B26'] = '=IF(AND(ISNUMBER(VALUE(SUBSTITUTE(B24,"M",""))),VALUE(SUBSTITUTE(B24,"M",""))>=5,B25<>"да"),"⚠ M≥5 без подтверждения испытаниями","")'
dv3 = DataValidation(type="list", formula1='"M0,M1,M2,M3,M4,M5,M6,M7,M8,M9"', allow_blank=True); w2.add_data_validation(dv3); dv3.add("B24")
# матрица решений (O..R / 10..13)
w2['N9'] = "МАТРИЦА РЕШЕНИЙ: строка — I_А, столбец — I_Ц"; w2['N9'].font = Font(bold=True, color=NAVY)
for j, lab in enumerate(LVL):
    c = w2.cell(10, 15 + j, lab); c.font = hdr_f; c.fill = PatternFill("solid", fgColor=ACC)  # O10..R10
for i, lab in enumerate(LVL):
    c = w2.cell(11 + i, 14, lab); c.font = hdr_f; c.fill = hdr_fill  # N11..N14
for i in range(4):
    for j in range(4):
        c = w2.cell(11 + i, 15 + j, RECMX4[i][j][0]); c.alignment = Alignment(wrap_text=True, vertical='top'); c.font = Font(size=8); c.border = thin
        w2.column_dimensions[get_column_letter(15 + j)].width = 26
w2.column_dimensions['N'].width = 14
w2['B28'] = '=IF(OR(B14="—",B15="—"),"н/д",INDEX(O11:R14,MATCH(B14,N11:N14,0),MATCH(B15,O10:R10,0)))'
w2['A28'] = "РЕКОМЕНДАЦИЯ (матрица)"; w2['A28'].font = Font(bold=True, size=11)
w2['B28'].font = Font(bold=True, size=11, color=NAVY); w2['B28'].alignment = Alignment(wrap_text=True)
w2.row_dimensions[28].height = 56
w2['A30'] = "БЛОКИРОВКА"; w2['A30'].font = Font(bold=True, color="C62828")
w2['B30'] = '=IF(B21=1,"СТОП: SF-2 подтверждён — безусловно положительное заключение запрещено; внеочередное рассмотрение руководством",IF(B19=1,"СТОП: SF-0 — аудирование прекращается, заключение не выносится","нет"))'
w2['B30'].alignment = Alignment(wrap_text=True); w2['B30'].font = Font(bold=True)
w2['A32'] = "КАТЕГОРИЯ ЗАКЛЮЧЕНИЯ"; w2['A32'].font = Font(bold=True, size=12)
w2['B32'] = ('=IF(B19=1,"IV. Отрицательное / эскалация (SF-0)",'
             'IF(B21=1,"IV. Отрицательное / эскалация (SF-2)",'
             'IF(OR(ISTEXT(B4),ISTEXT(B5),ISTEXT(B6),ISTEXT(B9)),"Недостаточно данных",'
             'IF(AND(B4>=85,B5>=85,B6>=85,B9>=85),"I. Положительное",'
             'IF(AND(B4>=70,B5>=70,B6>=70,B9>=85),"II. Положительное с замечаниями",'
             'IF(OR(B4<55,B5<55,B6<55,B9<85),"IV. Отрицательное / эскалация","III. Условно отрицательное"))))))')
w2['B32'].font = Font(bold=True, size=12); w2['B32'].alignment = Alignment(wrap_text=True)
w2.column_dimensions['A'].width = 34; w2.column_dimensions['B'].width = 44

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
block("Уровни индексов", ["Уровень", "Диапазон", "Действие"], [list(a) for a in d9['POROG_LVL']])
block("Зрелость M0–M9", ["Уровень", "Название", "Признак"], [list(a) for a in d9['M09']])
block("Решения R1–R7", ["Код", "Решение", "Условия"], [list(a) for a in q18['r17']])
block("Обстоятельства SF-2", ["№", "Формулировка"], [[i, a] for i, a in enumerate(d9['SF2'], 1)])
block("Категории заключения", ["Категория", "Условие", "Действие"], [list(a) for a in d9['CAT30']])
block("Пороги Coverage", ["Уровень опросника", "Порог"], [["Краткий (15)", 80], ["Средний (30)", 85], ["Полный (60)", 90]])
block("BCR", ["Диапазон", "Оценка"], [list(a) for a in d9['BCR']])
for j, wd in enumerate([12, 26, 70, 30], 1):
    w3.column_dimensions[get_column_letter(j)].width = wd

OUT = "output/1_Шаблон_расчёта_аудита_НИОКР_v18_Средний_30.xlsx"
wb.save(OUT)
import os
print("XLSX записан:", OUT, os.path.getsize(OUT), "B")
