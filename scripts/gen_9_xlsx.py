# -*- coding: utf-8 -*-
# Этап 18А: генератор Excel-калькулятора аудита НИОКР v2
# output/9_Калькулятор_аудита_НИОКР_v2.xlsx — 14 листов по спецификации ЧАСТИ I «Научно.docx»:
#   Панель_руководителя, Паспорт_НИОКР, Краткий_15, Средний_30, Полный_60, Доказательства,
#   Эксперты, AHP, Стоп_факторы, Зрелость_M0_M9, Экономика_LCC, План_мероприятий, Заключение, Справочники.
# Данные вопросов/весов/порогов: scripts/data/data9.json (scripts/parse_data9.py из «Научно.docx»).
import io, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

D = json.load(io.open('scripts/data/data9.json', encoding='utf-8'))
Q15, Q30, Q60 = D['Q15'], D['Q30'], D['Q60']
BLOCKS60 = D['BLOCKS60']
ANCH30 = D['ANCH30']
AUTO_REC = D['AUTO_REC']
OUT = 'output/9_Калькулятор_аудита_НИОКР_v2.xlsx'
CELLS = 'scripts/data/cells9.json'

# ---------- стили ----------
F_TITLE = Font(name='Calibri', size=14, bold=True, color='1F4E79')
F_SUB   = Font(name='Calibri', size=10, italic=True, color='404040')
F_HDR   = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
F_BOLD  = Font(name='Calibri', size=11, bold=True)
F_SEC   = Font(name='Calibri', size=11, bold=True, color='1F4E79')
F_RED   = Font(name='Calibri', size=11, bold=True, color='C00000')
F_SMALL = Font(name='Calibri', size=9, color='595959')
FILL_HDR  = PatternFill('solid', fgColor='1F4E79')
FILL_A    = PatternFill('solid', fgColor='DDEBF7')   # актуальность
FILL_B    = PatternFill('solid', fgColor='E2EFDA')   # целесообразность
FILL_V    = PatternFill('solid', fgColor='FFF2CC')   # результативность
FILL_FORM = PatternFill('solid', fgColor='F2F2F2')   # формулы
FILL_SUM  = PatternFill('solid', fgColor='D9E1F2')   # сводные блоки
FILL_WARN = PatternFill('solid', fgColor='FCE4E4')
FILL_OK   = PatternFill('solid', fgColor='E7F4E4')
THIN = Side(style='thin', color='BFBFBF')
BORD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CTR  = Alignment(horizontal='center', vertical='center', wrap_text=True)

def title(ws, text, sub=None, sub2=None):
    ws['A1'] = text; ws['A1'].font = F_TITLE
    if sub:  ws['A2'] = sub;  ws['A2'].font = F_SUB
    if sub2: ws['A3'] = sub2; ws['A3'].font = F_SUB

def header_row(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col+i, value=h)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = CTR; c.border = BORD

def sec(ws, row, text, col=1):
    c = ws.cell(row=row, column=col, value=text); c.font = F_SEC

def put(ws, coord, val, font=None, fill=None, fmt=None, align=None, border=True):
    c = ws[coord]; c.value = val
    if font: c.font = font
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    c.alignment = align or WRAP
    if border: c.border = BORD
    return c

def dv(ws, formula1, ranges, allow_blank=True):
    d = DataValidation(type='list', formula1=formula1, allow_blank=allow_blank, showDropDown=False)
    ws.add_data_validation(d)
    for r in ranges: d.add(r)

def widths(ws, wd):
    for col, w in wd.items(): ws.column_dimensions[col].width = w

FMT_PCT = '0.0"%"'; FMT_W = '0.00'; FMT_SC = '0.00'; FMT_CR = '0.000'; FMT_MONEY = '# ##0.00'

wb = Workbook(); wb.remove(wb.active)
wb.properties.title = 'Калькулятор аудита НИОКР v2 — актуальность, целесообразность, результативность'
wb.properties.creator = 'Аудит НИОКР · этап 18А (по методике «Научно-обоснованная практическая методика оценки НИОКР») '
wb.properties.description = 'Трёхуровневая система опросников 15/30/60; шкала 0–5 + НД + НП; Coverage 80/85/90%; AHP с расчётом CR; SF-0/1/2; зрелость M0–M9; LCC/NPV/BCR; R1–R7.'

# ==================================================================
# ЛИСТ 2 (создаём раньше панели — на неё ссылаемся): Паспорт_НИОКР
# ==================================================================
ws = wb.create_sheet('Паспорт_НИОКР'); ws.sheet_properties.tabColor = '1F4E79'
title(ws, 'ПАСПОРТ НИОКР — исходные данные',
      'Заполняется до начала аудита. «Уровень аудита» определяет, какой опросник и заключение используются на панели и в заключении.',
      'Решение о прекращении НИОКР принимается с учётом условий госконтракта, прав на РИД, обязательств перед заказчиком и правовых последствий (не только по индексам).')
PASS_FIELDS = [
    ('Наименование НИОКР', '', None),
    ('Шифр', '', None),
    ('Государственный контракт / договор', '', None),
    ('Государственный заказчик / заказчик', '', None),
    ('Головной исполнитель', '', None),
    ('Основные организации кооперации', '', None),
    ('Дата начала', '', 'DD.MM.YYYY'),
    ('Дата окончания (план)', '', 'DD.MM.YYYY'),
    ('Стоимость НИОКР, млн ₽', '', FMT_MONEY),
    ('Текущая стадия жизненного цикла', 'ОКР',
     '"фундаментальные исследования,прикладные исследования (НИР),ОКР,опытный образец,предварительные испытания,приёмочные испытания,установочная партия,серийное производство,завершена"'),
    ('Планируемая зрелость результата M (0–9)', 5, '"0,1,2,3,4,5,6,7,8,9"'),
    ('Уровень аудита', 'Средний', '"Краткий,Средний,Полный"'),
    ('Дата проведения аудита', '', 'DD.MM.YYYY'),
    ('Председатель комиссии', '', None),
    ('Секретарь комиссии', '', None),
]
header_row(ws, 5, ['', 'Поле', 'Значение'])
for i, (lab, val, dvl) in enumerate(PASS_FIELDS):
    r = 6 + i
    put(ws, f'B{r}', lab, font=F_BOLD, fill=FILL_FORM)
    c = put(ws, f'C{r}', val); c.fill = PatternFill('solid', fgColor='FFFFFF')
    if i == 8: c.number_format = FMT_MONEY
    if dvl: dv(ws, dvl, [f'C{r}'])
widths(ws, {'A': 2, 'B': 40, 'C': 52})
sec(ws, 15, 'Примечания')
NOTES_PASS = [
    'Смена стадии НИОКР → проверьте применимость критериев зрелости (лист «Зрелость_M0_M9») и «НП»-ответов.',
    'Уровни аудита: Краткий (15 вопросов, экспресс-скрининг, Coverage ≥ 80%) · Средний (30, углублённый, ≥ 85%) · Полный (60, комплексный, ≥ 90%).',
    'Руководитель НИОКР предоставляет материалы и пояснения, но не определяет итоговые оценки самостоятельно.',
]
for i, t in enumerate(NOTES_PASS): ws[f'B{16+i}'] = '• ' + t; ws[f'B{16+i}'].font = F_SMALL
PASS = {'name': 'C6', 'code': 'C7', 'contract': 'C8', 'customer': 'C9', 'exec': 'C10', 'coop': 'C11',
        'start': 'C12', 'end': 'C13', 'cost': 'C14', 'stage': 'C15', 'mplan': 'C16', 'level': 'C17',
        'adate': 'C18', 'chair': 'C19', 'secr': 'C20'}

# ==================================================================
# ЛИСТ 7: Эксперты
# ==================================================================
ws = wb.create_sheet('Эксперты'); ws.sheet_properties.tabColor = '7030A0'
title(ws, 'ЭКСПЕРТЫ — состав комиссии и коэффициенты компетентности',
      'Комиссия 7–9 человек: технический, экономический, производственный и профильный эксперты. Коэффициент компетентности rᵢ задаётся ДО оценивания (0,1–1,0).',
      'Колонки Э1–Э3 в опросниках соответствуют строкам 1–3 этого листа; итоговый балл = взвешенное среднее xⱼ* = Σrᵢxᵢⱼ / Σrᵢ (по числовым оценкам).')
header_row(ws, 4, ['№', 'Ф.И.О.', 'Роль в комиссии', 'Коэффициент компетентности rᵢ', 'Декларация независимости (дата)', 'Конфликт интересов', 'Примечание'])
for i in range(9):
    r = 5 + i
    put(ws, f'A{r}', i+1, align=CTR)
    for col in 'BCDEFG':
        c = ws[f'{col}{r}']; c.border = BORD; c.alignment = WRAP
    ws[f'C{r}'] = D['ROLES'][i][0] if i < len(D['ROLES']) else ''
    ws[f'D{r}'] = 1.0; ws[f'D{r}'].number_format = '0.0'
    ws[f'F{r}'] = 'Нет'
dv(ws, '"Нет,Да"', ['F5:F13'])
d = DataValidation(type='decimal', operator='between', formula1='0.1', formula2='1', allow_blank=True)
ws.add_data_validation(d); d.add('D5:D13')
put(ws, 'B15', '=IF(COUNTIF(F5:F13,"Да")>0,"⚠ Заявлен конфликт интересов — требуется отвод эксперта от соответствующей части проверки либо замена (решение председателя)","✔ Конфликты интересов не заявлены")', font=F_BOLD, fill=FILL_FORM)
ws.merge_cells('B15:G15')
sec(ws, 17, 'Принципы деятельности комиссии')
header_row(ws, 18, ['', 'Принцип', 'Содержание'])
for i, (a, b) in enumerate(D['PRINC']):
    r = 19 + i
    put(ws, f'B{r}', a, font=F_BOLD, fill=FILL_FORM)
    put(ws, f'C{r}', b); ws.merge_cells(f'C{r}:G{r}')
sec(ws, 28, 'Полномочия и ответственность ролей')
header_row(ws, 29, ['', 'Участник', 'Основные полномочия', 'Ответственность'])
for i, (a, b, c_) in enumerate(D['ROLES']):
    r = 30 + i
    put(ws, f'B{r}', a, font=F_BOLD, fill=FILL_FORM); put(ws, f'C{r}', b); put(ws, f'D{r}', c_)
    ws.merge_cells(f'D{r}:G{r}')
sec(ws, 40, 'Независимость: эксперт не оценивает собственную работу; самооценка руководителя НИОКР не подменяет независимую оценку комиссии; изменение экспертной оценки по административным мотивам не допускается — только при новых доказательствах, ошибке расчёта или повторном рассмотрении с фиксацией причины.')
ws['A40'].alignment = WRAP; ws['A40'].font = F_SMALL; ws.merge_cells('A40:G42')
widths(ws, {'A': 4, 'B': 30, 'C': 30, 'D': 12, 'E': 16, 'F': 12, 'G': 24})
EXP_R = ('Эксперты!$D$5', 'Эксперты!$D$6', 'Эксперты!$D$7')

# ==================================================================
# Общая функция листов опросников
# ==================================================================
def sheet_summary_formula(ws, col_R, col_S, col_T, r1, r2, out):
    return None

HDR_Q = ['№/код', 'Блок', 'Вопрос', 'Что проверяем / установить', 'Основные доказательства', 'Якоря / примечание', 'Вес',
         'Э1', 'Э2', 'Э3', 'Разброс', 'Флаг', 'ИТОГ', 'Обоснование / комментарий (обязательно для 0–2 и НП)',
         'Доказательство (ID, обязательно для 3–5)', 'Критичность', 'Контроль',
         'w·числ', 'w·x', 'w·прим', 'Изменение (кто/дата)', 'флаг⚑', '⚑0–2', '⚑контр']

def build_questions(ws, rows_data, first_row, fill_map):
    """rows_data: list of (code, block, question, colD, colE, colF, weight)"""
    r0 = first_row
    for i, (code, block, q, cd, ce, cf, wq) in enumerate(rows_data):
        r = r0 + i
        fill = fill_map.get(block[0] if block else '', None)
        put(ws, f'A{r}', code, align=CTR)
        put(ws, f'B{r}', block, align=CTR)
        put(ws, f'C{r}', q)
        put(ws, f'D{r}', cd, font=F_SMALL)
        put(ws, f'E{r}', ce, font=F_SMALL)
        put(ws, f'F{r}', cf, font=F_SMALL)
        put(ws, f'G{r}', wq, fmt=FMT_W, align=CTR)
        for col in 'HIJ':
            c = ws[f'{col}{r}']; c.border = BORD; c.alignment = CTR
        put(ws, f'K{r}', f'=IF(COUNT(H{r}:J{r})>=2,MAX(H{r}:J{r})-MIN(H{r}:J{r}),"")', fill=FILL_FORM, fmt=FMT_SC, align=CTR)
        put(ws, f'L{r}', (f'=IF(COUNT(H{r}:J{r})<2,"",IF(K{r}>=3,"⚠ разброс ≥3 — коллегиальное рассмотрение",'
                          f'IF(K{r}>=2,"обсуждение: разброс ≥2",'
                          f'IF(AND(COUNT(H{r}:J{r})>0,COUNTIF(H{r}:J{r},"НД")>0),"обсуждение: НД вместе с баллами",'
                          f'IF(AND(MIN(H{r}:J{r})<=1,MAX(H{r}:J{r})>=4),"обсуждение: 0–1 против 4–5","")))))'), fill=FILL_FORM, font=F_SMALL)
        m = (f'=IF(COUNT(H{r}:J{r})>0,IF(COUNT(H{r}:J{r})>=COUNTIF(H{r}:J{r},"НП"),'
             f'IFERROR(ROUND((IF(ISNUMBER(H{r}),H{r}*{EXP_R[0]},0)+IF(ISNUMBER(I{r}),I{r}*{EXP_R[1]},0)+IF(ISNUMBER(J{r}),J{r}*{EXP_R[2]},0))'
             f'/(IF(ISNUMBER(H{r}),{EXP_R[0]},0)+IF(ISNUMBER(I{r}),{EXP_R[1]},0)+IF(ISNUMBER(J{r}),{EXP_R[2]},0)),2),"НД"),"НП"),'
             f'IF(COUNTIF(H{r}:J{r},"НП")>0,"НП",IF(COUNTIF(H{r}:J{r},"НД")>0,"НД","")))')
        c = put(ws, f'M{r}', m, fill=FILL_FORM, font=F_BOLD, fmt=FMT_SC, align=CTR)
        for col in 'NO':
            cc = ws[f'{col}{r}']; cc.border = BORD; cc.alignment = WRAP
        cc = ws[f'P{r}']; cc.border = BORD; cc.alignment = WRAP
        put(ws, f'Q{r}', (f'=IF(AND(ISNUMBER(M{r}),M{r}<=2,LEN(N{r})=0),"⚠ балл 0–2: нет обоснования; ","")'
                          f'&IF(AND(ISNUMBER(M{r}),M{r}>=3,LEN(O{r})=0),"⚠ балл 3–5: нет доказательства; ","")'
                          f'&IF(AND(M{r}="НП",LEN(N{r})=0),"⚠ НП: нет письменного обоснования; ","")'), fill=FILL_FORM, font=F_SMALL)
        put(ws, f'R{r}', f'=IF(ISNUMBER(M{r}),G{r},0)', fill=FILL_FORM, fmt=FMT_W)
        put(ws, f'S{r}', f'=IF(ISNUMBER(M{r}),G{r}*M{r},0)', fill=FILL_FORM, fmt=FMT_W)
        put(ws, f'T{r}', f'=IF(M{r}="НП",0,G{r})', fill=FILL_FORM, fmt=FMT_W)
        cc = ws[f'U{r}']; cc.border = BORD; cc.alignment = WRAP
        put(ws, f'V{r}', f'=IF(L{r}="",0,1)', fill=FILL_FORM)
        put(ws, f'W{r}', f'=IF(AND(ISNUMBER(M{r}),M{r}<=2),1,0)', fill=FILL_FORM)
        put(ws, f'X{r}', f'=IF(LEN(Q{r})>0,1,0)', fill=FILL_FORM)
        if fill:
            for col in ('A', 'B', 'C'): ws[f'{col}{r}'].fill = fill
    last = r0 + len(rows_data) - 1
    dv(ws, '"0,1,2,3,4,5,НД,НП"', [f'H{r0}:J{last}'])
    dv(ws, '"обычное замечание,существенное замечание,критический недостаток"', [f'P{r0}:P{last}'])
    for col in ('R', 'S', 'T', 'V', 'W', 'X'):
        ws.column_dimensions[col].hidden = True
    rng = f'M{r0}:M{last}'
    ws.conditional_formatting.add(rng, CellIsRule(operator='lessThanOrEqual', formula=['2'], fill=PatternFill('solid', fgColor='F8CBAD')))
    ws.conditional_formatting.add(rng, CellIsRule(operator='greaterThanOrEqual', formula=['4'], fill=PatternFill('solid', fgColor='C6EFCE')))
    ws.conditional_formatting.add(f'Q{r0}:Q{last}', FormulaRule(formula=[f'LEN($Q{r0})>0'], font=Font(color='C00000', bold=True)))
    ws.conditional_formatting.add(f'L{r0}:L{last}', FormulaRule(formula=[f'LEN($L{r0})>0'], font=Font(color='BF6900', bold=True)))
    return last

def idx_formula(ws, S_range, R_range):
    return f'=IF(SUM({R_range})=0,"н/д",ROUND(SUM({S_range})/(5*SUM({R_range}))*100,1))'

def cov_formula(R_range, T_range):
    return f'=IF(SUM({T_range})=0,"н/д",ROUND(SUM({R_range})/SUM({T_range})*100,1))'

QW = {'A': 6, 'B': 6, 'C': 52, 'D': 30, 'E': 30, 'F': 34, 'G': 6.5, 'H': 5.5, 'I': 5.5, 'J': 5.5,
      'K': 8, 'L': 24, 'M': 7, 'N': 30, 'O': 16, 'P': 16, 'Q': 28, 'U': 16}
SUMFILL = FILL_SUM
RULES_TXT = 'Правила: 4–5 баллов — только при документальном подтверждении (колонка «Доказательство»); 0–2 — обязательное обоснование; НД снижает полноту доказательств (Coverage); НП — только с письменным обоснованием (исключается из расчёта и из знаменателя Coverage).'
SCALE_TXT = 'Шкала: 5 полностью соответствует · 4 в основном · 3 частично · 2 существенно не соответствует · 1 практически не соответствует · 0 не соответствует · НД нет данных · НП неприменимо.'

def summary_block(ws, r, labels_vals, fmts=None):
    for i, (lab, val, fmt) in enumerate(labels_vals):
        rr = r + i
        put(ws, f'B{rr}', lab, font=F_BOLD, fill=SUMFILL)
        c = put(ws, f'C{rr}', val, fill=FILL_FORM, font=F_BOLD)
        if fmt: c.number_format = fmt
        ws.merge_cells(f'D{rr}:H{rr}')
    return r + len(labels_vals) - 1

# ==================================================================
# ЛИСТ 3: Краткий_15
# ==================================================================
ws = wb.create_sheet('Краткий_15'); ws.sheet_properties.tabColor = '9DC3E6'
title(ws, 'ПРИЛОЖЕНИЕ № 1. КРАТКИЙ ОПРОСНИК — экспресс-оценка (15 вопросов: 5 А + 5 Б + 5 В)',
      'Назначение: первичный отбор НИОКР, выявление очевидных проблем, определение необходимости углублённой проверки. 30–60 минут; проводит руководитель или аудитор. Результат — НЕ решение о прекращении, а решение о необходимости углублённой проверки.',
      SCALE_TXT + ' ' + RULES_TXT)
header_row(ws, 5, HDR_Q)
rows = []
for q in Q15:
    blk = {'А': 'А', 'Б': 'Б', 'В': 'В'}[q['sec']]
    rows.append((q['n'], blk, q['q'], q['a5'], q['a3'], q['a0'], 1))
# подставим якоря в D/E/F: D=5 баллов, E=3 балла, F=0 баллов
last = build_questions(ws, rows, 6, {'А': FILL_A, 'Б': FILL_B, 'В': FILL_V})
# переименуем заголовки D/E/F для краткого
ws['D5'] = 'Якорь «5 баллов»'; ws['E5'] = 'Якорь «3 балла»'; ws['F5'] = 'Якорь «0 баллов»'
rA1, rA2 = 6, 10; rB1, rB2 = 11, 15; rC1, rC2 = 16, 20
sec(ws, 22, 'СВОДКА — КРАТКИЙ ОПРОСНИК (I = Σx / (5·n) × 100%, веса равные; Coverage = Σw оценённых / Σw применимых × 100%)')
SUM15 = [
    ('Индекс актуальности I_А, %',  idx_formula(ws, f'$S${rA1}:$S${rA2}', f'$R${rA1}:$R${rA2}'), FMT_PCT),
    ('Индекс целесообразности I_Ц, %', idx_formula(ws, f'$S${rB1}:$S${rB2}', f'$R${rB1}:$R${rB2}'), FMT_PCT),
    ('Индекс результативности I_Р, %', idx_formula(ws, f'$S${rC1}:$S${rC2}', f'$R${rC1}:$R${rC2}'), FMT_PCT),
    ('Комплексный индекс (геом. среднее, справочно)', '=IFERROR(ROUND(POWER(C23*C24*C25,1/3),1),"н/д")', FMT_PCT),
    ('Полнота доказательств Coverage, %', cov_formula(f'$R${rA1}:$R${rC2}', f'$T${rA1}:$T${rC2}'), FMT_PCT),
    ('Порог Coverage (краткий уровень)', 80, '0'),
    ('Проверка полноты доказательств', '=IF(C27="н/д","✘ нет данных",IF(C27<C28,"✘ ниже порога — положительное заключение не выносится","✔ порог пройден"))', None),
    ('Вопросов «к обсуждению» (флаги)', '=SUM($V$6:$V$20)', '0'),
    ('Вопросов с контрольными замечаниями', '=SUM($X$6:$X$20)', '0'),
    ('Минимальный индекс', '=IF(OR(C23="н/д",C24="н/д",C25="н/д"),"н/д",MIN(C23,C24,C25))', FMT_PCT),
    ('Вопросов с баллом ≤ 2', '=SUM($W$6:$W$20)', '0'),
]
summary_block(ws, 23, SUM15)
put(ws, 'B35', 'ЗАКЛЮЧЕНИЕ КРАТКОГО ОПРОСНИКА (авто)', font=F_RED, fill=SUMFILL)
put(ws, 'C35', ('=IF($C$32="н/д","Недостаточно данных для достоверной оценки",'
                'IF(Стоп_факторы!$C$20>0,"Выявлен SF-2: безусловно положительное заключение невозможно; эскалация руководству",'
                'IF($C$27<$C$28,"Оценка предварительная. Требуется получение доказательств (Coverage < 80%)",'
                'IF($C$32>=80,"Предварительно обоснованная НИОКР. Оснований для внепланового углублённого аудита не выявлено",'
                'IF($C$32>=60,"НИОКР имеет основания для продолжения, необходим мониторинг выявленных недостатков",'
                'IF($C$32>=40,"Выявлены признаки недостаточной актуальности или целесообразности. Рекомендован средний опросник",'
                '"Выявлены существенные сомнения в обоснованности НИОКР. Необходим внеплановый углублённый аудит"))))))'),
    fill=FILL_FORM, font=F_BOLD)
ws.merge_cells('C35:N35'); ws['C35'].alignment = WRAP
RULES15 = [
    ('=IF(AND(ISNUMBER(C23),C23<40),"⚠ I_А < 40% — НИОКР обязательно направляется на углублённую оценку","")', ),
    ('=IF(AND(ISNUMBER(C24),C24<40),"⚠ I_Ц < 40% — проверяется возможность альтернативного способа решения задачи","")', ),
    ('=IF(AND(ISNUMBER(C32),C32<60),"⚠ Индекс < 60% — проект не может получить безусловное заключение «положительно»","")', ),
]
for i, (f,) in enumerate(RULES15):
    put(ws, f'B{36+i}', f, fill=FILL_FORM, font=F_RED)
    ws.merge_cells(f'B{36+i}:N{36+i}')
widths(ws, QW); ws.freeze_panes = 'A6'
K15C = {'IA': 'C23', 'IC': 'C24', 'IR': 'C25', 'KOMPL': 'C26', 'COV': 'C27', 'THR': 'C28', 'OK': 'C29',
        'FLAGS': 'C30', 'CTRL': 'C31', 'MIN': 'C32', 'LOW2': 'C33', 'CONC': 'C35'}

# ==================================================================
# ЛИСТ 4: Средний_30
# ==================================================================
ws = wb.create_sheet('Средний_30'); ws.sheet_properties.tabColor = 'A9D08E'
title(ws, 'ПРИЛОЖЕНИЕ № 2. СРЕДНИЙ ОПРОСНИК — углублённая оценка (30 вопросов: А1–А10, Б1–Б10, В1–В10)',
      'Основной инструмент планового аудита портфеля. 2–4 часа, аудиторская группа. Веса разделов равны (по 10 условных единиц); I = Σw·x / (5·Σw) × 100%. Результат — оценка состояния + план корректирующих действий.',
      SCALE_TXT + ' ' + RULES_TXT)
header_row(ws, 5, HDR_Q)
rows = []
for q in Q30:
    blk = q['code'][0]
    anchor = ''
    if q['code'] in ANCH30:
        anchor = ' · '.join(f'{k} — {ANCH30[q["code"]][k]}' for k in ['5', '4', '3', '2', '1', '0'])
    rows.append((q['code'], blk, q['q'], q['check'], q['ev'], anchor, q['w']))
last = build_questions(ws, rows, 6, {'А': FILL_A, 'Б': FILL_B, 'В': FILL_V})
rA1, rA2 = 6, 15; rB1, rB2 = 16, 25; rC1, rC2 = 26, 35
sec(ws, 37, 'СВОДКА — СРЕДНИЙ ОПРОСНИК (знаменатель 50 = 5 баллов × сумма весов 10)')
SUM30 = [
    ('Индекс актуальности I_А, %',  idx_formula(ws, f'$S${rA1}:$S${rA2}', f'$R${rA1}:$R${rA2}'), FMT_PCT),
    ('Индекс целесообразности I_Ц, %', idx_formula(ws, f'$S${rB1}:$S${rB2}', f'$R${rB1}:$R${rB2}'), FMT_PCT),
    ('Индекс результативности I_Р, %', idx_formula(ws, f'$S${rC1}:$S${rC2}', f'$R${rC1}:$R${rC2}'), FMT_PCT),
    ('Комплексный индекс (геом. среднее, справочно)', '=IFERROR(ROUND(POWER(C38*C39*C40,1/3),1),"н/д")', FMT_PCT),
    ('Полнота доказательств Coverage, %', cov_formula(f'$R${rA1}:$R${rC2}', f'$T${rA1}:$T${rC2}'), FMT_PCT),
    ('Порог Coverage (средний уровень)', 85, '0'),
    ('Проверка полноты доказательств', '=IF(C42="н/д","✘ нет данных",IF(C42<C43,"✘ ниже порога — положительное заключение не выносится","✔ порог пройден"))', None),
    ('Вопросов «к обсуждению» (флаги)', '=SUM($V$6:$V$35)', '0'),
    ('Вопросов с контрольными замечаниями', '=SUM($X$6:$X$35)', '0'),
    ('Минимальный индекс', '=IF(OR(C38="н/д",C39="н/д",C40="н/д"),"н/д",MIN(C38,C39,C40))', FMT_PCT),
    ('Вопросов с баллом ≤ 2', '=SUM($W$6:$W$35)', '0'),
    ('Coverage раздела А (актуальность), %', cov_formula(f'$R${rA1}:$R${rA2}', f'$T${rA1}:$T${rA2}'), FMT_PCT),
    ('Coverage раздела Б (целесообразность), %', cov_formula(f'$R${rB1}:$R${rB2}', f'$T${rB1}:$T${rB2}'), FMT_PCT),
    ('Coverage раздела В (результативность), %', cov_formula(f'$R${rC1}:$R${rC2}', f'$T${rC1}:$T${rC2}'), FMT_PCT),
    ('Сумма весов А (контроль = 10)', f'=SUM($G${rA1}:$G${rA2})', FMT_W),
    ('Сумма весов Б (контроль = 10)', f'=SUM($G${rB1}:$G${rB2})', FMT_W),
    ('Сумма весов В (контроль = 10)', f'=SUM($G${rC1}:$G${rC2})', FMT_W),
]
summary_block(ws, 38, SUM30)
put(ws, 'B56', 'ЗАКЛЮЧЕНИЕ СРЕДНЕГО ОПРОСНИКА (авто; категории I–IV)', font=F_RED, fill=SUMFILL)
put(ws, 'C56', ('=IF($C$47="н/д","Недостаточно данных для достоверной оценки",'
                'IF(Стоп_факторы!$C$20>0,"IV. Отрицательное / эскалация — выявлен SF-2: передать на рассмотрение комиссии, оценить изменение концепции и альтернативы",'
                'IF($C$42<$C$43,"Недостаточно данных для достоверной оценки (Coverage ниже порога 85%)",'
                'IF($C$47<55,"IV. Отрицательное / эскалация — передать на рассмотрение комиссии, оценить изменение концепции и альтернативы",'
                'IF($C$47>=85,"I. Положительное — продолжить НИОКР",'
                'IF($C$47>=70,"II. Положительное с замечаниями — продолжить с корректирующими мероприятиями",'
                '"III. Условно отрицательное — утвердить план восстановления, провести повторную оценку"))))))'),
    fill=FILL_FORM, font=F_BOLD)
ws.merge_cells('C56:N56'); ws['C56'].alignment = WRAP
RULES30 = [
    '=IF(AND(ISNUMBER(C38),C38<55),"⚠ I_А < 55% — обязательна проверка потребности у заказчика","")',
    '=IF(AND(ISNUMBER(C39),C39<55),"⚠ I_Ц < 55% — обязателен анализ альтернативных решений","")',
    '=IF(AND(ISNUMBER(C40),C40<55),"⚠ I_Р < 55% — план восстановления результативности, повторная проверка","")',
    '=IF(Стоп_факторы!$C$19>0,"⚠ SF-1: при наличии критического технического риска составляется отдельное заключение независимо от итогового индекса","")',
    '=IF(AND(ISNUMBER(C47),C47<70),"⚠ Хотя бы один индекс < 70% — положительное заключение без замечаний не выносится","")',
]
for i, f in enumerate(RULES30):
    put(ws, f'B{57+i}', f, fill=FILL_FORM, font=F_RED)
    ws.merge_cells(f'B{57+i}:N{57+i}')
widths(ws, QW); ws.freeze_panes = 'A6'
C30C = {'IA': 'C38', 'IC': 'C39', 'IR': 'C40', 'KOMPL': 'C41', 'COV': 'C42', 'THR': 'C43', 'OK': 'C44',
        'FLAGS': 'C45', 'CTRL': 'C46', 'MIN': 'C47', 'LOW2': 'C48', 'CONC': 'C56'}
ROW30 = {}
for i, q in enumerate(Q30): ROW30[q['code']] = 6 + i

# ==================================================================
# ЛИСТ 5: Полный_60
# ==================================================================
ws = wb.create_sheet('Полный_60'); ws.sheet_properties.tabColor = 'FFD966'
title(ws, 'ПРИЛОЖЕНИЕ № 3. ПОЛНЫЙ ОПРОСНИК — комплексный научно-технический аудит (60 вопросов: 20 А + 20 Б + 20 В)',
      '1–3 рабочих дня, междисциплинарная экспертная группа (7–9 человек). Двухуровневые веса: блоки внутри направления (базовые или AHP) → направления. Для полного аудита эксперт указывает конкретное доказательство: дату, номер, место хранения либо идентификатор в ИС.',
      SCALE_TXT + ' ' + RULES_TXT)
header_row(ws, 5, HDR_Q)
ws['F5'] = 'Примечание / особое правило'
SPECIAL60 = {
    13: 'Вопросы 13–15: сравнивать функциональные задачи, архитектуру, техрешения, состав изделия и результаты — различие названий не доказывает отсутствие дублирования.',
    42: 'В2: нельзя оценивать только по отчёту разработчика — необходимы подтверждённые испытания/измерения.',
    49: 'В7: формальный отчёт об устранении недостаточен — проверка фактического закрытия замечаний.',
    50: 'В8: особенно важен, когда результат — готовность к следующему этапу освоения производства.',
    60: 'В10: учитывать прирост характеристик, импортозамещение, надёжность, эксплуатационные расходы, задел.',
}
rows = []
for q in Q60:
    blk = None
    for b, a, z, wgt in BLOCKS60:
        if a <= q['n'] <= z: blk = b
    rows.append((q['n'], blk, q['q'], q['est'], q['ev'], SPECIAL60.get(q['n'], ''), 1))
last = build_questions(ws, rows, 6, {'А': FILL_A, 'Б': FILL_B, 'В': FILL_V})
r1, r2 = 6, 65
sec(ws, 67, 'БЛОКИ И ИНДЕКСЫ — ПОЛНЫЙ ОПРОСНИК (Dk = Σq·x / (5·Σq) × 100%; направления — по весам блоков с ренормировкой при «н/д»)')
header_row(ws, 68, ['Блок', 'Содержание', 'Базовый вес', 'Вес AHP', 'Применённый вес', 'Dk, %', 'w эфф.', 'вклад'])
BNAME = {'А1': 'Обоснованность исходной потребности', 'А2': 'Перспективность и прогноз', 'А3': 'Научно-техническая новизна',
         'А4': 'Дублирование и стратегическая значимость', 'Б1': 'Обоснованность технической концепции', 'Б2': 'Экономическая обоснованность',
         'Б3': 'Ресурсы и реализуемость', 'Б4': 'Управление, риски и кооперация', 'В1': 'Выполнение ТЗ и достижение результатов',
         'В2': 'Документация и испытания', 'В3': 'Готовность к производству и внедрению', 'В4': 'Эффект и управление результатами'}
AHP_W = {'А1': 'AHP!$F$18', 'А2': 'AHP!$F$19', 'А3': 'AHP!$F$20', 'А4': 'AHP!$F$21',
         'Б1': 'AHP!$F$30', 'Б2': 'AHP!$F$31', 'Б3': 'AHP!$F$32', 'Б4': 'AHP!$F$33',
         'В1': 'AHP!$F$42', 'В2': 'AHP!$F$43', 'В3': 'AHP!$F$44', 'В4': 'AHP!$F$45'}
for i, (b, a, z, wgt) in enumerate(BLOCKS60):
    r = 69 + i
    put(ws, f'A{r}', b, font=F_BOLD, align=CTR)
    put(ws, f'B{r}', BNAME[b], font=F_SMALL)
    put(ws, f'C{r}', wgt, fmt=FMT_W, align=CTR)
    put(ws, f'D{r}', f'={AHP_W[b]}', fmt='0.0000', align=CTR, fill=FILL_FORM)
    put(ws, f'E{r}', f'=IF($C$82="Да",$D{r},$C{r})', fmt=FMT_W, align=CTR, fill=FILL_FORM)
    put(ws, f'F{r}', (f'=IF(SUMIF($B${r1}:$B${r2},$A{r},$R${r1}:$R${r2})=0,"н/д",'
                      f'ROUND(SUMIF($B${r1}:$B${r2},$A{r},$S${r1}:$S${r2})/(5*SUMIF($B${r1}:$B${r2},$A{r},$R${r1}:$R${r2}))*100,1))'),
        fmt=FMT_PCT, fill=FILL_FORM, font=F_BOLD)
    put(ws, f'G{r}', f'=IF(ISNUMBER($F{r}),$E{r},0)', fmt=FMT_W, fill=FILL_FORM)
    put(ws, f'H{r}', f'=IF(ISNUMBER($F{r}),$E{r}*$F{r},0)', fmt=FMT_W, fill=FILL_FORM)
    fill = {'А': FILL_A, 'Б': FILL_B, 'В': FILL_V}[b[0]]
    ws[f'A{r}'].fill = fill; ws[f'B{r}'].fill = fill
put(ws, 'B82', 'Использовать веса блоков из AHP (лист AHP, матрицы 2–4):', font=F_BOLD, fill=SUMFILL)
put(ws, 'C82', 'Нет', align=CTR); dv(ws, '"Да,Нет"', ['C82'])
SUM60 = [
    ('Индекс актуальности I_А, %', '=IF(SUM($G$69:$G$72)=0,"н/д",ROUND(SUM($H$69:$H$72)/SUM($G$69:$G$72),1))', FMT_PCT),
    ('Индекс целесообразности I_Ц, %', '=IF(SUM($G$73:$G$76)=0,"н/д",ROUND(SUM($H$73:$H$76)/SUM($G$73:$G$76),1))', FMT_PCT),
    ('Индекс результативности I_Р, %', '=IF(SUM($G$77:$G$80)=0,"н/д",ROUND(SUM($H$77:$H$80)/SUM($G$77:$G$80),1))', FMT_PCT),
    ('Комплексный индекс (геом. среднее, справочно)', '=IFERROR(ROUND(POWER(C84*C85*C86,1/3),1),"н/д")', FMT_PCT),
    ('Полнота доказательств Coverage, %', cov_formula(f'$R${r1}:$R${r2}', f'$T${r1}:$T${r2}'), FMT_PCT),
    ('Порог Coverage (полный уровень)', 90, '0'),
    ('Проверка полноты доказательств', '=IF(C88="н/д","✘ нет данных",IF(C88<C89,"✘ ниже порога — положительное заключение не выносится","✔ порог пройден"))', None),
    ('Вопросов «к обсуждению» (флаги)', f'=SUM($V${r1}:$V${r2})', '0'),
    ('Вопросов с контрольными замечаниями', f'=SUM($X${r1}:$X${r2})', '0'),
    ('Минимальный индекс', '=IF(OR(C84="н/д",C85="н/д",C86="н/д"),"н/д",MIN(C84,C85,C86))', FMT_PCT),
    ('Вопросов с баллом ≤ 2', f'=SUM($W${r1}:$W${r2})', '0'),
    ('Coverage направления А (вопросы 1–20), %', cov_formula('$R$6:$R$25', '$T$6:$T$25'), FMT_PCT),
    ('Coverage направления Б (вопросы 21–40), %', cov_formula('$R$26:$R$45', '$T$26:$T$45'), FMT_PCT),
    ('Coverage направления В (вопросы 41–60), %', cov_formula('$R$46:$R$65', '$T$46:$T$65'), FMT_PCT),
    ('Контроль весов: сумма А1–А4 = 1', '=SUM($C$69:$C$72)', FMT_W),
    ('Контроль весов: сумма Б1–Б4 = 1', '=SUM($C$73:$C$76)', FMT_W),
    ('Контроль весов: сумма В1–В4 = 1', '=SUM($C$77:$C$80)', FMT_W),
]
summary_block(ws, 84, SUM60)
put(ws, 'B102', 'ИТОГОВОЕ РЕШЕНИЕ ПОЛНОГО АУДИТА (R1–R7, авто)', font=F_RED, fill=SUMFILL)
put(ws, 'C102', ('=IF($C$93="н/д","Недостаточно данных для достоверной оценки",'
                 'IF(Стоп_факторы!$C$20>0,"R7 — передать на рассмотрение комиссии по прекращению или изменению проекта (выявлен SF-2)",'
                 'IF($C$88<$C$89,"Недостаточно данных для достоверной оценки (Coverage < 90%)",'
                 'IF(AND($C$84<55,$C$85<55,$C$86<55),"R7 — передать на рассмотрение комиссии по прекращению или изменению проекта (все три индекса < 55%)",'
                 'IF(AND($C$84>=85,$C$85>=85,$C$86>=85),"R1 — продолжить без существенных изменений",'
                 'IF($C$93>=70,"R2 — продолжить с корректировкой",'
                 'IF(Экономика_LCC!$C$48<>"Продолжить разработку","R6 — рассмотреть альтернативное решение (по сравнению альтернатив предпочтительнее: "&Экономика_LCC!$C$48&")",'
                 'IF(AND(ISNUMBER($F$75),$F$75<70),"R5 — пересмотреть сроки и ресурсы (блок Б3 «Ресурсы и реализуемость» < 70%)",'
                 'IF(AND($C$84<70,$C$85>=70),"R4 — изменить состав и объём работ (актуальность снижена при обоснованной целесообразности)",'
                 'IF(OR($C$85<70,$C$86<70),"R3 — изменить техническую концепцию (целесообразность или результативность < 70%)",'
                 '"R2 — продолжить с корректировкой"))))))))))'),
    fill=FILL_FORM, font=F_BOLD)
ws.merge_cells('C102:N102'); ws['C102'].alignment = WRAP
put(ws, 'B103', '=IF(AND(ISNUMBER(C93),C93<70),"⚠ Хотя бы один индекс < 70% — положительное заключение без замечаний не выносится, даже если комплексный индекс высокий","")', fill=FILL_FORM, font=F_RED)
ws.merge_cells('B103:N103')
put(ws, 'B104', 'R7 — рекомендация о рассмотрении вопроса комиссией, а НЕ самостоятельное решение о расторжении контракта или прекращении обязательств.', font=F_SMALL, fill=FILL_FORM)
ws.merge_cells('B104:N104')
widths(ws, QW); ws.column_dimensions['B'].width = 26; ws.freeze_panes = 'A6'
C60C = {'IA': 'C84', 'IC': 'C85', 'IR': 'C86', 'KOMPL': 'C87', 'COV': 'C88', 'THR': 'C89', 'OK': 'C90',
        'FLAGS': 'C91', 'CTRL': 'C92', 'MIN': 'C93', 'LOW2': 'C94', 'CONC': 'C102', 'DKB3': 'F75', 'TOGGLE': 'C82'}
BLK_ROW = {b: 69 + i for i, (b, a, z, wgt) in enumerate(BLOCKS60)}

# ==================================================================
# ЛИСТ 6: Доказательства
# ==================================================================
ws = wb.create_sheet('Доказательства'); ws.sheet_properties.tabColor = 'BDD7EE'
title(ws, 'ДОКАЗАТЕЛЬСТВА — реестр документов, протоколов, расчётов и иных подтверждений',
      'Каждая оценка 3–5 ссылается на доказательство (ID из этого реестра); для полного аудита указываются дата, номер и место хранения либо идентификатор в ИС.',
      'Наличие числовой оценки само по себе ещё не означает качественное доказательство: поля «оценка» и «доказательство» разделены.')
header_row(ws, 4, ['ID', 'Тип', 'Наименование', 'Дата', 'Номер', 'Место хранения / идентификатор в ИС', 'Связанный вопрос (уровень/код)', 'Кем проверено', 'Результат проверки', 'Комментарий', 'заполнено'])
for i in range(30):
    r = 5 + i
    put(ws, f'A{r}', f'Д-{i+1:03d}', align=CTR, font=F_SMALL)
    for col in 'BCDEFGHIJ':
        c = ws[f'{col}{r}']; c.border = BORD; c.alignment = WRAP
    put(ws, f'K{r}', f'=IF(LEN(C{r})>0,1,0)', fill=FILL_FORM)
    ws.column_dimensions['K'].hidden = True
dv(ws, '"документ,протокол,расчёт,испытание,отчёт,интервью,иное"', ['B5:B34'])
dv(ws, '"подтверждено,не подтверждено,ожидает проверки"', ['I5:I34'])
sec(ws, 36, 'Сводка')
put(ws, 'B37', 'Всего записей', font=F_BOLD, fill=SUMFILL); put(ws, 'C37', '=SUM(K5:K34)', fill=FILL_FORM, fmt='0')
put(ws, 'B38', 'Подтверждено', font=F_BOLD, fill=SUMFILL); put(ws, 'C38', '=COUNTIF(I5:I34,"подтверждено")', fill=FILL_FORM, fmt='0')
put(ws, 'B39', 'Не подтверждено', font=F_BOLD, fill=SUMFILL); put(ws, 'C39', '=COUNTIF(I5:I34,"не подтверждено")', fill=FILL_FORM, fmt='0')
put(ws, 'B40', 'Ожидает проверки', font=F_BOLD, fill=SUMFILL); put(ws, 'C40', '=COUNTIF(I5:I34,"ожидает проверки")', fill=FILL_FORM, fmt='0')
put(ws, 'B42', '=IF(C39>0,"⚠ Есть неподтверждённые доказательства: оценки 3–5, опирающиеся на них, требуют пересмотра","✔ Неподтверждённых доказательств нет")', font=F_BOLD, fill=FILL_FORM)
ws.merge_cells('B42:J42')
widths(ws, {'A': 8, 'B': 12, 'C': 36, 'D': 11, 'E': 12, 'F': 24, 'G': 18, 'H': 18, 'I': 16, 'J': 24})
ws.freeze_panes = 'A5'

# ==================================================================
# ЛИСТ 8: AHP
# ==================================================================
ws = wb.create_sheet('AHP'); ws.sheet_properties.tabColor = 'C55A11'
title(ws, 'AHP — метод анализа иерархий (Т. Саати): веса направлений и блоков, λmax, CI, RI, CR',
      'Шкала парных сравнений: 1 одинаковая значимость · 3 умеренное · 5 существенное · 7 очень сильное · 9 абсолютное преимущество; 2/4/6/8 — промежуточные; обратное сравнение 1/aᵢⱼ. Заполняйте только ВЫДЕЛЕННЫЕ ячейки верхней треугольной части.',
      'Критерий согласованности CR ≤ 0,10. CR нельзя задавать заранее или подменять фиксированным коэффициентом. Групповая оценка — геометрическое среднее индивидуальных матриц (индивидуальные матрицы сохранять).')
RI_REF = 'Справочники!$B$82:$C$91'

def ahp_matrix(ws, top, names, ri_n, wcol_out):
    """top — строка заголовка матрицы; names — список критериев (n=3 или 4). Возвращает строки данных."""
    n = len(names)
    hdr = 1 + top
    header_row(ws, top, ['Критерий'] + names + ['вес w', '(Aw)ᵢ', '(Aw)ᵢ/wᵢ'])
    first = hdr
    for i in range(n):
        r = first + i
        put(ws, f'A{r}', names[i], font=F_BOLD, fill=FILL_FORM)
        for j in range(n):
            col = get_column_letter(2 + j)
            cell = f'{col}{r}'
            if i == j: put(ws, cell, 1, align=CTR, fill=FILL_FORM)
            elif j > i:
                c = ws[cell]; c.value = 1; c.border = BORD; c.alignment = CTR
                c.fill = PatternFill('solid', fgColor='FFF9E0')
            else:
                up = f'{get_column_letter(2+i)}{first+j}'
                put(ws, cell, f'=IF({up}>0,1/{up},"")', align=CTR, fill=FILL_FORM, fmt='0.0000')
    last = first + n - 1
    wcol = get_column_letter(2 + n); awcol = get_column_letter(3 + n); rcol = get_column_letter(4 + n)
    for i in range(n):
        r = first + i
        parts = '+'.join(f'{get_column_letter(2+j)}{r}/SUM({get_column_letter(2+j)}${first}:{get_column_letter(2+j)}${last})' for j in range(n))
        put(ws, f'{wcol}{r}', f'=({parts})/{n}', fmt='0.0000', fill=FILL_FORM, font=F_BOLD, align=CTR)
        aw = '+'.join(f'{get_column_letter(2+j)}{r}*${wcol}${first+j}' for j in range(n))
        put(ws, f'{awcol}{r}', f'={aw}', fmt='0.0000', fill=FILL_FORM)
        put(ws, f'{rcol}{r}', f'=IFERROR({awcol}{r}/{wcol}{r},"")', fmt='0.0000', fill=FILL_FORM)
    s = last + 2
    put(ws, f'B{s}', 'λmax', font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{s}', f'=IF(COUNT({rcol}{first}:{rcol}{last})<{n},"",SUM({rcol}{first}:{rcol}{last})/{n})', fmt=FMT_CR, fill=FILL_FORM)
    put(ws, f'B{s+1}', 'CI = (λmax − n)/(n − 1)', font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{s+1}', f'=IF(C{s}="","",ROUND((C{s}-{n})/{n-1},4))', fmt=FMT_CR, fill=FILL_FORM)
    put(ws, f'B{s+2}', f'RI (n={n}, из справочника)', font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{s+2}', f'=VLOOKUP({n},{RI_REF},2,0)', fmt=FMT_CR, fill=FILL_FORM)
    put(ws, f'B{s+3}', 'CR = CI / RI', font=F_RED, fill=SUMFILL)
    put(ws, f'C{s+3}', f'=IFERROR(ROUND(C{s+1}/C{s+2},4),"")', fmt=FMT_CR, fill=FILL_FORM, font=F_BOLD)
    put(ws, f'B{s+4}', 'Проверка', font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{s+4}', f'=IF(NOT(ISNUMBER(C{s+3})),"заполните матрицу",IF(C{s+3}>0.1,"⚠ CR > 0,10 — пересмотрите противоречивые сравнения","✔ CR ≤ 0,10"))', fill=FILL_FORM, font=F_RED)
    ws.merge_cells(f'C{s+4}:H{s+4}')
    return first, last, f'$C${s+3}', [f'${wcol}${first+i}' for i in range(n)]

sec(ws, 4, 'Матрица 1. Веса трёх направлений (Актуальность / Целесообразность / Результативность)')
d_first, d_last, d_cr, d_w = ahp_matrix(ws, 5, ['Актуальность', 'Целесообразность', 'Результативность'], 3, None)
sec(ws, 16, 'Матрица 2. Веса блоков актуальности А1–А4 (полный опросник)')
a_first, a_last, a_cr, a_w = ahp_matrix(ws, 17, ['А1 Потребность', 'А2 Перспективность', 'А3 Новизна', 'А4 Дублирование/стратегия'], 4, None)
sec(ws, 28, 'Матрица 3. Веса блоков целесообразности Б1–Б4')
b_first, b_last, b_cr, b_w = ahp_matrix(ws, 29, ['Б1 Концепция', 'Б2 Экономика', 'Б3 Ресурсы', 'Б4 Управление/риски'], 4, None)
sec(ws, 40, 'Матрица 4. Веса блоков результативности В1–В4')
v_first, v_last, v_cr, v_w = ahp_matrix(ws, 41, ['В1 Техрезультаты', 'В2 Документация/испытания', 'В3 Производство/внедрение', 'В4 Эффект'], 4, None)
sec(ws, 52, 'Сводка AHP')
put(ws, 'B53', 'Максимальный CR по четырём матрицам', font=F_BOLD, fill=SUMFILL)
put(ws, 'C53', f'=IF(COUNT({d_cr[1:]},{a_cr[1:]},{b_cr[1:]},{v_cr[1:]})=0,"",MAX({d_cr[1:]},{a_cr[1:]},{b_cr[1:]},{v_cr[1:]}))'.replace('$', '$'), fmt=FMT_CR, fill=FILL_FORM, font=F_BOLD)
put(ws, 'B54', 'Проверка согласованности', font=F_BOLD, fill=SUMFILL)
put(ws, 'C54', '=IF(NOT(ISNUMBER(C53)),"заполните матрицы",IF(C53>0.1,"⚠ AHP: пересмотрите противоречивые парные сравнения","✔ AHP: суждения согласованы (CR ≤ 0,10)"))', fill=FILL_FORM, font=F_RED)
ws.merge_cells('C54:H54')
put(ws, 'B56', 'Веса блоков из матриц 2–4 подставляются на лист «Полный_60» (колонка «Вес AHP») и применяются при переключателе «Использовать веса AHP = Да».', font=F_SMALL)
ws.merge_cells('B56:H57'); ws['B56'].alignment = WRAP
widths(ws, {'A': 26, 'B': 16, 'C': 13, 'D': 13, 'E': 13, 'F': 11, 'G': 11, 'H': 11})
# точные адреса CR матриц и весов AHP:
CR_CELLS = {'dir': f'C{d_last+5}', 'A': f'C{a_last+5}', 'B': f'C{b_last+5}', 'V': f'C{v_last+5}'}
AHP_W_CELLS = {'dir': [x.replace('$', '') for x in d_w],
               'A': [x.replace('$', '') for x in a_w], 'Б': [x.replace('$', '') for x in b_w], 'В': [x.replace('$', '') for x in v_w]}

# ==================================================================
# ЛИСТ 9: Стоп_факторы
# ==================================================================
ws = wb.create_sheet('Стоп_факторы'); ws.sheet_properties.tabColor = 'C00000'
title(ws, 'СТОП-ФАКТОРЫ И КРИТИЧЕСКИЕ ОБСТОЯТЕЛЬСТВА (SF-0 / SF-1 / SF-2)',
      'Высокий интегральный балл не должен скрывать обстоятельства, делающие продолжение проекта в текущем виде необоснованным. SF-2 не означает автоматического прекращения НИОКР — решение принимается после анализа контрактных обязательств, прав на РИД, последствий для заказчика и возможности корректировки.',
      'SF-0 отсутствует → обычный порядок · SF-1 существенная неопределённость → обязательная дополнительная проверка · SF-2 критическое обстоятельство → эскалация руководству и запрет безусловно положительного заключения.')
header_row(ws, 5, ['Проверка / обстоятельство', 'Статус', 'SF (авто)', 'Дата', 'Материалы / обоснование', 'Действие'])
for i, s in enumerate(D['SF2']):
    r = 6 + i
    put(ws, f'A{r}', f'Проверка SF-2 №{i+1}: {s}', font=F_BOLD, fill=FILL_WARN)
    put(ws, f'B{r}', 'не выявлено', align=CTR)
    put(ws, f'C{r}', f'=IF(B{r}="не выявлено",0,IF(B{r}="выявлена неопределённость",1,2))', align=CTR, fill=FILL_FORM, font=F_BOLD)
    for col in 'DEF': ws[f'{col}{r}'].border = BORD; ws[f'{col}{r}'].alignment = WRAP
for i in range(5):
    r = 11 + i
    put(ws, f'A{r}', f'Иное обстоятельство {i+1} (вписать):')
    put(ws, f'B{r}', 'не выявлено', align=CTR)
    put(ws, f'C{r}', f'=IF(B{r}="не выявлено",0,IF(B{r}="выявлена неопределённость",1,2))', align=CTR, fill=FILL_FORM, font=F_BOLD)
    for col in 'DEF': ws[f'{col}{r}'].border = BORD; ws[f'{col}{r}'].alignment = WRAP
dv(ws, '"не выявлено,выявлена неопределённость,подтверждено"', ['B6:B15'])
sec(ws, 17, 'Сводка')
put(ws, 'B18', 'Максимальный уровень SF', font=F_BOLD, fill=SUMFILL); put(ws, 'C18', '=MAX(C6:C15)', fill=FILL_FORM, font=F_BOLD, align=CTR)
put(ws, 'B19', 'Количество SF-1', font=F_BOLD, fill=SUMFILL); put(ws, 'C19', '=COUNTIF(C6:C15,1)', fill=FILL_FORM, align=CTR)
put(ws, 'B20', 'Количество SF-2', font=F_RED, fill=SUMFILL); put(ws, 'C20', '=COUNTIF(C6:C15,2)', fill=FILL_FORM, font=F_RED, align=CTR)
put(ws, 'B21', 'Итог', font=F_BOLD, fill=SUMFILL)
put(ws, 'C21', ('=IF(C20>0,"SF-2: эскалация руководству; безусловно положительное заключение невозможно до рассмотрения критического обстоятельства",'
                'IF(C19>0,"SF-1: обязательная дополнительная проверка (отдельный план проверки; не равен автоматически критическому)","SF-0: обычный порядок оценки"))'),
    fill=FILL_FORM, font=F_RED)
ws.merge_cells('C21:F21'); ws['C21'].alignment = WRAP
widths(ws, {'A': 60, 'B': 22, 'C': 8, 'D': 11, 'E': 30, 'F': 30})
SF_C = {'max': 'C18', 'n1': 'C19', 'n2': 'C20', 'res': 'C21'}

# ==================================================================
# ЛИСТ 10: Зрелость_M0_M9
# ==================================================================
ws = wb.create_sheet('Зрелость_M0_M9'); ws.sheet_properties.tabColor = 'ED7D31'
title(ws, 'ЗРЕЛОСТЬ РЕЗУЛЬТАТА M0–M9 — готовность к следующему этапу',
      'Внутренняя шкала зрелости (не замена нормативно установленных стадий разработки и испытаний). Оценивается соответствие достигнутого уровня запланированному этапу (паспорт, строка «Планируемая зрелость M»).',
      'Отставание фактической зрелости от плановой — серьёзный сигнал для аудита (например, TRL 2 через два года при плановом прототипе TRL 5).')
header_row(ws, 4, ['Уровень', 'Состояние результата', 'Что должно быть подтверждено', 'Подтверждение есть', 'Доказательство (ID)'])
for i, (m, st, cf) in enumerate(D['M09']):
    r = 5 + i
    put(ws, f'A{r}', m, font=F_BOLD, align=CTR, fill=FILL_FORM)
    put(ws, f'B{r}', st, font=F_BOLD)
    put(ws, f'C{r}', cf)
    put(ws, f'D{r}', '—', align=CTR)
    ws[f'E{r}'].border = BORD
dv(ws, '"Да,Нет,—"', ['D5:D14'])
sec(ws, 16, 'Оценка')
MAT = [
    ('Текущий уровень зрелости M (0–9)', 0, '"0,1,2,3,4,5,6,7,8,9"'),
    ('Планируемый уровень M (по этапу)', 0, '"0,1,2,3,4,5,6,7,8,9"'),
]
for i, (lab, val, dvl) in enumerate(MAT):
    r = 17 + i
    put(ws, f'B{r}', lab, font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{r}', val, align=CTR); dv(ws, dvl, [f'C{r}'])
put(ws, 'B19', 'Отклонение (факт − план)', font=F_BOLD, fill=SUMFILL)
put(ws, 'C19', '=IF(AND(ISNUMBER(C17),ISNUMBER(C18)),C17-C18,"")', fill=FILL_FORM, align=CTR)
put(ws, 'B20', 'Примечание по этапу', font=F_BOLD, fill=SUMFILL)
put(ws, 'C20', '=IF(AND(ISNUMBER(C17),ISNUMBER(C18)),IF(C19>0,"опережение планового этапа",IF(C19<0,"⚠ отставание от планового этапа","соответствует плановому этапу")),"")', fill=FILL_FORM, font=F_BOLD)
ws.merge_cells('C20:F20')
put(ws, 'B22', 'Испытаниями подтверждён текущий уровень', font=F_BOLD, fill=SUMFILL)
put(ws, 'C22', 'Нет', align=CTR); dv(ws, '"Да,Нет"', ['C22'])
put(ws, 'B23', 'Автозамечание (контроль доказательств)', font=F_RED, fill=SUMFILL)
put(ws, 'C23', '=IF(AND(ISNUMBER(C17),C17>=5,C22<>"Да"),"⚠ Зрелость M"&C17&" заявлена без подтверждения испытаниями — недостаточно доказательств","OK")', fill=FILL_FORM, font=F_RED)
ws.merge_cells('C23:F23')
put(ws, 'B24', 'Уровней без подтверждения', font=F_BOLD, fill=SUMFILL)
put(ws, 'C24', '=COUNTIF(D5:D14,"Нет")', fill=FILL_FORM, align=CTR)
put(ws, 'B25', 'Замечание', font=F_BOLD, fill=SUMFILL)
put(ws, 'C25', '=IF(C24>0,"⚠ "&C24&" уровень(ей) без подтверждения — запросить доказательства","OK")', fill=FILL_FORM)
ws.merge_cells('C25:F25')
sec(ws, 27, 'Обязательные контрольные точки переоценки (динамическая оценка ΔIA, ΔIC)')
for i, t in enumerate(D['CKPT']):
    ws[f'B{28+i}'] = f'{i+1}) {t}'; ws[f'B{28+i}'].font = F_SMALL
put(ws, 'B37', 'Динамика: ΔI = I(t) − I(t−1); скорость V = (I(t) − I(0)) / t, где t — лет с начала НИОКР. Устойчивое снижение актуальности или целесообразности — сигнал к внеплановой переоценке (см. «Панель_руководителя», блок раннего предупреждения).', font=F_SMALL)
ws.merge_cells('B37:F39'); ws['B37'].alignment = WRAP
widths(ws, {'A': 9, 'B': 34, 'C': 52, 'D': 14, 'E': 14, 'F': 20})
MAT_C = {'cur': 'C17', 'plan': 'C18', 'dev': 'C19', 'note': 'C20', 'tests': 'C22', 'warn': 'C23'}

# ==================================================================
# ЛИСТ 11: Экономика_LCC
# ==================================================================
ws = wb.create_sheet('Экономика_LCC'); ws.sheet_properties.tabColor = '548235'
title(ws, 'ЭКОНОМИКА — LCC, NPV, BCR и сравнение альтернатив',
      'Для оборонных и государственных НИОКР отрицательный или неопределённый NPV НЕ является автоматическим основанием прекращения — отдельно оцениваются технический, стратегический, научный и инфраструктурный эффекты (нельзя искусственно переводить все эффекты в рубли).',
      'В затратах учитывать не только контрактную стоимость: трудозатраты, стенды, опытное производство, комплектующие, ПО, сертификацию, обучение, внедрение, сопровождение, эксплуатацию.')
put(ws, 'B4', 'Ставка дисконтирования r', font=F_BOLD, fill=SUMFILL); put(ws, 'C4', 0.10, fmt='0.00', align=CTR)
d = DataValidation(type='decimal', operator='between', formula1='0', formula2='1', allow_blank=True); ws.add_data_validation(d); d.add('C4')
put(ws, 'B5', 'Горизонт оценки T, лет', font=F_BOLD, fill=SUMFILL); put(ws, 'C5', 5, align=CTR)
sec(ws, 7, 'Полная стоимость жизненного цикла (LCC), 8 этапов')
header_row(ws, 8, ['', 'Этап', 'Состав затрат', 'Затраты, млн ₽', 'Год t', 'PV затрат'])
for i, (st, comp) in enumerate(D['LCC']):
    r = 9 + i
    put(ws, f'B{r}', st, font=F_BOLD, fill=FILL_FORM)
    put(ws, f'C{r}', comp, font=F_SMALL)
    ws[f'D{r}'].border = BORD; ws[f'D{r}'].number_format = FMT_MONEY
    ws[f'E{r}'].border = BORD; ws[f'E{r}'].alignment = CTR
    put(ws, f'F{r}', f'=IF(OR(D{r}="",E{r}=""),"",ROUND(D{r}/POWER(1+$C$4,E{r}),2))', fmt=FMT_MONEY, fill=FILL_FORM)
put(ws, 'B17', 'LCC итого (PV)', font=F_RED, fill=SUMFILL)
put(ws, 'F17', '=ROUND(SUM(F9:F16),2)', fmt=FMT_MONEY, fill=FILL_FORM, font=F_RED)
sec(ws, 19, 'NPV и BCR по денежным потокам')
header_row(ws, 20, ['t', 'Эффект Bt, млн ₽', 'Затраты Ct, млн ₽', 'PV(Bt)', 'PV(Ct)'])
for t in range(11):
    r = 21 + t
    put(ws, f'A{r}', t, align=CTR, fill=FILL_FORM)
    ws[f'B{r}'].border = BORD; ws[f'B{r}'].number_format = FMT_MONEY
    ws[f'C{r}'].border = BORD; ws[f'C{r}'].number_format = FMT_MONEY
    put(ws, f'D{r}', f'=IF(B{r}="","",ROUND(B{r}/POWER(1+$C$4,$A{r}),2))', fmt=FMT_MONEY, fill=FILL_FORM)
    put(ws, f'E{r}', f'=IF(C{r}="","",ROUND(C{r}/POWER(1+$C$4,$A{r}),2))', fmt=FMT_MONEY, fill=FILL_FORM)
ECON = [
    ('NPV = ΣPV(B) − ΣPV(C), млн ₽', '=ROUND(SUM(D21:D31)-SUM(E21:E31),2)', FMT_MONEY),
    ('PV выгод B, млн ₽', '=ROUND(SUM(D21:D31),2)', FMT_MONEY),
    ('PV затрат C, млн ₽', '=ROUND(SUM(E21:E31),2)', FMT_MONEY),
    ('BCR = PV(B) / PV(C)', '=IFERROR(ROUND(C33/C34,2),"—")', '0.00'),
    ('Интерпретация BCR', '=IF(NOT(ISNUMBER(C35)),"—",IF(C35>1.5,"Высокая экономическая привлекательность",IF(C35>=1,"Приемлемая",IF(C35>=0.7,"Требуется дополнительное обоснование","Низкая экономическая привлекательность"))))', None),
]
r = 32
for lab, f, fmt in ECON:
    put(ws, f'B{r}', lab, font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{r}', f, fill=FILL_FORM, font=F_BOLD, fmt=fmt)
    ws.merge_cells(f'C{r}:E{r}')
    r += 1
put(ws, 'B37', 'Ориентиры BCR — предлагаемые для внутреннего аудита, не универсальные нормативы. Дополнительно оценивать: стоимость изделия, стоимость достижения характеристики, стоимость владения комплектом, затраты на ремонт, масштабирование, замещение импортных компонентов, эффект повторного использования решений.', font=F_SMALL)
ws.merge_cells('B37:E39'); ws['B37'].alignment = WRAP
sec(ws, 41, 'Сравнение альтернатив (функция полезности Us = Σ wᵢ·xᵢ; оценки 0–5; для каждого варианта указывать источник данных и достоверность)')
header_row(ws, 42, ['Критерий', 'Вес'] + D['ALTNAMES'])
for i, cr in enumerate(D['ALTCRIT']):
    r = 43 + i
    put(ws, f'A{r}', cr, font=F_BOLD, fill=FILL_FORM)
    put(ws, f'B{r}', 1, align=CTR, fmt=FMT_W)
    for col in 'CDEF':
        c = ws[f'{col}{r}']; c.border = BORD; c.alignment = CTR
dv(ws, '"0,1,2,3,4,5"', ['C43:F51'])
put(ws, 'A52', 'Us (интегральная полезность)', font=F_RED, fill=SUMFILL)
for col in 'CDEF':
    put(ws, f'{col}52', f'=SUMPRODUCT($B$43:$B$51,{col}43:{col}51)', fmt=FMT_W, fill=FILL_FORM, font=F_BOLD, align=CTR)
put(ws, 'A53', 'Наилучший сценарий', font=F_RED, fill=SUMFILL)
put(ws, 'C53', '=INDEX($C$42:$F$42,MATCH(MAX(C52:F52),C52:F52,0))', fill=FILL_FORM, font=F_RED, align=CTR)
ws.merge_cells('C53:F53')
put(ws, 'A54', 'Преимущество над «Продолжить разработку»', font=F_BOLD, fill=SUMFILL)
put(ws, 'C54', '=ROUND(MAX(C52:F52)-C52,2)', fill=FILL_FORM, align=CTR)
ws.merge_cells('C54:F54')
put(ws, 'A55', 'Источник данных / достоверность (по сценариям)', font=F_SMALL, fill=SUMFILL)
for col in 'CDEF':
    ws[f'{col}55'].border = BORD; ws[f'{col}55'].alignment = WRAP
put(ws, 'A57', 'Неизвестные параметры нельзя автоматически оценивать максимальным баллом. Самая дешёвая альтернатива не обязательно является наиболее целесообразной. Экономическая выгода не компенсирует несоответствие обязательным требованиям безопасности (они рассматриваются раздельно). Минимум четыре сценария: S1 продолжить в существующем виде · S2 скорректировать ТЗ и продолжить · S3 приобрести/лицензировать/адаптировать · S4 прекратить либо заменить другим проектом.', font=F_SMALL)
ws.merge_cells('A57:F60'); ws['A57'].alignment = WRAP
widths(ws, {'A': 34, 'B': 30, 'C': 18, 'D': 18, 'E': 18, 'F': 18})
ECON_C = {'r': 'C4', 'npv': 'C32', 'pvb': 'C33', 'pvc': 'C34', 'bcr': 'C35', 'bcr_txt': 'C36',
          'us': ['C52', 'D52', 'E52', 'F52'], 'best': 'C53', 'margin': 'C54', 'lcc': 'F17'}
# fix: C48 не используется — панель/заключение ссылаются на C53 (best)
# (в формуле Полный_60!C102 использован адрес Экономика_LCC!$C$48 — исправим ниже через замену)

# ==================================================================
# ЛИСТ 12: План_мероприятий
# ==================================================================
ws = wb.create_sheet('План_мероприятий'); ws.sheet_properties.tabColor = '7F7F7F'
title(ws, 'ПЛАН КОРРЕКТИРУЮЩИХ МЕРОПРИЯТИЙ — контроль исполнения рекомендаций',
      'Для каждого мероприятия фиксируются статус, подтверждающий документ, остаточный риск и результат повторной проверки. Статус «Выполнено» — только при объективном подтверждении (докладная записка исполнителя недостаточна).',
      'Категории: К1 критические (безотлагательно) · К2 существенные (в приоритетном порядке) · К3 организационные (по плану) · К4 развивающие (плановые).')
header_row(ws, 4, ['№', 'Выявленная проблема', 'Мероприятие', 'Категория', 'Ответственный', 'Срок', 'Критерий устранения', 'Статус', 'Подтверждающий документ', 'Остаточный риск', 'Результат повторной проверки'])
for i in range(12):
    r = 5 + i
    put(ws, f'A{r}', i+1, align=CTR, fill=FILL_FORM)
    for col in 'BCDEFGHIJK':
        c = ws[f'{col}{r}']; c.border = BORD; c.alignment = WRAP
dv(ws, '"К1,К2,К3,К4"', ['D5:D16'])
dv(ws, '"Не начато,В работе,На проверке,Выполнено,Не выполнено,Перенесено,Отклонено"', ['H5:H16'])
sec(ws, 18, 'Сводка статусов')
for i, (st, _) in enumerate(D['STAT']):
    put(ws, f'B{19+i}', st, font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{19+i}', f'=COUNTIF($H$5:$H$16,"{st}")', fill=FILL_FORM, align=CTR)
put(ws, 'B27', '=IF(COUNTIF(H5:H16,"Не выполнено")>0,"⚠ Есть невыполненные мероприятия: повторное невыполнение — эскалация (рассмотрение ответственности и альтернатив)","✔ Невыполненных мероприятий нет")', font=F_RED, fill=FILL_FORM)
ws.merge_cells('B27:K27')
sec(ws, 29, 'Повторный аудит проводится: не выполнены критические мероприятия · новые существенные обстоятельства · изменена техконцепция · изменились требования заказчика · значительное изменение стоимости/сроков · переход на следующий значимый этап ЖЦ · подтверждение готовности к установочной партии или серии.')
ws['A29'].alignment = WRAP; ws['A29'].font = F_SMALL; ws.merge_cells('A29:K31')
widths(ws, {'A': 4, 'B': 30, 'C': 34, 'D': 9, 'E': 16, 'F': 11, 'G': 24, 'H': 13, 'I': 22, 'J': 16, 'K': 22})
PLAN_C = {'statuses': [f'C{19+i}' for i in range(7)]}

# ==================================================================
# ЛИСТ 13: Заключение
# ==================================================================
ws = wb.create_sheet('Заключение'); ws.sheet_properties.tabColor = '203864'
title(ws, 'ЗАКЛЮЧЕНИЕ ПО РЕЗУЛЬТАТАМ АУДИТА НИОКР (формируется автоматически)',
      'Единая аналитическая карта: три самостоятельных индекса + комплексный (справочно) + Coverage + AHP CR + SF + зрелость. Итог не сводится к одному среднему баллу.',
      'Прекращение НИОКР не производится только по интегральному индексу — требуется отдельное заключение (контрактные обязательства, правовые последствия, незавершённые результаты, РИД, возможность использования полученных результатов).')
LVL = 'Паспорт_НИОКР!$C$17'
put(ws, 'B4', 'НИОКР', font=F_BOLD, fill=SUMFILL); put(ws, 'C4', '=Паспорт_НИОКР!$C$6', fill=FILL_FORM)
put(ws, 'B5', 'Уровень аудита', font=F_BOLD, fill=SUMFILL); put(ws, 'C5', f'={LVL}', fill=FILL_FORM, font=F_BOLD)
put(ws, 'B6', 'Шифр', font=F_BOLD, fill=SUMFILL); put(ws, 'C6', '=Паспорт_НИОКР!$C$7', fill=FILL_FORM)
L3 = lambda k: (f'=IF($C$5="Краткий",Краткий_15!${K15C[k][0]}${K15C[k][1:]},'
                f'IF($C$5="Средний",Средний_30!${C30C[k][0]}${C30C[k][1:]},Полный_60!${C60C[k][0]}${C60C[k][1:]}))')
Z = [
    ('Индекс актуальности I_А, %', L3('IA'), FMT_PCT),
    ('Индекс целесообразности I_Ц, %', L3('IC'), FMT_PCT),
    ('Индекс результативности I_Р, %', L3('IR'), FMT_PCT),
    ('Комплексный индекс (∛(I_А·I_Ц·I_Р), справочно)', L3('KOMPL'), FMT_PCT),
    ('Полнота доказательств Coverage, %', L3('COV'), FMT_PCT),
    ('Порог Coverage для уровня', '=IF($C$5="Краткий",80,IF($C$5="Средний",85,90))', '0'),
    ('Проверка доказательности', L3('OK'), None),
    ('Минимальный индекс', L3('MIN'), FMT_PCT),
    ('Количество SF-2', '=Стоп_факторы!$C$20', '0'),
    ('Количество SF-1', '=Стоп_факторы!$C$19', '0'),
    ('Стоп-факторы: итог', '=Стоп_факторы!$C$21', None),
    ('Зрелость результата M', '=Зрелость_M0_M9!$C$17', '0'),
    ('Зрелость: примечание по этапу', '=Зрелость_M0_M9!$C$20', None),
    ('Зрелость: автозамечание по доказательствам', '=Зрелость_M0_M9!$C$23', None),
    ('AHP: максимальный CR', '=AHP!$C$53', FMT_CR),
    ('AHP: проверка согласованности', '=AHP!$C$54', None),
]
summary_block(ws, 7, Z)
put(ws, 'B24', 'ИТОГОВОЕ ЗАКЛЮЧЕНИЕ / РЕШЕНИЕ (авто по уровню)', font=F_RED, fill=SUMFILL)
put(ws, 'C24', L3('CONC'), fill=FILL_FORM, font=F_RED)
ws.merge_cells('C24:N24'); ws['C24'].alignment = WRAP; ws.row_dimensions[24].height = 42
sec(ws, 26, 'АВТОМАТИЧЕСКИЕ РЕКОМЕНДАЦИИ (средний опросник; условие — итоговый балл вопроса ≤ 2)')
header_row(ws, 27, ['Код', 'Условие', 'Сработала', 'Рекомендация'])
for i, (cond, rec) in enumerate(AUTO_REC):
    r = 28 + i
    code = cond.split('<=')[0].strip()
    row30 = ROW30[code]
    put(ws, f'A{r}', code, align=CTR, font=F_BOLD, fill=FILL_FORM)
    put(ws, f'B{r}', f'{code} ≤ 2', align=CTR, font=F_SMALL)
    put(ws, f'C{r}', f'=IF(AND($C$5="Средний",ISNUMBER(Средний_30!$M${row30}),Средний_30!$M${row30}<=2),"сработала","")', fill=FILL_FORM, align=CTR)
    put(ws, f'D{r}', rec, font=F_SMALL)
put(ws, 'B43', 'Всего сработало:', font=F_BOLD, fill=SUMFILL)
put(ws, 'C43', '=COUNTIF(C28:C42,"сработала")', fill=FILL_FORM, font=F_BOLD, align=CTR)
sec(ws, 45, 'ОБЩИЕ ПРАВИЛА (любой уровень)')
header_row(ws, 46, ['', 'Условие', 'Сработало', 'Действие'])
GEN = [
    ('I_А < 55%', '=IF(AND(ISNUMBER($C$7),$C$7<55),"сработала","")', 'Повторно подтвердить потребность и актуальность у заказчика'),
    ('I_Ц < 55%', '=IF(AND(ISNUMBER($C$8),$C$8<55),"сработала","")', 'Провести технико-экономический анализ альтернатив'),
    ('I_Р < 55%', '=IF(AND(ISNUMBER($C$9),$C$9<55),"сработала","")', 'Проверить технические результаты; утвердить план восстановления'),
    ('Coverage ниже порога', '=IF(AND(ISNUMBER($C$11),$C$11<$C$12),"сработала","")', 'Запросить доказательства; положительное заключение не выносится'),
]
for i, (cond, f, act) in enumerate(GEN):
    r = 47 + i
    put(ws, f'B{r}', cond, font=F_BOLD, fill=FILL_FORM)
    put(ws, f'C{r}', f, fill=FILL_FORM, align=CTR)
    put(ws, f'D{r}', act, font=F_SMALL)
put(ws, 'B51', 'Всего сработало:', font=F_BOLD, fill=SUMFILL)
put(ws, 'C51', '=COUNTIF(C47:C50,"сработала")', fill=FILL_FORM, font=F_BOLD, align=CTR)
sec(ws, 53, 'ШАБЛОН ЗАКЛЮЧЕНИЯ (20 пунктов унифицированной структуры)')
TMPL_F = [
    '=Паспорт_НИОКР!$C$6', '=Паспорт_НИОКР!$C$8', '=Паспорт_НИОКР!$C$9',
    '=Паспорт_НИОКР!$C$10&"; кооперация: "&Паспорт_НИОКР!$C$11',
    '=Паспорт_НИОКР!$C$12&" — "&Паспорт_НИОКР!$C$13',
    '=IF(ISNUMBER(Паспорт_НИОКР!$C$14),Паспорт_НИОКР!$C$14&" млн ₽","—")',
    '=Паспорт_НИОКР!$C$15',
    '(заполняется вручную: цель и ожидаемый результат)',
    '=IF(ISNUMBER($C$7),$C$7&"%","н/д")', '=IF(ISNUMBER($C$8),$C$8&"%","н/д")', '=IF(ISNUMBER($C$9),$C$9&"%","н/д")',
    '=IF(ISNUMBER($C$10),$C$10&"% (справочно)","н/д")',
    '=IF(ISNUMBER($C$11),$C$11&"% — "&$C$13,"н/д")',
    '=IF(ISNUMBER($C$21),"CR max = "&$C$21&" ("&$C$22&")","AHP не заполнен")',
    '=$C$17',
    '=IF($C$5="Краткий","к обсуждению: "&Краткий_15!$C$30&"; с замечаниями: "&Краткий_15!$C$31&"; балл ≤2: "&Краткий_15!$C$33,IF($C$5="Средний","к обсуждению: "&Средний_30!$C$45&"; с замечаниями: "&Средний_30!$C$46&"; балл ≤2: "&Средний_30!$C$48,"к обсуждению: "&Полный_60!$C$91&"; с замечаниями: "&Полный_60!$C$92&"; балл ≤2: "&Полный_60!$C$94))',
    '="Наилучший сценарий: "&Экономика_LCC!$C$53&" (преимущество над продолжением разработки: "&Экономика_LCC!$C$54&")"',
    '="Зрелость M"&$C$18&" ("&$C$19&"); "&$C$20',
    '=$C$24',
    '="Лист «План_мероприятий»; сработало авторекомендаций: "&$C$43&"; общих правил: "&$C$51',
]
for i, item in enumerate(D['TMPL']):
    r = 54 + i
    put(ws, f'B{r}', f'{i+1}. {item}', font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{r}', TMPL_F[i], fill=FILL_FORM)
    ws.merge_cells(f'C{r}:N{r}')
sec(ws, 75, 'ТРИ САМОСТОЯТЕЛЬНЫХ ВЫВОДА (форма Ф-06)')
def LVLNAME(cell):
    return f'IF({cell}>=85,"Высокий",IF({cell}>=70,"Приемлемый",IF({cell}>=55,"Проблемный","Критический")))'
F06 = [
    ('По актуальности', '=IF(ISNUMBER($C$7),"I_А = "&$C$7&"% — "&' + LVLNAME('$C$7') + '&" уровень. Сохраняется ли объективная потребность, новизна, отсутствие неоправданного дублирования — см. вопросы блока А и авторекомендации.","I_А — недостаточно данных")'),
    ('По целесообразности', '=IF(ISNUMBER($C$8),"I_Ц = "&$C$8&"% — "&' + LVLNAME('$C$8') + '&" уровень. Обоснованность концепции, затрат, ресурсов, сопоставление альтернатив — см. вопросы блока Б и лист «Экономика_LCC».","I_Ц — недостаточно данных")'),
    ('По результативности', '=IF(ISNUMBER($C$9),"I_Р = "&$C$9&"% — "&' + LVLNAME('$C$9') + '&" уровень. Достижение требований ТЗ, подтверждение характеристик, устранение замечаний, внедрение и производство — см. вопросы блока В.","I_Р — недостаточно данных")'),
]
for i, (lab, f) in enumerate(F06):
    r = 76 + i
    put(ws, f'B{r}', lab + ':', font=F_BOLD, fill=SUMFILL)
    put(ws, f'C{r}', f, fill=FILL_FORM)
    ws.merge_cells(f'C{r}:N{r}'); ws[f'C{r}'].alignment = WRAP
sec(ws, 80, 'ОГРАНИЧЕНИЯ И КАЛИБРОВКА')
LIMS = [
    'Веса и пороговые значения — проектная модель методики: до калибровки на исторических данных (20–30 завершённых НИОКР) не рассматриваются как доказанные универсальные нормативы.',
    'SF-2 не означает автоматического прекращения НИОКР — эскалация и запрет безусловно положительного заключения до рассмотрения обстоятельства.',
    'Большинство голосов комиссии не может превратить неподтверждённое утверждение в установленный факт: голосование определяет позицию, но не отменяет доказательности.',
    'Изменение экспертной оценки исключительно по административным мотивам не допускается (только новые доказательства / ошибка расчёта / повторное рассмотрение с фиксацией причины).',
    'Особое мнение эксперта прилагается к заключению и не может быть удалено или изменено без согласия автора.',
]
for i, t in enumerate(LIMS):
    ws[f'B{81+i}'] = '• ' + t; ws[f'B{81+i}'].font = F_SMALL; ws[f'B{81+i}'].alignment = WRAP
    ws.merge_cells(f'B{81+i}:N{81+i}')
widths(ws, {'A': 6, 'B': 44, 'C': 16, 'D': 60, 'E': 12, 'F': 12, 'G': 12, 'H': 12, 'I': 12, 'J': 12, 'K': 12, 'L': 12, 'M': 12, 'N': 12})
ZAKL_C = {'lvl': 'C5', 'IA': 'C7', 'IC': 'C8', 'IR': 'C9', 'kompl': 'C10', 'cov': 'C11', 'thr': 'C12',
          'ok': 'C13', 'min': 'C14', 'sf2': 'C15', 'sf1': 'C16', 'sfres': 'C17', 'M': 'C18', 'Mnote': 'C19',
          'Mwarn': 'C20', 'cr': 'C21', 'crchk': 'C22', 'conc': 'C24', 'nrec': 'C43', 'ngen': 'C51'}

# ==================================================================
# ЛИСТ 1: Панель_руководителя
# ==================================================================
wsp = wb.create_sheet('Панель_руководителя', 0); wsp.sheet_properties.tabColor = 'C00000'
title(wsp, 'ПАНЕЛЬ РУКОВОДИТЕЛЯ — сводные показатели аудита НИОКР',
      'Порядок заполнения: Паспорт → Эксперты → опросник выбранного уровня (Краткий_15 / Средний_30 / Полный_60) → Доказательства → Стоп_факторы → Зрелость → Экономика → (при необходимости) AHP → План_мероприятий. Заключение формируется автоматически.',
      'Не сводить всё к одному баллу: высокая актуальность не компенсирует отсутствие реализуемости или потребителя. Три индекса всегда рассматриваются самостоятельно.')
put(wsp, 'B4', 'НИОКР:', font=F_BOLD, fill=SUMFILL); put(wsp, 'C4', '=Заключение!$C$4', fill=FILL_FORM); wsp.merge_cells('C4:H4')
put(wsp, 'B5', 'Уровень аудита:', font=F_BOLD, fill=SUMFILL); put(wsp, 'C5', '=Заключение!$C$5', fill=FILL_FORM, font=F_BOLD)
put(wsp, 'B6', 'Дата аудита:', font=F_BOLD, fill=SUMFILL); put(wsp, 'C6', '=Паспорт_НИОКР!$C$18', fill=FILL_FORM)
sec(wsp, 8, 'СВОДНЫЕ ПОКАЗАТЕЛИ')
PANEL = [
    ('Актуальность I_А', '=Заключение!$C$7', FMT_PCT),
    ('Целесообразность I_Ц', '=Заключение!$C$8', FMT_PCT),
    ('Результативность I_Р', '=Заключение!$C$9', FMT_PCT),
    ('Комплексный индекс (справочно)', '=Заключение!$C$10', FMT_PCT),
    ('Полнота доказательств Coverage', '=Заключение!$C$11', FMT_PCT),
    ('Порог Coverage / проверка', '=Заключение!$C$12&"% — "&Заключение!$C$13', None),
    ('Стоп-факторы', '=Заключение!$C$17', None),
    ('Зрелость результата', '=IF(ISNUMBER(Заключение!$C$18),"M"&Заключение!$C$18&" ("&Заключение!$C$19&")","—")', None),
    ('AHP (макс. CR)', '=IF(ISNUMBER(Заключение!$C$21),"CR = "&Заключение!$C$21&" — "&Заключение!$C$22,"—")', None),
]
summary_block(wsp, 9, PANEL)
put(wsp, 'B19', 'ЗАКЛЮЧЕНИЕ:', font=F_RED, fill=SUMFILL)
put(wsp, 'C19', '=Заключение!$C$24', fill=FILL_FORM, font=F_RED)
wsp.merge_cells('C19:J19'); wsp['C19'].alignment = WRAP; wsp.row_dimensions[19].height = 42
sec(wsp, 21, 'РАННЕЕ ПРЕДУПРЕЖДЕНИЕ (динамика и триггеры)')
put(wsp, 'B22', 'Предыдущий аудит — I_А / I_Ц / I_Р (ввести):', font=F_BOLD, fill=SUMFILL)
for col in 'CDE':
    c = wsp[f'{col}22']; c.border = BORD; c.alignment = CTR; c.number_format = FMT_PCT
header_row(wsp, 23, ['Событие', 'Индикатор', 'Статус', 'Действие'])
EARLY_ROWS = [
    ('I_А снизился на 10 пунктов и более', '=IF(ISNUMBER($C$22),ROUND($C$9-$C$22,1),"")', '=IF(AND(ISNUMBER($B$24),$B$24<=-10),"⚠ Предупреждение","OK")', 'Повторная проверка потребности'),
    ('I_Ц снизился на 10 пунктов и более', '=IF(ISNUMBER($D$22),ROUND($C$10-$D$22,1),"")', '=IF(AND(ISNUMBER($B$25),$B$25<=-10),"⚠ Предупреждение","OK")', 'Пересмотр концепции и ресурсов'),
    ('I_Р снизился на 10 пунктов и более', '=IF(ISNUMBER($E$22),ROUND($C$11-$E$22,1),"")', '=IF(AND(ISNUMBER($B$26),$B$26<=-10),"⚠ Критический сигнал","OK")', 'Проверка технических результатов'),
    ('Coverage ниже 85%', '=Заключение!$C$11', '=IF(AND(ISNUMBER($B$27),$B$27<85),"⚠ Недостаточно данных","OK")', 'Запрос доказательств'),
    ('Появился SF-2', '=Заключение!$C$15', '=IF(AND(ISNUMBER($B$28),$B$28>0),"⚠ Эскалация","OK")', 'Внеплановое заседание комиссии'),
    ('Срок этапа превышен', '=IFERROR(IF(ISNUMBER(Паспорт_НИОКР!$C$13),Паспорт_НИОКР!$C$13-TODAY(),""),"")', '=IF(AND(ISNUMBER($B$29),$B$29<0),"⚠ Отклонение","OK")', 'Анализ критического пути'),
    ('Повторное невыполнение мероприятия', '=COUNTIF(План_мероприятий!$H$5:$H$16,"Не выполнено")', '=IF(AND(ISNUMBER($B$30),$B$30>0),"⚠ Эскалация","OK")', 'Рассмотрение ответственности и альтернатив'),
]
for i, (ev, ind, st_, act) in enumerate(EARLY_ROWS):
    r = 24 + i
    put(wsp, f'A{r}', ev, font=F_BOLD, fill=FILL_FORM)
    put(wsp, f'B{r}', ind, fill=FILL_FORM, align=CTR)
    put(wsp, f'C{r}', st_, fill=FILL_FORM, font=F_BOLD, align=CTR)
    put(wsp, f'D{r}', act, font=F_SMALL)
wsp.conditional_formatting.add('C24:C30', FormulaRule(formula=['LEFT($C24,1)="⚠"'], fill=PatternFill('solid', fgColor='F8CBAD'), font=Font(color='C00000', bold=True)))
sec(wsp, 32, 'СЧЁТЧИКИ КАЧЕСТВА ОЦЕНОК')
CNT = [
    ('К обсуждению — Краткий / Средний / Полный', '=Краткий_15!$C$30&" / "&Средний_30!$C$45&" / "&Полный_60!$C$91', None),
    ('Контрольные замечания — Краткий / Средний / Полный', '=Краткий_15!$C$31&" / "&Средний_30!$C$46&" / "&Полный_60!$C$92', None),
    ('Вопросов с баллом ≤ 2 — Краткий / Средний / Полный', '=Краткий_15!$C$33&" / "&Средний_30!$C$48&" / "&Полный_60!$C$94', None),
    ('Сработало авторекомендаций / общих правил', '=Заключение!$C$43&" / "&Заключение!$C$51', None),
    ('План: в работе / выполнено / не выполнено', '=COUNTIF(План_мероприятий!$H$5:$H$16,"В работе")&" / "&COUNTIF(План_мероприятий!$H$5:$H$16,"Выполнено")&" / "&COUNTIF(План_мероприятий!$H$5:$H$16,"Не выполнено")', None),
    ('Доказательства: всего / подтверждено / не подтверждено', '=Доказательства!$C$37&" / "&Доказательства!$C$38&" / "&Доказательства!$C$39', None),
]
summary_block(wsp, 33, CNT)
put(wsp, 'B40', 'Три уровня управления: проектный (руководитель НИОКР, ежемесячно) · экспертно-аудиторский (независимая оценка I_А/I_Ц/I_Р по контрольным рубежам) · портфельный (руководство, ежеквартально). Портфельный аудит (25 вопросов, 15 KPI, матрица приоритизации) — отдельный раздел методики v2.', font=F_SMALL)
wsp.merge_cells('B40:J42'); wsp['B40'].alignment = WRAP
widths(wsp, {'A': 40, 'B': 26, 'C': 16, 'D': 40, 'E': 14, 'F': 12, 'G': 12, 'H': 12})
wsp.freeze_panes = 'A4'

# ==================================================================
# ЛИСТ 14: Справочники
# ==================================================================
ws = wb.create_sheet('Справочники'); ws.sheet_properties.tabColor = 'A6A6A6'
title(ws, 'СПРАВОЧНИКИ — шкалы, веса, пороги, классификаторы и настройки',
      'Все пороги и веса являются проектными (пилотными): перед организационным применением калибруются на 20–30 ранее завершённых НИОКР и проверяются на устойчивость (веса ±10–20%, баллы ±1).',
      None)
r = 4
sec(ws, r, '1. Шкала экспертных оценок (единая для всех трёх опросников)'); r += 1
header_row(ws, r, ['Балл', 'Интерпретация', 'Правило выставления / требования к подтверждению']); r += 1
for b, i2, rule in D['SCALE']:
    put(ws, f'A{r}', b, align=CTR, font=F_BOLD, fill=FILL_FORM); put(ws, f'B{r}', i2); put(ws, f'C{r}', rule); r += 1
r += 1
sec(ws, r, '2. Уровни индексов (итоговая логика заключения)'); r += 1
header_row(ws, r, ['Уровень', 'Значение', 'Управленческое решение']); r += 1
for a, b2, c2 in D['POROG_LVL']:
    put(ws, f'A{r}', a, font=F_BOLD, fill=FILL_FORM); put(ws, f'B{r}', b2, align=CTR); put(ws, f'C{r}', c2); r += 1
r += 1
sec(ws, r, '3. Краткий опросник: пороги и формулировки заключений'); r += 1
header_row(ws, r, ['Индекс', 'Оценка', 'Рекомендация']); r += 1
for a, b2, c2 in D['POROG15']:
    put(ws, f'A{r}', a, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2, font=F_BOLD); put(ws, f'C{r}', c2); r += 1
header_row(ws, r, ['Результат', 'Формулировка заключения', '']); r += 1
for a, b2 in D['CONC15']:
    put(ws, f'A{r}', a, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2); r += 1
put(ws, f'A{r}', 'Правила: I_А < 40% → обязательная углублённая оценка; I_Ц < 40% → проверка альтернативного способа; любой индекс < 60% → нет безусловного «положительно». Результат краткого опросника — не решение о прекращении, а решение о необходимости углублённой проверки.', font=F_SMALL); r += 2
sec(ws, r, '4. Средний опросник: пороги, категории заключений (I–IV), формулировки 21.2'); r += 1
header_row(ws, r, ['Индекс', 'Интерпретация', 'Управленческое решение']); r += 1
for a, b2, c2 in D['POROG30']:
    put(ws, f'A{r}', a, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2, font=F_BOLD); put(ws, f'C{r}', c2); r += 1
header_row(ws, r, ['Категория', 'Условия', 'Управленческое решение']); r += 1
for a, b2, c2 in D['CAT30']:
    put(ws, f'A{r}', a, font=F_BOLD, fill=FILL_FORM); put(ws, f'B{r}', b2); put(ws, f'C{r}', c2); r += 1
header_row(ws, r, ['Результат (21.2)', 'Формулировка заключения', '']); r += 1
for a, b2 in D['CONC30']:
    put(ws, f'A{r}', a, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2); r += 1
put(ws, f'A{r}', 'Строгие условия: I_А < 55% → проверка потребности у заказчика; I_Ц < 55% → анализ альтернатив; уникальность < 40% → проверка дублирования; практическое применение < 40% → пересмотр плана внедрения; критический технический риск → отдельное заключение независимо от индекса.', font=F_SMALL); r += 2
sec(ws, r, '5. Полный опросник: решения комиссии (21.3) и варианты итоговых решений R1–R7'); r += 1
header_row(ws, r, ['Результат', 'Рекомендуемое решение', '']); r += 1
for a, b2 in D['CONC60']:
    put(ws, f'A{r}', a, font=F_BOLD, fill=FILL_FORM); put(ws, f'B{r}', b2); r += 1
header_row(ws, r, ['Код', 'Решение', 'Условия применения']); r += 1
for a, b2, c2 in D['R17']:
    put(ws, f'A{r}', a, font=F_BOLD, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2); put(ws, f'C{r}', c2); r += 1
put(ws, f'A{r}', 'R7 — рекомендация о рассмотрении вопроса, а не самостоятельное решение о расторжении государственного контракта или прекращении обязательств.', font=F_SMALL); r += 2
sec(ws, r, '6. AHP: шкала парных сравнений и случайные индексы RI'); r += 1
header_row(ws, r, ['Значение', 'Интерпретация', '']); r += 1
for a, b2 in D['AHPSCALE']:
    put(ws, f'A{r}', a, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2); r += 1
put(ws, f'A{r}', 'Расчёт: нормированный главный собственный вектор (Aw = λmax·w); CI = (λmax − n)/(n − 1); CR = CI/RI ≤ 0,10. Групповая оценка — геометрическое среднее индивидуальных матриц aᵢⱼ^group = (Π aᵢⱼ^(r))^(1/m); индивидуальные матрицы сохранять.', font=F_SMALL); r += 1
RI_ROW0 = r + 1
header_row(ws, r, ['n', 'RI (случайный индекс)', '']); r += 1
RI = [0, 0, 0.58, 0.90, 1.12, 1.24, 1.32, 1.41, 1.45, 1.49]
for n_, ri in enumerate(RI, start=1):
    put(ws, f'B{r}', n_, align=CTR, fill=FILL_FORM); put(ws, f'C{r}', ri, fmt='0.00', align=CTR); r += 1
RI_RANGE = f'$B${RI_ROW0}:$C${RI_ROW0+9}'
r += 1
sec(ws, r, '7. Стоп-факторы SF-0 / SF-1 / SF-2'); r += 1
header_row(ws, r, ['Код', 'Категория', 'Последствие']); r += 1
for a, b2, c2 in [['SF-0', 'Отсутствуют', 'Обычный порядок оценки'], ['SF-1', 'Существенная неопределённость', 'Обязательная дополнительная проверка'], ['SF-2', 'Критическое обстоятельство', 'Эскалация и запрет безусловно положительного заключения']]:
    put(ws, f'A{r}', a, font=F_BOLD, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2); put(ws, f'C{r}', c2); r += 1
r += 1
sec(ws, r, '8. Зрелость результата M0–M9'); r += 1
header_row(ws, r, ['Уровень', 'Состояние', 'Что должно быть подтверждено']); r += 1
for a, b2, c2 in D['M09']:
    put(ws, f'A{r}', a, font=F_BOLD, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2); put(ws, f'C{r}', c2); r += 1
r += 1
sec(ws, r, '9. Категории рекомендаций К1–К4 и статусы мероприятий'); r += 1
header_row(ws, r, ['Категория', 'Содержание', 'Срок реакции']); r += 1
for a, b2, c2 in D['K14']:
    put(ws, f'A{r}', a, font=F_BOLD, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2); put(ws, f'C{r}', c2); r += 1
header_row(ws, r, ['Статус', 'Значение', '']); r += 1
for a, b2 in D['STAT']:
    put(ws, f'A{r}', a, font=F_BOLD, fill=FILL_FORM); put(ws, f'B{r}', b2); r += 1
r += 1
sec(ws, r, '10. BCR: интерпретация'); r += 1
header_row(ws, r, ['BCR', 'Интерпретация', '']); r += 1
for a, b2 in D['BCR']:
    put(ws, f'A{r}', a, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', b2); r += 1
r += 1
sec(ws, r, '11. Правила расчёта (НД / НП / Coverage / индексы)'); r += 1
RULES_FULL = [
    'Индекс направления: I = Σ(wⱼ·xⱼ) / (5·Σwⱼ) × 100% — только по вопросам с числовой оценкой (НД и НП исключаются из числителя и знаменателя).',
    'НД — недостаточно данных: исключается из расчёта индекса, НО уменьшает полноту доказательств Coverage; аудитор запрашивает подтверждение; при непредставлении в срок — оценка 0 либо отдельный штраф за недостаточность доказательной базы.',
    'НП — неприменимо: исключается из расчёта индекса И из знаменателя Coverage, только при наличии письменного обоснования (согласуется председателем комиссии).',
    'Coverage = Σwⱼ (вопросы с числовой оценкой) / Σwⱼ (все применимые вопросы) × 100%. Пороги: краткий ≥ 80%, средний ≥ 85%, полный ≥ 90%. Ниже порога — «недостаточно данных для достоверной оценки», положительное заключение не выносится.',
    'Комплексный индекс I_компл = ∛(I_А · I_Ц · I_Р) — вспомогательный (геометрическое среднее не допускает чрезмерной компенсации); НЕ заменяет три самостоятельных индекса и не отменяет стоп-факторы.',
    'Полный опросник — двухуровневые веса: Dk по блокам (вопросы с собственными коэффициентами значимости), затем I_А = 0,20·DА1+0,20·DА2+0,20·DА3+0,40·DА4; I_Ц = 0,25·(DБ1+DБ2+DБ3+DБ4); I_Р = 0,30·DВ1+0,25·DВ2+0,30·DВ3+0,15·DВ4 (базовая конфигурация; может быть заменена весами AHP).',
    'Эксперт не вправе выставлять 4–5 баллов только на основании пояснений руководителя НИОКР: требуется документальное, расчётное, экспериментальное или иное проверяемое подтверждение. Баллы 0–2 — обязательный комментарий и фиксация недостатка.',
    'Итоговый балл вопроса = взвешенное среднее индивидуальных оценок xⱼ* = Σrᵢxᵢⱼ/Σrᵢ (rᵢ — коэффициент компетентности, задаётся до оценивания). Существенные расхождения (разброс ≥ 2) выносятся на коллегиальное обсуждение; при сохранении разногласий — протокол и особое мнение.',
    'Динамическая оценка: ΔI = I(t) − I(t−1); скорость V = (I(t) − I(0))/t. Обязательные контрольные точки переоценки: открытие НИОКР, завершение аванпроекта, технического проекта, основных этапов ОКР, изготовление опытного образца, предварительные и государственные испытания, установочная партия, решение о серийном производстве.',
]
for t in RULES_FULL:
    ws[f'A{r}'] = '• ' + t; ws[f'A{r}'].font = F_SMALL; ws[f'A{r}'].alignment = WRAP
    ws.merge_cells(f'A{r}:H{r+1}'); r += 2
r += 1
sec(ws, r, '12. Раннее предупреждение (триггеры)'); r += 1
header_row(ws, r, ['Событие', 'Индикатор', 'Действие']); r += 1
for a, b2, c2 in D['EARLY']:
    put(ws, f'A{r}', a, font=F_BOLD, fill=FILL_FORM); put(ws, f'B{r}', b2, align=CTR); put(ws, f'C{r}', c2); r += 1
put(ws, f'A{r}', 'Порог снижения на 10 пунктов — предварительная настройка; калибруется по фактической динамике проектов.', font=F_SMALL); r += 2
sec(ws, r, '13. Контрольные примеры для приёмки калькулятора (воспроизводятся скриптом scripts/verify_9_xlsx.py)'); r += 1
header_row(ws, r, ['№', 'Контрольный пример', 'Ожидаемое поведение']); r += 1
for i, (a, b2) in enumerate(D['CTRL']):
    put(ws, f'A{r}', i+1, align=CTR, fill=FILL_FORM); put(ws, f'B{r}', a); put(ws, f'C{r}', b2); r += 1
put(ws, f'A{r}', 'Дополнительно проверяются: граничные значения Coverage, нулевые знаменатели, пустые ячейки, некорректные оценки, изменение состава применимых вопросов. Журнал изменений оценок: колонка «Изменение (кто/дата)» в опросниках; полноценный журнал — в интерактивной модели v18.', font=F_SMALL)
ws.merge_cells(f'A{r}:H{r+1}'); ws[f'A{r}'].alignment = WRAP
widths(ws, {'A': 10, 'B': 34, 'C': 60, 'D': 20, 'E': 14, 'F': 14, 'G': 14, 'H': 14})

# ---------- порядок листов по спецификации ----------
order = ['Панель_руководителя', 'Паспорт_НИОКР', 'Краткий_15', 'Средний_30', 'Полный_60', 'Доказательства',
         'Эксперты', 'AHP', 'Стоп_факторы', 'Зрелость_M0_M9', 'Экономика_LCC', 'План_мероприятий',
         'Заключение', 'Справочники']
wb._sheets = [wb[n] for n in order]
wb.active = 0

# ---------- замена адреса лучшей альтернативы в формуле Полный_60 (C48 → C53) ----------
w60 = wb['Полный_60']
old = w60['C102'].value
w60['C102'] = old.replace('Экономика_LCC!$C$48', 'Экономика_LCC!$C$53')

# ---------- патч фактического RI-диапазона в формулах AHP (VLOOKUP) ----------
ws_a = wb['AHP']
for cell in ['C12', 'C25', 'C37', 'C49']:
    v = ws_a[cell].value
    assert '$B$82:$C$91' in v, (cell, v)
    ws_a[cell] = v.replace('$B$82:$C$91', RI_RANGE)
RI_FULL = 'Справочники!' + RI_RANGE

wb.save(OUT)

# ---------- дамп ключевых адресов для верификации ----------
cells = dict(
    PASS=PASS, K15=K15C, C30=C30C, C60=C60C, ROW30=ROW30, BLK_ROW=BLK_ROW,
    AHP=dict(cr=CR_CELLS, w=AHP_W_CELLS, cr_max='C53', cr_chk='C54'),
    SF=SF_C, MAT=MAT_C, ECON=ECON_C, PLAN=PLAN_C, ZAKL=ZAKL_C,
    RI_RANGE=RI_FULL,
)
io.open(CELLS, 'w', encoding='utf-8').write(json.dumps(cells, ensure_ascii=False, indent=1))
import os
print('GEN9 OK →', OUT, os.path.getsize(OUT), 'bytes; cells →', CELLS)
