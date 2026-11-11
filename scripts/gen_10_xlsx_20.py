#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Этап 20 (v20.0), файл 10. Excel-калькулятор и подробный чек-лист методики
ЭО-НИОКР v1.0 (22 вопроса, единая шкала 1–5 + Н/О, 5 стоп-факторов, 9 «красных»,
два индекса I_A/I_C, матрица 3×3, σ-согласованность, оверрайд стоп-фактора).
Группа до 10 экспертов. 16 листов: Методика · Чек-лист_22 · Э1–Э10 · Свод · Панель · Заключение.
Вся логика — формулами, без макросов (Excel 365/2021)."""
import json, io, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.chart import BarChart, Reference

HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(io.open(os.path.join(HERE, 'data', 'qeo22.json'), encoding='utf-8'))
OUT = os.path.join(HERE, '..', 'output', '10_Чек-лист_ЭО_НИОКР_22_вопроса_v20.xlsx')
NEXP = 10

NAVY = '1F3864'; ACC = '2E74B5'; LGRAY = 'E6EBF4'; YELL = 'FFF2CC'
GREEN = 'C6EFCE'; REDF = 'FFC7CE'; ORANG = 'FFEB9C'
F_TITLE = Font(bold=True, size=14, color='FFFFFF')
F_H = Font(bold=True, size=11, color='FFFFFF')
F_SEC = Font(bold=True, size=11, color=NAVY)
F_B = Font(bold=True)
FILL_T = PatternFill('solid', fgColor=NAVY)
FILL_H = PatternFill('solid', fgColor=ACC)
FILL_S = PatternFill('solid', fgColor=LGRAY)
FILL_IN = PatternFill('solid', fgColor=YELL)
WRAP = Alignment(wrap_text=True, vertical='top')
WRAPC = Alignment(wrap_text=True, vertical='center', horizontal='center')
THIN = Side(style='thin', color='B9C4D9')
BORD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

QA = [q for q in D['q'] if q['dir'] == 'A']
QC = [q for q in D['q'] if q['dir'] == 'C']
SF, KV = D['sf'], D['kv']
MX = D['matrix']

wb = openpyxl.Workbook()

def hdr(ws, row, ncol, text):
    ws.cell(row=row, column=1, value=text).font = F_H
    for c in range(1, ncol + 1):
        ws.cell(row=row, column=c).fill = FILL_H
    ws.row_dimensions[row].height = 18

def title(ws, text, ncol=8):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(row=1, column=1, value=text)
    c.font = F_TITLE; c.fill = FILL_T; c.alignment = Alignment(vertical='center')
    ws.row_dimensions[1].height = 26

# ============================ ЛИСТ «Методика» ============================
ws = wb.active; ws.title = 'Методика'
title(ws, 'МЕТОДИКА ЭО-НИОКР v1.0 — экспертная оценка актуальности и целесообразности НИОКР (модель v20.0, группа до 10 экспертов, 11.2026)')
ws['A3'] = 'Назначение: единый порядок экспертной оценки НИОКР. Оценка — профессиональное суждение; документальное подтверждение баллов НЕ требуется.'
ws['A3'].font = F_SEC
r = 5
hdr(ws, r, 3, '1. Единая пятибалльная шкала (все 22 вопроса)'); r += 1
for c, t in zip('ABC', ('Балл', 'Уровень', 'Единое пояснение для эксперта')):
    ws[c + str(r)] = t; ws[c + str(r)].font = F_B; ws[c + str(r)].fill = FILL_S
r += 1
for b, nm, tx in D['scale']:
    ws.cell(row=r, column=1, value=b).alignment = WRAPC
    ws.cell(row=r, column=2, value=nm)
    ws.cell(row=r, column=3, value=tx).alignment = WRAP
    r += 1
ws.cell(row=r, column=1, value='Н/О').alignment = WRAPC
ws.cell(row=r, column=2, value='Не оценивается')
ws.cell(row=r, column=3, value='Эксперт не обладает достаточными знаниями или компетенциями для суждения по вопросу. Н/О не является баллом и не включается в расчёт.').alignment = WRAP
r += 2
hdr(ws, r, 3, '2. Правила экспертного оценивания'); r += 1
for i, rule in enumerate(D['rules'], 1):
    ws.cell(row=r, column=1, value=i).alignment = WRAPC
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2, value=rule).alignment = WRAP
    r += 1
r += 1
hdr(ws, r, 3, '3. Минимальная полнота индивидуальной оценки'); r += 1
for c, t in zip('ABC', ('Направление', 'Всего вопросов', 'Минимум оценённых (допуск Н/О ≤ 2)')):
    ws[c + str(r)] = t; ws[c + str(r)].font = F_B; ws[c + str(r)].fill = FILL_S
r += 1
for k, nm in (('A', 'Актуальность'), ('C', 'Целесообразность')):
    mf = D['minfull'][k]
    ws.cell(row=r, column=1, value=nm); ws.cell(row=r, column=2, value=mf['total'])
    ws.cell(row=r, column=3, value=mf['min']); r += 1
ws.cell(row=r, column=1, value='Если оценённых вопросов меньше минимума — индекс направления не рассчитывается (привлечь эксперта / повторная оценка).').alignment = WRAP
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3); r += 2
hdr(ws, r, 3, '4. Расчёт индексов'); r += 1
for t in ('Экспертная группа: от 3 до 10 специалистов (рекомендуется 5); веса экспертов равные.',
          'Оценка проводится в два раунда: раунд 1 — независимое первичное заполнение (фиксируется), раунд 2 — возможные правки после обсуждения комиссии; исходные оценки раунда 1 сохраняются.',
          'I_A = ΣA_i / N_A (по 10 вопросам актуальности, Н/О исключены из знаменателя); диапазон 1,00–5,00.',
          'I_C = ΣC_j / N_C (по 12 вопросам целесообразности); диапазон 1,00–5,00.',
          'Групповой индекс = простое среднее индивидуальных индексов экспертов с действующей оценкой направления (веса равные).',
          'σ = √(Σ(I_k − I_ср)² / n) — популяционное стандартное отклонение (группа рассматривается как полная совокупность оценивающих).',
          'По каждому вопросу: среднее, min, max, разброс, σ; σ > 1,00 — флаг «РАЗНОГЛАСИЕ», обязательное обсуждение при сохранении исходных оценок.'):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=1, value='• ' + t).alignment = WRAP; r += 1
r += 1
hdr(ws, r, 3, '5. Уровни (зоны) индексов и матрица решений 3×3'); r += 1
for z in D['zones']:
    ws.cell(row=r, column=1, value=z['k']).alignment = WRAPC
    ws.cell(row=r, column=2, value=z['name'])
    ws.cell(row=r, column=3, value='%.2f – %.2f' % (z['lo'], z['hi'])).alignment = WRAPC
    r += 1
r += 1
ws.cell(row=r, column=1, value='I_A ↓ \\ I_C →').font = F_B
for j, zc in enumerate(('В', 'С', 'Н')):
    ws.cell(row=r, column=2 + j, value='Целесообразность: ' + zc).font = F_B
    ws.cell(row=r, column=2 + j).fill = FILL_S
r += 1
for za in ('В', 'С', 'Н'):
    ws.cell(row=r, column=1, value='Актуальность: ' + za).font = F_B
    ws.cell(row=r, column=1).fill = FILL_S
    for j, zc in enumerate(('В', 'С', 'Н')):
        ws.cell(row=r, column=2 + j, value=MX[za][zc]).alignment = WRAP
    r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws.cell(row=r, column=1, value='Оверрайд: при выявленном стоп-факторе рекомендация «Продолжить НИОКР» автоматически понижается до «Пересмотреть параметры НИОКР (стоп-фактор)». Итоговое решение принимает комиссия.').alignment = WRAP
r += 2
hdr(ws, r, 3, '6. Согласованность мнений (σ по индексу)'); r += 1
for lo, nm in ((0.50, 'до 0,50 — высокая согласованность'), (0.80, '0,51–0,80 — приемлемая'), (1.00, '0,81–1,00 — умеренные разногласия'), (99, 'более 1,00 — существенные разногласия: обсуждение, исходные оценки сохраняются')):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=1, value='• ' + nm); r += 1
r += 1
hdr(ws, r, 3, '7. Стоп-факторы (5) — ответ Да / Нет / Неопределённо; в индексы не входят'); r += 1
for i, s in enumerate(SF, 1):
    ws.cell(row=r, column=1, value=i).alignment = WRAPC
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2, value=s).alignment = WRAP; r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
ws.cell(row=r, column=1, value='Фактор выявлен, если хотя бы один эксперт ответил «Да»; неопределён, если «Да» нет, но есть «Неопределённо». Выявленный стоп-фактор обязателен к рассмотрению комиссией и блокирует автоматическую рекомендацию «Продолжить НИОКР».').alignment = WRAP
r += 2
hdr(ws, r, 3, '8. «Красные» вопросы (9) — шкала 1=Да … 5=Нет; в индексы не входят'); r += 1
for k in KV:
    ws.cell(row=r, column=1, value=k['n']).alignment = WRAPC
    ws.cell(row=r, column=2, value=k['q']).alignment = WRAP
    ws.cell(row=r, column=3, value=('негативный: сигнал = балл 1–2' if k['dir'] == 'neg' else 'позитивный: сигнал = балл 4–5')).alignment = WRAP
    r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
ws.cell(row=r, column=1, value='Групповой сигнал фиксируется при ≥ 2 экспертах с сигналом; единичный сигнал — особое мнение, в групповой сигнал не входит.').alignment = WRAP
ws.column_dimensions['A'].width = 16
ws.column_dimensions['B'].width = 46
ws.column_dimensions['C'].width = 68
ws.column_dimensions['D'].width = 34
ws.sheet_view.showGridLines = False

# ============================ ЛИСТ «Чек-лист_22» ============================
ws = wb.create_sheet('Чек-лист_22')
title(ws, 'ПОДРОБНЫЙ ЧЕК-ЛИСТ ЭКСПЕРТНОЙ ОЦЕНКИ НИОКР — 22 вопроса с полными якорями шкалы 1–5 (ЭО-НИОКР v1.0)', 11)
ws['A2'] = 'Баллы 1–5 по единой шкале (лист «Методика»). Якоря — опорные описания состояний признака; 2 и 4 выводятся из общих определений шкалы (принцип осторожности: при колебании — более низкий балл).'
ws['A2'].alignment = WRAP; ws.merge_cells('A2:K2')
H = ['Код', 'Направление', '№', 'Критерий', 'Вопрос для экспертной оценки',
     '1 — крайне низкая', '2 — низкая', '3 — средняя', '4 — высокая', '5 — очень высокая',
     'Примечание по разграничению']
for j, t in enumerate(H, 1):
    c = ws.cell(row=4, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
ws.row_dimensions[4].height = 30
r = 5
for q in D['q']:
    vals = [q['code'], 'Актуальность' if q['dir'] == 'A' else 'Целесообразность', q['n'],
            q['crit'], q['q'], q['a']['1'], q['a']['2'], q['a']['3'], q['a']['4'], q['a']['5'],
            q.get('note', '')]
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=j, value=v); c.alignment = WRAP; c.border = BORD
        if j <= 3: c.alignment = WRAPC
    if q['dir'] == 'C' and q['n'] == 1:
        for j in range(1, 12): ws.cell(row=r, column=j).fill = FILL_S
    ws.row_dimensions[r].height = 74
    r += 1
widths = [7, 15, 5, 24, 40, 34, 34, 34, 34, 34, 40]
for j, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = 'A5'
ws.sheet_view.showGridLines = False

# ============================ ЛИСТЫ Э1–Э10 ===========================
EXP_ROWS_A = range(9, 19)     # А1..А10 (строки 9-18)
EXP_ROWS_C = range(21, 33)    # Ц1..Ц12 (строки 21-32)
SF_ROW0, KV_ROW0 = 45, 53     # СФ1..СФ5 (45-49), КВ1..КВ9 (53-61)
for e in range(1, NEXP + 1):
    ws = wb.create_sheet('Э%d' % e)
    title(ws, 'АНКЕТА ЭКСПЕРТНОЙ ОЦЕНКИ НИОКР — эксперт %d (ЭО-НИОКР v1.0, 22 вопроса)' % e, 6)
    ws['A3'] = 'Наименование НИОКР:'; ws['B3'] = ''
    ws['A4'] = 'Шифр:'; ws['B4'] = ''
    ws['A5'] = 'Эксперт (Ф.И.О.):'; ws['B5'] = ''
    ws['C5'] = 'Компетенция:'; ws['D5'] = ''
    ws['E5'] = 'Дата:'; ws['F5'] = ''
    for cc in ('B3', 'B4', 'B5', 'D5', 'F5'):
        ws[cc].fill = FILL_IN
    ws['A3'].font = F_B; ws['A4'].font = F_B; ws['A5'].font = F_B; ws['C5'].font = F_B; ws['E5'].font = F_B
    ws.cell(row=7, column=1, value='РАЗДЕЛ А. АКТУАЛЬНОСТЬ — 10 вопросов (макс. сумма 50)').font = F_SEC
    for j, t in enumerate(('№', 'Критерий', 'Вопрос эксперту', 'Оценка 1–5 / Н/О', 'Комментарий (необязательно)'), 1):
        c = ws.cell(row=8, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
    for i, q in enumerate(QA):
        rr = EXP_ROWS_A.start + i
        ws.cell(row=rr, column=1, value=q['n']).alignment = WRAPC
        ws.cell(row=rr, column=2, value=q['crit']).alignment = WRAP
        ws.cell(row=rr, column=3, value=q['q']).alignment = WRAP
        ws.cell(row=rr, column=4).fill = FILL_IN
        ws.cell(row=rr, column=4).alignment = WRAPC
        ws.cell(row=rr, column=5).alignment = WRAP
        ws.row_dimensions[rr].height = 30
    ws.cell(row=19, column=1, value='РАЗДЕЛ Б. ЦЕЛЕСООБРАЗНОСТЬ — 12 вопросов (макс. сумма 60)').font = F_SEC
    for j, t in enumerate(('№', 'Критерий', 'Вопрос эксперту', 'Оценка 1–5 / Н/О', 'Комментарий (необязательно)'), 1):
        c = ws.cell(row=20, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
    for i, q in enumerate(QC):
        rr = EXP_ROWS_C.start + i
        ws.cell(row=rr, column=1, value=q['n']).alignment = WRAPC
        ws.cell(row=rr, column=2, value=q['crit']).alignment = WRAP
        ws.cell(row=rr, column=3, value=q['q']).alignment = WRAP
        ws.cell(row=rr, column=4).fill = FILL_IN
        ws.cell(row=rr, column=4).alignment = WRAPC
        ws.cell(row=rr, column=5).alignment = WRAP
        ws.row_dimensions[rr].height = 30
    # служебный блок
    ws['A34'] = 'I_A — индекс актуальности'; ws['A34'].font = F_B
    ws['B34'] = '=COUNT(D9:D18)'; ws['C34'] = '=SUM(D9:D18)'
    ws['D34'] = '=IF(B34<8,"НЕ РАССЧИТАН",C34/B34)'; ws['D34'].number_format = '0.00'
    ws['A35'] = 'I_C — индекс целесообразности'; ws['A35'].font = F_B
    ws['B35'] = '=COUNT(D21:D32)'; ws['C35'] = '=SUM(D21:D32)'
    ws['D35'] = '=IF(B35<10,"НЕ РАССЧИТАН",C35/B35)'; ws['D35'].number_format = '0.00'
    ws['A37'] = 'Число Н/О по актуальности (допуск ≤ 2):'; ws['B37'] = '=COUNTIF(D9:D18,"Н/О")'
    ws['A38'] = 'Число Н/О по целесообразности (допуск ≤ 2):'; ws['B38'] = '=COUNTIF(D21:D32,"Н/О")'
    ws['A40'] = 'Предупреждение:'; ws['A40'].font = F_B
    ws['B40'] = ('=IF(AND(B37>2,B38>2),"ВНИМАНИЕ: оба направления неполны (Н/О > 2)",'
                 'IF(OR(B37>2,B38>2),"ВНИМАНИЕ: Н/О > 2 в одном направлении — индекс по нему не рассчитывается",""))')
    ws['A41'] = 'Статус анкеты:'; ws['A41'].font = F_B
    ws['B41'] = ('=IF(COUNT(D9:D32)+COUNTIF(D9:D32,"Н/О")=0,"НЕ ЗАПОЛНЕНА",'
                 'IF(COUNT(D9:D32)+COUNTIF(D9:D32,"Н/О")<22,"ЗАПОЛНЕНИЕ НЕ ЗАВЕРШЕНО","ЗАВЕРШЕНА"))')
    # стоп-факторы
    ws.cell(row=43, column=1, value='СТОП-ФАКТОРЫ (оцениваются отдельно от 22 вопросов; ответ: Да / Нет / Неопределённо)').font = F_SEC
    for j, t in enumerate(('№', 'Стоп-фактор', 'Ответ', 'Комментарий'), 1):
        c = ws.cell(row=44, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
    for i, s in enumerate(SF):
        rr = SF_ROW0 + i
        ws.cell(row=rr, column=1, value=i + 1).alignment = WRAPC
        ws.cell(row=rr, column=2, value=s).alignment = WRAP
        ws.cell(row=rr, column=3).fill = FILL_IN; ws.cell(row=rr, column=3).alignment = WRAPC
        ws.cell(row=rr, column=4).alignment = WRAP
        ws.row_dimensions[rr].height = 28
    # красные вопросы
    ws.cell(row=51, column=1, value='«КРАСНЫЕ» ВОПРОСЫ (шкала: 1=Да, 2=Возможно да, 3=Неопределённо, 4=Возможно нет, 5=Нет)').font = F_SEC
    for j, t in enumerate(('№', 'Вопрос', 'Балл', 'Направление', 'Сигнал риска', 'Комментарий'), 1):
        c = ws.cell(row=52, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
    for i, k in enumerate(KV):
        rr = KV_ROW0 + i
        ws.cell(row=rr, column=1, value=k['n']).alignment = WRAPC
        ws.cell(row=rr, column=2, value=k['q']).alignment = WRAP
        ws.cell(row=rr, column=3).fill = FILL_IN; ws.cell(row=rr, column=3).alignment = WRAPC
        ws.cell(row=rr, column=4, value='Негативный' if k['dir'] == 'neg' else 'Позитивный').alignment = WRAPC
        ws.cell(row=rr, column=5, value='=IF(C%d="","",IF(D%d="Негативный",IF(C%d<=2,"СИГНАЛ",""),IF(C%d>=4,"СИГНАЛ","")))' % (rr, rr, rr, rr)).alignment = WRAPC
        ws.cell(row=rr, column=6).alignment = WRAP
        ws.row_dimensions[rr].height = 26
    # валидации
    dv = DataValidation(type='list', formula1='"1,2,3,4,5,Н/О"', allow_blank=True,
                        showErrorMessage=True, errorTitle='Недопустимое значение',
                        error='Допустимо: 1, 2, 3, 4, 5 или Н/О')
    ws.add_data_validation(dv)
    dv.add('D9:D18'); dv.add('D21:D32')
    dv2 = DataValidation(type='list', formula1='"Да,Нет,Неопределённо"', allow_blank=True,
                         showErrorMessage=True, errorTitle='Недопустимое значение',
                         error='Допустимо: Да, Нет, Неопределённо')
    ws.add_data_validation(dv2); dv2.add('C45:C49')
    dv3 = DataValidation(type='list', formula1='"1,2,3,4,5"', allow_blank=True,
                         showErrorMessage=True, errorTitle='Недопустимое значение',
                         error='Допустимо: 1, 2, 3, 4, 5')
    ws.add_data_validation(dv3); dv3.add('C53:C61')
    for j, w in enumerate((6, 30, 62, 16, 30, 24), 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'A9'
    ws.sheet_view.showGridLines = False

# ============================ ЛИСТ «Свод» ============================
ws = wb.create_sheet('Свод')
title(ws, 'СВОДНАЯ ЭКСПЕРТНАЯ ОЦЕНКА — расчётное ядро (групповые индексы, σ, сигналы, матрица; до 10 экспертов)', 22)
E0 = 5                    # строки экспертов 5..14
GRP = E0 + NEXP           # 15 — групповое значение
SIG = GRP + 1             # 16 — σ
CON = SIG + 1             # 17 — согласованность
QH = 19                   # заголовок раздела 2
QHEAD = 20                # шапка таблицы вопросов
Q0 = 21                   # вопросы 21..42
QCNT = Q0 + 22 + 1        # 44 — счётчик разногласий
SFH = 46; SFHEAD = 47; SF0 = 48   # СФ 48..52
SFF = SF0 + 5 + 1         # 54 — флаг
KVH = 56; KVHEAD = 57; KV0 = 58   # КВ 58..66
KVL = KV0 + 9 + 1         # 68 — перечень
MXH = 70; ZA = 71; ZC = 72; MXLAB = 74; MX0 = 75  # матрица 75..77
REC = 79; OVR = 80; LOWR = 82; DISR = 83
EXPL = lambda c, r1, r2: c + str(r1) + ':' + c + str(r2)
hdr(ws, 3, 4, '1. Индивидуальные и групповые индексы (до 10 экспертов)')
for j, t in enumerate(('Эксперт', 'I_A (актуальность)', 'I_C (целесообразность)', 'Статус'), 1):
    c = ws.cell(row=4, column=j, value=t); c.font = F_B; c.fill = FILL_S
for e in range(1, NEXP + 1):
    rr = E0 + e - 1
    ws.cell(row=rr, column=1, value='Эксперт %d (Э%d)' % (e, e))
    ws.cell(row=rr, column=2, value="=Э%d!D34" % e).number_format = '0.00'
    ws.cell(row=rr, column=3, value="=Э%d!D35" % e).number_format = '0.00'
    ws.cell(row=rr, column=4, value=('=IF(Э%d!B41="НЕ ЗАПОЛНЕНА","НЕ ЗАПОЛНЕНА",'
        'IF(AND(NOT(ISNUMBER(Э%d!D34)),NOT(ISNUMBER(Э%d!D35))),"ИНДЕКСЫ НЕ РАССЧИТАНЫ","ОК"))' % (e, e, e)))
ws.cell(row=GRP, column=1, value='ГРУППОВОЕ ЗНАЧЕНИЕ (среднее)').font = F_B
ws.cell(row=GRP, column=2, value='=IFERROR(AVERAGE(B%d:B%d),"н/д")' % (E0, GRP - 1)).number_format = '0.00'
ws.cell(row=GRP, column=3, value='=IFERROR(AVERAGE(C%d:C%d),"н/д")' % (E0, GRP - 1)).number_format = '0.00'
ws.cell(row=GRP, column=2).font = F_B; ws.cell(row=GRP, column=3).font = F_B
ws.cell(row=GRP, column=4, value='среднее по экспертам с действующим индексом')
ws.cell(row=SIG, column=1, value='σ (стандартное отклонение, популяционное)').font = F_B
ws.cell(row=SIG, column=2, value='=IF(COUNT(B%d:B%d)<2,"н/д",STDEVP(B%d:B%d))' % (E0, GRP - 1, E0, GRP - 1)).number_format = '0.00'
ws.cell(row=SIG, column=3, value='=IF(COUNT(C%d:C%d)<2,"н/д",STDEVP(C%d:C%d))' % (E0, GRP - 1, E0, GRP - 1)).number_format = '0.00'
ws.cell(row=CON, column=1, value='Согласованность мнений').font = F_B
for col in 'BC':
    ws[col + str(CON)] = ('=IF(NOT(ISNUMBER(%s%d)),"н/д",IF(%s%d<=0.5,"Высокая",IF(%s%d<=0.8,"Приемлемая",'
                          'IF(%s%d<=1,"Умеренные разногласия","Существенные разногласия"))))' % (col, SIG, col, SIG, col, SIG, col, SIG))
hdr(ws, QH, 22, '2. Повопросная сводка (22 вопроса): оценки экспертов, среднее, разброс, σ, флаги')
HEAD = ['Код'] + ['Э%d' % e for e in range(1, NEXP + 1)] + ['Среднее', 'Min', 'Max', 'Разброс', 'σ',
        'Флаг разногласия (σ>1)', 'Низкая оценка (≤2)', 'Низкие (помощ.)', 'Разногласия (помощ.)', 'Критерий', 'Вопрос']
for j, t in enumerate(HEAD, 1):
    c = ws.cell(row=QHEAD, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
CL = get_column_letter(1 + NEXP)        # K — последний эксперт
MEAN, MIN_, MAX_, RNG, SD, FLG, LOW = 2 + NEXP, 3 + NEXP, 4 + NEXP, 5 + NEXP, 6 + NEXP, 7 + NEXP, 8 + NEXP
LOWH, DISH, CRIT, QUES = 9 + NEXP, 10 + NEXP, 11 + NEXP, 12 + NEXP
LC = lambda c, rr: get_column_letter(c) + str(rr)
for i, q in enumerate(D['q']):
    rr = Q0 + i
    erow = (9 + i) if i < 10 else (21 + (i - 10))
    ws.cell(row=rr, column=1, value=q['code']).alignment = WRAPC
    for e in range(1, NEXP + 1):
        ws.cell(row=rr, column=1 + e, value='=IF(ISNUMBER(Э%d!D%d),Э%d!D%d,"")' % (e, erow, e, erow)).alignment = WRAPC
    rng = 'B%d:%s%d' % (rr, CL, rr)
    ws.cell(row=rr, column=MEAN, value='=IF(COUNT(%s)=0,"н/д",AVERAGE(%s))' % (rng, rng)).number_format = '0.00'
    ws.cell(row=rr, column=MIN_, value='=IF(COUNT(%s)=0,"—",MIN(%s))' % (rng, rng)).alignment = WRAPC
    ws.cell(row=rr, column=MAX_, value='=IF(COUNT(%s)=0,"—",MAX(%s))' % (rng, rng)).alignment = WRAPC
    ws.cell(row=rr, column=RNG, value='=IF(AND(ISNUMBER(%s),ISNUMBER(%s)),%s-%s,"—")' % (LC(MIN_, rr), LC(MAX_, rr), LC(MAX_, rr), LC(MIN_, rr))).alignment = WRAPC
    ws.cell(row=rr, column=SD, value='=IF(COUNT(%s)<2,"—",STDEVP(%s))' % (rng, rng)).number_format = '0.00'
    ws.cell(row=rr, column=FLG, value='=IF(AND(ISNUMBER(%s),%s>1),"РАЗНОГЛАСИЕ","")' % (LC(SD, rr), LC(SD, rr))).alignment = WRAPC
    ws.cell(row=rr, column=LOW, value='=IF(AND(ISNUMBER(%s),%s<=2),"НИЗКАЯ","")' % (LC(MEAN, rr), LC(MEAN, rr))).alignment = WRAPC
    ws.cell(row=rr, column=LOWH, value='=IF(%s="НИЗКАЯ",A%d&" ","")' % (LC(LOW, rr), rr))
    ws.cell(row=rr, column=DISH, value='=IF(%s="РАЗНОГЛАСИЕ",A%d&" ","")' % (LC(FLG, rr), rr))
    ws.cell(row=rr, column=CRIT, value=q['crit']).alignment = WRAP
    ws.cell(row=rr, column=QUES, value=q['q']).alignment = WRAP
ws.cell(row=QCNT, column=1, value='Вопросов с разногласиями (σ>1,00):').font = F_B
ws.cell(row=QCNT, column=2, value='=COUNTIF(%s,"РАЗНОГЛАСИЕ")' % EXPL(LC(FLG, 0)[:-1], Q0, Q0 + 21))
ws.cell(row=QCNT, column=3, value='=IF(B%d=0,"обсуждение не требуется","ОБСУЖДЕНИЕ ОБЯЗАТЕЛЬНО (сохраняя исходные оценки)")' % QCNT)
hdr(ws, SFH, 7, '3. Стоп-факторы — групповая обработка')
for j, t in enumerate(('№', 'Стоп-фактор', '«Да»', '«Неопр.»', '«Нет»', 'Статус', 'Для заключения'), 1):
    c = ws.cell(row=SFHEAD, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
for i in range(5):
    rr = SF0 + i
    er = SF_ROW0 + i
    ws.cell(row=rr, column=1, value=i + 1).alignment = WRAPC
    ws.cell(row=rr, column=2, value=SF[i]).alignment = WRAP
    ws.cell(row=rr, column=3, value='=' + '+'.join('COUNTIF(Э%d!C%d,"Да")' % (e, er) for e in range(1, NEXP + 1))).alignment = WRAPC
    ws.cell(row=rr, column=4, value='=' + '+'.join('COUNTIF(Э%d!C%d,"Неопределённо")' % (e, er) for e in range(1, NEXP + 1))).alignment = WRAPC
    ws.cell(row=rr, column=5, value='=' + '+'.join('COUNTIF(Э%d!C%d,"Нет")' % (e, er) for e in range(1, NEXP + 1))).alignment = WRAPC
    ws.cell(row=rr, column=6, value='=IF(C%d>0,"ВЫЯВЛЕН",IF(D%d>0,"НЕОПРЕДЕЛЁН","не выявлен"))' % (rr, rr)).alignment = WRAPC
    ws.cell(row=rr, column=7, value='=IF(F%d="ВЫЯВЛЕН","СФ%d; ","")' % (rr, i + 1))
ws.cell(row=SFF, column=1, value='Флаг стоп-факторов:').font = F_B
ws.cell(row=SFF, column=2, value='=IF(COUNTIF(F%d:F%d,"ВЫЯВЛЕН")>0,"СТОП-ФАКТОР ВЫЯВЛЕН",IF(COUNTIF(F%d:F%d,"НЕОПРЕДЕЛЁН")>0,"ЕСТЬ НЕОПРЕДЕЛЁННЫЕ ФАКТОРЫ","СТОП-ФАКТОРЫ НЕ ВЫЯВЛЕНЫ"))' % (SF0, SF0 + 4, SF0, SF0 + 4))
ws.cell(row=SFF, column=2).font = F_B
hdr(ws, KVH, 6, '4. «Красные» вопросы — групповые сигналы (сигнал при ≥ 2 экспертах)')
for j, t in enumerate(('№', 'Вопрос', 'Число сигналов', 'Групповой сигнал', 'Направление', 'Для заключения'), 1):
    c = ws.cell(row=KVHEAD, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
for i, k in enumerate(KV):
    rr = KV0 + i
    er = KV_ROW0 + i
    ws.cell(row=rr, column=1, value=k['n']).alignment = WRAPC
    ws.cell(row=rr, column=2, value=k['q']).alignment = WRAP
    ws.cell(row=rr, column=3, value='=' + '+'.join('COUNTIF(Э%d!E%d,"СИГНАЛ")' % (e, er) for e in range(1, NEXP + 1))).alignment = WRAPC
    ws.cell(row=rr, column=4, value='=IF(C%d>=2,"ГРУППОВОЙ СИГНАЛ",IF(C%d=1,"единичный сигнал",""))' % (rr, rr)).alignment = WRAPC
    ws.cell(row=rr, column=5, value='Негативный' if k['dir'] == 'neg' else 'Позитивный').alignment = WRAPC
    ws.cell(row=rr, column=6, value='=IF(D%d="ГРУППОВОЙ СИГНАЛ","КВ%d; ","")' % (rr, k['n']))
ws.cell(row=KVL, column=1, value='Перечень групповых «красных» сигналов:').font = F_B
kvcat = '&'.join('F%d' % (KV0 + i) for i in range(9))
ws.cell(row=KVL, column=2, value='=IF(%s="","не выявлены",%s)' % (kvcat, kvcat))
hdr(ws, MXH, 4, '5. Матрица решений 3×3 и предварительная рекомендация')
ws.cell(row=ZA, column=1, value='Зона I_A:')
ws.cell(row=ZA, column=2, value='=IF(NOT(ISNUMBER(B%d)),"н/д",IF(B%d>=4,"В",IF(B%d>=3,"С","Н")))' % (GRP, GRP, GRP))
ws.cell(row=ZC, column=1, value='Зона I_C:')
ws.cell(row=ZC, column=2, value='=IF(NOT(ISNUMBER(C%d)),"н/д",IF(C%d>=4,"В",IF(C%d>=3,"С","Н")))' % (GRP, GRP, GRP))
ws.cell(row=MXLAB, column=1, value='I_A ↓ \\ I_C →')
ws.cell(row=MXLAB, column=2, value='В'); ws.cell(row=MXLAB, column=3, value='С'); ws.cell(row=MXLAB, column=4, value='Н')
for j in range(1, 5): ws.cell(row=MXLAB, column=j).font = F_B; ws.cell(row=MXLAB, column=j).fill = FILL_S
for i, za in enumerate(('В', 'С', 'Н')):
    rr = MX0 + i
    ws.cell(row=rr, column=1, value=za).font = F_B; ws.cell(row=rr, column=1).fill = FILL_S
    for j, zc in enumerate(('В', 'С', 'Н')):
        ws.cell(row=rr, column=2 + j, value=MX[za][zc]).alignment = WRAP
ws.cell(row=REC, column=1, value='Предварительная рекомендация (матрица):'); ws.cell(row=REC, column=1).font = F_B
ws.cell(row=REC, column=2, value=('=IF(OR(B%d="н/д",B%d="н/д"),"РЕКОМЕНДАЦИЯ НЕ ФОРМИРУЕТСЯ (неполная оценка)",'
    'INDEX(B%d:D%d,MATCH(B%d,A%d:A%d,0),MATCH(B%d,B%d:D%d,0)))' % (ZA, ZC, MX0, MX0 + 2, ZA, MX0, MX0 + 2, ZC, MXLAB, MXLAB)))
ws.cell(row=OVR, column=1, value='Рекомендация с оверрайдом стоп-фактора:'); ws.cell(row=OVR, column=1).font = F_B
ws.cell(row=OVR, column=2, value='=IF(B%d="СТОП-ФАКТОР ВЫЯВЛЕН",IF(B%d="Продолжить НИОКР","ПОНИЖЕНИЕ: Пересмотреть параметры НИОКР (стоп-фактор)",B%d),B%d)' % (SFF, REC, REC, REC))
ws.cell(row=OVR, column=2).font = F_B
lowcat = '&'.join(LC(LOWH, Q0 + i) for i in range(22))
disccat = '&'.join(LC(DISH, Q0 + i) for i in range(22))
ws.cell(row=LOWR, column=1, value='Перечень вопросов с низкими оценками (≤2):'); ws.cell(row=LOWR, column=1).font = F_B
ws.cell(row=LOWR, column=2, value='=IF(%s="","не выявлены",%s)' % (lowcat, lowcat))
ws.cell(row=DISR, column=1, value='Перечень вопросов с разногласиями (σ>1,00):'); ws.cell(row=DISR, column=1).font = F_B
ws.cell(row=DISR, column=2, value='=IF(%s="","не выявлены",%s)' % (disccat, disccat))
ws.column_dimensions['A'].width = 26
for e in range(1, NEXP + 1):
    ws.column_dimensions[get_column_letter(1 + e)].width = 7
for j, w in enumerate((10, 7, 7, 9, 8, 14, 12, 14, 16, 26, 40), MEAN):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = 'B%d' % Q0
# условное форматирование свода
ws.conditional_formatting.add('B%d:C%d' % (E0, GRP - 1), CellIsRule(operator='lessThan', formula=['3'], fill=PatternFill('solid', fgColor=REDF)))
ws.conditional_formatting.add('B%d:C%d' % (E0, GRP - 1), CellIsRule(operator='between', formula=['3', '3.99'], fill=PatternFill('solid', fgColor=ORANG)))
ws.conditional_formatting.add('B%d:C%d' % (E0, GRP - 1), CellIsRule(operator='greaterThanOrEqual', formula=['4'], fill=PatternFill('solid', fgColor=GREEN)))
qr = '%s%d:%s%d' % (get_column_letter(MEAN), Q0, get_column_letter(MEAN), Q0 + 21)
ws.conditional_formatting.add(qr, CellIsRule(operator='lessThan', formula=['3'], fill=PatternFill('solid', fgColor=REDF)))
ws.conditional_formatting.add(qr, CellIsRule(operator='greaterThanOrEqual', formula=['4'], fill=PatternFill('solid', fgColor=GREEN)))
ws.conditional_formatting.add('%s%d:%s%d' % (get_column_letter(SD), Q0, get_column_letter(SD), Q0 + 21),
    ColorScaleRule(start_type='num', start_value=0, start_color='63BE7B', mid_type='num', mid_value=0.8, mid_color='FFEB84',
                   end_type='num', end_value=1.5, end_color='F8696B'))
ws.conditional_formatting.add('F%d:F%d' % (SF0, SF0 + 4), CellIsRule(operator='equal', formula=['"ВЫЯВЛЕН"'], fill=PatternFill('solid', fgColor=REDF), font=Font(bold=True, color='9C0006')))
ws.conditional_formatting.add('D%d:D%d' % (KV0, KV0 + 8), CellIsRule(operator='equal', formula=['"ГРУППОВОЙ СИГНАЛ"'], fill=PatternFill('solid', fgColor=REDF), font=Font(bold=True, color='9C0006')))
ws.conditional_formatting.add('%s%d:%s%d' % (get_column_letter(FLG), Q0, get_column_letter(FLG), Q0 + 21),
    CellIsRule(operator='equal', formula=['"РАЗНОГЛАСИЕ"'], fill=PatternFill('solid', fgColor=ORANG), font=Font(bold=True)))
ws.sheet_view.showGridLines = False

# ============================ ЛИСТ «Панель» ============================
ws = wb.create_sheet('Панель')
title(ws, 'АНАЛИТИЧЕСКАЯ ПАНЕЛЬ ЭО-НИОКР (визуализация расчётов листа «Свод»)', 8)
ws['A3'] = 'Индекс актуальности I_A (групп.)'; ws['B3'] = '=IF(ISNUMBER(Свод!B15),Свод!B15,"—")'; ws['B3'].number_format = '0.00'
ws['C3'] = '=Свод!B71'; ws['D3'] = '=Свод!B17'
ws['A4'] = 'Индекс целесообразности I_C (групп.)'; ws['B4'] = '=IF(ISNUMBER(Свод!C15),Свод!C15,"—")'; ws['B4'].number_format = '0.00'
ws['C4'] = '=Свод!B72'; ws['D4'] = '=Свод!C17'
for rr in (3, 4):
    ws.cell(row=rr, column=1).font = F_B
    ws.cell(row=rr, column=2).font = Font(bold=True, size=18, color=NAVY)
ws['A6'] = 'СТОП-ФАКТОРЫ:'; ws['A6'].font = F_B; ws['B6'] = '=Свод!B54'
ws['A7'] = '«Красные» сигналы:'; ws['A7'].font = F_B; ws['B7'] = '=Свод!B68'
ws['A8'] = 'ПРЕДВАРИТЕЛЬНАЯ РЕКОМЕНДАЦИЯ:'; ws['A8'].font = F_B
ws['B8'] = '=Свод!B80'; ws['B8'].font = Font(bold=True, size=13, color=NAVY)
ws.merge_cells('B8:E8')
hdr(ws, 10, 4, 'Профиль 22 критериев (средние баллы группы) и тепловая карта σ')
ws['A11'] = 'Код'; ws['B11'] = 'Средний балл'; ws['C11'] = 'σ'; ws['D11'] = 'Критерий'
for j in range(1, 5): ws.cell(row=11, column=j).font = F_B; ws.cell(row=11, column=j).fill = FILL_S
for i in range(22):
    rr = 12 + i
    ws.cell(row=rr, column=1, value='=Свод!A%d' % (21 + i)).alignment = WRAPC
    ws.cell(row=rr, column=2, value='=IF(ISNUMBER(Свод!L%d),Свод!L%d,"")' % (21 + i, 21 + i)).number_format = '0.00'
    ws.cell(row=rr, column=3, value='=IF(ISNUMBER(Свод!P%d),Свод!P%d,"")' % (21 + i, 21 + i)).number_format = '0.00'
    ws.cell(row=rr, column=4, value='=Свод!U%d' % (21 + i)).alignment = WRAP
ws.conditional_formatting.add('B12:B33', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=5, color=ACC))
ws.conditional_formatting.add('C12:C33', ColorScaleRule(start_type='num', start_value=0, start_color='63BE7B',
                                                        mid_type='num', mid_value=0.8, mid_color='FFEB84',
                                                        end_type='num', end_value=1.5, end_color='F8696B'))
ch = BarChart(); ch.type = 'col'; ch.title = 'Профиль 22 критериев (средние баллы)'
ch.y_axis.title = 'балл (1–5)'; ch.x_axis.title = 'код вопроса'
data = Reference(ws, min_col=2, min_row=11, max_row=33)
cats = Reference(ws, min_col=1, min_row=12, max_row=33)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ch.height = 9; ch.width = 24; ch.legend = None
ws.add_chart(ch, 'F11')
hdr(ws, 36, 4, 'Матрица решений (активная ячейка подсвечена)')
ws['A37'] = 'I_A ↓ \\ I_C →'; ws['B37'] = 'В'; ws['C37'] = 'С'; ws['D37'] = 'Н'
for j in range(1, 5): ws.cell(row=37, column=j).font = F_B; ws.cell(row=37, column=j).fill = FILL_S
for i, za in enumerate(('В', 'С', 'Н')):
    rr = 38 + i
    ws.cell(row=rr, column=1, value=za).font = F_B; ws.cell(row=rr, column=1).fill = FILL_S
    for j, zc in enumerate(('В', 'С', 'Н')):
        c = ws.cell(row=rr, column=2 + j, value='=Свод!%s%d' % (get_column_letter(2 + j), 75 + i))
        c.alignment = WRAP
ws.conditional_formatting.add('B38:D40', FormulaRule(formula=['AND($A38=Свод!$B$71,B$37=Свод!$B$72)'],
                              fill=PatternFill('solid', fgColor='FFE699'), font=Font(bold=True)))
for j, w in enumerate((30, 16, 10, 40, 4, 4, 4, 4), 1):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.conditional_formatting.add('B3', CellIsRule(operator='lessThan', formula=['3'], fill=PatternFill('solid', fgColor=REDF)))
ws.conditional_formatting.add('B3', CellIsRule(operator='greaterThanOrEqual', formula=['4'], fill=PatternFill('solid', fgColor=GREEN)))
ws.conditional_formatting.add('B4', CellIsRule(operator='lessThan', formula=['3'], fill=PatternFill('solid', fgColor=REDF)))
ws.conditional_formatting.add('B4', CellIsRule(operator='greaterThanOrEqual', formula=['4'], fill=PatternFill('solid', fgColor=GREEN)))
ws.conditional_formatting.add('B6', CellIsRule(operator='equal', formula=['"СТОП-ФАКТОР ВЫЯВЛЕН"'], fill=PatternFill('solid', fgColor=REDF), font=Font(bold=True, color='9C0006')))
ws.sheet_view.showGridLines = False

# ============================ ЛИСТ «Заключение» ============================
ws = wb.create_sheet('Заключение')
title(ws, 'ИТОГОВОЕ ЭКСПЕРТНОЕ ЗАКЛЮЧЕНИЕ ПО НИОКР (формируется автоматически из расчётных ячеек)', 4)
LOW = '&'.join('Свод!S%d' % (21 + i) for i in range(22))
DIS = '&'.join('Свод!T%d' % (21 + i) for i in range(22))
SFC = '&'.join('Свод!G%d' % (48 + i) for i in range(5))
KVC = '&'.join('Свод!F%d' % (58 + i) for i in range(9))
rows = [
 (3, '1. Объект и состав комиссии',
  '="НИОКР: "&Э1!B3&" (шифр "&Э1!B4&"). Эксперты: "&' + '&'.join('IF(Э%d!B5="","",Э%d!B5&"; ")' % (e, e) for e in range(1, NEXP)) + '&IF(Э%d!B5="","",Э%d!B5&".")' % (NEXP, NEXP)),
 (5, '2. Природа оценки',
  '="Настоящая оценка является экспертным суждением компетентных специалистов на дату "&IF(Э1!F5="","__",""&Э1!F5)&" и не является верификацией документированных фактов. Документальное подтверждение баллов методикой не требуется."'),
 (7, '3. Оценка актуальности',
  '=IF(Свод!B15="н/д","Индекс актуальности не рассчитан (недостаточное число оценок: Н/О более двух у каждого эксперта).","Индекс актуальности составляет "&ROUND(Свод!B15,2)&" балла, что соответствует "&IF(Свод!B15>=4,"высокому",IF(Свод!B15>=3,"среднему","низкому"))&" уровню (зона "&Свод!B71&"). "&IF(Свод!S21&Свод!S22&Свод!S23&Свод!S24&Свод!S25&Свод!S26&Свод!S27&Свод!S28&Свод!S29&Свод!S30="","Вопросов с оценкой ≤2 не выявлено.","Вопросы с оценкой ≤2: "&Свод!S21&Свод!S22&Свод!S23&Свод!S24&Свод!S25&Свод!S26&Свод!S27&Свод!S28&Свод!S29&Свод!S30&"."))'),
 (9, '4. Оценка целесообразности',
  '=IF(Свод!C15="н/д","Индекс целесообразности не рассчитан (недостаточное число оценок).","Индекс целесообразности составляет "&ROUND(Свод!C15,2)&" балла, что соответствует "&IF(Свод!C15>=4,"высокому",IF(Свод!C15>=3,"среднему","низкому"))&" уровню (зона "&Свод!B72&"). "&IF(Свод!S31&Свод!S32&Свод!S33&Свод!S34&Свод!S35&Свод!S36&Свод!S37&Свод!S38&Свод!S39&Свод!S40&Свод!S41&Свод!S42="","Вопросов с оценкой ≤2 не выявлено.","Вопросы с оценкой ≤2: "&Свод!S31&Свод!S32&Свод!S33&Свод!S34&Свод!S35&Свод!S36&Свод!S37&Свод!S38&Свод!S39&Свод!S40&Свод!S41&Свод!S42&"."))'),
 (11, '5. Согласованность экспертов',
  '="σ(I_A)="&IF(ISNUMBER(Свод!B16),ROUND(Свод!B16,2),"—")&" ("&Свод!B17&"); σ(I_C)="&IF(ISNUMBER(Свод!C16),ROUND(Свод!C16,2),"—")&" ("&Свод!C17&"). Вопросов с существенными разногласиями (σ>1,00): "&IF(' + DIS + '="","не выявлены.",' + DIS + '&" — требуется обсуждение с сохранением исходных оценок.")'),
 (13, '6. Стоп-факторы',
  '="Статус: "&Свод!B54&IF(' + SFC + '="","",". Выявлены: "&' + SFC + '&"— обязательное рассмотрение комиссией; автоматическая рекомендация «Продолжить НИОКР» при выявленном факторе не применяется.")&"."'),
 (15, '7. «Красные» вопросы',
  '="Групповые сигналы (≥2 экспертов): "&IF(' + KVC + '="","не выявлены.",' + KVC + '&"— комиссия обязана учесть при вынесении решения.")'),
 (17, '8. Предварительная рекомендация',
  '="По матрице решений: "&Свод!B79&". С учётом оверрайда стоп-факторов: "&Свод!B80&". Основание: I_A="&IF(ISNUMBER(Свод!B15),ROUND(Свод!B15,2),"н/д")&" ("&Свод!B71&"), I_C="&IF(ISNUMBER(Свод!C15),ROUND(Свод!C15,2),"н/д")&" ("&Свод!B72&")."'),
]
for rr, h, f in rows:
    ws.cell(row=rr, column=1, value=h).font = F_SEC
    c = ws.cell(row=rr, column=2, value=f); c.alignment = WRAP
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
    ws.row_dimensions[rr].height = 62
ws.cell(row=19, column=1, value='9. Решение экспертной комиссии').font = F_SEC
ws.cell(row=20, column=1, value='Рекомендация комиссии (выбор):')
dv = DataValidation(type='list', formula1='"' + ','.join(D['recs5']) + '"', allow_blank=True)
ws.add_data_validation(dv); dv.add('B20')
ws['B20'].fill = FILL_IN
ws.cell(row=21, column=1, value='Обоснование решения:')
ws.merge_cells('B21:D23'); ws['B21'].fill = FILL_IN; ws['B21'].alignment = WRAP
ws.cell(row=24, column=1, value='Особые мнения экспертов:')
ws.merge_cells('B24:D26'); ws['B24'].fill = FILL_IN; ws['B24'].alignment = WRAP
ws.cell(row=28, column=1, value='Председатель комиссии: ____________________')
ws.cell(row=29, column=1, value='Члены комиссии: ____________________')
ws.cell(row=30, column=1, value='Ответственный секретарь: ____________________')
ws.column_dimensions['A'].width = 34
for col in 'BCD': ws.column_dimensions[col].width = 46
ws.sheet_view.showGridLines = False

# защита расчётных листов (ввод — только анкеты и свободные поля заключения)
for name in ('Методика', 'Чек-лист_22', 'Свод', 'Панель'):
    wb[name].protection.sheet = True
    wb[name].protection.password = ''
for cell in ('B20', 'B21', 'B24'):
    wb['Заключение'][cell].protection = openpyxl.styles.Protection(locked=False)
wb['Заключение'].protection.sheet = True

wb.save(OUT)
print('OK ->', os.path.abspath(OUT))
print('листов:', wb.sheetnames)
