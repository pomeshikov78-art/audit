#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Этап 22 (v23.0), файл 10. Excel-калькулятор и подробный чек-лист методики
ЭО-НИОКР v1.0 + расширения v2.0 (22 вопроса, шкала 1–5 + Н/О, 5 стоп-факторов, 9 «красных»,
I_A/I_C, матрица 3×3, σ-согласованность, оверрайд стоп-фактора — базис v1.0 неизменен).
Расширения v2.0 (opt-in, лист «Калибровка»): дифференцированные веса экспертов,
индекс результативности I_R (блок Р1–Р6, минимум 5 из 6), калибровка порогов по итогам
апробации, портфель до 30 НИОКР, печатная анкета (прил. №1).
Группа до 10 экспертов. 19 листов: Методика · Чек-лист_22 · Калибровка · Э1–Э10 ·
Свод · Обоснованность · Панель · Портфель · Заключение · Анкета_бланк.
Вся логика — формулами, без макросов (Excel 365/2021)."""
import json, io, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.chart import BarChart, Reference

HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(io.open(os.path.join(HERE, 'data', 'qeo_v23.json'), encoding='utf-8'))['base']
ORG = json.load(io.open(os.path.join(HERE, 'data', 'org_form_v23.json'), encoding='utf-8'))
D21 = json.load(io.open(os.path.join(HERE, 'data', 'qeo_v23.json'), encoding='utf-8'))
RB = D21['rblock']['q']            # блок Р1–Р6 (расширение v2.0)
W21 = D21['weights']               # веса экспертов: default/min/max/step
CAL0 = D21['calib']                # базовые пороги зон и σ
PMAX = D21['portfolio']['max']     # ёмкость портфеля (30)
OUT = os.path.join(HERE, '..', 'output', '10_Чек-лист_ЭО_НИОКР_22_вопроса_v23.xlsx')
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
title(ws, 'МЕТОДИКА ЭО-НИОКР v2.2 (базис v1.0, расширения v2.0/v2.1) — экспертная оценка актуальности, целесообразности и результативности НИОКР (модель v23.0, группа до 10 экспертов, портфель до 30 НИОКР, 11.2026)')
ws['A3'] = 'Назначение: единый порядок экспертной оценки НИОКР. Оценка — профессиональное суждение; Печатная форма воспроизводит Приложение №2 организации, включая исходные формулировки. Правила интерпретации ответов и границы экспертного суждения — раздел 22 методики ЭО-НИОКР v2.2.'
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
for t in ('Экспертная группа: от 3 до 10 специалистов (рекомендуется 5); веса равные (базовый режим v1.0); дифференцированные веса — расширение v2.0 (лист «Калибровка», B9).',
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
hdr(ws, r, 3, '7. Стоп-факторы (5) — ответ 1–5 (Приложение №2); в индексы не входят'); r += 1
for i, s in enumerate(SF, 1):
    ws.cell(row=r, column=1, value=i).alignment = WRAPC
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2, value=s).alignment = WRAP; r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
ws.cell(row=r, column=1, value='Фактор выявлен при ответе 1–2 хотя бы у одного эксперта; ответ 3 требует уточнения комиссией; 4–5 — не выявлен. Пустой ответ не означает «нет». Выявленный стоп-фактор обязателен к рассмотрению комиссией и блокирует автоматическую рекомендацию «Продолжить НИОКР».').alignment = WRAP
r += 2
hdr(ws, r, 3, '8. «Красные» вопросы (9) — шкала 1=Да … 5=Нет; в индексы не входят'); r += 1
for k in KV:
    ws.cell(row=r, column=1, value=k['n']).alignment = WRAPC
    ws.cell(row=r, column=2, value=k['q']).alignment = WRAP
    ws.cell(row=r, column=3, value='1–2 — сигнал риска; 3 — уточнить; 4–5 — нет сигнала').alignment = WRAP
    r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
ws.cell(row=r, column=1, value='Групповой сигнал фиксируется при ≥ 2 экспертах с сигналом; единичный сигнал — особое мнение, в групповой сигнал не входит.').alignment = WRAP
r += 1
hdr(ws, r, 3, '9. Расширения v2.0 (лист «Калибровка»; по умолчанию ОТКЛЮЧЕНЫ — полная совместимость с v1.0)'); r += 1
for t in (
    'Веса экспертов (Калибровка!B9 = Да): w ∈ [0,25; 3,00], шаг 0,25; ввод на листах Э1–Э10 (ячейка D3). Групповой индекс = Σ(w_k·I_k)/Σw_k по экспертам с действующим индексом направления; σ = √(Σw_k·(I_k−I_ср)²/Σw_k). Повопросная статистика — без весов. Основание весов фиксируется в протоколе комиссии.',
    'Индекс результативности I_R (Калибровка!B10 = Да): блок Р1–Р6 на листах Э (строки 65–70), единая шкала 1–5 + Н/О; минимум 5 оценённых из 6. I_R = ΣР/N_Р; зоны — общие с I_A/I_C. При I_R в зоне Н и рекомендации «Продолжить НИОКР» формируется модификатор «с планом повышения результативности»; оверрайд стоп-фактора имеет приоритет над модификатором.',
    'Калибровка порогов (Калибровка!B4:B8): корректировка зон В/С и границ σ допускается ТОЛЬКО по итогам апробации (пилот ≥ 5 НИОКР) с оформлением протокола (раздел 18 методики ЭО-НИОКР v2.2). Базовые значения: В ≥ 4,00; С ≥ 3,00; σ 0,50 / 0,80 / 1,00.',
    'Портфель НИОКР (лист «Портфель»): реестр до 30 проектов, сравнительная матрица зон I_A/I_C, сводная статистика; групповые значения переносятся из листа «Свод» (строка-помощник 45) после оценки каждой НИОКР. В интерактивной HTML-модели портфель ведётся автоматически.',
    'Бумажная процедура (лист «Анкета_бланк»): единая форма организации — Приложение №2, две страницы А4, исходные тексты скана; заполняется экспертом независимо (раунд 1), затем секретарь переносит баллы в листы Э1–Э10.',
    'Обоснованность решения (v23.0, лист «Обоснованность»): совещательный статистический слой — 95% доверительные интервалы индексов (Стьюдент, df=n−1; в режиме весов — Kish n_eff), устойчивость зон, коэффициент конкордации Кендалла W с поправкой на связки и χ²-значимостью, jackknife-устойчивость рекомендации, повопросные медианы и ДИ; ворота G1–G5 и вердикт. Базовая рекомендация (Свод!B81) НЕ изменяется; решение принимает комиссия. Методика ЭО-НИОКР v2.2, раздел 21, приложение 13.',
    'Состав 22 вопросов, 5 стоп-факторов, 9 «красных» вопросов, шкала и якоря — НЕИЗМЕННЫ (базис ЭО-НИОКР v1.0). Расширения не подменяют решение комиссии.'
):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=1, value='• ' + t).alignment = WRAP; r += 1
r += 1
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
    title(ws, 'АНКЕТА ЭКСПЕРТНОЙ ОЦЕНКИ НИОКР — эксперт %d (ЭО-НИОКР v1.0 + расширения v2.0; 22 вопроса + блок Р)' % e, 6)
    ws['A3'] = 'Наименование НИОКР:'; ws['B3'] = ''
    ws['A4'] = 'Шифр:'; ws['B4'] = ''
    ws['A5'] = 'Эксперт (Ф.И.О.):'; ws['B5'] = ''
    ws['C5'] = 'Компетенция:'; ws['D5'] = ''
    ws['E5'] = 'Дата:'; ws['F5'] = ''
    ws['C3'] = 'Вес эксперта (v2.0):'; ws['D3'] = W21['default']; ws['C3'].font = F_B
    for cc in ('B3', 'B4', 'B5', 'D5', 'F5', 'D3'):
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
    ws.cell(row=43, column=1, value='СТОП-ФАКТОРЫ (оцениваются отдельно от 22 вопросов; ответ 1–5 по форме организации: 1=да, 2=возможно да, 3=неопределённо, 4=возможно нет, 5=нет)').font = F_SEC
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
        ws.cell(row=rr, column=4, value='1–2 = риск; 3 = уточнить').alignment = WRAPC
        ws.cell(row=rr, column=5, value='=IF(C%d="","",IF(C%d<=2,"СИГНАЛ",IF(C%d=3,"УТОЧНИТЬ","")))' % (rr, rr, rr)).alignment = WRAPC
        ws.cell(row=rr, column=6).alignment = WRAP
        ws.row_dimensions[rr].height = 26
    # блок Р — результативность (расширение v2.0)
    ws.cell(row=63, column=1, value='РАЗДЕЛ В. РЕЗУЛЬТАТИВНОСТЬ — 6 вопросов Р1–Р6 (расширение v2.0; учитывается при включении на листе «Калибровка», B10)').font = F_SEC
    for j, t in enumerate(('№', 'Критерий', 'Вопрос эксперту', 'Оценка 1–5 / Н/О', 'Комментарий (необязательно)'), 1):
        c = ws.cell(row=64, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
    for i, q in enumerate(RB):
        rr = 65 + i
        ws.cell(row=rr, column=1, value=q['n']).alignment = WRAPC
        ws.cell(row=rr, column=2, value=q['crit']).alignment = WRAP
        ws.cell(row=rr, column=3, value=q['q']).alignment = WRAP
        ws.cell(row=rr, column=4).fill = FILL_IN; ws.cell(row=rr, column=4).alignment = WRAPC
        ws.cell(row=rr, column=5).alignment = WRAP
        ws.row_dimensions[rr].height = 30
    ws['A72'] = 'I_R — индекс результативности (v2.0)'; ws['A72'].font = F_B
    ws['B72'] = '=COUNT(D65:D70)'; ws['C72'] = '=SUM(D65:D70)'
    ws['D72'] = '=IF(Калибровка!$B$10<>"Да","ОТКЛЮЧЕН",IF(B72<5,"НЕ РАССЧИТАН",C72/B72))'; ws['D72'].number_format = '0.00'
    ws['A73'] = 'Число Н/О по блоку Р (допуск ≤ 1):'; ws['B73'] = '=COUNTIF(D65:D70,"Н/О")'
    # валидации
    dv = DataValidation(type='list', formula1='"1,2,3,4,5,Н/О"', allow_blank=True,
                        showErrorMessage=True, errorTitle='Недопустимое значение',
                        error='Допустимо: 1, 2, 3, 4, 5 или Н/О')
    ws.add_data_validation(dv)
    dv.add('D9:D18'); dv.add('D21:D32')
    dv2 = DataValidation(type='list', formula1='"1,2,3,4,5"', allow_blank=True,
                         showErrorMessage=True, errorTitle='Недопустимое значение',
                         error='Допустимо: 1, 2, 3, 4, 5')
    ws.add_data_validation(dv2); dv2.add('C45:C49')
    dv3 = DataValidation(type='list', formula1='"1,2,3,4,5"', allow_blank=True,
                         showErrorMessage=True, errorTitle='Недопустимое значение',
                         error='Допустимо: 1, 2, 3, 4, 5')
    ws.add_data_validation(dv3); dv3.add('C53:C61')
    dv4 = DataValidation(type='list', formula1='"1,2,3,4,5,Н/О"', allow_blank=True,
                         showErrorMessage=True, errorTitle='Недопустимое значение',
                         error='Допустимо: 1, 2, 3, 4, 5 или Н/О')
    ws.add_data_validation(dv4); dv4.add('D65:D70')
    dv5 = DataValidation(type='decimal', operator='between', formula1=str(W21['min']), formula2=str(W21['max']),
                         allow_blank=True, showErrorMessage=True, errorTitle='Недопустимый вес',
                         error='Вес эксперта: от 0,25 до 3,00 (шаг 0,25)')
    ws.add_data_validation(dv5); dv5.add('D3')
    for j, w in enumerate((6, 30, 62, 16, 30, 24), 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'A9'
    ws.sheet_view.showGridLines = False

# ============================ ЛИСТ «Свод» ============================
ws = wb.create_sheet('Свод')
title(ws, 'СВОДНАЯ ЭКСПЕРТНАЯ ОЦЕНКА — расчётное ядро v23.0 (I_A/I_C/I_R, веса, σ, сигналы, матрица; до 10 экспертов)', 22)
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
hdr(ws, 3, 6, '1. Индивидуальные и групповые индексы (до 10 экспертов; I_R и веса — расширение v2.0)')
for j, t in enumerate(('Эксперт', 'I_A (актуальность)', 'I_C (целесообразность)', 'I_R (v2.0)', 'Вес (v2.0)', 'Статус'), 1):
    c = ws.cell(row=4, column=j, value=t); c.font = F_B; c.fill = FILL_S
for j, t in enumerate(('w·I_A', 'w_A', 'w·I_C', 'w_C', 'w·I_R', 'w_R', 'w·(I_A-ср)²', 'w·(I_C-ср)²', 'w·(I_R-ср)²'), 23):
    ws.cell(row=4, column=j, value=t).font = F_B
for e in range(1, NEXP + 1):
    rr = E0 + e - 1
    ws.cell(row=rr, column=1, value='Эксперт %d (Э%d)' % (e, e))
    ws.cell(row=rr, column=2, value="=Э%d!D34" % e).number_format = '0.00'
    ws.cell(row=rr, column=3, value="=Э%d!D35" % e).number_format = '0.00'
    ws.cell(row=rr, column=4, value='=IF(Калибровка!$B$10="Да",Э%d!D72,"—")' % e).number_format = '0.00'
    ws.cell(row=rr, column=5, value='=IF(ISNUMBER(Э%d!D3),Э%d!D3,1)' % (e, e)).number_format = '0.00'
    ws.cell(row=rr, column=6, value=('=IF(Э%d!B41="НЕ ЗАПОЛНЕНА","НЕ ЗАПОЛНЕНА",'
        'IF(AND(NOT(ISNUMBER(Э%d!D34)),NOT(ISNUMBER(Э%d!D35))),"ИНДЕКСЫ НЕ РАССЧИТАНЫ","ОК"))' % (e, e, e)))
    # скрытые помощники W..AE (23..31) для взвешенных групповых и σ
    ws.cell(row=rr, column=23, value='=IF(ISNUMBER(B%d),B%d*E%d,0)' % (rr, rr, rr))
    ws.cell(row=rr, column=24, value='=IF(ISNUMBER(B%d),E%d,0)' % (rr, rr))
    ws.cell(row=rr, column=25, value='=IF(ISNUMBER(C%d),C%d*E%d,0)' % (rr, rr, rr))
    ws.cell(row=rr, column=26, value='=IF(ISNUMBER(C%d),E%d,0)' % (rr, rr))
    ws.cell(row=rr, column=27, value='=IF(ISNUMBER(D%d),D%d*E%d,0)' % (rr, rr, rr))
    ws.cell(row=rr, column=28, value='=IF(ISNUMBER(D%d),E%d,0)' % (rr, rr))
    ws.cell(row=rr, column=29, value='=IF(ISNUMBER(B%d),E%d*(B%d-$B$%d)^2,0)' % (rr, rr, rr, GRP))
    ws.cell(row=rr, column=30, value='=IF(ISNUMBER(C%d),E%d*(C%d-$C$%d)^2,0)' % (rr, rr, rr, GRP))
    ws.cell(row=rr, column=31, value='=IF(ISNUMBER(D%d),E%d*(D%d-$D$%d)^2,0)' % (rr, rr, rr, GRP))
ws.column_dimensions.group('W', 'AE', hidden=True)
ws.cell(row=GRP, column=1, value='ГРУППОВОЕ ЗНАЧЕНИЕ (среднее / взвешенное)').font = F_B
ws.cell(row=GRP, column=2, value='=IF(Калибровка!$B$9="Да",IF(SUM(X5:X14)=0,"н/д",SUM(W5:W14)/SUM(X5:X14)),IFERROR(AVERAGE(B5:B14),"н/д"))').number_format = '0.00'
ws.cell(row=GRP, column=3, value='=IF(Калибровка!$B$9="Да",IF(SUM(Z5:Z14)=0,"н/д",SUM(Y5:Y14)/SUM(Z5:Z14)),IFERROR(AVERAGE(C5:C14),"н/д"))').number_format = '0.00'
ws.cell(row=GRP, column=4, value='=IF(Калибровка!$B$10<>"Да","—",IF(Калибровка!$B$9="Да",IF(SUM(AB5:AB14)=0,"н/д",SUM(AA5:AA14)/SUM(AB5:AB14)),IFERROR(AVERAGE(D5:D14),"н/д")))').number_format = '0.00'
ws.cell(row=GRP, column=2).font = F_B; ws.cell(row=GRP, column=3).font = F_B; ws.cell(row=GRP, column=4).font = F_B
ws.cell(row=GRP, column=6, value='веса — при включённом режиме v2.0 (Калибровка!B9)')
ws.cell(row=SIG, column=1, value='σ (стандартное отклонение, популяционное)').font = F_B
ws.cell(row=SIG, column=2, value='=IF(COUNT(B5:B14)<2,"н/д",IF(Калибровка!$B$9="Да",SQRT(SUM(AC5:AC14)/SUM(X5:X14)),STDEVP(B5:B14)))').number_format = '0.00'
ws.cell(row=SIG, column=3, value='=IF(COUNT(C5:C14)<2,"н/д",IF(Калибровка!$B$9="Да",SQRT(SUM(AD5:AD14)/SUM(Z5:Z14)),STDEVP(C5:C14)))').number_format = '0.00'
ws.cell(row=SIG, column=4, value='=IF(Калибровка!$B$10<>"Да","—",IF(COUNT(D5:D14)<2,"н/д",IF(Калибровка!$B$9="Да",SQRT(SUM(AE5:AE14)/SUM(AB5:AB14)),STDEVP(D5:D14))))').number_format = '0.00'
ws.cell(row=CON, column=1, value='Согласованность мнений').font = F_B
for col in 'BCD':
    ws[col + str(CON)] = ('=IF(NOT(ISNUMBER(%s%d)),"н/д",IF(%s%d<=Калибровка!$B$6,"Высокая",IF(%s%d<=Калибровка!$B$7,"Приемлемая",'
                          'IF(%s%d<=Калибровка!$B$8,"Умеренные разногласия","Существенные разногласия"))))' % (col, SIG, col, SIG, col, SIG, col, SIG))
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
    ws.cell(row=rr, column=FLG, value='=IF(AND(ISNUMBER(%s),%s>Калибровка!$B$8),"РАЗНОГЛАСИЕ","")' % (LC(SD, rr), LC(SD, rr))).alignment = WRAPC
    ws.cell(row=rr, column=LOW, value='=IF(AND(ISNUMBER(%s),%s<=2),"НИЗКАЯ","")' % (LC(MEAN, rr), LC(MEAN, rr))).alignment = WRAPC
    ws.cell(row=rr, column=LOWH, value='=IF(%s="НИЗКАЯ",A%d&" ","")' % (LC(LOW, rr), rr))
    ws.cell(row=rr, column=DISH, value='=IF(%s="РАЗНОГЛАСИЕ",A%d&" ","")' % (LC(FLG, rr), rr))
    ws.cell(row=rr, column=CRIT, value=q['crit']).alignment = WRAP
    ws.cell(row=rr, column=QUES, value=q['q']).alignment = WRAP
ws.cell(row=QCNT, column=1, value='Вопросов с разногласиями (σ>порога калибровки):').font = F_B
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
    ws.cell(row=rr, column=3, value='=' + '+'.join('COUNTIF(Э%d!C%d,1)+COUNTIF(Э%d!C%d,2)' % (e, er, e, er) for e in range(1, NEXP + 1))).alignment = WRAPC
    ws.cell(row=rr, column=4, value='=' + '+'.join('COUNTIF(Э%d!C%d,3)' % (e, er) for e in range(1, NEXP + 1))).alignment = WRAPC
    ws.cell(row=rr, column=5, value='=' + '+'.join('COUNTIF(Э%d!C%d,4)+COUNTIF(Э%d!C%d,5)' % (e, er, e, er) for e in range(1, NEXP + 1))).alignment = WRAPC
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
    ws.cell(row=rr, column=5, value='1–2 = риск; 3 = уточнить').alignment = WRAPC
    ws.cell(row=rr, column=6, value='=IF(D%d="ГРУППОВОЙ СИГНАЛ","КВ%d; ","")' % (rr, k['n']))
ws.cell(row=KVL, column=1, value='Перечень групповых «красных» сигналов:').font = F_B
kvcat = '&'.join('F%d' % (KV0 + i) for i in range(9))
ws.cell(row=KVL, column=2, value='=IF(%s="","не выявлены",%s)' % (kvcat, kvcat))
hdr(ws, MXH, 4, '5. Матрица решений 3×3 и предварительная рекомендация')
ws.cell(row=ZA, column=1, value='Зона I_A:')
ws.cell(row=ZA, column=2, value='=IF(NOT(ISNUMBER(B%d)),"н/д",IF(B%d>=Калибровка!$B$4,"В",IF(B%d>=Калибровка!$B$5,"С","Н")))' % (GRP, GRP, GRP))
ws.cell(row=ZC, column=1, value='Зона I_C:')
ws.cell(row=ZC, column=2, value='=IF(NOT(ISNUMBER(C%d)),"н/д",IF(C%d>=Калибровка!$B$4,"В",IF(C%d>=Калибровка!$B$5,"С","Н")))' % (GRP, GRP, GRP))
ws.cell(row=73, column=1, value='Зона I_R (v2.0):')
ws.cell(row=73, column=2, value='=IF(NOT(ISNUMBER(D%d)),"н/д",IF(D%d>=Калибровка!$B$4,"В",IF(D%d>=Калибровка!$B$5,"С","Н")))' % (GRP, GRP, GRP))
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
ws.cell(row=81, column=1, value='Рекомендация с учётом результативности (v2.0):'); ws.cell(row=81, column=1).font = F_B
ws.cell(row=81, column=2, value='=IF(AND(Калибровка!$B$10="Да",B73="Н",B80="Продолжить НИОКР"),"Продолжить НИОКР с планом повышения результативности (I_R ниже зоны С)",B80)')
ws.cell(row=81, column=2).font = F_B
lowcat = '&'.join(LC(LOWH, Q0 + i) for i in range(22))
disccat = '&'.join(LC(DISH, Q0 + i) for i in range(22))
ws.cell(row=LOWR, column=1, value='Перечень вопросов с низкими оценками (≤2):'); ws.cell(row=LOWR, column=1).font = F_B
ws.cell(row=LOWR, column=2, value='=IF(%s="","не выявлены",%s)' % (lowcat, lowcat))
ws.cell(row=DISR, column=1, value='Перечень вопросов с разногласиями (σ>порога):'); ws.cell(row=DISR, column=1).font = F_B
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
# ---- 6. Обоснованность решения (v23.0, совещательный слой) ----
hdr(ws, 85, 4, '6. Обоснованность решения (v23.0, совещательный слой — лист «Обоснованность»)')
ws.cell(row=86, column=1, value='W Кендалла (22 вопроса):').font = F_B
ws.cell(row=86, column=2, value='=Обоснованность!D51').number_format = '0.00'
ws.cell(row=87, column=1, value='Значимость W (χ²):')
ws.cell(row=87, column=2, value='=Обоснованность!I51')
ws.cell(row=88, column=1, value='ДИ I_A (95%), нижн./верхн.:').font = F_B
ws.cell(row=88, column=2, value='=Обоснованность!H44').number_format = '0.00'
ws.cell(row=88, column=3, value='=Обоснованность!I44').number_format = '0.00'
ws.cell(row=89, column=1, value='ДИ I_C (95%), нижн./верхн.:').font = F_B
ws.cell(row=89, column=2, value='=Обоснованность!H45').number_format = '0.00'
ws.cell(row=89, column=3, value='=Обоснованность!I45').number_format = '0.00'
ws.cell(row=90, column=1, value='Устойчивость зон (I_A / I_C):')
ws.cell(row=90, column=2, value='=Обоснованность!L44')
ws.cell(row=90, column=3, value='=Обоснованность!L45')
ws.cell(row=91, column=1, value='Изменений рекомендации (jackknife):')
ws.cell(row=91, column=2, value='=Обоснованность!B73')
ws.cell(row=92, column=1, value='ВЕРДИКТ (совещательный):').font = F_B
ws.cell(row=92, column=2, value='=Обоснованность!B84'); ws.cell(row=92, column=2).font = F_B
ws.merge_cells('B92:D92')
ws.cell(row=93, column=1, value='Базовая рекомендация (строка 81) не изменяется; окончательное решение принимает комиссия.').alignment = WRAP
ws.merge_cells('A93:D93')
ws.sheet_view.showGridLines = False

# ============================ ЛИСТ «Обоснованность» (v22) ============================
ws = wb.create_sheet('Обоснованность')
title(ws, 'ОБОСНОВАННОСТЬ РЕШЕНИЯ — математический аппарат v23.0 (ЭО-НИОКР v2.2 §21): ДИ, W Кендалла, χ², jackknife, ворота G1–G5', 12)
ws['A2'] = 'Совещательный статистический слой: базовая рекомендация (Свод!B81) НЕ изменяется. Решение принимает комиссия. Шкала 1–5 порядковая — ДИ по t-критерию является приближением для малых выборок (n≤10).'
ws['A2'].alignment = WRAP; ws.merge_cells('A2:L2'); ws.row_dimensions[2].height = 30
LC = lambda c, r: get_column_letter(c) + str(r)
# --- 1. критические значения ---
hdr(ws, 4, 8, '1. Критические значения и шкала интерпретации')
ws['A5'] = 'df'; ws['B5'] = 't(0,975; df)'
TVALS = [12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262]
for i, tv in enumerate(TVALS):
    ws.cell(row=6 + i, column=1, value=i + 1).alignment = WRAPC
    ws.cell(row=6 + i, column=2, value=tv).number_format = '0.000'
ws['D5'] = 'df'; ws['E5'] = 'χ²(0,95; df)'
for i, (df, cv) in enumerate([(9, 16.919), (11, 19.675), (21, 32.671)]):
    ws.cell(row=6 + i, column=4, value=df).alignment = WRAPC
    ws.cell(row=6 + i, column=5, value=cv).number_format = '0.000'
ws['G5'] = 'W'; ws['H5'] = 'Интерпретация'
for i, (a, b) in enumerate([('<0,3', 'слабая'), ('0,3–0,5', 'умеренная'), ('0,5–0,7', 'хорошая'), ('≥0,7', 'высокая')]):
    ws.cell(row=6 + i, column=7, value=a).alignment = WRAPC
    ws.cell(row=6 + i, column=8, value=b)
for c in ('A5', 'B5', 'D5', 'E5', 'G5', 'H5'): ws[c].font = F_B; ws[c].fill = FILL_S
# --- 2. повопросная статистика ---
hdr(ws, 16, 11, '2. Повопросная статистика: медиана, выборочное s, 95% ДИ среднего (эксперт — единица наблюдения)')
for j, h in enumerate(('Код', 'n', 'Среднее', 'Медиана', 's (выб.)', 'SE', 'df', 't', 'ДИ нижн.', 'ДИ верхн.', 'Ширина ДИ'), 1):
    c = ws.cell(row=17, column=j, value=h); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
for i in range(22):
    r = 18 + i; sr = 21 + i; rng = 'Свод!B%d:K%d' % (sr, sr)
    ws.cell(row=r, column=1, value='=Свод!A%d' % sr).alignment = WRAPC
    ws.cell(row=r, column=2, value='=COUNT(%s)' % rng)
    ws.cell(row=r, column=3, value='=IF(B%d=0,"н/д",Свод!L%d)' % (r, sr)).number_format = '0.00'
    ws.cell(row=r, column=4, value='=IF(B%d=0,"н/д",MEDIAN(%s))' % (r, rng)).number_format = '0.00'
    ws.cell(row=r, column=5, value='=IF(B%d<2,"н/д",STDEV(%s))' % (r, rng)).number_format = '0.000'
    ws.cell(row=r, column=6, value='=IF(B%d<2,"н/д",E%d/SQRT(B%d))' % (r, r, r)).number_format = '0.000'
    ws.cell(row=r, column=7, value='=IF(B%d<2,"н/д",B%d-1)' % (r, r))
    ws.cell(row=r, column=8, value='=IF(B%d<2,"н/д",INDEX($B$6:$B$14,MATCH(MIN(G%d,9),$A$6:$A$14,0)))' % (r, r)).number_format = '0.000'
    ws.cell(row=r, column=9, value='=IF(B%d<2,"н/д",C%d-H%d*F%d)' % (r, r, r, r)).number_format = '0.00'
    ws.cell(row=r, column=10, value='=IF(B%d<2,"н/д",C%d+H%d*F%d)' % (r, r, r, r)).number_format = '0.00'
    ws.cell(row=r, column=11, value='=IF(B%d<2,"н/д",J%d-I%d)' % (r, r, r)).number_format = '0.00'
ws['A40'] = 'ДИ = среднее ± t(0,975; n−1)·s/√n. Медиана и s — устойчивые/выборочные аналоги среднего и популяционной σ (Свод).'
ws['A40'].alignment = WRAP; ws.merge_cells('A40:K40')
# --- 3. ДИ групповых индексов ---
hdr(ws, 42, 12, '3. Доверительные интервалы групповых индексов (95%; в режиме весов — Kish n_eff, при равных весах сводится к s/√n)')
for j, h in enumerate(('Индекс', 'Значение', 'n', 's (выб.)', 'SE', 'df', 't', 'ДИ нижн.', 'ДИ верхн.', 'Зона (нижн.)', 'Зона (верхн.)', 'Устойчивость зоны'), 1):
    c = ws.cell(row=43, column=j, value=h); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
# скрытые помощники весов: строки 170-179 (эксперты), sums 180; B/C/D=IA, E/F/G=IC, H/I/J=IR
for e in range(1, NEXP + 1):
    hr = 169 + e; er = 4 + e
    ws.cell(row=hr, column=2, value='=IF(ISNUMBER(Свод!B%d),Свод!E%d,0)' % (er, er))
    ws.cell(row=hr, column=3, value='=B%d^2' % hr)
    ws.cell(row=hr, column=4, value='=IF(ISNUMBER(Свод!B%d),B%d*(Свод!B%d-Свод!$B$15)^2,0)' % (er, hr, er))
    ws.cell(row=hr, column=5, value='=IF(ISNUMBER(Свод!C%d),Свод!E%d,0)' % (er, er))
    ws.cell(row=hr, column=6, value='=E%d^2' % hr)
    ws.cell(row=hr, column=7, value='=IF(ISNUMBER(Свод!C%d),E%d*(Свод!C%d-Свод!$C$15)^2,0)' % (er, hr, er))
    ws.cell(row=hr, column=8, value='=IF(ISNUMBER(Свод!D%d),Свод!E%d,0)' % (er, er))
    ws.cell(row=hr, column=9, value='=H%d^2' % hr)
    ws.cell(row=hr, column=10, value='=IF(ISNUMBER(Свод!D%d),H%d*(Свод!D%d-Свод!$D$15)^2,0)' % (er, hr, er))
for j in range(2, 11):
    ws.cell(row=180, column=j, value='=SUM(%s170:%s179)' % (get_column_letter(j), get_column_letter(j)))
IDX = [('I_A (актуальность)', 'B', 'B', 'C', 'D', 44), ('I_C (целесообразность)', 'C', 'E', 'F', 'G', 45), ('I_R (результативность)', 'D', 'H', 'I', 'J', 46)]
for name, sc, swc, sw2c, vc, r in IDX:
    sw = '%s180' % swc; sw2 = '%s180' % sw2c; vv = '%s180' % vc
    rng = 'Свод!%s5:%s14' % (sc, sc)
    wguard = 'OR(%s<=0,%s<=0,%s-%s/%s<=0)' % (sw, sw2, sw, sw2, sw)
    ws.cell(row=r, column=1, value=name).font = F_B
    ws.cell(row=r, column=2, value='=IF(ISNUMBER(Свод!%s15),Свод!%s15,"н/д")' % (sc, sc)).number_format = '0.00'
    ws.cell(row=r, column=3, value='=COUNT(%s)' % rng)
    ws.cell(row=r, column=4, value='=IF(C%d<2,"н/д",IF(Калибровка!$B$9="Да",IF(%s,"н/д",SQRT(%s/(%s-%s/%s))),STDEV(%s)))' % (r, wguard, vv, sw, sw2, sw, rng)).number_format = '0.000'
    ws.cell(row=r, column=5, value='=IF(C%d<2,"н/д",IF(Калибровка!$B$9="Да",IF(%s,"н/д",SQRT(%s/(%s-%s/%s)*%s)/%s),D%d/SQRT(C%d)))' % (r, wguard, vv, sw, sw2, sw, sw2, sw, r, r)).number_format = '0.000'
    ws.cell(row=r, column=6, value='=IF(C%d<2,"н/д",IF(Калибровка!$B$9="Да",IF(OR(%s<=0,%s<=0),"н/д",MAX(1,MIN(9,INT(%s*%s/%s)-1))),C%d-1))' % (r, sw, sw2, sw, sw, sw2, r))
    ws.cell(row=r, column=7, value='=IF(OR(C%d<2,NOT(ISNUMBER(F%d))),"н/д",INDEX($B$6:$B$14,MATCH(MIN(F%d,9),$A$6:$A$14,0)))' % (r, r, r)).number_format = '0.000'
    ws.cell(row=r, column=8, value='=IF(OR(C%d<2,NOT(ISNUMBER(E%d))),"н/д",B%d-G%d*E%d)' % (r, r, r, r, r)).number_format = '0.00'
    ws.cell(row=r, column=9, value='=IF(OR(C%d<2,NOT(ISNUMBER(E%d))),"н/д",B%d+G%d*E%d)' % (r, r, r, r, r)).number_format = '0.00'
    ws.cell(row=r, column=10, value='=IF(NOT(ISNUMBER(H%d)),"н/д",IF(H%d>=Калибровка!$B$4,"В",IF(H%d>=Калибровка!$B$5,"С","Н")))' % (r, r, r))
    ws.cell(row=r, column=11, value='=IF(NOT(ISNUMBER(I%d)),"н/д",IF(I%d>=Калибровка!$B$4,"В",IF(I%d>=Калибровка!$B$5,"С","Н")))' % (r, r, r))
    ws.cell(row=r, column=12, value='=IF(OR(J%d="н/д",K%d="н/д"),"н/д",IF(J%d=K%d,"УСТОЙЧИВА","НЕУСТОЙЧИВА"))' % (r, r, r, r))
ws['A47'] = 'Зона устойчива, если ДИ целиком лежит в одной зоне; пересечение границы — повод для обсуждения или привлечения дополнительных экспертов. I_R — при включённом блоке Р (Калибровка!B10).'
ws['A47'].alignment = WRAP; ws.merge_cells('A47:L47')
# --- 4. W Кендалла ---
hdr(ws, 49, 9, '4. Согласованность экспертов — коэффициент конкордации Кендалла W (с поправкой на связки; complete cases)')
for j, h in enumerate(('Блок', 'k', 'Полных наборов (m)', 'W', 'Интерпретация', 'χ²', 'df', 'χ² крит.', 'Значимость'), 1):
    c = ws.cell(row=50, column=j, value=h); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
# скрытые помощники: оценки B..K, ранги M..V
# 22-блок: строки 100-121; count 122; Rj 123 (B..W); summary 124 (B=S,C=ΣT,D=den,E=m); Te 125
# А-блок: 130-139; count 140; Rj 141 (B..K); summary 142; Te 143
# Ц-блок: 150-161; count 162; Rj 163 (B..M); summary 164; Te 165
def tie_term(sc, r1, r2):
    return '+'.join('COUNTIF(%s%d:%s%d,%d)^3-COUNTIF(%s%d:%s%d,%d)' % (sc, r1, sc, r2, v, sc, r1, sc, r2, v) for v in range(1, 6))
for e in range(1, NEXP + 1):
    sc = get_column_letter(1 + e); rc = get_column_letter(12 + e)
    for i in range(22):
        r = 100 + i; erow = (9 + i) if i < 10 else (21 + (i - 10))
        ws.cell(row=r, column=1 + e, value='=IF(ISNUMBER(Э%d!D%d),Э%d!D%d,"")' % (e, erow, e, erow))
        ws.cell(row=r, column=12 + e, value='=IF(ISNUMBER(%s%d),COUNTIF(%s$100:%s$121,"<"&%s%d)+(COUNTIF(%s$100:%s$121,%s%d)+1)/2,"")' % (sc, r, sc, sc, sc, r, sc, sc, sc, r))
    ws.cell(row=122, column=1 + e, value='=COUNT(%s100:%s121)' % (sc, sc))
    ws.cell(row=125, column=1 + e, value='=IF(%s122=22,%s,0)' % (sc, tie_term(sc, 100, 121)))
    for i in range(10):
        r = 130 + i
        ws.cell(row=r, column=1 + e, value='=IF(ISNUMBER(%s%d),%s%d,"")' % (sc, 100 + i, sc, 100 + i))
        ws.cell(row=r, column=12 + e, value='=IF(ISNUMBER(%s%d),COUNTIF(%s$130:%s$139,"<"&%s%d)+(COUNTIF(%s$130:%s$139,%s%d)+1)/2,"")' % (sc, r, sc, sc, sc, r, sc, sc, sc, r))
    ws.cell(row=140, column=1 + e, value='=COUNT(%s130:%s139)' % (sc, sc))
    ws.cell(row=143, column=1 + e, value='=IF(%s140=10,%s,0)' % (sc, tie_term(sc, 130, 139)))
    for i in range(12):
        r = 150 + i
        ws.cell(row=r, column=1 + e, value='=IF(ISNUMBER(%s%d),%s%d,"")' % (sc, 110 + i, sc, 110 + i))
        ws.cell(row=r, column=12 + e, value='=IF(ISNUMBER(%s%d),COUNTIF(%s$150:%s$161,"<"&%s%d)+(COUNTIF(%s$150:%s$161,%s%d)+1)/2,"")' % (sc, r, sc, sc, sc, r, sc, sc, sc, r))
    ws.cell(row=162, column=1 + e, value='=COUNT(%s150:%s161)' % (sc, sc))
    ws.cell(row=165, column=1 + e, value='=IF(%s162=12,%s,0)' % (sc, tie_term(sc, 150, 161)))
for i in range(22):
    ws.cell(row=123, column=2 + i, value='=SUM(M%d:V%d)' % (100 + i, 100 + i))
for i in range(10):
    ws.cell(row=141, column=2 + i, value='=SUM(M%d:V%d)' % (130 + i, 130 + i))
for i in range(12):
    ws.cell(row=163, column=2 + i, value='=SUM(M%d:V%d)' % (150 + i, 150 + i))
# summaries: B=S, C=ΣT, D=den, E=m_used
ws['E124'] = '=COUNTIF(B122:K122,22)'
ws['B124'] = '=SUMPRODUCT((B123:W123-E124*23/2)^2)'
ws['C124'] = '=SUM(B125:K125)'
ws['D124'] = '=E124^2*(22^3-22)-E124*C124'
ws['E142'] = '=COUNTIF(B140:K140,10)'
ws['B142'] = '=SUMPRODUCT((B141:K141-E142*11/2)^2)'
ws['C142'] = '=SUM(B143:K143)'
ws['D142'] = '=E142^2*(10^3-10)-E142*C142'
ws['E164'] = '=COUNTIF(B162:K162,12)'
ws['B164'] = '=SUMPRODUCT((B163:M163-E164*13/2)^2)'
ws['C164'] = '=SUM(B165:K165)'
ws['D164'] = '=E164^2*(12^3-12)-E164*C164'
WB = [('Все 22 вопроса', 22, 124, 51), ('Актуальность (А1–А10)', 10, 142, 52), ('Целесообразность (Ц1–Ц12)', 12, 164, 53)]
for name, k, sr, r in WB:
    ws.cell(row=r, column=1, value=name).font = F_B
    ws.cell(row=r, column=2, value=k)
    ws.cell(row=r, column=3, value='=E%d' % sr)
    ws.cell(row=r, column=4, value='=IF(OR(C%d<2,D%d<=0),"н/д",12*B%d/D%d)' % (r, sr, sr, sr)).number_format = '0.000'
    ws.cell(row=r, column=5, value='=IF(D%d="н/д","—",IF(D%d<0.3,"слабая",IF(D%d<0.5,"умеренная",IF(D%d<0.7,"хорошая","высокая"))))' % (r, r, r, r))
    ws.cell(row=r, column=6, value='=IF(D%d="н/д","н/д",C%d*(B%d-1)*D%d)' % (r, r, r, r)).number_format = '0.00'
    ws.cell(row=r, column=7, value='=B%d-1' % r)
    ws.cell(row=r, column=8, value='=INDEX($E$6:$E$8,MATCH(G%d,$D$6:$D$8,0))' % r).number_format = '0.000'
    ws.cell(row=r, column=9, value='=IF(F%d="н/д","н/д",IF(F%d>=H%d,"значима","не значима"))' % (r, r, r))
ws['A54'] = 'W=12S/(m²(k³−k)−m·ΣT); S=Σ(R_j−m(k+1)/2)²; T=Σ(t³−t) по связкам эксперта; ранги средние при связках. χ²=m(k−1)W, df=k−1. Complete cases: эксперты с полным числовым набором блока. Шкала W — Schmidt (1997).'
ws['A54'].alignment = WRAP; ws.merge_cells('A54:I54'); ws.row_dimensions[54].height = 30
# --- 5. jackknife ---
hdr(ws, 56, 7, '5. Устойчивость рекомендации — jackknife (поочерёдное исключение одного эксперта; полный пересчёт ядра)')
ws['A57'] = 'Н/О всего:'; ws['B57'] = '=' + '+'.join('COUNTIF(Э%d!D9:D32,"Н/О")' % e for e in range(1, NEXP + 1))
ws['A58'] = 'Действующих экспертов (n):'; ws['B58'] = '=COUNTIF(Свод!F5:F14,"ОК")'
ws['A59'] = 'Полных наборов (22):'; ws['B59'] = '=COUNTIF(B122:K122,22)'
ws['A60'] = 'Полнота данных:'; ws['B60'] = '=IF(AND(B57=0,B59=B58,B58>=2),"ДА","НЕТ")'
ws['A61'] = 'СФ «Да» всего:'; ws['B61'] = '=Свод!C48+Свод!C49+Свод!C50+Свод!C51+Свод!C52'
for c in ('A57', 'A58', 'A59', 'A60', 'A61'): ws[c].font = F_B
for j, h in enumerate(('Эксперт', 'I_A (−e)', 'I_C (−e)', 'Зона A (−e)', 'Зона C (−e)', 'Рекомендация (−e)', 'Изменение'), 1):
    c = ws.cell(row=62, column=j, value=h); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
for e in range(1, NEXP + 1):
    r = 62 + e; er = 4 + e
    ws.cell(row=r, column=1, value='=Свод!A%d' % er).alignment = WRAPC
    ws.cell(row=r, column=14, value='=' + '+'.join('COUNTIF(Э%d!C%d,"Да")' % (e, 45 + i) for i in range(5)))  # N: yes_e
    ws.cell(row=r, column=2, value='=IF(OR($B$60<>"ДА",Свод!F%d<>"ОК"),"н/д",IF(Калибровка!$B$9="Да",IF(SUM(Свод!X5:X14)-Свод!X%d<=0,"н/д",(SUM(Свод!W5:W14)-Свод!W%d)/(SUM(Свод!X5:X14)-Свод!X%d)),($B$58*Свод!B15-Свод!B%d)/($B$58-1)))' % (er, er, er, er, er)).number_format = '0.00'
    ws.cell(row=r, column=3, value='=IF(OR($B$60<>"ДА",Свод!F%d<>"ОК"),"н/д",IF(Калибровка!$B$9="Да",IF(SUM(Свод!Z5:Z14)-Свод!Z%d<=0,"н/д",(SUM(Свод!Y5:Y14)-Свод!Y%d)/(SUM(Свод!Z5:Z14)-Свод!Z%d)),($B$58*Свод!C15-Свод!C%d)/($B$58-1)))' % (er, er, er, er, er)).number_format = '0.00'
    ws.cell(row=r, column=15, value='=IF(OR($B$60<>"ДА",Калибровка!$B$10<>"Да",Свод!F%d<>"ОК"),"н/д",IF(Калибровка!$B$9="Да",IF(SUM(Свод!AB5:AB14)-Свод!AB%d<=0,"н/д",(SUM(Свод!AA5:AA14)-Свод!AA%d)/(SUM(Свод!AB5:AB14)-Свод!AB%d)),IF(COUNT(Свод!D5:D14)<2,"н/д",(COUNT(Свод!D5:D14)*Свод!D15-Свод!D%d)/(COUNT(Свод!D5:D14)-1))))' % (er, er, er, er, er))  # O: IR-e
    ws.cell(row=r, column=4, value='=IF(NOT(ISNUMBER(B%d)),"н/д",IF(B%d>=Калибровка!$B$4,"В",IF(B%d>=Калибровка!$B$5,"С","Н")))' % (r, r, r))
    ws.cell(row=r, column=5, value='=IF(NOT(ISNUMBER(C%d)),"н/д",IF(C%d>=Калибровка!$B$4,"В",IF(C%d>=Калибровка!$B$5,"С","Н")))' % (r, r, r))
    ws.cell(row=r, column=16, value='=IF(NOT(ISNUMBER(O%d)),"н/д",IF(O%d>=Калибровка!$B$4,"В",IF(O%d>=Калибровка!$B$5,"С","Н")))' % (r, r, r))  # P: zoneR-e
    ws.cell(row=r, column=17, value='=IF(OR(D%d="н/д",E%d="н/д"),"н/д",INDEX(Свод!$B$75:$D$77,MATCH(D%d,Свод!$A$75:$A$77,0),MATCH(E%d,Свод!$B$74:$D$74,0)))' % (r, r, r, r))  # Q: matrix rec-e
    ws.cell(row=r, column=6, value='=IF(Q%d="н/д","н/д",IF(AND($B$61-N%d>0,Q%d="Продолжить НИОКР"),"ПОНИЖЕНИЕ: Пересмотреть параметры НИОКР (стоп-фактор)",IF(AND(Калибровка!$B$10="Да",P%d="Н",Q%d="Продолжить НИОКР"),"Продолжить НИОКР с планом повышения результативности (I_R ниже зоны С)",Q%d)))' % (r, r, r, r, r, r)).alignment = WRAP
    ws.cell(row=r, column=7, value='=IF(F%d="н/д","н/д",IF(F%d<>Свод!$B$81,1,0))' % (r, r))
ws['A73'] = 'Изменений рекомендации:'; ws['A73'].font = F_B
ws['B73'] = '=IF($B$60<>"ДА","н/д",SUM(G63:G72))'
ws['C73'] = '=IF(B73="н/д","н/д",IF(B73=0,"РЕКОМЕНДАЦИЯ УСТОЙЧИВА","РЕКОМЕНДАЦИЯ НЕУСТОЙЧИВА"))'; ws['C73'].font = F_B
ws.merge_cells('C73:G73')
ws['A74'] = 'Jackknife выполняется при полных данных (нет Н/О, все действующие эксперты оценили 22 вопроса). В режиме весов исключение пересчитывает взвешенное среднее. Оверрайд СФ — по оставшимся экспертам.'
ws['A74'].alignment = WRAP; ws.merge_cells('A74:G74'); ws.row_dimensions[74].height = 28
# --- 6. ворота и вердикт ---
hdr(ws, 76, 3, '6. Ворота обоснованности G1–G5 и совещательный вердикт')
for j, h in enumerate(('Ворота', 'Статус', 'Пояснение'), 1):
    c = ws.cell(row=77, column=j, value=h); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
GATES = [
    ('G1 Полнота оценок', '=IF(AND(B57=0,B59=B58,B58>=1),"ВЫПОЛНЕНО","НЕ ВЫПОЛНЕНО")', '=IF(B78="ВЫПОЛНЕНО","все 22 вопроса оценены, Н/О отсутствуют","Н/О: "&B57&"; полных наборов: "&B59&" из "&B58)'),
    ('G2 Согласованность (W≥0,5; χ² значим)', '=IF(OR(D51="н/д",D51<0.5,I51<>"значима"),"НЕ ВЫПОЛНЕНО","ВЫПОЛНЕНО")', '="W="&IF(D51="н/д","н/д",ROUND(D51,2))&"; "&I51'),
    ('G3 Устойчивость зон I_A и I_C (95% ДИ)', '=IF(AND(L44="УСТОЙЧИВА",L45="УСТОЙЧИВА"),"ВЫПОЛНЕНО","НЕ ВЫПОЛНЕНО")', '="I_A: "&L44&"; I_C: "&L45'),
    ('G4 Устойчивость рекомендации (jackknife)', '=IF(OR(B60<>"ДА",B73="н/д",B58<2),"НЕ ВЫПОЛНЕНО",IF(B73=0,"ВЫПОЛНЕНО","НЕ ВЫПОЛНЕНО"))', '="изменений: "&IF(B73="н/д","н/д",B73)&" из "&B58'),
    ('G5 Отсутствие неразрешённых разногласий', '=IF(AND(Свод!B44=0,Свод!B68="не выявлены"),"ВЫПОЛНЕНО","НЕ ВЫПОЛНЕНО")', '="разногласий (σ): "&Свод!B44&"; КВ: "&Свод!B68'),
]
for i, (name, f, note) in enumerate(GATES):
    r = 78 + i
    ws.cell(row=r, column=1, value=name).alignment = WRAP
    ws.cell(row=r, column=2, value=f).alignment = WRAPC
    ws.cell(row=r, column=3, value=note).alignment = WRAP
ws['A83'] = 'Ворот выполнено:'; ws['A83'].font = F_B
ws['B83'] = '=COUNTIF(B78:B82,"ВЫПОЛНЕНО")'
ws['A84'] = 'ВЕРДИКТ (совещательный):'; ws['A84'].font = F_B
ws['B84'] = '=IF(B58<3,"Недостаточно экспертов (менее 3) — статистическая проверка ограничена",IF(B83=5,"Решение статистически обосновано",IF(B83>=3,"Обосновано с оговорками","Требуется доработка экспертной оценки")))'
ws['B84'].font = Font(bold=True, size=12, color=NAVY); ws.merge_cells('B84:L84')
ws['A85'] = 'Вердикт: 5/5 — «Решение статистически обосновано»; 3–4/5 — «Обосновано с оговорками»; ≤2/5 — «Требуется доработка экспертной оценки»; n<3 — проверка ограничена. Статистический слой не заменяет решение комиссии.'
ws['A85'].alignment = WRAP; ws.merge_cells('A85:L85'); ws.row_dimensions[85].height = 28
# оформление
for j, w in enumerate((30, 12, 12, 12, 12, 10, 8, 12, 12, 12, 12, 16), 1):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = 'B18'
ws.row_dimensions.group(100, 194, hidden=True)
ws.column_dimensions.group('M', 'W', hidden=True)
ws.conditional_formatting.add('L44:L46', CellIsRule(operator='equal', formula=['"НЕУСТОЙЧИВА"'], fill=PatternFill('solid', fgColor=ORANG), font=Font(bold=True)))
ws.conditional_formatting.add('B78:B82', CellIsRule(operator='equal', formula=['"НЕ ВЫПОЛНЕНО"'], fill=PatternFill('solid', fgColor=REDF), font=Font(bold=True, color='9C0006')))
ws.conditional_formatting.add('G63:G72', CellIsRule(operator='equal', formula=['1'], fill=PatternFill('solid', fgColor=REDF), font=Font(bold=True)))
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
ws['A5'] = 'Индекс результативности I_R (v2.0)'; ws['A5'].font = F_B
ws['B5'] = '=IF(ISNUMBER(Свод!D15),Свод!D15,"—")'; ws['B5'].number_format = '0.00'
ws['C5'] = '=Свод!B73'; ws['D5'] = '=Свод!D17'
ws['A6'] = 'СТОП-ФАКТОРЫ:'; ws['A6'].font = F_B; ws['B6'] = '=Свод!B54'
ws['A7'] = '«Красные» сигналы:'; ws['A7'].font = F_B; ws['B7'] = '=Свод!B68'
ws['A8'] = 'ПРЕДВАРИТЕЛЬНАЯ РЕКОМЕНДАЦИЯ:'; ws['A8'].font = F_B
ws['B8'] = '=Свод!B81'; ws['B8'].font = Font(bold=True, size=13, color=NAVY)
ws.merge_cells('B8:E8')
ws['A9'] = 'ОБОСНОВАННОСТЬ (v23.0):'; ws['A9'].font = F_B
ws['B9'] = '=Обоснованность!B84'; ws['B9'].font = Font(bold=True, size=11, color=NAVY); ws['B9'].alignment = WRAP
ws.merge_cells('B9:E9'); ws.row_dimensions[9].height = 28
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
  '="По матрице решений: "&Свод!B79&". С учётом оверрайда стоп-факторов: "&Свод!B80&". Основание: I_A="&IF(ISNUMBER(Свод!B15),ROUND(Свод!B15,2),"н/д")&" ("&Свод!B71&"), I_C="&IF(ISNUMBER(Свод!C15),ROUND(Свод!C15,2),"н/д")&" ("&Свод!B72&")."&IF(Калибровка!B9="Да"," Применены дифференцированные веса экспертов (расширение v2.0).","")&IF(Калибровка!B10="Да"," Итоговая рекомендация с учётом результативности: "&Свод!B81&".","")'),
]
rows21 = [
 (10, '4а. Оценка результативности (v2.0)',
  '=IF(Калибровка!B10<>"Да","Режим оценки результативности отключён (базовая процедура ЭО-НИОКР v1.0).",IF(NOT(ISNUMBER(Свод!D15)),"Индекс результативности I_R не рассчитан (менее 5 из 6 вопросов Р у экспертов с действующей оценкой).","I_R="&ROUND(Свод!D15,2)&" балла (зона "&Свод!B73&"). "&IF(Свод!B73="Н","Результативность ниже зоны С: рекомендуется план повышения результативности (см. итоговую рекомендацию).","Результативность соответствует требованиям.")))'),
 (18, '8а. Режим расчёта и калибровка',
  '=IF(AND(Калибровка!B4=4,Калибровка!B5=3,Калибровка!B6=0.5,Калибровка!B7=0.8,Калибровка!B8=1),"Пороги зон и согласованности — базовые значения ЭО-НИОКР v1.0 (калибровка не применялась).","Пороги калиброваны по итогам апробации (лист «Калибровка»): В≥"&Калибровка!B4&"; С≥"&Калибровка!B5&"; σ: "&Калибровка!B6&" / "&Калибровка!B7&" / "&Калибровка!B8&".")&IF(Калибровка!B9="Да"," Веса экспертов: дифференцированные."," Веса экспертов: равные.")&IF(Калибровка!B10="Да"," Блок Р: включён."," Блок Р: отключён.")'),
]
for rr, h, f in rows + rows21:
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

# ============================ ЛИСТ «Калибровка» (v22) ============================
ws = wb.create_sheet('Калибровка')
title(ws, 'КАЛИБРОВКА ПОРОГОВ И РЕЖИМЫ РАСШИРЕНИЯ v2.0 (изменение порогов — только по итогам апробации, с протоколом)', 4)
ws['A3'] = 'Базовые значения ЭО-НИОКР v1.0: зона В ≥ 4,00; зона С ≥ 3,00; σ: 0,50 / 0,80 / 1,00. Расширения v2.0 по умолчанию отключены (совместимость с v1.0).'
ws['A3'].alignment = WRAP; ws.merge_cells('A3:D3'); ws['A3'].font = F_SEC
cal_rows = [
 (4, 'Порог зоны В (индекс ≥)', CAL0['zoneV'], 'Применяется к I_A, I_C и I_R'),
 (5, 'Порог зоны С (индекс ≥)', CAL0['zoneS'], 'Ниже порога С — зона Н'),
 (6, 'σ: высокая согласованность (≤)', CAL0['s1'], 'Граница уровня «Высокая»'),
 (7, 'σ: приемлемая согласованность (≤)', CAL0['s2'], 'Граница уровня «Приемлемая»'),
 (8, 'σ: флаг разногласия по вопросу (>)', CAL0['s3'], 'Выше — «РАЗНОГЛАСИЕ» по вопросу; до порога — «Умеренные разногласия»'),
 (9, 'Дифференцированные веса экспертов (v2.0)', 'Нет', 'Да — групповые индексы и σ взвешенные (веса на листах Э, ячейка D3)'),
 (10, 'Индекс результативности I_R, блок Р (v2.0)', 'Нет', 'Да — учитывается блок Р1–Р6 и индекс I_R (минимум 5 из 6)'),
]
for rr, lab, val, note in cal_rows:
    ws.cell(row=rr, column=1, value=lab).font = F_B
    c = ws.cell(row=rr, column=2, value=val); c.fill = FILL_IN; c.alignment = WRAPC
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    ws.cell(row=rr, column=3, value=note).alignment = WRAP
hdr(ws, 13, 4, 'Статистика распределения баллов по всем анкетам Э1–Э10 (материал для калибровки)')
R22 = ','.join('Э%d!%s' % (e, rg) for e in range(1, NEXP + 1) for rg in ('D9:D18', 'D21:D32'))
def cntif(v):
    return '+'.join('COUNTIF(Э%d!%s,%s)' % (e, rg, v) for e in range(1, NEXP + 1) for rg in ('D9:D18', 'D21:D32'))
stat = [
 (14, 'Оценённых вопросов (баллы 1–5)', '=COUNT(%s)' % R22, '0'),
 (15, 'Ответов Н/О', '=' + cntif('"Н/О"'), '0'),
 (16, 'Средний балл', '=IF(B14=0,"н/д",SUM(%s)/B14)' % R22, '0.00'),
 (17, 'σ баллов (популяционное)', '=IF(B14<2,"н/д",STDEVP(%s))' % R22, '0.00'),
 (18, 'Доля баллов 1–2', '=IF(B14=0,"н/д",(%s+%s)/B14)' % (cntif('1'), cntif('2')), '0.0%'),
 (19, 'Доля баллов 3', '=IF(B14=0,"н/д",%s/B14)' % cntif('3'), '0.0%'),
 (20, 'Доля баллов 4–5', '=IF(B14=0,"н/д",(%s+%s)/B14)' % (cntif('4'), cntif('5')), '0.0%'),
 (21, 'Завершённых анкет (из 10)', '=' + '+'.join('COUNTIF(Э%d!B41,"ЗАВЕРШЕНА")' % e for e in range(1, NEXP + 1)), '0'),
 (22, 'Заполнение группы, %', '=B21/10', '0%'),
]
for rr, lab, f, nf in stat:
    ws.cell(row=rr, column=1, value=lab)
    c = ws.cell(row=rr, column=2, value=f); c.number_format = nf
ws['A23'] = 'Корректировка порогов допускается только по итогам апробации (пилот ≥ 5 НИОКР) с оформлением протокола — раздел 18 методики ЭО-НИОКР v2.2.'
ws['A23'].alignment = WRAP; ws.merge_cells('A23:D23')
hdr(ws, 25, 4, 'Протокол калибровки (заполняется по итогам апробации)')
for rr, lab in ((26, '№ и дата протокола:'), (27, 'Основание (выборка апробации: НИОКР / анкет):'), (28, 'Утверждено (должность, Ф.И.О., дата):')):
    ws.cell(row=rr, column=1, value=lab).font = F_B
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
    ws.cell(row=rr, column=2).fill = FILL_IN
ws['A30'] = 'Порядок: 1) пилотная апробация; 2) анализ статистики выше; 3) предложение новых порогов; 4) протокол; 5) ввод значений B4:B8. Без протокола пороги не изменяются.'
ws['A30'].alignment = WRAP; ws.merge_cells('A30:D30')
dvz = DataValidation(type='decimal', operator='between', formula1='1', formula2='5', allow_blank=False,
                     showErrorMessage=True, errorTitle='Недопустимый порог', error='Порог зоны: число от 1 до 5')
ws.add_data_validation(dvz); dvz.add('B4'); dvz.add('B5')
dvs = DataValidation(type='decimal', operator='between', formula1='0', formula2='3', allow_blank=False,
                     showErrorMessage=True, errorTitle='Недопустимый порог σ', error='Порог σ: число от 0 до 3')
ws.add_data_validation(dvs); dvs.add('B6:B8')
dvd = DataValidation(type='list', formula1='"Да,Нет"', allow_blank=False,
                     showErrorMessage=True, errorTitle='Недопустимое значение', error='Допустимо: Да или Нет')
ws.add_data_validation(dvd); dvd.add('B9'); dvd.add('B10')
ws.column_dimensions['A'].width = 44; ws.column_dimensions['B'].width = 14
ws.column_dimensions['C'].width = 40; ws.column_dimensions['D'].width = 30
ws.sheet_view.showGridLines = False

# ============================ ЛИСТ «Портфель» (v22) ============================
ws = wb.create_sheet('Портфель')
title(ws, 'ПОРТФЕЛЬ НИОКР — реестр до 30 проектов, сравнительная матрица зон I_A/I_C (v23.0)', 13)
ws['A3'] = 'Значения I_A/I_C переносятся из листа «Свод» (строка-помощник 45) после завершения оценки каждой НИОКР. В интерактивной HTML-модели портфель ведётся автоматически.'
ws['A3'].alignment = WRAP; ws.merge_cells('A3:M3')
PH = ['№', 'Наименование НИОКР', 'Шифр', 'Исполнитель', 'Заказчик', 'I_A', 'I_C', 'Зона I_A', 'Зона I_C',
      'Стоп-фактор (Да/Нет)', 'Рекомендация (матрица + оверрайд)', 'I_R (v2.0)', 'Примечание']
for j, t in enumerate(PH, 1):
    c = ws.cell(row=4, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
ws.row_dimensions[4].height = 30
for i in range(PMAX):
    rr = 5 + i
    ws.cell(row=rr, column=1, value=i + 1).alignment = WRAPC
    for j in (2, 3, 4, 5, 6, 7, 10, 12, 13):
        ws.cell(row=rr, column=j).fill = FILL_IN
    ws.cell(row=rr, column=6).number_format = '0.00'
    ws.cell(row=rr, column=7).number_format = '0.00'
    ws.cell(row=rr, column=12).number_format = '0.00'
    ws.cell(row=rr, column=8, value='=IF(ISNUMBER(F%d),IF(F%d>=Калибровка!$B$4,"В",IF(F%d>=Калибровка!$B$5,"С","Н")),"")' % (rr, rr, rr)).alignment = WRAPC
    ws.cell(row=rr, column=9, value='=IF(ISNUMBER(G%d),IF(G%d>=Калибровка!$B$4,"В",IF(G%d>=Калибровка!$B$5,"С","Н")),"")' % (rr, rr, rr)).alignment = WRAPC
    mx = 'INDEX(Свод!$B$75:$D$77,MATCH(H%d,Свод!$A$75:$A$77,0),MATCH(I%d,Свод!$B$74:$D$74,0))' % (rr, rr)
    ws.cell(row=rr, column=11, value='=IF(OR(H%d="",I%d=""),"",IF(J%d="Да",IF(%s="Продолжить НИОКР","Пересмотреть параметры НИОКР (стоп-фактор)",%s),%s))' % (rr, rr, rr, mx, mx, mx)).alignment = WRAP
    ws.row_dimensions[rr].height = 26
hdr(ws, 36, 6, 'Сводная статистика портфеля')
pstat = [
 (37, 'НИОКР в реестре', '=COUNTA(B5:B34)', '0'),
 (38, 'Средний I_A по реестру', '=IF(COUNT(F5:F34)=0,"н/д",AVERAGE(F5:F34))', '0.00'),
 (39, 'Средний I_C по реестру', '=IF(COUNT(G5:G34)=0,"н/д",AVERAGE(G5:G34))', '0.00'),
 (40, 'НИОКР в квадранте В/В', '=COUNTIFS(H5:H34,"В",I5:I34,"В")', '0'),
 (41, 'НИОКР со стоп-фактором', '=COUNTIF(J5:J34,"Да")', '0'),
 (42, 'Рекомендовано «Продолжить НИОКР»', '=COUNTIF(K5:K34,"Продолжить НИОКР")', '0'),
]
for rr, lab, f, nf in pstat:
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=5)
    ws.cell(row=rr, column=1, value=lab).font = F_B
    c = ws.cell(row=rr, column=6, value=f); c.number_format = nf
hdr(ws, 44, 13, 'Перенос текущей НИОКР (из листов Э1/Свод) — скопируйте ЗНАЧЕНИЯ строки 45 в свободную строку реестра')
rr = 45
ws.cell(row=rr, column=1, value='→').alignment = WRAPC
ws.cell(row=rr, column=2, value='=Э1!B3')
ws.cell(row=rr, column=3, value='=Э1!B4')
ws.cell(row=rr, column=6, value='=IF(ISNUMBER(Свод!B15),Свод!B15,"")').number_format = '0.00'
ws.cell(row=rr, column=7, value='=IF(ISNUMBER(Свод!C15),Свод!C15,"")').number_format = '0.00'
ws.cell(row=rr, column=10, value='=IF(Свод!B54="СТОП-ФАКТОР ВЫЯВЛЕН","Да","Нет")')
ws.cell(row=rr, column=11, value='=Свод!B81').alignment = WRAP
ws.cell(row=rr, column=12, value='=IF(ISNUMBER(Свод!D15),Свод!D15,"")').number_format = '0.00'
dvp = DataValidation(type='decimal', operator='between', formula1='1', formula2='5', allow_blank=True,
                     showErrorMessage=True, errorTitle='Недопустимый индекс', error='Индекс: число от 1 до 5')
ws.add_data_validation(dvp); dvp.add('F5:G34'); dvp.add('L5:L34')
dvj = DataValidation(type='list', formula1='"Да,Нет"', allow_blank=True,
                     showErrorMessage=True, errorTitle='Недопустимое значение', error='Допустимо: Да или Нет')
ws.add_data_validation(dvj); dvj.add('J5:J34')
for col, vals in (('H', ('В', 'С', 'Н')), ('I', ('В', 'С', 'Н'))):
    for zv, fc in zip(vals, (GREEN, ORANG, REDF)):
        ws.conditional_formatting.add('%s5:%s34' % (col, col), CellIsRule(operator='equal', formula=['"%s"' % zv], fill=PatternFill('solid', fgColor=fc)))
ws.conditional_formatting.add('J5:J34', CellIsRule(operator='equal', formula=['"Да"'], fill=PatternFill('solid', fgColor=REDF), font=Font(bold=True, color='9C0006')))
ws.conditional_formatting.add('K5:K34', CellIsRule(operator='equal', formula=['"Продолжить НИОКР"'], fill=PatternFill('solid', fgColor=GREEN)))
for j, w in enumerate((5, 40, 14, 22, 22, 8, 8, 9, 9, 12, 40, 9, 24), 1):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = 'C5'
ws.sheet_view.showGridLines = False

# ============================ ЛИСТ «Анкета_бланк» (v23) ============================
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import Font as XFont, Border as XBorder, Side as XSide, Alignment as XAlign
ws = wb.create_sheet('Анкета_бланк')
# Воспроизведение скана: 8 колонок, две страницы А4; без дополнительных полей.
widths=[4.5,25,51,10,10,10,10,10]
for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
OB=XBorder(*(XSide(style='thin',color='000000'),)*4)
OF=XFont(name='Times New Roman',size=9)
OFC=XAlign(horizontal='center',vertical='center',wrap_text=True)
OFL=XAlign(horizontal='left',vertical='center',wrap_text=True)
def oc(r,c,v=None,bold=False,align='left',size=9):
    x=ws.cell(r,c,v);x.font=XFont(name='Times New Roman',size=size,bold=bold);x.alignment=OFC if align=='center' else OFL;x.border=OB;return x
def om(r,c1,c2,v,bold=False,align='center',size=9):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    for c in range(c1,c2+1):oc(r,c,bold=bold,align=align,size=size)
    oc(r,c1,v,bold,align,size)
def oheader(r,sec):
    om(r,1,8,sec['title'],True,'center',9.5);ws.row_dimensions[r].height=27;r+=1
    oc(r,1,'№',True,'center');oc(r,2,'Критерий',True,'center');oc(r,3,'Вопрос для экспертизы',True,'center');om(r,4,8,'Ответ',True);ws.row_dimensions[r].height=17;r+=1
    om(r,1,3,'',False);oc(r,4,sec['low'],True,'center',7);oc(r,8,sec['high'],True,'center',7);ws.row_dimensions[r].height=25;r+=1
    om(r,1,3,'',False)
    for c in range(4,9):oc(r,c,c-3,True,'center')
    ws.row_dimensions[r].height=14;return r+1
def omain(r,items,sec,height=32):
    r=oheader(r,sec)
    for q in items:
        oc(r,1,q['n'],False,'center');oc(r,2,q['criterion'],True);oc(r,3,q['question'])
        for c in range(4,9):oc(r,c)
        ws.row_dimensions[r].height=height;r+=1
    return r
ws['G1']='Приложение №2';ws['G1'].font=XFont(name='Times New Roman',size=11,bold=True);ws.merge_cells('G1:H1');ws['G1'].alignment=XAlign(horizontal='right')
ws['A2']='ОКР шифр "________________________________"';ws.merge_cells('A2:H2');ws['A2'].font=XFont(name='Times New Roman',size=11,bold=True);ws['A2'].alignment=OFC
ws.row_dimensions[1].height=18;ws.row_dimensions[2].height=24;ws.row_dimensions[3].height=8
r=omain(4,ORG['main'][:10],ORG['sections'][0],34)
r=omain(r,ORG['main'][10:16],ORG['sections'][1],34)
PAGE1_END=r-1
r=omain(r,ORG['main'][16:],ORG['sections'][1],29)
om(r,1,8,'СТОП-ФАКТОРЫ',True);r+=1
oc(r,1,'№',True,'center');om(r,2,3,'Стоп-фактор',True);om(r,4,8,'Ответ',True);r+=1
om(r,1,3,'');oc(r,4,'да',True,'center',7);oc(r,8,'нет',True,'center',7);ws.row_dimensions[r].height=24;r+=1
om(r,1,3,'')
for c in range(4,9):oc(r,c,c-3,True,'center')
r+=1
for n,s in enumerate(ORG['sf'],1):
    oc(r,1,n,False,'center');om(r,2,3,s,False,'left');
    for c in range(4,9):oc(r,c)
    ws.row_dimensions[r].height=30;r+=1
om(r,1,8,'«КРАСНЫЕ» ВОПРОСЫ',True);r+=1
oc(r,1,'№',True,'center');om(r,2,3,'Вопрос',True);om(r,4,8,'Ответ',True);r+=1
om(r,1,3,'');oc(r,4,'да',True,'center',7);oc(r,8,'нет',True,'center',7);ws.row_dimensions[r].height=24;r+=1
om(r,1,3,'')
for c in range(4,9):oc(r,c,c-3,True,'center')
r+=1
for n,s in enumerate(ORG['kv'],1):
    oc(r,1,n,False,'center');om(r,2,3,s,False,'left')
    for c in range(4,9):oc(r,c)
    ws.row_dimensions[r].height=26;r+=1
ws.row_breaks.append(Break(id=PAGE1_END))
ws.page_setup.orientation='portrait';ws.page_setup.paperSize=ws.PAPERSIZE_A4
ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=2
ws.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
ws.print_options.horizontalCentered=True
ws.page_margins.left=0.35;ws.page_margins.right=0.35;ws.page_margins.top=0.35;ws.page_margins.bottom=0.35
ws.print_area='A1:H%d'%(r-1);ws.print_title_rows=None
ws.sheet_view.showGridLines=False

# защита расчётных листов (ввод — только анкеты, калибровка, портфель и свободные поля заключения)
for name in ('Методика', 'Чек-лист_22', 'Свод', 'Обоснованность', 'Панель', 'Анкета_бланк'):
    wb[name].protection.sheet = True
    wb[name].protection.password = ''
for cell in ('B20', 'B21', 'B24'):
    wb['Заключение'][cell].protection = openpyxl.styles.Protection(locked=False)
wb['Заключение'].protection.sheet = True
for rr in list(range(4, 11)) + [26, 27, 28]:
    wb['Калибровка'].cell(row=rr, column=2).protection = openpyxl.styles.Protection(locked=False)
wb['Калибровка'].protection.sheet = True
for i in range(PMAX):
    rr = 5 + i
    for j in (2, 3, 4, 5, 6, 7, 10, 12, 13):
        wb['Портфель'].cell(row=rr, column=j).protection = openpyxl.styles.Protection(locked=False)
wb['Портфель'].protection.sheet = True
ORDER = ['Методика', 'Чек-лист_22', 'Калибровка'] + ['Э%d' % e for e in range(1, NEXP + 1)] + ['Свод', 'Обоснованность', 'Панель', 'Портфель', 'Заключение', 'Анкета_бланк']
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0

wb.save(OUT)
print('OK ->', os.path.abspath(OUT))
print('листов:', wb.sheetnames)
