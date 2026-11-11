# -*- coding: utf-8 -*-
"""Порождает gen_10_xlsx_21.py из gen_10_xlsx_20.py (этап 21, v21.0).
18 листов: +Калибровка, +Портфель, +Анкета_бланк; Э: вес D3 + блок Р65–70 + I_R D72;
Свод: I_R/вес/статус D/E/F, взвешенные групповые и σ (помощники W:AE, скрыты),
зоны и флаг разногласия от калибровки, модификатор I_R (B81)."""
import io, os
HERE = os.path.dirname(os.path.abspath(__file__))
s = io.open(os.path.join(HERE, 'gen_10_xlsx_20.py'), encoding='utf-8').read()
n0 = len(s)

def rep(a, b):
    global s
    assert s.count(a) == 1, 'NOT FOUND or non-unique: %r' % a[:90]
    s = s.replace(a, b)

# 1. шапка
rep('''"""Этап 20 (v20.0), файл 10. Excel-калькулятор и подробный чек-лист методики
ЭО-НИОКР v1.0 (22 вопроса, единая шкала 1–5 + Н/О, 5 стоп-факторов, 9 «красных»,
два индекса I_A/I_C, матрица 3×3, σ-согласованность, оверрайд стоп-фактора).
Группа до 10 экспертов. 16 листов: Методика · Чек-лист_22 · Э1–Э10 · Свод · Панель · Заключение.
Вся логика — формулами, без макросов (Excel 365/2021)."""''',
'''"""Этап 21 (v21.0), файл 10. Excel-калькулятор и подробный чек-лист методики
ЭО-НИОКР v1.0 + расширения v2.0 (22 вопроса, шкала 1–5 + Н/О, 5 стоп-факторов, 9 «красных»,
I_A/I_C, матрица 3×3, σ-согласованность, оверрайд стоп-фактора — базис v1.0 неизменен).
Расширения v2.0 (opt-in, лист «Калибровка»): дифференцированные веса экспертов,
индекс результативности I_R (блок Р1–Р6, минимум 5 из 6), калибровка порогов по итогам
апробации, портфель до 30 НИОКР, печатная анкета (прил. №1).
Группа до 10 экспертов. 18 листов: Методика · Чек-лист_22 · Калибровка · Э1–Э10 ·
Свод · Панель · Портфель · Заключение · Анкета_бланк.
Вся логика — формулами, без макросов (Excel 365/2021)."""''')

# 2. данные + имя файла
rep("""D = json.load(io.open(os.path.join(HERE, 'data', 'qeo22.json'), encoding='utf-8'))
OUT = os.path.join(HERE, '..', 'output', '10_Чек-лист_ЭО_НИОКР_22_вопроса_v20.xlsx')""",
"""D = json.load(io.open(os.path.join(HERE, 'data', 'qeo22.json'), encoding='utf-8'))
D21 = json.load(io.open(os.path.join(HERE, 'data', 'qeo_v21.json'), encoding='utf-8'))
RB = D21['rblock']['q']            # блок Р1–Р6 (расширение v2.0)
W21 = D21['weights']               # веса экспертов: default/min/max/step
CAL0 = D21['calib']                # базовые пороги зон и σ
PMAX = D21['portfolio']['max']     # ёмкость портфеля (30)
OUT = os.path.join(HERE, '..', 'output', '10_Чек-лист_ЭО_НИОКР_22_вопроса_v21.xlsx')""")

# 3. Методика: заголовок
rep("title(ws, 'МЕТОДИКА ЭО-НИОКР v1.0 — экспертная оценка актуальности и целесообразности НИОКР (модель v20.0, группа до 10 экспертов, 11.2026)')",
    "title(ws, 'МЕТОДИКА ЭО-НИОКР v1.0 + расширения v2.0 — экспертная оценка актуальности, целесообразности и результативности НИОКР (модель v21.0, группа до 10 экспертов, портфель до 30 НИОКР, 11.2026)')")

# 4. Методика: веса в разделе 4
rep("'Экспертная группа: от 3 до 10 специалистов (рекомендуется 5); веса экспертов равные.',",
    "'Экспертная группа: от 3 до 10 специалистов (рекомендуется 5); веса равные (базовый режим v1.0); дифференцированные веса — расширение v2.0 (лист «Калибровка», B9).',")

# 5. Методика: раздел 9 (расширения v2.0)
rep("""ws.column_dimensions['A'].width = 16
ws.column_dimensions['B'].width = 46""",
"""r += 1
hdr(ws, r, 3, '9. Расширения v2.0 (лист «Калибровка»; по умолчанию ОТКЛЮЧЕНЫ — полная совместимость с v1.0)'); r += 1
for t in (
    'Веса экспертов (Калибровка!B9 = Да): w ∈ [0,25; 3,00], шаг 0,25; ввод на листах Э1–Э10 (ячейка D3). Групповой индекс = Σ(w_k·I_k)/Σw_k по экспертам с действующим индексом направления; σ = √(Σw_k·(I_k−I_ср)²/Σw_k). Повопросная статистика — без весов. Основание весов фиксируется в протоколе комиссии.',
    'Индекс результативности I_R (Калибровка!B10 = Да): блок Р1–Р6 на листах Э (строки 65–70), единая шкала 1–5 + Н/О; минимум 5 оценённых из 6. I_R = ΣР/N_Р; зоны — общие с I_A/I_C. При I_R в зоне Н и рекомендации «Продолжить НИОКР» формируется модификатор «с планом повышения результативности»; оверрайд стоп-фактора имеет приоритет над модификатором.',
    'Калибровка порогов (Калибровка!B4:B8): корректировка зон В/С и границ σ допускается ТОЛЬКО по итогам апробации (пилот ≥ 5 НИОКР) с оформлением протокола (раздел 18 методики ЭО-НИОКР v2.0). Базовые значения: В ≥ 4,00; С ≥ 3,00; σ 0,50 / 0,80 / 1,00.',
    'Портфель НИОКР (лист «Портфель»): реестр до 30 проектов, сравнительная матрица зон I_A/I_C, сводная статистика; групповые значения переносятся из листа «Свод» (строка-помощник 45) после оценки каждой НИОКР. В интерактивной HTML-модели портфель ведётся автоматически.',
    'Бумажная процедура (лист «Анкета_бланк»): печатная форма индивидуальной анкеты — приложение №1 к методике; заполняется экспертом независимо (раунд 1), затем секретарь переносит баллы в листы Э1–Э10.',
    'Состав 22 вопросов, 5 стоп-факторов, 9 «красных» вопросов, шкала и якоря — НЕИЗМЕННЫ (базис ЭО-НИОКР v1.0). Расширения не подменяют решение комиссии.'
):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=1, value='• ' + t).alignment = WRAP; r += 1
r += 1
ws.column_dimensions['A'].width = 16
ws.column_dimensions['B'].width = 46""")

# 6. Э: заголовок
rep("title(ws, 'АНКЕТА ЭКСПЕРТНОЙ ОЦЕНКИ НИОКР — эксперт %d (ЭО-НИОКР v1.0, 22 вопроса)' % e, 6)",
    "title(ws, 'АНКЕТА ЭКСПЕРТНОЙ ОЦЕНКИ НИОКР — эксперт %d (ЭО-НИОКР v1.0 + расширения v2.0; 22 вопроса + блок Р)' % e, 6)")

# 7. Э: вес D3
rep("""    for cc in ('B3', 'B4', 'B5', 'D5', 'F5'):
        ws[cc].fill = FILL_IN""",
"""    ws['C3'] = 'Вес эксперта (v2.0):'; ws['D3'] = W21['default']; ws['C3'].font = F_B
    for cc in ('B3', 'B4', 'B5', 'D5', 'F5', 'D3'):
        ws[cc].fill = FILL_IN""")

# 8. Э: блок Р перед валидациями
rep("""    # валидации
    dv = DataValidation(type='list', formula1='"1,2,3,4,5,Н/О"', allow_blank=True,""",
"""    # блок Р — результативность (расширение v2.0)
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
    dv = DataValidation(type='list', formula1='"1,2,3,4,5,Н/О"', allow_blank=True,""")

# 9. Э: валидации блока Р и веса
rep("    ws.add_data_validation(dv3); dv3.add('C53:C61')",
"""    ws.add_data_validation(dv3); dv3.add('C53:C61')
    dv4 = DataValidation(type='list', formula1='"1,2,3,4,5,Н/О"', allow_blank=True,
                         showErrorMessage=True, errorTitle='Недопустимое значение',
                         error='Допустимо: 1, 2, 3, 4, 5 или Н/О')
    ws.add_data_validation(dv4); dv4.add('D65:D70')
    dv5 = DataValidation(type='decimal', operator='between', formula1=str(W21['min']), formula2=str(W21['max']),
                         allow_blank=True, showErrorMessage=True, errorTitle='Недопустимый вес',
                         error='Вес эксперта: от 0,25 до 3,00 (шаг 0,25)')
    ws.add_data_validation(dv5); dv5.add('D3')""")

# 10. Свод: заголовок и шапка раздела 1
rep("title(ws, 'СВОДНАЯ ЭКСПЕРТНАЯ ОЦЕНКА — расчётное ядро (групповые индексы, σ, сигналы, матрица; до 10 экспертов)', 22)",
    "title(ws, 'СВОДНАЯ ЭКСПЕРТНАЯ ОЦЕНКА — расчётное ядро v21.0 (I_A/I_C/I_R, веса, σ, сигналы, матрица; до 10 экспертов)', 22)")
rep("hdr(ws, 3, 4, '1. Индивидуальные и групповые индексы (до 10 экспертов)')",
    "hdr(ws, 3, 6, '1. Индивидуальные и групповые индексы (до 10 экспертов; I_R и веса — расширение v2.0)')")

# 11. Свод: таблица экспертов + групповые + σ + согласованность (взвешенный режим)
OLD_BLK = """for j, t in enumerate(('Эксперт', 'I_A (актуальность)', 'I_C (целесообразность)', 'Статус'), 1):
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
                          'IF(%s%d<=1,"Умеренные разногласия","Существенные разногласия"))))' % (col, SIG, col, SIG, col, SIG, col, SIG))"""
NEW_BLK = """for j, t in enumerate(('Эксперт', 'I_A (актуальность)', 'I_C (целесообразность)', 'I_R (v2.0)', 'Вес (v2.0)', 'Статус'), 1):
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
                          'IF(%s%d<=Калибровка!$B$8,"Умеренные разногласия","Существенные разногласия"))))' % (col, SIG, col, SIG, col, SIG, col, SIG))"""
rep(OLD_BLK, NEW_BLK)

# 12. Свод: флаг разногласия от калибровки
rep("ws.cell(row=rr, column=FLG, value='=IF(AND(ISNUMBER(%s),%s>1),\"РАЗНОГЛАСИЕ\",\"\")' % (LC(SD, rr), LC(SD, rr))).alignment = WRAPC",
    "ws.cell(row=rr, column=FLG, value='=IF(AND(ISNUMBER(%s),%s>Калибровка!$B$8),\"РАЗНОГЛАСИЕ\",\"\")' % (LC(SD, rr), LC(SD, rr))).alignment = WRAPC")
rep("ws.cell(row=QCNT, column=1, value='Вопросов с разногласиями (σ>1,00):').font = F_B",
    "ws.cell(row=QCNT, column=1, value='Вопросов с разногласиями (σ>порога калибровки):').font = F_B")
rep("ws.cell(row=DISR, column=1, value='Перечень вопросов с разногласиями (σ>1,00):'); ws.cell(row=DISR, column=1).font = F_B",
    "ws.cell(row=DISR, column=1, value='Перечень вопросов с разногласиями (σ>порога):'); ws.cell(row=DISR, column=1).font = F_B")

# 13. Свод: зоны от калибровки + зона I_R (строка 73)
rep("""ws.cell(row=ZA, column=1, value='Зона I_A:')
ws.cell(row=ZA, column=2, value='=IF(NOT(ISNUMBER(B%d)),"н/д",IF(B%d>=4,"В",IF(B%d>=3,"С","Н")))' % (GRP, GRP, GRP))
ws.cell(row=ZC, column=1, value='Зона I_C:')
ws.cell(row=ZC, column=2, value='=IF(NOT(ISNUMBER(C%d)),"н/д",IF(C%d>=4,"В",IF(C%d>=3,"С","Н")))' % (GRP, GRP, GRP))""",
"""ws.cell(row=ZA, column=1, value='Зона I_A:')
ws.cell(row=ZA, column=2, value='=IF(NOT(ISNUMBER(B%d)),"н/д",IF(B%d>=Калибровка!$B$4,"В",IF(B%d>=Калибровка!$B$5,"С","Н")))' % (GRP, GRP, GRP))
ws.cell(row=ZC, column=1, value='Зона I_C:')
ws.cell(row=ZC, column=2, value='=IF(NOT(ISNUMBER(C%d)),"н/д",IF(C%d>=Калибровка!$B$4,"В",IF(C%d>=Калибровка!$B$5,"С","Н")))' % (GRP, GRP, GRP))
ws.cell(row=73, column=1, value='Зона I_R (v2.0):')
ws.cell(row=73, column=2, value='=IF(NOT(ISNUMBER(D%d)),"н/д",IF(D%d>=Калибровка!$B$4,"В",IF(D%d>=Калибровка!$B$5,"С","Н")))' % (GRP, GRP, GRP))""")

# 14. Свод: модификатор I_R (строка 81)
rep("""ws.cell(row=OVR, column=2).font = F_B
lowcat = '&'.join(LC(LOWH, Q0 + i) for i in range(22))""",
"""ws.cell(row=OVR, column=2).font = F_B
ws.cell(row=81, column=1, value='Рекомендация с учётом результативности (v2.0):'); ws.cell(row=81, column=1).font = F_B
ws.cell(row=81, column=2, value='=IF(AND(Калибровка!$B$10="Да",B73="Н",B80="Продолжить НИОКР"),"Продолжить НИОКР с планом повышения результативности (I_R ниже зоны С)",B80)')
ws.cell(row=81, column=2).font = F_B
lowcat = '&'.join(LC(LOWH, Q0 + i) for i in range(22))""")

# 15. Панель: строка 5 (I_R) + итоговая рекомендация B8 ← B81
rep("ws['A6'] = 'СТОП-ФАКТОРЫ:'; ws['A6'].font = F_B; ws['B6'] = '=Свод!B54'",
"""ws['A5'] = 'Индекс результативности I_R (v2.0)'; ws['A5'].font = F_B
ws['B5'] = '=IF(ISNUMBER(Свод!D15),Свод!D15,"—")'; ws['B5'].number_format = '0.00'
ws['C5'] = '=Свод!B73'; ws['D5'] = '=Свод!D17'
ws['A6'] = 'СТОП-ФАКТОРЫ:'; ws['A6'].font = F_B; ws['B6'] = '=Свод!B54'""")
rep("ws['B8'] = '=Свод!B80'; ws['B8'].font = Font(bold=True, size=13, color=NAVY)",
    "ws['B8'] = '=Свод!B81'; ws['B8'].font = Font(bold=True, size=13, color=NAVY)")

# 16. Заключение: раздел 8 — итог с учётом I_R и весов
rep("""&" ("&Свод!B72&")."'),""",
    """&" ("&Свод!B72&")."&IF(Калибровка!B9="Да"," Применены дифференцированные веса экспертов (расширение v2.0).","")&IF(Калибровка!B10="Да"," Итоговая рекомендация с учётом результативности: "&Свод!B81&".","")'),""")

# 17. Заключение: разделы 4а и 8а
rep("""for rr, h, f in rows:
    ws.cell(row=rr, column=1, value=h).font = F_SEC
    c = ws.cell(row=rr, column=2, value=f); c.alignment = WRAP
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
    ws.row_dimensions[rr].height = 62""",
"""rows21 = [
 (10, '4а. Оценка результативности (v2.0)',
  '=IF(Калибровка!B10<>"Да","Режим оценки результативности отключён (базовая процедура ЭО-НИОКР v1.0).",IF(NOT(ISNUMBER(Свод!D15)),"Индекс результативности I_R не рассчитан (менее 5 из 6 вопросов Р у экспертов с действующей оценкой).","I_R="&ROUND(Свод!D15,2)&" балла (зона "&Свод!B73&"). "&IF(Свод!B73="Н","Результативность ниже зоны С: рекомендуется план повышения результативности (см. итоговую рекомендацию).","Результативность соответствует требованиям.")))'),
 (18, '8а. Режим расчёта и калибровка',
  '=IF(AND(Калибровка!B4=4,Калибровка!B5=3,Калибровка!B6=0.5,Калибровка!B7=0.8,Калибровка!B8=1),"Пороги зон и согласованности — базовые значения ЭО-НИОКР v1.0 (калибровка не применялась).","Пороги калиброваны по итогам апробации (лист «Калибровка»): В≥"&Калибровка!B4&"; С≥"&Калибровка!B5&"; σ: "&Калибровка!B6&" / "&Калибровка!B7&" / "&Калибровка!B8&".")&IF(Калибровка!B9="Да"," Веса экспертов: дифференцированные."," Веса экспертов: равные.")&IF(Калибровка!B10="Да"," Блок Р: включён."," Блок Р: отключён.")'),
]
for rr, h, f in rows + rows21:
    ws.cell(row=rr, column=1, value=h).font = F_SEC
    c = ws.cell(row=rr, column=2, value=f); c.alignment = WRAP
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
    ws.row_dimensions[rr].height = 62""")

# 18. Новые листы + защита + порядок
rep("""# защита расчётных листов (ввод — только анкеты и свободные поля заключения)
for name in ('Методика', 'Чек-лист_22', 'Свод', 'Панель'):
    wb[name].protection.sheet = True
    wb[name].protection.password = ''
for cell in ('B20', 'B21', 'B24'):
    wb['Заключение'][cell].protection = openpyxl.styles.Protection(locked=False)
wb['Заключение'].protection.sheet = True

wb.save(OUT)""",
"""# ============================ ЛИСТ «Калибровка» (v21) ============================
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
ws['A23'] = 'Корректировка порогов допускается только по итогам апробации (пилот ≥ 5 НИОКР) с оформлением протокола — раздел 18 методики ЭО-НИОКР v2.0.'
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

# ============================ ЛИСТ «Портфель» (v21) ============================
ws = wb.create_sheet('Портфель')
title(ws, 'ПОРТФЕЛЬ НИОКР — реестр до 30 проектов, сравнительная матрица зон I_A/I_C (v21.0)', 13)
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

# ============================ ЛИСТ «Анкета_бланк» (v21) ============================
from openpyxl.worksheet.properties import PageSetupProperties
ws = wb.create_sheet('Анкета_бланк')
title(ws, 'ПРИЛОЖЕНИЕ №1 — ИНДИВИДУАЛЬНАЯ АНКЕТА ЭКСПЕРТНОЙ ОЦЕНКИ НИОКР (бумажная процедура, ЭО-НИОКР v2.0)', 4)
ws['A2'] = 'Заполняется экспертом независимо (раунд 1). Баллы — профессиональное суждение; документальное подтверждение не требуется. При колебании между баллами выбирайте более низкий (принцип осторожности).'
ws['A2'].alignment = WRAP; ws.merge_cells('A2:D2'); ws.row_dimensions[2].height = 30
ws['A3'] = 'Наименование НИОКР: ______________________________________'; ws.merge_cells('A3:B3')
ws['C3'] = 'Шифр: ____________'
ws['A4'] = 'Эксперт (Ф.И.О.): ________________________________________'; ws.merge_cells('A4:B4')
ws['C4'] = 'Дата: ____________'
ws['A5'] = 'Компетенция: ____________________________________________'; ws.merge_cells('A5:B5')
ws['C5'] = 'Вес (v2.0): ______'
ws['A6'] = 'Шкала: 1 — крайне низкая; 2 — низкая; 3 — средняя; 4 — высокая; 5 — очень высокая; Н/О — не оценивается (не балл, в расчёт не входит).'
ws['A6'].alignment = WRAP; ws.merge_cells('A6:D6'); ws.row_dimensions[6].height = 28
r = 8
def frm_section(r, name, qs):
    ws.cell(row=r, column=1, value=name).font = F_SEC; r += 1
    for j, t in enumerate(('Код', 'Критерий / вопрос', 'Балл (1–5 / Н/О)', 'Комментарий'), 1):
        c = ws.cell(row=r, column=j, value=t); c.font = F_H; c.fill = FILL_H; c.alignment = WRAPC
    r += 1
    for q in qs:
        ws.cell(row=r, column=1, value=q['code']).alignment = WRAPC
        ws.cell(row=r, column=2, value=q['crit'] + '. ' + q['q']).alignment = WRAP
        for j in (1, 2, 3, 4): ws.cell(row=r, column=j).border = BORD
        ws.row_dimensions[r].height = 34
        r += 1
    return r + 1
r = frm_section(r, 'РАЗДЕЛ А. АКТУАЛЬНОСТЬ — 10 вопросов (минимум 8 оценённых)', QA)
r = frm_section(r, 'РАЗДЕЛ Б. ЦЕЛЕСООБРАЗНОСТЬ — 12 вопросов (минимум 10 оценённых)', QC)
ws.cell(row=r, column=1, value='СТОП-ФАКТОРЫ — обведите один ответ по каждому фактору').font = F_SEC; r += 1
for i, sfq in enumerate(SF, 1):
    ws.cell(row=r, column=1, value='СФ%d' % i).alignment = WRAPC
    ws.cell(row=r, column=2, value=sfq).alignment = WRAP
    ws.cell(row=r, column=3, value='Да / Нет / Неопр.').alignment = WRAPC
    for j in (1, 2, 3): ws.cell(row=r, column=j).border = BORD
    ws.row_dimensions[r].height = 30
    r += 1
r += 1
ws.cell(row=r, column=1, value='«КРАСНЫЕ» ВОПРОСЫ — балл 1–5 (1=Да, 2=Возможно да, 3=Неопределённо, 4=Возможно нет, 5=Нет)').font = F_SEC; r += 1
for k in KV:
    ws.cell(row=r, column=1, value='КВ%d' % k['n']).alignment = WRAPC
    ws.cell(row=r, column=2, value=k['q']).alignment = WRAP
    for j in (1, 2, 3, 4): ws.cell(row=r, column=j).border = BORD
    ws.row_dimensions[r].height = 28
    r += 1
r += 1
r = frm_section(r, 'РАЗДЕЛ В. РЕЗУЛЬТАТИВНОСТЬ — 6 вопросов (расширение v2.0; заполняется, если включён блок Р)', RB)
ws.cell(row=r, column=1, value='Особое мнение / комментарии:').font = F_B; r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=4)
for rr2 in range(r, r + 3):
    for j in range(1, 5): ws.cell(row=rr2, column=j).border = BORD
r += 4
ws.cell(row=r, column=1, value='Подпись эксперта: ____________________    Дата: ____________').font = F_B
ws.column_dimensions['A'].width = 8
ws.column_dimensions['B'].width = 72
ws.column_dimensions['C'].width = 18
ws.column_dimensions['D'].width = 30
ws.page_setup.orientation = 'portrait'
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws.print_area = 'A1:D%d' % r
ws.sheet_view.showGridLines = False

# защита расчётных листов (ввод — только анкеты, калибровка, портфель и свободные поля заключения)
for name in ('Методика', 'Чек-лист_22', 'Свод', 'Панель', 'Анкета_бланк'):
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
ORDER = ['Методика', 'Чек-лист_22', 'Калибровка'] + ['Э%d' % e for e in range(1, NEXP + 1)] + ['Свод', 'Панель', 'Портфель', 'Заключение', 'Анкета_бланк']
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0

wb.save(OUT)""")

io.open(os.path.join(HERE, 'gen_10_xlsx_21.py'), 'w', encoding='utf-8').write(s)
print('gen_10_xlsx_21.py written: %d -> %d chars' % (n0, len(s)))