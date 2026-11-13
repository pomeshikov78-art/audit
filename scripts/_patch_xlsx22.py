# -*- coding: utf-8 -*-
"""Этап 22 (v22.0): патч gen_10_xlsx_21.py -> gen_10_xlsx_22.py.
Новый лист «Обоснованность» (19-й): математический аппарат обоснованности решения.
Базовое правило решения не изменяется; слой совещательный."""
import os, re
HERE = os.path.dirname(os.path.abspath(__file__))
t = open(os.path.join(HERE, 'gen_10_xlsx_21.py'), encoding='utf-8').read()

def rep(old, new, cnt=1):
    global t
    n = t.count(old)
    assert n == cnt, 'COUNT %d != %d: %r' % (n, cnt, old[:80])
    t = t.replace(old, new)

# ---- версии ----
rep('"""Этап 21 (v21.0), файл 10.', '"""Этап 22 (v22.0), файл 10.')
rep('Группа до 10 экспертов. 18 листов:', 'Группа до 10 экспертов. 19 листов:')
rep('Свод · Панель · Портфель · Заключение · Анкета_бланк.', 'Свод · Обоснованность · Панель · Портфель · Заключение · Анкета_бланк.')
rep('10_Чек-лист_ЭО_НИОКР_22_вопроса_v21.xlsx', '10_Чек-лист_ЭО_НИОКР_22_вопроса_v22.xlsx')
rep('МЕТОДИКА ЭО-НИОКР v1.0 + расширения v2.0 — экспертная оценка', 'МЕТОДИКА ЭО-НИОКР v2.1 (базис v1.0, расширения v2.0/v2.1) — экспертная оценка')
rep('модель v21.0', 'модель v22.0')
rep('расчётное ядро v21.0', 'расчётное ядро v22.0')
rep('(v21.0)', '(v22.0)')
n = t.count('ЭО-НИОКР v2.0'); t = t.replace('ЭО-НИОКР v2.0', 'ЭО-НИОКР v2.1'); print('ЭО-НИОКР v2.0->v2.1:', n)
n = t.count('(v21)'); t = t.replace('(v21)', '(v22)'); print('(v21)->(v22):', n)

# ---- расширение в списке методики ----
rep("    'Состав 22 вопросов, 5 стоп-факторов, 9 «красных» вопросов, шкала и якоря — НЕИЗМЕННЫ (базис ЭО-НИОКР v1.0). Расширения не подменяют решение комиссии.'",
    "    'Обоснованность решения (v22.0, лист «Обоснованность»): совещательный статистический слой — 95% доверительные интервалы индексов (Стьюдент, df=n−1; в режиме весов — Kish n_eff), устойчивость зон, коэффициент конкордации Кендалла W с поправкой на связки и χ²-значимостью, jackknife-устойчивость рекомендации, повопросные медианы и ДИ; ворота G1–G5 и вердикт. Базовая рекомендация (Свод!B81) НЕ изменяется; решение принимает комиссия. Методика ЭО-НИОКР v2.1, раздел 21, приложение 13.',\n    'Состав 22 вопросов, 5 стоп-факторов, 9 «красных» вопросов, шкала и якоря — НЕИЗМЕННЫ (базис ЭО-НИОКР v1.0). Расширения не подменяют решение комиссии.'")

# ---- блок Свода + лист Обоснованность ----
OB = """# ---- 6. Обоснованность решения (v22.0, совещательный слой) ----
hdr(ws, 85, 4, '6. Обоснованность решения (v22.0, совещательный слой — лист «Обоснованность»)')
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
title(ws, 'ОБОСНОВАННОСТЬ РЕШЕНИЯ — математический аппарат v22.0 (ЭО-НИОКР v2.1 §21): ДИ, W Кендалла, χ², jackknife, ворота G1–G5', 12)
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

# ============================ ЛИСТ «Панель» ============================"""
rep("ws.sheet_view.showGridLines = False\n\n# ============================ ЛИСТ «Панель» ============================", OB)

# ---- Панель: строка вердикта ----
rep("ws.merge_cells('B8:E8')\nhdr(ws, 10, 4, 'Профиль 22 критериев",
    "ws.merge_cells('B8:E8')\nws['A9'] = 'ОБОСНОВАННОСТЬ (v22.0):'; ws['A9'].font = F_B\nws['B9'] = '=Обоснованность!B84'; ws['B9'].font = Font(bold=True, size=11, color=NAVY); ws['B9'].alignment = WRAP\nws.merge_cells('B9:E9'); ws.row_dimensions[9].height = 28\nhdr(ws, 10, 4, 'Профиль 22 критериев")

# ---- защита и порядок ----
rep("for name in ('Методика', 'Чек-лист_22', 'Свод', 'Панель', 'Анкета_бланк'):",
    "for name in ('Методика', 'Чек-лист_22', 'Свод', 'Обоснованность', 'Панель', 'Анкета_бланк'):")
rep("+ ['Свод', 'Панель', 'Портфель', 'Заключение', 'Анкета_бланк']",
    "+ ['Свод', 'Обоснованность', 'Панель', 'Портфель', 'Заключение', 'Анкета_бланк']")

assert not re.search(r'[一-鿿]', t), 'CJK!'
open(os.path.join(HERE, 'gen_10_xlsx_22.py'), 'w', encoding='utf-8').write(t)
print('gen_10_xlsx_22.py written:', len(t), 'chars')