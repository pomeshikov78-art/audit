# -*- coding: utf-8 -*-
"""Этап 22 (v22.0): патч verify_10_xlsx_21.py -> verify_10_xlsx_22.py.
Независимый пересчёт листа «Обоснованность» (formulas) + эталоны Python."""
import os, re
HERE = os.path.dirname(os.path.abspath(__file__))
t = open(os.path.join(HERE, 'verify_10_xlsx_21.py'), encoding='utf-8').read()

def rep(old, new, cnt=1):
    global t
    n = t.count(old)
    assert n == cnt, 'COUNT %d != %d: %r' % (n, cnt, old[:80])
    t = t.replace(old, new)

rep('SRC = os.path.join(HERE, \'..\', \'output\', \'10_Чек-лист_ЭО_НИОКР_22_вопроса_v21.xlsx\')',
    'SRC = os.path.join(HERE, \'..\', \'output\', \'10_Чек-лист_ЭО_НИОКР_22_вопроса_v22.xlsx\')')
rep("TMP = '/tmp/eo21_case.xlsx'", "TMP = '/tmp/eo22_case.xlsx'")
rep('pref = "\'[eo21_case.xlsx]%s\'!%s"', 'pref = "\'[eo22_case.xlsx]%s\'!%s"')
rep('eq(\'18 листов\', len(wbs.sheetnames), 18)', 'eq(\'19 листов\', len(wbs.sheetnames), 19)')
rep("ck('новые листы присутствуют', all(n in wbs.sheetnames for n in ('Портфель', 'Анкета_бланк')))",
    "ck('новые листы присутствуют', all(n in wbs.sheetnames for n in ('Портфель', 'Анкета_бланк', 'Обоснованность')))")

# ---- эталонные функции Python ----
REF = '''    return {'A': A, 'C': C, 'SF': SF, 'KV': KV, 'R': r, 'w': w}

# ================= эталоны Python для листа «Обоснованность» (v22) =================
import math, statistics
from collections import Counter
TCRIT = [12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262]
CHI2 = {9: 16.919, 11: 19.675, 21: 32.671}
_wb0 = openpyxl.load_workbook(SRC)
_ws0 = _wb0['Свод']
MX = {}
for _i, _za in enumerate(('В', 'С', 'Н')):
    MX[_za] = {}
    for _j, _zc in enumerate(('В', 'С', 'Н')):
        MX[_za][_zc] = _ws0.cell(row=75 + _i, column=2 + _j).value

def median_ref(a):
    return statistics.median(a) if a else None

def stdev_s_ref(a):
    return statistics.stdev(a) if len(a) >= 2 else None

def pop_sd(a):
    m = sum(a) / len(a)
    return math.sqrt(sum((x - m) ** 2 for x in a) / len(a))

def ci_ref(vals, weights=None):
    n = len(vals)
    if n < 2:
        return None
    if weights:
        sw = sum(weights); sw2 = sum(w * w for w in weights)
        m = sum(w * v for w, v in zip(weights, vals)) / sw
        v = sum(w * (x - m) ** 2 for w, x in zip(weights, vals))
        den = sw - sw2 / sw
        neff = sw * sw / sw2
        df = max(1, min(9, int(neff) - 1))
        se = math.sqrt(v / den * sw2) / sw if den > 0 else None
    else:
        m = sum(vals) / n
        se = stdev_s_ref(vals) / math.sqrt(n)
        df = n - 1
        neff = n
    if se is None:
        return dict(m=m, se=None, df=df, lo=None, hi=None, neff=neff)
    tc = TCRIT[min(9, df) - 1]
    return dict(m=m, se=se, df=df, t=tc, lo=m - tc * se, hi=m + tc * se, neff=neff)

def zone_ref(v, zv=4.0, zs=3.0):
    return 'В' if v >= zv else ('С' if v >= zs else 'Н')

def kendall_w_ref(rows):
    m = len(rows)
    if m < 2:
        return None
    k = len(rows[0])
    if k < 2:
        return None
    for r in rows:
        if len(r) != k or any(not isinstance(x, (int, float)) for x in r):
            return None
    Rj = [0.0] * k
    Tsum = 0
    for r in rows:
        ranks = []
        for j in range(k):
            less = sum(1 for u in range(k) if r[u] < r[j])
            eq = sum(1 for u in range(k) if r[u] == r[j])
            ranks.append(less + (eq + 1) / 2)
        for j in range(k):
            Rj[j] += ranks[j]
        for c in Counter(r).values():
            Tsum += c ** 3 - c
    Rbar = m * (k + 1) / 2
    S = sum((x - Rbar) ** 2 for x in Rj)
    den = m * m * (k ** 3 - k) - m * Tsum
    if den <= 0:
        return dict(W=None, chi2=None, df=k - 1, m=m, k=k, sig=False)
    W = 12 * S / den
    chi2 = m * (k - 1) * W
    crit = CHI2.get(k - 1)
    return dict(W=W, chi2=chi2, df=k - 1, crit=crit, sig=(chi2 >= crit) if crit else None, m=m, k=k)

def jack_ref(IA_e, IC_e, yes_e, zv=4.0, zs=3.0):
    n = len(IA_e)
    IA = sum(IA_e) / n; IC = sum(IC_e) / n
    yes_total = sum(yes_e)
    def rec(ia, ic, yes):
        r = MX[zone_ref(ia, zv, zs)][zone_ref(ic, zv, zs)]
        if yes > 0 and r == 'Продолжить НИОКР':
            return 'ПОНИЖЕНИЕ: Пересмотреть параметры НИОКР (стоп-фактор)'
        return r
    full = rec(IA, IC, yes_total)
    runs = []
    for i in range(n):
        ia = (n * IA - IA_e[i]) / (n - 1)
        ic = (n * IC - IC_e[i]) / (n - 1)
        r = rec(ia, ic, yes_total - yes_e[i])
        runs.append(dict(IA=ia, IC=ic, zA=zone_ref(ia, zv, zs), zC=zone_ref(ic, zv, zs), rec=r, changed=(r != full)))
    return dict(IA=IA, IC=IC, full=full, runs=runs, changes=sum(x['changed'] for x in runs))

'''
rep("    return {'A': A, 'C': C, 'SF': SF, 'KV': KV, 'R': r, 'w': w}\n\nprint('К-1:", REF + "print('К-1:")

# ---- новые случаи К-17…К-22 ----
CASES = '''
# ================= К-17…К-22: лист «Обоснованность» (v22) =================
print('К-17: демо-заполнение (3 эксперта) → повопросные медиана/s/ДИ, ДИ индексов, W, jackknife, ворота — независимый пересчёт')
A1 = [5, 4, 3, 4, 5, 4, 3, 4, 5, 4]; C1 = [4, 3, 4, 5, 4, 3, 4, 5, 4, 3, 4, 5]
A2 = [4, 3, 4, 5, 4, 3, 4, 5, 4, 3]; C2 = [3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4]
A3 = [3, 4, 5, 4, 3, 4, 5, 4, 3, 4]; C3 = [4, 4, 3, 3, 4, 4, 3, 3, 4, 4, 3, 3]
As = [A1, A2, A3]; Cs = [C1, C2, C3]
g = run({1: exp(a=A1, c=C1), 2: exp(a=A2, c=C2), 3: exp(a=A3, c=C3)})
OB = 'Обоснованность'
IA_e = [sum(a) / 10 for a in As]; IC_e = [sum(c) / 12 for c in Cs]
for i in range(22):
    vals = [As[e][i] if i < 10 else Cs[e][i - 10] for e in range(3)]
    r = 18 + i
    close('К-17 медиана вопроса %d' % (i + 1), g(OB, 'D%d' % r), median_ref(vals), 1e-6)
    ci = ci_ref(vals)
    close('К-17 ДИ нижн. вопроса %d' % (i + 1), g(OB, 'I%d' % r), ci['lo'], 1e-3)
    close('К-17 ДИ верхн. вопроса %d' % (i + 1), g(OB, 'J%d' % r), ci['hi'], 1e-3)
for i in (0, 11, 21):
    vals = [As[e][i] if i < 10 else Cs[e][i - 10] for e in range(3)]
    close('К-17 s вопроса %d' % (i + 1), g(OB, 'E%d' % (18 + i)), stdev_s_ref(vals), 1e-3)
ciA = ci_ref(IA_e); ciC = ci_ref(IC_e)
close('К-17 I_A среднее', g(OB, 'B44'), ciA['m'], 1e-6)
close('К-17 I_A SE', g(OB, 'E44'), ciA['se'], 1e-4)
eq('К-17 I_A df', g(OB, 'F44'), ciA['df'])
close('К-17 I_A ДИ нижн.', g(OB, 'H44'), ciA['lo'], 1e-3)
close('К-17 I_A ДИ верхн.', g(OB, 'I44'), ciA['hi'], 1e-3)
eq('К-17 I_A устойчивость зоны', g(OB, 'L44'), 'УСТОЙЧИВА' if zone_ref(ciA['lo']) == zone_ref(ciA['hi']) else 'НЕУСТОЙЧИВА')
close('К-17 I_C ДИ нижн.', g(OB, 'H45'), ciC['lo'], 1e-3)
close('К-17 I_C ДИ верхн.', g(OB, 'I45'), ciC['hi'], 1e-3)
eq('К-17 I_C устойчивость зоны', g(OB, 'L45'), 'УСТОЙЧИВА' if zone_ref(ciC['lo']) == zone_ref(ciC['hi']) else 'НЕУСТОЙЧИВА')
wall = kendall_w_ref([A1 + C1, A2 + C2, A3 + C3])
wA = kendall_w_ref(As); wC = kendall_w_ref(Cs)
close('К-17 W (22)', g(OB, 'D51'), wall['W'], 1e-4)
close('К-17 χ² (22)', g(OB, 'F51'), wall['chi2'], 1e-3)
eq('К-17 значимость W (22)', g(OB, 'I51'), 'значима' if wall['sig'] else 'не значима')
eq('К-17 полных наборов (22)', g(OB, 'C51'), wall['m'])
close('К-17 W (А)', g(OB, 'D52'), wA['W'], 1e-4)
close('К-17 W (Ц)', g(OB, 'D53'), wC['W'], 1e-4)
jk = jack_ref(IA_e, IC_e, [0, 0, 0])
for i in range(3):
    r = 63 + i
    close('К-17 jackknife I_A(−Э%d)' % (i + 1), g(OB, 'B%d' % r), jk['runs'][i]['IA'], 1e-6)
    close('К-17 jackknife I_C(−Э%d)' % (i + 1), g(OB, 'C%d' % r), jk['runs'][i]['IC'], 1e-6)
    eq('К-17 jackknife рекомендация(−Э%d)' % (i + 1), g(OB, 'F%d' % r), jk['runs'][i]['rec'])
    eq('К-17 jackknife изменение(−Э%d)' % (i + 1), g(OB, 'G%d' % r), 1 if jk['runs'][i]['changed'] else 0)
eq('К-17 jackknife изменений всего', g(OB, 'B73'), jk['changes'])
dis = any(pop_sd([As[e][i] if i < 10 else Cs[e][i - 10] for e in range(3)]) > 1.0 for i in range(22))
g1 = True
g2 = (wall['W'] is not None and wall['W'] >= 0.5 and wall['sig'] is True)
g3 = (zone_ref(ciA['lo']) == zone_ref(ciA['hi'])) and (zone_ref(ciC['lo']) == zone_ref(ciC['hi']))
g4 = (jk['changes'] == 0)
g5 = not dis
okc = sum([g1, g2, g3, g4, g5])
eq('К-17 G1', g(OB, 'B78'), 'ВЫПОЛНЕНО' if g1 else 'НЕ ВЫПОЛНЕНО')
eq('К-17 G2', g(OB, 'B79'), 'ВЫПОЛНЕНО' if g2 else 'НЕ ВЫПОЛНЕНО')
eq('К-17 G3', g(OB, 'B80'), 'ВЫПОЛНЕНО' if g3 else 'НЕ ВЫПОЛНЕНО')
eq('К-17 G4', g(OB, 'B81'), 'ВЫПОЛНЕНО' if g4 else 'НЕ ВЫПОЛНЕНО')
eq('К-17 G5', g(OB, 'B82'), 'ВЫПОЛНЕНО' if g5 else 'НЕ ВЫПОЛНЕНО')
eq('К-17 ворот выполнено', g(OB, 'B83'), okc)
verd = 'Решение статистически обосновано' if okc == 5 else ('Обосновано с оговорками' if okc >= 3 else 'Требуется доработка экспертной оценки')
eq('К-17 вердикт', g(OB, 'B84'), verd)
eq('К-17 Свод: вердикт сквозной', g('Свод', 'B92'), verd)
eq('К-17 Панель: вердикт сквозной', g('Панель', 'B9'), verd)

print('К-18: полное согласие рангов (3 эксперта, паттерн 1–5) → W=1, χ²=63, вердикт 5/5')
pat = [(i % 5) + 1 for i in range(22)]
pa = pat[:10]; pc = pat[10:]
g = run({1: exp(a=pa, c=pc), 2: exp(a=pa, c=pc), 3: exp(a=pa, c=pc)})
close('К-18 W=1', g(OB, 'D51'), 1.0, 1e-6)
close('К-18 χ²=63', g(OB, 'F51'), 63.0, 1e-4)
eq('К-18 χ² значим', g(OB, 'I51'), 'значима')
eq('К-18 ворот 5/5', g(OB, 'B83'), 5)
eq('К-18 вердикт обосновано', g(OB, 'B84'), 'Решение статистически обосновано')
eq('К-18 jackknife без изменений', g(OB, 'B73'), 0)
eq('К-18 зона I_A устойчива', g(OB, 'L44'), 'УСТОЙЧИВА')

print('К-19: неустойчивая зона (I_A 3.5/4.0/4.5) → ДИ пересекает границы, G3 не выполнено')
g = run({1: exp(a=[3, 3, 3, 3, 3, 4, 4, 4, 4, 4], c_all=4), 2: exp(a=[4] * 10, c_all=4), 3: exp(a=[4, 4, 4, 4, 4, 5, 5, 5, 5, 5], c_all=4)})
close('К-19 I_A=4.00', g(OB, 'B44'), 4.0, 1e-6)
ck('К-19 ДИ нижн. < 3', num(g(OB, 'H44')) < 3, g(OB, 'H44'), '<3')
ck('К-19 ДИ верхн. > 4', num(g(OB, 'I44')) > 4, g(OB, 'I44'), '>4')
eq('К-19 зона неустойчива', g(OB, 'L44'), 'НЕУСТОЙЧИВА')
eq('К-19 G3 не выполнено', g(OB, 'B80'), 'НЕ ВЫПОЛНЕНО')

print('К-20: jackknife меняет рекомендацию (E1 I_A=2, E2/E3 I_A=4.5, I_C=4) → G4 не выполнено')
g = run({1: exp(a=[2] * 10, c_all=4), 2: exp(a=[4, 4, 4, 4, 4, 5, 5, 5, 5, 5], c_all=4), 3: exp(a=[4, 4, 4, 4, 4, 5, 5, 5, 5, 5], c_all=4)})
jk20 = jack_ref([2.0, 4.5, 4.5], [4.0, 4.0, 4.0], [0, 0, 0])
ck('К-20 изменений ≥1', num(g(OB, 'B73')) >= 1, g(OB, 'B73'), '>=1')
eq('К-20 изменений = эталон', g(OB, 'B73'), jk20['changes'])
eq('К-20 G4 не выполнено', g(OB, 'B81'), 'НЕ ВЫПОЛНЕНО')
eq('К-20 рекомендация(−Э1) = эталон', g(OB, 'F63'), jk20['runs'][0]['rec'])
ck('К-20 вердикт не «обосновано»', g(OB, 'B84') != 'Решение статистически обосновано', g(OB, 'B84'), '!=обосновано')

print('К-21: Н/О → G1 не выполнено, W complete cases; разногласие σ>1 → G5; n<3 → вердикт ограничен')
d_no = exp(a=[4] * 10, c_all=4); d_no['A'][0] = 'Н/О'
g = run({1: d_no, 2: exp(a=[4] * 10, c_all=4), 3: exp(a=[4] * 10, c_all=4)})
eq('К-21 Н/О всего=1', g(OB, 'B57'), 1)
eq('К-21 G1 не выполнено', g(OB, 'B78'), 'НЕ ВЫПОЛНЕНО')
eq('К-21 W complete cases m=2', g(OB, 'C51'), 2)
eq('К-21 полнота данных НЕТ', g(OB, 'B60'), 'НЕТ')
eq('К-21 jackknife н/д при неполных', g(OB, 'B73'), 'н/д')
g = run({1: exp(a=[1] + [4] * 9, c_all=4), 2: exp(a=[5] + [4] * 9, c_all=4), 3: exp(a=[4] * 10, c_all=4)})
eq('К-21 разногласие А1 (σ>1)', g('Свод', 'Q21'), 'РАЗНОГЛАСИЕ')
eq('К-21 G5 не выполнено', g(OB, 'B82'), 'НЕ ВЫПОЛНЕНО')
g = run({1: exp(a=[4] * 10, c_all=4), 2: exp(a=[4] * 10, c_all=4)})
has('К-21 n<3 вердикт ограничен', g(OB, 'B84'), 'Недостаточно экспертов')

print('К-22: режим весов (Kish n_eff) + равные веса = невзвешенный + пустая книга')
g = run({1: exp(a=[5] * 10, c_all=4, w=1), 2: exp(a=[3] * 10, c_all=4, w=1), 3: exp(a=[4] * 10, c_all=4, w=2)}, cal={'B9': 'Да'})
ciw = ci_ref([5.0, 3.0, 4.0], [1, 1, 2])
close('К-22 взвешенный I_A=4', g(OB, 'B44'), 4.0, 1e-6)
close('К-22 взвешенный SE', g(OB, 'E44'), ciw['se'], 1e-4)
eq('К-22 df по n_eff=1', g(OB, 'F44'), ciw['df'])
close('К-22 взвешенный ДИ нижн.', g(OB, 'H44'), ciw['lo'], 1e-3)
g = run({1: exp(a=[5] * 10, c_all=4, w=1), 2: exp(a=[3] * 10, c_all=4, w=1), 3: exp(a=[4] * 10, c_all=4, w=1)}, cal={'B9': 'Да'})
cie = ci_ref([5.0, 3.0, 4.0])
close('К-22 равные веса SE = s/√n', g(OB, 'E44'), cie['se'], 1e-6)
eq('К-22 равные веса df = n−1', g(OB, 'F44'), cie['df'])
close('К-22 равные веса ДИ нижн.', g(OB, 'H44'), cie['lo'], 1e-6)
g = run({})
eq('К-22 пустая книга: W н/д', g(OB, 'D51'), 'н/д')
eq('К-22 пустая книга: I_A н/д', g(OB, 'B44'), 'н/д')
has('К-22 пустая книга: вердикт ограничен', g(OB, 'B84'), 'Недостаточно экспертов')
eq('К-22 пустая книга: ворот 1 (G5 вакуумно истинно, вердикт ограничен)', g(OB, 'B83'), 1)

'''
rep('wbs = openpyxl.load_workbook(SRC)\neq(\'19 листов\'', CASES + 'wbs = openpyxl.load_workbook(SRC)\neq(\'19 листов\'')

# ---- статика нового листа ----
rep("ck('свод: помощники скрыты', wbs['Свод'].column_dimensions['W'].hidden is True)",
    """ck('свод: помощники скрыты', wbs['Свод'].column_dimensions['W'].hidden is True)
wso = wbs['Обоснованность']
eq('обоснованность: t df1=12.706', wso['B6'].value, 12.706)
eq('обоснованность: t df9=2.262', wso['B14'].value, 2.262)
eq('обоснованность: χ² df9=16.919', wso['E6'].value, 16.919)
eq('обоснованность: χ² df11=19.675', wso['E7'].value, 19.675)
eq('обоснованность: χ² df21=32.671', wso['E8'].value, 32.671)
eq('обоснованность: k блока 22', wso['B51'].value, 22)
eq('обоснованность: k блока А=10', wso['B52'].value, 10)
eq('обоснованность: k блока Ц=12', wso['B53'].value, 12)
ck('обоснованность: формула W', isinstance(wso['D51'].value, str) and wso['D51'].value.startswith('=IF'))
ck('обоснованность: формула вердикта', isinstance(wso['B84'].value, str) and 'Недостаточно экспертов' in wso['B84'].value)
ck('обоснованность: помощники рангов скрыты (строки)', wso.row_dimensions[100].hidden is True)
ck('обоснованность: помощники скрыты (столбцы M)', wso.column_dimensions['M'].hidden is True)
ck('обоснованность: лист защищён', wso.protection.sheet is True)
eq('обоснованность: Свод B92 ссылка', wbs['Свод']['B92'].value, '=Обоснованность!B84')
eq('обоснованность: Панель B9 ссылка', wbs['Панель']['B9'].value, '=Обоснованность!B84')
ck('обоснованность: порядок после Свода', wbs.sheetnames.index('Обоснованность') == wbs.sheetnames.index('Свод') + 1)""")

assert not re.search(r'[一-鿿]', t), 'CJK!'
open(os.path.join(HERE, 'verify_10_xlsx_22.py'), 'w', encoding='utf-8').write(t)
print('verify_10_xlsx_22.py written:', len(t), 'chars')