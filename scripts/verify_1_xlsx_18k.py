# -*- coding: utf-8 -*-
"""Этап 18Г: верификация 1_Шаблон_расчёта_аудита_НИОКР_v18_Краткий_15.xlsx независимым пересчётом (formulas).
12 контрольных примеров: индексы, Coverage (порог 80), маршрутизация, матрица, SF-блокировки,
предварительная формулировка заключения (conc.K), замечания по колонке L."""
import shutil, json, io, re
import openpyxl, formulas

SRC = 'output/1_Шаблон_расчёта_аудита_НИОКР_v18_Краткий_15.xlsx'
TMP = '/tmp/s1k_case.xlsx'
q18 = json.load(io.open('scripts/data/q18.json', encoding='utf-8'))
K = q18['levels']['K']['q']
CONCK = q18['conc']['K']
CODES = [x['code'] for x in K]; SEC = {x['code']: x['sec'] for x in K}; W = {x['code']: float(x['w']) for x in K}
_html = io.open('output/Аудит_НИОКР_интерактивная_модель.html', encoding='utf-8').read()
MX4 = json.loads(re.search(r'RECMX4=(\[.*?\]);\n', _html, re.S).group(1))
R0 = 6
PASS = FAIL = 0; FAILED = []

def ck(name, cond, got=None, want=None):
    global PASS, FAIL
    if cond: PASS += 1; print("  ✔", name)
    else: FAIL += 1; FAILED.append(name); print("  ✘ %s | got=%r want=%r" % (name, got, want))

def num(v):
    try: return float(v)
    except Exception: return None

def close(name, v, want, tol=0.06):
    n = num(v); ck(name, n is not None and abs(n - want) <= tol, v, round(want, 2))

def eq(name, v, want):
    ck(name, v == want, v, want)

def has(name, v, sub):
    ck(name, isinstance(v, str) and sub in v, v, "contains " + repr(sub))

def refAns(a):
    ia = {}
    for sec, k in (("А", "IA"), ("Б", "IC"), ("В", "IR")):
        code = [c for c in CODES if SEC[c] == sec and a.get(c, "") not in ("", "НД", "НП")]
        numr = sum(W[c] * float(a[c]) for c in code); den = sum(W[c] for c in code)
        ia[k] = 100 * numr / (5 * den) if den else None
    nums = [c for c in CODES if a.get(c, "") not in ("", "НД", "НП")]
    nnp = sum(1 for c in CODES if a.get(c, "") == "НП")
    ia['cov'] = 100 * len(nums) / (15 - nnp) if (15 - nnp) else None
    if ia['IA'] and ia['IC'] and ia['IR']:
        ia['I'] = (ia['IA'] * ia['IC'] * ia['IR']) ** (1 / 3)
    else:
        ia['I'] = None
    return ia

def lvl(v):
    if v is None: return "—"
    return "Высокий" if v >= 85 else ("Приемлемый" if v >= 70 else ("Проблемный" if v >= 55 else "Критический"))

def mkcase(ans=None, sf=None, mat=None, extra=None):
    shutil.copy(SRC, TMP)
    wb = openpyxl.load_workbook(TMP)
    w1 = wb['Опросник_15']
    for i, c in enumerate(CODES):
        v = (ans or {}).get(c, "")
        w1.cell(R0 + i, 8, (int(v) if v.isdigit() else v) if v else None)
        if (extra or {}).get(c):
            j, txt = extra[c]
            w1.cell(R0 + i, 10 if j == "j" else 11, txt)
    w2 = wb['Расчёт']
    if sf:
        for addr, v in sf.items(): w2[addr] = v
    if mat:
        for addr, v in mat.items(): w2[addr] = v
    wb.save(TMP); wb.close()

def solve():
    xl = formulas.ExcelModel().loads(TMP).finish()
    return xl.calculate()

def val(sol, sheet, cell):
    pref = "'[S1K_CASE.XLSX]%s'!%s" % (sheet, cell)
    for k, v in sol.items():
        if k.upper().endswith(pref.upper()):
            try: return v.value[0, 0]
            except Exception: return v.value
    return None

print("=== СТРУКТУРА")
wb = openpyxl.load_workbook(SRC)
ck("лист Опросник_15", 'Опросник_15' in wb.sheetnames)
ck("лист Расчёт", 'Расчёт' in wb.sheetnames)
ck("лист Справочники", 'Справочники' in wb.sheetnames)
w1 = wb['Опросник_15']
ck("15 строк кодов", w1.cell(R0, 1).value == 'А1' and w1.cell(R0 + 14, 1).value == 'В15')
ck("веса проставлены", all(isinstance(w1.cell(R0 + i, 7).value, (int, float)) for i in range(15)))
ck("якоря заполнены", all(w1.cell(R0 + i, 4).value and w1.cell(R0 + i, 6).value for i in range(15)))
wb.close()

print("=== К1: все «5» → 100/100/100, R1, формулировка «обоснована»")
mkcase({c: "5" for c in CODES}, extra={c: ("e", "доказательство") for c in CODES})
s = solve(); ref = refAns({c: "5" for c in CODES})
close("I_А=100", val(s, 'Расчёт', 'B4'), ref['IA'])
close("I_Ц=100", val(s, 'Расчёт', 'B5'), ref['IC'])
close("I_Р=100", val(s, 'Расчёт', 'B6'), ref['IR'])
close("Комплексный I=100", val(s, 'Расчёт', 'B7'), 100)
close("Coverage=100", val(s, 'Расчёт', 'B9'), 100)
eq("уровень I_А Высокий", val(s, 'Расчёт', 'B14'), "Высокий")
has("маршрутизация Высокий", val(s, 'Расчёт', 'B17'), "Высокий — продолжить выполнение")
has("матрица R1", val(s, 'Расчёт', 'B29'), "R1")
eq("формулировка conc.K[0]", val(s, 'Расчёт', 'B33'), CONCK[0][1])
eq("примечаний 0", num(val(s, 'Расчёт', 'B11')), 0)

print("=== К2: все «0» → 0, Критический, R7, внеплановый аудит")
mkcase({c: "0" for c in CODES}, extra={c: ("j", "обоснование") for c in CODES})
s = solve()
close("I_А=0", val(s, 'Расчёт', 'B4'), 0)
close("Комплексный I=0", val(s, 'Расчёт', 'B7'), 0)
eq("уровень Критический", val(s, 'Расчёт', 'B14'), "Критический")
has("матрица R7", val(s, 'Расчёт', 'B29'), "R7")
eq("формулировка conc.K[3]", val(s, 'Расчёт', 'B33'), CONCK[3][1])
has("маршрутизация внеплановое", val(s, 'Расчёт', 'B17'), "внеплановое углублённое обследование")

print("=== К3: все «4» → 80.0, Приемлемый; min 80 → «обоснована»")
mkcase({c: "4" for c in CODES}, extra={c: ("e", "доказательство") for c in CODES})
s = solve(); ref = refAns({c: "4" for c in CODES})
close("I_А≈80", val(s, 'Расчёт', 'B4'), ref['IA'])
eq("уровень Приемлемый", val(s, 'Расчёт', 'B14'), "Приемлемый")
eq("формулировка conc.K[0] (min 80 ≥ 80)", val(s, 'Расчёт', 'B33'), CONCK[0][1])

print("=== К4: один «НД» (А1) → Coverage 93,3; I_А по 4 кодам")
a = {c: "4" for c in CODES}; a["А1"] = "НД"
mkcase(a, extra={c: ("e", "доказательство") for c in CODES if c != "А1"})
s = solve(); ref = refAns(a)
close("I_А без А1 = 80", val(s, 'Расчёт', 'B4'), ref['IA'])
close("Coverage 93.3", val(s, 'Расчёт', 'B9'), ref['cov'])

print("=== К5: «НП» с обоснованием (А2) → исключён из индекса и Coverage")
a = {c: "4" for c in CODES}; a["А2"] = "НП"
mkcase(a, extra=dict({c: ("e", "доказательство") for c in CODES}, **{"А2": ("j", "этап не предусматривается")}))
s = solve(); ref = refAns(a)
close("I_А без А2", val(s, 'Расчёт', 'B4'), ref['IA'])
close("Coverage 14/14", val(s, 'Расчёт', 'B9'), ref['cov'])
eq("без примечаний", num(val(s, 'Расчёт', 'B11')), 0)

print("=== К6: «НП» без обоснования → примечание")
a = {c: "4" for c in CODES}; a["А3"] = "НП"
mkcase(a, extra={c: ("e", "доказательство") for c in CODES if c != "А3"})
s = solve()
ck("≥1 примечание", num(val(s, 'Расчёт', 'B11')) >= 1, num(val(s, 'Расчёт', 'B11')), True)
has("текст НП", val(s, 'Опросник_15', 'L%d' % (R0 + CODES.index("А3"))), "НП: нет письменного обоснования")

print("=== К7: балл 5 без доказательства → примечание «4–5 без доказательства»")
a = {c: "3" for c in CODES}; a["В15"] = "5"
mkcase(a, extra={c: ("j", "обоснование") for c in CODES if c != "В15"})
s = solve()
has("текст 4–5", val(s, 'Опросник_15', 'L%d' % (R0 + CODES.index("В15"))), "4–5 без доказательства")

print("=== К8: SF-2 подтверждён → блокировка и эскалация")
mkcase({c: "5" for c in CODES}, extra={c: ("e", "доказательство") for c in CODES}, sf={"B22": 1})
s = solve()
has("блокировка SF-2", val(s, 'Расчёт', 'B31'), "SF-2")
has("эскалация в формулировке", val(s, 'Расчёт', 'B33'), "Эскалация")

print("=== К9: все «НП» → н/д everywhere")
mkcase({c: "НП" for c in CODES}, extra={c: ("j", "не применимо") for c in CODES})
s = solve()
eq("I_А н/д", val(s, 'Расчёт', 'B4'), "н/д")
eq("Coverage н/д", val(s, 'Расчёт', 'B9'), "н/д")
eq("матрица н/д", val(s, 'Расчёт', 'B29'), "н/д")
has("недостаточно данных", val(s, 'Расчёт', 'B33'), "Недостаточно данных")

print("=== К10: смешанный набор против независимого пересчёта")
a = {}
mix = [5, 4, 3, 2, 1]
for i, c in enumerate(CODES):
    a[c] = str(mix[i % 5])
mkcase(a, extra={c: (("j", "обоснование") if float(a[c]) <= 2 else ("e", "доказательство")) for c in CODES})
s = solve(); ref = refAns(a)
close("I_А = ref", val(s, 'Расчёт', 'B4'), ref['IA'])
close("I_Ц = ref", val(s, 'Расчёт', 'B5'), ref['IC'])
close("I_Р = ref", val(s, 'Расчёт', 'B6'), ref['IR'])
close("Комплексный = ref", val(s, 'Расчёт', 'B7'), ref['I'])
close("Coverage = ref", val(s, 'Расчёт', 'B9'), ref['cov'])
eq("уровень I_А = ref-lvl", val(s, 'Расчёт', 'B14'), lvl(ref['IA']))

print("=== К11: M≥5 без испытаний → замечание")
mkcase({c: "4" for c in CODES}, extra={c: ("e", "доказательство") for c in CODES}, mat={"B25": "M5", "B26": "нет"})
s = solve()
has("замечание M5", val(s, 'Расчёт', 'B27'), "M≥5 без подтверждения")

print("=== К12: матрица — Проблемный×Высокий")
a = {c: "3" for c in CODES if SEC[c] == "А"}
a.update({c: "5" for c in CODES if SEC[c] != "А"})
mkcase(a, extra={c: ("e", "доказательство") for c in CODES})
s = solve()
eq("уровень I_А Проблемный", val(s, 'Расчёт', 'B14'), "Проблемный")
eq("матрица = RECMX4[2][0]", val(s, 'Расчёт', 'B29'), MX4[2][0][0])

print()
print("ИТОГ VERIFY-1-XLSX-18K: PASS=%d FAIL=%d" % (PASS, FAIL))
for f in FAILED: print("  ✗", f)
