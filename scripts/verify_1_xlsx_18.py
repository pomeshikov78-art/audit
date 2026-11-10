# -*- coding: utf-8 -*-
"""Этап 18В: верификация 1_Шаблон_расчёта_аудита_НИОКР_v18_Средний_30.xlsx независимым пересчётом (formulas)."""
import shutil, json, io
import openpyxl, formulas

SRC = 'output/1_Шаблон_расчёта_аудита_НИОКР_v18_Средний_30.xlsx'
TMP = '/tmp/s1_case.xlsx'
q18 = json.load(io.open('scripts/data/q18.json', encoding='utf-8'))
S = q18['levels']['S']['q']
import re as _re
_html = io.open('output/Аудит_НИОКР_интерактивная_модель.html', encoding='utf-8').read()
MX4 = json.loads(_re.search(r'RECMX4=(\[.*?\]);\n', _html, _re.S).group(1))
CODES = [x['code'] for x in S]; SEC = {x['code']: x['sec'] for x in S}; W = {x['code']: float(x['w']) for x in S}
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
        code = [c for c in CODES if SEC[c] == sec and a.get(c, "") != "" and a.get(c, "") not in ("НД", "НП")]
        numr = sum(W[c] * float(a[c]) for c in code); den = sum(W[c] for c in code)
        ia[k] = 100 * numr / (5 * den) if den else None
    nums = [c for c in CODES if a.get(c, "") != "" and a.get(c, "") not in ("НД", "НП")]
    nnp = sum(1 for c in CODES if a.get(c, "") == "НП")
    cov = 100 * len(nums) / (30 - nnp) if (30 - nnp) else None
    if ia['IA'] and ia['IC'] and ia['IR']:
        ia['I'] = (ia['IA'] * ia['IC'] * ia['IR']) ** (1 / 3)
    else:
        ia['I'] = None
    ia['cov'] = cov
    return ia

def lvl(v):
    if v is None: return "—"
    return "Высокий" if v >= 85 else ("Приемлемый" if v >= 70 else ("Проблемный" if v >= 55 else "Критический"))

def mkcase(ans=None, sf=None, mat=None, extra=None):
    shutil.copy(SRC, TMP)
    wb = openpyxl.load_workbook(TMP)
    w1 = wb['Опросник_30']
    for i, c in enumerate(CODES):
        v = (ans or {}).get(c, "")
        w1.cell(R0 + i, 7, (int(v) if v.isdigit() else v) if v else None)
        if (extra or {}).get(c):
            j, txt = extra[c]
            w1.cell(R0 + i, 9 if j == "j" else 10, txt)
    w2 = wb['Расчёт']
    if sf:
        for addr, v in sf.items(): w2[addr] = v
    if mat:
        for addr, v in mat.items(): w2[addr] = v
    wb.save(TMP); wb.close()

def solve():
    xl = formulas.ExcelModel().loads(TMP).finish()
    sol = xl.calculate()
    return sol

def val(sol, sheet, cell):
    pref = "'[S1_CASE.XLSX]%s'!%s" % (sheet, cell)
    for k, v in sol.items():
        if k.upper().endswith(pref.upper()):
            try: return v.value[0, 0]
            except Exception: return v.value
    return None

def allnum(v):
    mkcase({c: v for c in CODES}, extra={c: ("j", "есть") for c in CODES})
    # обоснования не нужны при 4-5; при 0-2 — нужны, ставим и их
    if num(v) is not None and float(v) <= 2:
        mkcase({c: v for c in CODES}, extra={c: ("e", "обоснование") for c in CODES})
    return solve()

print("=== СТРУКТУРА")
wb = openpyxl.load_workbook(SRC)
ck("лист Опросник_30", 'Опросник_30' in wb.sheetnames)
ck("лист Расчёт", 'Расчёт' in wb.sheetnames)
ck("лист Справочники", 'Справочники' in wb.sheetnames)
w1 = wb['Опросник_30']
ck("30 строк кодов", [w1.cell(R0 + i, 1).value for i in range(30)][0] == 'А1' and w1.cell(R0 + 29, 1).value == 'В10')
ck("веса проставлены", all(isinstance(w1.cell(R0 + i, 6).value, (int, float)) for i in range(30)))
wb.close()

print("=== К1: все «5» → 100/100/100, R1, кат. I")
mkcase({c: "5" for c in CODES}, extra={c: ("e", "доказательство") for c in CODES})
s = solve()
ref = refAns({c: "5" for c in CODES})
close("I_А=100", val(s, 'Расчёт', 'B4'), ref['IA'])
close("I_Ц=100", val(s, 'Расчёт', 'B5'), ref['IC'])
close("I_Р=100", val(s, 'Расчёт', 'B6'), ref['IR'])
close("Комплексный I=100", val(s, 'Расчёт', 'B7'), 100)
close("Coverage=100", val(s, 'Расчёт', 'B9'), 100)
has("матрица R1", val(s, 'Расчёт', 'B28'), "R1")
has("категория I", val(s, 'Расчёт', 'B32'), "I. Положительное")
eq("примечаний 0", num(val(s, 'Расчёт', 'B11')), 0)

print("=== К2: все «0» → 0, R7, кат. IV")
mkcase({c: "0" for c in CODES}, extra={c: ("e", "обоснование") for c in CODES})
s = solve()
close("I_А=0", val(s, 'Расчёт', 'B4'), 0)
close("Комплексный I=0", val(s, 'Расчёт', 'B7'), 0)
eq("уровень Критический", val(s, 'Расчёт', 'B14'), "Критический")
has("матрица R7", val(s, 'Расчёт', 'B28'), "R7")
has("категория IV", val(s, 'Расчёт', 'B32'), "IV")

print("=== К3: все «4» → 80.0, «Приемлемый», R3-ячейка, кат. II")
mkcase({c: "4" for c in CODES}, extra={c: ("e", "доказательство") for c in CODES})
s = solve(); ref = refAns({c: "4" for c in CODES})
close("I_А≈80", val(s, 'Расчёт', 'B4'), ref['IA'])
eq("уровень Приемлемый", val(s, 'Расчёт', 'B14'), "Приемлемый")
has("матрица [1][1] R3", val(s, 'Расчёт', 'B28'), "R3")
has("категория II", val(s, 'Расчёт', 'B32'), "II.")

print("=== К4: один «НД» (А1) → Coverage 96,7; I_А по 9 кодам")
a = {c: "4" for c in CODES}; a["А1"] = "НД"
mkcase(a, extra={c: ("e", "обоснование") for c in CODES if c != "А1"})
s = solve(); ref = refAns(a)
close("I_А без А1 = 80", val(s, 'Расчёт', 'B4'), ref['IA'])
close("Coverage 96.7", val(s, 'Расчёт', 'B9'), ref['cov'])

print("=== К5: «НП» с обоснованием (А2) → исключён из индекса и Coverage")
a = {c: "4" for c in CODES}; a["А2"] = "НП"
mkcase(a, extra=dict({c: ("e", "обоснование") for c in CODES}, **{"А2": ("j", "этап не предусматривается")}))
s = solve(); ref = refAns(a)
close("I_А без А2", val(s, 'Расчёт', 'B4'), ref['IA'])
close("Coverage 28/29", val(s, 'Расчёт', 'B9'), ref['cov'])
eq("без примечаний", num(val(s, 'Расчёт', 'B11')), 0)

print("=== К6: «НП» без обоснования → примечание")
a = {c: "4" for c in CODES}; a["А3"] = "НП"
mkcase(a, extra={c: ("e", "обоснование") for c in CODES if c != "А3"})
s = solve()
eq("≥1 примечание", num(val(s, 'Расчёт', 'B11')) >= 1, True)
has("текст НП", val(s, 'Опросник_30', 'K%d' % (R0 + CODES.index("А3"))), "НП: нет письменного обоснования")

print("=== К7: балл 5 без доказательства → примечание «4–5 без доказательства»")
a = {c: "3" for c in CODES}; a["В5"] = "5"
mkcase(a, extra={c: ("e", "обоснование") for c in CODES if c != "В5"})
s = solve()
has("текст 4–5", val(s, 'Опросник_30', 'K%d' % (R0 + CODES.index("В5"))), "4–5 без доказательства")

print("=== К8: SF-2 подтверждён → блокировка и кат. IV")
mkcase({c: "5" for c in CODES}, extra={c: ("e", "доказательство") for c in CODES}, sf={"B21": 1})
s = solve()
has("блокировка", val(s, 'Расчёт', 'B30'), "SF-2")
has("категория IV/эскалация", val(s, 'Расчёт', 'B32'), "IV.")

print("=== К9: все «НП» → н/д everywhere")
mkcase({c: "НП" for c in CODES}, extra={c: ("j", "не применимо") for c in CODES})
s = solve()
eq("I_А н/д", val(s, 'Расчёт', 'B4'), "н/д")
eq("Coverage н/д", val(s, 'Расчёт', 'B9'), "н/д")
eq("матрица н/д", val(s, 'Расчёт', 'B28'), "н/д")
has("недостаточно данных", val(s, 'Расчёт', 'B32'), "Недостаточно данных")

print("=== К10: смешанный набор против независимого пересчёта")
a = {}
mix = [5, 4, 4, 3, 2, 4, 5, 3, 1, 4]
for i, c in enumerate(CODES):
    a[c] = str(mix[i % 10])
mkcase(a, extra={c: ("e", "обоснование") for c in CODES if float(a[c]) <= 2})
s = solve(); ref = refAns(a)
close("I_А = ref", val(s, 'Расчёт', 'B4'), ref['IA'])
close("I_Ц = ref", val(s, 'Расчёт', 'B5'), ref['IC'])
close("I_Р = ref", val(s, 'Расчёт', 'B6'), ref['IR'])
close("Комплексный = ref", val(s, 'Расчёт', 'B7'), ref['I'])
close("Coverage = ref", val(s, 'Расчёт', 'B9'), ref['cov'])
eq("уровень I_А = ref-lvl", val(s, 'Расчёт', 'B14'), lvl(ref['IA']))

print("=== К11: M≥5 без испытаний → замечание")
mkcase({c: "4" for c in CODES}, mat={"B24": "M5", "B25": "нет"})
s = solve()
has("замечание M5", val(s, 'Расчёт', 'B26'), "M≥5 без подтверждения")

print("=== К12: матрица — Проблемный×Высокий")
a = {c: "3" for c in CODES if SEC[c] == "А"}
a.update({c: "5" for c in CODES if SEC[c] != "А"})
mkcase(a, extra={c: ("e", "обоснование") for c in CODES})
s = solve()
eq("уровень I_А Проблемный", val(s, 'Расчёт', 'B14'), "Проблемный")
eq("матрица = RECMX4[2][0] (Проблемный A × Высокий Ц)", val(s, 'Расчёт', 'B28'), MX4[2][0][0])

print()
print("ИТОГ VERIFY-1-XLSX-18: PASS=%d FAIL=%d" % (PASS, FAIL))
for f in FAILED: print("  ✗", f)
