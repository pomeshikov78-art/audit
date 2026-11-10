# -*- coding: utf-8 -*-
"""Проверка расчётов шаблона: заполняем тестовые ответы, пересчитываем книгу
библиотекой formulas, сверяем интегральный балл и обработку «Н/П»."""
import shutil
import sys
import openpyxl
import formulas

SRC = "output/1_Чек-лист_аудит_НИОКР_шаблон_расчёта.xlsx"

B1 = ["Да, полностью", "В основном да", "В основном да", "Частично", "Да, полностью",
      "В основном да", "Н/П", "Частично", "В основном да"]                       # S=32, N=8 → P=0.800, B=24.000
B2 = ["В основном да", "Частично", "Частично", "В основном да", "Частично",
      "Частично", "В основном да", "Частично", "Скорее нет"]                    # S=28, N=9 → P=0.6222, B=21.778
B3 = ["Да, полностью", "В основном да", "В основном да", "Частично", "В основном да",
      "В основном да", "Частично", "В основном да", "Н/П"]                      # S=31, N=8 → P=0.775, B=27.125

def fill(path, b1, b2, b3, confirm_first="Нет"):
    wb = openpyxl.load_workbook(path)
    cl = wb["Чек-лист"]
    rows = list(range(7, 16)) + list(range(18, 27)) + list(range(29, 38))
    answers = b1 + b2 + b3
    for r, a in zip(rows, answers):
        cl.cell(row=r, column=4, value=a)
    # подтверждение: первая строка — «Нет» при 5 баллах (ожидаем флаг ⚠), остальные «Да»
    for i, r in enumerate(rows):
        cl.cell(row=r, column=6, value=(confirm_first if i == 0 else "Да"))
    ps = wb["Паспорт НИОКР"]
    ps["C5"] = "Тестовая НИОКР «Проверка расчёта»"
    ps["C6"] = "ТЕСТ-01"
    ps["C17"] = 100.0   # утверждённая
    ps["C19"] = 60.0    # факт
    ps["C20"] = 20.0    # прогноз до завершения
    for r in range(41, 44):
        ps.cell(row=r, column=4, value="Не выявлено")
    ps.cell(row=44, column=4, value="Выявлено")     # одно контрольное условие
    ps.cell(row=45, column=4, value="Требует проверки")
    wb.save(path)

def calc(path):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()
    out = {}
    for k, v in sol.items():
        out[k.upper()] = v
    return out

def get(sol, sheet, addr):
    key = f"'[1_ЧЕК-ЛИСТ_АУДИТ_НИОКР_ШАБЛОН_РАСЧЁТА.XLSX]{sheet.upper()}'!{addr}"
    v = sol.get(key)
    if v is None:
        # поиск по подстроке
        cands = [k for k in sol if k.endswith(f"]{sheet.upper()}'!{addr}")]
        if cands:
            v = sol[cands[0]]
    try:
        val = v.value[0, 0]
    except Exception:
        val = v
    return val

def scenario(name, b1, b2, b3, expect):
    path = f"/tmp/test_{name}.xlsx"
    shutil.copy(SRC, path)
    fill(path, b1, b2, b3)
    sol = calc(path)
    res = {}
    for addr in ["E16", "F16", "G16", "E27", "F27", "G27", "E38", "F38", "G38",
                 "G45", "H45", "G47", "G48", "G49", "G7"]:
        res[addr] = get(sol, "Чек-лист", addr)
    for addr in ["C21", "C22", "C23", "G35", "G36", "G37", "B48", "G30", "F30"]:
        res["PS!" + addr] = get(sol, "Паспорт НИОКР", addr)
    print(f"===== Сценарий: {name} =====")
    ok = True
    for k, exp in expect.items():
        got = res.get(k)
        try:
            match = (abs(float(got) - float(exp)) < 1e-6) if isinstance(exp, (int, float)) else (str(got) == str(exp))
        except Exception:
            match = str(got) == str(exp)
        status = "OK " if match else "FAIL"
        if not match:
            ok = False
        print(f"  {status} {k}: ожидание={exp!r} получено={got!r}")
    for k in ("G7", "PS!B48"):
        print(f"  info {k}: {res.get(k)!r}")
    return ok, res

if __name__ == "__main__":
    # Сценарий 1: базовый (Н/П в блоках I и III) → I = 24 + 21.777... + 27.125 = 72.9027...
    exp1 = {
        "E16": 32, "F16": 8, "G16": 0.8,
        "E27": 28, "F27": 9,
        "E38": 31, "F38": 8, "G38": 0.775,
        "H45": 100, "G48": "Хорошая",
        "PS!C21": 80.0, "PS!C22": -20.0, "PS!C23": -0.2,
        "PS!G36": "Хорошая",
    }
    ok1, r1 = scenario("base", B1, B2, B3, exp1)
    i1 = r1.get("G47")
    print("  I (расчёт) =", i1, "ожидаем ≈ 72.9028")
    ok1 = ok1 and abs(float(i1) - (24 + 35 * 28 / 45 + 27.125)) < 1e-6

    # Сценарий 2: блок III целиком «Н/П» → ренормировка I = 100·(B1+B2)/65
    b3_np = ["Н/П"] * 9
    exp2 = {"E38": 0, "F38": 0, "H45": 65, "G47": 100 * (24 + 35 * 28 / 45) / 65}
    ok2, r2 = scenario("renorm", B1, B2, b3_np, exp2)

    # Сценарий 3: всё «Н/П» → Недостаточно данных
    exp3 = {"H45": 0, "G47": "Н/П", "G48": "Недостаточно данных"}
    ok3, r3 = scenario("allnp", ["Н/П"] * 9, ["Н/П"] * 9, ["Н/П"] * 9, exp3)

    print("\nИТОГ:", "ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ" if (ok1 and ok2 and ok3) else "ЕСТЬ РАСХОЖДЕНИЯ")
    sys.exit(0 if (ok1 and ok2 and ok3) else 1)
