# -*- coding: utf-8 -*-
"""Генератор Excel-шаблона чек-листа аудита НИОКР.
Листы: Инструкция, Паспорт НИОКР, Чек-лист, Портфель, Справочник.
Формулы учитывают «Н/П» (исключение из числителя и знаменателя + ренормировка весов).
"""
import math
import sys

sys.path.insert(0, "scripts")
from data_checklist import (SCALE, BLOCKS, CONTROL_CONDITIONS, DECISIONS,
                            SCALE_VERDICTS, all_questions)

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule, DataBarRule

# ---------- палитра ----------
NAVY = "1F3864"       # тёмно-синий — заголовки
NAVY2 = "2E5395"      # синий — подзаголовки секций
ACCENT = "2E74B5"
INPUT = "FFF2CC"      # жёлтый — ячейки для ввода
CALC = "F2F2F2"       # серый — вычисляемые ячейки
GREEN_D = "C6EFCE"; GREEN_D_T = "006100"
GREEN_L = "E2EFDA"; GREEN_L_T = "375623"
YELLOW = "FFEB9C"; YELLOW_T = "9C6500"
ORANGE = "FCE4D6"; ORANGE_T = "974706"
RED = "FFC7CE"; RED_T = "9C0006"
GRAY = "D9D9D9"; GRAY_T = "595959"
BLUE_L = "DDEBF7"

F_TITLE = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
F_SUB = Font(name="Calibri", size=11, italic=True, color="595959")
F_SEC = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
F_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_B = Font(name="Calibri", size=11, bold=True)
F_N = Font(name="Calibri", size=10)
F_NB = Font(name="Calibri", size=10, bold=True)
F_SM = Font(name="Calibri", size=9, color="595959")
F_BIG = Font(name="Calibri", size=22, bold=True, color=NAVY)
F_LBL = Font(name="Calibri", size=11, bold=True, color=NAVY)

FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor=NAVY2)
FILL_HDR = PatternFill("solid", fgColor=ACCENT)
FILL_INPUT = PatternFill("solid", fgColor=INPUT)
FILL_CALC = PatternFill("solid", fgColor=CALC)
FILL_BLUE = PatternFill("solid", fgColor=BLUE_L)

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
WRAP_CTR = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT_TOP = Alignment(wrap_text=True, vertical="top", horizontal="left")
CTR = Alignment(vertical="center", horizontal="center", wrap_text=True)


def est_h(text, width_chars, base=15.5, pad=5, minh=18):
    """Оценка высоты строки для wrap-текста."""
    if not text:
        return minh
    lines = 0
    for part in str(text).split("\n"):
        lines += max(1, math.ceil(len(part) / width_chars))
    return max(minh, base * lines + pad)


def title_row(ws, row, text, ncols, font=F_TITLE, fill=FILL_TITLE, h=30, align=None):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = font
    c.fill = fill
    c.alignment = align or Alignment(vertical="center", horizontal="left", indent=1)
    ws.row_dimensions[row].height = h


def section_row(ws, row, first_col, last_col, text, h=22):
    ws.merge_cells(start_row=row, start_column=first_col, end_row=row, end_column=last_col)
    c = ws.cell(row=row, column=first_col, value=text)
    c.font = F_SEC
    c.fill = FILL_SEC
    c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    ws.row_dimensions[row].height = h


def hdr_cell(ws, row, col, text, fill=FILL_HDR, font=F_HDR):
    c = ws.cell(row=row, column=col, value=text)
    c.font = font
    c.fill = fill
    c.alignment = CTR
    c.border = BORDER
    return c


wb = Workbook()

# =====================================================================
# ЛИСТ «Инструкция»
# =====================================================================
ws = wb.active
ws.title = "Инструкция"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2.5
ws.column_dimensions["B"].width = 118
for col in "CDEFGH":
    ws.column_dimensions[col].width = 2.5

r = 1
title_row(ws, r, "ШАБЛОН ЧЕК-ЛИСТА АУДИТА НИОКР — ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ", 8, h=32); r += 1
c = ws.cell(row=r, column=2, value="Аудит актуальности, эффективности и целесообразности НИОКР. Версия 2.0 (финальная, математически выверенная). Файл не содержит макросов.")
c.font = F_SUB; r += 2

INSTR = [
    ("S", "1. Назначение и состав файла"),
    ("T", "Файл реализует методику балльной оценки отдельной НИОКР (этапа, завершённой разработки) по 27 контрольным вопросам трёх блоков с автоматическим расчётом интегрального показателя I (0–100 баллов), полной обработкой ответов «Н/П» и условным форматированием."),
    ("B", "«Паспорт НИОКР» — сведения об объекте аудита, сводные результаты (считаются автоматически), контрольные условия, итоговое решение комиссии."),
    ("B", "«Чек-лист» — 27 вопросов (3 блока по 9): заполняются ответы, подтверждения и комментарии; баллы, итоги и интегральный балл рассчитываются автоматически."),
    ("B", "«Портфель» — сводная матрица по нескольким НИОКР подразделения с автоматической категорией по интегральному баллу."),
    ("B", "«Справочник» — шкала ответов, веса блоков, итоговая шкала и списки для выпадающих меню. ИЗМЕНЯТЬ СТРУКТУРУ ЛИСТА НЕЛЬЗЯ — на него ссылаются формулы."),
    ("S", "2. Порядок заполнения"),
    ("N", "Заполните жёлтые ячейки листа «Паспорт НИОКР» (разделы 1–3)."),
    ("N", "На листе «Чек-лист» по каждому вопросу выберите ответ в столбце D (выпадающий список): «Да, полностью» / «В основном да» / «Частично» / «Скорее нет» / «Нет» / «Н/П»."),
    ("N", "В столбце F укажите наличие документального подтверждения («Да» / «Нет»). В столбце H — наименование подтверждающих документов, а для ответа «Н/П» — обязательное обоснование."),
    ("N", "Столбцы E (балл) и G (контроль правила) рассчитываются автоматически. Баллы 4 и 5 без документального подтверждения помечаются красным флажком «⚠»."),
    ("N", "Итоги блоков, интегральный балл I, оценка состояния и рекомендуемое решение рассчитываются автоматически в нижней части «Чек-листа» и дублируются в «Паспорте НИОКР»."),
    ("N", "Проверьте 7 контрольных условий в разделе 5 «Паспорта НИОКР» — они влияют на итоговое решение независимо от набранных баллов."),
    ("N", "Зафиксируйте итоговое решение комиссии (раздел 6 «Паспорта НИОКР») и перенесите результат в матрицу «Портфель»."),
    ("S", "3. Шкала ответов (значение «2 балла» намеренно не используется — чтобы исключить «усреднённые» оценки)"),
    ("B", "«Да, полностью» = 5; «В основном да» = 4; «Частично» = 3; «Скорее нет» = 1; «Нет» = 0; «Н/П» — исключается из расчёта."),
    ("S", "4. Правило работы с ответом «Н/П»"),
    ("T", "Ответ «Н/П» исключает вопрос ОДНОВРЕМЕННО из числителя и из знаменателя: процент выполнения блока P = S / (5·N), где S — сумма баллов только применимых вопросов, N — число применимых вопросов блока. Обоснование «Н/П» в столбце H обязательно."),
    ("T", "Если все вопросы блока получили «Н/П», вес блока исключается из расчёта, а интегральный балл пересчитывается с ренормировкой: I = 100 · ΣBⱼ(прим.) / ΣWⱼ(прим.). Если «Н/П» даны все 27 вопросов — выводится «Недостаточно данных»."),
    ("S", "5. Правило документального подтверждения (автоматически контролируется)"),
    ("T", "Отсутствие документального подтверждения не позволяет присваивать 4 или 5 баллов только на основании устных пояснений исполнителя. При выборе высокого балла и ответа «Нет» в столбце F формула в столбце G выводит красное предупреждение «⚠ 4–5 баллов без подтверждения документом». Комиссия обязана снизить балл или приложить документ."),
    ("S", "6. Методика расчёта (реализована формулами)"),
    ("B", "Sⱼ — сумма баллов применимых вопросов блока j;  Nⱼ — число применимых вопросов блока j (ответы «Н/П» не учитываются)."),
    ("B", "Pⱼ = Sⱼ / (5 · Nⱼ) — процент выполнения блока (0…1)."),
    ("B", "Bⱼ = Wⱼ · Pⱼ — взвешенный балл блока; веса: W(I) = 30, W(II) = 35, W(III) = 35, ΣW = 100."),
    ("B", "I = 100 · Σ Bⱼ(применимых) / Σ Wⱼ(применимых) — интегральный показатель, всегда в диапазоне 0–100."),
    ("B", "Итоговая шкала: 85–100 «Высокая»; 70–84 «Хорошая»; 50–69 «Удовлетворительная»; 30–49 «Неудовлетворительная»; 0–29 «Критическая»."),
    ("S", "7. Цветовая легенда"),
    ("B", "Жёлтый — ячейки для заполнения аудитором.  Серый/голубой — вычисляемые ячейки (не редактировать)."),
    ("B", "Зелёный (тёмный/светлый) — 5/4 балла; жёлтый — 3; оранжевый — 1; красный — 0 и предупреждения ⚠; серый — «Н/П»."),
    ("S", "8. Аудит нескольких НИОКР"),
    ("T", "Рекомендуемая практика: один файл — одна НИОКР (сохраните копию шаблона под именем «Аудит_НИОКР_<шифр>.xlsx»). Для нескольких разработок в одном файле: скопируйте листы «Чек-лист» и «Паспорт НИОКР» (ПКМ по ярлычку → «Переместить или скопировать» → «Создать копию»), переименуйте копии (например «ЧЛ_НИОКР-01», «П_НИОКР-01») и замените в формулах скопированного паспорта ссылки 'Чек-лист'! на имя копии. Итоги по каждой разработке занесите в матрицу листа «Портфель»."),
    ("S", "9. Предупреждения"),
    ("B", "Не удаляйте и не вставляйте строки внутри таблицы вопросов и расчётного блока — формулы ссылаются на фиксированные диапазоны. Не переименовывайте лист «Справочник»."),
    ("B", "Интегральный показатель — инструмент поддержки решения, а не самостоятельное основание для прекращения НИОКР или списания активов. Итоговое решение принимает комиссия с учётом контрольных условий (раздел 5 «Паспорта НИОКР»)."),
]
for kind, text in INSTR:
    cell = ws.cell(row=r, column=2, value=text)
    if kind == "S":
        cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        cell.fill = FILL_SEC
        cell.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r].height = 21
    elif kind == "N":
        cell.font = F_N
        cell.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        ws.row_dimensions[r].height = est_h(text, 112, minh=15)
        # маркер нумерации
        mc = ws.cell(row=r, column=1, value="•")
        mc.font = Font(name="Calibri", size=10, bold=True, color=ACCENT)
        mc.alignment = Alignment(horizontal="center", vertical="top")
    else:  # T / B
        cell.font = F_N if kind == "T" else F_SM
        cell.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        ws.row_dimensions[r].height = est_h(text, 112, minh=15)
    r += 1

# =====================================================================
# ЛИСТ «Справочник»  (создаём раньше чек-листа — на него ссылки)
# =====================================================================
ref = wb.create_sheet("Справочник")
ref.sheet_view.showGridLines = False
widths = {"A": 22, "B": 8, "C": 52, "D": 3, "E": 26, "F": 14, "G": 3, "H": 40}
for k, v in widths.items():
    ref.column_dimensions[k].width = v

title_row(ref, 1, "СПРАВОЧНИК — используется формулами и списками. Не изменяйте структуру листа!", 8, h=26)
hdr_cell(ref, 3, 1, "Вариант ответа"); hdr_cell(ref, 3, 2, "Балл"); hdr_cell(ref, 3, 2 + 1, "Интерпретация")
for i, (ans, ball, interp) in enumerate(SCALE):
    rr = 4 + i
    a = ref.cell(row=rr, column=1, value=ans); a.font = F_NB; a.border = BORDER
    b = ref.cell(row=rr, column=2, value=ball); b.font = F_N; b.border = BORDER; b.alignment = CTR
    c = ref.cell(row=rr, column=3, value=interp); c.font = F_N; c.border = BORDER; c.alignment = LEFT_TOP

ref.cell(row=11, column=5, value="Блок").font = F_HDR
ref.cell(row=11, column=5).fill = FILL_HDR; ref.cell(row=11, column=5).border = BORDER
ref.cell(row=11, column=6, value="Вес Wⱼ, баллы").font = F_HDR
ref.cell(row=11, column=6).fill = FILL_HDR; ref.cell(row=11, column=6).border = BORDER
for i, b in enumerate(BLOCKS):
    rr = 12 + i
    e = ref.cell(row=rr, column=5, value=f'{b["num"]}. {b["name"]}')
    e.font = F_N; e.border = BORDER
    f = ref.cell(row=rr, column=6, value=b["weight"])
    f.font = F_NB; f.border = BORDER; f.alignment = CTR
e = ref.cell(row=15, column=5, value="ИТОГО"); e.font = F_NB; e.border = BORDER
f = ref.cell(row=15, column=6, value="=SUM(F12:F14)"); f.font = F_NB; f.border = BORDER; f.alignment = CTR

hdr_cell(ref, 17, 1, "Интегральный балл I"); hdr_cell(ref, 17, 2, "Оценка"); hdr_cell(ref, 17, 3, "Рекомендуемое решение")
ranges_txt = ["85–100", "70–84", "50–69", "30–49", "0–29"]
for i, (thr, name, dec) in enumerate(SCALE_VERDICTS):
    rr = 18 + i
    a = ref.cell(row=rr, column=1, value=ranges_txt[i]); a.font = F_NB; a.border = BORDER; a.alignment = CTR
    b = ref.cell(row=rr, column=2, value=name); b.font = F_N; b.border = BORDER; b.alignment = CTR
    c = ref.cell(row=rr, column=3, value=dec); c.font = F_N; c.border = BORDER; c.alignment = LEFT_TOP
# покраска строк шкалы
scale_fills = [(GREEN_D, GREEN_D_T), (GREEN_L, GREEN_L_T), (YELLOW, YELLOW_T), (ORANGE, ORANGE_T), (RED, RED_T)]
for i, (fl, ft) in enumerate(scale_fills):
    rr = 18 + i
    for col in (1, 2, 3):
        ref.cell(row=rr, column=col).fill = PatternFill("solid", fgColor=fl)
        ref.cell(row=rr, column=col).font = Font(name="Calibri", size=10, bold=(col <= 2), color=ft)

# списки для валидации (столбец H)
def put_list(start_row, header, items):
    h = ref.cell(row=start_row, column=8, value=header)
    h.font = F_NB; h.fill = FILL_BLUE; h.border = BORDER
    for i, it in enumerate(items):
        c = ref.cell(row=start_row + 1 + i, column=8, value=it)
        c.font = F_N; c.border = BORDER
    return f"$H${start_row+1}:$H${start_row+len(items)}"

L_ANS = put_list(3, "Список: ответы", [s[0] for s in SCALE])
L_CONF = put_list(11, "Список: подтверждение", ["Да", "Нет"])
L_STATUS = put_list(15, "Список: статус НИОКР", ["Выполняется", "Завершена", "Приостановлена", "Прекращена"])
L_KU = put_list(21, "Список: контрольное условие", ["Выявлено", "Не выявлено", "Требует проверки"])
L_DEC = put_list(26, "Список: решение комиссии", DECISIONS)
L_PORT = put_list(36, "Список: решение по портфелю", ["Внедрить", "Продолжить", "Продолжить с корректировкой", "Модернизировать", "Объединить с другой НИОКР", "Приостановить", "Прекратить", "Списать"])
ref.cell(row=46, column=8, value="Диапазоны, на которые ссылаются формулы: шкала A4:B8 (VLOOKUP); веса F12:F14.").font = F_SM

# =====================================================================
# ЛИСТ «Чек-лист»
# =====================================================================
cl = wb.create_sheet("Чек-лист")
cl.sheet_view.showGridLines = False
cl_widths = {"A": 5, "B": 46, "C": 54, "D": 18, "E": 8, "F": 13, "G": 27, "H": 42}
for k, v in cl_widths.items():
    cl.column_dimensions[k].width = v

title_row(cl, 1, "ЧЕК-ЛИСТ: АУДИТ АКТУАЛЬНОСТИ, ЭФФЕКТИВНОСТИ И ЦЕЛЕСООБРАЗНОСТИ НИОКР", 8, h=30)
c = cl.cell(row=2, column=1, value='="НИОКР: "&IF(\'Паспорт НИОКР\'!C5="","— укажите наименование в паспорте —",\'Паспорт НИОКР\'!C5)&"    |    Шифр: "&IF(\'Паспорт НИОКР\'!C6="","—",\'Паспорт НИОКР\'!C6)')
cl.merge_cells("A2:H2"); c.font = F_LBL; c.alignment = Alignment(vertical="center", indent=1)
cl.row_dimensions[2].height = 20
c = cl.cell(row=3, column=1, value='="Дата аудита: "&IF(\'Паспорт НИОКР\'!C25="","—",TEXT(\'Паспорт НИОКР\'!C25,"DD.MM.YYYY"))&"    |    Комиссия: "&IF(\'Паспорт НИОКР\'!C26="","—",\'Паспорт НИОКР\'!C26)')
cl.merge_cells("A3:H3"); c.font = F_N; c.alignment = Alignment(vertical="center", indent=1)
cl.row_dimensions[3].height = 18
c = cl.cell(row=4, column=1, value="Заполняйте только жёлтые ячейки (D — ответ, F — документальное подтверждение, H — комментарий/обоснование «Н/П»). Баллы (E), контроль правила (G), итоги блоков и интегральный балл рассчитываются автоматически. ★ — вопросы, добавленные в финальной версии чек-листа.")
cl.merge_cells("A4:H4"); c.font = F_SUB; c.alignment = WRAP_TOP
cl.row_dimensions[4].height = 28

hdrs = ["№", "Контрольный вопрос", "Что должен установить аудитор", "Ответ",
        "Балл", "Документальное подтверждение", "Контроль правила подтверждения",
        "Комментарий аудитора; документы; обоснование «Н/П»"]
for j, h in enumerate(hdrs, start=1):
    hdr_cell(cl, 5, j, h)
cl.row_dimensions[5].height = 32

dv_ans = DataValidation(type="list", formula1=f"'Справочник'!{L_ANS}", allow_blank=True,
                        promptTitle="Ответ аудитора", prompt="Выберите один из шести вариантов. «Н/П» — с обоснованием в комментарии.",
                        showInputMessage=True)
dv_ans.error = "Выберите значение из списка: Да, полностью / В основном да / Частично / Скорее нет / Нет / Н/П"
dv_ans.errorTitle = "Недопустимый ответ"
cl.add_data_validation(dv_ans)
dv_conf = DataValidation(type="list", formula1=f"'Справочник'!{L_CONF}", allow_blank=True,
                         promptTitle="Подтверждение", prompt="Есть ли документальное подтверждение ответа? Без него 4–5 баллов не допускаются.",
                         showInputMessage=True)
cl.add_data_validation(dv_conf)

qnum = 0
block_rows = []   # (диапазон строк вопросов, строка итога)
r = 6
for b in BLOCKS:
    first_q = r + 1
    section_row(cl, r, 1, 8, f'БЛОК {b["num"]}. {b["name"].upper()} — вес {b["weight"]}%  (вопросы {qnum+1}–{qnum+len(b["questions"])});  цель: {b["goal"]}', h=24)
    r += 1
    for (q, establish, added, src) in b["questions"]:
        qnum += 1
        star = " ★" if added else ""
        a = cl.cell(row=r, column=1, value=qnum); a.font = F_NB; a.alignment = CTR; a.border = BORDER
        qc = cl.cell(row=r, column=2, value=q + star); qc.font = F_N if not added else Font(name="Calibri", size=10, bold=True)
        qc.alignment = LEFT_TOP; qc.border = BORDER
        ec = cl.cell(row=r, column=3, value=establish); ec.font = F_SM; ec.alignment = LEFT_TOP; ec.border = BORDER
        d = cl.cell(row=r, column=4); d.fill = FILL_INPUT; d.border = BORDER; d.alignment = CTR; d.font = F_NB
        d.comment = None
        e = cl.cell(row=r, column=5, value=f'=IF($D{r}="","",IF($D{r}="Н/П","Н/П",VLOOKUP($D{r},\'Справочник\'!$A$4:$B$8,2,0)))')
        e.font = F_NB; e.alignment = CTR; e.border = BORDER; e.fill = FILL_CALC
        f = cl.cell(row=r, column=6); f.fill = FILL_INPUT; f.border = BORDER; f.alignment = CTR; f.font = F_N
        g = cl.cell(row=r, column=7, value=f'=IF(OR($D{r}="",$D{r}="Н/П"),"",IF(AND(N($E{r})>=4,$F{r}="Нет"),"⚠ 4–5 баллов без подтверждения документом",IF(AND(N($E{r})>=4,$F{r}="Да"),"подтверждено","")))')
        g.font = F_SM; g.alignment = WRAP_CTR; g.border = BORDER; g.fill = FILL_CALC
        h = cl.cell(row=r, column=8); h.fill = FILL_INPUT; h.border = BORDER; h.alignment = LEFT_TOP; h.font = F_SM
        dv_ans.add(d)
        dv_conf.add(f)
        htext = max(est_h(q + star, 44), est_h(establish, 52))
        cl.row_dimensions[r].height = max(30, htext)
        r += 1
    last_q = r - 1
    # строка итога блока
    cl.cell(row=r, column=2, value=f'ИТОГ БЛОКА {b["num"]} →').font = F_LBL
    cl.cell(row=r, column=2).alignment = Alignment(vertical="center", horizontal="right")
    cl.cell(row=r, column=3, value="сумма баллов Sⱼ →  |  применимых вопросов Nⱼ →  |  процент выполнения Pⱼ →").font = F_SM
    cl.cell(row=r, column=3).alignment = Alignment(vertical="center", horizontal="right")
    s = cl.cell(row=r, column=5, value=f"=SUM(E{first_q}:E{last_q})"); s.font = F_B; s.alignment = CTR; s.fill = FILL_CALC; s.border = BORDER
    n = cl.cell(row=r, column=6, value=f"=COUNT(E{first_q}:E{last_q})"); n.font = F_B; n.alignment = CTR; n.fill = FILL_CALC; n.border = BORDER
    p = cl.cell(row=r, column=7, value=f'=IF(F{r}=0,"Н/П",E{r}/(5*F{r}))'); p.font = F_B; p.alignment = CTR; p.fill = FILL_CALC; p.border = BORDER
    p.number_format = "0.0%"
    note = cl.cell(row=r, column=8, value=f'=IF(F{r}=0,"⚠ Все вопросы блока «Н/П» — блок исключён из расчёта (вес {b["weight"]}% не учитывается)","")')
    note.font = Font(name="Calibri", size=9, bold=True, color=RED_T); note.alignment = WRAP_CTR; note.fill = FILL_CALC; note.border = BORDER
    for col in (1, 2, 3, 4):
        cl.cell(row=r, column=col).border = BORDER
        cl.cell(row=r, column=col).fill = FILL_BLUE
    cl.row_dimensions[r].height = 22
    block_rows.append((first_q, last_q, r, b))
    r += 1

# ---------- расчётный блок ----------
calc_hdr_row = r + 1
r += 1
section_row(cl, r, 1, 8, "РАСЧЁТ ИНТЕГРАЛЬНОГО ПОКАЗАТЕЛЯ С ОБРАБОТКОЙ «Н/П»  (все ячейки вычисляются автоматически)", h=24)
r += 1
calc_table_hdr = r
for j, h in enumerate(["", "Блок оценки", "Вес Wⱼ", "Сумма баллов Sⱼ", "Применимо Nⱼ", "Выполнение Pⱼ", "Взвешенный балл Bⱼ", "Вес, учтённый в расчёте I"], start=1):
    if j == 1:
        cl.cell(row=r, column=1).fill = FILL_HDR; cl.cell(row=r, column=1).border = BORDER
    else:
        hdr_cell(cl, r, j, h)
cl.row_dimensions[r].height = 30
r += 1
calc_first = r
for i, (first_q, last_q, sub_row, b) in enumerate(block_rows):
    cl.cell(row=r, column=2, value=f'{b["num"]}. {b["name"]}').font = F_NB
    cl.cell(row=r, column=2).border = BORDER
    cW = cl.cell(row=r, column=3, value=f"='Справочник'!F{12+i}"); cW.font = F_N; cW.alignment = CTR; cW.border = BORDER; cW.fill = FILL_CALC
    cS = cl.cell(row=r, column=4, value=f"=E{sub_row}"); cS.font = F_N; cS.alignment = CTR; cS.border = BORDER; cS.fill = FILL_CALC
    cN = cl.cell(row=r, column=5, value=f"=F{sub_row}"); cN.font = F_N; cN.alignment = CTR; cN.border = BORDER; cN.fill = FILL_CALC
    cP = cl.cell(row=r, column=6, value=f'=IF(E{r}=0,"Н/П",D{r}/(5*E{r}))'); cP.font = F_NB; cP.alignment = CTR; cP.border = BORDER; cP.fill = FILL_CALC
    cP.number_format = "0.0%"
    cB = cl.cell(row=r, column=7, value=f"=IF(E{r}=0,0,C{r}*F{r})"); cB.font = F_NB; cB.alignment = CTR; cB.border = BORDER; cB.fill = FILL_CALC
    cB.number_format = "0.00"
    cWap = cl.cell(row=r, column=8, value=f"=IF(E{r}=0,0,C{r})"); cWap.font = F_N; cWap.alignment = CTR; cWap.border = BORDER; cWap.fill = FILL_CALC
    cl.row_dimensions[r].height = 20
    r += 1
calc_last = r - 1
# итого
cl.cell(row=r, column=2, value="ИТОГО").font = F_LBL; cl.cell(row=r, column=2).border = BORDER
cW = cl.cell(row=r, column=3, value=f"=SUM(C{calc_first}:C{calc_last})"); cW.font = F_NB; cW.alignment = CTR; cW.border = BORDER
cN = cl.cell(row=r, column=5, value=f"=SUM(E{calc_first}:E{calc_last})"); cN.font = F_NB; cN.alignment = CTR; cN.border = BORDER
cB = cl.cell(row=r, column=7, value=f"=SUM(G{calc_first}:G{calc_last})"); cB.font = F_NB; cB.alignment = CTR; cB.border = BORDER
cB.number_format = "0.00"
cWap = cl.cell(row=r, column=8, value=f"=SUM(H{calc_first}:H{calc_last})"); cWap.font = F_NB; cWap.alignment = CTR; cWap.border = BORDER
for col in (4, 6):
    cl.cell(row=r, column=col, value="—").alignment = CTR
    cl.cell(row=r, column=col).border = BORDER
cl.row_dimensions[r].height = 20
total_row = r
r += 2

# интегральный балл
int_row = r
cl.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
c = cl.cell(row=r, column=2, value="ИНТЕГРАЛЬНЫЙ ПОКАЗАТЕЛЬ  I  (0–100)")
c.font = F_LBL; c.alignment = Alignment(vertical="center", horizontal="right")
cI = cl.cell(row=r, column=7, value=f'=IF(H{total_row}=0,"Н/П",100*G{total_row}/H{total_row})')
cI.font = F_BIG; cI.alignment = CTR; cI.border = BORDER; cI.number_format = "0.0"
cI.fill = FILL_CALC
cnote = cl.cell(row=r, column=8, value=f'=IF(H{total_row}=0,"⚠ Недостаточно данных: все вопросы «Н/П»",IF(H{total_row}<100,"Применена ренормировка весов: I = 100·ΣBⱼ / ΣWⱼ по применимым блокам","Расчёт по полной схеме весов 30/35/35"))')
cnote.font = F_SM; cnote.alignment = WRAP_CTR; cnote.border = BORDER
cl.row_dimensions[r].height = 40
r += 1
ver_row = r
cl.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
c = cl.cell(row=r, column=2, value="ОЦЕНКА СОСТОЯНИЯ НИОКР")
c.font = F_LBL; c.alignment = Alignment(vertical="center", horizontal="right")
cV = cl.cell(row=r, column=7, value=f'=IF(H{total_row}=0,"Недостаточно данных",IF(G{int_row}>=85,"Высокая",IF(G{int_row}>=70,"Хорошая",IF(G{int_row}>=50,"Удовлетворительная",IF(G{int_row}>=30,"Неудовлетворительная","Критическая")))))')
cV.font = Font(name="Calibri", size=13, bold=True); cV.alignment = CTR; cV.border = BORDER; cV.fill = FILL_CALC
cl.cell(row=r, column=8).border = BORDER
cl.row_dimensions[r].height = 26
r += 1
dec_row = r
cl.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
c = cl.cell(row=r, column=2, value="РЕКОМЕНДУЕМОЕ УПРАВЛЕНЧЕСКОЕ РЕШЕНИЕ")
c.font = F_LBL; c.alignment = Alignment(vertical="center", horizontal="right")
cl.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
cD = cl.cell(row=r, column=7, value=f'=IF(H{total_row}=0,"Собрать данные и повторить оценку",IF(G{int_row}>=85,"Продолжить финансирование, внедрять, масштабировать",IF(G{int_row}>=70,"Продолжить с устранением выявленных недостатков",IF(G{int_row}>=50,"Пересмотреть ТЗ, бюджет, сроки и план внедрения",IF(G{int_row}>=30,"Приостановить, провести дополнительную экспертизу","Рассмотреть прекращение работ и списание в установленном порядке")))))')
cD.font = Font(name="Calibri", size=11, bold=True, color=NAVY); cD.alignment = WRAP_CTR; cD.border = BORDER; cD.fill = FILL_CALC
cl.row_dimensions[r].height = 30
r += 2
c = cl.cell(row=r, column=2, value="Важно: интегральный показатель — инструмент поддержки решения, а не самостоятельное основание для прекращения НИОКР или списания активов. Контрольные условия проверяются в разделе 5 листа «Паспорт НИОКР».")
cl.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
c.font = F_SUB; c.alignment = WRAP_TOP; cl.row_dimensions[r].height = 26

CL = {"int_row": int_row, "ver_row": ver_row, "dec_row": dec_row, "calc_first": calc_first, "calc_last": calc_last, "total_row": total_row}
with open("scripts/_cl_rows.py", "w") as fo:
    fo.write(repr(CL))

cl.freeze_panes = "A6"

# ---------- условное форматирование чек-листа ----------
q_ranges = " ".join([f"D{f}:D{l}" for f, l, _, _ in block_rows])
e_ranges = " ".join([f"E{f}:E{l}" for f, l, _, _ in block_rows])
g_ranges = " ".join([f"G{f}:G{l}" for f, l, _, _ in block_rows])

answer_colors = [("Да, полностью", GREEN_D, GREEN_D_T), ("В основном да", GREEN_L, GREEN_L_T),
                 ("Частично", YELLOW, YELLOW_T), ("Скорее нет", ORANGE, ORANGE_T),
                 ("Нет", RED, RED_T), ("Н/П", GRAY, GRAY_T)]
for text, fill, fcolor in answer_colors:
    cl.conditional_formatting.add(q_ranges, CellIsRule(operator="equal", formula=[f'"{text}"'],
                                                       fill=PatternFill("solid", fgColor=fill),
                                                       font=Font(color=fcolor, bold=True)))
ball_colors = [(5, GREEN_D, GREEN_D_T), (4, GREEN_L, GREEN_L_T), (3, YELLOW, YELLOW_T),
               (1, ORANGE, ORANGE_T), (0, RED, RED_T)]
for val, fill, fcolor in ball_colors:
    cl.conditional_formatting.add(e_ranges, CellIsRule(operator="equal", formula=[str(val)],
                                                       fill=PatternFill("solid", fgColor=fill),
                                                       font=Font(color=fcolor, bold=True)))
cl.conditional_formatting.add(e_ranges, CellIsRule(operator="equal", formula=['"Н/П"'],
                                                   fill=PatternFill("solid", fgColor=GRAY),
                                                   font=Font(color=GRAY_T, bold=True)))
cl.conditional_formatting.add(g_ranges, FormulaRule(formula=[f'ISNUMBER(SEARCH("⚠",G{block_rows[0][0]}))'],
                                                    fill=PatternFill("solid", fgColor=RED),
                                                    font=Font(color=RED_T, bold=True)))
sub_ranges = " ".join([f"G{s}" for _, _, s, _ in block_rows])
cl.conditional_formatting.add(sub_ranges, ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                         mid_type="num", mid_value=0.5, mid_color="FFEB84",
                                                         end_type="num", end_value=1, end_color="63BE7B"))
cl.conditional_formatting.add(f"F{calc_first}:F{calc_last}", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                                            mid_type="num", mid_value=0.5, mid_color="FFEB84",
                                                                            end_type="num", end_value=1, end_color="63BE7B"))
cl.conditional_formatting.add(f"G{int_row}", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                            mid_type="num", mid_value=50, mid_color="FFEB84",
                                                            end_type="num", end_value=100, end_color="63BE7B"))
for txt, fill, fcolor in [("Высокая", GREEN_D, GREEN_D_T), ("Хорошая", GREEN_L, GREEN_L_T),
                          ("Удовлетворительная", YELLOW, YELLOW_T), ("Неудовлетворительная", ORANGE, ORANGE_T),
                          ("Критическая", RED, RED_T), ("Недостаточно данных", GRAY, GRAY_T)]:
    cl.conditional_formatting.add(f"G{ver_row}", CellIsRule(operator="equal", formula=[f'"{txt}"'],
                                                            fill=PatternFill("solid", fgColor=fill),
                                                            font=Font(color=fcolor, bold=True, size=13)))

# =====================================================================
# ЛИСТ «Паспорт НИОКР»
# =====================================================================
ps = wb.create_sheet("Паспорт НИОКР")
ps.sheet_view.showGridLines = False
ps_widths = {"A": 2.5, "B": 44, "C": 46, "D": 15, "E": 12, "F": 12, "G": 15, "H": 44}
for k, v in ps_widths.items():
    ps.column_dimensions[k].width = v

def input_cell(ws_, row, col, fmt=None):
    c = ws_.cell(row=row, column=col)
    c.fill = FILL_INPUT; c.border = BORDER; c.font = F_N
    c.alignment = Alignment(vertical="center", wrap_text=True)
    if fmt:
        c.number_format = fmt
    return c

def label_cell(ws_, row, col, text, font=F_NB):
    c = ws_.cell(row=row, column=col, value=text)
    c.font = font; c.border = BORDER; c.alignment = Alignment(vertical="center", wrap_text=True)
    return c

title_row(ps, 1, "ПАСПОРТ НИОКР — ОБЪЕКТ АУДИТА", 8, h=30, align=Alignment(vertical="center", horizontal="left", indent=1))
# сдвинем заголовок: используемmerge B..H
ps.merge_cells("B2:H2")
c = ps.cell(row=2, column=2, value="Заполните жёлтые ячейки. Раздел 4 «Результаты аудита» рассчитывается автоматически по листу «Чек-лист».")
c.font = F_SUB; c.alignment = Alignment(vertical="center", indent=1)

section_row(ps, 4, 2, 8, "1. ОБЩИЕ СВЕДЕНИЯ", h=22)
fields1 = [
    (5, "Наименование НИОКР", None),
    (6, "Шифр НИОКР", None),
    (7, "Исполнитель (головной, соисполнители)", None),
    (8, "Заказчик (государственный / корпоративный)", None),
    (9, "Основание выполнения (ГОЗ, госконтракт, инициативная, внутренний заказ)", None),
    (10, "Договор: номер, дата, стоимость", None),
    (11, "Сроки выполнения: начало — окончание по договору", None),
    (12, "Текущий статус", "status"),
    (13, "Ответственное подразделение, руководитель", None),
    (14, "Руководитель НИОКР (должность, ФИО)", None),
]
for row, lbl, kind in fields1:
    label_cell(ps, row, 2, lbl)
    input_cell(ps, row, 3)
    ps.merge_cells(start_row=row, start_column=3, end_row=row, end_column=8)
    ps.row_dimensions[row].height = 20
dv_status = DataValidation(type="list", formula1=f"'Справочник'!{L_STATUS}", allow_blank=True)
ps.add_data_validation(dv_status)
dv_status.add(ps.cell(row=12, column=3))

section_row(ps, 16, 2, 8, "2. СТОИМОСТЬ И ФИНАНСИРОВАНИЕ, млн руб.", h=22)
fin_rows = [
    (17, "Утверждённая (первоначальная) стоимость", "input"),
    (18, "Скорректированная стоимость", "input"),
    (19, "Фактические затраты", "input"),
    (20, "Прогноз затрат до завершения", "input"),
    (21, "Прогнозная итоговая стоимость", "=IF(COUNT(C19,C20)=2,C19+C20,\"\")"),
    (22, "Отклонение прогноза от утверждённой стоимости", "=IF(OR(C21=\"\",C17=\"\"),\"\",C21-C17)"),
    (23, "Отклонение, %", "=IF(OR(C21=\"\",C17=\"\",C17=0),\"\",C21/C17-1)"),
]
for row, lbl, f in fin_rows:
    label_cell(ps, row, 2, lbl)
    if f == "input":
        c = input_cell(ps, row, 3, fmt="#,##0.00")
    else:
        c = ps.cell(row=row, column=3, value=f)
        c.fill = FILL_CALC; c.border = BORDER; c.font = F_NB
        c.alignment = Alignment(vertical="center", horizontal="left")
        c.number_format = "0.0%" if row == 23 else "#,##0.00"
    note = ps.cell(row=row, column=4, value="автоматически" if f != "input" else "")
    note.font = F_SM
    ps.merge_cells(start_row=row, start_column=4, end_row=row, end_column=8)
    ps.row_dimensions[row].height = 18

section_row(ps, 24, 2, 8, "3. ДАННЫЕ АУДИТА", h=20)
label_cell(ps, 25, 2, "Дата проведения аудита")
c = input_cell(ps, 25, 3, fmt="DD.MM.YYYY")
ps.merge_cells(start_row=25, start_column=3, end_row=25, end_column=8)
label_cell(ps, 26, 2, "Состав комиссии")
input_cell(ps, 26, 3)
ps.merge_cells(start_row=26, start_column=3, end_row=26, end_column=8)
ps.row_dimensions[25].height = 18; ps.row_dimensions[26].height = 18

# ---- раздел 4: результаты аудита (авто) ----
section_row(ps, 28, 2, 8, "4. РЕЗУЛЬТАТЫ АУДИТА (рассчитываются автоматически по листу «Чек-лист»)", h=22)
hdrs4 = ["Блок оценки", "Вес Wⱼ", "Сумма Sⱼ", "Применимо Nⱼ", "Выполнение Pⱼ", "Взвешенный балл Bⱼ", "Диапазон"]
for j, h in enumerate(hdrs4, start=2):
    hdr_cell(ps, 29, j, h)
ps.row_dimensions[29].height = 28
for i, b in enumerate(BLOCKS):
    row = 30 + i
    cf = CL["calc_first"] + i
    label_cell(ps, row, 2, f'{b["num"]}. {b["name"]}')
    for col, formula, fmt in [
        (3, f"='Чек-лист'!C{cf}", "0"),
        (4, f"='Чек-лист'!D{cf}", "0"),
        (5, f"='Чек-лист'!E{cf}", "0"),
        (6, f"='Чек-лист'!F{cf}", "0.0%"),
        (7, f"='Чек-лист'!G{cf}", "0.00"),
    ]:
        c = ps.cell(row=row, column=col, value=formula)
        c.fill = FILL_CALC; c.border = BORDER; c.font = F_N; c.alignment = CTR
        c.number_format = fmt
    c = ps.cell(row=row, column=8, value=f"0 – {b['weight']:.0f} баллов")
    c.font = F_SM; c.border = BORDER; c.alignment = CTR
    ps.row_dimensions[row].height = 19
row = 33
label_cell(ps, row, 2, "ИТОГО")
c = ps.cell(row=row, column=3, value="=SUM(C30:C32)"); c.font = F_NB; c.alignment = CTR; c.border = BORDER
c = ps.cell(row=row, column=7, value="=SUM(G30:G32)"); c.font = F_NB; c.alignment = CTR; c.border = BORDER; c.number_format = "0.00"
for col in (4, 5, 6):
    ps.cell(row=row, column=col, value="—").alignment = CTR
    ps.cell(row=row, column=col).border = BORDER
c = ps.cell(row=row, column=8, value="0 – 100 баллов"); c.font = F_SM; c.border = BORDER; c.alignment = CTR

ps.merge_cells("B35:F35")
c = ps.cell(row=35, column=2, value="ИНТЕГРАЛЬНЫЙ ПОКАЗАТЕЛЬ I (0–100)")
c.font = F_LBL; c.alignment = Alignment(vertical="center", horizontal="right")
cI = ps.cell(row=35, column=7, value=f"='Чек-лист'!G{CL['int_row']}")
cI.font = F_BIG; cI.alignment = CTR; cI.border = BORDER; cI.number_format = "0.0"; cI.fill = FILL_CALC
c = ps.cell(row=35, column=8, value=f"='Чек-лист'!H{CL['int_row']}")
c.font = F_SM; c.alignment = WRAP_CTR; c.border = BORDER
ps.row_dimensions[35].height = 42

ps.merge_cells("B36:F36")
c = ps.cell(row=36, column=2, value="Оценка состояния НИОКР")
c.font = F_LBL; c.alignment = Alignment(vertical="center", horizontal="right")
cV = ps.cell(row=36, column=7, value=f"='Чек-лист'!G{CL['ver_row']}")
cV.font = Font(name="Calibri", size=13, bold=True); cV.alignment = CTR; cV.border = BORDER; cV.fill = FILL_CALC
ps.cell(row=36, column=8).border = BORDER
ps.row_dimensions[36].height = 26

ps.merge_cells("B37:F37")
c = ps.cell(row=37, column=2, value="Рекомендуемое решение (автоматически по шкале)")
c.font = F_LBL; c.alignment = Alignment(vertical="center", horizontal="right")
ps.merge_cells("G37:H37")
cD = ps.cell(row=37, column=7, value=f"='Чек-лист'!G{CL['dec_row']}")
cD.font = F_NB; cD.alignment = WRAP_CTR; cD.border = BORDER; cD.fill = FILL_CALC
ps.row_dimensions[37].height = 30

# подсветка оценки
for txt, fill, fcolor in [("Высокая", GREEN_D, GREEN_D_T), ("Хорошая", GREEN_L, GREEN_L_T),
                          ("Удовлетворительная", YELLOW, YELLOW_T), ("Неудовлетворительная", ORANGE, ORANGE_T),
                          ("Критическая", RED, RED_T), ("Недостаточно данных", GRAY, GRAY_T)]:
    ps.conditional_formatting.add("G36", CellIsRule(operator="equal", formula=[f'"{txt}"'],
                                                    fill=PatternFill("solid", fgColor=fill),
                                                    font=Font(color=fcolor, bold=True, size=13)))
ps.conditional_formatting.add("G35", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                    mid_type="num", mid_value=50, mid_color="FFEB84",
                                                    end_type="num", end_value=100, end_color="63BE7B"))
ps.conditional_formatting.add("F30:F33", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                        mid_type="num", mid_value=0.5, mid_color="FFEB84",
                                                        end_type="num", end_value=1, end_color="63BE7B"))

# ---- раздел 5: контрольные условия ----
section_row(ps, 39, 2, 8, "5. КОНТРОЛЬНЫЕ УСЛОВИЯ, ВЛИЯЮЩИЕ НА ИТОГОВОЕ РЕШЕНИЕ (проверяются независимо от баллов)", h=22)
hdr_cell(ps, 40, 2, "№")
hdr_cell(ps, 40, 3, "Контрольное условие")
hdr_cell(ps, 40, 4, "Статус")
ps.merge_cells("E40:H40")
hdr_cell(ps, 40, 5, "Последствие при выявлении")
ps.row_dimensions[40].height = 24
dv_ku = DataValidation(type="list", formula1=f"'Справочник'!{L_KU}", allow_blank=True)
ps.add_data_validation(dv_ku)
for i, (cond, cons) in enumerate(CONTROL_CONDITIONS):
    row = 41 + i
    c = ps.cell(row=row, column=2, value=i + 1); c.font = F_N; c.alignment = CTR; c.border = BORDER
    c = ps.cell(row=row, column=3, value=cond); c.font = F_N; c.alignment = LEFT_TOP; c.border = BORDER
    c = ps.cell(row=row, column=4); c.fill = FILL_INPUT; c.border = BORDER; c.alignment = CTR; c.font = F_NB
    dv_ku.add(c)
    ps.merge_cells(start_row=row, start_column=5, end_row=row, end_column=8)
    c = ps.cell(row=row, column=5, value=cons); c.font = F_SM; c.alignment = LEFT_TOP; c.border = BORDER
    for col in range(6, 9):
        ps.cell(row=row, column=col).border = BORDER
    ps.row_dimensions[row].height = max(20, est_h(cond, 44, minh=20))
ps.conditional_formatting.add("D41:D47", CellIsRule(operator="equal", formula=['"Выявлено"'],
                                                     fill=PatternFill("solid", fgColor=RED),
                                                     font=Font(color=RED_T, bold=True)))
ps.conditional_formatting.add("D41:D47", CellIsRule(operator="equal", formula=['"Требует проверки"'],
                                                     fill=PatternFill("solid", fgColor=YELLOW),
                                                     font=Font(color=YELLOW_T, bold=True)))
ps.conditional_formatting.add("D41:D47", CellIsRule(operator="equal", formula=['"Не выявлено"'],
                                                     fill=PatternFill("solid", fgColor=GREEN_L),
                                                     font=Font(color=GREEN_L_T)))
row = 48
ps.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
c = ps.cell(row=row, column=2, value='=IF(COUNTIF(D41:D47,"Выявлено")>0,"⚠ ВЫЯВЛЕНО КОНТРОЛЬНЫХ УСЛОВИЙ: "&COUNTIF(D41:D47,"Выявлено")&" — итоговое решение комиссия принимает с ОБЯЗАТЕЛЬНЫМ учётом последствий (строки выше). Балльная оценка не отменяет действие контрольных условий.",IF(COUNTIF(D41:D47,"Требует проверки")>0,"Есть условия со статусом «Требует проверки» — завершите проверку до подписания акта.","Контрольные условия не выявлены."))')
c.font = Font(name="Calibri", size=10, bold=True, color=RED_T); c.alignment = WRAP_TOP; c.border = BORDER; c.fill = FILL_CALC
ps.row_dimensions[row].height = 40

# ---- раздел 6: итоговое решение комиссии ----
section_row(ps, 50, 2, 8, "6. ИТОГОВОЕ РЕШЕНИЕ КОМИССИИ", h=22)
label_cell(ps, 51, 2, "Решение комиссии (выберите из списка)")
c = input_cell(ps, 51, 3); c.font = F_NB
ps.merge_cells(start_row=51, start_column=3, end_row=51, end_column=8)
dv_dec = DataValidation(type="list", formula1=f"'Справочник'!{L_DEC}", allow_blank=True)
ps.add_data_validation(dv_dec)
dv_dec.add(ps.cell(row=51, column=3))
ps.row_dimensions[51].height = 22

label_cell(ps, 52, 2, "Обоснование решения")
input_cell(ps, 52, 3); ps.merge_cells("C52:H53"); ps.row_dimensions[52].height = 20; ps.row_dimensions[53].height = 20
label_cell(ps, 54, 2, "Мероприятия по результатам аудита (содержание, ответственный, срок)")
input_cell(ps, 54, 3); ps.merge_cells("C54:H55"); ps.row_dimensions[54].height = 20; ps.row_dimensions[55].height = 20
label_cell(ps, 56, 2, "Выявленные недостатки и проблемы")
input_cell(ps, 56, 3); ps.merge_cells("C56:H57"); ps.row_dimensions[56].height = 20; ps.row_dimensions[57].height = 20
label_cell(ps, 58, 2, "Сильные стороны и положительные результаты")
input_cell(ps, 58, 3); ps.merge_cells("C58:H59"); ps.row_dimensions[58].height = 20; ps.row_dimensions[59].height = 20

label_cell(ps, 61, 2, "Председатель комиссии (ФИО, подпись)")
input_cell(ps, 61, 3); ps.merge_cells(start_row=61, start_column=3, end_row=61, end_column=5)
label_cell(ps, 61, 6, "Дата", font=F_SM)
input_cell(ps, 61, 7); ps.merge_cells(start_row=61, start_column=7, end_row=61, end_column=8)
label_cell(ps, 62, 2, "Члены комиссии (ФИО, подписи)")
input_cell(ps, 62, 3); ps.merge_cells(start_row=62, start_column=3, end_row=62, end_column=8)
ps.row_dimensions[61].height = 22; ps.row_dimensions[62].height = 22

# =====================================================================
# ЛИСТ «Портфель»
# =====================================================================
pf = wb.create_sheet("Портфель")
pf.sheet_view.showGridLines = False
pf_widths = {"A": 5, "B": 13, "C": 40, "D": 15, "E": 12, "F": 11, "G": 11, "H": 34, "I": 15, "J": 30, "K": 30}
for k, v in pf_widths.items():
    pf.column_dimensions[k].width = v

title_row(pf, 1, "ПОРТФЕЛЬ НИОКР ПОДРАЗДЕЛЕНИЯ — СВОДНАЯ МАТРИЦА РЕЗУЛЬТАТОВ АУДИТА", 11, h=30)
c = pf.cell(row=2, column=1, value="Подразделение: ____________________    Дата актуальности сведений: __.__.20__    Строки 8–10 — иллюстративный пример формы (значения условные), удалите их при заполнении.")
pf.merge_cells("A2:K2"); c.font = F_SUB; c.alignment = Alignment(vertical="center", indent=1)

hdrs = ["№", "Шифр", "Наименование НИОКР", "Статус", "Стоимость, млн ₽", "Готовность, %",
        "Интегральный балл I", "Категория состояния (автоматически)", "Контрольные условия",
        "Управленческое решение", "Риски / примечания"]
for j, h in enumerate(hdrs, start=1):
    hdr_cell(pf, 4, j, h)
pf.row_dimensions[4].height = 32

cat_formula = ('=IF($G{r}="","",IF($G{r}>=85,"Высокая — внедрять, масштабировать",'
               'IF($G{r}>=70,"Хорошая — продолжить с устранением недостатков",'
               'IF($G{r}>=50,"Удовлетворительная — пересмотреть ТЗ, бюджет, сроки",'
               'IF($G{r}>=30,"Неудовлетворительная — приостановить, доп. экспертиза",'
               '"Критическая — рассмотреть прекращение и списание")))))')

examples = [
    (1, "НИОКР-01", "Разработка А (пример)", "Выполняется", 120, 0.85, 91, "Внедрить"),
    (2, "НИОКР-02", "Разработка Б (пример)", "Выполняется", 250, 0.55, 74, "Продолжить"),
    (3, "НИОКР-03", "Разработка В (пример)", "Приостановлена", 80, 0.30, 48, "Продолжить с корректировкой"),
    (4, "НИОКР-04", "Разработка Г (пример)", "Выполняется", 65, 0.20, 25, "Прекратить"),
]
dv_pdec = DataValidation(type="list", formula1=f"'Справочник'!{L_PORT}", allow_blank=True)
dv_pstat = DataValidation(type="list", formula1=f"'Справочник'!{L_STATUS}", allow_blank=True)
dv_ku2 = DataValidation(type="list", formula1=f"'Справочник'!{L_KU}", allow_blank=True)
pf.add_data_validation(dv_pdec); pf.add_data_validation(dv_pstat); pf.add_data_validation(dv_ku2)

first_data = 5
for i in range(18):
    row = first_data + i
    ex = examples[i] if i < len(examples) else None
    c = pf.cell(row=row, column=1, value=(ex[0] if ex else None)); c.font = F_N; c.alignment = CTR; c.border = BORDER
    for col, val in ((2, ex[1] if ex else None), (3, ex[2] if ex else None)):
        c = pf.cell(row=row, column=col, value=val); c.font = F_N; c.alignment = LEFT_TOP; c.border = BORDER
        c.fill = FILL_INPUT
    c = pf.cell(row=row, column=4, value=ex[3] if ex else None); c.fill = FILL_INPUT; c.border = BORDER; c.alignment = CTR; c.font = F_N
    dv_pstat.add(c)
    c = pf.cell(row=row, column=5, value=ex[4] if ex else None); c.fill = FILL_INPUT; c.border = BORDER; c.alignment = CTR; c.font = F_N; c.number_format = "#,##0.0"
    c = pf.cell(row=row, column=6, value=ex[5] if ex else None); c.fill = FILL_INPUT; c.border = BORDER; c.alignment = CTR; c.font = F_N; c.number_format = "0%"
    c = pf.cell(row=row, column=7, value=ex[6] if ex else None); c.fill = FILL_INPUT; c.border = BORDER; c.alignment = CTR; c.font = F_NB; c.number_format = "0.0"
    c = pf.cell(row=row, column=8, value=cat_formula.format(r=row)); c.fill = FILL_CALC; c.border = BORDER; c.alignment = LEFT_TOP; c.font = F_SM
    c = pf.cell(row=row, column=9, value=None); c.fill = FILL_INPUT; c.border = BORDER; c.alignment = CTR; c.font = F_N
    dv_ku2.add(c)
    c = pf.cell(row=row, column=10, value=ex[7] if ex else None); c.fill = FILL_INPUT; c.border = BORDER; c.alignment = LEFT_TOP; c.font = F_N
    dv_pdec.add(c)
    c = pf.cell(row=row, column=11, value=None); c.fill = FILL_INPUT; c.border = BORDER; c.alignment = LEFT_TOP; c.font = F_SM
    pf.row_dimensions[row].height = 20
last_data = first_data + 17

pf.conditional_formatting.add(f"G{first_data}:G{last_data}", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                                            mid_type="num", mid_value=50, mid_color="FFEB84",
                                                                            end_type="num", end_value=100, end_color="63BE7B"))
pf.conditional_formatting.add(f"F{first_data}:F{last_data}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=ACCENT))
for txt, fill, fcolor in [("Высокая", GREEN_D, GREEN_D_T), ("Хорошая", GREEN_L, GREEN_L_T),
                          ("Удовлетворительная", YELLOW, YELLOW_T), ("Неудовлетворительная", ORANGE, ORANGE_T),
                          ("Критическая", RED, RED_T)]:
    pf.conditional_formatting.add(f"H{first_data}:H{last_data}", FormulaRule(formula=[f'ISNUMBER(SEARCH("{txt}",$H{first_data}))'],
                                                                             fill=PatternFill("solid", fgColor=fill),
                                                                             font=Font(color=fcolor)))

# сводные счётчики
sr = last_data + 2
section_row(pf, sr, 1, 11, "СВОДНЫЕ ПОКАЗАТЕЛИ ПОРТФЕЛЯ (рассчитываются автоматически)", h=22)
counters = [
    ("Всего НИОКР в матрице", f'=COUNTA(C{first_data}:C{last_data})'),
    ("Выполняется", f'=COUNTIF(D{first_data}:D{last_data},"Выполняется")'),
    ("Завершено", f'=COUNTIF(D{first_data}:D{last_data},"Завершена")'),
    ("Приостановлено", f'=COUNTIF(D{first_data}:D{last_data},"Приостановлена")'),
    ("С высокой оценкой (I ≥ 85)", f'=COUNTIF(G{first_data}:G{last_data},">=85")'),
    ("Требует пересмотра (50 ≤ I < 70)", f'=COUNTIFS(G{first_data}:G{last_data},">=50",G{first_data}:G{last_data},"<70")'),
    ("Кандидаты на приостановку/прекращение (I < 50)", f'=COUNTIFS(G{first_data}:G{last_data},">=0",G{first_data}:G{last_data},"<50")'),
    ("Предложено к прекращению/списанию", f'=COUNTIF(J{first_data}:J{last_data},"Прекратить")+COUNTIF(J{first_data}:J{last_data},"Списать")'),
    ("Суммарная стоимость, млн ₽", f'=SUM(E{first_data}:E{last_data})'),
    ("Средний интегральный балл", f'=IF(COUNT(G{first_data}:G{last_data})=0,"—",ROUND(AVERAGE(G{first_data}:G{last_data}),1))'),
    ("Выявлены контрольные условия", f'=COUNTIF(I{first_data}:I{last_data},"Выявлено")'),
]
for i, (lbl, formula) in enumerate(counters):
    col_b = 2 + (i % 2) * 5   # B или G
    row = sr + 1 + (i // 2)
    if i % 2 == 0:
        c = pf.cell(row=row, column=2, value=lbl)
        pf.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
        c.font = F_N; c.alignment = Alignment(vertical="center", indent=1)
        v = pf.cell(row=row, column=6, value=formula)
    else:
        c = pf.cell(row=row, column=7, value=lbl)
        pf.merge_cells(start_row=row, start_column=7, end_row=row, end_column=10)
        c.font = F_N; c.alignment = Alignment(vertical="center", indent=1)
        v = pf.cell(row=row, column=11, value=formula)
    v.font = F_NB; v.alignment = CTR; v.fill = FILL_CALC; v.border = BORDER
    pf.row_dimensions[row].height = 18
note_r = sr + 1 + math.ceil(len(counters) / 2) + 1
c = pf.cell(row=note_r, column=2, value="Портфельная карта позволяет руководству видеть не только состояние отдельных НИОКР, но и эффективность управления всем научно-техническим портфелем: приоритетные разработки, проблемные и дублирующие работы, кандидаты на внедрение в ближайший год, направления перераспределения ресурсов.")
pf.merge_cells(start_row=note_r, start_column=2, end_row=note_r + 1, end_column=11)
c.font = F_SUB; c.alignment = WRAP_TOP
pf.freeze_panes = "A5"

# ---------- порядок листов ----------
order = ["Инструкция", "Паспорт НИОКР", "Чек-лист", "Портфель", "Справочник"]
wb._sheets = [wb[name] for name in order]
wb.active = 1

# печать
for sheet in wb.worksheets:
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True

import os
os.makedirs("output", exist_ok=True)
OUT = "output/1_Чек-лист_аудит_НИОКР_шаблон_расчёта.xlsx"
wb.save(OUT)
print("Сохранено:", OUT)
print("Строки чек-листа:", CL)
