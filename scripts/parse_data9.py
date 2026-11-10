# -*- coding: utf-8 -*-
# Этап 18А: парсер таблиц методики «Научно.docx» (scripts/data/nauchno.txt) → scripts/data/data9.json
# Все структуры, необходимые для генератора Excel-калькулятора v2 (9_Калькулятор_аудита_НИОКР_v2.xlsx).
import io, re, json

txt = io.open('scripts/data/nauchno.txt', encoding='utf-8').read()
tables = []; cur = None
for ln in txt.split('\n'):
    if ln == '[TABLE]': cur = []
    elif ln == '[/TABLE]':
        if cur is not None: tables.append(cur)
        cur = None
    elif cur is not None: cur.append([c.strip() for c in ln.split(' | ')])

def find(pred): return [t for t in tables if t and pred(t)]
def w(x): return float(x.replace(',', '.'))

# ---------- КРАТКИЙ 15 (Приложение №1, разделы А/Б/В по 5 вопросов, якоря 5/3/0) ----------
k15 = find(lambda t: t[0][:2] == ['№', 'Вопрос'] and '5 баллов' in t[0])
assert len(k15) == 3 and all(len(t) == 6 for t in k15), [len(t) for t in k15]
Q15 = []
for sec, t in zip(['А', 'Б', 'В'], k15):
    for row in t[1:]:
        assert len(row) == 5, row
        Q15.append(dict(sec=sec, n=int(row[0]), q=row[1], a5=row[2], a3=row[3], a0=row[4]))
assert len(Q15) == 15 and [q['n'] for q in Q15] == list(range(1, 16))

# ---------- СРЕДНИЙ 30 (Приложение №2: А1-А10, Б1-Б10, В1-В10, веса) ----------
s30 = find(lambda t: len(t[0]) == 5 and t[0][1] == 'Контрольный вопрос' and t[0][4] == 'Вес')
assert len(s30) == 3 and all(len(t) == 11 for t in s30)
Q30 = []
for t in s30:
    for row in t[1:]:
        assert re.fullmatch(r'[АБВ]\d{1,2}', row[0]), row
        Q30.append(dict(code=row[0], q=row[1], check=row[2], ev=row[3], w=w(row[4])))
assert len(Q30) == 30
for pre in 'АБВ':
    ws = [q['w'] for q in Q30 if q['code'][0] == pre]
    assert len(ws) == 10 and abs(sum(ws) - 10.0) < 1e-9, (pre, sum(ws))

def anchors(marker):
    i = txt.index(marker); seg = txt[i:i+1200]; out = {}
    for m in re.finditer(r'^([0-5]) — (.+)$', seg, re.M):
        out[str(int(m.group(1)))] = m.group(2).strip()
        if len(out) == 6: break
    return out
ANCH30 = {'А1': anchors('А1. Подтверждение потребности'),
          'А6': anchors('А6. Анализ аналогов'),
          'А7': anchors('А7. Отсутствие дублирования')}
for k, v in ANCH30.items(): assert len(v) == 6, (k, list(v))

# ---------- ПОЛНЫЙ 60 (Приложение №3: 12 блоков) ----------
f60 = find(lambda t: len(t[0]) == 4 and t[0][1] == 'Вопрос' and t[0][2] == 'Что необходимо установить')
assert [len(t)-1 for t in f60] == [4,4,4,8,4,4,4,8,6,4,6,4]
Q60 = []
for t in f60:
    for row in t[1:]:
        assert re.fullmatch(r'\d{1,2}', row[0]), row
        Q60.append(dict(n=int(row[0]), q=row[1], est=row[2], ev=row[3]))
assert [q['n'] for q in Q60] == list(range(1, 61))
BLOCKS60 = [['А1',1,4,0.20],['А2',5,8,0.20],['А3',9,12,0.20],['А4',13,20,0.40],
            ['Б1',21,24,0.25],['Б2',25,28,0.25],['Б3',29,32,0.25],['Б4',33,40,0.25],
            ['В1',41,46,0.30],['В2',47,50,0.25],['В3',51,56,0.30],['В4',57,60,0.15]]
assert abs(sum(b[3] for b in BLOCKS60 if b[0][0]=='А')-1.0)<1e-9
assert abs(sum(b[3] for b in BLOCKS60 if b[0][0]=='Б')-1.0)<1e-9
assert abs(sum(b[3] for b in BLOCKS60 if b[0][0]=='В')-1.0)<1e-9

# ---------- Автоматические рекомендации (средний, 15 правил) ----------
ar = find(lambda t: t[0][:2] == ['Условие', 'Автоматическая рекомендация'])
assert len(ar) == 1 and len(ar[0]) == 16, [len(t) for t in ar]
AUTO_REC = [[row[0].replace(' ≤ ', '<= ').strip(), row[1]] for row in ar[0][1:]]
assert len(AUTO_REC) == 15

# ---------- SF-2: примеры критических обстоятельств (5) ----------
i = txt.index('Примеры проверок SF-2:'); seg = txt[i:i+900]
SF2 = [ln.strip().rstrip(';.') for ln in seg.split('\n')[1:] if ln.strip() and ln[0].islower()][:5]
assert len(SF2) == 5, SF2

# ---------- Зрелость M0-M9 ----------
m = find(lambda t: t[0][0] == 'Уровень зрелости')
assert len(m) == 1 and len(m[0]) == 11
M09 = [[r[0], r[1], r[2]] for r in m[0][1:]]

# ---------- Решения R1-R7 ----------
r = find(lambda t: t[0][0] == 'Код' and t[0][1] == 'Решение')
assert len(r) == 1 and len(r[0]) == 8
R17 = [[x[0], x[1], x[2]] for x in r[0][1:]]

# ---------- Шкала 0-5 + НД + НП (средний опросник) ----------
sc = find(lambda t: t[0][0] == 'Балл' and t[0][1] == 'Интерпретация')
assert len(sc) == 1 and len(sc[0]) == 9
SCALE = [[x[0], x[1], x[2]] for x in sc[0][1:]]

# ---------- LCC: 8 этапов ----------
lcc = find(lambda t: t[0][0] == 'Этап' and t[0][1] == 'Состав затрат')
assert len(lcc) == 1 and len(lcc[0]) == 9
LCC = [[x[0], x[1]] for x in lcc[0][1:]]

# ---------- Альтернативы: 9 критериев × 4 сценария ----------
alt = find(lambda t: t[0][0] == 'Критерий' and 'Продолжить разработку' in t[0])
assert len(alt) == 1 and len(alt[0]) == 10
ALTCRIT = [x[0] for x in alt[0][1:]]
ALTNAMES = alt[0][0][1:]  # 4 сценария: Продолжить разработку / Модернизировать / Закупить / Адаптировать
assert len(ALTNAMES) == 4 and len(ALTCRIT) == 9

# ---------- BCR интерпретация ----------
bcr = [t for t in find(lambda t: t[0][0] == 'BCR') if len(t[0]) == 2]
assert len(bcr) == 1 and len(bcr[0]) == 5
BCR = [[x[0], x[1]] for x in bcr[0][1:]]

# ---------- Категории заключения среднего аудита (I-IV, исправленная редакция) ----------
cc = find(lambda t: any(x[0].startswith('I.') for x in t))
assert len(cc) == 1 and len(cc[0]) == 5
CAT30 = [[x[0], x[1], x[2]] for x in cc[0][1:]]

# ---------- Заключения: краткого (21.1) и среднего (21.2) ----------
kb = find(lambda t: t[0] == ['Результат', 'Формулировка заключения'])
assert len(kb) == 2 and all(len(t) == 6 for t in kb)
CONC15 = [[x[0], x[1]] for x in kb[0][1:]]   # краткий: оба 80-100 / 60-79 / один 40-59 / один <40 / нет документов
CONC30 = [[x[0], x[1]] for x in kb[1][1:]]   # средний (21.2): оба 85-100 / 70-84 / один 55-69 / 40-54 / <40

# ---------- Решения полного опросника (21.3) ----------
kf = find(lambda t: t[0] == ['Результат', 'Рекомендуемое решение'])
assert len(kf) == 1 and len(kf[0]) == 8
CONC60 = [[x[0], x[1]] for x in kf[0][1:]]

# ---------- Пороги краткого (17.2) ----------
kp = find(lambda t: t[0] == ['Индекс', 'Оценка', 'Рекомендация'])
assert len(kp) == 1 and len(kp[0]) == 5
POROG15 = [[x[0], x[1], x[2]] for x in kp[0][1:]]

# ---------- Пороги среднего (18.2) ----------
sp = find(lambda t: t[0] == ['Индекс', 'Интерпретация', 'Управленческое решение'])
assert len(sp) == 1 and len(sp[0]) == 6
POROG30 = [[x[0], x[1], x[2]] for x in sp[0][1:]]

# ---------- Уровни итогового заключения (Высокий/Приемлемый/Проблемный/Критический) ----------
lv = find(lambda t: t[0] == ['Уровень', 'Значение', 'Управленческое решение'])
assert len(lv) == 1 and len(lv[0]) == 5
POROG_LVL = [[x[0], x[1], x[2]] for x in lv[0][1:]]

# ---------- Роли комиссии (9), принципы (8) ----------
roles = find(lambda t: t[0][:3] == ['№', 'Участник', 'Основные полномочия'])
assert len(roles) == 1 and len(roles[0]) == 10
ROLES = [[x[1], x[2], x[3]] for x in roles[0][1:]]
pr = find(lambda t: t[0] == ['Принцип', 'Содержание'])
assert len(pr) == 1 and len(pr[0]) == 9
PRINC = [[x[0], x[1]] for x in pr[0][1:]]

# ---------- Категории рекомендаций К1-К4, статусы мероприятий (7) ----------
k14 = find(lambda t: t[0][0] == 'Категория' and t[0][2] == 'Срок реакции')
assert len(k14) == 1 and len(k14[0]) == 5
K14 = [[x[0], x[1], x[2]] for x in k14[0][1:]]
st = find(lambda t: t[0] == ['Статус', 'Значение'])
assert len(st) == 1 and len(st[0]) == 8
STAT = [[x[0], x[1]] for x in st[0][1:]]

# ---------- Раннее предупреждение (7 триггеров) ----------
ew = find(lambda t: t[0] == ['Событие', 'Индикатор', 'Действие'])
assert len(ew) == 1 and len(ew[0]) == 8
EARLY = [[x[0], x[1], x[2]] for x in ew[0][1:]]

# ---------- Контрольные примеры приёмки (8) ----------
ctl = find(lambda t: t[0] == ['№', 'Контрольный пример', 'Ожидаемое поведение'])
assert len(ctl) == 1 and len(ctl[0]) == 9
CTRL = [[x[1], x[2]] for x in ctl[0][1:]]

# ---------- Шкала AHP ----------
ahp = find(lambda t: t[0][0] == 'Значение' and any('Одинаковая значимость' in ' '.join(x) for x in t))
assert len(ahp) == 2 and all(len(t) == 7 for t in ahp)
AHPSCALE = [[x[0], x[1]] for x in ahp[1][1:]]  # исправленная редакция («Интерпретация»)

# ---------- Контрольные точки переоценки (8) ----------
i = txt.index('обязательные контрольные точки переоценки')
seg = txt[i:i+500]
CKPT = [ln.strip().rstrip(';') for ln in seg.split('\n')[1:] if ln.strip() and ln[0].islower()][:8]
assert len(CKPT) == 8, CKPT

# ---------- Шаблон заключения (11.3, 20 пунктов) ----------
i = txt.index('ЗАКЛЮЧЕНИЕ ПО РЕЗУЛЬТАТАМ АУДИТА НИОКР')
TMPL = []
for ln in txt[i:].split('\n')[1:]:
    ln = ln.strip().rstrip('.')
    if not ln or ln.startswith('11.4'): break
    TMPL.append(ln)
assert len(TMPL) == 20, (len(TMPL), TMPL)

data = dict(Q15=Q15, Q30=Q30, ANCH30=ANCH30, Q60=Q60, BLOCKS60=BLOCKS60, AUTO_REC=AUTO_REC,
            SF2=SF2, M09=M09, R17=R17, SCALE=SCALE, LCC=LCC, ALTCRIT=ALTCRIT, ALTNAMES=ALTNAMES,
            BCR=BCR, CAT30=CAT30, CONC15=CONC15, CONC30=CONC30, CONC60=CONC60, POROG15=POROG15,
            POROG30=POROG30, POROG_LVL=POROG_LVL, ROLES=ROLES, PRINC=PRINC, K14=K14, STAT=STAT,
            EARLY=EARLY, CTRL=CTRL, AHPSCALE=AHPSCALE, CKPT=CKPT, TMPL=TMPL)
io.open('scripts/data/data9.json', 'w', encoding='utf-8').write(json.dumps(data, ensure_ascii=False, indent=1))
print("PARSE OK →", len(json.dumps(data, ensure_ascii=False)), "chars JSON")
for k in sorted(data): print(" ", k, len(data[k]))
