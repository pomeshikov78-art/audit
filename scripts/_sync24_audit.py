#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Этап 23 (v23.0): аудит синхронности информации пакета v23.0.
Только стандартная библиотека: факты извлекаются из OOXML (zip+XML), HTML и MD,
затем сверяются с утверждениями во всех документах поставки."""
import io, os, re, json, zipfile, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, '..')
OUT = os.path.join(ROOT, 'output')
R = []          # отчёт
BAD = []        # расхождения


def rd(p):
    return io.open(p, encoding='utf-8').read()


def fact(name, val):
    R.append(('ФАКТ', name, val))
    return val


def check(doc, claim, ok, detail=''):
    R.append(('OK ' if ok else 'РАСХОЖДЕНИЕ', '%s: %s' % (doc, claim), detail))
    if not ok:
        BAD.append('%s: %s %s' % (doc, claim, detail))


# ---------- ФАКТЫ ИЗ АРТЕФАКТОВ ----------
html_p = os.path.join(OUT, 'Аудит_НИОКР_интерактивная_модель.html')
html = rd(html_p)
F = {}
F['html_bytes'] = fact('HTML размер, Б', os.path.getsize(html_p))
F['html_ver'] = fact('HTML version', (re.search(r'version:\s*"([\d.]+)"', html) or [None, '?'])[1])
F['metodika_ver'] = fact('HTML редакция методики', 'v2.2' if 'ЭО-НИОКР v2.2' in html else 'НЕТ v2.2')
snap = os.path.join(OUT, 'versions', 'Аудит_НИОКР_модель_v23.0.html')
F['snap_eq'] = fact('HTML == снимок v23.0', rd(snap) == html)

# CSV портфеля: число колонок в заголовке
m = re.search(r'portCSV[\s\S]{0,4000}?', html)
hdr = re.findall(r'\[[^\[\]]*?Обоснованность[^\[\]]*?\]', html)
cols = None
for h in hdr:
    parts = [x for x in re.split(r',', h) if x.strip()]
    if 12 <= len(parts) <= 25:
        cols = len(parts)
        break
F['csv_cols'] = fact('CSV портфеля колонок (по заголовку с «Обоснованность»)', cols)

# XLSX
xp = os.path.join(OUT, '10_Чек-лист_ЭО_НИОКР_22_вопроса_v23.xlsx')
with zipfile.ZipFile(xp) as z:
    wb = z.read('xl/workbook.xml').decode('utf-8')
    sheets = re.findall(r'<sheet\b[^>]*?\bname="([^"]+)"', wb)
F['xlsx_sheets'] = fact('Excel листов', len(sheets))
F['xlsx_names'] = fact('Excel состав', ', '.join(sheets))
F['xlsx_bytes'] = fact('Excel размер, Б', os.path.getsize(xp))

# PPTX
pp = os.path.join(OUT, '12_Презентация_ЭО_НИОКР_v23.pptx')
with zipfile.ZipFile(pp) as z:
    slides = sorted([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)],
                    key=lambda n: int(re.search(r'(\d+)', n).group(1)))
    texts = []
    for n in slides:
        xml = z.read(n).decode('utf-8')
        texts.append(' '.join(re.findall(r'<a:t>([^<]*)</a:t>', xml)))
F['pptx_slides'] = fact('Презентация слайдов', len(slides))
F['pptx_bytes'] = fact('Презентация размер, Б', os.path.getsize(pp))
ppt_all = '\n'.join(texts)

# Методика MD
mp = os.path.join(OUT, '11_Методика_ЭО_НИОКР_22_вопроса_v23.md')
met = rd(mp)
h2 = re.findall(r'^## (.+)$', met, re.M)
secs = [h for h in h2 if re.match(r'^\s*(РАЗДЕЛ|Раздел)?\s*\d+\.', h) or re.match(r'^\d+', h)]
apps = sorted(set(int(x) for x in re.findall(r'Приложение №\s*(\d+)', met)))
F['met_h2'] = fact('Методика: заголовков ##', len(h2))
F['met_apps'] = fact('Методика: номера приложений', apps)
F['met_bytes'] = fact('Методика MD размер, Б', os.path.getsize(mp))
s21 = met[met.find('## 21.'):] if '## 21.' in met else ''
F['met_s21_subs'] = fact('Методика §21 подразделов (### 21.x)', len(re.findall(r'^### 21\.\d+', met, re.M)))
F['met_sections'] = fact('Методика: разделов (## N.)', len(re.findall(r'^## \d+\.', met, re.M)))

# Демо-статистика
ST = json.load(io.open(os.path.join(HERE, 'data', 'demo_stats_v23.json'), encoding='utf-8'))
F['demo_n'] = fact('Демо: проектов', len(ST['just']))

# ---------- ДОКУМЕНТЫ ----------
DOCS = {
    'README.md (корень)': rd(os.path.join(ROOT, 'README.md')),
    'README_итоговый_пакет.md': rd(os.path.join(OUT, 'README_итоговый_пакет.md')),
    'VERSIONS.md': rd(os.path.join(OUT, 'VERSIONS.md')),
    'Инструкция v23': rd(os.path.join(OUT, '6_Инструкция_пользователя_модели_v23.md')),
    'Сценарий v23': rd(os.path.join(OUT, '13_Сценарий_апробации_и_калибровки_v23.md')),
    'Методика v23': met,
    'Презентация v23 (текст)': ppt_all,
}
# текущая часть итогового README (до архива v21.1)
cur = DOCS['README_итоговый_пакет.md'].split('# Архив редакции v21.1')[0]

CLAIMS = [
    ('19 листов', lambda t: '19 лист' in t),
    ('14 слайдов', lambda t: '14 слайд' in t),
    ('354 проверок', lambda t: '354' in t),
    ('345 проверок', lambda t: '345' in t),
    ('28/28 зонд', lambda t: '28/28' in t or '28 / 28' in t),
    ('бандл v19-v23', lambda t: 'audit_v19-v23.bundle' in t),
]
for name in ['README.md (корень)', 'README_итоговый_пакет.md', 'Инструкция v23']:
    t = cur if name.startswith('README_итоговый') else DOCS[name]
    for cl, fn in CLAIMS:
        check(name, cl, fn(t))

# устаревшие ссылки в ТЕКУЩИХ частях
STALE = ['audit_v19-v21.bundle', '18 листов', '243 провер', '180 провер', '12 слайдов',
         '20 разделов + приложения №1–12', 'текущая версия — v21', 'ТЕКУЩАЯ ВЕРСИЯ — v21']
for name in ['README.md (корень)', 'Инструкция v23', 'Сценарий v23']:
    t = DOCS[name]
    for st in STALE:
        check(name, 'нет устаревшего «%s»' % st, st.lower() not in t.lower())
for st in STALE:
    check('README_итоговый (текущая часть)', 'нет устаревшего «%s»' % st, st.lower() not in cur.lower())

# факты против утверждений
check('Все документы', 'Excel листов = 19 (факт)', F['xlsx_sheets'] == 19, str(F['xlsx_sheets']))
check('Все документы', 'слайдов = 14 (факт)', F['pptx_slides'] == 14, str(F['pptx_slides']))
check('Методика', 'приложений 1..13 (факт)', apps_ok := (F['met_apps'] == list(range(1, 14))), str(F['met_apps']))
check('Методика', '§21 подразделов = 8 (факт)', F['met_s21_subs'] == 8, str(F['met_s21_subs']))
check('Методика', 'разделов = 22 (факт)', F['met_sections'] == 22, str(F['met_sections']))
check('HTML', 'version = 23.0', F['html_ver'] == '23.0', F['html_ver'])
check('HTML', 'редакция методики v2.2', F['metodika_ver'] == 'v2.2')
check('HTML', 'снимок идентичен', F['snap_eq'] is True)
check('HTML', 'CSV колонок = 18 (факт)', F['csv_cols'] == 18, str(F['csv_cols']))

# числа стенда: README / сценарий / презентация против JSON
def num(x):
    return ('%.4f' % x).rstrip('0').rstrip('.')

for p in ST['just']:
    sh = p['shifr']
    w4 = '%.4f' % p['W']
    w4c = w4.replace('.', ',')
    for name in ['README.md (корень)', 'Сценарий v23', 'Презентация v23 (текст)']:
        t = DOCS[name]
        if sh not in t:
            check(name, 'шифр %s упомянут' % sh, False)
            continue
        check(name, 'W %s = %s' % (sh, w4c), w4c in t)
        check(name, 'I_A/I_C %s' % sh,
              ('%.2f' % p['IA']).replace('.', ',') in t and ('%.2f' % p['IC']).replace('.', ',') in t)
    # вердикт
    for name in ['README.md (корень)', 'Сценарий v23']:
        check(name, 'вердикт %s' % sh, p['verdict'] in DOCS[name])

# ссылки на файлы в текущей поставке существуют
missing = []
for fn in re.findall(r'`([^`]+\.(?:html|xlsx|md|docx|pptx|bundle|json|py|js))`', cur):
    base = os.path.basename(fn)
    # git-бандл — транспортный артефакт публикации: он сознательно не отслеживается git
    # (иначе репозиторий содержал бы сам себя), поэтому в клоне его отсутствие нормально.
    if base.endswith('.bundle'):
        continue
    if base.startswith(('gen_', 'verify_', 'test_', '_qa', '_sync', '_patch', '_stats', 'model_', 'qeo', 'org_', 'demo_', 'audit_')):
        cand = [os.path.join(HERE, base), os.path.join(HERE, 'data', base),
                os.path.join(HERE, 'tests', base), os.path.join(OUT, base)]
    else:
        cand = [os.path.join(OUT, base), os.path.join(ROOT, base),
                os.path.join(OUT, 'versions', base)]
    if not any(os.path.exists(c) for c in cand):
        missing.append(base)
check('README_итоговый (текущая часть)', 'все упомянутые файлы существуют', not missing, ', '.join(sorted(set(missing))))

# корневой README: журнал итераций и правило обновления
root = DOCS['README.md (корень)']
check('README.md (корень)', 'журнал итераций 19→23', all(k in root for k in ['Этап 19', 'Этап 20', 'Этап 21', 'Этап 22', 'Этап 23']))
check('README.md (корень)', 'правило обновления после каждой итерации', 'после каждой итерации' in root)
check('README.md (корень)', 'адрес репозитория задачи', 'https://github.com/pomeshikov78-art/audit' in root)

# выпуск синхронизации v22.1
check('README.md (корень)', 'выпуск v23.0', 'v23.0' in root)
check('README_итоговый (текущая часть)', 'выпуск v23.0', 'v23.0' in cur)
check('VERSIONS.md', 'пункт 33 (этап 23)', '\n33. (этап 23' in DOCS['VERSIONS.md'])
check('Инструкция v23', 'выпуск v23.0', 'v23.0' in DOCS['Инструкция v23'])
for name in ['Сценарий v23', 'Презентация v23 (текст)']:
    t = DOCS[name]
    # номера разделов, версии и даты обязаны сохранять точку — исключаем их
    prot = [r'v\d+(?:\.\d+)*', r'\d{1,2}\.\d{1,2}\.\d{4}', r'§\s*\d+(?:\.\d+)*',
            r'(?m)^#{1,6}\s*\d+(?:\.\d+)*', r'\bраздел(?:а|е|ов|ах)?\s+\d+(?:\.\d+)*',
            r'\b\d+\.\d+(?=\s*(?:методики|методике|методику))']
    tt = t
    for pat in prot:
        tt = re.sub(pat, ' ', tt)
    bad = [x for x in re.findall(r'(?<![\w.])\d+\.\d+(?![\w.])', tt)]
    check(name, 'десятичная запятая (нет чисел с точкой)', not bad, ', '.join(bad[:6]))

# CJK
for name, t in DOCS.items():
    check(name, 'без CJK', not re.search(r'[\u4e00-\u9fff]', t))


# ---------- единая форма организации v23 ----------
ORG = json.load(io.open(os.path.join(HERE, 'data', 'org_form_v23.json'), encoding='utf-8'))
F['org_main'] = fact('Единая форма: основных вопросов', len(ORG['main']))
F['org_sf'] = fact('Единая форма: СФ', len(ORG['sf']))
F['org_kv'] = fact('Единая форма: КВ', len(ORG['kv']))
check('Единая форма', '22 основных = 10 А + 12 Ц', F['org_main'] == 22)
check('Единая форма', '5 СФ + 9 КВ', F['org_sf'] == 5 and F['org_kv'] == 9)
check('HTML', 'печатная форма Приложение №2', 'Приложение №2' in html)
check('HTML', 'нет печатного блока Р в единой форме', 'БЛОК РЕЗУЛЬТАТИВНОСТИ' not in html.split('function ankFormHTML')[1].split('function printTab')[0])
for fn in ['14_Единая_форма_анкеты_Приложение_2_v23.pdf', '14_Единая_форма_анкеты_Приложение_2_v23.docx']:
    check('Поставка', 'файл единой формы существует: ' + fn, os.path.exists(os.path.join(OUT, fn)))
for name in ['README.md (корень)', 'README_итоговый_пакет.md', 'Инструкция v23', 'Методика v23']:
    t = cur if name.startswith('README_итоговый') else DOCS[name]
    check(name, 'упомянута единая форма', 'единая форма' in t.lower() or 'Приложение №2' in t)

# ---------- ВЫВОД ----------
print('=' * 78)
for k, n, v in R:
    if k == 'ФАКТ':
        print('  %-52s %s' % (n, v))
print('-' * 78)
nok = sum(1 for k, _, _ in R if k == 'OK ')
for k, n, v in R:
    if k == 'РАСХОЖДЕНИЕ':
        print('  ✗ %-60s %s' % (n, v))
print('=' * 78)
print('ПРОВЕРОК: %d | OK: %d | РАСХОЖДЕНИЙ: %d' % (nok + len(BAD), nok, len(BAD)))
sys.exit(1 if BAD else 0)
