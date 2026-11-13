#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Этап 23 (выпуск синхронизации v22.1):
1) десятичный разделитель-запятая во всех генерируемых числах (презентация, сценарий);
2) в презентацию добавлен слайд с таблицей обоснованности ВСЕХ 5 НИОКР стенда
   (синхронно с корневым README и §6.1 сценария);
3) версии выпуска -> v22.1.
Артефакты модели/Excel/методики не изменяются."""
import io, os, re

HERE = os.path.dirname(os.path.abspath(__file__))

RU_HELPER = '''
# --- синхронизация числового формата: десятичная запятая (этап 23) ---
_RU_PROTECT = [
    r'v\\d+(?:\\.\\d+)*',                 # версии: v1.0, v22.1
    r'\\d{1,2}\\.\\d{1,2}\\.\\d{4}',          # даты: 11.11.2026
    r'§\\s*\\d+(?:\\.\\d+)*',              # §16.3
    r'(?m)^#{1,6}\\s*\\d+(?:\\.\\d+)*',    # номера заголовков: ### 6.1.
    r'\\bраздел(?:а|е|ов|ах)?\\s+\\d+(?:\\.\\d+)*',
    r'\\b(?:test|gen|verify|model|demo_stats|qeo|_qa|_patch|_stats)[\\w]*_v\\d+[\\w.]*',
    r'\\b\\d+\\.\\d+(?=\\s*(?:методики|методике|методику))',
]


def ru_num(t):
    """Меняет десятичную точку на запятую, не трогая версии, даты, номера разделов."""
    t = str(t)
    keep = []

    def stash(m):
        keep.append(m.group(0))
        return '\\x00%d\\x00' % (len(keep) - 1)

    for pat in _RU_PROTECT:
        t = re.sub(pat, stash, t)
    t = re.sub(r'(?<![\\w.])\\d+\\.\\d+(?![\\w.])', lambda m: m.group(0).replace('.', ','), t)
    return re.sub('\\x00(\\d+)\\x00', lambda m: keep[int(m.group(1))], t)

'''


def patch(src, dst, pairs, inject=None):
    s = io.open(os.path.join(HERE, src), encoding='utf-8').read()
    log = []
    for old, new, n in pairs:
        c = s.count(old)
        assert c == n, 'ANCHOR %d!=%d: %r' % (c, n, old[:70])
        s = s.replace(old, new)
        log.append(old[:56].replace('\n', ' '))
    if inject:
        anchor, code = inject
        assert s.count(anchor) == 1
        s = s.replace(anchor, anchor + code)
    assert 'import re' in s or 'import json, io, os' in s
    if 'import re' not in s:
        s = s.replace('import json, io, os', 'import json, io, os, re', 1)
    assert not re.search(r'[\u4e00-\u9fff]', s), 'CJK'
    io.open(os.path.join(HERE, dst), 'w', encoding='utf-8').write(s)
    print('OK ->', dst, '|', len(s), 'chars |', len(log), 'замен')
    for l in log:
        print('   ', l)


# ======================= ПРЕЗЕНТАЦИЯ =======================
NEW_SLIDE = '''# ---- 7Б. Обоснованность на учебном стенде (этап 23: синхронизация) ----
s = slide('Обоснованность на учебном стенде', 'Те же 5 НИОКР, что в корневом README и §6.1 сценария апробации · равные веса · базовые пороги')
rows = [['НИОКР', 'n', 'W', 'ДИ I_A (зона)', 'ДИ I_C (зона)', 'Jack', 'Ворот', 'Вердикт']]
for p in ST['just']:
    vd = 'Обосновано' if p['okCount'] == 5 else ('С оговорками' if p['okCount'] >= 3 else 'Доработка')
    ng = ', '.join(x['id'] for x in p['gates'] if not x['ok'])
    rows.append([
        p['shifr'], str(p['n']), '%.4f' % p['W'],
        '%.2f–%.2f (%s, %s)' % (p['ciA']['lo'], p['ciA']['hi'], p['zA'], 'уст' if p['ciA']['stable'] else 'НЕ уст'),
        '%.2f–%.2f (%s, %s)' % (p['ciC']['lo'], p['ciC']['hi'], p['zC'], 'уст' if p['ciC']['stable'] else 'НЕ уст'),
        '%d/%d' % (p['jackChanges'], p['jackN']),
        '%d/5%s' % (p['okCount'], (' (' + ng + ')') if ng else ''),
        vd,
    ])
table(s, rows, top=1.35, size=11.5, col_w=[2.0, 0.45, 0.85, 2.25, 2.25, 0.6, 1.1, 2.6], height=3.1)
bullets(s, [
    'Рекомендации базового правила НЕ меняются ни у одной НИОКР — слой показывает только надёжность оценки.',
    'Типовая оговорка: точечная оценка в зоне В, но нижняя граница ДИ уже в зоне С (ворота G3) — решение принимается с оговоркой либо расширяется группа.',
    'МАТ-ЛИТ-25 (4 эксперта) — единственная НИОКР стенда с 5 из 5 ворот: рост числа экспертов сужает интервалы.',
], top=4.7, size=14, height=2.2)

'''
patch('gen_12_pptx_22.py', 'gen_12_pptx_23.py', [
    ('"""Этап 22 (v22.0), файл 12: презентация пакета ЭО-НИОКР v2.1 для руководства.\n13 слайдов',
     '"""Этап 23 (выпуск синхронизации v22.1), файл 12: презентация пакета ЭО-НИОКР v2.1.\n14 слайдов', 1),
    ("'ЭО-НИОКР v2.1 · пакет v22.0 · 11.11.2026 · слайд %d'",
     "'ЭО-НИОКР v2.1 · пакет v22.1 · 11.11.2026 · слайд %d'", 1),
    ("r3.text = 'Пакет v22.0 ·", "r3.text = 'Пакет v22.1 ·", 1),
    ("['Презентация', '13 слайдов (этот файл)'", "['Презентация', '14 слайдов (этот файл)'", 1),
    ("    'Поставка v22.0: HTML-модель", "    'Поставка v22.1: HTML-модель", 1),
    ('# ---- 7. Калибровка ----', NEW_SLIDE + '# ---- 7. Калибровка ----', 1),
    # числовой формат через обёртки
    ('        r = p.add_run(); r.text = (\'•  \' if lvl == 0 else \'–  \') + txt',
     '        r = p.add_run(); r.text = ru_num((\'•  \' if lvl == 0 else \'–  \') + txt)', 1),
    ('            c = t.cell(i, j); c.text = str(val)',
     '            c = t.cell(i, j); c.text = ru_num(val)', 1),
    ("    p = tf.paragraphs[0]; r = p.add_run(); r.text = title",
     "    p = tf.paragraphs[0]; r = p.add_run(); r.text = ru_num(title)", 1),
    ("        p2 = tf.add_paragraph(); r2 = p2.add_run(); r2.text = sub",
     "        p2 = tf.add_paragraph(); r2 = p2.add_run(); r2.text = ru_num(sub)", 1),
], inject=('N = [0]\n', RU_HELPER))

# ======================= СЦЕНАРИЙ =======================
patch('gen_13_scenario_22.py', 'gen_13_scenario_23.py', [
    ('"""Этап 22 (v22.0), файл 13: сценарий пилотной апробации и калибровки ЭО-НИОКР v2.1.',
     '"""Этап 23 (выпуск синхронизации v22.1), файл 13: сценарий апробации и калибровки ЭО-НИОКР v2.1.', 1),
    ("w('**Пакет v22.0 (этап 22).", "w('**Пакет v22.1 (этапы 22–23).", 1),
    ("*Конец сценария апробации и калибровки (пакет v22.0, 11.11.2026).",
     "*Конец сценария апробации и калибровки (пакет v22.1, 11.11.2026).", 1),
    ("md = '\\n'.join(L)", "L = [ru_num(x) for x in L]\nmd = '\\n'.join(L)", 1),
], inject=('L = []\n', RU_HELPER))
