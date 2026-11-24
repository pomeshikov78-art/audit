# -*- coding: utf-8 -*-
import io
p='scripts/verify_10_xlsx_22.py';s=io.open(p,encoding='utf-8').read()
def rep(a,b,n=1):
 global s
 c=s.count(a);assert c==n,(c,n,a[:80]);s=s.replace(a,b)
rep('10_Чек-лист_ЭО_НИОКР_22_вопроса_v22.xlsx','10_Чек-лист_ЭО_НИОКР_22_вопроса_v23.xlsx')
rep("TMP = '/tmp/eo22_case.xlsx'","TMP = '/tmp/eo23_case.xlsx'")
rep('eo22_case.xlsx','eo23_case.xlsx',1)
rep("KV_DIR = ['neg', 'neg', 'neg', 'neg', 'neg', 'neg', 'pos', 'pos', 'neg']","KV_DIR = ['risk_yes'] * 9  # v23: единая шкала, риск = 1–2 для всех 9")
rep("SF = sf or ['Нет'] * 5","SF = sf or [5] * 5  # v23: числовая шкала, 5 = «Нет»")
rep("KV = kv or [(5 if d == 'neg' else 1) for d in KV_DIR]","KV = kv or [5] * 9  # v23: 5 = «Нет» = нет сигнала")
rep("sf=['Да', 'Нет', 'Нет', 'Нет', 'Нет']","sf=[1, 5, 5, 5, 5]",2)
rep("счётчик «Да» СФ1 = 1","счётчик «выявлен» СФ1 = 1 (ответы 1–2)")
rep("К-6: «красные» групповые сигналы (КВ1 neg ≤2 у двоих; КВ7 pos ≥4 у двоих)","К-6: «красные» групповые сигналы (единая шкала: КВ1 и КВ7 ≤2 у двоих)")
rep("kv2 = [1, 5, 5, 5, 5, 5, 4, 5, 5]","kv2 = [1, 5, 5, 5, 5, 5, 2, 5, 5]")
rep("КВ7: 2 сигнала (позитивный)","КВ7: 2 сигнала (единая шкала, 1–2)")
# static blank block
a=s.index("wsa = wbs['Анкета_бланк']");b=s.index("wsk = wbs['Калибровка']",a)
s=s[:a]+"""wsa = wbs['Анкета_бланк']
bl = [wsa.cell(row=rr, column=1).value for rr in range(1, wsa.max_row + 1)]
import re
bl_codes = [v for v in bl if isinstance(v, int)]
eq('бланк: 22 основных + 5 СФ + 9 КВ = 36 строк ответов', len(bl_codes), 36)
ck('бланк: заголовок Приложение №2', any(isinstance(v, str) and 'Приложение №2' in v for row in wsa.iter_rows() for v in [c.value for c in row]))
ck('бланк: без подписи эксперта', not any(isinstance(v, str) and 'Подпись эксперта' in v for row in wsa.iter_rows() for v in [c.value for c in row]))
ck('бланк: без блока Р', not any(isinstance(v, str) and 'БЛОК РЕЗУЛЬТАТИВНОСТИ' in v for row in wsa.iter_rows() for v in [c.value for c in row]))
eq('бланк: область печати A1:H59', str(wsa.print_area), "'Анкета_бланк'!$A$1:$H$59")
eq('бланк: разрыв после строки 27', [b.id for b in wsa.row_breaks.brk], [27])
# числовая валидация СФ на листах экспертов
wse1 = wbs['Э1']
sf_dv = [dv for dv in wse1.data_validations.dataValidation if 'C45' in str(dv.sqref)]
ck('Э1: валидация СФ C45:C49 найдена', len(sf_dv) >= 1)
if sf_dv:
    eq('Э1: СФ шкала 1–5', sf_dv[0].formula1, '"1,2,3,4,5"')
has('Э1: КВ7 риск при ≤2', wse1['E59'].value, '<=2')
has('свод: СФ1 считает ответы 1 и 2', wbs['Свод']['C48'].value, 'COUNTIF')
"""+s[b:]
io.open('scripts/verify_10_xlsx_23.py','w',encoding='utf-8').write(s)
print('verify_10_xlsx_23.py',len(s))
