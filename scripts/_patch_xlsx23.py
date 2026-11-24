#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Этап 24: Excel v23 — единая двухстраничная форма Приложения №2, СФ 1–5."""
import io,os,re
H=os.path.dirname(os.path.abspath(__file__))
t=io.open(os.path.join(H,'gen_10_xlsx_22.py'),encoding='utf-8').read()
def rep(a,b,n=1):
 global t
 c=t.count(a);assert c==n,(c,n,a[:90]);t=t.replace(a,b)
rep("'data', 'qeo22.json'","'data', 'qeo_v23.json'")
rep("D = json.load(io.open(os.path.join(HERE, 'data', 'qeo_v23.json'), encoding='utf-8'))","D = json.load(io.open(os.path.join(HERE, 'data', 'qeo_v23.json'), encoding='utf-8'))['base']\nORG = json.load(io.open(os.path.join(HERE, 'data', 'org_form_v23.json'), encoding='utf-8'))")
rep("'data', 'qeo_v21.json'","'data', 'qeo_v23.json'")
rep('10_Чек-лист_ЭО_НИОКР_22_вопроса_v22.xlsx','10_Чек-лист_ЭО_НИОКР_22_вопроса_v23.xlsx')
rep('ЭО-НИОКР v2.1','ЭО-НИОКР v2.2',t.count('ЭО-НИОКР v2.1'))
rep('v22.0','v23.0',t.count('v22.0'))
rep('модель v22.0','модель v23.0',t.count('модель v22.0'))
rep('ответ: Да / Нет / Неопределённо','ответ 1–5 по форме организации: 1=да, 2=возможно да, 3=неопределённо, 4=возможно нет, 5=нет')
rep('ответ Да / Нет / Неопределённо','ответ 1–5 (Приложение №2)')
rep('formula1=\'"Да,Нет,Неопределённо"\'','formula1=\'"1,2,3,4,5"\'')
rep("error='Допустимо: Да, Нет, Неопределённо'","error='Допустимо: 1, 2, 3, 4, 5'")
rep("'=' + '+'.join('COUNTIF(Э%d!C%d,\"Да\")' % (e, er) for e in range(1, NEXP + 1))", "'=' + '+'.join('COUNTIF(Э%d!C%d,1)+COUNTIF(Э%d!C%d,2)' % (e, er, e, er) for e in range(1, NEXP + 1))")
rep("'=' + '+'.join('COUNTIF(Э%d!C%d,\"Неопределённо\")' % (e, er) for e in range(1, NEXP + 1))", "'=' + '+'.join('COUNTIF(Э%d!C%d,3)' % (e, er) for e in range(1, NEXP + 1))")
rep("'=' + '+'.join('COUNTIF(Э%d!C%d,\"Нет\")' % (e, er) for e in range(1, NEXP + 1))", "'=' + '+'.join('COUNTIF(Э%d!C%d,4)+COUNTIF(Э%d!C%d,5)' % (e, er, e, er) for e in range(1, NEXP + 1))")
rep('value=\'=IF(C%d=\"\",\"\",IF(D%d=\"Негативный\",IF(C%d<=2,\"СИГНАЛ\",\"\"),IF(C%d>=4,\"СИГНАЛ\",\"\")))\' % (rr, rr, rr, rr)', 'value=\'=IF(C%d=\"\",\"\",IF(C%d<=2,\"СИГНАЛ\",IF(C%d=3,\"УТОЧНИТЬ\",\"\")))\' % (rr, rr, rr)')
rep("value='Негативный' if k['dir'] == 'neg' else 'Позитивный'", "value='1–2 = риск; 3 = уточнить'",2)
rep("value=('негативный: сигнал = балл 1–2' if k['dir'] == 'neg' else 'позитивный: сигнал = балл 4–5')", "value='1–2 — сигнал риска; 3 — уточнить; 4–5 — нет сигнала'")
rep('Фактор выявлен, если хотя бы один эксперт ответил «Да»; неопределён, если «Да» нет, но есть «Неопределённо».', 'Фактор выявлен при ответе 1–2 хотя бы у одного эксперта; ответ 3 требует уточнения комиссией; 4–5 — не выявлен. Пустой ответ не означает «нет».')
rep('Бумажная процедура (лист «Анкета_бланк»): печатная форма индивидуальной анкеты — приложение №1 к методике;', 'Бумажная процедура (лист «Анкета_бланк»): единая форма организации — Приложение №2, две страницы А4, исходные тексты скана;')
rep('документальное подтверждение баллов НЕ требуется.', 'Печатная форма воспроизводит Приложение №2 организации, включая исходные формулировки. Правила интерпретации ответов и границы экспертного суждения — раздел 22 методики ЭО-НИОКР v2.2.')
# Замена печатного листа целиком.
i=t.index("# ============================ ЛИСТ «Анкета_бланк» (v22) ============================")
j=t.index('# защита расчётных листов',i)
FORM=r'''# ============================ ЛИСТ «Анкета_бланк» (v23) ============================
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import Font as XFont, Border as XBorder, Side as XSide, Alignment as XAlign
ws = wb.create_sheet('Анкета_бланк')
# Воспроизведение скана: 8 колонок, две страницы А4; без дополнительных полей.
widths=[4.5,25,51,10,10,10,10,10]
for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
OB=XBorder(*(XSide(style='thin',color='000000'),)*4)
OF=XFont(name='Times New Roman',size=9)
OFC=XAlign(horizontal='center',vertical='center',wrap_text=True)
OFL=XAlign(horizontal='left',vertical='center',wrap_text=True)
def oc(r,c,v=None,bold=False,align='left',size=9):
    x=ws.cell(r,c,v);x.font=XFont(name='Times New Roman',size=size,bold=bold);x.alignment=OFC if align=='center' else OFL;x.border=OB;return x
def om(r,c1,c2,v,bold=False,align='center',size=9):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    for c in range(c1,c2+1):oc(r,c,bold=bold,align=align,size=size)
    oc(r,c1,v,bold,align,size)
def oheader(r,sec):
    om(r,1,8,sec['title'],True,'center',9.5);ws.row_dimensions[r].height=27;r+=1
    oc(r,1,'№',True,'center');oc(r,2,'Критерий',True,'center');oc(r,3,'Вопрос для экспертизы',True,'center');om(r,4,8,'Ответ',True);ws.row_dimensions[r].height=17;r+=1
    om(r,1,3,'',False);oc(r,4,sec['low'],True,'center',7);oc(r,8,sec['high'],True,'center',7);ws.row_dimensions[r].height=25;r+=1
    om(r,1,3,'',False)
    for c in range(4,9):oc(r,c,c-3,True,'center')
    ws.row_dimensions[r].height=14;return r+1
def omain(r,items,sec,height=32):
    r=oheader(r,sec)
    for q in items:
        oc(r,1,q['n'],False,'center');oc(r,2,q['criterion'],True);oc(r,3,q['question'])
        for c in range(4,9):oc(r,c)
        ws.row_dimensions[r].height=height;r+=1
    return r
ws['G1']='Приложение №2';ws['G1'].font=XFont(name='Times New Roman',size=11,bold=True);ws.merge_cells('G1:H1');ws['G1'].alignment=XAlign(horizontal='right')
ws['A2']='ОКР шифр "________________________________"';ws.merge_cells('A2:H2');ws['A2'].font=XFont(name='Times New Roman',size=11,bold=True);ws['A2'].alignment=OFC
ws.row_dimensions[1].height=18;ws.row_dimensions[2].height=24;ws.row_dimensions[3].height=8
r=omain(4,ORG['main'][:10],ORG['sections'][0],34)
r=omain(r,ORG['main'][10:16],ORG['sections'][1],34)
PAGE1_END=r-1
r=omain(r,ORG['main'][16:],ORG['sections'][1],29)
om(r,1,8,'СТОП-ФАКТОРЫ',True);r+=1
oc(r,1,'№',True,'center');om(r,2,3,'Стоп-фактор',True);om(r,4,8,'Ответ',True);r+=1
om(r,1,3,'');oc(r,4,'да',True,'center',7);oc(r,8,'нет',True,'center',7);ws.row_dimensions[r].height=24;r+=1
om(r,1,3,'')
for c in range(4,9):oc(r,c,c-3,True,'center')
r+=1
for n,s in enumerate(ORG['sf'],1):
    oc(r,1,n,False,'center');om(r,2,3,s,False,'left');
    for c in range(4,9):oc(r,c)
    ws.row_dimensions[r].height=30;r+=1
om(r,1,8,'«КРАСНЫЕ» ВОПРОСЫ',True);r+=1
oc(r,1,'№',True,'center');om(r,2,3,'Вопрос',True);om(r,4,8,'Ответ',True);r+=1
om(r,1,3,'');oc(r,4,'да',True,'center',7);oc(r,8,'нет',True,'center',7);ws.row_dimensions[r].height=24;r+=1
om(r,1,3,'')
for c in range(4,9):oc(r,c,c-3,True,'center')
r+=1
for n,s in enumerate(ORG['kv'],1):
    oc(r,1,n,False,'center');om(r,2,3,s,False,'left')
    for c in range(4,9):oc(r,c)
    ws.row_dimensions[r].height=26;r+=1
ws.row_breaks.append(Break(id=PAGE1_END))
ws.page_setup.orientation='portrait';ws.page_setup.paperSize=ws.PAPERSIZE_A4
ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=2
ws.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
ws.print_options.horizontalCentered=True
ws.page_margins.left=0.35;ws.page_margins.right=0.35;ws.page_margins.top=0.35;ws.page_margins.bottom=0.35
ws.print_area='A1:H%d'%(r-1);ws.print_title_rows=None
ws.sheet_view.showGridLines=False

'''
t=t[:i]+FORM+t[j:]
assert not re.search(r'[\u4e00-\u9fff]',t)
io.open(os.path.join(H,'gen_10_xlsx_23.py'),'w',encoding='utf-8').write(t)
print('OK gen_10_xlsx_23.py',len(t))
