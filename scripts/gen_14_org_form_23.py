#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Этап 24 (v23.0): единая форма организации по скану Приложения №2.
Два автономных бланка: PDF (ровно 2 страницы А4) и редактируемый DOCX.
Тексты и шкалы — scripts/data/org_form_v23.json; исходный скан сохранён отдельно."""
import os, io, json
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from xml.sax.saxutils import escape
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

H=os.path.dirname(os.path.abspath(__file__)); O=os.path.join(H,'..','output')
D=json.load(io.open(os.path.join(H,'data','org_form_v23.json'),encoding='utf-8'))
PDF=os.path.join(O,'14_Единая_форма_анкеты_Приложение_2_v23.pdf')
DOCX=os.path.join(O,'14_Единая_форма_анкеты_Приложение_2_v23.docx')
import matplotlib
fontdir=os.path.join(os.path.dirname(matplotlib.__file__),'mpl-data','fonts','ttf')
pdfmetrics.registerFont(TTFont('DejaVu',os.path.join(fontdir,'DejaVuSerif.ttf')))
pdfmetrics.registerFont(TTFont('DejaVu-Bold',os.path.join(fontdir,'DejaVuSerif-Bold.ttf')))
pdfmetrics.registerFontFamily('DejaVu',normal='DejaVu',bold='DejaVu-Bold')
W,Hh=A4; M=12*mm; AW=W-2*M
CW=[7*mm,34*mm,AW-7*mm-34*mm-5*12*mm]+[12*mm]*5
st=ParagraphStyle('cell',fontName='DejaVu',fontSize=7.7,leading=9.1,spaceAfter=0)
stb=ParagraphStyle('bold',parent=st,fontName='DejaVu-Bold')
stc=ParagraphStyle('center',parent=stb,alignment=1)
stsmall=ParagraphStyle('small',parent=stc,fontSize=6.3,leading=7.2)

def P(t,style=st): return Paragraph(escape(str(t)),style)
def row(cells,heights=None): return cells

def header(sec):
    return [
        [P(sec['title'],stc),'','','','','','',''],
        [P('№',stc),P('Критерий',stc),P('Вопрос для экспертизы',stc),P('Ответ',stc),'','','',''],
        ['','','',P(sec['low'],stsmall),'','','',P(sec['high'],stsmall)],
        ['','','']+[P(str(i),stc) for i in range(1,6)],
    ]

def mainrows(items):
    return [[P(q['n'],stc),P(q['criterion'],stb),P(q['question'])]+['']*5 for q in items]

def special(title,label,items):
    rows=[[P(title,stc),'','','','','','',''],
          [P('№',stc),P(label,stc),'',P('Ответ',stc),'','','',''],
          ['','','',P('да',stsmall),'','','',P('нет',stsmall)],
          ['','','']+[P(str(i),stc) for i in range(1,6)]]
    for n,t in enumerate(items,1): rows.append([P(n,stc),P(t),'']+['']*5)
    return rows

def spans(rows):
    cmds=[]
    for i,r in enumerate(rows):
        txt=r[0].text if isinstance(r[0],Paragraph) else str(r[0])
        if txt.startswith(('НАПРАВЛЕНИЕ','СТОП-ФАКТОРЫ','«КРАСНЫЕ»')): cmds.append(('SPAN',(0,i),(7,i)))
        elif txt=='№':
            if r[1].text in ('Стоп-фактор','Вопрос'): cmds.append(('SPAN',(1,i),(2,i)))
            cmds.append(('SPAN',(3,i),(7,i)))
        elif r[1]=='' and r[2]=='' and isinstance(r[3],Paragraph) and r[3].text in ('да',D['sections'][0]['low'],D['sections'][1]['low']): cmds.append(('SPAN',(0,i),(2,i)))
        elif r[1]=='' and r[2]=='': cmds.append(('SPAN',(0,i),(2,i)))
        elif isinstance(r[1],Paragraph) and r[2]=='' and r[3]=='': cmds.append(('SPAN',(1,i),(2,i)))
    return cmds

def table(rows):
    t=Table(rows,colWidths=CW,hAlign='CENTER')
    cmds=[('GRID',(0,0),(-1,-1),0.65,colors.black),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
          ('LEFTPADDING',(0,0),(-1,-1),2.2),('RIGHTPADDING',(0,0),(-1,-1),2.2),
          ('TOPPADDING',(0,0),(-1,-1),1.8),('BOTTOMPADDING',(0,0),(-1,-1),1.8)]+spans(rows)
    t.setStyle(TableStyle(cmds));return t

from reportlab.pdfgen import canvas
c=canvas.Canvas(PDF,pagesize=A4,pageCompression=1)
c.setTitle('Приложение №2 — единая форма анкеты экспертной оценки НИОКР')
c.setAuthor('Организация; редактируемая версия по исходному скану')
def drawpage(rows,first=False):
    y=Hh-M
    if first:
        c.setFont('DejaVu-Bold',10);c.drawRightString(W-M,y,'Приложение №2');y-=10*mm
        c.drawCentredString(W/2,y,'ОКР шифр "________________________________"');y-=8*mm
    t=table(rows);tw,th=t.wrap(AW,Hh)
    assert th<=y-M+1,(first,th,y-M)
    t.drawOn(c,M,y-th)
    c.showPage()
p1=header(D['sections'][0])+mainrows(D['main'][:10])+header(D['sections'][1])+mainrows(D['main'][10:16])
p2=header(D['sections'][1])+mainrows(D['main'][16:])+special('СТОП-ФАКТОРЫ','Стоп-фактор',D['sf'])+special('«КРАСНЫЕ» ВОПРОСЫ','Вопрос',D['kv'])
drawpage(p1,True);drawpage(p2);c.save()

# DOCX: тот же состав, явный разрыв страницы после Ц6.
doc=Document();sec=doc.sections[0];sec.page_width=Cm(21);sec.page_height=Cm(29.7)
sec.top_margin=Cm(1.1);sec.bottom_margin=Cm(1.1);sec.left_margin=Cm(1.2);sec.right_margin=Cm(1.2)
sec.header_distance=Cm(0.5);sec.footer_distance=Cm(0.5)
normal=doc.styles['Normal'];normal.font.name='Times New Roman';normal.font.size=Pt(8.5)
normal.paragraph_format.space_after=Pt(0);normal.paragraph_format.space_before=Pt(0)

def dpar(text,bold=False,align='left',size=8.5):
    p=doc.add_paragraph();p.alignment={'left':WD_ALIGN_PARAGRAPH.LEFT,'center':WD_ALIGN_PARAGRAPH.CENTER,'right':WD_ALIGN_PARAGRAPH.RIGHT}[align]
    r=p.add_run(text);r.bold=bold;r.font.size=Pt(size);r.font.name='Times New Roman';return p

def setcell(cell,text,bold=False,align='left',size=8.2):
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p=cell.paragraphs[0];p.alignment={'left':WD_ALIGN_PARAGRAPH.LEFT,'center':WD_ALIGN_PARAGRAPH.CENTER}[align]
    p.paragraph_format.space_after=Pt(0);p.paragraph_format.space_before=Pt(0)
    r=p.add_run(str(text));r.bold=bold;r.font.size=Pt(size);r.font.name='Times New Roman'
    tcPr=cell._tc.get_or_add_tcPr();mar=OxmlElement('w:tcMar')
    for side in ('top','left','bottom','right'):
        x=OxmlElement('w:'+side);x.set(qn('w:w'),'35');x.set(qn('w:type'),'dxa');mar.append(x)
    tcPr.append(mar)

def dtable(rows):
    t=doc.add_table(rows=len(rows),cols=8);t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    tblPr=t._tbl.tblPr;borders=OxmlElement('w:tblBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        x=OxmlElement('w:'+side);x.set(qn('w:val'),'single');x.set(qn('w:sz'),'6');x.set(qn('w:color'),'000000');borders.append(x)
    tblPr.append(borders)
    for i,w in enumerate(CW): t.columns[i].width=Cm(w/mm*10/10)  # pt->cm ниже
    # reportlab widths в pt; 1 pt = 0.0352778 cm
    for i,w in enumerate(CW): t.columns[i].width=Cm(w*0.0352778)
    for ri,r in enumerate(rows):
        texts=[x.text if isinstance(x,Paragraph) else str(x) for x in r]
        bolds=[isinstance(x,Paragraph) and x.style.name in ('bold','center','small') for x in r]
        aligns=['center' if isinstance(x,Paragraph) and x.style.name in ('center','small') else 'left' for x in r]
        for ci in range(8): setcell(t.cell(ri,ci),texts[ci],bolds[ci],aligns[ci],7.2 if isinstance(r[ci],Paragraph) and r[ci].style.name=='small' else 8.2)
        t.rows[ri].height=Cm(0.42)
    for i,r in enumerate(rows):
        txt=r[0].text if isinstance(r[0],Paragraph) else str(r[0])
        if txt.startswith(('НАПРАВЛЕНИЕ','СТОП-ФАКТОРЫ','«КРАСНЫЕ»')): t.cell(i,0).merge(t.cell(i,7))
        elif txt=='№':
            if r[1].text in ('Стоп-фактор','Вопрос'): t.cell(i,1).merge(t.cell(i,2))
            t.cell(i,3).merge(t.cell(i,7))
        elif r[1]=='' and r[2]=='':
            t.cell(i,0).merge(t.cell(i,2))
            if isinstance(r[3],Paragraph) and r[3].text in ('да',D['sections'][0]['low'],D['sections'][1]['low']): pass
        elif isinstance(r[1],Paragraph) and r[2]=='' and r[3]=='': t.cell(i,1).merge(t.cell(i,2))
    return t

dpar('Приложение №2',True,'right',10.5)
dpar('ОКР шифр "________________________________"',True,'center',10.5)
dtable(p1);doc.add_page_break();dtable(p2)
doc.save(DOCX)
print('PDF:',os.path.abspath(PDF),os.path.getsize(PDF),'bytes')
print('DOCX:',os.path.abspath(DOCX),os.path.getsize(DOCX),'bytes')
