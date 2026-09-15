"""Округлення — лише на базову ціну 1–2 ос.; крок за додаткового гостя додається точно."""
import openpyxl, shutil
from openpyxl.utils import get_column_letter as gl

SRC, DST = 'uploaded.xlsx', 'v2.2.xlsx'
shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)

# ─────────── ПРАЙС ───────────
p = wb['Прайс']
R0, R1 = 9, 26
WD0, WD1, SPACER, WE0, WE1 = 4, 11, 12, 13, 20
H = dict(ro='W', mult='X', wmult='Y', maxg='Z', incl='AA',
         pkgd='AB', pkge='AC', spa='AD', okd='AE', oke='AF')
n_price = 0
for r in range(R0, R1 + 1):
    for c in range(WD0, WE1 + 1):
        if c == SPACER: continue
        wknd = c >= WE0
        g   = f'{gl(c)}$8'                      # кількість гостей із шапки
        ok  = H['oke'] if wknd else H['okd']
        pkg = H['pkge'] if wknd else H['pkgd']
        base = f'{H["ro"]}{r}*{H["mult"]}{r}' + (f'*{H["wmult"]}{r}' if wknd else '')
        p.cell(row=r, column=c).value = (
            f'=IF(OR(${ok}{r}=0,{g}>${H["maxg"]}{r}),"",'
            f'ROUND({base}*(1-$U$5)/ОКРУГЛЕННЯ,0)*ОКРУГЛЕННЯ'
            f'+MAX(0,{g}-${H["incl"]}{r})*КРОК_ГОСТЯ*(1-$U$5)'
            f'+{g}*${pkg}{r}*(1-$U$5*$U$6)'
            f'+${H["spa"]}{r}*(1-$U$5*$U$7))')
        n_price += 1

# ─────────── КОНТРОЛЬ І ПОРІВНЯННЯ ───────────
k = wb['Контроль і порівняння']
KR0, KR1 = 7, 24
N = '($D$4+$F$4)'
def kcols(start):
    keys = ['ro','mult','wmult','maxg','incl','ming','pa_b','pk_b','pa_v','pk_v','spa','okd','oke']
    return {key: gl(start + i) for i, key in enumerate(keys)}
n_ctl = 0
for col_out, start, wknd in ((2, 16, False), (3, 16, True), (4, 30, False), (5, 30, True)):
    c = kcols(start); Hc = gl(start)
    for r in range(KR0, KR1 + 1):
        ok   = c['oke'] if wknd else c['okd']
        pa   = c['pa_v'] if wknd else c['pa_b']
        pk   = c['pk_v'] if wknd else c['pk_b']
        base = f'{c["ro"]}{r}*{c["mult"]}{r}' + (f'*{c["wmult"]}{r}' if wknd else '')
        k.cell(row=r, column=col_out).value = (
            f'=IF(OR({ok}{r}=0,{N}=0,{N}>{c["maxg"]}{r},{N}<{c["ming"]}{r}),"",'
            f'ROUND({base}*(1-{Hc}$3)/ОКРУГЛЕННЯ,0)*ОКРУГЛЕННЯ'
            f'+MAX(0,{N}-{c["incl"]}{r})*КРОК_ГОСТЯ*(1-{Hc}$3)'
            f'+($D$4*{pa}{r}+$F$4*{pk}{r})*(1-{Hc}$3*{Hc}$4)'
            f'+{c["spa"]}{r}*(1-{Hc}$3*{Hc}$5))')
        n_ctl += 1

# ─────────── КАЛЬКУЛЯТОР ───────────
c = wb['Калькулятор']
c['D22'] = 'Проживання 1–2 ос. до знижки, грн'
c['F22'] = '=F11*IF($B$9="Вихідні",F13,1)'
c['F23'] = ('=IF($F$6<>"OK","",ROUND(F22*(1-F15)/ОКРУГЛЕННЯ,0)*ОКРУГЛЕННЯ'
            '+F14*КРОК_ГОСТЯ*(1-F15)'
            '+($B$10*F18+$B$11*F19)*(1-F15*F16)'
            '+F20*(1-F15*F17)+F21)')

# ─────────── ПОЯСНЕННЯ В ІНТЕРФЕЙСІ ───────────
p['A6'] = ('=IF($U$3=0,"⚠ Прайс-лист не обрано або не знайдено — оберіть значення у клітинці B3.",'
           'IF($U$13=0,"⚠ У цьому прайс-листі ще не заповнені базові RO-ціни — сітка буде порожньою. '
           'Заповніть його блок на листі «Керування прайсом».",'
           '"Усі ціни — за ніч, за дорослих гостей, у гривні. Дитячі тарифи рахує лист «Калькулятор». '
           'Крок округлення діє на базову ціну за 1–2 особи; кожен наступний гість додає рівно крок за гостя × знижку."))')
wb['Керування прайсом']['H70'] = (
    'Логіка: RO-ціна задається лише для 4 опорних категорій, решта формуються множниками. '
    'Крок округлення (G7) застосовується до базової ціни за 1–2 особи після знижки; '
    'крок за 3-го й наступних гостей (G6) додається точно, помножений на знижку, і не округлюється. '
    'BB / BBSPA додають пакет за кожного гостя, ROSPA — один SPA-візит на об\'єкт, '
    'ROSPABB — сніданки за кількістю гостей + один SPA-візит.')

wb.save(DST)
print(f'оновлено формул: Прайс {n_price}, Контроль {n_ctl}, Калькулятор 2')
print('приклад Прайс D9:'); print(' ', wb['Прайс']['D9'].value)

# ═══════════ ВІДНОВЛЕННЯ НАЛАШТУВАНЬ, ВТРАЧЕНИХ ПРИ РЕДАГУВАННІ ═══════════
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.load_workbook(DST)
RED_SOFT, RED_INK = 'FCE8E4', 'A8321E'
def redrule():
    return dict(fill=PatternFill('solid', fgColor=RED_SOFT),
                font=Font(name='Calibri', size=11, bold=True, color=RED_INK))

# 1. Червона підсвітка опорних категорій
for sheet, rng, anchor, tint in (('Прайс','A9:C26','$A9',False),
                                 ('Контроль і порівняння','A7:I24','$A7',True),
                                 ('Моделювання','A8:C25','$A8',False),
                                 ('Матриця категорій','A7:B24','$A7',True)):
    ws = wb[sheet]
    style = redrule() if tint else dict(font=Font(name='Calibri', size=11, bold=True, color=RED_INK))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'IFERROR(INDEX(ДОВ_РОЛЬ,MATCH({anchor},КАТЕГОРІЇ,0))="Базова",FALSE)'],
        stopIfTrue=False, **style))

# 2. Випадаючі списки
kc = wb['Контроль і порівняння']
for rng, src in (('B3','=ПРАЙС_ЛИСТИ'), ('F3','=ПРАЙС_ЛИСТИ'),
                 ('D3','=ТАРИФИ_НАЗВИ'), ('H3','=ТАРИФИ_НАЗВИ')):
    dv = DataValidation(type='list', formula1=src, allow_blank=False)
    kc.add_data_validation(dv); dv.add(rng)
dvg = DataValidation(type='whole', operator='between', formula1='0', formula2='20',
                     allow_blank=False, showErrorMessage=True, errorTitle='Кількість гостей',
                     error='Введіть ціле число від 0 до 20.')
kc.add_data_validation(dvg); dvg.add('D4'); dvg.add('F4')
mo = wb['Моделювання']
dvm = DataValidation(type='decimal', operator='between', formula1='0', formula2='10',
                     allow_blank=True, showErrorMessage=True, errorTitle='Множник',
                     error='Множник має бути числом від 0 до 10.')
mo.add_data_validation(dvm); dvm.add('H8:I25'); dvm.add('N8:O25')

# 3. Друк і закріплення
pr = wb['Прайс']
pr.print_area = 'A1:T26'; pr.print_title_rows = '1:8'; pr.freeze_panes = 'D9'
pr.page_setup.orientation = 'landscape'; pr.page_setup.paperSize = pr.PAPERSIZE_A4
pr.page_setup.fitToWidth = 1; pr.page_setup.fitToHeight = 0
pr.sheet_properties.pageSetUpPr.fitToPage = True
ca = wb['Калькулятор']
ca.print_area = 'A1:F29'; ca.page_setup.fitToWidth = 1; ca.page_setup.fitToHeight = 0
ca.sheet_properties.pageSetUpPr.fitToPage = True
kc.print_area = 'A1:I24'; kc.print_title_rows = '1:6'
kc.page_setup.orientation = 'landscape'; kc.page_setup.fitToWidth = 1; kc.page_setup.fitToHeight = 0
kc.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(DST)
print('відновлено: 4 правила підсвітки, 7 випадаючих списків, 3 області друку, закріплення «Прайсу»')
