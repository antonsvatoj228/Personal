from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter as gl
from style import *
from names import *

R0 = 7
R1 = R0 + N_CAT - 1
HKEYS = ['ro','mult','wmult','maxg','incl','ming','pa_b','pk_b','pa_v','pk_v','spa','okd','oke']
H1_0, H2_0 = 16, 30                      # P..AB  та  AD..AP
N_TOT = '($D$4+$F$4)'

def hcols(start):
    return {k: gl(start + i) for i, k in enumerate(HKEYS)}

def build(wb, data):
    ws = wb.create_sheet('Контроль і порівняння')
    title(ws,'A1','ПОРІВНЯННЯ ТА КОНТРОЛЬ ПРАЙС-ЛИСТІВ — BUKOVEL','A1:I1')
    for w,c in zip([40,13,13,13,13,12,12,11,11],'ABCDEFGHI'): ws.column_dimensions[c].width = w

    for addr,lab,val,fmt in (('B3','Прайс-лист 1','Осінь 2026 (00-10%)',None),
                             ('D3','Тариф 1','Rack Rate',None),
                             ('F3','Прайс-лист 2','Осінь 2026 (10-20%)',None),
                             ('H3','Тариф 2','Rack Rate',None),
                             ('B4','Пакет','RO',None),
                             ('D4','Дорослих',2,'0'),
                             ('F4','Дітей',0,'0')):
        lcell = gl(ws[addr].column - 1) + addr[1:]
        label(ws,lcell,lab,bold=True); inp(ws,addr,val,fmt)
    for rng,src in (('B3','=ПРАЙС_ЛИСТИ'),('F3','=ПРАЙС_ЛИСТИ'),('D3','=ТАРИФИ_НАЗВИ'),
                    ('H3','=ТАРИФИ_НАЗВИ'),('B4','=СПИСОК_ПАКЕТІВ')):
        dv = DataValidation(type='list', formula1=src, allow_blank=False)
        ws.add_data_validation(dv); dv.add(rng)
    dv_g = DataValidation(type='whole', operator='between', formula1='0', formula2='20',
                          allow_blank=False, showErrorMessage=True,
                          errorTitle='Кількість гостей', error='Введіть ціле число від 0 до 20.')
    ws.add_data_validation(dv_g); dv_g.add('D4'); dv_g.add('F4')

    banner(ws,'A5','Порівняння показує ціну за одну ніч без тварин — для буднього і вихідного дня одночасно. '
                   'Базові (опорні) категорії виділені червоним: саме вони задають ціну решті через множники. '
                   'Значення поза місткістю категорії не показуються.','A5:I5',34)

    for i,h in enumerate(['Категорія','Прайс 1\nбудні','Прайс 1\nвихідні','Прайс 2\nбудні','Прайс 2\nвихідні',
                          'Δ грн\nбудні','Δ грн\nвихідні','Δ %\nбудні','Δ %\nвихідні']):
        col_head(ws,f'{gl(i+1)}6',h)
    ws.row_dimensions[6].height = 34

    _globals(ws)
    _helpers(ws,'$B$3','$D$3',H1_0,'K')
    _helpers(ws,'$F$3','$H$3',H2_0,'L')

    for i in range(N_CAT):
        r, dr, band = R0+i, 5+i, i % 2 == 1
        ws[f'A{r}'] = f"=IF('Довідник категорій'!A{dr}=\"\",\"\",'Довідник категорій'!A{dr})"
        out(ws,f'A{r}',None,band=band); ws[f'A{r}'].font = f(b=True)
        ws[f'B{r}'] = _price(H1_0,'K',False,r); ws[f'C{r}'] = _price(H1_0,'K',True,r)
        ws[f'D{r}'] = _price(H2_0,'L',False,r); ws[f'E{r}'] = _price(H2_0,'L',True,r)
        ws[f'F{r}'] = f'=IF(OR(B{r}="",D{r}=""),"",D{r}-B{r})'
        ws[f'G{r}'] = f'=IF(OR(C{r}="",E{r}=""),"",E{r}-C{r})'
        ws[f'H{r}'] = f'=IF(OR(B{r}="",D{r}="",B{r}=0),"",D{r}/B{r}-1)'
        ws[f'I{r}'] = f'=IF(OR(C{r}="",E{r}="",C{r}=0),"",E{r}/C{r}-1)'
        for c in 'BCDEFG': out(ws,f'{c}{r}',None,MONEY,band=band)
        for c in 'HI':     out(ws,f'{c}{r}',None,PCT,band=band)

    ws.conditional_formatting.add(f'A{R0}:I{R1}',
        FormulaRule(formula=[f'IFERROR(INDEX(ДОВ_РОЛЬ,MATCH($A{R0},КАТЕГОРІЇ,0))="Базова",FALSE)'],
                    fill=fill(RED_SOFT), font=f(11,True,RED_INK), stopIfTrue=False))
    for rng in (f'F{R0}:G{R1}', f'H{R0}:I{R1}'):
        ws.conditional_formatting.add(rng, CellIsRule(operator='lessThan', formula=['0'],
                                                      font=f(11,color=RED_INK)))
    _checklist(ws)

    for c in range(14, H2_0 + len(HKEYS)):
        ws.column_dimensions[gl(c)].hidden = True
    ws.freeze_panes = f'B{R0}'
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = GREEN_M
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f'A1:I{R1}'; ws.print_title_rows = '1:6'
    return ws


def _globals(ws):
    ws['N1'] = 'ТЕХНІЧНІ ДАНІ — НЕ РЕДАГУВАТИ'; ws['N1'].font = f(10,True,'6B7771')
    for r,fml,lab in ((3,"=IFERROR(MATCH('Прайс'!$B$3,ПЛ_РЯДОК,0),0)",'колонка активного прайсу'),
                      (5,'=IF(OR($B$4="BB",$B$4="BBSPA",$B$4="ROSPABB"),1,0)','потрібен пакет'),
                      (6,'=IF($N$5=0,"",IF($B$4="ROSPABB","BB",$B$4))','ключ пакета'),
                      (7,'=IF(OR($B$4="ROSPA",$B$4="ROSPABB"),1,0)','потрібен SPA')):
        ws[f'N{r}'] = fml; ws[f'N{r}'].font = f(9,color='6B7771')
        ws[f'O{r}'] = lab; ws[f'O{r}'].font = f(9,it=True,color='9AA6A0')


def _helpers(ws, pl_cell, tar_cell, start, tag):
    c = hcols(start); H = gl(start)
    ws[f'{H}1'] = f'{tag}: службові'; ws[f'{H}1'].font = f(9,True,'6B7771')
    ws[f'{H}2'] = f'=IFERROR(MATCH({pl_cell},ПЛ_РЯДОК,0),0)'
    ws[f'{H}3'] = (f'=IF($B$4="BBSPA",ЗНИЖКА_BBSPA,IFERROR(SUMIFS(ТАРИФИ_ЗНИЖКИ,ТАРИФИ_НАЗВИ,{tar_cell}),0))')
    ws[f'{H}4'] = f'=IF({H}$2=0,0,IFERROR(INDEX(ПЕРЕМИКАЧІ,IF($B$4="BBSPA",2,1),{H}$2),0))'
    ws[f'{H}5'] = f'=IF({H}$2=0,0,IFERROR(INDEX(ПЕРЕМИКАЧІ,3,{H}$2),0))'
    for r in range(2,6): ws[f'{H}{r}'].font = f(9,color='6B7771')
    for i in range(N_CAT):
        r = R0 + i
        m = f'MATCH($A{r},КАТЕГОРІЇ,0)'
        ws[f'{c["ro"]}{r}']   = (f'=IF(OR({H}$2=0,$A{r}=""),0,IFERROR(INDEX(БАЗА_RO,'
                                f'MATCH(INDEX(ДОВ_ОПОРА,{m}),БАЗА_КАТ,0),{H}$2),0))')
        ws[f'{c["mult"]}{r}'] = f'=IF(OR({H}$2=0,$A{r}=""),0,IFERROR(INDEX(МНОЖНИКИ,MATCH($A{r},МНОЖ_КАТ,0),{H}$2),0))'
        ws[f'{c["wmult"]}{r}']= f'=IF(OR({H}$2=0,$A{r}=""),0,IFERROR(INDEX(МНОЖНИКИ,MATCH($A{r},МНОЖ_КАТ,0),{H}$2+1),0))'
        ws[f'{c["maxg"]}{r}'] = f'=IF($A{r}="",0,IFERROR(INDEX(ДОВ_МАКС,{m}),0))'
        ws[f'{c["incl"]}{r}'] = f'=IF($A{r}="",0,IFERROR(INDEX(ДОВ_ВКЛ,{m}),0))'
        ws[f'{c["ming"]}{r}'] = f'=IF($A{r}="",0,IFERROR(INDEX(ДОВ_МІН,{m}),0))'
        for key,day,off in (('pa_b','Будні',0),('pk_b','Будні',1),('pa_v','Вихідні',0),('pk_v','Вихідні',1)):
            ws[f'{c[key]}{r}'] = (f'=IF(OR({H}$2=0,$N$5=0),0,IFERROR(INDEX(ПАКЕТИ_ЦІНИ,'
                                  f'MATCH($N$6&"|{day}",ПАКЕТИ_КЛЮЧ,0),{H}$2+{off}),0))')
        ws[f'{c["spa"]}{r}']  = (f'=IF(OR({H}$2=0,$N$7=0,$A{r}=""),0,IFERROR(INDEX(SPA_ЦІНИ,'
                                f'MATCH(INDEX(ДОВ_SPA,{m}),SPA_РОЛЬ,0),{H}$2),0))')
        ws[f'{c["okd"]}{r}']  = (f'=IF(OR($A{r}="",{c["ro"]}{r}=0,{c["mult"]}{r}=0,'
                                f'AND($N$5=1,$D$4>0,{c["pa_b"]}{r}=0),AND($N$5=1,$F$4>0,{c["pk_b"]}{r}=0),'
                                f'AND($N$7=1,{c["spa"]}{r}=0)),0,1)')
        ws[f'{c["oke"]}{r}']  = (f'=IF(OR({c["okd"]}{r}=0,{c["wmult"]}{r}=0,'
                                f'AND($N$5=1,$D$4>0,{c["pa_v"]}{r}=0),AND($N$5=1,$F$4>0,{c["pk_v"]}{r}=0)),0,1)')
        for k in c.values(): ws[f'{k}{r}'].font = f(9,color='6B7771')


def _price(start, tag, wknd, r):
    c = hcols(start); H = gl(start)
    ok  = c['oke'] if wknd else c['okd']
    pa  = c['pa_v'] if wknd else c['pa_b']
    pk  = c['pk_v'] if wknd else c['pk_b']
    core = f'{c["ro"]}{r}*{c["mult"]}{r}' + (f'*{c["wmult"]}{r}' if wknd else '')
    return (f'=IF(OR({ok}{r}=0,{N_TOT}=0,{N_TOT}>{c["maxg"]}{r},{N_TOT}<{c["ming"]}{r}),"",'
            f'ROUND({core}*(1-{H}$3)/ОКРУГЛЕННЯ,0)*ОКРУГЛЕННЯ'
            f'+MAX(0,{N_TOT}-{c["incl"]}{r})*КРОК_ГОСТЯ*(1-{H}$3)'
            f'+($D$4*{pa}{r}+$F$4*{pk}{r})*(1-{H}$3*{H}$4)'
            f'+{c["spa"]}{r}*(1-{H}$3*{H}$5))')


def _checklist(ws):
    block_head(ws,'K3','КОНТРОЛЬ АКТИВНОГО ПРАЙС-ЛИСТА','K3:M3')
    ws.column_dimensions['K'].width = 46; ws.column_dimensions['L'].width = 15
    ws.column_dimensions['M'].width = 3
    checks = [
      (4,'Незаповнені базові RO-ціни',
       '=4-COUNT(INDEX(БАЗА_RO,1,$N$3),INDEX(БАЗА_RO,2,$N$3),INDEX(БАЗА_RO,3,$N$3),INDEX(БАЗА_RO,4,$N$3))'),
      (5,'Незаповнені множники',
       f'={2*N_CAT}-COUNT(INDEX(МНОЖНИКИ,0,$N$3))-COUNT(INDEX(МНОЖНИКИ,0,$N$3+1))'),
      (6,'Незаповнені пакети BB',
       '=4-COUNT(INDEX(ПАКЕТИ_ЦІНИ,1,$N$3),INDEX(ПАКЕТИ_ЦІНИ,1,$N$3+1),'
       'INDEX(ПАКЕТИ_ЦІНИ,2,$N$3),INDEX(ПАКЕТИ_ЦІНИ,2,$N$3+1))'),
      (7,'Незаповнений фіксований SPA-візит',
       '=2-COUNT(INDEX(SPA_ЦІНИ,1,$N$3),INDEX(SPA_ЦІНИ,2,$N$3))'),
      (8,'Дублікати назв прайс-листів',
       '=SUMPRODUCT((ПРАЙС_ЛИСТИ<>"")*(COUNTIF(ПРАЙС_ЛИСТИ,ПРАЙС_ЛИСТИ&"")>1))'),
    ]
    for r,lab,fml in checks:
        label(ws,f'K{r}',lab); ws[f'L{r}'] = fml; out(ws,f'L{r}',None,'0')
        ws.conditional_formatting.add(f'L{r}', CellIsRule(operator='greaterThan', formula=['0'],
                                     fill=fill(RED_SOFT), font=f(11,True,RED_INK)))
    ws['K10'] = ('=IF(SUM($L$4:$L$8)=0,"ГОТОВО ДО ПУБЛІКАЦІЇ","ПОТРІБНЕ ЗАПОВНЕННЯ")')
    ws['K10'].font = f(12,True); ws['K10'].alignment = Alignment(horizontal='center',vertical='center')
    ws.merge_cells('K10:L10'); ws.row_dimensions[10].height = 26
    ws.conditional_formatting.add('K10:L10', FormulaRule(formula=['$K$10="ГОТОВО ДО ПУБЛІКАЦІЇ"'],
                                  fill=fill('E1EFE7'), font=f(12,True,'2E6B4F'), stopIfTrue=True))
    ws.conditional_formatting.add('K10:L10', FormulaRule(formula=['$K$10="ПОТРІБНЕ ЗАПОВНЕННЯ"'],
                                  fill=fill(RED_SOFT), font=f(12,True,RED_INK), stopIfTrue=True))
    info = [
      (12,'Пакет BBSPA заповнено?',
       '=IF(COUNT(INDEX(ПАКЕТИ_ЦІНИ,3,$N$3),INDEX(ПАКЕТИ_ЦІНИ,3,$N$3+1),'
       'INDEX(ПАКЕТИ_ЦІНИ,4,$N$3),INDEX(ПАКЕТИ_ЦІНИ,4,$N$3+1))=4,"так","ні — пакет недоступний")',None),
      (13,'Прайс-листів у моделі','=COUNTIF(ПРАЙС_ЛИСТИ,"<>")','0'),
      (14,'Категорій у моделі','=COUNTIF(КАТЕГОРІЇ,"<>")','0'),
      (15,'Крок за 3-го гостя, грн','=КРОК_ГОСТЯ',MONEY),
      (16,'Крок округлення, грн','=ОКРУГЛЕННЯ',MONEY),
    ]
    for r,lab,fml,fmt in info:
        label(ws,f'K{r}',lab); ws[f'L{r}'] = fml; out(ws,f'L{r}',None,fmt,band=True)
    label(ws,'K18','BBSPA не входить до чек-листа: пакет може бути свідомо не задіяний у сезоні. '
                   'Його стан показано окремим рядком вище.')
    ws['K18'].font = f(9,it=True,color='6B7771')
    ws['K18'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells('K18:L19'); ws.row_dimensions[18].height = 26
