"""Перебудовує лист «Контроль і порівняння» у наявному файлі Буковеля, зберігаючи селектори."""
import sys, shutil, openpyxl
sys.path.insert(0, '/home/user/Personal/build')
import sheet_control
from style import protect

SRC, DST = 'up3.xlsx', 'buk_v23.xlsx'
shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)
old = wb['Контроль і порівняння']
idx = wb.sheetnames.index('Контроль і порівняння')
sel = {c: old[c].value for c in ('B3','D3','F3','H3','B4','D4','F4')}
was_protected = old.protection.sheet
print('збережені селектори:', sel, '| захист був:', was_protected)

wb.remove(old)
ws = sheet_control.build(wb, None)
wb._sheets.remove(ws); wb._sheets.insert(idx, ws)
for c, v in sel.items():
    if v is not None: ws[c] = v
if not was_protected:
    ws.protection.sheet = False
wb.save(DST)
print('лист перебудовано, позиція', idx)
print('заголовки A6:O6:', [ws.cell(row=6,column=i).value for i in range(1,16)])
