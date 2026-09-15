import json, datetime, sys
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from style import *
from names import *
import sheet_price, sheet_calc, sheet_keruvannya, sheet_control, sheet_rest, sheet_dovidnyk

VERSION = 'v2.0'
TODAY   = datetime.date(2026, 9, 15)

PERIODS = {   # текстові періоди оригіналу → справжні дати
    'Осінь 2026 (00-10%)': (datetime.date(2026,10,1), datetime.date(2026,11,30)),
    'Осінь 2026 (10-20%)': (datetime.date(2026,9,1),  datetime.date(2026,9,30)),
}

def load():
    d = json.load(open('source_data.json', encoding='utf8'))
    for pl in d['pricelists']:
        pl['name'] = pl['name'].strip()
        df, dt = PERIODS.get(pl['name'], (None, None))
        pl['date_from'], pl['date_to'] = df, dt
    # ── ЄДИНА ЗМІНА ДАНИХ: «Апартаменти Люкс 25 м²» — макс. місткість 3 → 2 ──
    for row in d['cats']:
        if row[0] == 'Апартаменти Люкс 25 м²':
            assert row[6] == 3.0, f'очікували макс. гостей 3, отримали {row[6]}'
            row[6] = 2.0
    for row in d['cats']:
        for i in (4,5,6):
            row[i] = int(row[i]) if row[i] is not None else None
    return d

def main(outpath):
    data = load()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheet_dovidnyk.build(wb, data)
    sheet_keruvannya.build(wb, data)
    sheet_price.build(wb, data)
    sheet_calc.build(wb, data)
    sheet_control.build(wb, data)
    sheet_rest.build_model(wb, data)
    sheet_rest.build_matrix(wb, data)
    sheet_rest.build_base(wb, data)
    sheet_rest.build_extra(wb, data)
    sheet_rest.build_log(wb, data, VERSION, TODAY)

    for name, ref in DEFINED.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    order = ['Прайс','Калькулятор','Керування прайсом','Контроль і порівняння','Моделювання',
             'Матриця категорій','Базові ціни','Довідник категорій','Додаткові послуги','Журнал змін']
    wb._sheets = [wb[n] for n in order]

    for name in order:
        if name != 'Журнал змін':
            protect(wb[name])
    wb['Журнал змін'].protection.sheet = False

    wb.save(outpath)
    print('збережено:', outpath)
    print('листів:', len(wb.sheetnames), '| іменованих діапазонів:', len(DEFINED))

if __name__ == '__main__':
    main(sys.argv[1])
