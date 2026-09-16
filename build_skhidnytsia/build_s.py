import json, datetime, sys, re
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from style import *
from names_s import *
import s_price, s_calc, s_keruvannya, s_rest, s_dovidnyk

TODAY = datetime.date(2026, 9, 16)

def parse_period(txt):
    """'01.09.2025–30.11.2025' → (date, date); інакше (None, None)."""
    if not txt: return None, None
    m = re.findall(r'(\d{2})[.](\d{2})[.](\d{4})', str(txt))
    if len(m) == 2:
        try:
            return (datetime.date(int(m[0][2]), int(m[0][1]), int(m[0][0])),
                    datetime.date(int(m[1][2]), int(m[1][1]), int(m[1][0])))
        except ValueError:
            return None, None
    return None, None

def load():
    d = json.load(open('data_s.json', encoding='utf8'))
    for pl in d['pricelists']:
        pl['name'] = str(pl['name']).strip()
        df, dt = parse_period(pl.get('period'))
        pl['date_from'], pl['date_to'] = df, dt
        # текстовий період зберігаємо, якщо дати не розпізналися
        pl['period_text'] = None if df else pl.get('period')
        if pl.get('status') == 'Неактивний': pl['status'] = 'Неактивний'
    for row in d['cats']:
        for i in (3,4,5):
            row[i] = int(row[i]) if row[i] is not None else None
    return d

def main(outpath):
    data = load()
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    s_dovidnyk.build(wb, data)
    s_keruvannya.build(wb, data)
    s_price.build(wb, data)
    s_calc.build(wb, data)
    s_rest.build_control(wb, data)
    s_rest.build_model(wb, data)
    s_rest.build_base(wb, data)
    s_rest.build_extra(wb, data)
    s_rest.build_log(wb, data, TODAY)
    for name, ref in DEFINED.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))
    order = ['Прайс','Калькулятор','Керування прайсом','Контроль і порівняння','Моделювання',
             'Базові ціни','Довідник категорій','Додаткові послуги','Журнал змін']
    wb._sheets = [wb[n] for n in order]
    for n in order:
        if n != 'Журнал змін': protect(wb[n])
    wb['Журнал змін'].protection.sheet = False
    wb.save(outpath)
    print('збережено:', outpath)
    print('листів:', len(wb.sheetnames), '| діапазонів:', len(DEFINED),
          '| прайс-листів:', N_PL, '| категорій:', N_CAT)

if __name__ == '__main__':
    main(sys.argv[1])
