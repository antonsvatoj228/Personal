"""Адреси та іменовані діапазони Phoenix Східниця."""
N_PL, N_CAT, N_TAR = 100, 12, 8
COL0 = 9
LAST_COL = COL0 + N_PL*2 - 1          # 208 = GZ
MAXG = 8                               # колонок гостей у сітці «Прайс»

R_TECH = 2
R_PL0 = 6
R_PL_LAST = R_PL0 + N_PL - 1           # 105
R_BLK_NAME, R_BLK_PERIOD, R_BLK_STATUS, R_BLK_UPD = 5, 6, 7, 8
R_BASE0 = 12;  R_BASE1 = R_BASE0 + N_CAT - 1     # 12..23
R_DA0   = 27;  R_DA1   = R_DA0 + N_CAT - 1       # 27..38
R_DB0   = 43;  R_DB1   = R_DB0 + N_CAT - 1       # 43..54
R_PKG0, R_PKG1 = 58, 61
R_SPA0, R_SPA1 = 65, 66
R_SW0,  R_SW1  = 70, 72
R_TAR0 = 14; R_TAR1 = R_TAR0 + N_TAR - 1         # 14..21

SECTION_ROWS = (10, 25, 40, 56, 63, 68)
GUTTER_ROWS  = (9, 24, 39, 55, 62, 67, 73)

KM = "'Керування прайсом'"
DK = "'Довідник категорій'"

def col_letter(i):
    s = ''
    while i > 0:
        i, r = divmod(i-1, 26); s = chr(65+r) + s
    return s
LASTL = col_letter(LAST_COL)
DR1 = 4 + N_CAT                                   # останній рядок довідника = 16

DEFINED = {
    'ПРАЙС_ЛИСТИ':    f'{KM}!$A${R_PL0}:$A${R_PL_LAST}',
    'ПЛ_ДАТА_З':      f'{KM}!$B${R_PL0}:$B${R_PL_LAST}',
    'ПЛ_ДАТА_ПО':     f'{KM}!$C${R_PL0}:$C${R_PL_LAST}',
    'ПЛ_СТАТУС':      f'{KM}!$D${R_PL0}:$D${R_PL_LAST}',
    'ПЛ_РЯДОК':       f'{KM}!$I${R_TECH}:${LASTL}${R_TECH}',
    'БАЗА_ЦІНИ':      f'{KM}!$I${R_BASE0}:${LASTL}${R_BASE1}',
    'БАЗА_КАТ':       f'{KM}!$H${R_BASE0}:$H${R_BASE1}',
    'ПАКЕТИ_ЦІНИ':    f'{KM}!$I${R_PKG0}:${LASTL}${R_PKG1}',
    'ПАКЕТИ_КЛЮЧ':    f'{KM}!$H${R_PKG0}:$H${R_PKG1}',
    'SPA_ЦІНИ':       f'{KM}!$I${R_SPA0}:${LASTL}${R_SPA1}',
    'SPA_РІВЕНЬ':     f'{KM}!$H${R_SPA0}:$H${R_SPA1}',
    'ПЕРЕМИКАЧІ':     f'{KM}!$I${R_SW0}:${LASTL}${R_SW1}',
    'ТАРИФИ_НАЗВИ':   f'{KM}!$F${R_TAR0}:$F${R_TAR1}',
    'ТАРИФИ_ЗНИЖКИ':  f'{KM}!$G${R_TAR0}:$G${R_TAR1}',
    'СПИСОК_ПАКЕТІВ': f'{KM}!$F$25:$F$29',
    'СПИСОК_ДНІВ':    f'{KM}!$F$32:$F$33',
    'СПИСОК_СТАТУСІВ':f'{KM}!$F$36:$F$38',
    'СПИСОК_SPA':     f'{KM}!$H${R_SPA0}:$H${R_SPA1}',
    'ОКРУГЛЕННЯ':     f'{KM}!$G$6',
    'ЗНИЖКА_BBSPA':   f'{KM}!$G$7',
    'ТВАРИНИ':        f'{KM}!$G$8',
    'ЕТАЛОН_ПЛ':      f'{KM}!$G$9',
    'ПЕРШИЙ_ВИХІДНИЙ':f'{KM}!$G$10',
    'ДОВ_КОД':        f'{DK}!$A$5:$A${DR1}',
    'КАТЕГОРІЇ':      f'{DK}!$B$5:$B${DR1}',
    'ДОВ_BOOKING':    f'{DK}!$C$5:$C${DR1}',
    'ДОВ_МІН':        f'{DK}!$D$5:$D${DR1}',
    'ДОВ_ВКЛ':        f'{DK}!$E$5:$E${DR1}',
    'ДОВ_МАКС':       f'{DK}!$F$5:$F${DR1}',
    'ДОВ_ДОПЛ_ДОР':   f'{DK}!$G$5:$G${DR1}',
    'ДОВ_ДОПЛ_ДІТ':   f'{DK}!$H$5:$H${DR1}',
    'ДОВ_SPA':        f'{DK}!$I$5:$I${DR1}',
}

def pl_col(i): return COL0 + i*2
