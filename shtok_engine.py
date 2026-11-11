# -*- coding: utf-8 -*-
"""
P&L-агент ТОО «Saurida Oil & Gas» (СОГ).
Схема как в CSS/Градексе, но структура проще — без проектов:
  6010 доход: вид (ур.2) -> контрагент (ур.3)   -> Кредит (кол.6)   => доход по контрагентам
  7010 себестоимость: статья (ур.2)             -> Дебет  (кол.5)
  7110 расходы по реализации: статья (ур.2)      -> Дебет  (кол.5)   => прямые производственные затраты по статьям
  7211 административные: статья (ур.2)            -> Дебет  (кол.5)
  7212 не идущие на вычет: статья (ур.2)          -> Дебет  (кол.5)   => отдельный блок ниже
  6110 депозит, 6250 курс.доход                  -> низ
  7300/7310 финансирование, 7430 курс.расход     -> низ
Налоги с ЗП (ОПВ+ОСМС+соц.отчисления+соц.налог) сворачиваются в «Налоги по ЗП».
"""
import re, openpyxl
from collections import defaultdict, OrderedDict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

THRESHOLD = 300_000
_TAX_KEYS = ('обязательные пенсионные взносы', 'отчисления осмс',
             'социальные отчисления', 'социальный налог')
TAX_LINE = 'Налоги по заработной плате'


def norm(s):
    return re.sub(r'\s+', ' ', str(s or '').strip().lower()).replace('ё', 'е')

def _num(x):
    return float(x) if isinstance(x, (int, float)) else 0.0


def group_taxes(d):
    """ОПВ+ОСМС+соц → одна строка «Налоги по заработной плате» (после строки ЗП)."""
    out = OrderedDict(); tax = 0.0
    for art, amt in sorted(d.items(), key=lambda x: -x[1]):
        if any(k in norm(art) for k in _TAX_KEYS):
            tax += amt
        else:
            out[art] = out.get(art, 0.0) + amt
    if tax:
        res = OrderedDict(); placed = False
        for art, amt in out.items():
            res[art] = amt
            if not placed and 'заработн' in norm(art) and 'плат' in norm(art):
                res[TAX_LINE] = tax; placed = True
        if not placed:
            res[TAX_LINE] = tax
        return res
    return out


def parse_1c(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for r in range(1, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if name is None:
            continue
        name = str(name).strip()
        if name in ('<...>', ''):
            continue
        od = ws.row_dimensions[r].outline_level if r in ws.row_dimensions else 0
        deb = _num(ws.cell(row=r, column=5).value)
        cred = _num(ws.cell(row=r, column=6).value)
        rows.append((r, od, name, deb, cred))

    R = {
        'income': defaultdict(float),       # контрагент -> выручка (6010 ур.3)
        'direct': defaultdict(float),       # статья -> сумма (7010+7110 ур.2)
        'admin': defaultdict(float),        # статья -> сумма (7211 ур.2)
        'admin_7212': defaultdict(float),   # статья -> сумма (7212 ур.2)
        'deposit': 0.0, 'fx_income': 0.0, 'other_income': 0.0,
        'financing': 0.0, 'fx_expense': 0.0, 'other_expense': 0.0,
        'tot_income_6': 0.0, 'tot_expense_7': 0.0,
        'direct_target': 0.0, 'admin_target': 0.0, 'admin7212_target': 0.0, 'admin_grand': 0.0,
    }
    GENERIC = ('основная номенклатурная группа', '<...>', 'основное подразделение')
    acct = None; cur_service = None; cur_article = None
    code_od = {}   # код счёта -> его уровень (чтобы ловить статьи относительно родителя)
    for i, (r, od, name, deb, cred) in enumerate(rows):
        nxt = rows[i + 1][1] if i + 1 < len(rows) else 0
        is_leaf = nxt <= od
        m = re.match(r'^(\d{4}),', name)
        if m:
            acct = m.group(1); code_od[acct] = od
            if od == 0:
                if acct.startswith('6'):
                    R['tot_income_6'] += cred
                elif acct.startswith('7') and acct < '7700':
                    R['tot_expense_7'] += deb
            # цели/низ — на строке самого счёта, независимо от уровня (у разных компаний уровни разные)
            if acct in ('7010', '7110'):
                R['direct_target'] += deb
            elif acct == '7200' and od == 0:
                R['admin_grand'] = R.get('admin_grand', 0) + deb   # весь админ 7200 (= вычит. + 7212)
            elif acct == '7212':
                R['admin7212_target'] += deb
            elif acct == '6110':
                R['deposit'] += cred
            elif acct == '6250':
                R['fx_income'] += cred
            elif acct == '6290':
                R['other_income'] += cred
            elif acct in ('7310', '7320'):
                R['financing'] += deb
            elif acct == '7430':
                R['fx_expense'] += deb
            elif acct == '7480':
                R['other_expense'] += deb
            continue
        a = acct or ''
        # доход 6010: вид услуги (первый уровень под счётом) -> контрагент/группа (лист).
        if a == '6010':
            base = code_od.get('6010', 1)
            if od == base + 1:
                cur_service = name
            if is_leaf and cred:
                lbl = cur_service if norm(name) in GENERIC else name
                R['income'][lbl] += cred
            continue
        # прямые затраты 7010/7011/7110: статья (ур.2) -> проект (лист). Метка «проект: статья».
        if a in ('7010', '7011', '7110', '7111'):
            base = max(code_od.get('7010', 1), code_od.get('7011', 0), code_od.get('7110', 0))
            if od == base + 1:
                cur_article = name
            if is_leaf and deb:
                if cur_article and od > base + 1:
                    lbl = f'{name}: {cur_article}' if norm(name) not in GENERIC else cur_article
                else:
                    lbl = name
                R['direct'][lbl] += deb
            continue
        # админ (вычитаемый): статьи под счетами 7210/7211 (кроме 7212). Статья = прямой ребёнок счёта.
        if a in ('7210', '7211'):
            if od == code_od.get(a, 1) + 1 and deb:
                R['admin'][name] += deb
            continue
        if a == '7212':
            if od == code_od.get('7212', 1) + 1 and deb:
                R['admin_7212'][name] += deb
            continue

    # цель вычитаемого админа = весь 7200 минус 7212
    R['admin_target'] = R.get('admin_grand', 0) - R['admin7212_target']
    # сверка остатков по целям
    for key, tgt in [('direct', 'direct_target'), ('admin', 'admin_target'), ('admin_7212', 'admin7212_target')]:
        resid = R[tgt] - sum(R[key].values())
        if abs(resid) > 1:
            lbl = 'Прочие (свёрнуто <...> в 1С)'
            R[key][lbl] = R[key].get(lbl, 0) + resid
    return R


# =====================================================================
#  ПОСТРОЕНИЕ P&L
# =====================================================================
_THIN = Side(style='thin', color='D9D9D9')
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_F = 'Arial'

def _sn(c):
    c.number_format = '#,##0;(#,##0);-'; c.alignment = Alignment(horizontal='right')
    c.font = Font(name=_F, size=10); c.border = _BORDER

def build_workbook(periods, out_path, company='ТОО «Штокпром»'):
    # контрагенты и статьи — объединение по всем периодам
    cust = defaultdict(float)
    for _, R in periods:
        for c, v in R['income'].items():
            cust[c] += v
    customers = [c for c, _ in sorted(cust.items(), key=lambda x: -x[1])]

    direct_raw = defaultdict(float); admin_raw = defaultdict(float); a7212_raw = defaultdict(float)
    for _, R in periods:
        for c, v in R['direct'].items():
            direct_raw[c] += v
        for c, v in R['admin'].items():
            admin_raw[c] += v
        for c, v in R['admin_7212'].items():
            a7212_raw[c] += v
    direct_lines = list(group_taxes(direct_raw).keys())
    admin_lines = list(group_taxes(admin_raw).keys())
    lines_7212 = list(group_taxes(a7212_raw).keys())

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'P&L 2026'
    ncol = len(periods)
    dcols = [get_column_letter(3 + i) for i in range(ncol)]
    tcol = get_column_letter(3 + ncol)
    r = 1

    ws.cell(row=r, column=1, value=company).font = Font(name=_F, size=12, bold=True)
    for i, (nm, _) in enumerate(periods):
        c = ws.cell(row=r, column=3 + i, value=nm); c.font = Font(name=_F, size=10, bold=True); c.alignment = Alignment(horizontal='center')
    c = ws.cell(row=r, column=3 + ncol, value='ИТОГО'); c.font = Font(name=_F, size=10, bold=True); c.alignment = Alignment(horizontal='center')
    r += 1

    def section(title):
        nonlocal r
        cell = ws.cell(row=r, column=2, value=title); cell.font = Font(name=_F, size=10, bold=True, color='FFFFFF')
        for cc in range(2, 4 + ncol):
            ws.cell(row=r, column=cc).fill = PatternFill('solid', fgColor='2F5496'); ws.cell(row=r, column=cc).border = _BORDER
        r += 1

    def leaf(label, getter, indent=1):
        nonlocal r
        lc = ws.cell(row=r, column=2, value=('   ' * indent) + label); lc.font = Font(name=_F, size=10); lc.border = _BORDER
        for i, (_, R) in enumerate(periods):
            _sn(ws.cell(row=r, column=3 + i, value=round(getter(R), 2)))
        tc = ws.cell(row=r, column=3 + ncol, value=(f'=SUM({dcols[0]}{r}:{dcols[-1]}{r})' if ncol > 1 else f'={dcols[0]}{r}'))
        _sn(tc); tc.font = Font(name=_F, size=10, bold=True)
        r += 1; return r - 1

    def total(label, rowrefs=None, formula=None, fill='D9E1F2'):
        nonlocal r
        lc = ws.cell(row=r, column=2, value=label); lc.font = Font(name=_F, size=10, bold=True); lc.border = _BORDER
        if fill:
            for cc in range(2, 4 + ncol):
                ws.cell(row=r, column=cc).fill = PatternFill('solid', fgColor=fill)
        for col in dcols + [tcol]:
            f = formula(col) if formula else ('=' + '+'.join(f'{col}{rr}' for rr in rowrefs) if rowrefs else '=0')
            c = ws.cell(row=r, column=column_index_from_string(col), value=f); _sn(c); c.font = Font(name=_F, size=10, bold=True)
        r += 1; return r - 1

    # ---- ДОХОД ОТ РЕАЛИЗАЦИИ (по контрагентам) ----
    section('ДОХОД ОТ РЕАЛИЗАЦИИ')
    inc_rows = [leaf(c, lambda R, c=c: R['income'].get(c, 0.0)) for c in customers]
    row_inc = total('ИТОГО ДОХОД ОТ РЕАЛИЗАЦИИ', rowrefs=inc_rows, fill='BDD7EE'); r += 1

    # ---- ПРЯМЫЕ ПРОИЗВОДСТВЕННЫЕ ЗАТРАТЫ (по статьям) ----
    section('ПРЯМЫЕ РАСХОДЫ')
    dir_rows = [leaf(a, lambda R, a=a: group_taxes(R['direct']).get(a, 0.0)) for a in direct_lines]
    row_dir = total('ИТОГО ПРЯМЫЕ ЗАТРАТЫ', rowrefs=dir_rows, fill='BDD7EE')
    row_gross = total('ВАЛОВАЯ ПРИБЫЛЬ', formula=lambda col: f'={col}{row_inc}-{col}{row_dir}', fill='FFF2CC'); r += 1

    # ---- АДМИНИСТРАТИВНЫЕ (7211, по статьям) ----
    section('АДМИНИСТРАТИВНЫЕ РАСХОДЫ (относимые на вычеты)')
    adm_rows = [leaf(a, lambda R, a=a: group_taxes(R['admin']).get(a, 0.0)) for a in admin_lines]
    row_adm = total('ИТОГО АДМИНИСТРАТИВНЫЕ', rowrefs=adm_rows, fill='BDD7EE'); r += 1

    # ---- ПРОЧИЕ ДОХОДЫ И РАСХОДЫ ----
    section('ПРОЧИЕ ДОХОДЫ И РАСХОДЫ')
    r_dep = leaf('Доходы по вознаграждениям (депозиты)', lambda R: R['deposit'])
    r_fxnet = leaf('Курсовая разница (нетто: + доход / − расход)', lambda R: R['fx_income'] - R['fx_expense'])
    r_oi = leaf('Прочие доходы', lambda R: R['other_income'])
    r_fin = leaf('Расходы на финансирование (по кредиту)', lambda R: R['financing'])
    r_oe = leaf('Прочие расходы', lambda R: R['other_expense'])
    r += 1

    # ---- НАЛОГООБЛАГАЕМЫЙ ДОХОД (БЕЗ 7212) ----
    row_pbt = total('НАЛОГООБЛАГАЕМЫЙ ДОХОД (до КПН)',
        formula=lambda col: (f'={col}{row_gross}-{col}{row_adm}'
                             f'+{col}{r_dep}+{col}{r_fxnet}+{col}{r_oi}'
                             f'-{col}{r_fin}-{col}{r_oe}'), fill='FCE4D6')
    row_kpn = total('КПН 20%', formula=lambda col: f'=MAX(0,{col}{row_pbt})*0.2', fill='FCE4D6')
    row_net = total('ЧИСТЫЙ ДОХОД', formula=lambda col: f'={col}{row_pbt}-{col}{row_kpn}', fill='C6E0B4')
    r += 1

    # ---- 7212 не идущие на вычет (ниже чистого дохода) ----
    row_7212 = None
    if lines_7212:
        section('РАСХОДЫ, НЕ ИДУЩИЕ НА ВЫЧЕТ (сч. 7212)')
        r7 = [leaf(a, lambda R, a=a: group_taxes(R['admin_7212']).get(a, 0.0)) for a in lines_7212]
        row_7212 = total('ИТОГО не идущие на вычет (7212)', rowrefs=r7, fill='BDD7EE'); r += 1

    # ---- КОНТРОЛЬ ---- (результат 1С включает 7212, поэтому сверяем P&L с учётом 7212)
    section('КОНТРОЛЬ СВЕРКИ С ОСВ 1С')
    r_i6 = leaf('Обороты 6ххх (доходы, Кт)', lambda R: R['tot_income_6'])
    r_e7 = leaf('Обороты 7ххх (расходы, Дт)', lambda R: R['tot_expense_7'])
    r_res = total('Результат 1С (6ххх − 7ххх)', formula=lambda col: f'={col}{r_i6}-{col}{r_e7}', fill=None)
    r_pl = total('P&L результат (с учётом 7212)',
        formula=lambda col: f'={col}{row_pbt}' + (f'-{col}{row_7212}' if row_7212 else ''), fill=None)
    total('Разница (P&L − 1С)', formula=lambda col: f'={col}{r_pl}-{col}{r_res}', fill='FFFF00')
    r += 1

    # ---- ИТОГОВАЯ СТРОКА: Чистый доход за вычетом не идущих на вычет ----
    total('ЧИСТЫЙ ДОХОД С УЧЁТОМ 7212 (= Чистый − 7212)',
        formula=lambda col: f'={col}{row_net}' + (f'-{col}{row_7212}' if row_7212 else ''), fill='C6E0B4')

    ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 52
    for col in dcols + [tcol]:
        ws.column_dimensions[col].width = 18
    ws.freeze_panes = 'C2'
    wb.save(out_path)
    return {'customers': len(customers), 'direct': len(direct_lines), 'admin': len(admin_lines)}


def build_journal(out_path, periods):
    wb = openpyxl.load_workbook(out_path)
    if 'Журнал сопоставления' in wb.sheetnames:
        del wb['Журнал сопоставления']
    js = wb.create_sheet('Журнал сопоставления')
    js.append(['Период', 'Раздел', 'Статья/судно в P&L', 'Сумма ₸'])
    for c in js[1]:
        c.font = Font(name=_F, size=10, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='2F5496'); c.border = _BORDER
    for nm, R in periods:
        for c, v in sorted(R['income'].items(), key=lambda x: -x[1]):
            js.append([nm, 'Доход (контрагент)', c, round(v, 2)])
        for a, v in group_taxes(R['direct']).items():
            js.append([nm, 'Прямые затраты', a, round(v, 2)])
        for a, v in group_taxes(R['admin']).items():
            js.append([nm, 'Административные (7211)', a, round(v, 2)])
        for a, v in group_taxes(R['admin_7212']).items():
            js.append([nm, '7212 (не на вычет)', a, round(v, 2)])
        for lbl, key in [('Депозиты', 'deposit'), ('Курс.доход', 'fx_income'),
                         ('Финансирование', 'financing'), ('Курс.расход', 'fx_expense')]:
            if R.get(key):
                js.append([nm, 'Прочие дох/расх', lbl, round(R[key], 2)])
        js.append([nm, 'КОНТРОЛЬ', 'Обороты 6ххх', round(R['tot_income_6'], 2)])
        js.append([nm, 'КОНТРОЛЬ', 'Обороты 7ххх', round(R['tot_expense_7'], 2)])
    for i, w in enumerate([12, 26, 48, 18], 1):
        js.column_dimensions[get_column_letter(i)].width = w
    for row in js.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=_F, size=9); c.border = _BORDER
        row[3].number_format = '#,##0'
    js.freeze_panes = 'A2'
    wb.save(out_path)
    return js.max_row - 1
