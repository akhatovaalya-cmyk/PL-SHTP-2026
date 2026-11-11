# -*- coding: utf-8 -*-
"""
P&L-агент ТОО «Штокпром» — вариант ПО ПРОЕКТАМ.
Структура 1С:
  6010 доход: вид услуги (ур.2) -> проект (ур.3, «NCOC: TAR-2026») -> Кредит (кол.6)
  7010 себестоимость: статья (ур.2) -> проект (ур.3, «NCOC»/«Рекадо»/...) -> Дебет (кол.5)
  7210/7211 админ: статья (ур.+1) -> Дебет ; 7212 — отдельный блок ниже чистого дохода.
  6250 курс.доход, 7430 курс.расход — нетто одной строкой в «Прочие».

P&L собирается блоками по проектам: для каждого проекта
    Доход -> Прямые расходы (по статьям) -> Операционная прибыль (= Доход − Прямые),
порядок NCOC -> NCOC TAR -> прочие. Дальше: административные, прочие дох/расх,
налогооблагаемый доход (без 7212), КПН 20 %, чистый доход, блок 7212, контроль сверки.

Проектная разметка в выгрузке разная: весь доход помечен «NCOC: TAR-2026» (→ NCOC TAR),
а затраты — «NCOC» / «Рекадо» / без проекта. Маппинг тегов в canon_project() —
правится одной строкой, если нужно перекинуть доход/статью в другой проект.
"""
import re, openpyxl
from collections import defaultdict, OrderedDict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

THRESHOLD = 300_000
_TAX_KEYS = ('обязательные пенсионные взносы', 'отчисления осмс',
             'социальные отчисления', 'социальный налог')
TAX_LINE = 'Налоги по заработной плате'
GENERIC = ('основная номенклатурная группа', '<...>', 'основное подразделение', '')

# Порядок блоков-проектов. Проекты не из списка добавляются после, по убыванию суммы.
PROJECT_ORDER = ['NCOC', 'NCOC TAR', 'Рекадо']
NOPROJECT = 'Основное подразделение'


def norm(s):
    return re.sub(r'\s+', ' ', str(s or '').strip().lower()).replace('ё', 'е')

def _num(x):
    return float(x) if isinstance(x, (int, float)) else 0.0


def canon_project(tag):
    """Проектный тег из 1С -> каноническое имя проекта.
    Доход размечен «NCOC: TAR-2026» (содержит TAR -> NCOC TAR),
    затраты — «NCOC», «Рекадо», либо общие (-> Без проекта).
    Чтобы перекинуть доход/статью в другой проект — поменяйте правило здесь."""
    n = norm(tag)
    if n in GENERIC:
        return NOPROJECT
    if 'tar' in n or 'тар' in n:
        return 'NCOC TAR'
    if 'рекадо' in n or 'rekado' in n:
        return 'Рекадо'
    if 'ncoc' in n:
        return 'NCOC'
    return str(tag).strip()


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
        'proj_income': defaultdict(float),                      # проект -> выручка (6010)
        'proj_direct': defaultdict(lambda: defaultdict(float)), # проект -> {статья -> сумма} (7010/7110)
        'admin': defaultdict(float),                            # статья -> сумма (7211)
        'admin_7212': defaultdict(float),                       # статья -> сумма (7212)
        'deposit': 0.0, 'fx_income': 0.0, 'other_income': 0.0,
        'financing': 0.0, 'fx_expense': 0.0, 'other_expense': 0.0,
        'tot_income_6': 0.0, 'tot_expense_7': 0.0,
        'income_target': 0.0, 'direct_target': 0.0,
        'admin_target': 0.0, 'admin7212_target': 0.0, 'admin_grand': 0.0,
    }
    acct = None; cur_service = None; cur_article = None
    code_od = {}
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
            # цели/низ — на строке самого счёта, независимо от уровня
            if acct == '6010':
                R['income_target'] += cred
            elif acct in ('7010', '7110'):
                R['direct_target'] += deb
            elif acct == '7200' and od == 0:
                R['admin_grand'] = R.get('admin_grand', 0) + deb
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
        # доход 6010: вид услуги (ур.+1) -> проект (лист). Выручка -> по проекту.
        if a == '6010':
            base = code_od.get('6010', 1)
            if od == base + 1:
                cur_service = name
            if is_leaf and cred:
                if od > base + 1:                 # лист = проект под видом услуги
                    proj = canon_project(name)
                else:                             # лист на уровне вида услуги (нет проекта)
                    proj = canon_project(cur_service) if norm(name) in GENERIC else canon_project(name)
                R['proj_income'][proj] += cred
            continue
        # прямые затраты 7010/7011/7110: статья (ур.+1) -> проект (лист).
        if a in ('7010', '7011', '7110', '7111'):
            base = max(code_od.get('7010', 1), code_od.get('7011', 0), code_od.get('7110', 0))
            if od == base + 1:
                cur_article = name
            if is_leaf and deb:
                if od > base + 1 and cur_article:   # лист = проект под статьёй
                    proj = canon_project(name)
                    art = cur_article
                else:                               # лист на уровне статьи (нет проекта)
                    proj = NOPROJECT
                    art = name
                R['proj_direct'][proj][art] += deb
            continue
        # админ (вычитаемый): статьи под 7210/7211 (кроме 7212)
        if a in ('7210', '7211'):
            if od == code_od.get(a, 1) + 1 and deb:
                R['admin'][name] += deb
            continue
        if a == '7212':
            if od == code_od.get('7212', 1) + 1 and deb:
                R['admin_7212'][name] += deb
            continue

    # ---- сверка остатков по целям ----
    # доход: сумма по проектам должна = обороту 6010
    inc_sum = sum(R['proj_income'].values())
    resid_i = R['income_target'] - inc_sum
    if abs(resid_i) > 1:
        R['proj_income'][NOPROJECT] += resid_i
    # прямые: сумма по всем проектам должна = обороту 7010/7110
    dir_sum = sum(v for d in R['proj_direct'].values() for v in d.values())
    resid_d = R['direct_target'] - dir_sum
    if abs(resid_d) > 1:
        R['proj_direct'][NOPROJECT]['Прочие (свёрнуто <...> в 1С)'] += resid_d
    # админ
    R['admin_target'] = R.get('admin_grand', 0) - R['admin7212_target']
    for key, tgt in [('admin', 'admin_target'), ('admin_7212', 'admin7212_target')]:
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


def _project_list(periods):
    """Упорядоченный список проектов (по PROJECT_ORDER, прочие — по убыванию суммы)."""
    tot = defaultdict(float)
    for _, R in periods:
        for p, v in R['proj_income'].items():
            tot[p] += v
        for p, d in R['proj_direct'].items():
            tot[p] += sum(d.values())
    present = set(tot.keys())
    ordered = [p for p in PROJECT_ORDER if p in present]
    rest = sorted((p for p in present if p not in PROJECT_ORDER), key=lambda p: -tot[p])
    return ordered + rest


def _project_articles(periods, proj):
    """Статьи прямых расходов проекта (по убыванию суммы)."""
    agg = defaultdict(float)
    for _, R in periods:
        for art, v in R['proj_direct'].get(proj, {}).items():
            agg[art] += v
    return [a for a, _ in sorted(agg.items(), key=lambda x: -x[1])]


def build_workbook(periods, out_path, company='ТОО «Штокпром»'):
    projects = _project_list(periods)

    admin_raw = defaultdict(float); a7212_raw = defaultdict(float)
    for _, R in periods:
        for c, v in R['admin'].items():
            admin_raw[c] += v
        for c, v in R['admin_7212'].items():
            a7212_raw[c] += v
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

    def section(title, color='2F5496'):
        nonlocal r
        cell = ws.cell(row=r, column=2, value=title); cell.font = Font(name=_F, size=10, bold=True, color='FFFFFF')
        for cc in range(2, 4 + ncol):
            ws.cell(row=r, column=cc).fill = PatternFill('solid', fgColor=color); ws.cell(row=r, column=cc).border = _BORDER
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

    # ---- ДОХОД ОТ РЕАЛИЗАЦИИ (сверху, одним блоком) ----
    section('ДОХОД ОТ РЕАЛИЗАЦИИ')
    inc_rows = [leaf(p, lambda R, p=p: R['proj_income'].get(p, 0.0)) for p in projects]
    row_inc = total('ИТОГО ДОХОД ОТ РЕАЛИЗАЦИИ', rowrefs=inc_rows, fill='BDD7EE'); r += 1

    # ---- ПРЯМЫЕ РАСХОДЫ (подблоками по проектам) ----
    section('ПРЯМЫЕ РАСХОДЫ (по проектам)')
    dir_projs = [p for p in projects if _project_articles(periods, p)]
    dsub_rows = []
    for proj in dir_projs:
        sc = ws.cell(row=r, column=2, value=f'  Проект «{proj}»')
        sc.font = Font(name=_F, size=10, bold=True, italic=True, color='375623'); sc.border = _BORDER
        for cc in range(2, 4 + ncol):
            ws.cell(row=r, column=cc).fill = PatternFill('solid', fgColor='E2EFDA'); ws.cell(row=r, column=cc).border = _BORDER
        r += 1
        arts = _project_articles(periods, proj)
        art_rows = [leaf(a, lambda R, p=proj, a=a: R['proj_direct'].get(p, {}).get(a, 0.0), indent=2)
                    for a in arts]
        r_sub = total(f'Итого «{proj}»', rowrefs=art_rows, fill='E2EFDA')
        dsub_rows.append(r_sub)
    row_dir = total('ИТОГО ПРЯМЫЕ РАСХОДЫ', rowrefs=dsub_rows, fill='BDD7EE')
    row_gross = total('ВАЛОВАЯ (ОПЕРАЦИОННАЯ) ПРИБЫЛЬ',
                      formula=lambda col: f'={col}{row_inc}-{col}{row_dir}', fill='FFF2CC'); r += 1

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

    # ---- КОНТРОЛЬ ---- (результат 1С включает 7212)
    section('КОНТРОЛЬ СВЕРКИ С ОСВ 1С')
    r_i6 = leaf('Обороты 6ххх (доходы, Кт)', lambda R: R['tot_income_6'])
    r_e7 = leaf('Обороты 7ххх (расходы, Дт)', lambda R: R['tot_expense_7'])
    r_res = total('Результат 1С (6ххх − 7ххх)', formula=lambda col: f'={col}{r_i6}-{col}{r_e7}', fill=None)
    r_pl = total('P&L результат (с учётом 7212)',
        formula=lambda col: f'={col}{row_pbt}' + (f'-{col}{row_7212}' if row_7212 else ''), fill=None)
    total('Разница (P&L − 1С)', formula=lambda col: f'={col}{r_pl}-{col}{r_res}', fill='FFFF00')
    r += 1

    # ---- ИТОГОВАЯ СТРОКА ----
    total('ЧИСТЫЙ ДОХОД С УЧЁТОМ 7212 (= Чистый − 7212)',
        formula=lambda col: f'={col}{row_net}' + (f'-{col}{row_7212}' if row_7212 else ''), fill='C6E0B4')

    ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 54
    for col in dcols + [tcol]:
        ws.column_dimensions[col].width = 18
    ws.freeze_panes = 'C2'
    wb.save(out_path)
    return {'projects': len(projects), 'admin': len(admin_lines)}


def build_journal(out_path, periods):
    wb = openpyxl.load_workbook(out_path)
    if 'Журнал сопоставления' in wb.sheetnames:
        del wb['Журнал сопоставления']
    js = wb.create_sheet('Журнал сопоставления')
    js.append(['Период', 'Раздел', 'Проект', 'Статья в P&L', 'Сумма ₸'])
    for c in js[1]:
        c.font = Font(name=_F, size=10, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='2F5496'); c.border = _BORDER
    for nm, R in periods:
        for p, v in sorted(R['proj_income'].items(), key=lambda x: -x[1]):
            js.append([nm, 'Доход (проект)', p, '—', round(v, 2)])
        for p in _project_list([(nm, R)]):
            for a, v in sorted(R['proj_direct'].get(p, {}).items(), key=lambda x: -x[1]):
                js.append([nm, 'Прямые затраты', p, a, round(v, 2)])
        for a, v in group_taxes(R['admin']).items():
            js.append([nm, 'Административные (7211)', '—', a, round(v, 2)])
        for a, v in group_taxes(R['admin_7212']).items():
            js.append([nm, '7212 (не на вычет)', '—', a, round(v, 2)])
        for lbl, key in [('Депозиты', 'deposit'), ('Курс.доход', 'fx_income'),
                         ('Финансирование', 'financing'), ('Курс.расход', 'fx_expense')]:
            if R.get(key):
                js.append([nm, 'Прочие дох/расх', '—', lbl, round(R[key], 2)])
        js.append([nm, 'КОНТРОЛЬ', '—', 'Обороты 6ххх', round(R['tot_income_6'], 2)])
        js.append([nm, 'КОНТРОЛЬ', '—', 'Обороты 7ххх', round(R['tot_expense_7'], 2)])
    for i, w in enumerate([12, 26, 16, 44, 18], 1):
        js.column_dimensions[get_column_letter(i)].width = w
    for row in js.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=_F, size=9); c.border = _BORDER
        row[4].number_format = '#,##0'
    js.freeze_panes = 'A2'
    wb.save(out_path)
    return js.max_row - 1
