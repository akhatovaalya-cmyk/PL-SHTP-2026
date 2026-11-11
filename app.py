# -*- coding: utf-8 -*-
"""P&L-агент ТОО «ТОО «Штокпром»» (ОСГ). Загрузка 1С ОСВ → P&L."""
import io, re, tempfile
import streamlit as st
import openpyxl
import shtok_engine as eng
import recalc_native

st.set_page_config(page_title='P&L агент — Штокпром', page_icon='📊', layout='wide')
st.title('📊 P&L агент — ТОО «Штокпром» (по проектам)')
st.success('✅ ВЕРСИЯ 2026-11-11 · блоки по проектам (NCOC → NCOC TAR → прочие) · операционная прибыль на проект · 7212 отдельно · курсовая нетто')
st.caption('Загрузите выгрузку 1С ОСВ. Берутся счета 6ххх (доходы) и 7ххх (расходы): '
           'доход и прямые затраты раскладываются по проектам (проектная аналитика 1С), '
           'на каждый проект — Доход → Прямые расходы по статьям → Операционная прибыль. '
           'Ниже: административные (7211 расшифрован, налоги с ЗП свёрнуты), 7212 отдельным блоком '
           'ниже чистого дохода, курсовая разница нетто. Всё сходится с 1С.')

MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь']

def guess(fn):
    s = fn.lower()
    for st_, m in [('январ','Январь'),('феврал','Февраль'),('март','Март'),('апрел','Апрель'),('мая','Май'),('май','Май'),
                   ('июн','Июнь'),('июл','Июль'),('август','Август'),('сентябр','Сентябрь'),('октябр','Октябрь'),('ноябр','Ноябрь'),('декабр','Декабрь')]:
        if st_ in s:
            return m
    y = re.search(r'(20\d\d)', s); return y.group(1) if y else ''

files = st.file_uploader('Выгрузки 1С ОСВ (.xlsx) — можно несколько', type=['xlsx'], accept_multiple_files=True)
labels = []
if files:
    st.write('**Названия колонок:**')
    for i, uf in enumerate(files):
        c1, c2 = st.columns([3, 2]); c1.write(f'📄 {uf.name}')
        labels.append(c2.text_input('Колонка', value=guess(uf.name) or f'Период {i+1}', key=f'l{i}', label_visibility='collapsed'))

st.divider()
if st.button('▶️ Сформировать P&L', type='primary', disabled=not files):
    try:
        periods = []
        for uf, lbl in zip(files, labels):
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
            with open(tmp, 'wb') as f:
                f.write(uf.getvalue())
            R = eng.parse_1c(tmp); periods.append((lbl or 'Период', R))
            st.info(f'✅ {uf.name} → «{lbl}»: доход {R["tot_income_6"]:,.0f}  расход {R["tot_expense_7"]:,.0f}  '
                    f'до КПН {R["tot_income_6"]-R["tot_expense_7"]:,.0f}')
        out = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
        eng.build_workbook(periods, out); eng.build_journal(out, periods); recalc_native.recalc_native(out)
        with open(out, 'rb') as f:
            data = f.read()
        st.success('Готово! P&L собран и сходится с 1С.')
        st.download_button('⬇️ Скачать P&L (.xlsx)', data=data, file_name='Shtokprom_P_L.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        st.error(f'Ошибка: {e}')
        import traceback; st.code(traceback.format_exc())

with st.expander('ℹ️ Как работает разноска'):
    st.markdown('''
- **Проекты.** Доход (6010) и прямые затраты (7010/7110) берутся с проектной
  аналитикой 1С и собираются блоками: на каждый проект — **Доход → Прямые
  расходы по статьям → Операционная прибыль** (= Доход − Прямые). Порядок блоков:
  **NCOC → NCOC TAR → прочие**.
- **Разметка проектов.** Теги 1С сводятся к имени проекта в функции `canon_project`
  (в `shtok_engine.py`): «NCOC: TAR-2026» → *NCOC TAR*, «NCOC» → *NCOC*, «Рекадо» →
  *Рекадо*, общие («Основная номенклатурная группа» и свёрнутые) → *Без проекта*.
  Чтобы перекинуть доход/статью в другой блок — правится одной строкой здесь.
- **Административные** (7211) — расшифровка по статьям; налоги с ЗП свёрнуты.
- **7212** — отдельный блок ниже чистого дохода (не уменьшает налогооблагаемый доход).
- **Курсовая разница** (6250/7430) — одной нетто-строкой (+ доход / − расход).
- Налоги с ЗП (ОПВ+ОСМС+соц.отчисления+соц.налог) — одной строкой «Налоги по заработной плате».
- Внизу — **контроль сверки** с оборотами 6ххх/7ххх 1С, «Разница» = 0.
''')
