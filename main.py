import streamlit as st
import itertools
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="FastOcenka prod", layout="wide")
st.markdown("<h3 style='text-align: left;'>FastOcenka prod</h3>", unsafe_allow_html=True)

MAX_COUNT = 10

with st.expander("Инструкция"):
    st.write(f"""
    - Введите **от 2 до {MAX_COUNT}** чисел через запятую, точку с запятой, двоеточие, пробел или с новой строки.
    - Допустимы как целые, так и дробные числа (через точку или запятую).
    - Будут рассчитаны все возможные комбинации средних для групп по 2, 3, 4 ... до N чисел.
    - Результаты можно выгрузить в Excel (все комбинации — на одном листе).
    - Используйте кнопку **"Подогнать"** для расчёта нового значения любого элемента, чтобы среднее выбранной группы стало равным нужному вам.
    """)

input_data = st.text_area(
    "Введите числа:",
    "10.2; 10.47\n17.5 8.5"
)

def to_excel_onelayer(tables):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        ws_name = "Комбинации"
        startrow = 0
        workbook = writer.book
        worksheet = writer.book.add_worksheet(ws_name)
        writer.sheets[ws_name] = worksheet
        for title, df in tables:
            worksheet.write(startrow, 0, title)
            df.to_excel(writer, index=False, sheet_name=ws_name, startrow=startrow + 1)
            startrow += len(df) + 3  # +3 строки между таблицами
        # Автоширина столбцов
        if tables:
            df = tables[-1][1]
            for col_idx, col in enumerate(df.columns):
                maxlen = max([len(str(val)) for table in tables for val in table[1][col].astype(str)] + [len(str(col))])
                worksheet.set_column(col_idx, col_idx, min(30, maxlen + 3))
    output.seek(0)
    return output

# 1. Кнопка расчёта — только заполняет session_state
if st.button("Рассчитать"):
    numbers = []
    for s in re.split(r'[,;:\n\s]+', input_data):
        s = s.replace(',', '.').strip()
        if not s: continue
        try:
            numbers.append(float(s))
        except Exception:
            continue

    if len(numbers) < 2:
        st.warning("Введите как минимум два числа.")
        st.stop()
    if len(numbers) > MAX_COUNT:
        st.warning(f"Максимально поддерживается {MAX_COUNT} чисел для расчёта.")
        st.stop()

    excel_tables = []
    all_combos = []

    for k in range(2, len(numbers)+1):
        comb_list = list(itertools.combinations(enumerate(numbers, 1), k))
        res = []
        for comb in comb_list:
            inds = [i for i, _ in comb]
            combo_name = "".join(str(i) for i in inds)
            vals = [v for _, v in comb]
            mean_val = round(sum(vals) / k, 4)
            res.append({'Комбинация индексов': combo_name, 'Среднее': mean_val, 'Значения': str(vals), 'Индексы': str(inds)})
            all_combos.append({'k': k, 'inds': inds, 'name': combo_name, 'vals': vals, 'mean': mean_val})
        df = pd.DataFrame(res)
        excel_tables.append((f"Средние для комбинаций по {k} значениям", df))

    # Сохраняем в session_state
    st.session_state.numbers = numbers
    st.session_state.excel_tables = excel_tables
    st.session_state.all_combos = all_combos

# 2. ВСЕГДА показываем результаты, если они уже посчитаны
if 'numbers' in st.session_state and st.session_state.numbers:
    st.write("**Введённые значения:**")
    st.dataframe(
        pd.DataFrame({'№': range(1, len(st.session_state.numbers)+1), 'Значение': st.session_state.numbers}),
        use_container_width=True
    )

if 'excel_tables' in st.session_state and st.session_state.excel_tables:
    for title, df in st.session_state.excel_tables:
        st.write(f"**{title}:**")
        st.dataframe(df[['Комбинация индексов', 'Среднее']], use_container_width=True)

    # Кнопка выгрузки Excel
    excel_data = to_excel_onelayer(st.session_state.excel_tables)
    st.download_button(
        label="Скачать все таблицы в Excel",
        data=excel_data,
        file_name="combinations_mean.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- Диалоговая подгонка через @st.dialog ---
@st.dialog("Подгонка среднего значения для любого элемента")
def fit_dialog():
    all_combos = st.session_state.all_combos
    options = [
        f"k={c['k']} | индексы: {c['name']} | среднее: {c['mean']} | значения: {c['vals']}"
        for c in all_combos
    ]
    selected = st.selectbox("Комбинация для подгонки", options=options, key="fit_combo")
    combo_idx = options.index(selected)
    combo = all_combos[combo_idx]
    elem_names = [f"Позиция {i+1} (№{ind}) = {val}" for i, (ind, val) in enumerate(zip(combo['inds'], combo['vals']))]
    elem_to_fit = st.selectbox("Выберите элемент для подгонки", options=elem_names, key="fit_elem")
    elem_idx = elem_names.index(elem_to_fit)
    desired = st.number_input("Желаемое среднее значение", value=combo['mean'], step=0.01, format="%.4f", key="desired_mean")
    if st.button("Выполнить подгонку", key="do_fit"):
        k = combo['k']
        vals = combo['vals']
        inds = combo['inds']
        sum_others = sum(v for idx, v in enumerate(vals) if idx != elem_idx)
        new_val = round(desired * k - sum_others, 4)
        new_vals = list(vals)
        new_vals[elem_idx] = new_val
        st.success(
            f"Чтобы среднее стало {desired}, нужно изменить элемент на позиции {elem_idx+1} (№{inds[elem_idx]}) на **{new_val}**"
        )
        st.write(f"Новый набор: {new_vals}")
        st.write(f"Новое среднее: {round(sum(new_vals)/k, 4)}")
        st.write(f"Оригинальный набор: {vals}")

# --- Кнопка запуска подгонки ---
st.markdown("---")
if st.button("Подогнать"):
    if 'all_combos' not in st.session_state or not st.session_state.all_combos:
        st.warning("Сначала выполните расчёт.")
        st.stop()
    fit_dialog()
