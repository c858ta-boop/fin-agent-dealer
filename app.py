import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Анализ влияния первичных статей расходов на общий бюджет.")

# Панель настроек в боковой панели
with st.sidebar:
    st.header("⚙️ Настройки анализа")
    target_column = st.text_input("Название столбца со статьями расходов:", value="Статья расходов")
    value_column = st.text_input("Название столбца со значениями (суммами):", value="Всего расходы")
    header_row = st.number_input("Строка с заголовками (в Excel нумерация с 1):", min_value=1, value=2)
    total_row_name = st.text_input("Название строки общего итога:", value="Всего по ДЦ")

# Блок загрузки файлов
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 Загрузите СТАРЫЙ отчет (прошлый месяц)", type=["xlsx"])
with col2:
    new_file = st.file_uploader("📂 Загрузите НОВЫЙ отчет (текущий месяц)", type=["xlsx"])

def is_colored(cell):
    if cell and cell.fill and cell.fill.fill_type:
        color = cell.fill.start_color.index
        if color and str(color) not in ['00000000', '0', 'FFFFFFFF', 'System_Color_Window']:
            return True
    return False

if old_file and new_file:
    st.success("Файлы успешно загружены! Запускаю диагностику структуры...")
    
    old_bytes = old_file.read()
    new_bytes = new_file.read()
    
    xl_old = pd.ExcelFile(BytesIO(old_bytes))
    xl_new = pd.ExcelFile(BytesIO(new_bytes))
    common_sheets = list(set(xl_old.sheet_names).intersection(set(xl_new.sheet_names)))
    
    if not common_sheets:
        st.error("❌ Ошибка: В файлах нет листов с одинаковыми названиями!")
    else:
        all_expenses_changes = []
        total_old_dc = 0.0
        total_new_dc = 0.0
        total_row_found = False
        header_idx = int(header_row)
        pandas_header_index = header_idx - 1
        
        clean_total_target = str(total_row_name).strip().lower().replace('c', 'с').replace('x', 'х')
        
        for sheet_name in common_sheets:
            # Читаем данные через pandas
            df_old = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet_name, header=pandas_header_index)
            df_new = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_name, header=pandas_header_index)
            
            # ВЫВОД ДИАГНОСТИКИ НА ЭКРАН
            st.warning(f"🔍 ДИАГНОСТИКА СТРУКТУРЫ ДЛЯ ЛИСТА: '{sheet_name}'")
            st.write("**Найденные столбцы в вашем файле:**", list(df_new.columns))
            st.write("**Первые 5 строк из файла (для проверки):**")
            st.dataframe(df_new.head(5), use_container_width=True)
            
            df_old.columns = [str(c).strip() for c in df_old.columns]
            df_new.columns = [str(c).strip() for c in df_new.columns]
            
            if target_column in df_new.columns and value_column in df_new.columns:
                # Извлекаем общий итог
                for i, row in df_old.iterrows():
                    if row[target_column] is not None:
                        cell_clean_text = str(row[target_column]).strip().lower().replace('c', 'с').replace('x', 'х')
                        if clean_total_target == cell_clean_text or clean_total_target in cell_clean_text or ("всего" in cell_clean_text and "дц" in cell_clean_text):
                            try: total_old_dc += float(row[value_column])
                            except: pass
                            total_row_found = True
                            
                for i, row in df_new.iterrows():
                    if row[target_column] is not None:
                        cell_clean_text = str(row[target_column]).strip().lower().replace('c', 'с').replace('x', 'х')
                        if clean_total_target == cell_clean_text or clean_total_target in cell_clean_text or ("всего" in cell_clean_text and "дц" in cell_clean_text):
                            try: total_new_dc += float(row[value_column])
                            except: pass
            else:
                st.error(f"❌ Столбцы '{target_column}' или '{value_column}' не найдены среди колонок листа!")
        
        # Расчет дельты по ДЦ
        dc_delta = total_new_dc - total_old_dc
        st.subheader("📊 Текущий результат подсчета по ДЦ")
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Расходы за прошлый месяц", f"{total_old_dc:,.2f} руб.")
        with c2: st.metric("Расходы за текущий месяц", f"{total_new_dc:,.2f} руб.")
        with c3: st.metric("Общее изменение расходов ДЦ", f"{dc_delta:+,.2f} руб.", delta_color="inverse")
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для запуска диагностики.")
