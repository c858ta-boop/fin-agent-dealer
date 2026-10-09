import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Анализ влияния первичных статей расходов (исключая цветные суммирующие строки отделов) на общий бюджет.")

# Панель настроек в боковой панели
with st.sidebar:
    st.header("⚙️ Настройки анализа")
    target_column = st.text_input("Название столбца со статьями расходов:", value="Статья расходов")
    value_column = st.text_input("Название столбца со значениями (суммами):", value="Всего расходы")
    header_row = st.number_input("Строка с заголовками (в Excel нумерация с 1):", min_value=1, value=2)
    total_row_name = st.text_input("Название строки общего итога:", value="Всего по ДЦ")
    st.caption("ℹ️ Алгоритм автоматически исключает из ТОП-10 строки, имеющие цветовую заливку.")

# Блок загрузки файлов
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 Загрузите СТАРЫЙ отчет (прошлый месяц)", type=["xlsx"])
with col2:
    new_file = st.file_uploader("📂 Загрузите НОВЫЙ отчет (текущий месяц)", type=["xlsx"])

def is_colored(cell):
    """Проверяет, есть ли у ячейки цветная заливка (игнорирует белый/прозрачный)"""
    if cell and cell.fill and cell.fill.fill_type:
        color = cell.fill.start_color.index
        if color and str(color) not in ['00000000', '0', 'FFFFFFFF', 'System_Color_Window']:
            return True
    return False

# Основная логика приложения
if old_file and new_file:
    st.success("Файлы успешно загружены! Начинаю факторный анализ...")
    
    old_bytes = old_file.read()
    new_bytes = new_file.read()
    
    wb_old = openpyxl.load_workbook(BytesIO(old_bytes), data_only=True, read_only=True)
    wb_new = openpyxl.load_workbook(BytesIO(new_bytes), data_only=True, read_only=True)
    
    common_sheets = list(set(wb_old.sheetnames).intersection(set(wb_new.sheetnames)))
    
    if not common_sheets:
        st.error("❌ Ошибка: В файлах нет листов с одинаковыми названиями!")
    else:
        all_expenses_changes = []
        total_old_dc = 0.0
        total_new_dc = 0.0
        total_row_found = False
        header_idx = int(header_row)
        pandas_header_index = header_idx - 1
        
        # Фиксируем искомое название итога в нижнем регистре без лишних пробелов
        clean_total_target = str(total_row_name).strip().lower()
        
        # Сканируем каждый общий лист
        for sheet_name in common_sheets:
            # Находим цветные строки через openpyxl на этом листе
            wb_old_full = openpyxl.load_workbook(BytesIO(old_bytes), data_only=True)
            ws_old_xl = wb_old_full[sheet_name]
            target_col_idx_old = None
            for col in range(1, ws_old_xl.max_column + 1):
                if str(ws_old_xl.cell(row=header_idx, column=col).value).strip() == target_column:
                    target_col_idx_old = col
                    break
            colored_old_rows = set()
            if target_col_idx_old:
                for r_idx in range(header_idx + 1, ws_old_xl.max_row + 1):
                    cell = ws_old_xl.cell(row=r_idx, column=target_col_idx_old)
                    if cell and cell.value is not None and is_colored(cell):
                        colored_old_rows.add(r_idx - header_idx - 1)
                        
            wb_new_full = openpyxl.load_workbook(BytesIO(new_bytes), data_only=True)
            ws_new_xl = wb_new_full[sheet_name]
            target_col_idx_new = None
            for col in range(1, ws_new_xl.max_column + 1):
                if str(ws_new_xl.cell(row=header_idx, column=col).value).strip() == target_column:
                    target_col_idx_new = col
                    break
            colored_new_rows = set()
            if target_col_idx_new:
                for r_idx in range(header_idx + 1, ws_new_xl.max_row + 1):
                    cell = ws_new_xl.cell(row=r_idx, column=target_col_idx_new)
                    if cell and cell.value is not None and is_colored(cell):
                        colored_new_rows.add(r_idx - header_idx - 1)

            # Читаем данные через pandas для математических расчетов
            df_old = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet_name, header=pandas_header_index)
            df_new = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_name, header=pandas_header_index)
            
            df_old.columns = [str(c).strip() for c in df_old.columns]
            df_new.columns = [str(c).strip() for c in df_new.columns]
            
            if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                
                # УМНЫЙ ПОИСК ИТОГА: убираем жесткое равенство, ищем совпадение текста с очисткой пробелов
                for i, row in df_old.iterrows():
                    if row[target_column] is not None:
                        cell_clean_text = str(row[target_column]).strip().lower()
                        if clean_total_target == cell_clean_text or clean_total_target in cell_clean_text:
                            try: total_old_dc += float(row[value_column])
                            except: pass
                            total_row_found = True
                            
                for i, row in df_new.iterrows():
                    if row[target_column] is not None:
                        cell_clean_text = str(row[target_column]).strip().lower()
                        if clean_total_target == cell_clean_text or clean_total_target in cell_clean_text:
                            try: total_new_dc += float(row[value_column])
                            except: pass
                
                # Отфильтровываем цветные суммирующие строки отделов
                df_old_clean = df_old.drop(index=list(colored_old_rows), errors='ignore').dropna(subset=[target_column, value_column])
                df_new_clean = df_new.drop(index=list(colored_new_rows), errors='ignore').dropna(subset=[target_column, value_column])
                
                dict_old = pd.Series(df_old_clean[value_column].values, index=df_old_clean[target_column]).to_dict()
                dict_new = pd.Series(df_new_clean[value_column].values, index=df_new_clean[target_column]).to_dict()
                
                sheet_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                for article in sheet_articles:
                    article_str = str(article).strip()
                    if article_str == "" or any(word in article_str.lower() for word in ["итого", "всего", "баланс", "результат", "свод"]):
                        continue
                    try: val_old = float(dict_old.get(article, 0) or 0)
                    except: val_old = 0.0
                    try: val_new = float(dict_new.get(article, 0) or 0)
                    except: val_new = 0.0
                    
                    item_delta = val_new - val_old
                    abs_delta = abs(item_delta)
                    if abs_delta > 0:
                        all_expenses_changes.append({
                            "Лист": sheet_name,
                            "Статья расходов": article_str,
                            "Было (руб.)": val_old,
                            "Стало (руб.)": val_new,
                            "Изменение (руб.)": item_delta,
                            "sort_key": abs_delta
                        })
        
        # Расчет дельты по ДЦ
        dc_delta = total_new_dc - total_old_dc
        st.subheader("📊 Общий финансовый результат по ДЦ")
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Расходы за прошлый месяц", f"{total_old_dc:,.2f} руб.")
        with c2: st.metric("Расходы за текущий месяц", f"{total_new_dc:,.2f} руб.")
        with c3: st.metric("Общее изменение расходов ДЦ", f"{dc_delta:+,.2f} руб.", delta_color="inverse")
            
        if not total_row_found:
            st.warning(f"⚠️ Строка '{total_row_name}' не найдена в файлах.")
        
        if all_expenses_changes:
            df_total_changes = pd.DataFrame(all_expenses_changes)
            top_10_changes = df_total_changes.sort_values(by="sort_key", ascending=False).head(10)
            
            base_denom = total_old_dc if total_old_dc > 0 else 1.0
            top_10_changes["Доля во влиянии на общую разницу"] = top_10_changes.apply(lambda r: f"{r['Изменение (руб.)'] / base_denom * 100:+.2f}%", axis=1)
            
            top_10_display = top_10_changes.drop(columns=["sort_key"]).reset_index(drop=True)
            top_10_display.index = top_10_display.index + 1
            
            st.subheader("📋 Директорский отчет: ТОП-10 чистых статей расходов")
            st.write("Суммирующие строки отделов отфильтрованы по цвету заливки ячеек.")
            st.dataframe(top_10_display, use_container_width=True)
            
            st.write("---")
            st.subheader("🖨️ Печать и экспорт в PDF")
            st.write("Нажмите комбинацию клавиш **Ctrl + P** (или **Cmd + P** на Mac) прямо на этой странице браузера, чтобы сохранить этот отчет в PDF.")
            
            html_preview = "<html><head><meta charset='utf-8'><style>"
            html_preview += "body { font-family: Arial, sans-serif; padding: 20px; color: #333; }"
            html_preview += "h2 { color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; font-size: 18px; margin-top:0; }"
            html_preview += "table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 11px; }"
            html_preview += "th { background: #1E3A8A; color: white; padding: 6px; text-align: left; }"
            html_preview += "td { padding: 6px; border-bottom: 1px solid #E5E7EB; }"
            html_preview += "</style></head><body><div style='background: white;'>"
