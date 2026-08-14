import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Инструмент выявления ТОП-10 изменений в статьях расходов (исключая цветные суммирующие строки отделов).")

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
    """Проверяет, есть ли у ячейки цветная заливка"""
    if cell and cell.fill and cell.fill.fill_type:
        color = cell.fill.start_color.index
        if color and str(color) not in ['00000000', '0', 'FFFFFFFF', 'System_Color_Window']:
            return True
    return False

def get_colored_rows(file_bytes, sheet_name, header_idx, target_col_name):
    """Быстро находит строки с цветовой заливкой с защитой от пустых ячеек"""
    colored_rows = set()
    try:
        # Убираем read_only для более стабильной обработки сложных ячеек Excel
        wb = openpyxl.load_workbook(file_bytes, data_only=True)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            target_col_idx = None
            
            # Читаем строку заголовков
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=header_idx, column=col).value
                if cell_val is not None and str(cell_val).strip() == target_col_name:
                    target_col_idx = col
                    break
                    
            # Если столбец найден, проверяем только его ячейки построчно
            if target_col_idx:
                for row_idx in range(header_idx + 1, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=target_col_idx)
                    # Если ячейка не пустая и покрашена — запоминаем индекс строки
                    if cell and cell.value is not None and is_colored(cell):
                        colored_rows.add(row_idx - header_idx - 1)
    except:
        pass
    return colored_rows

# Запуск основного интерфейса только при наличии обоих файлов
if old_file and new_file:
    st.success("Файлы успешно загружены! Начинаю факторный анализ...")
    
    # Клонируем файлы в независимые буферы памяти для стабильного чтения
    old_bytes = old_file.read()
    new_bytes = new_file.read()
    
    # Считываем книги
    wb_old = openpyxl.load_workbook(BytesIO(old_bytes), data_only=True, read_only=True)
    wb_new = openpyxl.load_workbook(BytesIO(new_bytes), data_only=True, read_only=True)
    
    common_sheets = list(set(wb_old.sheetnames).intersection(set(wb_new.sheetnames)))
    
    if not common_sheets:
        st.error("❌ Ошибка: В файлах нет листов с одинаковыми названиями!")
    
    if common_sheets:
        all_expenses_changes = []
        total_old_dc = 0.0
        total_new_dc = 0.0
        total_row_found = False
        header_idx = int(header_row)
        pandas_header_index = header_idx - 1
        
        # Сканируем каждый общий лист
        for sheet_name in common_sheets:
            # Сначала через защищенную функцию вычисляем цветные строки
            colored_old_rows = get_colored_rows(BytesIO(old_bytes), sheet_name, header_idx, target_column)
            colored_new_rows = get_colored_rows(BytesIO(new_bytes), sheet_name, header_idx, target_column)
            
            # Читаем данные через pandas для математических расчетов
            df_old = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet_name, header=pandas_header_index)
            df_new = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_name, header=pandas_header_index)
            
            df_old.columns = [str(c).strip() for c in df_old.columns]
            df_new.columns = [str(c).strip() for c in df_new.columns]
            
            # Проверяем наличие целевых столбцов на листах
            if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                
                # Фиксируем общий итог "Всего по ДЦ" до очистки цвета
                for i, row in df_old.iterrows():
                    if row[target_column] is not None and str(row[target_column]).strip().lower() == total_row_name.lower().strip():
                        try: total_old_dc += float(row[value_column])
                        except: pass
                        total_row_found = True
                for i, row in df_new.iterrows():
                    if row[target_column] is not None and str(row[target_column]).strip().lower() == total_row_name.lower().strip():
                        try: total_new_dc += float(row[value_column])
                        except: pass
                
                # Выкидываем цветные строки отделов и пустые ячейки
                df_old_clean = df_old.drop(index=list(colored_old_rows), errors='ignore').dropna(subset=[target_column, value_column])
                df_new_clean = df_new.drop(index=list(colored_new_rows), errors='ignore').dropna(subset=[target_column, value_column])
                
                dict_old = pd.Series(df_old_clean[value_column].values, index=df_old_clean[target_column]).to_dict()
                dict_new = pd.Series(df_new_clean[value_column].values, index=df_new_clean[target_column]).to_dict()
                
                # Рассчитываем изменения по статьям расходов
                sheet_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                for article in sheet_articles:
                    article_str = str(article).strip()
                    if article_str == "" or any(word in article_str.lower() for word in ["итого", "всего", "баланс", "результат", "свод"]):
                        continue
                    try: val_old = float(dict_old.get(article, 0))
                    except: val_old = 0.0
                    try: val_new = float(dict_new.get(article, 0))
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
        
        # Вывод финансового результата по ДЦ
        dc_delta = total_new_dc - total_old_dc
        st.subheader("📊 Общий финансовый результат по ДЦ")
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Расходы за прошлый месяц", f"{total_old_dc:,.2f} руб.")
        with c2: st.metric("Расходы за текущий месяц", f"{total_new_dc:,.2f} руб.")
        with c3: st.metric("Общее изменение расходов ДЦ", f"{dc_delta:+,.2f} руб.", delta_color="inverse")
            
        if not total_row_found:
            st.warning(f"⚠️ Строка '{total_row_name}' не найдена в файлах.")
        
        # Построение и отображение ТОП-10 изменений расходов
        if all_expenses_changes:
            df_total_changes = pd.DataFrame(all_expenses_changes)
            top_10_changes = df_total_changes.sort_values(by="sort_key", ascending=False).head(10)
            
            # Рассчитываем процент строго в одну строку
            base_denom = total_old_dc if total_old_dc > 0 else 1.0
            top_10_changes["Доля во влиянии на общую разницу"] = top_10_changes.apply(lambda r: f"{r['Изменение (руб.)'] / base_denom * 100:+.2f}%", axis=1)
            
            top_10_display = top_10_changes.drop(columns=["sort_key"]).reset_index(drop=True)
            top_10_display.index = top_10_display.index + 1
            
            st.subheader("📋 Директорский отчет: ТОП-10 чистых статей расходов")
            st.write("Промежуточные итоги отделов успешно отфильтрованы по цвету заливки ячеек.")
            st.dataframe(top_10_display, use_container_width=True)
            
            # Генерация HTML-версии отчета для сохранения в PDF
            st.write("---")
            st.subheader("📥 Экспорт отчета")
            
            html_report = "<html><head><meta charset='utf-8'><style>"
            html_report += "body { font-family: Arial, sans-serif; margin: 30px; color: #333; }"
            html_report += "h2 { color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; font-size: 18px; }"
            html_report += ".metric-box { background: #F3F4F6; padding: 15px; border-radius: 5px; margin-bottom: 20px; }"
            html_report += "table { width: 100%; border-collapse: collapse; margin-top: 15px; }"
            html_report += "th { background: #1E3A8A; color: white; padding: 10px; text-align: left; font-size: 13px; }"
            html_report += "td { padding: 10px; border-bottom: 1px solid #E5E7EB; font-size: 12px; }"
            html_report += "tr:nth-child(even) { background: #F9FAFB; }"
            html_report += ".right { text-align: right; }"
