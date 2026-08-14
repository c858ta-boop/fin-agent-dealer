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
            ws_old = wb_old[sheet_name]
            ws_new = wb_new[sheet_name]
            
            # Ищем индексы нужных столбцов по строке заголовков
            target_col_idx_old, value_col_idx_old = None, None
            target_col_idx_new, value_col_idx_new = None, None
            
            header_row_cells_old = list(ws_old.iter_rows(min_row=header_idx, max_row=header_idx, values_only=False))
            for idx, cell in enumerate(header_row_cells_old, start=1):
                if str(cell.value).strip() == target_column:
                    target_col_idx_old = idx
                    break
            
            header_row_cells_new = list(ws_new.iter_rows(min_row=header_idx, max_row=header_idx, values_only=False))
            for idx, cell in enumerate(header_row_cells_new, start=1):
                if str(cell.value).strip() == target_column:
                    target_col_idx_new = idx
                    break
            
            # Если структура совпала, быстро находим цветные строки на этом листе
            colored_old_rows = set()
            if target_col_idx_old:
                for r_idx in range(header_idx + 1, ws_old.max_row + 1):
                    cell = ws_old.cell(row=r_idx, column=target_col_idx_old)
                    if cell.fill and cell.fill.fill_type:
                        col_idx = cell.fill.start_color.index
                        if col_idx and str(col_idx) not in ['00000000', '0', 'FFFFFFFF', 'System_Color_Window']:
                            colored_old_rows.add(r_idx - header_idx - 1)
                            
            colored_new_rows = set()
            if target_col_idx_new:
                for r_idx in range(header_idx + 1, ws_new.max_row + 1):
                    cell = ws_new.cell(row=r_idx, column=target_col_idx_new)
                    if cell.fill and cell.fill.fill_type:
                        col_idx = cell.fill.start_color.index
                        if col_idx and str(col_idx) not in ['00000000', '0', 'FFFFFFFF', 'System_Color_Window']:
                            colored_new_rows.add(r_idx - header_idx - 1)
            
            # Читаем данные через pandas для математических расчетов
            df_old = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet_name, header=pandas_header_index)
            df_new = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_name, header=pandas_header_index)
            
            df_old.columns = [str(c).strip() for c in df_old.columns]
            df_new.columns = [str(c).strip() for c in df_new.columns]
            
            # Проверяем наличие целевых столбцов на листах
            if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                
                # Фиксируем общий итог "Всего по ДЦ" до очистки цвета
                for i, row in df_old.iterrows():
                    if str(row[target_column]).strip().lower() == total_row_name.lower().strip():
                        try: total_old_dc += float(row[value_column])
                        except: pass
                        total_row_found = True
                for i, row in df_new.iterrows():
                    if str(row[target_column]).strip().lower() == total_row_name.lower().strip():
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
            
            # Рассчитываем процент строго в одну строку без if-else
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
            
            # Создаем HTML без двойных фигурных скобок в f-строке, чтобы избежать синтаксического сбоя
            html_report = "<html><head><meta charset='utf-8'><style>"
            html_report += "body { font-family: Arial, sans-serif; margin: 30px; color: #333; }"
            html_report += "h2 { color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; font-size: 18px; }"
            html_report += ".metric-box { background: #F3F4F6; padding: 15px; border-radius: 5px; margin-bottom: 20px; }"
            html_report += "table { width: 100%; border-collapse: collapse; margin-top: 15px; }"
            html_report += "th { background: #1E3A8A; color: white; padding: 10px; text-align: left; font-size: 13px; }"
            html_report += "td { padding: 10px; border-bottom: 1px solid #E5E7EB; font-size: 12px; }"
