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

# Основная логика по условию загрузки файлов
if old_file and new_file:
    st.success("Файлы успешно загружены! Начинаю факторный анализ...")
    
    try:
        # ЗАЩИТА СТРИМА: Клонируем файлы в независимые буферы памяти BytesIO
        old_data = BytesIO(old_file.read())
        new_data = BytesIO(new_file.read())
        
        # Считываем книги
        wb_old = openpyxl.load_workbook(old_data, data_only=True)
        wb_new = openpyxl.load_workbook(new_data, data_only=True)
        
        common_sheets = list(set(wb_old.sheetnames).intersection(set(wb_new.sheetnames)))
        
        if not common_sheets:
            st.error("❌ Ошибка: В файлах нет листов с одинаковыми названиями!")
        else:
            all_expenses_changes = []
            total_old_dc = 0.0
            total_new_dc = 0.0
            total_row_found = False
            header_idx = int(header_row)
            
            # 1. СНАЧАЛА ИЩЕМ ОБЩИЙ ИТОГ ПО ВСЕМ СТОЛБЦАМ И СТРОКАМ БЕЗ ОГРАНИЧЕНИЙ НА ЦВЕТ
            # Сканируем старую книгу
            for sheet_name in wb_old.sheetnames:
                ws = wb_old[sheet_name]
                # Определяем индекс столбца значений
                val_col_idx = None
                for col in range(1, ws.max_column + 1):
                    if str(ws.cell(row=header_idx, column=col).value).strip() == value_column:
                        val_col_idx = col
                        break
                if val_col_idx:
                    for r in range(header_idx + 1, ws.max_row + 1):
                        for c in range(1, ws.max_column + 1):
                            if str(ws.cell(row=r, column=c).value).strip().lower() == total_row_name.lower().strip():
                                try:
                                    total_old_dc += float(ws.cell(row=r, column=val_col_idx).value or 0)
                                    total_row_found = True
                                except:
                                    pass
                                break

            # Сканируем новую книгу
            for sheet_name in wb_new.sheetnames:
                ws = wb_new[sheet_name]
                val_col_idx = None
                for col in range(1, ws.max_column + 1):
                    if str(ws.cell(row=header_idx, column=col).value).strip() == value_column:
                        val_col_idx = col
                        break
                if val_col_idx:
                    for r in range(header_idx + 1, ws.max_row + 1):
                        for c in range(1, ws.max_column + 1):
                            if str(ws.cell(row=r, column=c).value).strip().lower() == total_row_name.lower().strip():
                                try:
                                    total_new_dc += float(ws.cell(row=r, column=val_col_idx).value or 0)
                                except:
                                    pass
                                break

            # 2. ТЕПЕРЬ СОБИРАЕМ ОБЫЧНЫЕ СТРОКИ С ФИЛЬТРАЦИЕЙ ЦВЕТА КАК РАНЬШЕ
            for sheet_name in common_sheets:
                ws_old_sheet = wb_old[sheet_name]
                ws_new_sheet = wb_new[sheet_name]
                
                target_col_idx_old, value_col_idx_old = None, None
                target_col_idx_new, value_col_idx_new = None, None
                
                for col in range(1, ws_old_sheet.max_column + 1):
                    val = str(ws_old_sheet.cell(row=header_idx, column=col).value).strip()
                    if val == target_column: target_col_idx_old = col
                    if val == value_column: value_col_idx_old = col
                        
                for col in range(1, ws_new_sheet.max_column + 1):
                    val = str(ws_new_sheet.cell(row=header_idx, column=col).value).strip()
                    if val == target_column: target_col_idx_new = col
                    if val == value_column: value_col_idx_new = col
                
                if target_col_idx_old and value_col_idx_old and target_col_idx_new and value_col_idx_new:
                    dict_old = {}
                    for r in range(header_idx + 1, ws_old_sheet.max_row + 1):
                        cell_art = ws_old_sheet.cell(row=r, column=target_col_idx_old)
                        cell_val = ws_old_sheet.cell(row=r, column=value_col_idx_old)
                        if cell_art.value is not None:
                            art_str = cell_art.value.strip() if hasattr(cell_art.value, 'strip') else str(cell_art.value).strip()
                            if art_str.lower() == total_row_name.lower().strip() or is_colored(cell_art):
                                continue
                            dict_old[art_str] = cell_val.value

                    dict_new = {}
                    for r in range(header_idx + 1, ws_new_sheet.max_row + 1):
                        cell_art = ws_new_sheet.cell(row=r, column=target_col_idx_new)
                        cell_val = ws_new_sheet.cell(row=r, column=value_col_idx_new)
                        if cell_art.value is not None:
                            art_str = cell_art.value.strip() if hasattr(cell_art.value, 'strip') else str(cell_art.value).strip()
                            if art_str.lower() == total_row_name.lower().strip() or is_colored(cell_art):
                                continue
                            dict_new[art_str] = cell_val.value
                    
                    sheet_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                    for article in sheet_articles:
                        if article == "" or any(word in article.lower() for word in ["итого", "всего", "баланс", "результат", "свод"]):
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
                                "Статья расходов": article,
                                "Было (руб.)": val_old,
                                "Стало (руб.)": val_new,
                                "Изменение (руб.)": item_delta,
                                "sort_key": abs_delta
                            })
            
            # Вывод результатов
            dc_delta = total_new_dc - total_old_dc
            
            st.subheader("📊 Общий финансовый результат по ДЦ")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Расходы за прошлый месяц", f"{total_old_dc:,.2f} руб.")
            with c2:
                st.metric("Расходы за текущий месяц", f"{total_new_dc:,.2f} руб.")
            with c3:
                st.metric("Общее изменение расходов ДЦ", f"{dc_delta:+,.2f} руб.", delta_color="inverse")
                
            if not total_row_found:
                st.warning(f"⚠️ Строка '{total_row_name}' не найдена в файлах.")
            
            if all_expenses_changes:
                df_total_changes = pd.DataFrame(all_expenses_changes)
                top_10_changes = df_total_changes.sort_values(by="sort_key", ascending=False).head(10)
                
                base_denom = total_old_dc if total_old_dc > 0 else 1.0
                top_10_changes["Доля во влиянии на общую разницу"] = top_10_changes.apply(
                    lambda r: f"{r['Изменение (руб.)'] / base_denom * 100:+.2f}%", axis=1
                )
                
                top_10_display = top_10_changes.drop(columns=["sort_key"]).reset_index(drop=True)
                top_10_display.index = top_10_display.index + 1
                
                st.subheader("📋 Директорский отчет: ТОП-10 чистых статей расходов")
                st.write("Суммирующие строки отделов отфильтрованы по цвету заливки. Показываются только прямые статьи расходов:")
                st.dataframe(top_10_display, use_container_width=True)
                
                st.write("---")
                st.subheader("🖨️ Печать и экспорт в PDF")
