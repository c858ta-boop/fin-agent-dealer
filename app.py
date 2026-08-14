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
        wb = openpyxl.load_workbook(file_bytes, data_only=True)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            target_col_idx = None
            
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=header_idx, column=col).value
                if cell_val is not None and str(cell_val).strip() == target_col_name:
                    target_col_idx = col
                    break
                    
            if target_col_idx:
                for row_idx in range(header_idx + 1, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=target_col_idx)
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
        
        for sheet_name in common_sheets:
            colored_old_rows = get_colored_rows(BytesIO(old_bytes), sheet_name, header_idx, target_column)
            colored_new_rows = get_colored_rows(BytesIO(new_bytes), sheet_name, header_idx, target_column)
            
            df_old = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet_name, header=pandas_header_index)
            df_new = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_name, header=pandas_header_index)
            
            df_old.columns = [str(c).strip() for c in df_old.columns]
            df_new.columns = [str(c).strip() for c in df_new.columns]
            
            if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                
                for i, row in df_old.iterrows():
                    if row[target_column] is not None and str(row[target_column]).strip().lower() == total_row_name.lower().strip():
                        try: total_old_dc += float(row[value_column])
                        except: pass
                        total_row_found = True
                for i, row in df_new.iterrows():
                    if row[target_column] is not None and str(row[target_column]).strip().lower() == total_row_name.lower().strip():
                        try: total_new_dc += float(row[value_column])
                        except: pass
                
                df_old_clean = df_old.drop(index=list(colored_old_rows), errors='ignore').dropna(subset=[target_column, value_column])
                df_new_clean = df_new.drop(index=list(colored_new_rows), errors='ignore').dropna(subset=[target_column, value_column])
                
                dict_old = pd.Series(df_old_clean[value_column].values, index=df_old_clean[target_column]).to_dict()
                dict_new = pd.Series(df_new_clean[value_column].values, index=df_new_clean[target_column]).to_dict()
                
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
            st.write("Промежуточные итоги отделов успешно отфильтрованы по цвету заливки ячеек.")
            st.dataframe(top_10_display, use_container_width=True)
            
            # 📄 НОВЫЙ БЛОК: Генерация печатной формы прямо на экран
            st.write("---")
            st.subheader("🖨️ Печать и экспорт в PDF")
            st.write("Нажмите комбинацию клавиш **Ctrl + P** (или **Cmd + P** на Mac) прямо на этой странице браузера, чтобы мгновенно сохранить этот отчет в PDF.")
            
            # Собираем красивый HTML-блок для вывода на экран
            html_preview = "<div style='font-family: Arial, sans-serif; padding: 20px; border: 1px solid #E5E7EB; border-radius: 5px; background: white;'>"
            html_preview += "<h2 style='color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; font-size: 18px; margin-top:0;'>Финансовый отчет Дилерского Центра</h2>"
            html_preview += f"<p>• Расходы прошлого месяца: <b>{total_old_dc:,.2f} руб.</b></p>"
            html_preview += f"<p>• Расходы текущего месяца: <b>{total_new_dc:,.2f} руб.</b></p>"
            html_preview += f"<p>• Общее изменение расходов ДЦ: <b style='color: #1E3A8A;'>{dc_delta:+,.2f} руб.</b></p>"
            html_preview += "<h3 style='color: #1E3A8A; font-size: 14px;'>ТОП-10 главных изменений в статьях расходов:</h3>"
            html_preview += "<table style='width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px;'>"
            html_preview += "<tr style='background: #1E3A8A; color: white;'>"
            html_preview += "<th style='padding: 8px; text-align: left;'>№</th><th style='padding: 8px; text-align: left;'>Лист</th><th style='padding: 8px; text-align: left;'>Статья расходов</th><th style='padding: 8px; text-align: right;'>Было (руб.)</th><th style='padding: 8px; text-align: right;'>Стало (руб.)</th><th style='padding: 8px; text-align: right;'>Изменение (руб.)</th><th style='padding: 8px; text-align: right;'>Доля во влиянии</th></tr>"
            
            for idx, row in top_10_display.iterrows():
                bg_color = "#F9FAFB" if idx % 2 == 0 else "#FFFFFF"
                html_preview += f"<tr style='background: {bg_color}; border-bottom: 1px solid #E5E7EB;'>"
