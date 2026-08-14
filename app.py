import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Инструмент выявления ТОП-10 изменений в статьях расходов (исключая цветные суммирующие строки).")

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

def get_colored_rows(file_bytes, sheet_name, header_idx, target_col_name):
    """Быстро находит строки, в которых ячейка со статьей расходов имеет цветную заливку"""
    colored_rows = set()
    try:
        wb = openpyxl.load_workbook(file_bytes, data_only=True, read_only=True)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # Ищем индекс нужного столбца по строке заголовков
            target_col_idx = None
            header_row_cells = list(ws.iter_rows(min_row=header_idx, max_row=header_idx, values_only=False))[0]
            for idx, cell in enumerate(header_row_cells, start=1):
                if str(cell.value).strip() == target_col_name:
                    target_col_idx = idx
                    break
            
            # Если столбец найден, проверяем только его ячейки построчно
            if target_col_idx:
                for row_idx in range(header_idx + 1, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=target_col_idx)
                    if cell.fill and cell.fill.fill_type:
                        color = cell.fill.start_color.index
                        if color and str(color) not in ['00000000', '0', 'FFFFFFFF', 'System_Color_Window']:
                            # В pandas индексы строк будут сдвинуты, так как header_row стала заголовком.
                            # Строка row_idx в Excel соответствует индексу (row_idx - header_idx - 1) в df
                            colored_rows.add(row_idx - header_idx - 1)
    except Exception as e:
        pass
    return colored_rows

def convert_df_to_html_report(total_old, total_new, delta, df_top10):
    """Создает простой HTML-отчет для печати в PDF"""
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; margin: 30px; color: #333; }}
            h2 {{ color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; font-size: 18px; }}
            .metric-box {{ background: #F3F4F6; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th {{ background: #1E3A8A; color: white; padding: 10px; text-align: left; font-size: 13px; }}
            td {{ padding: 10px; border-bottom: 1px solid #E5E7EB; font-size: 12px; }}
            tr:nth-child(even) {{ background: #F9FAFB; }}
            .right {{ text-align: right; }}
        </style>
    </head>
    <body>
        <h2>Финансовый отчет Дилерского Центра</h2>
        <p><b>Факторный анализ изменений в статьях расходов</b></p>
        
        <div class="metric-box">
            <p>• Расходы за прошлый месяц: <b>{total_old:,.2f} руб.</b></p>
            <p>• Расходы за текущий месяц: <b>{total_new:,.2f} руб.</b></p>
            <p>• Изменение расходов всего ДЦ: <b>{delta:+,.2f} руб.</b></p>
        </div>
        
        <h2>ТОП-10 главных изменений в статьях расходов</h2>
        <p><i>В отчет включены только прямые статьи расходов (цветные промежуточные итоги исключены).</i></p>
        
        <table>
            <tr>
                <th>№</th>
                <th>Лист</th>
                <th>Статья расходов</th>
                <th class="right">Было (руб.)</th>
                <th class="right">Стало (руб.)</th>
                <th class="right">Изменение (руб.)</th>
                <th class="right">Доля во влиянии на общую разницу</th>
            </tr>
    """
    for idx, row in df_top10.iterrows():
        html += f"""
            <tr>
                <td>{idx}</td>
                <td>{row['Лист']}</td>
                <td>{row['Статья расходов']}</td>
                <td class="right">{row['Было (руб.)']:,.2f}</td>
                <td class="right">{row['Стало (руб.)']:,.2f}</td>
                <td class="right">{row['Изменение (руб.)']:+,.2f}</td>
                <td class="right">{row['Доля во влиянии на общую разницу']}</td>
            </tr>
        """
    html += """
        </table>
    </body>
    </html>
    """
    return html

if old_file and new_file:
    st.info("Файлы получены. Запускаю глубокий анализ с фильтрацией по цвету...")
    
    try:
        pandas_header_index = int(header_row) - 1
        
        # Фиксируем данные файлов в памяти, чтобы прочитать дважды без сбоев
        old_bytes_1 = old_file.read()
        new_bytes_1 = new_file.read()
        
        old_excel = pd.ExcelFile(BytesIO(old_bytes_1))
        new_excel = pd.ExcelFile(BytesIO(new_bytes_1))
        common_sheets = list(set(old_excel.sheet_names).intersection(set(new_excel.sheet_names)))
        
        all_expenses_changes = []
        total_old_dc = 0.0
        total_new_dc = 0.0
        total_row_found = False
        
        for sheet in common_sheets:
            df_old = pd.read_excel(BytesIO(old_bytes_1), sheet_name=sheet, header=pandas_header_index)
            df_new = pd.read_excel(BytesIO(new_bytes_1), sheet_name=sheet, header=pandas_header_index)
            
            df_old.columns = [str(c).strip() for c in df_old.columns]
            df_new.columns = [str(c).strip() for c in df_new.columns]
            
            if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                
                # Быстро сканируем openpyxl только этот лист на предмет цветных строк
                colored_old_rows = get_colored_rows(BytesIO(old_bytes_1), sheet, int(header_row), target_column)
                colored_new_rows = get_colored_rows(BytesIO(new_bytes_1), sheet, int(header_row), target_column)
                
                # Вытаскиваем "Всего по ДЦ" до очистки (даже если финотдел покрасил этот тотал)
                for i, row in df_old.iterrows():
                    if str(row[target_column]).strip().lower() == total_row_name.lower().strip():
                        try: total_old_dc += float(row[value_column])
                        except: pass
                        total_row_found = True
                for i, row in df_new.iterrows():
                    if str(row[target_column]).strip().lower() == total_row_name.lower().strip():
                        try: total_new_dc += float(row[value_column])
                        except: pass
                
                # Фильтруем датафреймы — выкидываем цветные строки
                df_old_clean = df_old.drop(index=list(colored_old_rows), errors='ignore').dropna(subset=[target_column, value_column])
                df_new_clean = df_new.drop(index=list(colored_new_rows), errors='ignore').dropna(subset=[target_column, value_column])
                
                dict_old = pd.Series(df_old_clean[value_column].values, index=df_old_clean[target_column]).to_dict()
                dict_new = pd.Series(df_new_clean[value_column].values, index=df_new_clean[target_column]).to_dict()
                
                sheet_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                for article in sheet_articles:
                    article_str = str(article).strip()
                    
                    # Стандартный текстовый фильтр итогов
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
                            "Лист": sheet,
                            "Статья расходов": article_str,
                            "Было (руб.)": val_old,
                            "Стало (руб.)": val_new,
                            "Изменение (руб.)": item_delta,
                            "сортировка_влияния": abs_delta
                        })
        
        dc_delta = total_new_dc - total_old_dc
        
        st.subheader("📊 Общий финансовый результат по ДЦ")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Расходы за прошлый месяц", f"{total_old_dc:,.2f} руб.")
        with c2:
            st.metric("Расходы за текущий месяц", f"{total_new_dc:,.2f} руб.")
        with c3:
            st.metric("Общее изменение расходов ДЦ", f"{dc_delta:+,.2f} руб.", delta_color="inverse")
            
        if all_expenses_changes:
            df_total_changes = pd.DataFrame(all_expenses_changes)
