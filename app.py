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
    if cell.fill and cell.fill.fill_type:
        color = cell.fill.start_color.index
        if color and str(color) not in ['00000000', '0', 'FFFFFFFF', 'System_Color_Window']:
            return True
    return False

def convert_df_to_html_report(total_old, total_new, delta, df_top10):
    """Создает простой HTML-отчет для печати в PDF"""
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; margin: 30px; color: #333; }}
            h2 {{ color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; }}
            .metric-box {{ background: #F3F4F6; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th {{ background: #1E3A8A; color: white; padding: 10px; text-align: left; font-size: 14px; }}
            td {{ padding: 10px; border-bottom: 1px solid #E5E7EB; font-size: 13px; }}
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
                <th class="right">Доля в ДЦ</th>
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
                <td class="right">{row['Доля в ДЦ']}</td>
            </tr>
        """
    html += """
        </table>
    </body>
    </html>
    """
    return html

# Основная логика приложения
if old_file and new_file:
    st.success("Файлы успешно загружены! Начинаю факторный анализ...")
    
    # Клонируем файлы в независимые буферы памяти для стабильного чтения
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
        
        for sheet_name in common_sheets:
            ws_old = wb_old[sheet_name]
            ws_new = wb_new[sheet_name]
            
            target_col_idx_old, value_col_idx_old = None, None
            target_col_idx_new, value_col_idx_new = None, None
            
            for col in range(1, ws_old.max_column + 1):
                val = str(ws_old.cell(row=header_idx, column=col).value).strip()
                if val == target_column: target_col_idx_old = col
                if val == value_column: value_col_idx_old = col
                    
            for col in range(1, ws_new.max_column + 1):
                val = str(ws_new.cell(row=header_idx, column=col).value).strip()
                if val == target_column: target_col_idx_new = col
                if val == value_column: value_col_idx_new = col
            
            if target_col_idx_old and value_col_idx_old and target_col_idx_new and value_col_idx_new:
                dict_old = {}
                for r in range(header_idx + 1, ws_old.max_row + 1):
                    cell_art = ws_old.cell(row=r, column=target_col_idx_old)
                    cell_val = ws_old.cell(row=r, column=value_col_idx_old)
                    if cell_art.value is not None:
                        art_str = str(cell_art.value).strip()
                        if art_str.lower() == total_row_name.lower().strip():
                            try: total_old_dc += float(cell_val.value or 0)
                            except: pass
                            total_row_found = True
                            continue
                        if is_colored(cell_art):
                            continue
                        dict_old[art_str] = cell_val.value

                dict_new = {}
                for r in range(header_idx + 1, ws_new.max_row + 1):
                    cell_art = ws_new.cell(row=r, column=target_col_idx_new)
                    cell_val = ws_new.cell(row=r, column=value_col_idx_new)
                    if cell_art.value is not None:
                        art_str = str(cell_art.value).strip()
                        if art_str.lower() == total_row_name.lower().strip():
                            try: total_new_dc += float(cell_val.value or 0)
                            except: pass
                            continue
                        if is_colored(cell_art):
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
                            "abs_vliyanie": abs_delta  # Заменяем имя на строго английское во избежание багов
                        })
        
        dc_delta = total_new_dc - total_old_dc
        
        st.subheader("📊 Общий финансовый результат по ДЦ")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Расходы за прошлый ...", f"{total_old_dc:,.2f} руб.")
        with c2:
            st.metric("Расходы за текущий ...", f"{total_new_dc:,.2f} руб.")
        with c3:
            st.metric("Общее изменение расходов ДЦ", f"{dc_delta:+,.2f} руб.", delta_color="inverse")
            
        if not total_row_found:
            st.warning(f"⚠️ Строка '{total_row_name}' не найдена в файлах.")
        
        if all_expenses_changes:
            df_total_changes = pd.DataFrame(all_expenses_changes)
            
            # Сортируем по английскому ключу без риска опечаток
            top_10_changes = df_total_changes.sort_values(by="abs_vliyanie", ascending=False).head(10)
            
            # Рассчитываем процент текстом со знаком %
            if total_old_dc > 0:
                top_10_changes["Доля в ДЦ"] = top_10_changes.apply(
                    lambda row: f"{row['Изменение (руб.)'] / total_old_dc * 100:+.2f}%", axis=1
                )
            else:
                top_10_changes["Доля в ДЦ"] = "0.00%"
            
            top_10_display = top_10_changes.drop(columns=["abs_vliyanie"], errors='ignore').reset_index(drop=True)
            top_10_display.index = top_10_display.index + 1
            
            st.subheader("📋 Директорский отчет: ТОП-10 чистых статей расходов")
