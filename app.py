import streamlit as st
import pandas as pd

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Инструмент выявления ТОП-10 изменений в статьях расходов относительно прошлого месяца.")

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
    st.info("Файлы получены. Запускаю базовый расчёт таблиц...")
    
    try:
        pandas_header_index = int(header_row) - 1
        
        old_excel = pd.ExcelFile(old_file)
        new_excel = pd.ExcelFile(new_file)
        common_sheets = list(set(old_excel.sheet_names).intersection(set(new_excel.sheet_names)))
        
        all_expenses_changes = []
        total_old_dc = 0.0
        total_new_dc = 0.0
        total_row_found = False
        
        for sheet in common_sheets:
            df_old = pd.read_excel(old_file, sheet_name=sheet, header=pandas_header_index)
            df_new = pd.read_excel(new_file, sheet_name=sheet, header=pandas_header_index)
            
            df_old.columns = [str(c).strip() for c in df_old.columns]
            df_new.columns = [str(c).strip() for c in df_new.columns]
            
            if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                
                df_old_clean = df_old.dropna(subset=[target_column, value_column])
                df_new_clean = df_new.dropna(subset=[target_column, value_column])
                
                dict_old = pd.Series(df_old_clean[value_column].values, index=df_old_clean[target_column]).to_dict()
                dict_new = pd.Series(df_new_clean[value_column].values, index=df_new_clean[target_column]).to_dict()
                
                for k, v in dict_old.items():
                    if str(k).strip().lower() == total_row_name.lower().strip():
                        try: total_old_dc += float(v)
                        except: pass
                        total_row_found = True
                        
                for k, v in dict_new.items():
                    if str(k).strip().lower() == total_row_name.lower().strip():
                        try: total_new_dc += float(v)
                        except: pass
                
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
            top_10_changes = df_total_changes.sort_values(by="сортировка_влияния", ascending=False).head(10)
            
            if total_old_dc > 0:
                top_10_changes["Доля во влиянии на общую разницу"] = top_10_changes.apply(
                    lambda row: f"{row['Изменение (руб.)'] / total_old_dc * 100:+.2f}%", axis=1
                )
            else:
                top_10_changes["Доля во влиянии на общую разницу"] = "0.00%"
            
            top_10_display = top_10_changes.drop(columns=["сортировка_влияния"]).reset_index(drop=True)
            top_10_display.index = top_10_display.index + 1
            
            st.subheader("📋 Директорский отчет: ТОП-10 изменений")
            st.dataframe(top_10_display, use_container_width=True)
            
            # КНОПКА СКАЧИВАНИЯ ОТЧЕТА ДЛЯ ПЕЧАТИ В PDF
            st.write("---")
            st.subheader("📥 Экспорт отчета")
            
            html_report = convert_df_to_html_report(total_old_dc, total_new_dc, dc_delta, top_10_display)
            
            st.download_button(
                label="📄 Скачать отчет для сохранения в PDF",
                data=html_report,
                file_name="Director_Financial_Report.html",
                mime="text/html"
            )
            st.caption("💡 Как сделать PDF: Откройте скачанный файл и нажмите Ctrl+P (или Cmd+P на Mac) -> выберите 'Сохранить как PDF'.")
            
        else:
            st.info("📊 Изменений по статьям расходов между отчетами не обнаружено.")
            
    except Exception as e:
        st.error(f"⚠️ Произошла непредвиденная ошибка: {e}")
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для глубокого факторного анализа.")
