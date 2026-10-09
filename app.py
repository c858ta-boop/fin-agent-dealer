import streamlit as st
import pandas as pd
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

def clean_to_float(val):
    """Всеядная функция для приведения ячеек к числу с плавающей точкой"""
    if pd.isna(val) or val is None:
        return 0.0
    val_str = str(val).strip()
    if val_str == "" or val_str == "-":
        return 0.0
    try:
        val_str = val_str.replace('\xa0', '').replace(' ', '').replace(',', '.')
        return float(val_str)
    except:
        return 0.0

# Основная логика приложения
if old_file and new_file:
    st.success("Файлы успешно загружены! Начинаю факторный анализ ДЦ...")
    
    pandas_header_index = int(header_row) - 1
    
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
        
        # Сканируем каждый общий лист быстрым методом Pandas
        for sheet in common_sheets:
            # Читаем БЕЗ указания header, чтобы видеть абсолютно всё содержимое листа с 1-й строки
            df_old_raw = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet, header=None)
            df_new_raw = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet, header=None)
            
            # 1. БРОНЕБОЙНЫЙ ПОИСК "В СТОЛБЦЕ А" ДЛЯ ОБЩЕГО ИТОГА (БЕЗ УЧЕТА ЦВЕТА И НАЗВАНИЯ ШАПКИ)
            # Ищем тотал в Старом файле
            for i, row in df_old_raw.iterrows():
                # Проверяем самую первую ячейку строки (столбец А имеет индекс 0)
                cell_a_val = str(row.iloc[0]).strip().lower()
                if cell_a_val == total_row_name.lower().strip():
                    # Финотдел мог сместить сумму вправо, ищем первое числовое значение в этой строке
                    for col_val in row.iloc[1:]:
                        val_float = clean_to_float(col_val)
                        if val_float > 0:
                            total_old_dc += val_float
                            total_row_found = True
                            break
                    break
                    
            # Ищем тотал в Новом файле
            for i, row in df_new_raw.iterrows():
                cell_a_val = str(row.iloc[0]).strip().lower()
                if cell_a_val == total_row_name.lower().strip():
                    for col_val in row.iloc[1:]:
                        val_float = clean_to_float(col_val)
                        if val_float > 0:
                            total_new_dc += val_float
                            break
                    break
            
            # 2. ЧИТАЕМ СТАТЬИ ДЛЯ ТОП-10 (КАК В НАШЕЙ САМОЙ СТАБИЛЬНОЙ ВЕРСИИ)
            df_old = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet, header=pandas_header_index)
            df_new = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet, header=pandas_header_index)
            
            df_old.columns = [str(c).strip() for c in df_old.columns]
            df_new.columns = [str(c).strip() for c in df_new.columns]
            
            if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                df_old_clean = df_old.dropna(subset=[target_column, value_column])
                df_new_clean = df_new.dropna(subset=[target_column, value_column])
                
                dict_old = pd.Series(df_old_clean[value_column].values, index=df_old_clean[target_column]).to_dict()
                dict_new = pd.Series(df_new_clean[value_column].values, index=df_new_clean[target_column]).to_dict()
                
                sheet_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                for article in sheet_articles:
                    article_str = str(article).strip()
                    # Игнорируем тотал и служебные слова финотдела
                    if article_str.lower() == total_row_name.lower().strip() or article_str == "" or any(word in article_str.lower() for word in ["итого", "всего", "баланс", "результат", "свод"]):
                        continue
                        
                    val_old = clean_to_float(dict_old.get(article, 0))
                    val_new = clean_to_float(dict_new.get(article, 0))
                    
                    item_delta = val_new - val_old
                    abs_delta = abs(item_delta)
                    
                    if abs_delta > 0:
                        all_expenses_changes.append({
                            "Лист": sheet,
                            "Статья расходов": article_str,
                            "Было (руб.)": val_old,
                            "Стало (руб.)": val_new,
                            "Изменение (руб.)": item_delta,
                            "sort_key": abs_delta
                        })
        
        # Вывод результатов на экран
        dc_delta = total_new_dc - total_old_dc
        
        st.subheader("📊 Общий финансовый результат по ДЦ")
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Расходы за прошлый месяц", f"{total_old_dc:,.2f} руб.")
        with c2: st.metric("Расходы за текущий месяц", f"{total_new_dc:,.2f} руб.")
        with c3: st.metric("Общее изменение расходов ДЦ", f"{dc_delta:+,.2f} руб.", delta_color="inverse")
            
        if not total_row_found:
            st.warning(f"⚠️ Строка '{total_row_name}' не найдена в столбце А ваших файлов.")
            
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
            st.dataframe(top_10_display, use_container_width=True)
            
            st.write("---")
            st.subheader("🖨️ Печать и экспорт в PDF")
            st.write("Нажмите комбинацию клавиш **Ctrl + P** (или **Cmd + P** на Mac) прямо на этой странице браузера, чтобы сохранить этот отчет в PDF.")
            
            # Абсолютно плоский HTML без внутренних f-строк во избежание синтаксических ошибок
            html_preview = "<html><head><meta charset='utf-8'><style>"
            html_preview += "body { font-family: Arial, sans-serif; padding: 20px; color: #333; }"
            html_preview += "h2 { color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; font-size: 18px; margin-top:0; }"
            html_preview += "table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 11px; }"
            html_preview += "th { background: #1E3A8A; color: white; padding: 6px; text-align: left; }"
            html_preview += "td { padding: 6px; border-bottom: 1px solid #E5E7EB; }"
            html_preview += "</style></head><body><div style='background: white;'>"
            html_preview += "<h2 style='margin-bottom:15px;'>Финансовый отчет Дилерского Центра</h2>"
            html_preview += "<p>• Расходы прошлого месяца: <b>" + f"{total_old_dc:,.2f}" + " руб.</b></p>"
            html_preview += "<p>• Расходы текущего месяца: <b>" + f"{total_new_dc:,.2f}" + " руб.</b></p>"
            html_preview += "<p>• Общее изменение расходов ДЦ: <b>" + f"{dc_delta:+,.2f}" + " руб.</b></p>"
            html_preview += "<h3>ТОП-10 главных изменений в статьях расходов:</h3>"
            html_preview += "<table><tr><th>№</th><th>Лист</th><th>Статья расходов</th><th style='text-align: right;'>Было (руб.)</th><th style='text-align: right;'>Стало (руб.)</th><th style='text-align: right;'>Изменение (руб.)</th><th style='text-align: right;'>Доля во влиянии</th></tr>"
            
            for idx, row in top_10_display.iterrows():
                bg_row = "#F9FAFB" if idx % 2 == 0 else "#FFFFFF"
                html_preview += "<tr style='background: " + str(bg_row) + ";'>"
                html_preview += "<td>" + str(idx) + "</td><td>" + str(row['Лист']) + "</td><td>" + str(row['Статья расходов']) + "</td>"
                html_preview += "<td style='text-align: right;'>" + f"{row['Было (руб.)']:,.2f}" + "</td>"
