import streamlit as st
import pandas as pd

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Анализ вклада отдельных статей расходов в изменение общего результата по ДЦ.")

# Панель настроек в боковой панели
with st.sidebar:
    st.header("⚙️ Настройки анализа")
    target_column = st.text_input("Название столбца со статьями расходов:", value="Статья расходов")
    value_column = st.text_input("Название столбца со значениями (суммами):", value="Всего расходы")
    header_row = st.number_input("Строка с заголовками (в Excel нумерация с 1):", min_value=1, value=2)
    total_row_name = st.text_input("Название строки общего итога:", value="Всего по ДЦ")
    st.caption("ℹ️ Настройки оптимизированы под структуру ваших ежемесячных отчетов.")

# Блок загрузки файлов
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 Загрузите СТАРЫЙ отчет (прошлый месяц)", type=["xlsx"])
with col2:
    new_file = st.file_uploader("📂 Загрузите НОВЫЙ отчет (текущий месяц)", type=["xlsx"])

if old_file and new_file:
    st.success("Файлы успешно загружены! Начинаю факторный анализ расходов...")
    
    try:
        old_excel = pd.ExcelFile(old_file)
        new_excel = pd.ExcelFile(new_file)
        
        common_sheets = list(set(old_excel.sheet_names).intersection(set(new_excel.sheet_names)))
        
        if not common_sheets:
            st.error("❌ Ошибка: В файлах нет листов с одинаковыми названиями!")
        else:
            all_expenses_changes = []
            total_old_dc = 0.0
            total_new_dc = 0.0
            total_row_found = False
            
            pandas_header_index = int(header_row) - 1
            
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
                    
                    # 1. Ищем и фиксируем строку "Всего по ДЦ" для определения общего изменения
                    for k, v in dict_old.items():
                        if str(k).strip().lower() == total_row_name.lower().strip():
                            try: total_old_dc += float(v)
                            except: pass
                            total_row_found = True
                            
                    for k, v in dict_new.items():
                        if str(k).strip().lower() == total_row_name.lower().strip():
                            try: total_new_dc += float(v)
                            except: pass
                    
                    # 2. Собираем обычные статьи расходов
                    sheet_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                    
                    for article in sheet_articles:
                        article_str = str(article).strip()
                        
                        # Пропускаем строку тотала и технический мусор, чтобы они не двоились в анализе
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
                                "Абсолютное влияние (руб.)": abs_delta
                            })
            
            # Выводим сводные результаты по ДЦ
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
                st.warning(f"⚠️ Строка '{total_row_name}' не найдена в файлах. Общий итог рассчитан как сумма листов.")
            
            # Формируем финальный ТОП-10 статей по их вкладу
            if all_expenses_changes:
                df_total_changes = pd.DataFrame(all_expenses_changes)
                
                # Сортируем по силе абсолютного численного влияния
                top_10_changes = df_total_changes.sort_values(by="Абсолютное влияние (руб.)", ascending=False).head(10)
                
                # Считаем процент вклада каждой статьи в общее изменение ДЦ
                if abs(dc_delta) > 0:
                    top_10_changes["Доля во влиянии на общую разницу"] = top_10_changes.apply(
                        lambda row: f"{(row['Абсолютное влияние (руб.)'] / abs(dc_delta)) * 100:.1f}%", axis=1
                    )
                else:
                    top_10_changes["Доля во влиянии на общую разницу"] = "0.0%"
                
                top_10_display = top_10_changes.drop(columns=["Абсолютное влияние (руб.)"], errors='ignore').reset_index(drop=True)
                top_10_display.index = top_10_display.index + 1
                
                st.subheader(f"📋 Директорский отчет: ТОП-10 виновников изменения на {dc_delta:,.2f} руб.")
                st.write("Эти 10 статей оказали самое мощное численное воздействие на финальный результат расходов компании:")
                
                st.dataframe(
                    top_10_display.style.format({
                        "Было (руб.)": "{:,.2f}", 
                        "Стало (руб.)": "{:,.2f}", 
                        "Изменение (руб.)": "{:+,.2f}"
                    }),
                    use_container_width=True
                )
            else:
                st.info("📊 Изменений по расходам между отчетами не найдено.")
                
    except Exception as e:
        st.error(f"⚠️ Произошла ошибка при анализе структуры Excel: {e}")
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для глубокого факторного анализа.")
