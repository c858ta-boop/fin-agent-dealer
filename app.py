import streamlit as st
import pandas as pd

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Инструмент автоматического выявления финансовых аномалий и критических отклонений (>10%).")

# Блок инструкций для настройки соответствия
with st.sidebar:
    st.header("⚙️ Настройки анализа")
    target_column = st.text_input("Название столбца со статьями:", value="Статья")
    value_column = st.text_input("Название столбца со значениями:", value="Сумма")
    st.caption("ℹ️ Убедитесь, что в ваших Excel файлах названия этих столбцов совпадают.")

# Блок загрузки файлов
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 Загрузите СТАРЫЙ отчет (прошлый месяц)", type=["xlsx"])
with col2:
    new_file = st.file_uploader("📂 Загрузите НОВЫЙ отчет (текущий месяц)", type=["xlsx"])

if old_file and new_file:
    st.success("Файлы успешно загружены! Запускаю глубокий финансовый аудит...")
    
    try:
        old_excel = pd.ExcelFile(old_file)
        new_excel = pd.ExcelFile(new_file)
        
        # Находим общие листы
        common_sheets = list(set(old_excel.sheet_names).intersection(set(new_excel.sheet_names)))
        
        if not common_sheets:
            st.error("❌ Ошибка: В файлах нет листов с одинаковыми названиями!")
        else:
            st.subheader("📋 Директорское резюме: Критические отклонения (>10%)")
            
            anomalies_found = False
            
            for sheet in common_sheets:
                df_old = pd.read_excel(old_file, sheet_name=sheet)
                df_new = pd.read_excel(new_file, sheet_name=sheet)
                
                # Приводим названия столбцов к нижнему регистру для надежности
                df_old.columns = [str(c).strip() for c in df_old.columns]
                df_new.columns = [str(c).strip() for c in df_new.columns]
                
                if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                    
                    # Переводим в словари для быстрого сопоставления {Статья: Сумма}
                    dict_old = pd.Series(df_old[value_column].values, index=df_old[target_column]).to_dict()
                    dict_new = pd.Series(df_new[value_column].values, index=df_new[target_column]).to_dict()
                    
                    # Считаем общий итог по листу для масштаба
                    total_old = sum(float(v) for v in dict_old.values() if pd.notna(v) and isinstance(v, (int, float)))
                    total_new = sum(float(v) for v in dict_new.values() if pd.notna(v) and isinstance(v, (int, float)))
                    sheet_delta = abs(total_new - total_old)
                    
                    if sheet_delta == 0:
                        continue
                        
                    # Проверяем каждую статью
                    all_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                    
                    sheet_anomalies = []
                    for article in all_articles:
                        if pd.isna(article) or str(article).strip() == "":
                            continue
                            
                        val_old = float(dict_old.get(article, 0)) if pd.notna(dict_old.get(article, 0)) else 0.0
                        val_new = float(dict_new.get(article, 0)) if pd.notna(dict_new.get(article, 0)) else 0.0
                        
                        item_delta = val_new - val_old
                        
                        # Влияние изменения этой строки на общее изменение листа в %
                        influence_percent = (abs(item_delta) / sheet_delta) * 100 if sheet_delta > 0 else 0
                        
                        if influence_percent >= 10.0 and abs(item_delta) > 0:
                            sheet_anomalies.append({
                                "Статья": article,
                                "Было": val_old,
                                "Стало": val_new,
                                "Изменение": item_delta,
                                "Влияние на итог листа": f"{influence_percent:.1f}%"
                            })
                    
                    if sheet_anomalies:
                        anomalies_found = True
                        with st.expander(f"⚠️ Лист '{sheet}': Общее изменение по листу: {item_delta:,.2f} руб.", expanded=True):
                            anom_df = pd.DataFrame(sheet_anomalies)
                            
                            # Подсвечиваем негативные изменения (рост расходов или падение доходов) красным
                            st.dataframe(anom_df.style.format({"Было": "{:,.2f}", "Стало": "{:,.2f}", "Изменение": "{:,.2f}"}))
                            
            if not anomalies_found:
                st.info("✅ Аномалий не обнаружено. Ни одна статья не изменилась более чем на 10% от общего финансового результата.")
                
    except Exception as e:
        st.error(f"⚠️ Произошла ошибка при анализе структуры Excel: {e}")
        st.info("💡 Подсказка: убедитесь, что в боковом меню правильно указаны названия столбцов со статьями и суммами.")
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для запуска автономного анализа.")
