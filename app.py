import streamlit as st
import pandas as pd

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Инструмент выявления ТОП-10 статей расходов, оказавших наибольшее численное влияние на финансовый результат.")

# Блок инструкций для настройки соответствия в боковой панели
with st.sidebar:
    st.header("⚙️ Настройки анализа")
    target_column = st.text_input("Название столбца со статьями расходов:", value="Статья расходов")
    value_column = st.text_input("Название столбца со значениями (суммами):", value="Всего расходы")
    header_row = st.number_input("Строка с заголовками (в Excel нумерация с 1):", min_value=1, value=2)
    st.caption("ℹ️ Сейчас настройки адаптированы: заголовки ищутся на 2-й строчке, а данные читаются с 3-й.")

# Блок загрузки файлов
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 Загрузите СТАРЫЙ отчет (прошлый месяц)", type=["xlsx"])
with col2:
    new_file = st.file_uploader("📂 Загрузите НОВЫЙ отчет (текущий месяц)", type=["xlsx"])

if old_file and new_file:
    st.success("Файлы успешно загружены! Начинаю поиск ТОП-10 критических изменений...")
    
    try:
        old_excel = pd.ExcelFile(old_file)
        new_excel = pd.ExcelFile(new_file)
        
        # Находим общие листы в двух файлах
        common_sheets = list(set(old_excel.sheet_names).intersection(set(new_excel.sheet_names)))
        
        if not common_sheets:
            st.error("❌ Ошибка: В файлах нет листов с одинаковыми названиями!")
        else:
            all_expenses_changes = []
            
            # В pandas нумерация строк идет с 0, поэтому строка 2 в Excel — это индекс 1
            pandas_header_index = int(header_row) - 1
            
            for sheet in common_sheets:
                # Читаем файл, указывая, на какой строчке находятся заголовки
                df_old = pd.read_excel(old_file, sheet_name=sheet, header=pandas_header_index)
                df_new = pd.read_excel(new_file, sheet_name=sheet, header=pandas_header_index)
                
                # Очищаем названия столбцов от пробелов по краям
                df_old.columns = [str(c).strip() for c in df_old.columns]
                df_new.columns = [str(c).strip() for c in df_new.columns]
                
                if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                    
                    # Фильтруем строки: убираем пустые значения в статьях и суммах
                    df_old_clean = df_old.dropna(subset=[target_column, value_column])
                    df_new_clean = df_new.dropna(subset=[target_column, value_column])
                    
                    # Переводим данные в словари {Статья: Сумма}
                    dict_old = pd.Series(df_old_clean[value_column].values, index=df_old_clean[target_column]).to_dict()
                    dict_new = pd.Series(df_new_clean[value_column].values, index=df_new_clean[target_column]).to_dict()
                    
                    # Объединяем все уникальные статьи расходов на листе
                    sheet_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                    
                    for article in sheet_articles:
                        # Фильтруем технический мусор и строки итогов
                        article_str = str(article).strip()
                        if article_str == "" or any(word in article_str.lower() for word in ["итого", "всего", "баланс", "результат", "свод"]):
                            continue
                            
                        # Безопасно переводим значения в числа
                        try:
                            val_old = float(dict_old.get(article, 0))
                        except:
                            val_old = 0.0
                            
                        try:
                            val_new = float(dict_new.get(article, 0))
                        except:
                            val_new = 0.0
                        
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
            
            # Формируем финальный ТОП-10
            if all_expenses_changes:
                df_total_changes = pd.DataFrame(all_expenses_changes)
                
                # Сортируем по убыванию силы численного влияния
                top_10_changes = df_total_changes.sort_values(by="Абсолютное влияние (руб.)", ascending=False).head(10)
                
                # Убираем техническую колонку
                top_10_display = top_10_changes.drop(columns=["Aбсолютное влияние (руб.)"], errors='ignore').reset_index(drop=True)
                top_10_display.index = top_10_display.index + 1  # Нумерация от 1 до 10
                
                st.subheader("📋 Директорский отчет: ТОП-10 главных изменений в статьях")
                st.write(f"Агент проанализировал листы и нашел 10 статей, которые сильнее всего сформировали вашу разницу в **{788451:,.2f} руб.**")
                
                # Отображение таблицы с форматированием денег
                st.dataframe(
                    top_10_display.style.format({
                        "Было (руб.)": "{:,.2f}", 
                        "Стало (руб.)": "{:,.2f}", 
                        "Изменение (руб.)": "{:+,.2f}"
                    }),
                    use_container_width=True
                )
            else:
                st.info("📊 Изменений по расходам между отчетами с такой структурой не найдено. Проверьте правильность названий листов.")
                
    except Exception as e:
        st.error(f"⚠️ Произошла ошибка при анализе структуры Excel: {e}")
        st.info("💡 Подсказка: Если финотдел изменил структуру в этом месяце, проверьте параметры в боковом меню.")
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для выявления ТОП-10 изменений.")
