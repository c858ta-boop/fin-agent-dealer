import streamlit as st
import pandas as pd

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый Автономный Агент Дилерского Центра")
st.write("Инструмент выявления ТОП-10 статей расходов, оказавших наибольшее численное влияние на финансовый результат.")

# Блок инструкций для настройки соответствия в боковой панели
with st.sidebar:
    st.header("⚙️ Настройки анализа")
    target_column = st.text_input("Название столбца со статьями расходов:", value="Статья")
    value_column = st.text_input("Название столбца со значениями (суммами):", value="Сумма")
    st.caption("ℹ️ Убедитесь, что в ваших Excel файлах названия этих столбцов совпадают.")

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
            # Сюда мы соберем изменения вообще по всем листам, чтобы найти ТОП-10 по всему предприятию
            all_expenses_changes = []
            
            for sheet in common_sheets:
                df_old = pd.read_excel(old_file, sheet_name=sheet)
                df_new = pd.read_excel(new_file, sheet_name=sheet)
                
                # Очищаем названия столбцов от пробелов
                df_old.columns = [str(c).strip() for c in df_old.columns]
                df_new.columns = [str(c).strip() for c in df_new.columns]
                
                if target_column in df_old.columns and value_column in df_old.columns and target_column in df_new.columns and value_column in df_new.columns:
                    
                    # Переводим данные в словари {Статья: Сумма}
                    dict_old = pd.Series(df_old[value_column].values, index=df_old[target_column]).to_dict()
                    dict_new = pd.Series(df_new[value_column].values, index=df_new[target_column]).to_dict()
                    
                    # Объединяем все уникальные статьи расходов на листе
                    sheet_articles = set(dict_old.keys()).union(set(dict_new.keys()))
                    
                    for article in sheet_articles:
                        # Пропускаем пустые строки или строки итогов (если они называются "Итого", "Всего" и т.д.)
                        if pd.isna(article) or str(article).strip() == "" or any(word in str(article).lower() for word in ["итого", "всего", "баланс", "результат"]):
                            continue
                            
                        val_old = float(dict_old.get(article, 0)) if pd.notna(dict_old.get(article, 0)) else 0.0
                        val_new = float(dict_new.get(article, 0)) if pd.notna(dict_new.get(article, 0)) else 0.0
                        
                        item_delta = val_new - val_old
                        # Нам нужно абсолютное влияние (неважно, вырос расход или упал, главное — на какую сумму)
                        abs_delta = abs(item_delta)
                        
                        if abs_delta > 0:
                            all_expenses_changes.append({
                                "Лист": sheet,
                                "Статья расходов": article,
                                "Было (руб.)": val_old,
                                "Стало (руб.)": val_new,
                                "Изменение (руб.)": item_delta,
                                "Абсолютное влияние (руб.)": abs_delta
                            })
            
            # Если изменения найдены, формируем финальный ТОП-10
            if all_expenses_changes:
                df_total_changes = pd.DataFrame(all_expenses_changes)
                
                # Сортируем по убыванию абсолютного влияния и берем первые 10 строк
                top_10_changes = df_total_changes.sort_values(by="Абсолютное влияние (руб.)", ascending=False).head(10)
                
                # Удаляем временную колонку абсолютного значения перед показом директору
                top_10_display = top_10_changes.drop(columns=["Абсолютное влияние (руб.)"]).reset_index(drop=True)
                top_10_display.index = top_10_display.index + 1  # Чтобы нумерация шла от 1 до 10
                
                st.subheader("📋 Директорский отчет: ТОП-10 главных изменений в статьях")
                st.write("Ниже представлены 10 статей, изменения по которым сильнее всего отразились на итоговом бюджете предприятия:")
                
                # Выводим красивую интерактивную таблицу
                st.dataframe(
                    top_10_display.style.format({
                        "Было (руб.)": "{:,.2f}", 
                        "Стало (руб.)": "{:,.2f}", 
                        "Изменение (руб.)": "{:+,.2f}" # Покажет плюс или минус перед цифрой изменения
                    }),
                    use_container_width=True
                )
            else:
                st.info("📊 Изменений по расходам между отчетами не найдено.")
                
    except Exception as e:
        st.error(f"⚠️ Произошла ошибка при анализе структуры Excel: {e}")
        st.info("💡 Подсказка: Проверьте правильность названий ключевых столбцов в левом боковом меню.")
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для выявления ТОП-10 изменений.")
