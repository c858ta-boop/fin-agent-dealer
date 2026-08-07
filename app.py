import streamlit as st
import pandas as pd

st.set_page_config(page_title="Финансовый ИИ-Агент", layout="wide")

st.title("🚗 Финансовый ИИ-Агент Дилерского Центра")
st.write("Агент сверяет два отчета Excel, находит отклонения > 10% от финрезультата и делает директорское резюме.")

# Блок загрузки файлов
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 Загрузите СТАРЫЙ отчет (прошлый месяц)", type=["xlsx"])
with col2:
    new_file = st.file_uploader("📂 Загрузите НОВЫЙ отчет (текущий месяц)", type=["xlsx"])

if old_file and new_file:
    st.success("Файлы успешно загружены! Начинаю финансовый аудит...")
    
    try:
        # Читаем списки листов
        old_excel = pd.ExcelFile(old_file)
        new_excel = pd.ExcelFile(new_file)
        
        st.subheader("📊 Анализ изменений по листам:")
        
        # Общие листы для сравнения
        common_sheets = set(old_excel.sheet_names).intersection(set(new_excel.sheet_names))
        
        for sheet in common_sheets:
            with st.expander(f"Лист: {sheet}"):
                df_old = pd.read_excel(old_file, sheet_name=sheet)
                df_new = pd.read_excel(new_file, sheet_name=sheet)
                
                st.write("Старый отчет (первые строки):", df_old.head(2))
                st.write("Новый отчет (первые строки):", df_new.head(2))
                
                # ТЗ для ИИ (заглушка, пока не подключили ключ)
                st.info("🤖 ИИ-Аналитика: Здесь будет выводиться текстовый разбор аномалий от нейросети LLaMA.")
                
    except Exception as e:
        st.error(f"Ошибка при чтении файлов: {e}. Убедитесь, что структура файлов совпадает.")
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для начала анализа.")
