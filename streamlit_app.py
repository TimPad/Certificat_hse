import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import tempfile
import os
from typing import Dict, Any, Optional, Tuple

# Настройка страницы
st.set_page_config(
    page_title="Система Обработки Сертификатов",
    page_icon="📜",
    layout="wide",
    initial_sidebar_state="expanded"
)

def deduplicate_lines(text):
    """Удаляет дублирующиеся строки из текста, сохраняя порядок первого появления"""
    if pd.isna(text) or not isinstance(text, str):
        return text
    
    lines = text.split('\n')
    seen_lines = set()
    unique_lines = []
    
    for line in lines:
        line_clean = line.strip()
        if line_clean and line_clean not in seen_lines:
            seen_lines.add(line_clean)
            unique_lines.append(line)
    
    return '\n'.join(unique_lines)

# Кэширование для загрузки справочных данных
@st.cache_data
def load_reference_data(skills_content: bytes) -> Dict[str, str]:
    """Загружает справочные данные из Excel файла с агрегированными навыками используя составные ключи"""
    # Создаем временный файл
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(skills_content)
        tmp_file_path = tmp_file.name
    
    try:
        skills_df = pd.read_excel(tmp_file_path)
        
        # Создаем словарь маппинга с составным ключом: {"дисциплина—уровень_оценки": описание_навыков}
        grade_mapping = {}
        cleaned_count = 0
        
        for _, row in skills_df.iterrows():
            discipline = row['Дисциплина']
            level = row['Уровень_оценки']
            description = row['Описание_навыков']
            
            # Удаляем дублирующиеся строки из описания навыков
            original_description = description
            clean_description = deduplicate_lines(description)
            
            if original_description != clean_description:
                cleaned_count += 1
            
            # Создаем составной ключ "дисциплина—уровень_оценки"
            composite_key = f"{discipline}—{level}"
            grade_mapping[composite_key] = clean_description
        
        # Логируем информацию об очистке (для отладки) - убрано для уменьшения логирования
        # if cleaned_count > 0:
        #     st.info(f"🧹 Очищено {cleaned_count} описаний от дубликатов")
        
        return grade_mapping
    finally:
        # Удаляем временный файл
        os.unlink(tmp_file_path)

def process_student_data(df: pd.DataFrame, grade_mapping: Dict[str, str]) -> Tuple[pd.DataFrame, list]:
    """Обрабатывает данные студентов используя составные ключи"""
    
    results = []
    processing_log = []
    
    # Упрощенное логирование
    processing_log.append(f"📊 Обрабатываем {len(df)} студентов")
    
    for index, row in df.iterrows():
        student_results = []
        processed_keys = set()
        
        for discipline_num in range(1, 4):
            discipline_col = f"Дисциплина {discipline_num}"
            grade_5_col = f"Оценка 5 баллов Дисциплина {discipline_num}"
            
            if discipline_col not in df.columns or grade_5_col not in df.columns:
                continue
                
            discipline_value = str(row[discipline_col]).strip()
            grade_value = str(row[grade_5_col]).strip()
            
            if pd.isna(discipline_value) or pd.isna(grade_value) or discipline_value == 'nan' or grade_value == 'nan':
                continue
            
            lookup_key = f"{discipline_value}—{grade_value}"
            
            if lookup_key in processed_keys:
                continue
            
            if lookup_key in grade_mapping:
                skill_description = grade_mapping[lookup_key]
                
                short_name_col = f"Название Дисциплины {discipline_num}"
                if short_name_col in df.columns:
                    display_name = str(row[short_name_col]).strip()
                    formatted_discipline = display_name.capitalize() if display_name != 'nan' and display_name else discipline_value
                else:
                    formatted_discipline = discipline_value
                
                formatted_result = f"📚 {formatted_discipline}:\n{skill_description}"
                student_results.append(formatted_result)
                processed_keys.add(lookup_key)
        
        final_result = "\n\n".join(student_results) if student_results else "Навыки не найдены."
        results.append(final_result)
    
    processing_log.append(f"✅ Успешно обработано")
    
    df_result = df.copy()
    df_result['Итоговый результат'] = results
    
    # Удаляем колонки, начинающиеся с "Название Дисциплины "
    columns_to_remove = [col for col in df_result.columns if col.startswith("Название Дисциплины ")]
    
    if columns_to_remove:
        df_result = df_result.drop(columns=columns_to_remove)
    
    return df_result, processing_log


def main():
    st.title("📜 Система Обработки Сертификатов")
    st.markdown("---")
    
    # Боковая панель с информацией
    with st.sidebar:
        st.header("ℹ️ О системе")
        st.info("""
        **Система автоматически:**
        - Обрабатывает данные экзаменов студентов
        - Сопоставляет оценки с критериями
        - Генерирует текст для сертификатов
        """)
        
        # Кнопка скачивания примера файла
        st.header("📥 Скачать примеры")
        
        # Загружаем пример Excel файла для скачивания
        try:
            # Используем путь относительно текущего файла
            current_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Excel пример
            excel_example_path = os.path.join(current_dir, 'Сертификаты пример.xlsx')
            
            if os.path.exists(excel_example_path):
                with open(excel_example_path, 'rb') as example_file:
                    excel_example_data = example_file.read()
                
                st.download_button(
                    label="📊 Скачать пример Excel файла",
                    data=excel_example_data,
                    file_name="Сертификаты пример.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Скачайте этот файл как шаблон для ваших данных студентов"
                )
            else:
                st.warning("⚠️ Пример Excel файла не найден в репозитории")
            
            # Excel справочник агрегированных навыков
            skills_example_path = os.path.join(current_dir, 'агрегированные_навыки.xlsx')
            
            if os.path.exists(skills_example_path):
                with open(skills_example_path, 'rb') as skills_file:
                    skills_data = skills_file.read()
                
                st.download_button(
                    label="📄 Скачать справочник навыков",
                    data=skills_data,
                    file_name="агрегированные_навыки.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Скачайте этот файл как пример справочника с агрегированными навыками"
                )
            else:
                st.warning("⚠️ Справочник навыков не найден в репозитории")
                
            st.success("💡 Используйте эти файлы как образцы для вашей работы!")
            
        except Exception as e:
            st.error(f"❌ Ошибка при загрузке примеров файлов: {str(e)}")
        
        st.header("📋 Требования к файлам")
        st.markdown("""
        **📊 Excel файл должен содержать:**
        - Колонки `Учащийся`
        - `Название Дисциплины 1/2/3` или `Дисциплина 1/2/3`
        - `Оценка 5 баллов Дисциплина 1/2/3`
        - Оценки: `Удовлетворительно`, `Хорошо`, `Отлично`
        
        **📄 Excel файл навыков должен содержать:**
        - Колонки: Дисциплина, Уровень_оценки, Описание_навыков
        - Уровни оценки: Удовлетворительно, Хорошо, Отлично
        - Текстовые описания навыков для каждой оценки
        
        💡 **Скачайте примеры выше для понимания формата!**
        """)
    
    # Основной интерфейс
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📊 Загрузка данных студентов")
        excel_file = st.file_uploader(
            "Выберите Excel файл с данными студентов",
            type=['xlsx', 'xls'],
            help="Файл должен содержать данные о студентах и их оценках"
        )
    
    with col2:
        st.subheader("📄 Загрузка справочника навыков")
        skills_file = st.file_uploader(
            "Выберите Excel файл с агрегированными навыками",
            type=['xlsx', 'xls'],
            help="Файл должен содержать агрегированные навыки с колонками: Дисциплина, Уровень_оценки, Описание_навыков"
        )
    
    # Обработка файлов
    if excel_file and skills_file:
        try:
            # Загружаем данные
            with st.spinner("📥 Загружаем файлы..."):
                df = pd.read_excel(excel_file)
                skills_content = skills_file.read()
                grade_mapping = load_reference_data(skills_content)
            
            st.success("✅ Файлы успешно загружены!")
            
            # Показываем информацию о данных
            st.subheader("📈 Информация о данных")
            
            col1, col2, col3 = st.columns([1, 1, 1])
            with col1:
                st.metric("Количество студентов", len(df))
            with col2:
                st.metric("Количество колонок", len(df.columns))
            with col3:
                st.metric("Дисциплин в справочнике", len(grade_mapping))
            
            # Превью данных
            with st.expander("👀 Просмотр данных Excel"):
                st.dataframe(df.head())
            
            with st.expander("📋 Справочник навыков"):
                ref_df = pd.DataFrame.from_dict(grade_mapping, orient='index')
                st.dataframe(ref_df)
            
            # Кнопка обработки
            if st.button("🚀 Обработать данные", type="primary"):
                with st.spinner("⚙️ Обрабатываем данные..."):
                    result_df, processing_log = process_student_data(df, grade_mapping)
                
                st.success("✅ Обработка завершена!")
                
                # Показываем результаты
                st.subheader("📊 Результаты обработки")
                
                # Вкладки для результатов
                tab1, tab2, tab3 = st.tabs(["📄 Результаты", "📋 Лог обработки", "💾 Скачать"])
                
                with tab1:
                    st.dataframe(result_df, use_container_width=True)
                
                with tab2:
                    st.text_area(
                        "Детальный лог обработки:",
                        value="\n".join(processing_log),
                        height=300
                    )
                
                with tab3:
                    # Подготовка файла для скачивания
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl', mode='w') as writer:
                        result_df.to_excel(writer, index=False)
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Скачать результаты Excel",
                        data=output.getvalue(),
                        file_name="Сертификаты_с_результатами.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("💡 Файл готов для скачивания!")
        
        except Exception as e:
            st.error(f"❌ Ошибка при обработке файлов: {str(e)}")
            st.exception(e)
    
    elif excel_file:
        st.info("📄 Загрузите также Excel файл с агрегированными навыками для продолжения")
    elif skills_file:
        st.info("📊 Загрузите также Excel файл с данными студентов для продолжения")
    else:
        st.info("📁 Загрузите оба файла для начала обработки")
    
    # Футер
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666;'>
            <p>Система Обработки Сертификатов v1.0.0 | 
            Создано с помощью Streamlit 🚀</p>
        </div>
        """, 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
