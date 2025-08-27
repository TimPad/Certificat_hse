import streamlit as st
import pandas as pd
from docx import Document
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

# Кэширование для загрузки справочных данных
@st.cache_data
def load_reference_data(docx_content: bytes) -> Dict[str, Dict[str, str]]:
    """Загружает справочные данные из Word документа"""
    # Создаем временный файл
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
        tmp_file.write(docx_content)
        tmp_file_path = tmp_file.name
    
    try:
        doc = Document(tmp_file_path)
        table = doc.tables[0]
        
        grade_mapping = {}
        for row in table.rows[1:]:  # Пропускаем заголовок
            cells = row.cells
            discipline = cells[0].text.strip()
            
            # Удаляем префикс "Независимый экзамен по ", если есть
            if discipline.startswith("Независимый экзамен по "):
                discipline = discipline.replace("Независимый экзамен по ", "")
            discipline = discipline.strip()

            satisfactory = cells[1].text.strip()
            good = cells[2].text.strip()
            excellent = cells[3].text.strip()

            grade_mapping[discipline] = {
                '3': satisfactory,
                '4': good,
                '5': excellent
            }
        
        return grade_mapping
    finally:
        # Удаляем временный файл
        os.unlink(tmp_file_path)

def process_student_data(df: pd.DataFrame, grade_mapping: Dict[str, Dict[str, str]]) -> Tuple[pd.DataFrame, list]:
    """Обрабатывает данные студентов и возвращает результаты"""
    
    # Словарь сопоставления названий дисциплин
    discipline_name_mapping = {
        'Цифровая грамотность': 'цифровой грамотности',
        'Алгоритмическое мышление и программирование': 'программированию. Базовый уровень',
        'Анализу данных, искусственный интеллект и генеративные модели': 
            'анализу данных, искусственному интеллекту и генеративным моделям. Базовый уровень'
    }
    
    # Словарь сопоставления оценок
    grade_column_mapping = {
        'Удовлетворительно': '3',
        'Хорошо': '4',
        'Отлично': '5'
    }
    
    results = []
    processing_log = []
    
    # Определяем какие колонки использовать
    has_clean_columns = 'Название Дисциплины 1' in df.columns
    
    for index, row in df.iterrows():
        student_results = []
        student_name = row.iloc[0] if len(row) > 0 else f"Студент {index + 1}"
        
        processing_log.append(f"👤 Обрабатываем студента: {student_name}")
        
        # Обрабатываем каждую из трех дисциплин
        for discipline_num in range(1, 4):
            try:
                if has_clean_columns:
                    clean_disc_name_col = f"Название Дисциплины {discipline_num}"
                    if clean_disc_name_col in df.columns:
                        clean_discipline = str(row[clean_disc_name_col]).strip()
                    else:
                        continue
                else:
                    # Используем старую колонку
                    old_disc_name_col = f"Дисциплина {discipline_num}"
                    if old_disc_name_col not in df.columns:
                        continue
                    discipline_name = row[old_disc_name_col]
                    clean_discipline = str(discipline_name).strip()
                    if clean_discipline.startswith("Независимый экзамен по "):
                        clean_discipline = clean_discipline.replace("Независимый экзамен по ", "").strip()
                
                grade_5_col = f"Оценка 5 баллов Дисциплина {discipline_num}"
                if grade_5_col not in df.columns:
                    continue
                
                grade_value = row[grade_5_col]
                
                processing_log.append(f"  📚 Дисциплина {discipline_num}: '{clean_discipline}', Оценка: {grade_value}")
                
                # Пропускаем, если дисциплина или оценка отсутствуют
                if pd.isna(clean_discipline) or pd.isna(grade_value):
                    processing_log.append(f"    ⏭️ Пропускаем: отсутствует название дисциплины или оценка")
                    continue
                
                # Очищаем текст оценки
                grade_text = str(grade_value).strip()
                
                # Получаем ключ колонки для справочника
                if grade_text in grade_column_mapping:
                    grade_key = grade_column_mapping[grade_text]
                else:
                    processing_log.append(f"    ❌ Неизвестная оценка: {grade_text}")
                    continue
                
                # Ищем соответствующую дисциплину в справочнике
                mapped_discipline = None
                
                # Прямое сопоставление
                if clean_discipline in discipline_name_mapping:
                    mapped_discipline = discipline_name_mapping[clean_discipline]
                
                # Поиск в справочнике
                if not mapped_discipline:
                    for ref_discipline in grade_mapping.keys():
                        if clean_discipline.lower() in ref_discipline.lower() or ref_discipline.lower() in clean_discipline.lower():
                            mapped_discipline = ref_discipline
                            break
                
                if not mapped_discipline:
                    processing_log.append(f"    ❌ Дисциплина '{clean_discipline}' не найдена в справочнике")
                    continue
                
                # Получаем результирующий текст
                if mapped_discipline in grade_mapping and grade_key in grade_mapping[mapped_discipline]:
                    result_text = grade_mapping[mapped_discipline][grade_key]
                    
                    # Форматируем название дисциплины
                    formatted_discipline = clean_discipline.capitalize()
                    formatted_result = f"{formatted_discipline}:\n{result_text}"
                    
                    student_results.append(formatted_result)
                    processing_log.append(f"    ✅ Успешно обработано")
                else:
                    processing_log.append(f"    ❌ Не найден текст для оценки {grade_key}")
                    
            except Exception as e:
                processing_log.append(f"    ❌ Ошибка при обработке дисциплины {discipline_num}: {str(e)}")
        
        # Объединяем результаты студента
        final_result = "\n\n".join(student_results) if student_results else "Нет данных для обработки"
        results.append(final_result)
    
    # Добавляем результаты в DataFrame
    df_result = df.copy()
    df_result['Итоговый результат'] = results
    
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
        
        st.header("📋 Требования к файлам")
        st.markdown("""
        **Excel файл должен содержать:**
        - Колонки `Учащийся`
        - `Название Дисциплины 1/2/3` или `Дисциплина 1/2/3`
        - `Оценка 5 баллов Дисциплина 1/2/3`
        
        **Word файл должен содержать:**
        - Таблицу с критериями оценок
        - Колонки: Дисциплина, Удовлетворительно, Хорошо, Отлично
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
        st.subheader("📄 Загрузка справочника критериев")
        word_file = st.file_uploader(
            "Выберите Word файл со справочником",
            type=['docx'],
            help="Файл должен содержать таблицу с критериями для каждой оценки"
        )
    
    # Обработка файлов
    if excel_file and word_file:
        try:
            # Загружаем данные
            with st.spinner("📥 Загружаем файлы..."):
                df = pd.read_excel(excel_file)
                word_content = word_file.read()
                grade_mapping = load_reference_data(word_content)
            
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
            
            with st.expander("📋 Справочник критериев"):
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
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False)
                    
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
        st.info("📄 Загрузите также Word файл со справочником для продолжения")
    elif word_file:
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