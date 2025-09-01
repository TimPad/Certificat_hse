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

# Кэширование для загрузки справочных данных
@st.cache_data
def load_reference_data(skills_content: bytes) -> Dict[str, Dict[str, str]]:
    """Загружает справочные данные из Excel файла с агрегированными навыками"""
    # Создаем временный файл
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(skills_content)
        tmp_file_path = tmp_file.name
    
    try:
        skills_df = pd.read_excel(tmp_file_path)
        
        # Создаем словарь маппинга: {дисциплина: {уровень_оценки: описание_навыков}}
        grade_mapping = {}
        
        for _, row in skills_df.iterrows():
            discipline = row['Дисциплина']
            level = row['Уровень_оценки']
            description = row['Описание_навыков']
            
            if discipline not in grade_mapping:
                grade_mapping[discipline] = {}
            
            # Маппим уровни оценки на ключи для совместимости
            level_key_mapping = {
                'Удовлетворительно': '3',
                'Хорошо': '4',
                'Отлично': '5'
            }
            
            if level in level_key_mapping:
                grade_mapping[discipline][level_key_mapping[level]] = description
        
        return grade_mapping
    finally:
        # Удаляем временный файл
        os.unlink(tmp_file_path)

def process_student_data(df: pd.DataFrame, grade_mapping: Dict[str, Dict[str, str]]) -> Tuple[pd.DataFrame, list]:
    """Обрабатывает данные студентов и возвращает результаты"""
    
    # Словарь сопоставления оценок
    grade_column_mapping = {
        'Удовлетворительно': '3',
        'Хорошо': '4',
        'Отлично': '5'
    }
    
    results = []
    processing_log = []
    
    processing_log.append(f"📊 Начинаем обработку {len(df)} студентов...")
    processing_log.append(f"🗂️ Найдено дисциплин в справочнике: {len(grade_mapping)}")
    processing_log.append(f"📋 Дисциплины в справочнике: {list(grade_mapping.keys())}")
    processing_log.append(f"📁 Колонки в Excel файле: {list(df.columns)}")
    
    for index, row in df.iterrows():
        student_results = []
        student_email = row['Почта'] if 'Почта' in df.columns else f"Студент {index + 1}"
        
        processing_log.append(f"\n👤 Обрабатываем студента: {student_email}")
        
        # Множество для отслеживания уже обработанных дисциплин (во избежание дублирования)
        processed_disciplines = set()
        
        # Обрабатываем каждую из трех дисциплин
        for discipline_num in range(1, 4):
            try:
                discipline_col = f"Дисциплина {discipline_num}"
                grade_5_col = f"Оценка 5 баллов Дисциплина {discipline_num}"
                
                # Проверяем наличие колонок
                if discipline_col not in df.columns or grade_5_col not in df.columns:
                    processing_log.append(f"    ⚠️ Пропускаем дисциплину {discipline_num}: колонки не найдены")
                    continue
                
                full_discipline = str(row[discipline_col]).strip()
                grade_value = row[grade_5_col]
                
                # Пропускаем пустые значения
                if pd.isna(full_discipline) or pd.isna(grade_value) or not full_discipline:
                    processing_log.append(f"    ⏭️ Пропускаем: пустая дисциплина или оценка (дисциплина {discipline_num})")
                    continue
                
                clean_grade = str(grade_value).strip()
                
                # Проверяем, известна ли оценка
                if clean_grade not in grade_column_mapping:
                    processing_log.append(f"    ❌ Неизвестная оценка: '{clean_grade}' (ожидалось: Удовлетворительно/Хорошо/Отлично)")
                    continue
                
                grade_key = grade_column_mapping[clean_grade]
                
                # Проверяем, не была ли дисциплина уже обработана
                if full_discipline in processed_disciplines:
                    processing_log.append(f"    ⚠️ Пропускаем дублированную дисциплину: '{full_discipline}'")
                    continue
                
                # Проверяем наличие дисциплины в справочнике
                if full_discipline not in grade_mapping:
                    processing_log.append(f"    ❌ Дисциплина '{full_discipline}' не найдена в справочнике")
                    continue
                
                if grade_key not in grade_mapping[full_discipline]:
                    processing_log.append(f"    ⚠️ Нет описания навыков для оценки '{clean_grade}' по дисциплине '{full_discipline}'")
                    continue
                
                # Добавляем результат
                result_text = grade_mapping[full_discipline][grade_key]
                formatted_result = f"{full_discipline}:\n{result_text}"
                student_results.append(formatted_result)
                processed_disciplines.add(full_discipline)  # Отмечаем как обработанную
                
                processing_log.append(f"    ✅ Добавлено: '{full_discipline}' с оценкой '{clean_grade}'")
                
            except Exception as e:
                processing_log.append(f"    ❌ Ошибка при обработке дисциплины {discipline_num}: {str(e)}")
        
        # Формируем итоговый результат (разделение двойным переносом)
        final_result = "\n\n".join(student_results) if student_results else ""
        results.append(final_result)
        
        if final_result:
            processing_log.append(f"  🎯 Итоговый результат:\n{final_result}")
        else:
            processing_log.append(f"  🎯 Итоговый результат: пусто")
    
    processing_log.append(f"\n✅ Обработка завершена: обработано {len(df)} студентов")
    
    # Создаём результирующий DataFrame
    df_result = df.copy()
    df_result['Итоговый результат'] = results
    
    # Удаляем старые колонки с названиями дисциплин (если остались)
    columns_to_drop = [col for col in df_result.columns if col.startswith('Название Дисциплины ')]
    if columns_to_drop:
        df_result = df_result.drop(columns=columns_to_drop)
        processing_log.append(f"🧹 Удалены колонки: {columns_to_drop}")
    
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
        - Колонку `Почта`
        - `Дисциплина 1/2/3`
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
                    st.dataframe(result_df, width="stretch")
                
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
