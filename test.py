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
        # Используем текстовые уровни напрямую ('Удовлетворительно', 'Хорошо', 'Отлично')
        grade_mapping = {}
        
        for _, row in skills_df.iterrows():
            discipline = str(row['Дисциплина']).strip() if 'Дисциплина' in skills_df.columns else ''
            level = str(row['Уровень_оценки']).strip() if 'Уровень_оценки' in skills_df.columns else ''
            description = str(row['Описание_навыков']).strip() if 'Описание_навыков' in skills_df.columns else ''
            
            # Пропускаем пустые строки
            if not discipline or not level or not description:
                continue
                
            if discipline not in grade_mapping:
                grade_mapping[discipline] = {}
            
            if level in ['Удовлетворительно', 'Хорошо', 'Отлично']:
                grade_mapping[discipline][level] = description
        
        return grade_mapping
    finally:
        # Удаляем временный файл
        try:
            os.unlink(tmp_file_path)
        except:
            pass

def process_student_data(df: pd.DataFrame, grade_mapping: Dict[str, Dict[str, str]]) -> Tuple[pd.DataFrame, list]:
    """Обрабатывает данные студентов и возвращает результаты с новой логикой"""
    
    results = []
    processing_log = []
    
    processing_log.append(f"📊 Начинаем обработку {len(df)} студентов...")
    processing_log.append(f"🗂️ Найдено дисциплин в справочнике: {len(grade_mapping)}")
    processing_log.append(f"📋 Дисциплины в справочнике: {list(grade_mapping.keys())}")
    processing_log.append(f"📁 Колонки в Excel файле: {list(df.columns)}")
    
    for index, row in df.iterrows():
        student_results = []  # List для навыков с оформлением
        student_email = str(row['Почта']).strip() if 'Почта' in df.columns and pd.notna(row['Почта']) else f"Студент {index + 1}"
        
        processing_log.append(f"\n👤 Обрабатываем студента: {student_email}")
        
        # Множество для отслеживания уже обработанных пар (дисциплина, оценка)
        processed_pairs = set()
        
        # Обрабатываем каждую из трех дисциплин
        for discipline_num in range(1, 4):
            try:
                discipline_col = f"Дисциплина {discipline_num}"
                grade_5_col = f"Оценка 5 баллов Дисциплина {discipline_num}"
                
                # Проверяем наличие колонок
                if discipline_col not in df.columns or grade_5_col not in df.columns:
                    processing_log.append(f"    ⚠️ Пропускаем дисциплину {discipline_num}: колонки не найдены")
                    continue
                
                full_discipline = str(row[discipline_col]).strip() if pd.notna(row[discipline_col]) else ""
                grade_value = row[grade_5_col]
                
                # Пропускаем пустые значения
                if not full_discipline or pd.isna(grade_value):
                    processing_log.append(f"    ⏭️ Пропускаем: пустая дисциплина или оценка (дисциплина {discipline_num})")
                    continue
                
                clean_grade = str(grade_value).strip()
                
                # Проверяем, известна ли оценка
                valid_grades = ['Удовлетворительно', 'Хорошо', 'Отлично']
                if clean_grade not in valid_grades:
                    processing_log.append(f"    ❌ Неизвестная оценка: '{clean_grade}' (ожидалось: {', '.join(valid_grades)})")
                    continue
                
                # Ключ: (дисциплина, оценка)
                discipline_grade_pair = (full_discipline, clean_grade)
                
                # Проверяем, была ли эта комбинация уже обработана
                if discipline_grade_pair in processed_pairs:
                    processing_log.append(f"    ⚠️ Пропускаем дублированную комбинацию: '{full_discipline}' с оценкой '{clean_grade}'")
                    continue
                
                # Проверяем наличие в справочнике
                if full_discipline not in grade_mapping:
                    processing_log.append(f"    ❌ Дисциплина '{full_discipline}' не найдена в справочнике")
                    continue
                
                if clean_grade not in grade_mapping[full_discipline]:
                    processing_log.append(f"    ⚠️ Нет описания навыков для оценки '{clean_grade}' по дисциплине '{full_discлина}'")
                    continue
                
                # Добавляем форматированный навык в список
                result_text = grade_mapping[full_discipline][clean_grade]
                formatted_result = f"- {full_discipline} ({clean_grade}): {result_text}"
                student_results.append(formatted_result)
                processed_pairs.add(discipline_grade_pair)  # Отмечаем как обработанную
                
                processing_log.append(f"    ✅ Добавлено: '{full_discipline}' с оценкой '{clean_grade}'")
                
            except Exception as e:
                processing_log.append(f"    ❌ Ошибка при обработке дисциплины {discipline_num}: {str(e)}")
        
        # Формируем итоговый результат (список с переносами строк)
        final_result = "\n".join(student_results) if student_results else ""
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
            current_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
            
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
            help="Файл должен содержать данные о студентах и их оценках",
            key="excel_file"
        )
    
    with col2:
        st.subheader("📄 Загрузка справочника навыков")
        skills_file = st.file_uploader(
            "Выберите Excel файл с агрегированными навыками",
            type=['xlsx', 'xls'],
            help="Файл должен содержать агрегированные навыки с колонками: Дисциплина, Уровень_оценки, Описание_навыков",
            key="skills_file"
        )
    
    # Обработка файлов
    if excel_file and skills_file:
        try:
            # Загружаем данные
            with st.spinner("📥 Загружаем файлы..."):
                df = pd.read_excel(excel_file)
                skills_content = skills_file.getvalue()  # Используем getvalue() вместо read()
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
                # Преобразуем словарь в DataFrame для отображения
                ref_data = []
                for discipline, levels in grade_mapping.items():
                    for level, description in levels.items():
                        ref_data.append({
                            'Дисциплина': discipline,
                            'Уровень': level,
                            'Описание': description
                        })
                ref_df = pd.DataFrame(ref_data)
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
                        result_df.to_excel(writer, index=False, sheet_name='Результаты')
                    output.seek(0)  # Важно: перемещаем указатель в начало
                    
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
