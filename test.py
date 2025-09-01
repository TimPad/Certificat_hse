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
            if 'Дисциплина' in skills_df.columns and 'Уровень_оценки' in skills_df.columns and 'Описание_навыков' in skills_df.columns:
                discipline = str(row['Дисциплина']).strip() if pd.notna(row['Дисциплина']) else ''
                level = str(row['Уровень_оценки']).strip() if pd.notna(row['Уровень_оценки']) else ''
                description = str(row['Описание_навыков']).strip() if pd.notna(row['Описание_навыков']) else ''
                
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
    
    # Ограничиваем размер лога для избежания ошибок WebSocket
    max_log_entries = 1000
    
    processing_log.append(f"📊 Начинаем обработку {len(df)} студентов...")
    processing_log.append(f"🗂️ Найдено дисциплин в справочнике: {len(grade_mapping)}")
    processing_log.append(f"📁 Колонки в Excel файле: {list(df.columns)}")
    
    log_count = len(processing_log)
    
    for index, row in df.iterrows():
        # Проверяем лимит лога
        if log_count >= max_log_entries:
            processing_log.append("⚠️ Лог обрезан для предотвращения ошибок (слишком много записей)")
            break
            
        student_results = []
        student_email = str(row['Почта']).strip() if 'Почта' in df.columns and pd.notna(row['Почта']) and str(row['Почта']).strip() else f"Студент {index + 1}"
        
        processing_log.append(f"👤 Студент {index + 1}: {student_email}")
        log_count += 1
        
        # Множество для отслеживания уже обработанных пар (дисциплина, оценка)
        processed_pairs = set()
        
        # Обрабатываем каждую из трех дисциплин
        for discipline_num in range(1, 4):
            if log_count >= max_log_entries:
                break
                
            try:
                discipline_col = f"Дисциплина {discipline_num}"
                grade_5_col = f"Оценка 5 баллов Дисциплина {discipline_num}"
                
                # Проверяем наличие колонок
                if discipline_col not in df.columns or grade_5_col not in df.columns:
                    continue
                
                full_discipline = str(row[discipline_col]).strip() if pd.notna(row[discipline_col]) else ""
                grade_value = row[grade_5_col]
                
                # Пропускаем пустые значения
                if not full_discipline or pd.isna(grade_value):
                    continue
                
                clean_grade = str(grade_value).strip()
                
                # Проверяем, известна ли оценка
                valid_grades = ['Удовлетворительно', 'Хорошо', 'Отлично']
                if clean_grade not in valid_grades:
                    continue
                
                # Ключ: (дисциплина, оценка)
                discipline_grade_pair = (full_discipline, clean_grade)
                
                # Проверяем, была ли эта комбинация уже обработана
                if discipline_grade_pair in processed_pairs:
                    continue
                
                # Проверяем наличие в справочнике
                if full_discipline not in grade_mapping:
                    continue
                
                if clean_grade not in grade_mapping[full_discipline]:
                    continue
                
                # Добавляем форматированный навык в список
                result_text = grade_mapping[full_discipline][clean_grade]
                formatted_result = f"- {full_discipline} ({clean_grade})"
                student_results.append(formatted_result)
                processed_pairs.add(discipline_grade_pair)
                
            except Exception:
                continue
        
        # Формируем итоговый результат
        final_result = "\n".join(student_results) if student_results else ""
        results.append(final_result)
    
    processing_log.append(f"✅ Обработка завершена: обработано {len(df)} студентов")
    
    # Создаём результирующий DataFrame
    df_result = df.copy()
    df_result['Итоговый результат'] = results
    
    # Удаляем старые колонки с названиями дисциплин (если остались)
    columns_to_drop = [col for col in df_result.columns if col.startswith('Название Дисциплины ')]
    if columns_to_drop:
        df_result = df_result.drop(columns=columns_to_drop)
    
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
        **📊 Excel файл должен содержать:**
        - Колонку `Почта`
        - `Дисциплина 1/2/3`
        - `Оценка 5 баллов Дисциплина 1/2/3`
        - Оценки: `Удовлетворительно`, `Хорошо`, `Отлично`
        
        **📄 Excel файл навыков должен содержать:**
        - Колонки: Дисциплина, Уровень_оценки, Описание_навыков
        - Уровни оценки: Удовлетворительно, Хорошо, Отлично
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
                skills_content = skills_file.getvalue()
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
                            'Описание': description[:100] + "..." if len(description) > 100 else description
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
                    # Ограничиваем размер отображаемого лога
                    log_text = "\n".join(processing_log[-500:])  # Показываем последние 500 записей
                    if len(processing_log) > 500:
                        log_text = f"⚠️ Показаны последние 500 записей из {len(processing_log)}\n\n" + log_text
                    
                    st.text_area(
                        "Детальный лог обработки:",
                        value=log_text,
                        height=300
                    )
                
                with tab3:
                    # Подготовка файла для скачивания
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='Результаты')
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
    
    elif excel_file:
        st.info("📄 Загрузите также Excel файл с агрегированными навыками для продолжения")
    elif skills_file:
        st.info("📊 Загрузите также Excel файл с данными студентов для продолжения")
    else:
        st.info("📁 Загрузите оба файла для начала обработки")

if __name__ == "__main__":
    main()
