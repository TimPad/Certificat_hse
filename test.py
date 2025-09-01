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
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(skills_content)
        tmp_file_path = tmp_file.name

    try:
        skills_df = pd.read_excel(tmp_file_path)

        # Убираем дубли (по дисциплине и уровню)
        skills_df = skills_df.drop_duplicates(subset=["Дисциплина", "Уровень_оценки"])

        grade_mapping = {}

        for _, row in skills_df.iterrows():
            discipline = str(row['Дисциплина']).strip().lower()
            level = str(row['Уровень_оценки']).strip()
            description = str(row['Описание_навыков']).strip()

            if discipline not in grade_mapping:
                grade_mapping[discipline] = {}

            level_key_mapping = {
                'Удовлетворительно': '3',
                'Хорошо': '4',
                'Отлично': '5'
            }

            if level in level_key_mapping:
                grade_mapping[discipline][level_key_mapping[level]] = description

        return grade_mapping
    finally:
        os.unlink(tmp_file_path)


def process_student_data(df: pd.DataFrame, grade_mapping: Dict[str, Dict[str, str]]) -> Tuple[pd.DataFrame, list]:
    """Обрабатывает данные студентов и возвращает результаты"""

    grade_column_mapping = {
        'Удовлетворительно': '3',
        'Хорошо': '4',
        'Отлично': '5'
    }

    results = []
    processing_log = []

    processing_log.append(f"📊 Начинаем обработку {len(df)} студентов...")
    processing_log.append(f"🗂️ Найдено дисциплин в справочнике: {len(grade_mapping)}")

    for index, row in df.iterrows():
        student_results = []
        student_email = row.get('Почта', f"Студент {index + 1}")

        processing_log.append(f"\n👤 Обрабатываем студента: {student_email}")

        discipline_results = {}

        for discipline_num in range(1, 4):
            try:
                discipline_col = f"Дисциплина {discipline_num}"
                grade_5_col = f"Оценка 5 баллов Дисциплина {discipline_num}"

                if discipline_col not in df.columns or grade_5_col not in df.columns:
                    continue

                full_discipline = str(row[discipline_col]).strip().lower() if pd.notna(row[discipline_col]) else ""
                grade_value = row[grade_5_col]

                if not full_discipline or pd.isna(grade_value):
                    continue

                clean_grade = str(grade_value).strip()

                if clean_grade not in grade_column_mapping:
                    continue

                grade_key = grade_column_mapping[clean_grade]

                display_name_col = f"Название Дисциплины {discipline_num}"
                if display_name_col in df.columns and pd.notna(row[display_name_col]):
                    display_name = str(row[display_name_col]).strip()
                else:
                    display_name = row[discipline_col]

                if full_discipline not in grade_mapping:
                    continue

                if grade_key not in grade_mapping[full_discipline]:
                    continue

                if full_discipline in discipline_results:
                    existing_grade = discipline_results[full_discipline][0]
                    if int(grade_key) > int(existing_grade):
                        result_text = grade_mapping[full_discipline][grade_key]
                        discipline_results[full_discipline] = (grade_key, result_text, display_name)
                else:
                    result_text = grade_mapping[full_discipline][grade_key]
                    discipline_results[full_discipline] = (grade_key, result_text, display_name)

            except Exception as e:
                processing_log.append(f"    ❌ Ошибка при обработке дисциплины {discipline_num}: {str(e)}")

        student_results = []
        for discipline, (grade, description, display_name) in discipline_results.items():
            formatted_result = f"{display_name}:\n{description}"
            student_results.append(formatted_result)

        # убираем дубликаты по тексту результата
        student_results = list(dict.fromkeys(student_results))

        final_result = "\n\n".join(student_results) if student_results else ""
        results.append(final_result)

    df_result = df.copy()
    df_result['Итоговый результат'] = results

    columns_to_drop = [col for col in df_result.columns if col.startswith('Название Дисциплины ')]
    if columns_to_drop:
        df_result = df_result.drop(columns=columns_to_drop)

    return df_result, processing_log


def main():
    st.title("📜 Система Обработки Сертификатов")
    st.markdown("---")

    with st.sidebar:
        st.header("ℹ️ О системе")
        st.info("""
        **Система автоматически:**
        - Обрабатывает данные экзаменов студентов
        - Сопоставляет оценки с критериями
        - Генерирует текст для сертификатов
        """)

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📊 Загрузка данных студентов")
        excel_file = st.file_uploader("Выберите Excel файл с данными студентов", type=['xlsx', 'xls'])

    with col2:
        st.subheader("📄 Загрузка справочника навыков")
        skills_file = st.file_uploader("Выберите Excel файл с агрегированными навыками", type=['xlsx', 'xls'])

    if excel_file and skills_file:
        try:
            df = pd.read_excel(excel_file)
            skills_content = skills_file.read()
            grade_mapping = load_reference_data(skills_content)

            if st.button("🚀 Обработать данные", type="primary"):
                result_df, processing_log = process_student_data(df, grade_mapping)

                st.success("✅ Обработка завершена!")
                st.subheader("📊 Результаты обработки")
                st.dataframe(result_df, width="stretch")

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)

                st.download_button(
                    label="📥 Скачать результаты Excel",
                    data=output.getvalue(),
                    file_name="Сертификаты_с_результатами.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ Ошибка при обработке файлов: {str(e)}")


if __name__ == "__main__":
    main()
