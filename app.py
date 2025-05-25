import streamlit as st
from pathlib import Path
from docx import Document
from parser import parse_document
from build_document import build_document
import os
import time

# Минималистичные стили для оформления
st.markdown("""
    <style>
    .main-title {
        font-size: 2rem;
        font-weight: bold;
        color: #333;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .section-header {
        font-size: 1.2rem;
        font-weight: 600;
        color: #444;
        margin-top: 1.5rem;
        margin-bottom: 0.5rem;
    }
    .stats-item {
        font-size: 1rem;
        color: #555;
        margin: 0.3rem 0;
    }
    div[data-testid="stDownloadButton"] button {
        background-color: #4a90e2;
        color: white !important;
        padding: 0.5rem 1rem;
        font-weight: 500;
        border-radius: 3px;
        width: 100%;
        border: none;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #3a78c2;
        color: white !important;
    }
    div[data-testid="stButton"] button {
        background-color: #4a90e2;
        color: white !important;
        padding: 0.5rem 1rem;
        font-weight: 500;
        border-radius: 3px;
        width: 100%;
        border: none;
    }
    div[data-testid="stButton"] button:hover {
        background-color: #3a78c2;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# Заголовок приложения
st.markdown('<div class="main-title">Автоматизация форматирования документов по ГОСТ</div>', unsafe_allow_html=True)

# Настройки парсинга
st.markdown('<div class="section-header">Настройки парсинга</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    use_markers = st.checkbox("Использовать маркеры (например, [H1], [OL]) для классификации", value=False)
with col2:
    use_regex = st.checkbox("Использовать регулярные выражения для классификации", value=True)

# Настройки форматирования
st.markdown('<div class="section-header">Настройки форматирования</div>', unsafe_allow_html=True)
col3, col4 = st.columns(2)
with col3:
    replace_quotes = st.checkbox("Заменить кавычки на «» (ёлочки)", value=True)
    add_space_before = st.checkbox("Добавить пустую строку перед таблицами и рисунками", value=True)
with col4:
    add_space_after = st.checkbox("Добавить пустую строку после таблиц и рисунков", value=True)
    format_tables = st.checkbox("Применить форматирование к таблицам (стиль Table Grid)", value=True)
add_space_after_headings = st.checkbox("Добавить пустую строку после заголовков всех уровней", value=True)
add_space_before_headings_2_to_4 = st.checkbox("Добавить пустую строку перед заголовками", value=True)

# Инструкция по использованию
with st.expander("Инструкция по использованию"):
    st.markdown("""
    ### Общие шаги по использованию приложения
    1. **Загрузите документ**: Нажмите на кнопку "Загрузите DOCX-файл" и выберите файл в формате `.docx`, который вы хотите отформатировать.
    2. **Настройте параметры**:
       - Включите или отключите опции парсинга ("Использовать маркеры" и "Использовать регулярные выражения").
       - Настройте параметры форматирования (замена кавычек, добавление пустых строк, форматирование таблиц и т.д.).
    3. **Нажмите "Форматировать"**: После настройки параметров нажмите кнопку "Форматировать", чтобы начать обработку. Дождитесь завершения (это займёт несколько секунд).
    4. **Просмотрите статистику**: После обработки вы увидите статистику документа (время выполнения, количество абзацев, таблиц, рисунков, приложений).
    5. **Скачайте результат**: Нажмите на кнопку "Скачать отформатированный документ", чтобы сохранить обработанный файл в формате `.docx`.

    ### Использование маркеров для форматирования
    Добавьте маркеры в начале строк в документе, чтобы указать тип элемента:
    - `[H1]`, `[H2]`, `[H3]`, `[H4]` — для заголовков 1–4 уровня.
    - `[OL]`, `[OL:1]` — для нумерованных списков (уровень 0, 1).
    - `[UL]`, `[UL:1]` — для маркированных списков.
    - `[TABLE]` — для подписи таблицы.
    - `[FIGURE]` — для подписи рисунка.

    Маркеры будут удалены, а текст отформатирован по ГОСТ 7.32–2017.

    **Пример:**
    ```
    [H1] Введение
    [UL] Тест
    [OL:1] Подпункт
    [TABLE] Пример структуры
    ```

    **Изображения**: Вставьте изображение (PNG/JPEG) в документ через Word (Вставка → Изображения → Этот компьютер).  
    Подпись укажите в следующем параграфе через `[FIGURE]` или текст, начинающийся с "Рисунок".  
    **Пример:**
    ```
    <Вставьте изображение>
    [FIGURE] Схема процесса
    ```
    """)

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите DOCX-файл", type=["docx"])

# Инициализация переменных для результатов
if 'formatted' not in st.session_state:
    st.session_state.formatted = False
    st.session_state.elements = None
    st.session_state.stats = None
    st.session_state.parse_time = None
    st.session_state.format_duration = None
    st.session_state.total_time = None

# Кнопка "Форматировать"
if uploaded_file:
    temp_path = Path("temp.docx")
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getvalue())

    if st.button("Форматировать", use_container_width=True):
        with st.spinner("Обработка документа..."):
            # Измеряем время парсинга
            start_parse = time.time()
            st.session_state.elements, st.session_state.stats = parse_document(
                temp_path,
                use_markers=use_markers,
                use_regex=use_regex
            )
            end_parse = time.time()
            st.session_state.parse_time = end_parse - start_parse

            # Измеряем время форматирования
            start_format = time.time()
            doc = Document(temp_path)
            formatted_doc = build_document(
                st.session_state.elements,
                doc=doc,
                output_path="formatted_output.docx",
                config={
                    "replace_quotes": replace_quotes,
                    "add_space_before": add_space_before,
                    "add_space_after": add_space_after,
                    "format_tables": format_tables,
                    "add_space_after_headings": add_space_after_headings,
                    "add_space_before_headings_2_to_4": add_space_before_headings_2_to_4  # Добавили новую настройку
                }
            )
            end_format = time.time()
            st.session_state.format_duration = end_format - start_format

            # Подсчёт общего количества абзацев
            paragraph_types = ["heading_1", "heading_2", "heading_3", "heading_4", "list_ordered", "list_bullet", "paragraph", "formula", "table_caption", "figure_caption"]
            total_paragraphs = sum(1 for el in st.session_state.elements if el["type"] in paragraph_types)

            # Подсчёт приложений (заголовков, начинающихся с "ПРИЛОЖЕНИЯ" или "ПРИЛОЖЕНИЕ")
            appendix_count = sum(1 for el in st.session_state.elements if el["type"] == "heading_1" and el.get("text", "").upper().strip().startswith(("ПРИЛОЖЕНИ", "ПРИЛОЖЕНИЕ")))

            # Форматируем время для отображения
            def format_time(seconds: float) -> str:
                if seconds < 1:
                    return f"{seconds * 1000:.0f} мс"
                elif seconds < 60:
                    return f"{seconds:.2f} сек"
                else:
                    minutes = int(seconds // 60)
                    remaining_secs = seconds % 60
                    return f"{minutes} мин {remaining_secs:.2f} сек"

            parse_time_display = format_time(st.session_state.parse_time)
            format_time_display = format_time(st.session_state.format_duration)
            st.session_state.total_time = st.session_state.parse_time + st.session_state.format_duration
            total_time_display = format_time(st.session_state.total_time)

            # Отображение времени выполнения
            st.markdown('<div class="section-header">Время выполнения</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="stats-item">Время парсинга: {parse_time_display}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="stats-item">Время форматирования: {format_time_display}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="stats-item">Общее время: {total_time_display}</div>', unsafe_allow_html=True)

            # Отображение статистики
            st.markdown('<div class="section-header">Статистика документа</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="stats-item">Общее количество абзацев: {total_paragraphs}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="stats-item">Таблицы: {st.session_state.stats["tables"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="stats-item">Рисунки: {st.session_state.stats["figures"]}</div>', unsafe_allow_html=True)

            # Дерево элементов
            with st.expander("Дерево элементов", expanded=False):
                for el in st.session_state.elements:
                    kind = el.get("type", "paragraph")
                    preview = el.get("text") or el.get("caption", "") or ""
                    indent = "  " * el.get("list_level", 0)
                    if kind.startswith("list_"):
                        ordered = "нумерованный" if el.get("ordered") else "маркированный"
                        level = el.get("list_level", 0)
                        st.write(f"{indent}**{kind} ({ordered}, уровень {level})**: {preview[:120]}")
                    elif kind == "figure":
                        img_status = "с изображением" if el.get("image") else "без изображения"
                        st.write(f"{indent}**{kind} (№{el.get('number')}, {img_status})**: Подпись='{preview[:120]}'")
                    else:
                        st.write(f"{indent}**{kind}**: {preview[:120]}")

            # Кнопка скачивания
            with open("formatted_output.docx", "rb") as f:
                st.download_button(
                    label="Скачать отформатированный документ",
                    data=f,
                    file_name="formatted_output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download-button",
                    use_container_width=True
                )

            # Удаляем временные файлы
            temp_path.unlink()
            Path("formatted_output.docx").unlink()
            st.session_state.formatted = True

else:
    st.info("Пожалуйста, загрузите DOCX-файл, настройте параметры и нажмите кнопку 'Форматировать'.")