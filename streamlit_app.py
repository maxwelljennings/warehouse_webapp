import streamlit as st
import pandas as pd
import json
import os
import io
from collections import defaultdict, Counter, OrderedDict
from typing import List, Dict, Optional, Set, Any
import traceback

# --- Логика анализа ---

# Функция чтения локального layout.json (без изменений)
def read_layout_data() -> Optional[Dict]:
    file_path = "layout.json"
    try:
        if not os.path.exists(file_path):
             st.warning(f"Plik '{file_path}' nie znaleziony. Używana jest pusta struktura.")
             return {"Version": "1.0", "Layout": {}}
        with open(file_path, 'r', encoding='utf-8') as f: data = json.load(f)
        if not isinstance(data, dict) or "Layout" not in data or not isinstance(data.get("Layout"), dict): raise ValueError("Plik layout.json nie zawiera oczekiwanej struktury JSON.")
        return data
    except Exception as e: st.error(f"Błąd odczytu pliku layout.json: {e}"); return None

# Функция чтения пустых локаций (без изменений)
def read_empty_locations_from_b3(uploaded_file) -> Optional[Set[str]]:
    if uploaded_file is None: st.error("Plik pustych lokalizacji nie został załadowany."); return None
    file_name = uploaded_file.name; st.info(f"Odczyt pliku pustych lokalizacji: {file_name} (dane od B3)")
    try:
        excel_engine = None; is_csv = False; file_ext = os.path.splitext(file_name)[1].lower()
        if file_ext == '.csv': is_csv = True
        elif file_ext == '.xls': excel_engine = 'xlrd'; import xlrd
        elif file_ext == '.xlsx': excel_engine = 'openpyxl'
        else: st.error(f"Nieobsługiwany format pliku: {file_ext}."); return None
        df: Optional[pd.DataFrame] = None; file_content = io.BytesIO(uploaded_file.getvalue())
        if is_csv:
            detected_encoding = None
            for enc in ['utf-8', 'windows-1251']:
                try: file_content.seek(0); df = pd.read_csv(file_content, delimiter=';', header=2, usecols=[1], encoding=enc, skipinitialspace=True, on_bad_lines='skip'); st.write(f"(CSV odczytany z kodowaniem {enc})"); detected_encoding = enc; break
                except: continue
            if df is None: raise ValueError("Nie udało się odczytać pliku CSV.")
        else: df = pd.read_excel(file_content, header=None, skiprows=2, usecols=[1], sheet_name=0, engine=excel_engine)
        if df is None or df.empty: st.warning("Plik pustych lokalizacji nie zawiera danych w kolumnie B od 3. wiersza."); return set()
        col_name = df.columns[0]; locations = df[col_name].dropna().astype(str).str.strip(); empty_locations_set = set(locations[locations != ''])
        count = len(empty_locations_set); st.success(f"Załadowano {count} unikalnych ID pustych lokalizacji z {file_name}.")
        return empty_locations_set
    except ImportError: st.error("Brak biblioteki 'xlrd' do odczytu .xls."); return None
    except Exception as e: st.error(f"Błąd odczytu pliku pustych lokalizacji: {e}"); return None

# ИЗМЕНЕНО: Функция анализа теперь возвращает сгруппированные данные по секциям
def analyze_sections(layout_data: Dict, empty_locations_set: Set[str]) -> Dict[str, Dict]:
    """Analizuje strukturę i puste miejsca, grupując dane według sekcji."""
    sections_data = defaultdict(lambda: {
        "capacity": 0,
        "locations": {}, # Словарь: {position: location_details}
        "hall": "",
        "row": ""
    })
    layout = layout_data.get("Layout", {})

    for hall, rows in layout.items():
        if not isinstance(rows, dict): continue
        for row, levels in rows.items():
            if not isinstance(levels, dict): continue
            for level_str, level_data in levels.items():
                if not isinstance(level_data, dict): continue
                locations = level_data.get("Locations", [])
                if not isinstance(locations, list): continue
                for loc_info in locations:
                    if not isinstance(loc_info, dict): continue
                    loc_num = loc_info.get("Number")
                    section_id = loc_info.get("SectionID")
                    if loc_num is None or not section_id: continue # Обрабатываем только локации в секциях

                    location_id = f"{hall}.{row}{loc_num}.{level_str}"
                    is_accessible = loc_info.get("IsAccessible", True)
                    is_corridor = loc_info.get("IsCorridor", False) # Коридоры не должны быть в секциях, но проверим
                    section_capacity = loc_info.get("SectionCapacity", 0)
                    pos_in_section = loc_info.get("PositionInSection", 0)
                    is_empty = location_id in empty_locations_set

                    if not is_accessible or is_corridor: continue # Пропускаем недоступные/коридоры

                    location_details = {
                        "id": location_id,
                        "is_empty": is_empty,
                    }

                    # Сохраняем детали локации по ее позиции в секции
                    sections_data[section_id]["locations"][pos_in_section] = location_details
                    # Сохраняем общую информацию о секции (один раз)
                    if not sections_data[section_id]["capacity"]:
                        sections_data[section_id]["capacity"] = section_capacity
                        sections_data[section_id]["hall"] = hall
                        sections_data[section_id]["row"] = row # Сохраняем ряд для сортировки

    # Убираем секции, которые оказались невалидными (например, не все позиции заполнены)
    valid_sections = {}
    for sec_id, data in sections_data.items():
        if len(data["locations"]) == data["capacity"] and data["capacity"] in [2, 3]:
             valid_sections[sec_id] = data

    st.write(f"(Analiza: Znaleziono {len(valid_sections)} poprawnych sekcji 2- i 3-miejscowych)")
    return valid_sections


# НОВАЯ Функция форматирования с визуализацией
def format_analysis_results_visual(analyzed_sections: Dict[str, Dict]) -> str:
    """Formatuje wyniki analizy wizualnie, grupując według hali."""
    if not analyzed_sections:
        return "Nie znaleziono sekcji do analizy pod kątem palet niestandardowych."

    results_by_hall = defaultdict(list) # Словарь: {hall: [list_of_visual_strings]}

    # Константы для визуализации
    empty_slot = "[ PUSTE ]"
    busy_slot = "[ZAJĘTE]"
    loc_width = 15 # Ширина для ID локации

    # Обрабатываем каждую секцию
    for section_id, data in analyzed_sections.items():
        capacity = data["capacity"]
        locations = data["locations"] # Словарь {pos: details}
        hall = data["hall"]
        row_num_part = data["row"] # Используем для сортировки внутри зала

        loc1 = locations.get(1)
        loc2 = locations.get(2)
        loc3 = locations.get(3) if capacity == 3 else None

        loc1_id = loc1['id'] if loc1 else ""
        loc2_id = loc2['id'] if loc2 else ""
        loc3_id = loc3['id'] if loc3 else ""

        loc1_state = empty_slot if loc1 and loc1['is_empty'] else busy_slot
        loc2_state = empty_slot if loc2 and loc2['is_empty'] else busy_slot
        loc3_state = empty_slot if loc3 and loc3['is_empty'] else busy_slot

        visual_lines = [] # Строки для текущей секции

        # --- Случай секции на 3 места ---
        if capacity == 3 and loc1 and loc2 and loc3:
            # 1. Все 3 пустые -> 2 возможности
            if loc1['is_empty'] and loc2['is_empty'] and loc3['is_empty']:
                line1 = f"`{loc1_id:<{loc_width}}` `{loc2_id:<{loc_width}}` `{loc3_id:<{loc_width}}`"
                line2 = f"{empty_slot} | {empty_slot} | {empty_slot}"
                line3 = f"-> **Można postawić DWIE palety niestandardowe** (na {loc1_id} i {loc2_id} ORAZ na {loc2_id} i {loc3_id})"
                visual_lines.extend([line1, line2, line3])

            # 2. Пустые 1 и 2 -> 1 возможность
            elif loc1['is_empty'] and loc2['is_empty'] and not loc3['is_empty']:
                line1 = f"`{loc1_id:<{loc_width}}` `{loc2_id:<{loc_width}}` `{loc3_id:<{loc_width}}`"
                line2 = f"{empty_slot} | {empty_slot} | {busy_slot}"
                line3 = f"-> **Można postawić paletę niestandardową** (na {loc1_id} i {loc2_id})"
                visual_lines.extend([line1, line2, line3])

            # 3. Пустые 2 и 3 -> 1 возможность
            elif not loc1['is_empty'] and loc2['is_empty'] and loc3['is_empty']:
                line1 = f"`{loc1_id:<{loc_width}}` `{loc2_id:<{loc_width}}` `{loc3_id:<{loc_width}}`"
                line2 = f"{busy_slot} | {empty_slot} | {empty_slot}"
                line3 = f"-> **Można postawić paletę niestandardową** (na {loc2_id} i {loc3_id})"
                visual_lines.extend([line1, line2, line3])

            # 4. Пустые 1 и 3, занята 2 -> Требуется перемещение
            elif loc1['is_empty'] and not loc2['is_empty'] and loc3['is_empty']:
                line1 = f"`{loc1_id:<{loc_width}}` `{loc2_id:<{loc_width}}` `{loc3_id:<{loc_width}}`"
                line2 = f"{empty_slot} | {busy_slot} | {empty_slot}"
                line3 = f"-> **Przesuń paletę z `{loc2_id}`** (na `{loc1_id}` LUB `{loc3_id}`)"
                line4 = f"   Wtedy można postawić paletę niestandardową."
                visual_lines.extend([line1, line2, line3, line4])

        # --- Случай секции на 2 места ---
        elif capacity == 2 and loc1 and loc2:
            # 1. Обе пустые -> 1 возможность
            if loc1['is_empty'] and loc2['is_empty']:
                line1 = f"`{loc1_id:<{loc_width}}` `{loc2_id:<{loc_width}}`"
                line2 = f"{empty_slot} | {empty_slot}"
                line3 = f"-> **Można postawić paletę niestandardową** (na {loc1_id} i {loc2_id})"
                visual_lines.extend([line1, line2, line3])

        # Добавляем собранные строки для этой секции в соответствующий зал
        if visual_lines:
            # Добавляем пустую строку перед каждой новой секцией для разделения
            results_by_hall[hall].append(("", visual_lines)) # Кортеж для сортировки

    # --- Формируем итоговый текст ---
    if not results_by_hall:
        return "Nie znaleziono możliwości umieszczenia palet niestandardowych."

    output_lines = []
    # Сортируем залы по имени
    for hall in sorted(results_by_hall.keys()):
        output_lines.append(f"\n--- HALA: {hall} ---")
        # Сортируем секции внутри зала по номеру ряда (извлекаем число из ID первой локации)
        # Это простая сортировка, может быть не идеальной для сложных имен рядов
        def sort_key(item):
            visuals = item[1]
            first_line = visuals[0]
            try:
                # Извлекаем ID первой локации из первой строки `F.A11.0` -> 11
                loc_id = first_line.split('`')[1]
                row_part = loc_id.split('.')[1] # A11
                num_part = ''.join(filter(str.isdigit, row_part))
                return int(num_part) if num_part else 0
            except:
                return 0 # Если не удалось извлечь номер

        sorted_sections = sorted(results_by_hall[hall], key=sort_key)

        for _, visual_lines_for_section in sorted_sections:
             output_lines.append("") # Пустая строка-разделитель
             output_lines.extend(visual_lines_for_section)


    return "\n".join(output_lines)


# --- Streamlit UI ---
st.set_page_config(page_title="Wyszukiwarka Miejsc Magazynowych", layout="wide")
st.title("Wyszukiwarka Miejsc dla Palet Niestandardowych")

st.sidebar.header("1. Załaduj plik")

uploaded_empty_loc_file = st.sidebar.file_uploader(
    "Załaduj plik pustych lokalizacji (Excel/CSV)",
    type=["xlsx", "xls", "csv"]
)

st.sidebar.info("Plik struktury magazynu (layout.json) jest ładowany automatycznie.")

st.sidebar.header("2. Uruchom analizę")
run_button = st.sidebar.button("Uruchom analizę", disabled=(not uploaded_empty_loc_file))

st.header("Wyniki analizy")
results_placeholder = st.empty()
results_placeholder.info("Załaduj plik pustych lokalizacji i kliknij 'Uruchom analizę'.")

# --- Логика выполнения при нажатии кнопки ---
analysis_results_text = "" # Переменная для хранения текста для печати

if run_button:
    results_placeholder.info("Wczytywanie struktury magazynu...")
    layout_data = read_layout_data()

    if layout_data is None:
         results_placeholder.error("Nie udało się wczytać pliku struktury magazynu (layout.json).")
    else:
        results_placeholder.info("Przetwarzanie pliku pustych lokalizacji...")
        empty_locations = read_empty_locations_from_b3(uploaded_empty_loc_file)

        if empty_locations is not None:
            results_placeholder.info("Wykonywanie analizy...")
            try:
                # ИЗМЕНЕНО: Вызываем новую функцию анализа
                analyzed_sections = analyze_sections(layout_data, empty_locations)
                # ИЗМЕНЕНО: Вызываем новую функцию форматирования
                analysis_results_text = format_analysis_results_visual(analyzed_sections)
                # Используем markdown для отображения, т.к. используем ` `` ` для форматирования
                results_placeholder.markdown(analysis_results_text, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                analysis_results_text = ""

        else:
            results_placeholder.warning("Analiza nie może zostać wykonana (błąd odczytu pliku pustych lokalizacji).")
            analysis_results_text = ""

# --- Кнопка Печати ---
if analysis_results_text and "Nie znaleziono" not in analysis_results_text:
    st.sidebar.header("3. Drukuj")
    print_button_html = """
    <style>
    /* Стили для кнопки печати */
    .print-button { display: inline-block; padding: 0.5em 1em; border: 1px solid #ccc; border-radius: 4px; background-color: #f0f0f0; color: #333; text-align: center; text-decoration: none; cursor: pointer; font-size: 1em; font-family: inherit; }
    .print-button:hover { background-color: #e0e0e0; }
    .print-button:active { background-color: #d0d0d0; }
    /* Стили для печати: убираем боковую панель и заголовки Streamlit */
    @media print {
        header[data-testid="stHeader"], div[data-testid="stSidebarNav"] { display: none; }
        section[data-testid="stSidebar"] { display: none; }
        div[data-testid="stToolbar"] { display: none; }
        div.stButton button { display: none; } /* Скрываем кнопки Streamlit при печати */
        button.print-button { display: none; } /* Скрываем саму кнопку печати */
        /* Уменьшаем отступы основного контента для печати */
        div[data-testid="stAppViewContainer"] > section { padding: 1cm !important; }
        /* Можно добавить стили для текста, если нужно */
        pre, code { white-space: pre-wrap !important; word-wrap: break-word !important; }
    }
    </style>
    <button class="print-button" onclick="window.print()">Drukuj wyniki</button>
    """
    st.sidebar.markdown(print_button_html, unsafe_allow_html=True)
