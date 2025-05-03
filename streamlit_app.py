import streamlit as st
import pandas as pd
import json
import os
import io
from collections import defaultdict, Counter
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
    file_name = uploaded_file.name
    st.info(f"Odczyt pliku pustych lokalizacji: {file_name} (dane od B3)")
    try:
        excel_engine = None; is_csv = False
        file_ext = os.path.splitext(file_name)[1].lower()
        if file_ext == '.csv': is_csv = True
        elif file_ext == '.xls': excel_engine = 'xlrd'; import xlrd
        elif file_ext == '.xlsx': excel_engine = 'openpyxl'
        else: st.error(f"Nieobsługiwany format pliku: {file_ext}."); return None
        df: Optional[pd.DataFrame] = None
        file_content = io.BytesIO(uploaded_file.getvalue())
        if is_csv:
            detected_encoding = None
            for enc in ['utf-8', 'windows-1251']:
                try: file_content.seek(0); df = pd.read_csv(file_content, delimiter=';', header=2, usecols=[1], encoding=enc, skipinitialspace=True, on_bad_lines='skip'); st.write(f"(CSV odczytany z kodowaniem {enc})"); detected_encoding = enc; break
                except: continue
            if df is None: raise ValueError("Nie udało się odczytać pliku CSV.")
        else: df = pd.read_excel(file_content, header=None, skiprows=2, usecols=[1], sheet_name=0, engine=excel_engine)
        if df is None or df.empty: st.warning("Plik pustych lokalizacji nie zawiera danych w kolumnie B od 3. wiersza."); return set()
        col_name = df.columns[0]
        locations = df[col_name].dropna().astype(str).str.strip()
        empty_locations_set = set(locations[locations != ''])
        count = len(empty_locations_set)
        st.success(f"Załadowano {count} unikalnych ID pustych lokalizacji z {file_name}.")
        return empty_locations_set
    except ImportError as e: st.error(f"Błąd importu biblioteki: {e}."); return None
    except ValueError as e: st.error(f"Błąd danych lub formatu pliku pustych lokalizacji: {e}"); return None
    except Exception as e: st.error(f"Nieoczekiwany błąd podczas odczytu pliku pustych lokalizacji: {e}"); return None

# ИЗМЕНЕНО: Функция анализа теперь возвращает словарь секций с возможностями
def analyze_opportunities_internal(layout_data: Dict, empty_locations_set: Set[str]) -> Dict[str, Dict]:
    """
    Analizuje możliwości i zwraca słownik sekcji z możliwościami
    dla palet niestandardowych.
    Klucz: SectionID
    Wartość: Słownik {'capacity': int, 'locations': list[dict], 'opportunity_type': str}
    opportunity_type: 'Direct' lub 'MoveRequired'
    """
    relevant_sections = defaultdict(lambda: {"capacity": 0, "locations": [], "opportunity_type": None})
    layout = layout_data.get("Layout", {})

    # 1. Zbierz dane o wszystkich lokalizacjach we wszystkich sekcjach
    all_sections_data = defaultdict(lambda: {"capacity": 0, "locations": []})
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
                    section_id = loc_info.get("SectionID")
                    if not section_id: continue # Interesują nas tylko sekcje

                    loc_num = loc_info.get("Number")
                    if loc_num is None: continue

                    location_id = f"{hall}.{row}{loc_num}.{level_str}"
                    is_accessible = loc_info.get("IsAccessible", True)
                    is_corridor = loc_info.get("IsCorridor", False) # Коридоры не должны быть в секциях, но проверим
                    section_capacity = loc_info.get("SectionCapacity", 0)
                    pos_in_section = loc_info.get("PositionInSection", 0)
                    is_empty = location_id in empty_locations_set

                    location_details = {
                        "id": location_id,
                        "number": loc_num,
                        "is_accessible": is_accessible,
                        "is_corridor": is_corridor,
                        "is_empty": is_empty,
                        "position": pos_in_section,
                        # Добавляем capacity в детали локации для удобства
                        "section_capacity": section_capacity
                    }

                    if not all_sections_data[section_id]["capacity"]:
                        all_sections_data[section_id]["capacity"] = section_capacity
                    all_sections_data[section_id]["locations"].append(location_details)

    # 2. Przeanalizuj zebrane sekcje pod kątem możliwości niestandardowych
    for section_id, section_info in all_sections_data.items():
        locations_in_section = sorted(section_info["locations"], key=lambda x: x["position"])
        capacity = section_info["capacity"]
        loc_by_pos = {loc["position"]: loc for loc in locations_in_section}

        found_opportunity = None # 'Direct' lub 'MoveRequired'

        # Sprawdź NonStandard_MoveRequired (Capacity 3) - ma priorytet
        if capacity == 3:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2); loc3 = loc_by_pos.get(3)
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"] and
                loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and not loc2["is_empty"]):
                found_opportunity = "MoveRequired"
                # Добавляем информацию о том, какую локацию двигать, прямо в детали средней локации
                loc2['needs_move'] = True
                loc2['move_target_for'] = (loc1['id'], loc3['id']) # Куда можно поставить после сдвига

        # Sprawdź NonStandard_Direct (jeśli MoveRequired не найден)
        if not found_opportunity:
            if capacity == 2:
                loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2)
                if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                    loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"]):
                    found_opportunity = "Direct"
            elif capacity == 3:
                loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2); loc3 = loc_by_pos.get(3)
                # Проверяем обе пары
                pair12_ok = (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                             loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"])
                pair23_ok = (loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"] and
                             loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"])
                if pair12_ok or pair23_ok:
                    found_opportunity = "Direct"

        # Если найдена любая нестандартная возможность, добавляем секцию в результат
        if found_opportunity:
            relevant_sections[section_id]["capacity"] = capacity
            relevant_sections[section_id]["locations"] = locations_in_section # Сохраняем все локации секции
            relevant_sections[section_id]["opportunity_type"] = found_opportunity

    st.write(f"(Analiza: Znaleziono {len(relevant_sections)} sekcji z możliwościami dla palet niestandardowych)")
    return relevant_sections


# НОВАЯ ФУНКЦИЯ: Форматирование секции в HTML для визуализации
def format_section_visually(section_id: str, section_data: Dict) -> str:
    """Generuje HTML do wizualizacji pojedynczej sekcji."""

    locations = section_data.get("locations", [])
    capacity = section_data.get("capacity", 0)
    opportunity_type = section_data.get("opportunity_type", "Unknown")

    if not locations:
        return ""

    # Определяем заголовок секции
    header_text = ""
    if opportunity_type == "Direct":
        header_text = f"Sekcja {section_id}: Można postawić OD RAZU"
    elif opportunity_type == "MoveRequired":
        header_text = f"Sekcja {section_id}: Można postawić PO PRZESUNIĘCIU"
    else:
        header_text = f"Sekcja {section_id}" # На всякий случай

    # Начинаем HTML блок для секции
    html = f"<div class='section-container'><h4>{header_text}</h4><div class='location-row'>"

    # Генерируем HTML для каждой локации
    for i, loc in enumerate(locations):
        loc_id = loc.get("id", "N/A")
        loc_num = loc.get("number", "?")
        is_empty = loc.get("is_empty", False)
        is_accessible = loc.get("is_accessible", True)
        needs_move = loc.get("needs_move", False) # Проверяем флаг для перемещения
        pos = loc.get("position", 0)

        # Определяем класс CSS для стиля
        cell_class = "location-cell"
        if not is_accessible:
            cell_class += " inaccessible"
        elif needs_move:
             cell_class += " needs-move" # Желтый для перемещаемой
        elif is_empty:
            cell_class += " empty" # Зеленый для пустых
        else:
            cell_class += " occupied" # Красный/серый для занятых

        # Определяем стиль рамки (балки)
        border_style = ""
        if pos == 1:
            border_style += " border-left: 3px solid #333;" # Левая балка
        if pos == capacity:
            border_style += " border-right: 3px solid #333;" # Правая балка
        if pos > 1:
             border_style += " border-left: 1px dotted #aaa;" # Внутренняя левая
        # Добавляем общие рамки сверху/снизу и справа для внутренних
        border_style += " border-top: 1px solid #ccc; border-bottom: 1px solid #ccc;"
        if pos < capacity:
             border_style += " border-right: 1px dotted #aaa;" # Внутренняя правая


        # Текст внутри ячейки
        cell_text = str(loc_num)
        if needs_move:
            cell_text += " 🚚" # Иконка грузовика для перемещаемой

        # Собираем HTML для ячейки
        html += f"<div class='{cell_class}' style='{border_style}' title='{loc_id}'>{cell_text}</div>"

    html += "</div></div>" # Закрываем location-row и section-container
    return html

# --- CSS Стили для визуализации ---
# Определяем стили один раз
CSS_STYLES = """
<style>
.section-container {
    border: 1px solid #eee;
    border-radius: 5px;
    padding: 10px;
    margin-bottom: 15px;
    background-color: #f9f9f9;
}
.section-container h4 {
    margin-top: 0;
    margin-bottom: 10px;
    color: #333;
}
.location-row {
    display: flex; /* Располагаем ячейки в ряд */
    flex-wrap: nowrap; /* Запрещаем перенос строки */
    width: fit-content; /* Ширина по содержимому */
    margin-left: auto; /* Центрируем ряд, если нужно */
    margin-right: auto;
}
.location-cell {
    min-width: 60px; /* Минимальная ширина ячейки */
    padding: 15px 5px; /* Внутренние отступы (верт./гориз.) */
    text-align: center;
    font-weight: bold;
    font-size: 0.9em;
    color: #333;
    box-sizing: border-box; /* Чтобы padding и border входили в ширину */
    /* Общие рамки будут добавлены инлайн */
}
.location-cell.empty {
    background-color: #c8e6c9; /* Светло-зеленый */
    color: #2e7d32;
}
.location-cell.occupied {
    background-color: #ffebee; /* Светло-красный */
    color: #c62828;
}
.location-cell.needs-move {
    background-color: #fff9c4; /* Светло-желтый */
    color: #f57f17;
    /* Можно добавить анимацию или другой индикатор */
    /* animation: blink 1s linear infinite; */
}
/* @keyframes blink { 50% { opacity: 0.6; } } */

.location-cell.inaccessible {
    background-color: #e0e0e0; /* Серый */
    color: #757575;
    text-decoration: line-through; /* Перечеркнутый текст */
}
.legend {
    margin-top: 20px;
    padding: 10px;
    border: 1px dashed #ccc;
    font-size: 0.9em;
}
.legend-item {
    display: inline-block;
    margin-right: 15px;
}
.legend-color {
    display: inline-block;
    width: 15px;
    height: 15px;
    margin-right: 5px;
    vertical-align: middle;
    border: 1px solid #999;
}
</style>
"""

# --- Streamlit UI ---
st.set_page_config(page_title="Wyszukiwarka Miejsc Magazynowych", layout="wide")
st.title("Wyszukiwarka Wolnych Miejsc dla Palet Niestandardowych")

# Вставляем CSS стили в начало
st.markdown(CSS_STYLES, unsafe_allow_html=True)

st.sidebar.header("1. Załaduj plik")

uploaded_empty_loc_file = st.sidebar.file_uploader(
    "Załaduj plik pustych lokalizacji (Excel/CSV)",
    type=["xlsx", "xls", "csv"]
)

st.sidebar.info("Plik struktury magazynu (layout.json) jest ładowany automatycznie.")

st.sidebar.header("2. Uruchom analizę")
run_button = st.sidebar.button("Uruchom analizę", disabled=(not uploaded_empty_loc_file))

st.header("Wyniki analizy wizualnej")
results_placeholder = st.empty() # Используем для вывода HTML секций
results_placeholder.info("Załaduj plik pustych lokalizacji i kliknij 'Uruchom analizę'.")

# --- Логика выполнения при нажатии кнопки ---
analysis_performed = False # Флаг, что анализ был выполнен

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
                # ИЗМЕНЕНО: Получаем словарь секций
                relevant_sections_data = analyze_opportunities_internal(layout_data, empty_locations)

                # Очищаем плейсхолдер перед выводом новых результатов
                results_placeholder.empty()

                if not relevant_sections_data:
                    results_placeholder.warning("Nie znaleziono sekcji z możliwościami dla palet niestandardowych.")
                else:
                    analysis_performed = True # Анализ успешен и есть результаты
                    # Выводим легенду
                    st.markdown("""
                    <div class="legend">
                        <b>Legenda:</b>
                        <span class="legend-item"><span class="legend-color" style="background-color: #c8e6c9;"></span> Wolne</span>
                        <span class="legend-item"><span class="legend-color" style="background-color: #ffebee;"></span> Zajęte</span>
                        <span class="legend-item"><span class="legend-color" style="background-color: #fff9c4;"></span> 🚚 Do przesunięcia</span>
                        <span class="legend-item"><span class="legend-color" style="background-color: #e0e0e0; text-decoration: line-through;"></span> Niedostępne</span>
                    </div>
                    """, unsafe_allow_html=True)

                    # Итерируем и выводим каждую секцию визуально
                    # Сортируем секции по ID для консистентного вывода
                    sorted_section_ids = sorted(relevant_sections_data.keys())
                    for section_id in sorted_section_ids:
                        section_html = format_section_visually(section_id, relevant_sections_data[section_id])
                        st.markdown(section_html, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                traceback.print_exc() # Выводим traceback в лог Streamlit для отладки

        else:
            results_placeholder.warning("Analiza nie może zostać wykonana (błąd odczytu pliku pustych lokalizacji).")

# --- Кнопка Печати ---
# Показываем кнопку только если анализ был выполнен и были найдены секции
if analysis_performed:
    st.sidebar.header("3. Drukuj")
    print_button_html = """
    <style>
    .print-button { /* Стили кнопки */ }
    /* ... (скопируйте стили кнопки из предыдущего ответа) ... */
     .print-button {
        display: inline-block; padding: 0.5em 1em; border: 1px solid #ccc;
        border-radius: 4px; background-color: #f0f0f0; color: #333;
        text-align: center; text-decoration: none; cursor: pointer;
        font-size: 1em; font-family: inherit; width: 100%; /* Растянем на всю ширину сайдбара */
        margin-top: 10px;
    }
    .print-button:hover { background-color: #e0e0e0; }
    .print-button:active { background-color: #d0d0d0; }
    </style>
    <button class="print-button" onclick="window.print()">Drukuj wyniki</button>
    """
    st.sidebar.markdown(print_button_html, unsafe_allow_html=True)
