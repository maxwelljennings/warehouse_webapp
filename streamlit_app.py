import streamlit as st
import pandas as pd
import json
import os
import io
from collections import defaultdict, Counter
from typing import List, Dict, Optional, Set, Any
import traceback

# --- Logika analizy (bez zmian w funkcjach odczytu i analizy) ---

def read_layout_data() -> Optional[Dict]:
    """Czyta plik layout.json z repozytorium."""
    file_path = "layout.json"
    try:
        if not os.path.exists(file_path):
             st.warning(f"Plik '{file_path}' nie znaleziony. Używana jest pusta struktura.")
             return {"Version": "1.0", "Layout": {}}
        with open(file_path, 'r', encoding='utf-8') as f: data = json.load(f)
        if not isinstance(data, dict) or "Layout" not in data or not isinstance(data.get("Layout"), dict):
             raise ValueError("Plik layout.json nie zawiera oczekiwanej struktury JSON.")
        return data
    except Exception as e: st.error(f"Błąd odczytu pliku layout.json: {e}"); return None

def read_empty_locations_from_b3(uploaded_file) -> Optional[Set[str]]:
    if uploaded_file is None: st.error("Plik pustych lokalizacji nie został załadowany."); return None
    file_name = uploaded_file.name; st.info(f"Odczyt pliku: {file_name} (dane od B3)")
    try:
        excel_engine = None; is_csv = False
        file_ext = os.path.splitext(file_name)[1].lower()
        if file_ext == '.csv': is_csv = True
        elif file_ext == '.xls': excel_engine = 'xlrd'; import xlrd
        elif file_ext == '.xlsx': excel_engine = 'openpyxl'
        else: st.error(f"Nieobsługiwany format: {file_ext}."); return None
        df: Optional[pd.DataFrame] = None; file_content = io.BytesIO(uploaded_file.getvalue())
        if is_csv:
            detected_encoding = None
            for enc in ['utf-8', 'windows-1251']:
                try: file_content.seek(0); df = pd.read_csv(file_content, delimiter=';', header=2, usecols=[1], encoding=enc, skipinitialspace=True, on_bad_lines='skip'); detected_encoding = enc; break
                except: continue
            if df is None: raise ValueError("Nie udało się odczytać CSV.")
        else: df = pd.read_excel(file_content, header=None, skiprows=2, usecols=[1], sheet_name=0, engine=excel_engine)
        if df is None or df.empty: st.warning("Plik nie zawiera danych w kolumnie B od 3. wiersza."); return set()
        col_name = df.columns[0]; locations = df[col_name].dropna().astype(str).str.strip()
        empty_locations_set = set(locations[locations != ''])
        st.success(f"Załadowano {len(empty_locations_set)} pustych lokalizacji z {file_name}.")
        return empty_locations_set
    except Exception as e: st.error(f"Błąd odczytu pliku pustych lokalizacji: {e}"); return None

def analyze_opportunities_internal(layout_data: Dict, empty_locations_set: Set[str]) -> List[Dict[str, Any]]:
    opportunities = []
    sections_data = defaultdict(lambda: {"capacity": 0, "locations": []})
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
                    if loc_num is None: continue
                    location_id = f"{hall}.{row}{loc_num}.{level_str}"
                    is_accessible = loc_info.get("IsAccessible", True)
                    is_corridor = loc_info.get("IsCorridor", False)
                    section_id = loc_info.get("SectionID")
                    section_capacity = loc_info.get("SectionCapacity", 0)
                    pos_in_section = loc_info.get("PositionInSection", 0)
                    is_empty = location_id in empty_locations_set
                    # Добавляем hall в детали локации для группировки
                    location_details = {"id": location_id, "hall": hall, "number": loc_num, "is_accessible": is_accessible, "is_corridor": is_corridor, "is_empty": is_empty, "position": pos_in_section, "section_id": section_id, "section_capacity": section_capacity}
                    if section_id and is_accessible and not is_corridor: # Анализируем только доступные не-коридорные секции
                        if not sections_data[section_id]["capacity"]: sections_data[section_id]["capacity"] = section_capacity
                        sections_data[section_id]["locations"].append(location_details)

    for section_id, section_info in sections_data.items():
        locations_in_section = sorted(section_info["locations"], key=lambda x: x["position"])
        capacity = section_info["capacity"]
        loc_by_pos = {loc["position"]: loc for loc in locations_in_section}
        hall = locations_in_section[0]['hall'] if locations_in_section else 'N/A' # Получаем зал из первой локации

        if capacity == 2:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2)
            if (loc1 and loc1["is_empty"] and loc2 and loc2["is_empty"]):
                opportunities.append({"OpportunityType": "NonStandard_Direct", "Hall": hall, "PrimaryLocation": loc1["id"], "SecondaryLocation": loc2["id"], "SectionID": section_id, "Notes": "", "SectionLocations": [loc1, loc2]})
        elif capacity == 3:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2); loc3 = loc_by_pos.get(3)
            # Проверяем возможность 1+2
            if (loc1 and loc1["is_empty"] and loc2 and loc2["is_empty"]):
                 opportunities.append({"OpportunityType": "NonStandard_Direct", "Hall": hall, "PrimaryLocation": loc1["id"], "SecondaryLocation": loc2["id"], "SectionID": section_id, "Notes": "", "SectionLocations": [loc1, loc2, loc3]}) # Передаем все 3 локации для визуал.
            # Проверяем возможность 2+3
            if (loc2 and loc2["is_empty"] and loc3 and loc3["is_empty"]):
                 opportunities.append({"OpportunityType": "NonStandard_Direct", "Hall": hall, "PrimaryLocation": loc2["id"], "SecondaryLocation": loc3["id"], "SectionID": section_id, "Notes": "", "SectionLocations": [loc1, loc2, loc3]}) # Передаем все 3 локации
            # Проверяем возможность 1+3 с перемещением 2
            if (loc1 and loc1["is_empty"] and loc3 and loc3["is_empty"] and loc2 and not loc2["is_empty"]):
                 opportunities.append({"OpportunityType": "NonStandard_MoveRequired", "Hall": hall, "PrimaryLocation": loc1["id"], "SecondaryLocation": loc3["id"], "SectionID": section_id, "Notes": f"Przesuń paletę z {loc2['id']}", "SectionLocations": [loc1, loc2, loc3]}) # Передаем все 3 локации

    return opportunities


# НОВАЯ Функция форматирования с визуализацией
def format_visual_results(opportunities: List[Dict[str, Any]]) -> str:
    """Formatuje wyniki analizy wizualnie dla palet niestandardowych."""

    if not opportunities:
        return "<h3>Nie znaleziono obecnie miejsc dla palet niestandardowych.</h3>"

    # Сортируем по залу, затем по первой локации
    opportunities.sort(key=lambda x: (x.get("Hall", ""), x.get("PrimaryLocation", "")))

    html_parts = []
    current_hall = None

    # CSS Стили для ячеек
    styles = """
    <style>
        .hall-header {
            font-size: 1.3em;
            font-weight: bold;
            margin-top: 15px;
            margin-bottom: 10px;
            border-bottom: 1px solid #ccc;
            padding-bottom: 5px;
        }
        .opportunity-block {
            margin-bottom: 15px;
            padding: 10px;
            border: 1px solid #eee;
            border-radius: 5px;
            background-color: #f9f9f9;
        }
        .section-vis {
            display: flex;
            margin-bottom: 5px;
            border: 2px solid #666; /* Имитация балок */
            border-radius: 3px;
            width: fit-content; /* Ширина по содержимому */
        }
        .location-cell {
            padding: 5px 8px; /* Уменьшены паддинги */
            border-right: 1px dotted #ccc;
            text-align: center;
            font-size: 0.9em; /* Уменьшен шрифт */
            min-width: 70px; /* Минимальная ширина ячейки */
            white-space: nowrap; /* Предотвращаем перенос текста */
        }
        .location-cell:last-child {
            border-right: none;
        }
        .empty { background-color: #c8e6c9; } /* Светло-зеленый */
        .occupied { background-color: #ffcdd2; } /* Светло-красный */
        .to-move { background-color: #ffecb3; font-weight: bold; } /* Желтоватый для перемещаемой */
        .action-text {
            font-weight: bold;
            margin-top: 5px;
        }
        .notes-text {
            font-style: italic;
            color: #d32f2f; /* Красный для примечания */
            font-size: 0.9em;
        }
    </style>
    """
    html_parts.append(styles)

    for opp in opportunities:
        hall = opp.get("Hall", "Nieznana Hala")
        opp_type = opp.get("OpportunityType")
        loc1_id = opp.get("PrimaryLocation", "")
        loc2_id = opp.get("SecondaryLocation", "")
        notes = opp.get("Notes", "")
        section_locations = opp.get("SectionLocations", []) # Список словарей локаций секции

        # Выводим заголовок зала при смене
        if hall != current_hall:
            html_parts.append(f"<div class='hall-header'>Hala {hall}</div>")
            current_hall = hall

        html_parts.append("<div class='opportunity-block'>")

        # --- Визуализация секции ---
        html_parts.append("<div class='section-vis'>")
        location_to_move = ""
        if "Przesuń paletę z " in notes:
            location_to_move = notes.split("Przesuń paletę z ")[1]

        for loc_detail in section_locations:
            loc_id = loc_detail.get("id", "")
            is_empty = loc_detail.get("is_empty", False)
            css_class = "location-cell"
            if is_empty:
                css_class += " empty"
            else:
                # Если это локация, которую нужно переместить
                if loc_id == location_to_move:
                    css_class += " to-move"
                else:
                    css_class += " occupied"

            html_parts.append(f"<div class='{css_class}'>{loc_id}</div>")
        html_parts.append("</div>") # Конец section-vis

        # --- Текст действия ---
        if opp_type == "NonStandard_Direct":
            html_parts.append(f"<div class='action-text'>Można postawić (2 miejsca): {loc1_id} i {loc2_id}</div>")
        elif opp_type == "NonStandard_MoveRequired":
            html_parts.append(f"<div class='action-text'>Można postawić (2 miejsca): {loc1_id} i {loc2_id}</div>")
            if notes:
                 html_parts.append(f"<div class='notes-text'>{notes}</div>")

        html_parts.append("</div>") # Конец opportunity-block

    return "".join(html_parts)


# --- Streamlit UI ---
st.set_page_config(page_title="Wyszukiwarka Miejsc Magazynowych", layout="wide")
st.title("Wyszukiwarka Wolnych Miejsc dla Palet Niestandardowych")

st.sidebar.header("1. Załaduj plik")
uploaded_empty_loc_file = st.sidebar.file_uploader(
    "Załaduj plik pustych lokalizacji (Excel/CSV)",
    type=["xlsx", "xls", "csv"]
)
st.sidebar.info("Plik struktury magazynu (layout.json) jest ładowany automatycznie.")

st.sidebar.header("2. Uruchom analizę")
run_button = st.sidebar.button("Uruchom analizę", disabled=(not uploaded_empty_loc_file))

st.header("Wyniki analizy")
results_placeholder = st.empty() # Используем empty для возможности замены содержимого
results_placeholder.info("Załaduj plik pustych lokalizacji i kliknij 'Uruchom analizę'.")

# --- Логика выполнения ---
analysis_results_html = "" # Переменная для хранения HTML результатов

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
                opportunities = analyze_opportunities_internal(layout_data, empty_locations)
                # ИЗМЕНЕНО: Генерируем HTML
                analysis_results_html = format_visual_results(opportunities)
                # Отображаем HTML
                results_placeholder.markdown(analysis_results_html, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                analysis_results_html = "" # Очищаем при ошибке

        else:
            results_placeholder.warning("Analiza nie może zostać wykonana (błąd odczytu pliku pustych lokalizacji).")
            analysis_results_html = "" # Очищаем при ошибке

# --- Кнопка Печати ---
# Показываем кнопку только если есть результаты для печати (HTML не пуст и не содержит сообщение об отсутствии мест)
if analysis_results_html and "Nie znaleziono obecnie miejsc" not in analysis_results_html:
    st.sidebar.header("3. Drukuj")
    print_button_html = """
    <style>
    .print-button { /* Стили кнопки */ }
    /* ... (скопируйте стили для .print-button из предыдущего ответа) ... */
     .print-button {
        display: inline-block; padding: 0.5em 1em; border: 1px solid #ccc; border-radius: 4px;
        background-color: #f0f0f0; color: #333; text-align: center; text-decoration: none;
        cursor: pointer; font-size: 1em; font-family: inherit; width: 90%; /* Ширина кнопки */ margin-top: 10px;
    }
    .print-button:hover { background-color: #e0e0e0; }
    .print-button:active { background-color: #d0d0d0; }
    </style>
    <button class="print-button" onclick="window.print()">Drukuj wyniki</button>
    """
    st.sidebar.markdown(print_button_html, unsafe_allow_html=True)
