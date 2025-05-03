import streamlit as st
import pandas as pd
import json
import os
import io
from collections import defaultdict, Counter
from typing import List, Dict, Optional, Set, Any
import traceback

# --- Логика анализа ---

# Функция чтения локального layout.json (bez zmian)
def read_layout_data() -> Optional[Dict]:
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

# Функция чтения пустых локаций (bez zmian)
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
                try:
                    file_content.seek(0)
                    df = pd.read_csv(file_content, delimiter=';', header=2, usecols=[1], encoding=enc, skipinitialspace=True, on_bad_lines='skip')
                    st.write(f"(CSV odczytany z kodowaniem {enc})"); detected_encoding = enc; break
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
    except Exception as e: st.error(f"Błąd odczytu pliku pustych lokalizacji: {e}"); return None

# ИЗМЕНЕНО: Функция анализа теперь добавляет детали всех локаций секции
def analyze_opportunities_internal(layout_data: Dict, empty_locations_set: Set[str]) -> List[Dict[str, Any]]:
    """Wykonuje analizę możliwości, dodając szczegóły lokalizacji sekcji."""
    opportunities = []
    sections_data = defaultdict(lambda: {"capacity": 0, "locations": [], "hall": ""}) # Dodano hall
    layout = layout_data.get("Layout", {})

    # Krok 1: Zbierz dane o sekcjach
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
                    if section_id:
                        loc_num = loc_info.get("Number")
                        if loc_num is None: continue
                        location_id = f"{hall}.{row}{loc_num}.{level_str}"
                        is_empty = location_id in empty_locations_set
                        location_details = {
                            "id": location_id,
                            "position": loc_info.get("PositionInSection", 0),
                            "is_empty": is_empty,
                            "is_accessible": loc_info.get("IsAccessible", True),
                            "is_corridor": loc_info.get("IsCorridor", False)
                        }
                        if not sections_data[section_id]["capacity"]:
                            sections_data[section_id]["capacity"] = loc_info.get("SectionCapacity", 0)
                            sections_data[section_id]["hall"] = hall # Zapisz halę dla sekcji
                        sections_data[section_id]["locations"].append(location_details)

    # Krok 2: Analizuj każdą sekcję
    for section_id, section_info in sections_data.items():
        locations_in_section = sorted(section_info["locations"], key=lambda x: x["position"])
        capacity = section_info["capacity"]
        hall = section_info["hall"]
        loc_by_pos = {loc["position"]: loc for loc in locations_in_section}

        # Pomijamy sekcje, które nie są dostępne lub są korytarzami (chociaż takich nie powinno być)
        if not all(loc.get("is_accessible", False) and not loc.get("is_corridor", True) for loc in locations_in_section):
            continue

        # Sprawdź możliwości dla palet niestandardowych
        if capacity == 2:
            loc1 = loc_by_pos.get(1)
            loc2 = loc_by_pos.get(2)
            if loc1 and loc1["is_empty"] and loc2 and loc2["is_empty"]:
                opportunities.append({
                    "OpportunityType": "NonStandard_Direct",
                    "Hall": hall,
                    "Locations": [loc1, loc2], # Przekazujemy listę lokalizacji
                    "Notes": ""
                })
        elif capacity == 3:
            loc1 = loc_by_pos.get(1)
            loc2 = loc_by_pos.get(2)
            loc3 = loc_by_pos.get(3)

            # Sprawdź, czy wszystkie 3 są wolne (miejsce na 2 palety niestandardowe)
            if loc1 and loc1["is_empty"] and loc2 and loc2["is_empty"] and loc3 and loc3["is_empty"]:
                 opportunities.append({
                    "OpportunityType": "NonStandard_Direct_3_Full", # Specjalny typ
                    "Hall": hall,
                    "Locations": [loc1, loc2, loc3],
                    "Notes": ""
                 })
            else:
                # Sprawdź parę 1-2
                if loc1 and loc1["is_empty"] and loc2 and loc2["is_empty"]:
                     opportunities.append({
                        "OpportunityType": "NonStandard_Direct",
                        "Hall": hall,
                        "Locations": [loc1, loc2, loc3], # Przekazujemy wszystkie 3 dla kontekstu
                        "Pair": (1, 2), # Wskazujemy, która para jest wolna
                        "Notes": ""
                     })
                # Sprawdź parę 2-3
                if loc2 and loc2["is_empty"] and loc3 and loc3["is_empty"]:
                     opportunities.append({
                        "OpportunityType": "NonStandard_Direct",
                        "Hall": hall,
                        "Locations": [loc1, loc2, loc3],
                        "Pair": (2, 3),
                        "Notes": ""
                     })

            # Sprawdź możliwość przesunięcia (1 wolna, 2 zajęta, 3 wolna)
            if loc1 and loc1["is_empty"] and loc3 and loc3["is_empty"] and loc2 and not loc2["is_empty"]:
                 opportunities.append({
                    "OpportunityType": "NonStandard_MoveRequired",
                    "Hall": hall,
                    "Locations": [loc1, loc2, loc3], # Przekazujemy wszystkie 3
                    "OccupiedPos": 2, # Wskazujemy zajętą pozycję
                    "Notes": f"Przesuń paletę z {loc2['id']}" # Generujemy notatkę
                 })

    return opportunities


# ИЗМЕНЕНО: Новая функция форматирования с визуализацией
def format_analysis_results_visual(opportunities: List[Dict[str, Any]]) -> str:
    """Formatuje wyniki analizy w sposób wizualny, pogrupowane według hali."""

    # Grupujemy możliwości według hali
    opportunities_by_hall = defaultdict(lambda: {"direct": [], "move": []})
    for opp in opportunities:
        hall = opp.get("Hall", "Nieznana Hala")
        opp_type = opp.get("OpportunityType")
        locations = opp.get("Locations", [])
        notes = opp.get("Notes", "")
        pair = opp.get("Pair") # Dla NonStandard_Direct w sekcji 3
        occupied_pos = opp.get("OccupiedPos") # Dla NonStandard_MoveRequired

        # Sortujemy lokalizacje wewnątrz sekcji dla pewności
        locations.sort(key=lambda x: x.get("position", 0))

        if opp_type == "NonStandard_Direct":
            opportunities_by_hall[hall]["direct"].append({
                "locations": locations,
                "pair": pair # Przekazujemy parę
            })
        elif opp_type == "NonStandard_Direct_3_Full":
             opportunities_by_hall[hall]["direct"].append({
                "locations": locations,
                "pair": (1, 3) # Specjalny znacznik dla pełnej sekcji 3
             })
        elif opp_type == "NonStandard_MoveRequired":
            opportunities_by_hall[hall]["move"].append({
                "locations": locations,
                "occupied_pos": occupied_pos,
                "notes": notes
            })

    if not opportunities_by_hall:
        return "Nie znaleziono obecnie miejsc dla palet niestandardowych."

    output_lines = []
    # Sortujemy hale alfabetycznie
    sorted_halls = sorted(opportunities_by_hall.keys())

    for hall in sorted_halls:
        hall_data = opportunities_by_hall[hall]
        direct_ops = hall_data["direct"]
        move_ops = hall_data["move"]

        # Pomijamy halę, jeśli nie ma w niej żadnych możliwości
        if not direct_ops and not move_ops:
            continue

        output_lines.append("") # Pusta linia przed nową halą
        output_lines.append(f"--- SALA: {hall} ---")
        output_lines.append("-" * (len(hall) + 10)) # Linia podkreślająca

        if direct_ops:
            output_lines.append("\n**Miejsca GOTOWE na paletę niestandardową:**")
            # Sortujemy możliwości wewnątrz hali (np. po pierwszej lokalizacji)
            direct_ops.sort(key=lambda x: x["locations"][0].get("id", ""))
            for op in direct_ops:
                locs = op["locations"]
                pair = op.get("pair")
                loc_ids = [l.get('id', '???') for l in locs]

                if len(locs) == 2: # Sekcja 2-miejscowa
                    # [ LOK1 | LOK2 ] -> 1x Paleta Niestandardowa
                    vis = f"[{loc_ids[0]:<12} | {loc_ids[1]:<12}]"
                    output_lines.append(f"{vis}  ->  1x Paleta Niestandardowa")
                elif len(locs) == 3: # Sekcja 3-miejscowa
                    if pair == (1, 3): # Wszystkie 3 wolne
                         # [ LOK1 | LOK2 | LOK3 ] -> 2x Paleta Niestandardowa
                         vis = f"[{loc_ids[0]:<12} | {loc_ids[1]:<12} | {loc_ids[2]:<12}]"
                         output_lines.append(f"{vis}  ->  **2x Paleta Niestandardowa**")
                    elif pair == (1, 2): # Wolna para 1-2
                         # [ LOK1 | LOK2 | .... ] -> 1x Paleta Niestandardowa (na L1-L2)
                         vis = f"[{loc_ids[0]:<12} | {loc_ids[1]:<12} | {' ' * 12}]" # Puste miejsce dla L3
                         output_lines.append(f"{vis}  ->  1x Paleta Niestandardowa (na {loc_ids[0]}-{loc_ids[1]})")
                    elif pair == (2, 3): # Wolna para 2-3
                         # [ .... | LOK2 | LOK3 ] -> 1x Paleta Niestandardowa (na L2-L3)
                         vis = f"[{' ' * 12} | {loc_ids[1]:<12} | {loc_ids[2]:<12}]" # Puste miejsce dla L1
                         output_lines.append(f"{vis}  ->  1x Paleta Niestandardowa (na {loc_ids[1]}-{loc_ids[2]})")
            output_lines.append("") # Pusta linia po sekcji

        if move_ops:
            output_lines.append("\n**Miejsca dostępne PO PRZESUNIĘCIU palety:**")
            move_ops.sort(key=lambda x: x["locations"][0].get("id", ""))
            for op in move_ops:
                locs = op["locations"]
                occupied_pos = op["occupied_pos"]
                notes = op["notes"]
                loc_ids = {l.get('position', 0): l.get('id', '???') for l in locs}

                # Wizualizacja stanu początkowego
                initial_state_vis = "["
                initial_state_vis += f"{loc_ids.get(1, ''):<12} (wolne) | "
                initial_state_vis += f"{loc_ids.get(2, ''):<12} (ZAJĘTE)| "
                initial_state_vis += f"{loc_ids.get(3, ''):<12} (wolne)]"

                # Wizualizacja stanu końcowego
                final_state_vis = f"[{loc_ids.get(1, ''):<12} | {'-'*12} | {loc_ids.get(3, ''):<12}]" # Środkowa pusta

                output_lines.append(f"1. Stan obecny:   {initial_state_vis}")
                output_lines.append(f"2. Akcja:         **{notes}**") # Np. Przesuń paletę z F.A4.0
                output_lines.append(f"3. Wynik:         {final_state_vis}  ->  1x Paleta Niestandardowa")
                output_lines.append("") # Pusta linia między możliwościami przesunięcia

    return "\n".join(output_lines)


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
results_placeholder = st.empty()
results_placeholder.info("Załaduj plik pustych lokalizacji i kliknij 'Uruchom analizę'.")

# --- Логика выполнения при нажатии кнопки ---
analysis_results_text = "" # Zmienna do przechowywania tekstu do druku

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
                # ИЗМЕНЕНО: Вызываем новую функцию форматирования
                analysis_results_text = format_analysis_results_visual(opportunities)
                # Используем markdown для лучшего отображения жирного шрифта и пробелов
                results_placeholder.markdown(f"<pre style='font-size: 0.9em; line-height: 1.3;'>{analysis_results_text}</pre>", unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                analysis_results_text = ""

        else:
            results_placeholder.warning("Analiza nie może zostać wykonana (błąd odczytu pliku pustych lokalizacji).")
            analysis_results_text = ""

# --- Кнопка Печати ---
if analysis_results_text and "Nie znaleziono obecnie miejsc" not in analysis_results_text:
    st.sidebar.header("3. Drukuj")
    print_button_html = """
    <style>
    .print-button { /* Стили кнопки */ }
    /* Добавляем стили для печати, чтобы скрыть ненужные элементы */
    @media print {
      header, .stSidebar, .css-1rs6os.e17vqv3v0, .css-1rs6os.e17vqv3v0 > *, .stButton, .stDownloadButton, .stFileUploader, .stSpinner, .stAlert, .stInfo, .stSuccess, .stWarning, .stError {
        display: none !important;
      }
      /* Стили для основного контента при печати */
      .main .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        width: 100% !important; /* Занять всю ширину листа */
        max-width: 100% !important;
      }
      pre { /* Стили для блока с результатами */
          font-size: 9pt !important; /* Уменьшаем шрифт для печати */
          line-height: 1.2 !important;
          white-space: pre-wrap !important; /* Разрешаем перенос строк */
          word-wrap: break-word !important;
      }
      h1, h2, h3 { /* Стили для заголовков при печати */
          margin-top: 0.5em !important;
          margin-bottom: 0.2em !important;
          font-size: 11pt !important;
      }
    }
    </style>
    <button class="print-button" onclick="window.print()">Drukuj wyniki</button>
    """
    st.sidebar.markdown(print_button_html, unsafe_allow_html=True)
