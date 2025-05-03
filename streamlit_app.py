import streamlit as st
import pandas as pd
import json
import os
import io # Для работы с загруженными файлами в памяти
from collections import defaultdict, Counter
from typing import List, Dict, Optional, Set, Any
import traceback # Для отладки ошибок

# --- Логика анализа ---

# Функция чтения локального layout.json
def read_layout_data() -> Optional[Dict]:
    """Czyta plik layout.json z repozytorium."""
    file_path = "layout.json"
    try:
        if not os.path.exists(file_path):
             st.warning(f"Plik '{file_path}' nie znaleziony. Używana jest pusta struktura.")
             return {"Version": "1.0", "Layout": {}}
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, dict) or "Layout" not in data or not isinstance(data.get("Layout"), dict):
             raise ValueError("Plik layout.json nie zawiera oczekiwanej struktury JSON.")
        # Убираем сообщение об успехе загрузки layout, т.к. это происходит автоматически
        # st.success("Plik struktury magazynu (layout.json) pomyślnie załadowany z repozytorium.")
        return data
    except FileNotFoundError:
        st.error(f"Plik layout.json nie znaleziony. Upewnij się, że jest w repozytorium GitHub.")
        return None
    except json.JSONDecodeError as e:
        st.error(f"Błąd parsowania pliku layout.json: {e}")
        return None
    except ValueError as e:
        st.error(f"Błąd walidacji struktury pliku layout.json: {e}")
        return None
    except Exception as e:
        st.error(f"Nieoczekiwany błąd podczas odczytu pliku layout.json: {e}")
        return None

# Функция чтения пустых локаций (без изменений)
def read_empty_locations_from_b3(uploaded_file) -> Optional[Set[str]]:
    if uploaded_file is None:
        st.error("Plik pustych lokalizacji (Excel/CSV) nie został załadowany.")
        return None
    file_name = uploaded_file.name
    st.info(f"Odczyt pliku pustych lokalizacji: {file_name} (dane od B3)")
    try:
        excel_engine = None; is_csv = False
        file_ext = os.path.splitext(file_name)[1].lower()
        if file_ext == '.csv': is_csv = True
        elif file_ext == '.xls':
            excel_engine = 'xlrd'
            try: import xlrd
            except ImportError: raise ImportError("Wymagana biblioteka 'xlrd' do odczytu plików .xls.")
        elif file_ext == '.xlsx': excel_engine = 'openpyxl'
        else: st.error(f"Nieobsługiwany format pliku: {file_ext}. Użyj .xlsx, .xls lub .csv."); return None
        df: Optional[pd.DataFrame] = None
        file_content = io.BytesIO(uploaded_file.getvalue())
        if is_csv:
            detected_encoding = None
            for enc in ['utf-8', 'windows-1251']:
                try:
                    file_content.seek(0)
                    df = pd.read_csv(file_content, delimiter=';', header=2, usecols=[1], encoding=enc, skipinitialspace=True, on_bad_lines='skip')
                    st.write(f"(CSV odczytany z kodowaniem {enc})"); detected_encoding = enc; break
                except UnicodeDecodeError: continue
                except Exception as e_csv:
                     try: file_content.seek(0); pd.read_csv(file_content, delimiter=';', header=2, encoding=enc, nrows=1); raise e_csv
                     except Exception: continue
            if df is None: raise ValueError("Nie udało się odczytać pliku CSV. Sprawdź kodowanie (UTF-8/Win-1251), separator (;) i obecność danych.")
        else: df = pd.read_excel(file_content, header=None, skiprows=2, usecols=[1], sheet_name=0, engine=excel_engine)
        if df is None or df.empty: st.warning("Plik pustych lokalizacji nie zawiera danych w kolumnie B począwszy od 3. wiersza."); return set()
        col_name = df.columns[0]
        locations = df[col_name].dropna().astype(str).str.strip()
        empty_locations_set = set(locations[locations != ''])
        count = len(empty_locations_set)
        st.success(f"Załadowano {count} unikalnych ID pustych lokalizacji z {file_name}.")
        return empty_locations_set
    except ImportError as e: st.error(f"Błąd importu biblioteki: {e}."); return None
    except ValueError as e: st.error(f"Błąd danych lub formatu pliku pustych lokalizacji: {e}"); return None
    except Exception as e: st.error(f"Nieoczekiwany błąd podczas odczytu pliku pustych lokalizacji: {e}"); return None

# Функция анализа (без изменений)
def analyze_opportunities_internal(layout_data: Dict, empty_locations_set: Set[str]) -> List[Dict[str, Any]]:
    # ... (код функции analyze_opportunities_internal остается таким же) ...
    opportunities = []
    sections_data = defaultdict(lambda: {"capacity": 0, "locations": []})
    layout = layout_data.get("Layout", {})
    processed_locations_count = 0
    for hall, rows in layout.items():
        if not isinstance(rows, dict): continue
        for row, levels in rows.items():
            if not isinstance(levels, dict): continue
            for level_str, level_data in levels.items():
                if not isinstance(level_data, dict): continue
                locations = level_data.get("Locations", [])
                if not isinstance(locations, list): continue
                for loc_info in locations:
                    processed_locations_count += 1
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
                    location_details = {"id": location_id, "number": loc_num, "is_accessible": is_accessible, "is_corridor": is_corridor, "is_empty": is_empty, "position": pos_in_section, "section_id": section_id, "section_capacity": section_capacity}
                    if section_id:
                        if not sections_data[section_id]["capacity"]: sections_data[section_id]["capacity"] = section_capacity
                        sections_data[section_id]["locations"].append(location_details)
                    # Убираем проверку на Standard здесь, т.к. она не нужна для вывода
                    # else:
                    #      if is_accessible and not is_corridor and is_empty: opportunities.append({"OpportunityType": "Standard", "PrimaryLocation": location_id, "SecondaryLocation": "", "SectionID": "", "Notes": ""})
    # Убираем вывод о количестве обработанных локаций/секций из лога Streamlit
    # st.write(f"(Analiza: Przetworzono {processed_locations_count} lokalizacji, {len(sections_data)} sekcji)")
    for section_id, section_info in sections_data.items():
        locations_in_section = sorted(section_info["locations"], key=lambda x: x["position"])
        capacity = section_info["capacity"]
        loc_by_pos = {loc["position"]: loc for loc in locations_in_section}
        # Убираем добавление Standard возможностей
        # for loc in locations_in_section:
        #     if loc["is_accessible"] and not loc["is_corridor"] and loc["is_empty"]: opportunities.append({"OpportunityType": "Standard", "PrimaryLocation": loc["id"], "SecondaryLocation": "", "SectionID": section_id, "Notes": ""})
        if capacity == 2:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2)
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"]): opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc2["id"], "SectionID": section_id, "Notes": ""})
        elif capacity == 3:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2); loc3 = loc_by_pos.get(3)
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"]): opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc2["id"], "SectionID": section_id, "Notes": ""})
            if (loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"] and loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"]): opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc2["id"], "SecondaryLocation": loc3["id"], "SectionID": section_id, "Notes": ""})
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"] and loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and not loc2["is_empty"]): opportunities.append({"OpportunityType": "NonStandard_MoveRequired", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc3["id"], "SectionID": section_id, "Notes": f"Требуется перемещение палеты из {loc2['id']}"}) # Примечание оставляем, оно используется ниже
    # Убираем вывод о количестве найденных возможностей из лога Streamlit
    # st.write(f"(Analiza: Znaleziono {len(opportunities)} możliwości)")
    return opportunities


# ИЗМЕНЕНО: Функция форматирования ТОЛЬКО для нестандартных палет
def format_analysis_results_simple(opportunities: List[Dict[str, Any]]) -> str:
    """Formatuje wyniki analizy tylko dla palet niestandardowych."""

    non_standard_direct = []
    non_standard_move = []

    # Собираем только нужные типы
    for opp in opportunities:
        opp_type = opp.get("OpportunityType")
        if opp_type == "NonStandard_Direct":
            non_standard_direct.append((opp.get("PrimaryLocation", ""), opp.get("SecondaryLocation", "")))
        elif opp_type == "NonStandard_MoveRequired":
            loc1 = opp.get("PrimaryLocation", "")
            loc3 = opp.get("SecondaryLocation", "") # Это loc3 из анализа
            notes = opp.get("Notes", "")
            move_note = ""
            # Переводим примечание
            if "Требуется перемещение палеты из " in notes:
                move_loc = notes.replace("Требуется перемещение палеты из ", "").strip()
                move_note = f"(Przesuń paletę z {move_loc})" # Польский перевод
            non_standard_move.append((loc1, loc3, move_note))

    # Если нет ни одного из нужных типов
    if not non_standard_direct and not non_standard_move:
        return "Nie znaleziono obecnie miejsc dla palet niestandardowych."

    # Формируем итоговый текст
    output_lines = []
    output_lines.append("--- Miejsca dla palet NIESTANDARDOWYCH (na 2 miejsca) ---")
    output_lines.append("")

    if non_standard_direct:
        output_lines.append("Można postawić OD RAZU w:")
        non_standard_direct.sort(key=lambda x: x[0]) # Сортируем
        for loc1, loc2 in non_standard_direct:
            output_lines.append(f"  - Lokalizacje: {loc1} i {loc2}")
        output_lines.append("") # Пустая строка после списка
    else:
        output_lines.append("Brak wolnych par miejsc, gdzie można postawić paletę od razu.")
        output_lines.append("")

    if non_standard_move:
        output_lines.append("Można postawić PO PRZESUNIĘCIU palety:")
        non_standard_move.sort(key=lambda x: x[0]) # Сортируем
        for loc1, loc3, move_note in non_standard_move:
            output_lines.append(f"  - Z: {loc1} i {loc3}  {move_note}") # Добавляем примечание
        output_lines.append("") # Пустая строка после списка
    else:
        # Не пишем ничего, если таких нет
        pass

    output_lines.append("-------------------------------------------------------")
    return "\n".join(output_lines)


# --- Streamlit UI ---
st.set_page_config(page_title="Wyszukiwarka Miejsc Magazynowych", layout="wide")
st.title("Wyszukiwarka Wolnych Miejsc dla Palet Niestandardowych") # Обновлено название

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
                opportunities = analyze_opportunities_internal(layout_data, empty_locations)
                # ИЗМЕНЕНО: Сохраняем результат форматирования
                analysis_results_text = format_analysis_results_simple(opportunities)
                results_placeholder.text_area("Wyniki:", value=analysis_results_text, height=400) # Уменьшил высоту

                # УБРАНА КНОПКА СКАЧИВАНИЯ EXCEL
                # if opportunities: ... st.sidebar.download_button(...)

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                analysis_results_text = "" # Очищаем текст при ошибке

        else:
            results_placeholder.warning("Analiza nie może zostać wykonana (błąd odczytu pliku pustych lokalizacji).")
            analysis_results_text = "" # Очищаем текст при ошибке

# --- Кнопка Печати ---
# Показываем кнопку только если есть результаты для печати
if analysis_results_text and "Nie znaleziono obecnie miejsc" not in analysis_results_text:
    st.sidebar.header("3. Drukuj")
    # Используем HTML и JavaScript для создания кнопки печати
    print_button_html = """
    <style>
    .print-button {
        display: inline-block;
        padding: 0.5em 1em;
        border: 1px solid #ccc;
        border-radius: 4px;
        background-color: #f0f0f0;
        color: #333;
        text-align: center;
        text-decoration: none;
        cursor: pointer;
        font-size: 1em; /* Размер шрифта как у обычного текста */
        font-family: inherit; /* Шрифт как у остального интерфейса */
    }
    .print-button:hover {
        background-color: #e0e0e0;
    }
    .print-button:active {
        background-color: #d0d0d0;
    }
    </style>
    <button class="print-button" onclick="window.print()">Drukuj wyniki</button>
    """
    st.sidebar.markdown(print_button_html, unsafe_allow_html=True)
