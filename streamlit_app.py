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
        st.success("Plik struktury magazynu (layout.json) pomyślnie załadowany z repozytorium.")
        return data
    except Exception as e: st.error(f"Błąd odczytu pliku layout.json: {e}"); return None

# Функция чтения пустых локаций (без изменений)
def read_empty_locations_from_b3(uploaded_file) -> Optional[Set[str]]:
    if uploaded_file is None: st.error("Plik pustych lokalizacji (Excel/CSV) nie został załadowany."); return None
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

# ИЗМЕНЕНО: Функция анализа теперь возвращает более подробную информацию для визуализации
def analyze_opportunities_internal(layout_data: Dict, empty_locations_set: Set[str]) -> List[Dict[str, Any]]:
    """
    Analizuje możliwości i zwraca listę słowników z detalami dla wizualizacji.
    Każdy słownik reprezentuje sekcję z możliwością i zawiera listę jej lokalizacji.
    """
    opportunities_details = [] # Список для хранения деталей секций с возможностями
    sections_data = defaultdict(lambda: {"capacity": 0, "locations": [], "has_opportunity": False})
    layout = layout_data.get("Layout", {})

    # 1. Сбор данных по секциям
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
                    if not section_id: continue # Интересуют только секции

                    loc_num = loc_info.get("Number")
                    if loc_num is None: continue

                    location_id = f"{hall}.{row}{loc_num}.{level_str}"
                    is_accessible = loc_info.get("IsAccessible", True)
                    is_corridor = loc_info.get("IsCorridor", False)
                    section_capacity = loc_info.get("SectionCapacity", 0)
                    pos_in_section = loc_info.get("PositionInSection", 0)
                    is_empty = location_id in empty_locations_set

                    location_details = {
                        "id": location_id,
                        "is_accessible": is_accessible,
                        "is_corridor": is_corridor,
                        "is_empty": is_empty,
                        "position": pos_in_section,
                    }

                    if not sections_data[section_id]["capacity"]:
                        sections_data[section_id]["capacity"] = section_capacity
                    sections_data[section_id]["locations"].append(location_details)

    # 2. Анализ каждой секции на наличие возможностей для НЕСТАНДАРТНЫХ палет
    for section_id, section_info in sections_data.items():
        locations_in_section = sorted(section_info["locations"], key=lambda x: x["position"])
        capacity = section_info["capacity"]
        loc_by_pos = {loc["position"]: loc for loc in locations_in_section}
        found_opportunity_in_section = False

        # Проверка NonStandard_Direct (Capacity 2)
        if capacity == 2:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2)
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"]):
                section_info["opportunity_type"] = "NonStandard_Direct"
                section_info["involved_positions"] = [1, 2]
                found_opportunity_in_section = True

        # Проверка NonStandard_Direct и MoveRequired (Capacity 3)
        elif capacity == 3:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2); loc3 = loc_by_pos.get(3)
            can_place_12 = (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                            loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"])
            can_place_23 = (loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"] and
                            loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"])
            can_place_13_move2 = (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                                  loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"] and
                                  loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and not loc2["is_empty"]) # loc2 ЗАНЯТА

            if can_place_12 or can_place_23:
                 section_info["opportunity_type"] = "NonStandard_Direct"
                 # Определяем, какие именно пары доступны
                 involved = []
                 if can_place_12: involved.extend([1, 2])
                 if can_place_23: involved.extend([2, 3])
                 section_info["involved_positions"] = sorted(list(set(involved))) # [1, 2] или [2, 3] или [1, 2, 3]
                 found_opportunity_in_section = True
            elif can_place_13_move2:
                 section_info["opportunity_type"] = "NonStandard_MoveRequired"
                 section_info["involved_positions"] = [1, 3] # Позиции, куда ставим
                 section_info["move_required_from_pos"] = 2 # Позиция, откуда двигаем
                 section_info["move_required_from_id"] = loc2.get("id", "NIEZNANA") if loc2 else "NIEZNANA"
                 found_opportunity_in_section = True

        # Если в секции найдена возможность, добавляем ее в итоговый список
        if found_opportunity_in_section:
            section_info["section_id_internal"] = section_id # Сохраняем для группировки, но не показываем
            opportunities_details.append(section_info)

    # Сортируем результат (например, по ID первой локации в секции) для консистентности
    opportunities_details.sort(key=lambda x: x["locations"][0]["id"] if x["locations"] else "")
    st.write(f"(Analiza: Znaleziono {len(opportunities_details)} sekcji z możliwościami dla palet niestandardowych)")
    return opportunities_details


# НОВАЯ Функция форматирования для ВИЗУАЛИЗАЦИИ
def format_opportunities_visual(opportunities_details: List[Dict[str, Any]]) -> str:
    """Formatuje wyniki analizy jako bloki HTML/CSS dla wizualizacji."""

    if not opportunities_details:
        return "<h3>Nie znaleziono obecnie miejsc dla palet niestandardowych.</h3>"

    # CSS Стили для блоков
    # Добавляем стили для печати (@media print)
    css_styles = """
    <style>
        .opportunity-box {
            border: 3px solid #555; /* "Балка" вокруг секции */
            border-radius: 5px;
            margin-bottom: 15px;
            padding: 10px;
            background-color: #f9f9f9;
            page-break-inside: avoid; /* Стараемся не разрывать блок при печати */
        }
        .opportunity-title {
            font-weight: bold;
            margin-bottom: 8px;
            font-size: 1.1em;
        }
        .location-row {
            display: flex; /* Располагаем ячейки в ряд */
            gap: 5px; /* Небольшой отступ между ячейками */
        }
        .location-cell {
            border: 1px solid #ccc;
            padding: 8px;
            text-align: center;
            min-width: 100px; /* Минимальная ширина ячейки */
            flex: 1; /* Растягиваем ячейки */
            border-radius: 3px;
            font-weight: bold;
            display: flex; /* Центрируем текст вертикально */
            flex-direction: column;
            justify-content: center;
            min-height: 50px; /* Минимальная высота */
        }
        .location-id {
            font-size: 0.9em;
            word-wrap: break-word; /* Перенос длинных ID */
        }
        .cell-note {
            font-size: 0.8em;
            font-weight: normal;
            margin-top: 3px;
            color: #d9534f; /* Красный цвет для примечания */
        }
        .empty-cell {
            background-color: #dff0d8; /* Светло-зеленый - пусто */
            border-color: #b2dba1;
            color: #3c763d;
        }
        .occupied-cell {
            background-color: #fcf8e3; /* Светло-желтый - занято (для перемещения) */
            border-color: #f8e7b5;
            color: #8a6d3b;
        }
        /* Стили для печати */
        @media print {
            body {
                font-size: 10pt; /* Уменьшаем шрифт для печати */
                color: black; /* Черный текст для печати */
            }
            .opportunity-box {
                border: 2px solid black !important; /* Четкие черные границы */
                background-color: white !important; /* Белый фон */
                margin-bottom: 10px;
                padding: 5px;
            }
            .location-cell {
                border: 1px solid black !important;
                min-width: 80px; /* Уменьшаем ширину для А4 */
                padding: 4px;
                min-height: 40px;
            }
            .empty-cell {
                background-color: #e9f5e5 !important; /* Чуть бледнее зеленый */
                color: black !important;
            }
            .occupied-cell {
                background-color: #fff9e6 !important; /* Чуть бледнее желтый */
                color: black !important;
            }
            .cell-note {
                 color: black !important; /* Черный текст примечания */
                 font-style: italic;
            }
            /* Скрываем элементы Streamlit, которые не нужны при печати */
            header, .stSidebar, .stButton, .stDownloadButton, .stFileUploader, .stInfo, .stSuccess, .stWarning, .stError {
                display: none !important;
            }
            /* Растягиваем основной контент */
            .main .block-container {
                 max-width: 100% !important;
                 padding: 1cm !important; /* Поля для печати */
            }
        }
    </style>
    """

    html_parts = [css_styles] # Начинаем с CSS

    # Группируем возможности по типу
    direct_opportunities = [opp for opp in opportunities_details if opp.get("opportunity_type") == "NonStandard_Direct"]
    move_opportunities = [opp for opp in opportunities_details if opp.get("opportunity_type") == "NonStandard_MoveRequired"]

    # 1. Вывод возможностей "Можно поставить OD RAZU"
    if direct_opportunities:
        html_parts.append("<h2>Można postawić OD RAZU:</h2>")
        for opp in direct_opportunities:
            html_parts.append('<div class="opportunity-box">')
            # Собираем ID локаций для заголовка (например, F.A1.0 - F.A2.0)
            loc_ids = sorted([loc["id"] for loc in opp["locations"]])
            title_range = f"{loc_ids[0]} - {loc_ids[-1]}" if len(loc_ids) > 1 else loc_ids[0]
            html_parts.append(f'<div class="opportunity-title">Para miejsc: {title_range}</div>')
            html_parts.append('<div class="location-row">')
            # Отображаем все ячейки секции
            for loc in opp["locations"]:
                # Все ячейки в этой возможности должны быть пустыми
                cell_class = "location-cell empty-cell"
                html_parts.append(f'<div class="{cell_class}">')
                html_parts.append(f'<span class="location-id">{loc["id"]}</span>')
                html_parts.append('</div>') # end location-cell
            html_parts.append('</div>') # end location-row
            html_parts.append('</div>') # end opportunity-box

    # 2. Вывод возможностей "Można postawić PO PRZESUNIĘCIU"
    if move_opportunities:
        html_parts.append("<h2>Można postawić PO PRZESUNIĘCIU:</h2>")
        for opp in move_opportunities:
            html_parts.append('<div class="opportunity-box">')
            loc_ids = sorted([loc["id"] for loc in opp["locations"]])
            title_range = f"{loc_ids[0]} - {loc_ids[-1]}" if len(loc_ids) > 1 else loc_ids[0]
            move_from_pos = opp.get("move_required_from_pos")
            move_from_id = opp.get("move_required_from_id", "NIEZNANA")
            html_parts.append(f'<div class="opportunity-title">Zwolnij miejsce w: {title_range}</div>')
            html_parts.append('<div class="location-row">')
            # Отображаем все ячейки секции
            for loc in opp["locations"]:
                cell_class = "location-cell"
                note_html = ""
                if loc["position"] == move_from_pos:
                    cell_class += " occupied-cell" # Ячейка, которую нужно освободить
                    note_html = f'<div class="cell-note">(Przesuń paletę z {loc["id"]})</div>'
                else:
                    # Остальные должны быть пустыми в этом сценарии
                    cell_class += " empty-cell"

                html_parts.append(f'<div class="{cell_class}">')
                html_parts.append(f'<span class="location-id">{loc["id"]}</span>')
                html_parts.append(note_html) # Добавляем примечание, если есть
                html_parts.append('</div>') # end location-cell
            html_parts.append('</div>') # end location-row
            html_parts.append('</div>') # end opportunity-box

    # Если не было найдено ни одного типа возможностей
    if not direct_opportunities and not move_opportunities:
         html_parts.append("<h3>Nie znaleziono obecnie miejsc dla palet niestandardowych.</h3>")


    return "\n".join(html_parts)


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
results_placeholder = st.empty() # Используем empty для возможности замены контента
results_placeholder.info("Załaduj plik pustych lokalizacji i kliknij 'Uruchom analizę'.")

# --- Логика выполнения при нажатии кнопки ---
analysis_run_success = False # Флаг для показа кнопки печати

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
                # ИЗМЕНЕНО: Получаем детальные данные
                opportunities_details = analyze_opportunities_internal(layout_data, empty_locations)
                # ИЗМЕНЕНО: Генерируем HTML для визуализации
                analysis_results_html = format_opportunities_visual(opportunities_details)
                # ИЗМЕНЕНО: Отображаем HTML
                results_placeholder.markdown(analysis_results_html, unsafe_allow_html=True)
                analysis_run_success = True # Анализ прошел успешно

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                analysis_run_success = False

        else:
            results_placeholder.warning("Analiza nie może zostać wykonana (błąd odczytu pliku pustych lokalizacji).")
            analysis_run_success = False

# --- Кнопка Печати ---
# Показываем кнопку только если анализ прошел успешно и есть что печатать
if analysis_run_success:
    st.sidebar.header("3. Drukuj")
    print_button_html = """
    <style>
    /* Стили для кнопки печати (можно вынести в общий CSS выше) */
    .print-button { display: inline-block; padding: 0.5em 1em; border: 1px solid #ccc; border-radius: 4px; background-color: #f0f0f0; color: #333; text-align: center; text-decoration: none; cursor: pointer; font-size: 1em; font-family: inherit; }
    .print-button:hover { background-color: #e0e0e0; }
    .print-button:active { background-color: #d0d0d0; }
    </style>
    <button class="print-button" onclick="window.print()">Drukuj wyniki</button>
    """
    st.sidebar.markdown(print_button_html, unsafe_allow_html=True)
