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
        if not isinstance(data, dict) or "Layout" not in data or not isinstance(data.get("Layout"), dict):
             raise ValueError("Plik layout.json nie zawiera oczekiwanej struktury JSON.")
        # st.success("Plik struktury magazynu (layout.json) pomyślnie załadowany z repozytorium.")
        return data
    except FileNotFoundError: st.error(f"Plik layout.json nie znaleziony."); return None
    except json.JSONDecodeError as e: st.error(f"Błąd parsowania pliku layout.json: {e}"); return None
    except ValueError as e: st.error(f"Błąd walidacji struktury pliku layout.json: {e}"); return None
    except Exception as e: st.error(f"Nieoczekiwany błąd podczas odczytu pliku layout.json: {e}"); return None

# Функция чтения пустых локаций (без изменений)
def read_empty_locations_from_b3(uploaded_file) -> Optional[Set[str]]:
    if uploaded_file is None: st.error("Plik pustych lokalizacji (Excel/CSV) nie został załadowany."); return None
    file_name = uploaded_file.name
    st.info(f"Odczyt pliku pustych lokalizacji: {file_name} (dane od B3)")
    try:
        excel_engine = None; is_csv = False
        file_ext = os.path.splitext(file_name)[1].lower()
        if file_ext == '.csv': is_csv = True
        elif file_ext == '.xls':
            excel_engine = 'xlrd'; import xlrd
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
                except UnicodeDecodeError: continue
                except Exception as e_csv:
                     try: file_content.seek(0); pd.read_csv(file_content, delimiter=';', header=2, encoding=enc, nrows=1); raise e_csv
                     except Exception: continue
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

# Функция анализа (без изменений в логике, но убраны Standard возможности)
def analyze_opportunities_internal(layout_data: Dict, empty_locations_set: Set[str]) -> List[Dict[str, Any]]:
    """Wykonuje analizę możliwości tylko dla palet niestandardowych."""
    opportunities = []
    sections_data = defaultdict(lambda: {"capacity": 0, "locations": [], "hall": ""}) # Добавим зал
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
                        if not sections_data[section_id]["capacity"]:
                            sections_data[section_id]["capacity"] = section_capacity
                            sections_data[section_id]["hall"] = hall # Сохраняем зал секции
                        sections_data[section_id]["locations"].append(location_details)

    # Анализируем собранные секции на нестандартные возможности
    for section_id, section_info in sections_data.items():
        locations_in_section = sorted(section_info["locations"], key=lambda x: x["position"])
        capacity = section_info["capacity"]
        hall = section_info["hall"] # Получаем зал
        loc_by_pos = {loc["position"]: loc for loc in locations_in_section}

        # Проверяем только нестандартные
        if capacity == 2:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2)
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"]):
                opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc2["id"], "SectionID": section_id, "Notes": "", "Hall": hall, "Locations": [loc1, loc2]}) # Добавляем детали локаций
        elif capacity == 3:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2); loc3 = loc_by_pos.get(3)
            # Возможность 1: 1 и 2 пустые
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"]):
                opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc2["id"], "SectionID": section_id, "Notes": "", "Hall": hall, "Locations": [loc1, loc2, loc3]}) # Добавляем все 3 локации
            # Возможность 2: 2 и 3 пустые
            if (loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"] and
                loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"]):
                 # Проверяем, не дублирует ли это предыдущую возможность (если все 3 пусты)
                 is_duplicate = any(o["SectionID"] == section_id and o["PrimaryLocation"] == loc1["id"] for o in opportunities if loc1)
                 if not is_duplicate:
                      opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc2["id"], "SecondaryLocation": loc3["id"], "SectionID": section_id, "Notes": "", "Hall": hall, "Locations": [loc1, loc2, loc3]})
            # Возможность 3: 1 и 3 пустые, 2 занята
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and
                loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"] and
                loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and not loc2["is_empty"]):
                opportunities.append({"OpportunityType": "NonStandard_MoveRequired", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc3["id"], "SectionID": section_id, "Notes": f"Przesuń paletę z {loc2['id']}", "Hall": hall, "Locations": [loc1, loc2, loc3]})

    # st.write(f"(Analiza: Znaleziono {len(opportunities)} możliwości)") # Убрано
    return opportunities


# НОВАЯ Функция форматирования с визуализацией секций
def format_sections_for_display(opportunities: List[Dict[str, Any]]) -> Dict[str, List[str]]:
    """Formatuje wyniki analizy jako wizualne przedstawienie sekcji, pogrupowane według hali."""

    output_by_hall = defaultdict(list)

    # Группируем возможности по SectionID, чтобы обработать каждую секцию один раз
    opportunities_by_section = defaultdict(list)
    for opp in opportunities:
        # Используем SectionID как ключ, но сохраняем всю информацию о возможности
        opportunities_by_section[opp["SectionID"]].append(opp)

    processed_sections = set() # Чтобы избежать дублирования секций

    for section_id, opp_list in opportunities_by_section.items():
        if section_id in processed_sections:
            continue

        # Берем информацию о секции из первой возможности (они все относятся к одной секции)
        first_opp = opp_list[0]
        hall = first_opp.get("Hall", "Nieznany")
        all_locations_in_section = sorted(first_opp.get("Locations", []), key=lambda x: x.get("position", 0))
        capacity = len(all_locations_in_section)

        # Формируем визуальное представление секции
        section_vis = []
        loc_map = {} # Для быстрого доступа к деталям локации по ID
        for loc in all_locations_in_section:
            loc_id = loc.get("id", "?")
            loc_map[loc_id] = loc
            is_empty = loc.get("is_empty", False)
            is_accessible = loc.get("is_accessible", True)
            is_corridor = loc.get("is_corridor", False)

            if not is_accessible or is_corridor:
                state_char = "[XXX]" # Недоступна или коридор
            elif is_empty:
                state_char = "[   ]" # Пусто
            else:
                state_char = "[ P ]" # Занято (Paleta)
            section_vis.append(f"{loc_id}{state_char}")

        section_str = "  ".join(section_vis)

        # Определяем тип возможности для этой секции
        notes = []
        can_place_direct = 0
        can_place_move = False
        move_note_text = ""

        for opp in opp_list: # Проходим по всем возможностям для этой секции
            opp_type = opp.get("OpportunityType")
            if opp_type == "NonStandard_Direct":
                can_place_direct += 1
            elif opp_type == "NonStandard_MoveRequired":
                can_place_move = True
                move_note_text = opp.get("Notes", "") # Берем примечание о перемещении

        # Формируем итоговую строку для секции
        result_line = f"{section_str}   ->   "

        if can_place_move:
            # Приоритет на перемещение, т.к. оно уникально
            result_line += f"**Zwolnij miejsce na 1 paletę Niestandardową** ({move_note_text})"
        elif can_place_direct > 0:
            if capacity == 3 and can_place_direct == 2: # Все 3 места пусты
                 result_line += "**Zwolnij miejsce na 2 palety Niestandardowe**"
            elif can_place_direct >= 1: # Либо секция из 2х пуста, либо 2 из 3х пусты
                 result_line += "**Zwolnij miejsce na 1 paletę Niestandardową**"
            else: # Не должно произойти, но на всякий случай
                 result_line = None # Не показываем эту секцию
        else: # Нет подходящих возможностей
             result_line = None # Не показываем эту секцию


        if result_line:
            output_by_hall[hall].append(result_line)
            processed_sections.add(section_id) # Отмечаем секцию как обработанную

    # Сортируем списки внутри каждой галлереи (опционально)
    for hall in output_by_hall:
        output_by_hall[hall].sort()

    return output_by_hall


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
analysis_results_for_print = {} # Словарь для хранения результатов для печати

if run_button:
    results_placeholder.info("Wczytywanie struktury magazynu...")
    layout_data = read_layout_data()

    if layout_data is None:
         results_placeholder.error("Nie udało się wczytać pliku struktury magazynu (layout.json).")
         analysis_results_for_print = {} # Очищаем при ошибке
    else:
        results_placeholder.info("Przetwarzanie pliku pustych lokalizacji...")
        empty_locations = read_empty_locations_from_b3(uploaded_empty_loc_file)

        if empty_locations is not None:
            results_placeholder.info("Wykonywanie analizy...")
            try:
                opportunities = analyze_opportunities_internal(layout_data, empty_locations)
                # ИЗМЕНЕНО: Используем новую функцию форматирования
                analysis_results_for_print = format_sections_for_display(opportunities)

                # Очищаем предыдущий вывод
                results_placeholder.empty()

                if not analysis_results_for_print:
                    results_placeholder.warning("Nie znaleziono miejsc dla palet niestandardowych.")
                else:
                    # Выводим результаты по залам
                    st.markdown("--- Miejsca dla palet NIESTANDARDOWYCH (na 2 miejsca) ---")
                    # Сортируем залы по имени
                    sorted_halls = sorted(analysis_results_for_print.keys())
                    for hall in sorted_halls:
                        st.subheader(f"Hala: {hall}")
                        # Используем text_area для сохранения форматирования и прокрутки
                        st.text_area(f"Możliwości w hali {hall}:",
                                     value="\n".join(analysis_results_for_print[hall]),
                                     height=len(analysis_results_for_print[hall]) * 30 + 10, # Динамическая высота
                                     key=f"hall_{hall}_results" # Уникальный ключ
                                     )
                        st.markdown("---") # Разделитель между залами

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                analysis_results_for_print = {} # Очищаем при ошибке

        else:
            results_placeholder.warning("Analiza nie może zostać wykonana (błąd odczytu pliku pustych lokalizacji).")
            analysis_results_for_print = {} # Очищаем при ошибке

# --- Кнопка Печати ---
# Показываем кнопку только если есть результаты для печати
if analysis_results_for_print:
    st.sidebar.header("3. Drukuj")
    print_button_html = """
    <style>
    /* Стили для кнопки печати */
    .print-button { display: inline-block; padding: 0.5em 1em; border: 1px solid #ccc; border-radius: 4px; background-color: #f0f0f0; color: #333; text-align: center; text-decoration: none; cursor: pointer; font-size: 1em; font-family: inherit; }
    .print-button:hover { background-color: #e0e0e0; }
    .print-button:active { background-color: #d0d0d0; }
    /* Стили для печати: скрываем боковую панель и ненужные элементы */
    @media print {
        .stApp > header, .stApp > footer, div[data-testid="stSidebar"], div[data-testid="stToolbar"], .stDeployButton { display: none !important; }
        /* Можно добавить стили для основного контента, если нужно */
        div[data-testid="stVerticalBlock"] { margin: 0 !important; padding: 0 !important; }
        div[data-testid="stTextArea"] > label { display: none !important; } /* Скрыть заголовки text_area */
        div[data-testid="stTextArea"] > div > textarea {
             height: auto !important; /* Автоматическая высота для печати */
             overflow-y: visible !important;
             border: none !important; /* Убрать рамку у text_area */
             font-size: 9pt !important; /* Уменьшить шрифт для печати */
             line-height: 1.2 !important;
        }
         h3 { font-size: 11pt !important; margin-top: 10px !important; margin-bottom: 5px !important; } /* Стили для заголовков залов */
         hr { display: none !important; } /* Скрыть разделители */
    }
    </style>
    <button class="print-button" onclick="window.print()">Drukuj wyniki</button>
    """
    st.sidebar.markdown(print_button_html, unsafe_allow_html=True)
