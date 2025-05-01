import streamlit as st
import pandas as pd
import json
import os
import io # Для работы с загруженными файлами в памяти
from collections import defaultdict, Counter
from typing import List, Dict, Optional, Set, Any
import traceback # Для отладки ошибок

# --- Логика анализа ---

# ИЗМЕНЕНО: Функция чтения локального layout.json
def read_layout_data() -> Optional[Dict]:
    """Читает файл layout.json, находящийся в том же каталоге."""
    file_path = "layout.json" # Имя файла в репозитории
    try:
        # Проверяем, существует ли файл (важно для локального запуска/отладки)
        if not os.path.exists(file_path):
             # В облаке Streamlit файл должен быть, если он в репо
             st.warning(f"Файл '{file_path}' не найден в репозитории. Используется пустая структура.")
             # Возвращаем пустую структуру, чтобы приложение не падало
             return {"Version": "1.0", "Layout": {}} 
             # Или можно вернуть None и обрабатывать это дальше:
             # st.error(f"Критическая ошибка: Файл '{file_path}' не найден в репозитории.")
             # return None


        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # Валидация
        if not isinstance(data, dict) or "Layout" not in data or not isinstance(data.get("Layout"), dict):
             raise ValueError("Файл layout.json не содержит ожидаемую структуру JSON (отсутствует ключ 'Layout').")
        st.success("Файл структуры склада (layout.json) успешно загружен из репозитория.")
        return data
    except FileNotFoundError: # Эта ошибка не должна возникать в облаке, если файл добавлен в репо
        st.error(f"Файл layout.json не найден. Убедитесь, что он добавлен в репозиторий GitHub.")
        return None
    except json.JSONDecodeError as e:
        st.error(f"Ошибка парсинга файла layout.json: {e}")
        return None
    except ValueError as e:
        st.error(f"Ошибка валидации структуры файла layout.json: {e}")
        return None
    except Exception as e:
        st.error(f"Неожиданная ошибка при чтении файла layout.json: {e}")
        # traceback.print_exc()
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
                    else:
                         if is_accessible and not is_corridor and is_empty: opportunities.append({"OpportunityType": "Standard", "PrimaryLocation": location_id, "SecondaryLocation": "", "SectionID": "", "Notes": ""})
    st.write(f"(Analiza: Przetworzono {processed_locations_count} lokalizacji, {len(sections_data)} sekcji)")
    for section_id, section_info in sections_data.items():
        locations_in_section = sorted(section_info["locations"], key=lambda x: x["position"])
        capacity = section_info["capacity"]
        loc_by_pos = {loc["position"]: loc for loc in locations_in_section}
        for loc in locations_in_section:
            if loc["is_accessible"] and not loc["is_corridor"] and loc["is_empty"]: opportunities.append({"OpportunityType": "Standard", "PrimaryLocation": loc["id"], "SecondaryLocation": "", "SectionID": section_id, "Notes": ""})
        if capacity == 2:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2)
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"]): opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc2["id"], "SectionID": section_id, "Notes": ""})
        elif capacity == 3:
            loc1 = loc_by_pos.get(1); loc2 = loc_by_pos.get(2); loc3 = loc_by_pos.get(3)
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"]): opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc2["id"], "SectionID": section_id, "Notes": ""})
            if (loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and loc2["is_empty"] and loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"]): opportunities.append({"OpportunityType": "NonStandard_Direct", "PrimaryLocation": loc2["id"], "SecondaryLocation": loc3["id"], "SectionID": section_id, "Notes": ""})
            if (loc1 and loc1["is_accessible"] and not loc1["is_corridor"] and loc1["is_empty"] and loc3 and loc3["is_accessible"] and not loc3["is_corridor"] and loc3["is_empty"] and loc2 and loc2["is_accessible"] and not loc2["is_corridor"] and not loc2["is_empty"]): opportunities.append({"OpportunityType": "NonStandard_MoveRequired", "PrimaryLocation": loc1["id"], "SecondaryLocation": loc3["id"], "SectionID": section_id, "Notes": f"Требуется перемещение палеты из {loc2['id']}"})
    st.write(f"(Analiza: Znaleziono {len(opportunities)} możliwości)")
    return opportunities


# Функция форматирования результатов (без изменений)
def format_analysis_results_simple(opportunities: List[Dict[str, Any]]) -> str:
    # ... (код функции format_analysis_results_simple остается таким же) ...
    if not opportunities: return "Nie znaleziono wolnych miejsc do umieszczenia palet."
    standard_locations = []; non_standard_direct = []; non_standard_move = []
    direct_pairs_locations = set(); move_pairs_locations = set()
    for opp in opportunities:
        if opp.get("OpportunityType") == "NonStandard_Direct":
            loc1 = opp.get("PrimaryLocation", ""); loc2 = opp.get("SecondaryLocation", "")
            direct_pairs_locations.add(loc1); direct_pairs_locations.add(loc2)
            non_standard_direct.append((loc1, loc2))
    for opp in opportunities:
         if opp.get("OpportunityType") == "NonStandard_MoveRequired":
              loc1 = opp.get("PrimaryLocation", ""); loc3 = opp.get("SecondaryLocation", ""); notes = opp.get("Notes", "")
              move_pairs_locations.add(loc1); move_pairs_locations.add(loc3); move_note = ""
              if "Требуется перемещение палеты из " in notes: move_loc = notes.replace("Требуется перемещение палеты из ", "").strip(); move_note = f"(Należy przesunąć paletę z {move_loc})"
              non_standard_move.append((loc1, loc3, move_note))
    for opp in opportunities:
        if opp.get("OpportunityType") == "Standard":
            loc1 = opp.get("PrimaryLocation", "")
            if loc1 not in direct_pairs_locations and loc1 not in move_pairs_locations: standard_locations.append(loc1)
    output_lines = []; output_lines.append("--- Dostępne miejsca ---"); output_lines.append("")
    if standard_locations:
        output_lines.append("STANDARDOWE palety można umieścić tutaj:"); standard_locations.sort(); line = []
        for i, loc in enumerate(standard_locations):
            line.append(loc)
            if (i + 1) % 5 == 0 or i == len(standard_locations) - 1: output_lines.append("  " + ", ".join(line)); line = []
        output_lines.append("")
    else: output_lines.append("Brak wolnych miejsc dla STANDARDOWYCH palet."); output_lines.append("")
    if non_standard_direct:
        output_lines.append("NIESTANDARDOWE palety (na 2 miejsca) można umieścić tutaj:"); non_standard_direct.sort(key=lambda x: x[0])
        for loc1, loc2 in non_standard_direct: output_lines.append(f"  - {loc1} i {loc2}")
        output_lines.append("")
    else: output_lines.append("Brak wolnych par miejsc dla NIESTANDARDOWYCH palet."); output_lines.append("")
    if non_standard_move:
        output_lines.append("NIESTANDARDOWE palety (na 2 miejsca) można umieścić, JEŚLI PRZESUNIESZ:"); non_standard_move.sort(key=lambda x: x[0])
        for loc1, loc3, move_note in non_standard_move: output_lines.append(f"  - {loc1} i {loc3} {move_note}")
        output_lines.append("")
    output_lines.append("------------------------------------")
    return "\n".join(output_lines)


# --- Streamlit UI (Изменения) ---
st.set_page_config(page_title="Wyszukiwarka Miejsc Magazynowych", layout="wide")
st.title("Wyszukiwarka Wolnych Miejsc Magazynowych")

st.sidebar.header("1. Załaduj plik")

# УБРАН ЗАГРУЗЧИК ДЛЯ layout.json
# uploaded_layout_file = st.sidebar.file_uploader(...)

# Загрузчик для файла пустых локаций (остается)
uploaded_empty_loc_file = st.sidebar.file_uploader(
    "Załaduj plik pustych lokalizacji (Excel/CSV)",
    type=["xlsx", "xls", "csv"]
)

st.sidebar.info("Plik struktury magazynu (layout.json) jest ładowany automatycznie z repozytorium.") # Добавлено инфо

st.sidebar.header("2. Uruchom analizę")
# ИЗМЕНЕНО: Кнопка активна только если загружен файл пустых локаций
run_button = st.sidebar.button("Uruchom analizę", disabled=(not uploaded_empty_loc_file))

st.header("Wyniki analizy")
results_placeholder = st.empty()
# ИЗМЕНЕНО: Обновлен текст-подсказка
results_placeholder.info("Załaduj plik pustych lokalizacji i kliknij 'Uruchom analizę' w panelu bocznym.")

# --- Логика выполнения при нажатии кнопки (Изменения) ---
if run_button:
    results_placeholder.info("Wczytywanie struktury magazynu...")
    QApplication.processEvents() # Даем интерфейсу обновиться

    # ИЗМЕНЕНО: Читаем локальный layout.json
    layout_data = read_layout_data()

    # Проверяем результат чтения layout_data
    if layout_data is None:
         results_placeholder.error("Nie udało się wczytać pliku struktury magazynu (layout.json). Analiza przerwana.")
    else:
        results_placeholder.info("Przetwarzanie pliku pustych lokalizacji...")
        QApplication.processEvents() # Даем интерфейсу обновиться
        empty_locations = read_empty_locations_from_b3(uploaded_empty_loc_file)

        # ИЗМЕНЕНО: Проверяем оба результата перед анализом
        if empty_locations is not None: # layout_data уже проверен выше
            results_placeholder.info("Wykonywanie analizy...")
            QApplication.processEvents() # Даем интерфейсу обновиться
            try:
                opportunities = analyze_opportunities_internal(layout_data, empty_locations)
                results_text = format_analysis_results_simple(opportunities)
                results_placeholder.text_area("Wyniki:", value=results_text, height=500)

                if opportunities:
                     df_report = pd.DataFrame(opportunities)
                     df_report = df_report[["OpportunityType", "PrimaryLocation", "SecondaryLocation", "SectionID", "Notes"]]
                     output_excel = io.BytesIO()
                     with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                          df_report.to_excel(writer, index=False, sheet_name='Opportunities')
                     excel_data = output_excel.getvalue()
                     st.sidebar.download_button(
                          label="Pobierz pełny raport (Excel)",
                          data=excel_data,
                          file_name="raport_analizy.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                     )

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                # traceback.print_exc()

        else:
            # Ошибка чтения файла пустых локаций уже выведена
            results_placeholder.warning("Analiza nie może zostać wykonana z powodu błędów odczytu pliku pustych lokalizacji.")
