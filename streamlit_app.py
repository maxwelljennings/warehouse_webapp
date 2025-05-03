import streamlit as st
import pandas as pd
import json
import os
import io
from collections import defaultdict
from typing import List, Dict, Optional, Set, Any
import traceback

# --- Логика анализа ---

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
        # st.success("Plik struktury magazynu (layout.json) pomyślnie załadowany z repozytorium.") # Убрано
        return data
    except Exception as e:
        st.error(f"Błąd odczytu pliku layout.json: {e}")
        return None

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

# ИЗМЕНЕНО: Функция анализа теперь возвращает словарь сгруппированный по залам
def analyze_opportunities_internal(layout_data: Dict, empty_locations_set: Set[str]) -> Dict[str, List[Dict[str, Any]]]:
    """Wykonuje analizę możliwości i grupuje wyniki według hali."""
    # Словарь для хранения результатов по залам
    opportunities_by_hall = defaultdict(list)
    layout = layout_data.get("Layout", {})

    # Итерация по структуре для сбора данных о секциях
    for hall, rows in layout.items():
        if not isinstance(rows, dict): continue
        for row, levels in rows.items():
            if not isinstance(levels, dict): continue
            for level_str, level_data in levels.items():
                if not isinstance(level_data, dict): continue

                # Группируем локации текущего уровня по SectionID
                locations_in_level = level_data.get("Locations", [])
                if not isinstance(locations_in_level, list): continue

                sections_in_level = defaultdict(list)
                for loc_info in locations_in_level:
                    if not isinstance(loc_info, dict): continue
                    section_id = loc_info.get("SectionID")
                    if section_id: # Собираем только локации внутри секций
                        loc_num = loc_info.get("Number")
                        if loc_num is None: continue
                        location_id = f"{hall}.{row}{loc_num}.{level_str}"
                        is_accessible = loc_info.get("IsAccessible", True)
                        is_corridor = loc_info.get("IsCorridor", False)
                        pos_in_section = loc_info.get("PositionInSection", 0)
                        is_empty = location_id in empty_locations_set
                        section_capacity = loc_info.get("SectionCapacity", 0) # Получаем емкость

                        sections_in_level[section_id].append({
                            "id": location_id,
                            "position": pos_in_section,
                            "is_empty": is_empty,
                            "is_accessible": is_accessible,
                            "is_corridor": is_corridor,
                            "capacity": section_capacity # Сохраняем емкость
                        })

                # Анализ собранных секций на текущем уровне
                for section_id, locations_in_section in sections_in_level.items():
                    # Проверяем, что все локации секции доступны и не коридоры
                    if not all(loc["is_accessible"] and not loc["is_corridor"] for loc in locations_in_section):
                        continue # Пропускаем секцию с недоступными/коридорами

                    locations_in_section.sort(key=lambda x: x["position"])
                    capacity = locations_in_section[0]["capacity"] if locations_in_section else 0
                    loc_by_pos = {loc["position"]: loc for loc in locations_in_section}

                    if capacity == 3:
                        loc1 = loc_by_pos.get(1)
                        loc2 = loc_by_pos.get(2)
                        loc3 = loc_by_pos.get(3)
                        if not (loc1 and loc2 and loc3): continue # Пропускаем неполные секции

                        # Случай 1: Все 3 свободны (2 нест. палеты)
                        if loc1["is_empty"] and loc2["is_empty"] and loc3["is_empty"]:
                            opportunities_by_hall[hall].append({
                                "type": "WOLNA_SEKCJA_3",
                                "locations": [loc1["id"], loc2["id"], loc3["id"]]
                            })
                        # Случай 2: 1 и 3 свободны, 2 занята (нужно перемещение для 1 нест. палеты)
                        elif loc1["is_empty"] and loc3["is_empty"] and not loc2["is_empty"]:
                             opportunities_by_hall[hall].append({
                                "type": "PRZESUN_SRODEK",
                                "locations": [loc1["id"], loc2["id"], loc3["id"]],
                                "move_from": loc2["id"] # Указываем, откуда двигать
                            })
                        # Другие комбинации в секции из 3 не дают возможности для *нестандартной* палеты

                    elif capacity == 2:
                        loc1 = loc_by_pos.get(1)
                        loc2 = loc_by_pos.get(2)
                        if not (loc1 and loc2): continue # Пропускаем неполные секции

                        # Случай 3: Обе свободны (1 нест. палета)
                        if loc1["is_empty"] and loc2["is_empty"]:
                            opportunities_by_hall[hall].append({
                                "type": "WOLNA_SEKCJA_2",
                                "locations": [loc1["id"], loc2["id"]]
                            })

    return opportunities_by_hall

# НОВАЯ функция форматирования с группировкой по залам
def create_results_markdown(opportunities_by_hall: Dict[str, List[Dict[str, Any]]]) -> str:
    """Formatuje wyniki analizy w Markdown z grupowaniem według hali."""

    if not opportunities_by_hall:
        return "## Nie znaleziono miejsc dla palet niestandardowych."

    markdown_lines = []
    markdown_lines.append("# Dostępne miejsca dla palet NIESTANDARDOWYCH")
    markdown_lines.append("*(Paleta niestandardowa zajmuje 2 miejsca)*")
    markdown_lines.append("---")

    # Сортируем залы по имени
    sorted_halls = sorted(opportunities_by_hall.keys())

    for hall in sorted_halls:
        opportunities_in_hall = opportunities_by_hall[hall]
        if not opportunities_in_hall: continue # Пропускаем зал без возможностей

        markdown_lines.append(f"## Hala: {hall}")
        markdown_lines.append("") # Пустая строка

        # Сортируем возможности внутри зала по первой локации
        opportunities_in_hall.sort(key=lambda x: x["locations"][0])

        has_direct_placement = False
        has_move_required = False

        # Сначала выводим возможности, где можно ставить сразу
        direct_placement_lines = []
        for opp in opportunities_in_hall:
            opp_type = opp["type"]
            locations = opp["locations"]
            if opp_type == "WOLNA_SEKCJA_3":
                direct_placement_lines.append(f"- **Cała sekcja 3-miejscowa wolna:** `{locations[0]}`, `{locations[1]}`, `{locations[2]}`")
                direct_placement_lines.append(f"  *(Można umieścić 2 palety niestandardowe)*")
                has_direct_placement = True
            elif opp_type == "WOLNA_SEKCJA_2":
                direct_placement_lines.append(f"- **Sekcja 2-miejscowa wolna:** `{locations[0]}`, `{locations[1]}`")
                direct_placement_lines.append(f"  *(Można umieścić 1 paletę niestandardową)*")
                has_direct_placement = True

        if has_direct_placement:
            markdown_lines.append("**Można postawić OD RAZU:**")
            markdown_lines.extend(direct_placement_lines)
            markdown_lines.append("") # Пустая строка

        # Затем выводим возможности, требующие перемещения
        move_required_lines = []
        for opp in opportunities_in_hall:
            opp_type = opp["type"]
            locations = opp["locations"]
            if opp_type == "PRZESUN_SRODEK":
                move_from_loc = opp["move_from"]
                loc1 = locations[0]
                loc3 = locations[2]
                move_required_lines.append(f"- **Przesuń paletę z:** `{move_from_loc}`")
                move_required_lines.append(f"  **Aby zwolnić miejsca:** `{loc1}` **i** `{loc3}`")
                move_required_lines.append(f"  *(Dla 1 palety niestandardowej)*")
                has_move_required = True

        if has_move_required:
             markdown_lines.append("**Można postawić PO PRZESUNIĘCIU palety:**")
             markdown_lines.extend(move_required_lines)
             markdown_lines.append("") # Пустая строка

        # Если в зале не было ни одного типа, добавим сообщение
        if not has_direct_placement and not has_move_required:
             markdown_lines.append("*Brak możliwości w tej hali.*")
             markdown_lines.append("")

        markdown_lines.append("---") # Разделитель между залами

    return "\n".join(markdown_lines)


# --- Streamlit UI ---
st.set_page_config(page_title="Wyszukiwarka Miejsc Magazynowych", layout="wide")
st.title("Wyszukiwarka Wolnych Miejsc dla Palet Niestandardowych")

st.sidebar.header("1. Załaduj plik")

uploaded_empty_loc_file = st.sidebar.file_uploader(
    "Załaduj plik pustych lokalizacji (Excel/CSV)",
    type=["xlsx", "xls", "csv"],
    help="Wybierz plik wygenerowany przez WMS. Dane powinny zaczynać się w komórce B3."
)

st.sidebar.info("Plik struktury magazynu (layout.json) jest ładowany automatycznie.")

st.sidebar.header("2. Uruchom analizę")
run_button = st.sidebar.button("Uruchom analizę", disabled=(not uploaded_empty_loc_file))

st.header("Wyniki analizy")
results_placeholder = st.empty()
results_placeholder.info("Załaduj plik pustych lokalizacji i kliknij 'Uruchom analizę'.")

# --- Логика выполнения при нажатии кнопки ---
analysis_results_markdown = "" # Переменная для хранения Markdown для печати

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
                opportunities_by_hall = analyze_opportunities_internal(layout_data, empty_locations)
                # ИЗМЕНЕНО: Генерируем Markdown
                analysis_results_markdown = create_results_markdown(opportunities_by_hall)
                # ИЗМЕНЕНО: Отображаем Markdown
                results_placeholder.markdown(analysis_results_markdown, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Błąd podczas wykonywania analizy: {e}")
                results_placeholder.error(f"Wystąpił błąd podczas analizy. Szczegóły: {e}")
                analysis_results_markdown = "" # Очищаем текст при ошибке

        else:
            results_placeholder.warning("Analiza nie może zostać wykonana (błąd odczytu pliku pustych lokalizacji).")
            analysis_results_markdown = "" # Очищаем текст при ошибке

# --- Кнопка Печати ---
# Показываем кнопку только если есть результаты для печати (проверяем наличие markdown)
if analysis_results_markdown and "## Nie znaleziono miejsc" not in analysis_results_markdown:
    st.sidebar.header("3. Drukuj")
    print_button_html = """
    <style>
    /* Стили для кнопки печати */
    @media print {
      /* Скрываем боковую панель и другие элементы при печати */
      div[data-testid="stSidebar"], header, footer {
        display: none !important;
      }
      /* Заставляем основной контент занимать всю ширину */
      div[data-testid="stAppViewContainer"] > section {
         margin-left: 0 !important;
         width: 100% !important;
         padding: 1cm !important; /* Добавляем поля для печати */
      }
       /* Скрываем саму кнопку печати при печати */
      .print-button-container {
          display: none !important;
      }
      /* Убираем лишние отступы у markdown */
       div[data-testid="stMarkdownContainer"] {
           padding-top: 0 !important;
       }
       /* Увеличиваем шрифт для читаемости при печати */
       body, p, li, h1, h2, h3 {
            font-size: 12pt !important;
       }
       code { /* Стиль для локаций в `обратных кавычках` */
            font-family: monospace !important;
            font-size: 11pt !important;
            background-color: #f0f0f0 !important; /* Светлый фон для выделения */
            padding: 1px 3px;
            border-radius: 3px;
            color: black !important; /* Черный текст */
            -webkit-print-color-adjust: exact !important; /* Chrome/Safari: принудительно печатать фон */
            print-color-adjust: exact !important; /* Стандарт: принудительно печатать фон */
       }
       strong { /* Стиль для выделенного текста */
            font-weight: bold !important;
       }
    }
    .print-button {
        display: inline-block; padding: 0.6em 1.2em; border: none; border-radius: 5px;
        background-color: #0068c9; color: white; text-align: center;
        text-decoration: none; cursor: pointer; font-size: 1em; font-family: inherit;
        width: 100%; /* Кнопка на всю ширину сайдбара */
        margin-bottom: 1em;
    }
    .print-button:hover { background-color: #00509e; }
    .print-button:active { background-color: #003b7a; }
    </style>
    <div class="print-button-container">
        <button class="print-button" onclick="window.print()">Drukuj wyniki</button>
    </div>
    """
    st.sidebar.markdown(print_button_html, unsafe_allow_html=True)
