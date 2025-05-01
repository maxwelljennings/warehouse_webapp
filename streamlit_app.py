import streamlit as st
import pandas as pd
import json
import os
import io # Для работы с загруженными файлами в памяти
from collections import defaultdict, Counter
from typing import List, Dict, Optional, Set, Any
import traceback # Для отладки ошибок

# --- Константы (можно убрать, т.к. пути не фиксированы) ---
# BASE_DIR, LAYOUT_FILENAME и т.д. больше не нужны в таком виде

# --- Логика анализа (переносим и адаптируем функции) ---

# Адаптированная функция чтения layout JSON из загруженного файла
def read_layout_data(uploaded_file) -> Optional[Dict]:
    if uploaded_file is None:
        st.error("Файл структуры склада (layout.json) не загружен.")
        return None
    try:
        # Читаем содержимое загруженного файла
        stringio = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
        data = json.load(stringio)
        # Валидация
        if not isinstance(data, dict) or "Layout" not in data or not isinstance(data.get("Layout"), dict):
             raise ValueError("Файл не содержит ожидаемую структуру JSON (отсутствует ключ 'Layout').")
        st.success("Файл структуры склада успешно загружен и прочитан.")
        return data
    except json.JSONDecodeError as e:
        st.error(f"Ошибка парсинга JSON файла структуры: {e}")
        return None
    except ValueError as e:
        st.error(f"Ошибка валидации структуры файла: {e}")
        return None
    except Exception as e:
        st.error(f"Непредвиденная ошибка при чтении файла структуры: {e}")
        # traceback.print_exc() # Для отладки можно раскомментировать
        return None

# Адаптированная функция чтения пустых локаций (с B3) из загруженного файла
def read_empty_locations_from_b3(uploaded_file) -> Optional[Set[str]]:
    if uploaded_file is None:
        st.error("Файл пустых локаций (Excel/CSV) не загружен.")
        return None

    file_name = uploaded_file.name
    st.info(f"Чтение файла пустых локаций: {file_name} (данные с B3)")

    try:
        excel_engine = None
        is_csv = False
        file_ext = os.path.splitext(file_name)[1].lower()

        if file_ext == '.csv':
            is_csv = True
        elif file_ext == '.xls':
            excel_engine = 'xlrd'
            try: import xlrd
            except ImportError: raise ImportError("Необходима библиотека 'xlrd' для чтения .xls файлов.")
        elif file_ext == '.xlsx':
            excel_engine = 'openpyxl'
        else:
            st.error(f"Неподдерживаемый формат файла: {file_ext}. Используйте .xlsx, .xls или .csv.")
            return None

        df: Optional[pd.DataFrame] = None
        # Используем BytesIO для чтения загруженного файла пандасом
        file_content = io.BytesIO(uploaded_file.getvalue())

        if is_csv:
            detected_encoding = None
            for enc in ['utf-8', 'windows-1251']:
                try:
                    # Сбрасываем позицию чтения BytesIO перед каждой попыткой
                    file_content.seek(0)
                    df = pd.read_csv(file_content, delimiter=';', header=2, usecols=[1], encoding=enc, skipinitialspace=True, on_bad_lines='skip')
                    st.write(f"(CSV прочитан с кодировкой {enc})")
                    detected_encoding = enc
                    break
                except UnicodeDecodeError: continue
                except Exception as e_csv:
                     try: # Проверка, если ошибка не в кодировке
                          file_content.seek(0)
                          pd.read_csv(file_content, delimiter=';', header=2, encoding=enc, nrows=1)
                          raise e_csv # Ошибка не в кодировке/формате
                     except Exception:
                          continue # Ошибка в формате/кодировке, пробуем дальше
            if df is None:
                 raise ValueError("Не удалось прочитать CSV. Проверьте кодировку (UTF-8/Win-1251), разделитель (;) и наличие данных.")
        else: # Excel
            df = pd.read_excel(file_content, header=None, skiprows=2, usecols=[1], sheet_name=0, engine=excel_engine)

        if df is None or df.empty:
             st.warning("Файл пустых локаций не содержит данных в столбце B начиная с 3 строки.")
             return set()

        col_name = df.columns[0]
        locations = df[col_name].dropna().astype(str).str.strip()
        empty_locations_set = set(locations[locations != ''])

        count = len(empty_locations_set)
        st.success(f"Загружено {count} уникальных ID пустых локаций из {file_name}.")
        return empty_locations_set

    except ImportError as e:
         st.error(f"Ошибка импорта библиотеки: {e}. Убедитесь, что все зависимости установлены.")
         return None
    except ValueError as e:
         st.error(f"Ошибка данных или формата файла пустых локаций: {e}")
         return None
    except Exception as e:
        st.error(f"Непредвиденная ошибка при чтении файла пустых локаций: {e}")
        # traceback.print_exc() # Для отладки
        return None


# Функция анализа (точно такая же, как в CLI/GUI версии)
def analyze_opportunities_internal(layout_data: Dict, empty_locations_set: Set[str]) -> List[Dict[str, Any]]:
    """Выполняет анализ возможностей."""
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

    st.write(f"(Анализ: Обработано {processed_locations_count} локаций, {len(sections_data)} секций)")

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

    st.write(f"(Анализ: Найдено {len(opportunities)} возможностей)")
    return opportunities

# Функция форматирования результатов (упрощенная версия для оператора)
def format_analysis_results_simple(opportunities: List[Dict[str, Any]]) -> str:
    """Formats analysis results in a simplified way for operators."""
    if not opportunities:
        return "Свободных мест для размещения палет не найдено."

    standard_locations = []
    non_standard_direct = []
    non_standard_move = []

    # Предварительная обработка для дедупликации
    direct_pairs_locations = set()
    for opp in opportunities:
        if opp.get("OpportunityType") == "NonStandard_Direct":
            direct_pairs_locations.add(opp.get("PrimaryLocation", ""))
            direct_pairs_locations.add(opp.get("SecondaryLocation", ""))
            non_standard_direct.append((opp.get("PrimaryLocation", ""), opp.get("SecondaryLocation", "")))

    move_pairs_locations = set()
    for opp in opportunities:
         if opp.get("OpportunityType") == "NonStandard_MoveRequired":
              loc1 = opp.get("PrimaryLocation", "")
              loc3 = opp.get("SecondaryLocation", "") # Это loc3 из анализа
              notes = opp.get("Notes", "")
              move_pairs_locations.add(loc1)
              move_pairs_locations.add(loc3)
              move_note = ""
              if "Требуется перемещение палеты из " in notes:
                  move_loc = notes.replace("Требуется перемещение палеты из ", "").strip()
                  move_note = f"(Нужно подвинуть палету с {move_loc})"
              non_standard_move.append((loc1, loc3, move_note))


    for opp in opportunities:
        if opp.get("OpportunityType") == "Standard":
            loc1 = opp.get("PrimaryLocation", "")
            # Добавляем, только если не участвует в уже найденных парах
            if loc1 not in direct_pairs_locations and loc1 not in move_pairs_locations:
                standard_locations.append(loc1)

    # Формируем итоговый текст
    output_lines = []
    output_lines.append("--- Возможности для размещения ---")
    output_lines.append("")

    if standard_locations:
        output_lines.append("СТАНДАРТНЫЕ палеты можно поставить сюда:")
        standard_locations.sort()
        line = []
        for i, loc in enumerate(standard_locations):
            line.append(loc)
            if (i + 1) % 5 == 0 or i == len(standard_locations) - 1:
                output_lines.append("  " + ", ".join(line))
                line = []
        output_lines.append("")
    else:
        output_lines.append("Свободных мест для СТАНДАРТНЫХ палет нет.")
        output_lines.append("")

    if non_standard_direct:
        output_lines.append("НЕСТАНДАРТНЫЕ палеты (на 2 места) можно поставить сюда:")
        non_standard_direct.sort(key=lambda x: x[0])
        for loc1, loc2 in non_standard_direct:
            output_lines.append(f"  - {loc1} и {loc2}")
        output_lines.append("")
    else:
        output_lines.append("Свободных пар мест для НЕСТАНДАРТНЫХ палет нет.")
        output_lines.append("")

    if non_standard_move:
        output_lines.append("НЕСТАНДАРТНЫЕ палеты (на 2 места) можно поставить, ЕСЛИ ПОДВИНУТЬ:")
        non_standard_move.sort(key=lambda x: x[0])
        for loc1, loc3, move_note in non_standard_move:
            output_lines.append(f"  - {loc1} и {loc3} {move_note}")
        output_lines.append("")

    output_lines.append("------------------------------------")
    return "\n".join(output_lines)


# --- Streamlit UI ---
st.set_page_config(page_title="Анализатор склада", layout="wide")
st.title("Анализатор возможностей размещения на складе")

st.sidebar.header("1. Загрузите файлы")

# Загрузчик для файла структуры
uploaded_layout_file = st.sidebar.file_uploader(
    "Загрузите файл структуры склада (layout.json)",
    type=["json"]
)

# Загрузчик для файла пустых локаций
uploaded_empty_loc_file = st.sidebar.file_uploader(
    "Загрузите файл пустых локаций (Excel/CSV)",
    type=["xlsx", "xls", "csv"]
)

st.sidebar.header("2. Запустите анализ")
run_button = st.sidebar.button("Запустить анализ", disabled=(not uploaded_layout_file or not uploaded_empty_loc_file))

st.header("Результаты анализа")
results_placeholder = st.empty() # Место для вывода результатов или сообщений
results_placeholder.info("Загрузите оба файла и нажмите 'Запустить анализ' в боковой панели.")

# --- Логика выполнения при нажатии кнопки ---
if run_button:
    results_placeholder.info("Обработка файлов...")

    # Читаем файлы
    layout_data = read_layout_data(uploaded_layout_file)
    empty_locations = read_empty_locations_from_b3(uploaded_empty_loc_file)

    if layout_data and empty_locations is not None: # empty_locations может быть пустым set(), это нормально
        results_placeholder.info("Выполнение анализа...")
        try:
            opportunities = analyze_opportunities_internal(layout_data, empty_locations)
            results_text = format_analysis_results_simple(opportunities)
            # Используем st.text_area для отображения большого текста с возможностью прокрутки
            results_placeholder.text_area("Результаты:", value=results_text, height=500)

            # Добавляем кнопку для скачивания полного отчета Excel (опционально)
            if opportunities:
                 df_report = pd.DataFrame(opportunities)
                 df_report = df_report[["OpportunityType", "PrimaryLocation", "SecondaryLocation", "SectionID", "Notes"]]
                 # Конвертируем DataFrame в Excel в памяти
                 output_excel = io.BytesIO()
                 with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                      df_report.to_excel(writer, index=False, sheet_name='Opportunities')
                 excel_data = output_excel.getvalue()
                 st.sidebar.download_button(
                      label="Скачать полный отчет (Excel)",
                      data=excel_data,
                      file_name="analysis_report.xlsx",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                 )

        except Exception as e:
            st.error(f"Ошибка во время выполнения анализа: {e}")
            results_placeholder.error(f"Произошла ошибка во время анализа. Детали: {e}")
            # traceback.print_exc() # Для отладки

    else:
        # Если чтение файлов не удалось, сообщение об ошибке уже было выведено
        results_placeholder.warning("Анализ не может быть выполнен из-за ошибок чтения файлов.")