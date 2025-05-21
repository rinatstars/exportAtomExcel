import json
import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd
import re
import numpy as np


# Коэффициенты пересчета содержания элементов в содержание оксидов
oxide_coefficients = {
    "P": 2.291,   # P -> P2O5
    "Ba": 1.117,  # Ba -> BaO
    "Sr": 1.183,  # Sr -> SrO
    "Ti": 1.668,  # Ti -> TiO2
    "Mn": 1.291,  # Mn -> MnO
    "V": 1.785,   # V -> V2O5
    "Cr": 1.462,  # Cr -> Cr2O3
    "Co": 1.271,  # Co -> CoO
    "Ni": 1.273,  # Ni -> NiO
    "Zr": 1.351,  # Zr -> ZrO2
    "Nb": 1.431,  # Nb -> Nb2O5
    "Sc": 1.534,  # Sc -> Sc2O3
    "Ce": 1.228,  # Ce -> CeO2
    "La": 1.173,  # La -> La2O3
    "Y": 1.270,   # Y -> Y2O3
    "Yb": 1.139,  # Yb -> Yb2O3
    "Be": 2.775,  # Be -> BeO
    "Li": 2.153,  # Li -> Li2O
    "W": 1.261,   # W -> WO3
    "Mo": 1.5,    # Mo -> MoO3
    "Sn": 1.270,  # Sn -> SnO2
    "Cu": 1.252,  # Cu -> CuO
    "Pb": 1.077,  # Pb -> PbO
    "Zn": 1.245,  # Zn -> ZnO
    "Cd": 1.142,  # Cd -> CdO
    "Bi": 1.115,   # Bi -> Bi2O3
    "Ag": 1.074,  # Ag -> Ag2O
    "Ge": 1.441,  # Ge -> GeO2
    "Ga": 1.344,  # Ga -> Ga2O3
    "As": 1.320,  # As -> As2O3
    "Sb": 1.197,   # Sb -> Sb2O3
    "B": 3.220,   # B -> B2O3
}

# Маппинг для замены имен элементов на оксиды
oxide_names = {
    "P": "P2O5",
    "Ba": "BaO",
    "Sr": "SrO",
    "Ti": "TiO2",
    "Mn": "MnO",
    "V": "V2O5",
    "Cr": "Cr2O3",
    "Co": "CoO",
    "Ni": "NiO",
    "Zr": "ZrO2",
    "Nb": "Nb2O5",
    "Sc": "Sc2O3",
    "Ce": "CeO2",
    "La": "La2O3",
    "Y": "Y2O3",
    "Yb": "Yb2O3",
    "Be": "BeO",
    "Li": "Li2O",
    "W": "WO3",
    "Mo": "MoO3",
    "Sn": "SnO2",
    "Cu": "CuO",
    "Pb": "PbO",
    "Zn": "ZnO",
    "Cd": "CdO",
    "Bi": "Bi2O3",
    "Ag": "Ag2O",
    "Ge": "GeO2",
    "Ga": "Ga2O3",
    "As": "As2O3",
    "Sb": "Sb2O3",
    "B": "B2O3"
}


# Функция для округления до n значащих цифр
def round_to_n_significant_figures(num, n):
    """
    Округляет число до заданного количества значащих цифр.
    Если число равно 0, возвращается 0.
    """
    if num == 0:
        return 0
    else:
        return round(num, n - int(np.floor(np.log10(abs(num)))) - 1)


# Функция для обработки данных из XML
def process_data(xml_root, settings):
    """
    Обрабатывает данные XML, используя настройки для коэффициентов и границ.
    Возвращает заголовок документа, заголовки столбцов и строки данных для записи в Excel.
    """
    coefficients = settings.get("coefficients", {})  # Получение коэффициентов из настроек
    oxide_conversion = settings.get("oxide_conversion", {})  # Получение настроек пересчета в оксиды
    thresholds = settings.get("thresholds", {})  # Получение пороговых значений
    significant_figures = settings.get("significant_figures", 2)  # Значащие цифры по умолчанию

    # Получение данных заголовка (например, заказчик, исполнитель)
    header_info = xml_root.find("header")

    def get_text_or_default(element, default="Введите данные"):
        if element is not None and element.text:
            text = element.text.strip()
            return text if text else default
        return default

    # Извлечение заказчика из комментария
    customer_comment = get_text_or_default(header_info.find("customer"), "Введите заказчика")
    customer_match = re.search(r'Заказчик:\s*(.+)', customer_comment)
    customer = customer_match.group(1).strip() if customer_match else "Введите заказчика"

    # Извлечение остальных данных заголовка (исполнитель, дата и методика)
    executor = get_text_or_default(header_info.find("executor"), "Введите исполнителя")
    date = get_text_or_default(header_info.find("date"), "Введите дату").replace('/', '.')
    method = get_text_or_default(header_info.find("method"), "Введите методику")
    device = get_text_or_default(header_info.find("device"), "Введите прибор")
    organization = get_text_or_default(header_info.find("organization"), "Введите организацию")

    # Формирование шапки документа
    document_header = [
        ["Заказчик:", customer],
        ["Исполнитель:", executor],
        ["Дата анализа:", date],
        ["Методика:", method],
        ["Прибор:", device],
        ["Методика:", organization]
    ]

    # Заголовки таблицы
    element_names = []
    rows = []

    # Обработка каждой пробы
    for probe in xml_root.findall("probe"):
        probe_row = [probe.find("name").text]  # Первая колонка - название пробы
        for element in probe.findall("element"):
            name = element.find("name").text

            # Проверяем наличие элемента <value>
            value_element = element.find("value")
            value = value_element.text.strip() if value_element is not None and value_element.text is not None else ""

            # Проверка, нужно ли пересчитывать в оксид
            if oxide_conversion.get(name, True):
                coefficient = coefficients.get(name, 1) * oxide_coefficients.get(name, 1)
                name = oxide_names.get(name, name)  # Замена названия элемента на название оксида
            else:
                coefficient = coefficients.get(name, 1)

            # Добавляем имена элементов в заголовки, если их еще нет
            if name not in element_names:
                element_names.append(name)

            # Получаем коэффициент и границы для текущего элемента
            #coefficient = coefficients.get(name, 1)
            lower_bound = thresholds.get(name, {}).get("lower", float('-inf'))
            upper_bound = thresholds.get(name, {}).get("upper", float('inf'))

            # Обработка значений с символами "<" или ">"
            num_value = None

            # Проверка, содержит ли value диапазон
            if '< C <' in value:
                # Обработка диапазона
                match = re.match(r'(\d*\.?\d*)\s*<\s*(\w+)\s*<\s*(\d*\.?\d*)', value)
                if match:
                    lower_bound_value = float(match.group(1))
                    upper_bound_value = float(match.group(3))
                    lower_bound_value *= coefficient  # Применение коэффициента
                    upper_bound_value *= coefficient  # Применение коэффициента

                    # Проверка границ для нижней и верхней границы
                    if upper_bound_value < lower_bound:
                        probe_row.append(f"< {round_to_n_significant_figures(lower_bound, significant_figures)}")
                    elif lower_bound_value < lower_bound:
                        probe_row.append(f"{round_to_n_significant_figures(lower_bound, significant_figures)} < C < {round_to_n_significant_figures(upper_bound_value, significant_figures)}")
                    elif lower_bound_value > upper_bound:
                        probe_row.append(f"> {round_to_n_significant_figures(upper_bound, significant_figures)}")
                    elif upper_bound_value > upper_bound:
                        probe_row.append(
                            f"{round_to_n_significant_figures(lower_bound_value, significant_figures)} < C < {round_to_n_significant_figures(upper_bound, significant_figures)}")
                    else:
                        probe_row.append(
                            f"{round_to_n_significant_figures(lower_bound_value, significant_figures)} < C < {round_to_n_significant_figures(upper_bound_value, significant_figures)}")

            elif '< ' in value:
                # Обработка значения со знаком "<"
                num_value = float(value.replace('< ', '').strip())
                num_value *= coefficient  # Применение коэффициента

                # Проверка границ
                if num_value < lower_bound:
                    probe_row.append(f"< {round_to_n_significant_figures(lower_bound, significant_figures)}")
                else:
                    probe_row.append(f"< {round_to_n_significant_figures(num_value, significant_figures)}")
            elif '> ' in value:
                # Обработка значения со знаком ">"
                num_value = float(value.replace('> ', '').strip())
                num_value *= coefficient  # Применение коэффициента

                # Проверка границ
                if num_value > upper_bound:
                    probe_row.append(f"> {round_to_n_significant_figures(upper_bound, significant_figures)}")
                else:
                    probe_row.append(f"> {round_to_n_significant_figures(num_value, significant_figures)}")
            else:
                # Обработка обычного значения
                try:
                    if value:  # Проверка, что значение не пустое
                        num_value = float(value) * coefficient  # Применение коэффициента
                    else:
                        num_value = None  # Если значение пустое
                except ValueError:
                    probe_row.append("N/A")  # Если значение не может быть преобразовано в число
                    continue

                # Проверка границ для обычных значений
                if num_value is not None:
                    if num_value < lower_bound:
                        probe_row.append(f"< {round_to_n_significant_figures(lower_bound, significant_figures)}")
                    elif num_value > upper_bound:
                        probe_row.append(f"> {round_to_n_significant_figures(upper_bound, significant_figures)}")
                    else:
                        probe_row.append(round_to_n_significant_figures(num_value, significant_figures))
                else:
                    probe_row.append("")  # Пустое значение, если данных нет

        rows.append(probe_row)

    # Добавляем заголовки в итоговые данные
    headers = ["Проба"] + element_names
    return document_header, headers, rows


# Функция для записи данных в Excel
def write_to_excel(document_header, headers, rows, output_file):
    """
    Записывает данные и шапку документа в Excel файл.
    Если шапка передана как пустая строка, она не добавляется.
    """
    # Создаем DataFrame из строк данных
    df = pd.DataFrame(rows, columns=headers)

    # Записываем в Excel
    with pd.ExcelWriter(output_file) as writer:
        if document_header != "": # Если шапка не пуста, добавляем её
            # Добавляем шапку документа
            header_df = pd.DataFrame(document_header, columns=["Описание", "Значение"])
            header_df.to_excel(writer, sheet_name="Результаты", index=False, header=False)

            # Добавляем данные после шапки
            df.to_excel(writer, sheet_name="Результаты", index=False, startrow=len(header_df) + 1)
        else:
            # Запись только данных
            df.to_excel(writer, sheet_name="Результаты", index=False)


# Функция для поиска последнего добавленного XML-файла в папке
def find_latest_xml_file(folder_path):
    """
    Ищет и возвращает последний добавленный XML файл в указанной папке.
    Если файлов нет, выбрасывает исключение.
    """
    # Ищем все файлы с расширением .xml
    xml_files = glob.glob(os.path.join(folder_path, '*.xml'))
    if not xml_files:
        raise FileNotFoundError("Нет доступных XML-файлов в указанной папке.")

    # Находим последний добавленный файл
    latest_file = max(xml_files, key=os.path.getmtime)
    return latest_file


# Основная функция
def main():
    """
    Основная функция программы:
    1. Загружает настройки.
    2. Находит последний XML файл.
    3. Обрабатывает XML.
    4. Записывает данные в Excel.
    5. Удаляет XML файл, если это указано в настройках.
    """

    # Загрузка настроек
    settings_file = "settings.json"
    try:
        with open(settings_file, 'r', encoding='utf-8') as f:
            settings = json.load(f)
    except:
        raise FileNotFoundError("!!! Проверьте файл с настройками settings.json !!!")

    # Поиск последнего XML файла в указанной папке
    xml_folder = settings.get("xml_folder", "")
    delete_xml = settings.get("delete_xml_after_processing", False) # Удалять XML после обработки?
    # Формирование имени файла на основе формата из настроек
    filename_format = settings.get("output_filename_format", "{date}_{method}_{executor}") # Формат имени файла
    xml_file = find_latest_xml_file(xml_folder) # Поиск последнего файла
    tilde = re.search(r'~\d+(?=\.xml$)', os.path.basename(xml_file)) # Поиск символа тильды с числом в названии файла
    print(xml_file)

    # Парсинг XML
    if xml_file:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Обработка данных
        document_header, headers, rows = process_data(root, settings)

        # Формируем шапку документа на основе настроек
        document_header_formated = []
        if settings["header_fields"].get("customer", True):
            document_header_formated.append(document_header[0])
        if settings["header_fields"].get("executor", True):
            document_header_formated.append(document_header[1])
        if settings["header_fields"].get("date", True):
            document_header_formated.append(document_header[2])
        if settings["header_fields"].get("method", True):
            document_header_formated.append(document_header[3])
        if settings["header_fields"].get("device", True):
            document_header_formated.append(document_header[4])
        if settings["header_fields"].get("organization", True):
            document_header_formated.append(document_header[5])


        # Формирование имени файла на основе формата из настроек
        filename_format = settings.get("output_filename_format", "{date}_{method}_{executor}")
        output_filename = filename_format.format(
            customer=document_header[0][1],
            executor=document_header[1][1],
            date=document_header[2][1],
            method=document_header[3][1],
            device=document_header[4][1],
            organization=document_header[5][1]
        )

        # Полный путь к выходному файлу
        output_file = os.path.join(xml_folder, f"{output_filename}{tilde.group() if tilde else ''}.xlsx")

        # Запись в Excel
        if settings["include_header"]:
            write_to_excel(document_header_formated, headers, rows, output_file)
        else:
            write_to_excel("", headers, rows, output_file)

        # Удаление XML файла после обработки, если это указано в настройках
        if delete_xml:
            os.remove(xml_file)
            print(f"XML файл {xml_file} был удален.")
        else:
            print(f"XML файл {xml_file} был сохранен.")
    else:
        print("Не найдено подходящих XML файлов.")


if __name__ == "__main__":
    main()
