import pandas as pd
import json
# Укажите путь к вашему файлу Excel
file_path = 'io.xlsx'

# Словарь замены для base-type
base_type_mappingServer = {
    'AI': 'Types.APS.PPZ_AI_PLC',
    'DI': 'Types.APS.PPZ_DI_PLC',
    'LG': 'Types.APS.PPZ_LG_PLC',
    'Calc_AI': 'Types.APS.PPZ_Calc_AI_PLC',
    'NOT DI': 'Types.APS.PPZ_DI_PLC',
    'FR': 'Types.APS.PPZ_LG_PLC',
    'SR': 'Types.APS.PPZ_LG_PLC',
}

base_type_mappingHMI = {
    'AI': 'BasePPZ_AI',
    'DI': 'BasePPZ_DI',
    'LG': 'BasePPZ_LG',
    'Calc_AI': 'BasePPZ_Calc_AI',
    'NOT DI': 'BasePPZ_DI',
    'FR': 'BasePPZ_LG',
    'SR': 'BasePPZ_LG',
}

# base_type_mappingHMI_ID = {
#     'AI': 'acef902a-d32a-4bd9-a2cb-28fc0665c2e4',
#     'DI': '9a8dc133-9fd1-435c-b69c-ab007c2265b8',
#     'LG': '3723cb79-9e67-44bc-bd96-1bc761ad667c',
#     'Calc_AI': '54fb7908-2a7c-4090-8d38-b304118a7af6',
#     'NOT DI': '9a8dc133-9fd1-435c-b69c-ab007c2265b8',
# }



# Чтение JSON-файла
with open('HMI_Id.json', 'r', encoding='utf-8') as json_file:
    base_type_mappingHMI_ID = json.load(json_file)


# # Получение значений из словаря
# values = list(base_type_mappingHMI_ID.values())
#
# # Вывод значений
# print(values)


def read_elements_from_excel(excel_file_path, sheet_name):
    elements_listServer = []
    elements_listHMI = []

    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

        for index, row in df.iterrows():
            base_type = row[17]  # Номер столбца, начиная с 0 (17 R)
            # Применяем замену для base-type
            if base_type in base_type_mappingServer:
                base_typeServer = base_type_mappingServer[base_type]
            else:
                base_typeServer = "ТУТ НЕ ОБРАБОТАЛОСЬ"

            elementServer = {
                'name': row[13],  # Номер столбца, начиная с 0 (13 - N)
                'base-type': base_typeServer,
                'description': row[2], #2 - C
                'title': row[6], # 6 - G
                'access-level': 'public',
                'access-scope': 'global',
                'aspect': 'Aspects.PLC',
                'uuid': '00000000-0000-0000-0000-000000000000'
            }
            elements_listServer.append(elementServer)

            if base_type in base_type_mappingHMI:
                base_typeHMI = base_type_mappingHMI[base_type]
                base_IdHMI = base_type_mappingHMI_ID[base_type]
            else:
                base_typeHMI = "ТУТ НЕ ОБРАБОТАЛОСЬ"
                base_IdHMI = '00000000-0000-0000-0000-000000000000'

            elementHMI = {
                'name': row[14],  # Номер столбца, начиная с 0
                'display-name': row[14],
                'base-type': base_typeHMI,
                'base-type-id': base_IdHMI,
                'ver': '5',
                'uuid': '00000000-0000-0000-0000-000000000000'
            }
            elements_listHMI.append(elementHMI)


    except Exception as e:
        print(f"Произошла ошибка при чтении данных из файла Excel: {e}")

    return [elements_listServer, elements_listHMI]


def dataExcelReadr():
    sheet_name = 'APS'
    elements_list = read_elements_from_excel(file_path, sheet_name)
    return elements_list
    # print(f"Количество элементов прочитанных из страницы 'APS': {len(elements_list)}")
    #
    # # Вывод массива элементов
    # for element in elements_list:
    #     print(element)
