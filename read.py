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
base_type_mappingServer_PS = {
    'AI': 'Types.APS.PPZ_PS_AI_PLC',
    'DI': 'Types.APS.PPZ_PS_DI_PLC',
    'LG': 'Types.APS.PPZ_PS_LG_PLC',
    'Calc_AI': 'Types.APS.PPZ_PS_Calc_AI_PLC',
    'NOT DI': 'Types.APS.PPZ_PS_DI_PLC',
    'FR': 'Types.APS.PPZ_PS_LG_PLC',
    'SR': 'Types.APS.PPZ_PS_LG_PLC',
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

# Определяем список значений, которые нам интересны
valid_values = ['АОбс', 'АОсс', 'НОбс', 'НОсс', "ПС"]

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


def read_elements_from_excel(excel_file_path, sheet_name):
    elements_listServer = []
    elements_listHMI = []

    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

        for index, row in df.iterrows():
            typeAPS = row[6]
            print(typeAPS)
            if typeAPS in valid_values:
                print(row[17])
                base_type = row[17]  # Номер столбца, начиная с 0 (17 R)
                # Применяем замену для base-type
                # НАКИНУТЬ ПРОВЕРКУ НА ПС
                if base_type in base_type_mappingServer:
                    if typeAPS == "ПС":
                        base_typeServer = base_type_mappingServer_PS[base_type]
                    else:
                        base_typeServer = base_type_mappingServer[base_type]
                else:
                    base_typeServer = "ТУТ НЕ ОБРАБОТАЛОСЬ"

                elementServer = {
                    'name': row[14],  # Номер столбца, начиная с 0 (14 - 0)
                    'base-type': base_typeServer,
                    'description': row[2],  # 2 - C
                    'title': typeAPS,  # 6 - G
                    'access-level': 'public',
                    'access-scope': 'global',
                    'aspect': 'Aspects.PLC',
                    'uuid': '00000000-0000-0000-0000-000000000000'
                }
                elements_listServer.append(elementServer)
                print(f"длина серпвер {len(elements_listServer)}")
                if base_type in base_type_mappingHMI:
                    base_typeHMI = base_type_mappingHMI[base_type]
                    base_IdHMI = base_type_mappingHMI_ID[base_type]
                else:
                    base_typeHMI = "ТУТ НЕ ОБРАБОТАЛОСЬ"
                    base_IdHMI = '00000000-0000-0000-0000-000000000000'

                    # ИСРПАВИТЬ HMI ДЛЯ ПС
                elementHMI = {
                    'name': row[14],  # Номер столбца, начиная с 0
                    'display-name': row[14],
                    'base-type': base_typeHMI,
                    'base-type-id': base_IdHMI,
                    'ver': '5',
                    'uuid': '00000000-0000-0000-0000-000000000000',
                    'delay': row[7],
                    'ust': "ЛОЖЬ" if base_type == "Not DI" else "ИСТИНА",
                    "typeAPS": typeAPS
                }
                elements_listHMI.append(elementHMI)
                print(f"длина HMI {len(elements_listHMI)}")

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
