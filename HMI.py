import xml.etree.ElementTree as ET
from xml.dom import minidom

from read import dataExcelReadr

def HMIGenerate():
    # Создаем элемент <type>
    type_element = ET.Element('type', {
        'access-modifier': 'private',
        'name': 'PPZ_LIST',
        'display-name': 'PPZ_LIST',
        'uuid': 'c7093b5f-db85-4eab-bae9-e0722fa04139',
        'base-type': 'Form',
        'base-type-id': 'ffaf5544-6200-45f4-87ec-9dd24558a9d5',
        'ver': '5'
    })

    # Список данных для элементов <designed>
    designed_elements = [
        {'target': 'X', 'value': '0', 'ver': '5'},
        {'target': 'Y', 'value': '0', 'ver': '5'},
        {'target': 'ZValue', 'value': '0', 'ver': '5'},
        {'target': 'Rotation', 'value': '0', 'ver': '5'},
        {'target': 'Scale', 'value': '1', 'ver': '5'},
        {'target': 'Visible', 'value': 'true', 'ver': '5'},
        {'target': 'Opacity', 'value': '1', 'ver': '5'},
        {'target': 'Enabled', 'value': 'true', 'ver': '5'},
        {'target': 'Tooltip', 'value': '', 'ver': '5'},
        {'target': 'Width', 'value': '1505', 'ver': '5'},
        {'target': 'Height', 'value': '785', 'ver': '5'},
        {'target': 'PenColor', 'value': '4278190080', 'ver': '5'},
        {'target': 'PenStyle', 'value': '0', 'ver': '5'},
        {'target': 'PenWidth', 'value': '1', 'ver': '5'},
        {'target': 'BrushColor', 'value': '0xffcacaca', 'ver': '5'},
        {'target': 'BrushStyle', 'value': '1', 'ver': '5'},
        {'target': 'WindowX', 'value': '0', 'ver': '5'},
        {'target': 'WindowY', 'value': '0', 'ver': '5'},
        {'target': 'WindowWidth', 'value': '1920', 'ver': '5'},
        {'target': 'WindowHeight', 'value': '1080', 'ver': '5'},
        {'target': 'WindowCaption', 'value': 'MainForm', 'ver': '5'},
        {'target': 'ShowWindowCaption', 'value': 'true', 'ver': '5'},
        {'target': 'ShowWindowMinimize', 'value': 'true', 'ver': '5'},
        {'target': 'ShowWindowMaximize', 'value': 'true', 'ver': '5'},
        {'target': 'ShowWindowClose', 'value': 'true', 'ver': '5'},
        {'target': 'AlwaysOnTop', 'value': 'false', 'ver': '5'},
        {'target': 'WindowSizeMode', 'value': '0', 'ver': '5'},
        {'target': 'WindowBorderStyle', 'value': '1', 'ver': '5'},
        {'target': 'WindowState', 'value': '0', 'ver': '5'},
        {'target': 'WindowScalingMode', 'value': '0', 'ver': '5'},
        {'target': 'MonitorNumber', 'value': '0', 'ver': '5'},
        {'target': 'WindowPosition', 'value': '0', 'ver': '5'},
        {'target': 'WindowCloseMode', 'value': '0', 'ver': '5'},
        {'target': 'WindowIconPath', 'value': '', 'ver': '5'}
    ]

    # Добавляем элементы <designed>
    for elem in designed_elements:
        ET.SubElement(type_element, 'designed', {'target': elem['target'], 'value': elem['value'], 'ver': elem['ver']})



    # Список объектов (пример, можно добавлять больше объектов по аналогии)
    objects_list = dataExcelReadr()[1]

    # Определяем высоту
    height = 22

    # Добавляем объекты <object> к элементу <type>
    for index, obj_elem in enumerate(objects_list):  # Используем enumerate для получения индекса
        object_element = ET.SubElement(type_element, 'object', {
            'access-modifier': 'private',
            'name': obj_elem['name'],
            'display-name': obj_elem['display-name'],
            'uuid': obj_elem['uuid'],
            'base-type': obj_elem['base-type'],
            'base-type-id': obj_elem['base-type-id'],
            'ver': obj_elem['ver'],
            'description': '',
            'cardinal': '1'
        })

        # Пример добавления элементов <designed>
        designed_attributes = [
            {'target': 'X', 'value': '7', 'ver': '5'},
            {'target': 'Y', 'value': str(index * height), 'ver': '5'},
            {'target': 'Rotation', 'value': '0', 'ver': '5'},
            {'target': 'Width', 'value': '1360', 'ver': '5'},
            {'target': 'Height', 'value': str(height), 'ver': '5'}  # Устанавливаем Height равным height
        ]

        for attr in designed_attributes:
            ET.SubElement(object_element, 'designed', {'target': attr['target'], 'value': attr['value'], 'ver': attr['ver']})

        # Пример добавления <init> элементов
        init_attributes = [
            {'target': '_initUst', 'value': obj_elem['name'] if obj_elem['base-type'] in ['BasePPZ_AI', 'BasePPZ_Calc_AI'] else 'ИСТИНА'},
            {'target': '_initDelay', 'value': '0'},
            {'target': '_initNum', 'value': str(index + 1)},  # Индекс + 1
            {'target': '_init_AS', 'value': obj_elem['name']}  # Значение name объекта
        ]

        for init in init_attributes:
            ET.SubElement(object_element, 'init', {'target': init['target'], 'ver': '5', 'value': init['value']})

        # Пример добавления <do-trace> элементов
        do_trace_targets = ['Select_All', 'Sbros', 'Sbros_All_PPZ']

        for target in do_trace_targets:
            #cdata_value = f"unit.Global_main.{target}.Value"
            modified_target = (
                'Select_PPZ' if target == 'Select_All' else
                'Sbros_PPZ' if target == 'Sbros' else  # Убираем '_PPZ' для 'Sbros_PPZ'
                #target[:-4] if target == 'Sbros_All_PPZ' else  # Убираем '_PPZ' для 'Sbros_All_PPZ'
                target  # Остальные значения остаются без изменений
            )
            cdata_value = f"unit.Global_main.{modified_target}.Value"
            do_trace_element = ET.SubElement(object_element, 'do-trace',
                                             {'access-modifier': 'private', 'target': target, 'ver': '5'})
            body = ET.SubElement(do_trace_element, 'body')
            body.text = f"<![CDATA[{cdata_value}]]>"

    # Создаем дерево XML
    tree = ET.ElementTree(type_element)

    # Используем minidom для форматирования и сохранения
    xml_str = ET.tostring(type_element, encoding='utf-8', xml_declaration=True)
    pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ")

    # Выполняем замену &lt; на < и &gt; на >
    pretty_xml = pretty_xml.replace("&lt;", "<").replace("&gt;", ">")

    # Записываем отформатированный XML в файл
    with open('APS_HMI.omobj', 'w', encoding='utf-8') as f:
        f.write(pretty_xml)

