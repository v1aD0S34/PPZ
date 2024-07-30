import xml.etree.ElementTree as ET
from xml.dom import minidom

from read import dataExcelReadr


def ServerGenerate():
    # Создаем корневой элемент
    omx = ET.Element('omx')
    omx.set('xmlns', 'system')
    omx.set('migration', '29')
    omx.set('xmlns:ct', 'automation.control')

    # Создаем элемент ct:object и добавляем его в корень
    ct_object = ET.SubElement(omx, 'ct:object')
    ct_object.set('name', 'PPZ')
    ct_object.set('access-level', 'public')
    ct_object.set('access-scope', 'global')
    ct_object.set('uuid', '230984dd-8c65-4ea8-8227-5f24a0b77619')

    # Создаем список элементов и добавляем их в ct_object
    # elements_list = [
    #     {
    #         'name': 'ИМЯ1',
    #         'base-type': 'ТИП1',
    #         'access-level': 'public',
    #         'access-scope': 'global',
    #         'aspect': 'Aspects.PLC',
    #         'uuid': '86198009-918d-497b-80bc-7341ef67222c',
    #         'description': 'ОПИСАНИЕ1',
    #         'title': 'ТИП2'
    #     }
    # ]

    elements_list = dataExcelReadr()[0]
    # for item in elements_list:
    #     print(item)


    # Добавляем каждый элемент из списка в ct_object
    for elem_data in elements_list:
        element = ET.SubElement(ct_object, 'ct:object')
        element.set('name', elem_data['name'])
        element.set('base-type', elem_data['base-type'])
        element.set('access-level', elem_data['access-level'])
        element.set('access-scope', elem_data['access-scope'])
        element.set('aspect', elem_data['aspect'])
        element.set('uuid', elem_data['uuid'])

        # Добавляем атрибуты к каждому элементу
        description = ET.SubElement(element, 'attribute')
        description.set('type', 'unit.System.Attributes.Description')
        description.set('value', elem_data['description'])

        title = ET.SubElement(element, 'attribute')
        title.set('type', 'unit.System.Attributes.Title')
        title.set('value', elem_data['title'])

    # Создаем XML дерево и записываем его в файл с форматированием  по тегам и без объявления версии XML
    xmlstr = minidom.parseString(ET.tostring(omx)).toprettyxml(indent="  ")

    # with open('output.omx-export', 'w', encoding='utf-8') as f:
    #     f.write(xmlstr)

    # Создаем XML дерево и записываем его в файл без объявления версии XML
    with open('APS_PLC_SERVER.omx-export', 'w', encoding='utf-8') as f:
        f.write(minidom.parseString(ET.tostring(omx)).toprettyxml(indent="  ")[23:])
