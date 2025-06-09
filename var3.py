# Решение через парсинг XML через библиотеку xml.etree.ElementTree (как будто бы канон)

# import xml
import xml.etree.ElementTree as ET

# чтение из файла
tree = ET.parse('new_data.xml')
root = tree.getroot()

# чтение из строки
# data_as_string = '<Cell ss:StyleID="s61"><Data ss:Type="String">Модель</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>'
# root = ET.fromstring(data_as_string)

# for type_tag in tree.findall('data/Cell/NamedCell'):
#     value = type_tag.find('Name').text
#     print(value)

for cell in root.findall('.//Cell'):
    named_cell = cell.find('NamedCell')
    if named_cell is not None and named_cell.get('Name') == '_FilterDatabase':
        cell.remove(named_cell)
tree.write('modified_data.xml', encoding="utf-8")

# print(root.tag)