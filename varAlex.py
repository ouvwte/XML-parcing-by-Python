# решение Лехи через рекурсию
# import xml
import xml.etree.ElementTree as ET

def print_tree(element, offset=0):
    if element.text():
        print(f'{' ' * offset} {element.tag}:{element.tag}')
    for child in element.child:
        print_element(child, offset+1)

tree = ET.parse('new_data.xml')
root = tree.getroot()