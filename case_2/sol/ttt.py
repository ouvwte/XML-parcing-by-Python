import xml.etree.ElementTree as ET

def extract_styles(target):
    try:
        tree = ET.parse('C:/Users/BrunovDD/Desktop/py/Парсим XML через Python/XML_Parcing/case_2/data/origin_xml.xml')
        root = tree.getroot()

        all_styles = root.findall('.//Style')
        print(f"count all styles: {len(all_styles)}")
        
        new_root = ET.Element('Styles')
        result = []
        for style in all_styles:
            style_id = style.get('ss:ID')
            if style_id in needed:
                result.append(style)
        
        for style in result:
            new_root.append(style)

        new_tree = ET.ElementTree(new_root)
        new_tree.write('C:/Users/BrunovDD/Desktop/py/Парсим XML через Python/XML_Parcing/case_2/result/styles_itog.xml', encoding="utf-8") # xml_declaration=True

        not_found = set(needed) -  {style.get('ss:ID') for style in result}
        if not_found:
            print(f"Not found styles {', '.join(not_found)}")

    except ET.ParseError as e:
        print(f"XML parsing error: {e}")
    except FileNotFoundError:
        print("File not found")
    except Exception as e:
        print(f"Error! Text: {e}")


needed = ["m461349208", "s79", "s80", "s83", "s84", "m461347300", "s70", "s71", "s47", "s48", "s39", "s64", "s152", "m461347280"]

extract_styles(needed)
