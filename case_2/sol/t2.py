import xml.etree.ElementTree as ET

def extract_styles_fixed(input_file, output_file, target_styles):
    """
    Извлекает указанные стили из XML файла с правильной обработкой структуры
    """
    
    try:
        # Читаем файл как текст и добавляем XML декларацию если нужно
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Если нет XML декларации, добавляем её
        if not content.strip().startswith('<?xml'):
            content = '<?xml version="1.0" encoding="UTF-8"?>\n' + content
        
        # Парсим XML
        root = ET.fromstring(content)
        
        # Проверяем, что корневой элемент - Styles
        if root.tag != 'Styles':
            # Ищем элемент Styles внутри корня
            styles_root = root.find('Styles')
            if styles_root is not None:
                root = styles_root
            else:
                # Если Styles не найден, используем корень как есть
                print("Предупреждение: корневой элемент не 'Styles'")
        
        # Находим все элементы Style
        all_styles = root.findall('.//Style')
        print(f"Всего стилей в файле: {len(all_styles)}")
        
        # Создаем новый корневой элемент для выходного файла
        new_root = ET.Element('Styles')
        
        # Ищем нужные стили
        found_styles = []
        for style in all_styles:
            style_id = style.get('ss:ID')
            if style_id in target_styles:
                # Клонируем элемент для нового дерева
                new_style = ET.Element('Style')
                new_style.set('ss:ID', style_id)
                
                # Копируем все атрибуты
                for attr_name, attr_value in style.attrib.items():
                    new_style.set(attr_name, attr_value)
                
                # Копируем все дочерние элементы
                for child in style:
                    new_style.append(child)
                
                new_root.append(new_style)
                found_styles.append(style_id)
                print(f"Найден стиль: {style_id}")
        
        # Создаем дерево для нового файла
        new_tree = ET.ElementTree(new_root)
        
        # Сохраняем в файл с форматированием
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write('<!-- Extracted styles -->\n')
            
            # Преобразуем в строку и форматируем
            xml_str = ET.tostring(new_root, encoding='unicode')
            
            # Добавляем отступы для читаемости
            formatted_xml = format_xml(xml_str)
            f.write(formatted_xml)
        
        print(f"\nУспешно извлечено {len(found_styles)} стилей из {len(target_styles)} запрошенных")
        print(f"Результат сохранен в: {output_file}")
        
        # Проверяем, все ли стили найдены
        not_found = set(target_styles) - set(found_styles)
        if not_found:
            print(f"Не найдены стили: {', '.join(not_found)}")
            
    except ET.ParseError as e:
        print(f"Ошибка парсинга XML: {e}")
        print("Пробуем альтернативный метод...")
        # Если парсинг не удался, используем более простой метод
        extract_styles_simple(input_file, output_file, target_styles)
    except FileNotFoundError:
        print(f"Файл не найден: {input_file}")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

def format_xml(xml_str):
    """
    Простое форматирование XML для читаемости
    """
    import re
    # Добавляем переносы после тегов
    xml_str = re.sub(r'>\s*<', '>\n<', xml_str)
    
    # Добавляем отступы
    lines = xml_str.split('\n')
    formatted_lines = []
    indent_level = 0
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Уменьшаем отступ для закрывающих тегов
        if line.startswith('</'):
            indent_level = max(0, indent_level - 1)
        
        formatted_lines.append('  ' * indent_level + line)
        
        # Увеличиваем отступ для открывающих тегов (кроме самозакрывающихся)
        if line.startswith('<') and not line.startswith('</') and not line.endswith('/>'):
            indent_level += 1
    
    return '\n'.join(formatted_lines)

def extract_styles_simple(input_file, output_file, target_styles):
    """
    Альтернативный метод с использованием регулярных выражений
    """
    print("Используем упрощенный метод с регулярными выражениями...")
    
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Шаблон для поиска полных блоков Style
    pattern = r'<Style\s+ss:ID="([^"]+)"([^>]*)>([\s\S]*?)</Style>'
    
    matches = re.findall(pattern, content)
    
    selected_styles = []
    for style_id, attributes, style_content in matches:
        if style_id in target_styles:
            full_style = f'<Style ss:ID="{style_id}"{attributes}>{style_content}</Style>'
            selected_styles.append(full_style)
            print(f"Найден стиль: {style_id}")
    
    if selected_styles:
        # Создаем корректный XML
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write('<Styles>\n')
            for style in selected_styles:
                # Добавляем отступ для красоты
                formatted_style = re.sub(r'\n', '\n  ', style)
                f.write(f'  {formatted_style}\n')
            f.write('</Styles>\n')
        
        print(f"Успешно извлечено {len(selected_styles)} стилей")
    else:
        print("Стили не найдены!")

# Основная часть скрипта
if __name__ == "__main__":
    # Укажите пути к файлам
    input_file = "styles.xml"  # путь к вашему исходному файлу
    output_file = "selected_styles.xml"  # путь для сохранения результата
    
    # Список стилей для извлечения
    target_styles = ["s79", "s80", "s83"]
    
    # Запускаем извлечение
    extract_styles_fixed(input_file, output_file, target_styles)