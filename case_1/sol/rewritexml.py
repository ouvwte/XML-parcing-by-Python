with open("data.xml", 'r', encoding="utf-8") as file:
    lines = file.readlines()

new_file = [line.replace('ss:', '') for line in lines]

with open("new_data.xml", "w", encoding="utf-8") as file:
    file.writelines(new_file)
