# Мое решение
def izvlech_slov(line):
    start = 46
    if len(line) > start:
        substring = line[start:]
        end_index = substring.find('<')
        if end_index != -1:
            return substring[:end_index]
    return None

spisok = []

with open("testing.txt", "r", encoding="utf-8") as file:
    for line in file:
        word = izvlech_slov(line.strip())
        if word:
            spisok.append(word)
# print(spisok)
with open("itogTesting.txt", "w", encoding="utf-8") as file:
    print(*spisok, file=file, sep="\n")