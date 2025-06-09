# Решение Влада (+ Леха) через регулярное выражение
import re

pattern = re.compile(r'^<(.+) .*><(.+) .*>(.+)<\/\2.+\/\1>$')

with open("testing.txt", "rt", encoding="utf-8") as file:
    while line := file.readline():
        rus_word = pattern.findall(line)[0][-1]
        print(rus_word)