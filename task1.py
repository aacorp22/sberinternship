import sys
import pandas as pd
import xml.etree.ElementTree as ET


# получение названия файла через командную строку
try:
    filename = sys.argv[1]
except IndexError:
    filename = input("Введите путь к файлу: ")

# загрузка excel файла
try:
    all_file = pd.read_excel(filename, header=None)
except FileNotFoundError:
    print("Файл не найден. Проверьте название и путь к файлу")
    sys.exit(1)

info_lines = all_file[:3][[0, 1]] # первые 3 строки
data_table = all_file[4:]
data_table.columns = data_table.iloc[0]
data_table = data_table[1:].reset_index(drop=True) # таблица с данными
data = data_table.to_dict('records') # список словарей, чтобы быстрее итерироваться по данным
