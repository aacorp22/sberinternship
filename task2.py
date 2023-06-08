import sys
import html5lib
import pandas as pd
import xml.etree.ElementTree as ET

# получение названия файла через командную строку
try:
    filename = sys.argv[1]
except IndexError:
    filename = input('Введите путь к файлу: ')

# чтение excel файла
try:
    all_file = pd.read_excel(filename, header=None)
except FileNotFoundError:
    print('Файл не найден. Проверьте название и путь к файлу')
    sys.exit(1)


def get_conversion_rate(date: str) -> float:
    # date должен быть в формате дд.мм.гггг
    url = f"https://www.cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To={date}"
    try:
        html_tables = pd.read_html(url, decimal=',', thousands='.')
    except ValueError:
        print(f'За дату {date} по ссылке {url} ничего не нашлось!')
        return None

    data = html_tables[0].to_dict("records")
    for item in data:
        if item['Валюта'] == 'Доллар США':
            return float(item['Курс']) / float(item['Единиц'])


def calculate_svalueusd_value(sdata: str, svalue: str) -> str:
    # sdata в формате дд.мм.гггг
    conv_rate = get_conversion_rate(sdata)
    try:
        svalueusd = round(float(svalue) / conv_rate, 2)
        return str(svalueusd)
    except TypeError:
        sys.exit("Некорректный курс конверсии!")


info_lines = all_file[:3][[0, 1]] # первые 3 строки
data_table = all_file[4:]
data_table.columns = data_table.iloc[0]
data_table = data_table[1:].reset_index(drop=True) # таблица с данными
data_table['IE Code'] = '0' + data_table['IE Code'].map(str) # пропадает leading 0 при чтении
data_table['SB Amount'] = data_table['SB Amount'].map(str)
data = data_table.to_dict('records') # список словарей, чтобы быстрее итерироваться по данным

root = ET.Element('CERTDATA') # формирование структуры xml
filename = ET.SubElement(root, 'FILENAME')
filename.text = info_lines[1][2]
envelope = ET.SubElement(root, 'ENVELOPE')

for item in data:
    ecert = ET.SubElement(envelope, 'ECERT')
    ET.SubElement(ecert, 'CERTNO').text = item['Ref no']
    ET.SubElement(ecert, 'CERTDATE').text = item['Issuance Date'].strftime('%Y-%m-%d')
    ET.SubElement(ecert, 'STATUS').text = item['Status']
    ET.SubElement(ecert, 'IEC').text = item['IE Code']
    ET.SubElement(ecert, 'EXPNAME').text = '"' + item['Client'] + '"'# добавлены кавычки
    ET.SubElement(ecert, 'BILLID').text = item['Bill Ref no']
    ET.SubElement(ecert, 'SDATE').text = item['SB Date'].strftime('%Y-%m-%d')
    ET.SubElement(ecert, 'SCC').text = item['SB Currency']
    ET.SubElement(ecert, 'SVALUE').text = item['SB Amount']
    ET.SubElement(ecert, "SVALUEUSD").text = calculate_svalueusd_value(
        item["SB Date"].strftime("%d.%m.%Y"), item["SB Amount"]
    )

# создание XML-файла
tree = ET.ElementTree(root)
ET.indent(tree, space='\t', level=0)
tree.write('result_data/output_task2.xml', encoding='utf-8', xml_declaration=True)
