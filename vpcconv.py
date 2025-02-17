import csv
import re
import pandas as pd
import time
import os


start_time = time.time()
filename = "input.txt"


def extract_data(lines):
    '''Функция для извлечения данных из строки файла input.txt.'''
    data = {
        'DCS_TAG': '',
        'TagComment': '',
        'Type': '',
        'ITEM': '',
        'ItemComment': '',
        'RangeHI': '',
        'RangeLO': '',
        'UNIT': '',
        'DataType': '',
        'Array1': '',
        'Array2': '',
        'MarType': '',
        'Process': '',
        'VM_UNIT': '',
        'VAR': '',
        'VMArray1': ''
    }

    for line in lines:
        uname_match = re.search(r"/UNAME\s+=([^\s]+)\s+([^\s]+)", line)
        if uname_match == None:
            uname_match = re.search(r"/UNAME\s+=([^\s]+)", line)
        utype_match = re.search(r"/UTYPE\s+=([^\s]+)", line)
        vname_match = re.search(r"/VNAME\s+=([^\s]+)\s+([^\s]+)", line)
        if vname_match == None:
            vname_match = re.search(r"/VNAME\s+=([^\s]+)", line)
        scale_match = re.search(r"/SCALE\s+=([^\s]+)\s+([^\s]+)", line)
        vtype_match = re.search(r"/VTYPE\s+=([^\s]+)\s+([^\s]+)\s+([^\s]+)",
                                line)
        refer_match = re.search(r"/REFER\s+=([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)", line)

        if uname_match:
            if uname_match.lastindex > 1:
                data['DCS_TAG'] = uname_match.group(1)
                data['TagComment'] = uname_match.group(2)
            else:
                data['DCS_TAG'] = uname_match.group(1)
                data['TagComment'] = 'non comment'
        elif utype_match:
            data['Type'] = utype_match.group(1)
        elif vname_match:
            if vname_match.lastindex > 1:
                data['ITEM'] = vname_match.group(1)
                data['ItemComment'] = vname_match.group(2)
            else:
                data['ITEM'] = vname_match.group(1)
                data['ItemComment'] = 'non comment'
        elif scale_match:
            data['RangeHI'] = scale_match.group(1)
            data['RangeLO'] = scale_match.group(2)
        elif vtype_match:
            data['UNIT'] = vtype_match.group(1)
            data['DataType'] = vtype_match.group(2)
            data['Array1'] = vtype_match.group(3)
        elif refer_match:
            data['Array2'] = refer_match.group(1)
            data['MarType'] = refer_match.group(2)
            data['Process'] = refer_match.group(3)
            data['VM_UNIT'] = refer_match.group(4)
            data['VAR'] = refer_match.group(5)
            data['VMArray1'] = refer_match.group(6)

    return data


# Список заголовков
headers = ['DCS_TAG', 'TagComment', 'Type', 'ITEM', 'ItemComment', 'RangeHI',
           'RangeLO', 'UNIT', 'DataType', 'Array1', 'Array2', 'MarType',
           'Process', 'VM_UNIT', 'VAR', 'VMArray1']

# Список для хранения всех данных
all_data = []

# Чтение файла и извлечение данных
if os.path.isfile(filename):
    print(f"Файл '{filename}' существует в текущей директории.")
    with open('input.txt', 'r', encoding='utf-8') as file:
        item_lines = []
        for line in file:
            if "/ITEM" in line and item_lines:
                data = extract_data(item_lines)
                all_data.append(data)
                item_lines = []
            item_lines.append(line)
        if item_lines:
            data = extract_data(item_lines)
            all_data.append(data)

    # Запись данных в CSV файл
    with open('output.csv', 'w', newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.DictWriter(csv_file, fieldnames=headers)
        csv_writer.writeheader()
        for data in all_data:
            csv_writer.writerow(data)

    print("Данные успешно извлечены и сохранены в файл output.csv")

    # Чтение CSV файла
    df = pd.read_csv('output.csv')

    # Запись в Excel файл
    df.to_excel('output.xlsx', index=False)

    print("Данные успешно сохранены в файл output.xlsx")

    end_time = time.time()

    execution_time_ms = (end_time - start_time) * 1000

    print(f"Время выполнения программы в миллисекундах: {execution_time_ms:.3f}")

else:
    print(f"Файл '{filename}' не найден в текущей директории.")
