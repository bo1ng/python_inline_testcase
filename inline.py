import os
import array
import json
import re
import xlrd
import csv


json_path = "D:\MyManuals\Python\Test_inline_case\eltexClass.json"
eltexSLA_path = 'D:\MyManuals\Python\Test_inline_case\eltexSLA.xlsx'
csv_path = r"D:\MyManuals\Python\Test_inline_case\result.csv"

# Чтение json
with open(json_path) as json_file:
    data = json.load(json_file)
#Выделение метрик в отдельный словарь metrics
metrics = data['metric']
#Выделение устройств в отдельный словарь devices
devices = data['devices']
    #обращение к ключу словаря
    #print(list(metrics.keys())[0])
    #обращение к значению ключа словаря - 0 имя метрики, 1 - множитель
    #print((list(metrics.values())[0])[0])
#Длина словаря (25)
len_dict = len(metrics)

# ИМИТАЦИЯ ПОЛУЧЕНИЯ ДАННЫХ
eltexSLA = xlrd.open_workbook(eltexSLA_path)
sheet = eltexSLA.sheet_by_index(0)
rows = sheet.nrows
array = []

for device in devices:
    ip = device['ip']
    n=0
    while n <= rows - 1:
        str = sheet.cell_value(n, 0)
        #получаем название метрики из строки xlsx-файла
        stat = re.sub(r'\d+|\s|[:]|[.]', '', str)
        #получаем значение метрики из строки xlsx-файла
        vl = re.sub(r"\d*: \b\w+\b.\d*", "", str)
        #получаем номер теста
        num = re.sub(r"\d*: \b\w+\b[.]", "", str)
        num = int(re.sub(r"\s\d*", "", num))
        i = 0
        #цикл по словарю
        while i <= len_dict:
            #проверка соответствия названия метрики из xlsx-файла названию метрики в словаре
            #если совпало, то изменяем имя метрики согласно словарю и умножаем значение метрики на множитель
            if stat == list(metrics.keys())[i]:
                statname = (list(metrics.values())[i])[0]
                vl = int(vl)*(list(metrics.values())[i])[1]
                break
            i=i+1
        res = ip,num,statname, vl
        array.append(res)
        n = n+1

#Вывод на экран
print(*array, sep='\n')

#Вывод в csv
file = open(csv_path, 'w+', newline='')
with file:
    header = ['ip'],['testnumber'],['metric'],['value']
    write = csv.writer(file, delimiter = ',')
    write.writerow(("IpAddress","Test Number","Metric", "Value"))
    write.writerows(array)