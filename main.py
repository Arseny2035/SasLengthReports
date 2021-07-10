import xlrd
import xlsxwriter
import numpy as np

#######################################################################################

def countLength(strLength):
    ### расчитываем протяженность трубы исходя из имеющихся координат

    strLength = strLength.replace(',0 ', ',')
    a = []
    for word in strLength.split(','):
        if float(word) > 0:
            a.append(float(word))
    xy = 0
    length = 0
    ### Пробегаемся по точкам (координатам) и суммируем длину участков трубы
    while xy <= len(a)-4:
        y1 = a[xy]
        x1 = a[xy + 1]
        y2 = a[xy + 2]
        x2 = a[xy + 3]
        length += 6371 * np.arccos(np.sin(np.radians(x1)) * np.sin(np.radians(x2)) +
                                           np.cos(np.radians(x1)) * np.cos(np.radians(x2)) *
                                           np.cos(np.radians(y2 - y1)))
        xy += 2

    length = round(length, 3)
    return length

##########################################################################################

filename = "C:\\SAS2020\\Экспорт меток 2021-07-05\\Апшеронск 2021-07-05.xls"

wb = xlrd.open_workbook(filename, formatting_info=True)
sheet = wb.sheet_by_index(0)

outWorkbook = xlsxwriter.Workbook("D:\\Out.xls")
outSheet = outWorkbook.add_worksheet()

arr = np.empty((sheet.nrows, sheet.ncols), dtype=object)

for cols in range(sheet.ncols):
    for rows in range(sheet.nrows):
        arr[rows, cols] = sheet.row_values(rows)[cols]

### Задаем шапку итоговой таблицы
outSheet.write(0, 0, "Тип")
outSheet.write(0, 1, "Нас. пункт")
outSheet.write(0, 2, "Улица")
outSheet.write(0, 3, "Тип водов") ### Магистральный/квартальный
outSheet.write(0, 4, "Водовод") ### К какому сектору водоводов относится данная труба
outSheet.write(0, 5, "Принадл")### affilations - кто "собственник" трубы
outSheet.write(0, 6, "Коммент") ### Поле с данными о диаметре, инвентарн номере, материале и т.д.
outSheet.write(0, 7, "Протяж, км") ### В километрах
###

for cols in range(sheet.ncols):
    # Тип
    if arr[0, cols] == "ns1:name":
        for rows in range(1, sheet.nrows):
            outSheet.write(rows, 0, arr[rows, cols])

    # Нас. пункт
    if arr[0, cols] == "ns1:name2":
        for rows in range(1, sheet.nrows):
            outSheet.write(rows, 1, arr[rows, cols])

    # Улица
    if arr[0, cols] == "ns1:name18":
        for rows in range(1, sheet.nrows):
            outSheet.write(rows, 2, arr[rows, cols])

    # Тип водовода
    if arr[0, cols] == "ns1:name6":
        for rows in range(1, sheet.nrows):
            outSheet.write(rows, 3, arr[rows, cols])

    # Водовод
    if arr[0, cols] == "ns1:name10":
        for rows in range(1, sheet.nrows):
            outSheet.write(rows, 4, arr[rows, cols])

    # Принадлежность
    if arr[0, cols] == "ns1:name14":
        for rows in range(1, sheet.nrows):
            outSheet.write(rows, 5, arr[rows, cols])

    # Комментарий
    if arr[0, cols] == "ns1:description":
        for rows in range(1, sheet.nrows):
            outSheet.write(rows, 6, arr[rows, cols])

    # Протяженность, км
    if arr[0, cols] == "ns1:coordinates":
        for rows in range(1, sheet.nrows):
            outSheet.write(rows, 7, countLength(arr[rows, cols]))

###
cell_format_top = outWorkbook.add_format({'bold': True, 'font_color': 'blue'})

# Задаем ширину столбцов
outSheet.set_column(0, 0, 16)# Тип
outSheet.set_column(1, 1, 18)# Нас. пункт
outSheet.set_column(2, 2, 23)# Улица
outSheet.set_column(3, 3, 12)# Тип водовода
outSheet.set_column(4, 4, 9)# Водовод
outSheet.set_column(5, 5, 9)# Принадлежность
outSheet.set_column(6, 6, 12)# Комментарий
outSheet.set_column(7, 7, 8)# Протяженность, км

outWorkbook.close()




# vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
#
# print(vals(0)[1])
# l = list()
#
# for i in vals:
#     l.extend(i)
# print(l[1,3])


# for rownum in range(sheet.nrows):
#     marks = sheet.row_values(rownum)
#
# for c_mark in marks:
#     print(c_mark)




