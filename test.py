import re
import xlrd
import xlsxwriter
import numpy as np

location = "C:\\SAS2020\\Экспорт меток 2021-07-05\\Апшеронск 2021-07-05.xls"
wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)
print(sheet.cell_value(0, 0))

outWorkbook = xlsxwriter.Workbook("D:\\Out.xls")
outSheet = outWorkbook.add_worksheet()

### Задаем шапку итоговой таблицы
outSheet.write(0, 0, "Нас. пункт")
outSheet.write(0, 1, "Улица")
outSheet.write(0, 2, "Тип водовода") ### Магистральный/квартальный
outSheet.write(0, 3, "Водовод") ### К какому сектору водоводов относится данная труба
outSheet.write(0, 4, "Принадлежность")### affilations - кто "собственник" трубы
outSheet.write(0, 5, "Комментарий") ### Поле с данными о диаметре, инвентарн номере, материале и т.д.
outSheet.write(0, 6, "Протяженность") ### В километрах
###

i = 1
length = 0
affilations = []
while i < sheet.nrows:

    ### расчитываем протяженность трубы исходя из имеющихся координат
    if sheet.cell_value(i, 25) != "":
        strLength = sheet.cell_value(i, 25)
        strLength = strLength.replace(',0 ', ',')
        a = []
        for word in strLength.split(','):
            if float(word) > 0:
                a.append(float(word))
        xy = 0
        length = 0
        ### Пробегаемся по точкам (координатам) и суммируем длину участков трубы
        while xy <= len(a)-4:
            print('xy: ', xy)
            y1 = a[xy]
            x1 = a[xy + 1]
            y2 = a[xy + 2]
            x2 = a[xy + 3]
            length += 6371 * np.arccos(np.sin(np.radians(x1)) * np.sin(np.radians(x2)) +
                                               np.cos(np.radians(x1)) * np.cos(np.radians(x2)) *
                                               np.cos(np.radians(y2 - y1)))
            xy += 2
        ###
    ###

    ### Выводим в итоговую таблицу строку с очередной трубой
    outSheet.write(i, 0, sheet.cell_value(i, 4))
    outSheet.write(i, 1, sheet.cell_value(i, 20))
    outSheet.write(i, 2, sheet.cell_value(i, 9))
    outSheet.write(i, 3, sheet.cell_value(i, 12))
    outSheet.write(i, 4, sheet.cell_value(i, 16))
    outSheet.write(i, 5, sheet.cell_value(i, 21))
    outSheet.write(i, 6, length)
    print(sheet.cell_value(i, 20))

    affilation = sheet.cell_value(i, 16)
    if affilation not in affilations:
        affilations.append([affilation])
    affilations[affilation] += 1


    i = i + 1
    ###

outWorkbook.close()
