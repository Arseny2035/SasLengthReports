import xlrd
import numpy as np


filename = "C:\\Users\\xxx\\Downloads\\SAS.Planet.Release.201212" \
           "\\Схема сетей (актуализация)\\2021-07-05-marks.xls"

wb = xlrd.open_workbook(filename, formatting_info=True)

sheet = wb.sheet_by_index(0)

arr = np.empty((sheet.nrows, sheet.ncols), dtype=object)

for y in range(sheet.ncols):
    for x in range(sheet.nrows):
        arr[x, y] = sheet.row_values(x)[y]



fin_arr





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




