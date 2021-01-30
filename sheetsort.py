import xlwt
wbk = xlwt.Workbook()

import xlrd
ExcelFile = xlrd.open_workbook('/Users/Oraida/Desktop/最终结果.xls')
sheet_names = []
for i in range(324):
    sheet_names.append(ExcelFile.sheet_by_index(i).name)
sheet_names.sort(reverse=True)
print(sheet_names)
for sheet_order in sheet_names:
    print(sheet_order)