import xlwt

xls = xlwt.Workbook()
sht1 = xls.add_sheet("sheet1")
sht1.write(0, 0, "列1")
sht1.write(0, 1, "列2")

sht1.write(1, 0, "a2")
sht1.write(1, 1, "b2")

xls.save('/Users/Oraida/Desktop/first.xls')
