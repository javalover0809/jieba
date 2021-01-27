import jieba.analyse

jieba.suggest_freq('养老服务', True)
jieba.suggest_freq('养老机构', True)
jieba.suggest_freq('机构养老', True)
jieba.suggest_freq('社区养老', True)
jieba.suggest_freq('居家养老', True)
jieba.suggest_freq('养老社区', True)
jieba.suggest_freq('养老公寓', True)
jieba.suggest_freq('养老保险', True)
jieba.suggest_freq('社会养老', True)


import xlrd
ExcelFile = xlrd.open_workbook('/Users/Oraida/Desktop/yanglao.xls')
read_excel = ExcelFile.sheet_by_name('整理324')

import xlwt
wbk = xlwt.Workbook()
sht1 = wbk.add_sheet("sheet1", cell_overwrite_ok=True)

for year in range(2003, 2020):
    data = ''
    for i in range(324):
        rows = read_excel.row_values(i)
        date = rows[0].split(".")[0]
        if( str(date) == str(year) ):
            data = data + rows[1]
        sht1.write(year, 0, date)
        sht1.write(year, 1, data)
    wbk.save('/Users/Oraida/Desktop/yearData.csv')