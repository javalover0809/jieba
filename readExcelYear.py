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

import xlwt


wbk = xlwt.Workbook()
import xlrd

ExcelFile = xlrd.open_workbook('/Users/Oraida/Desktop/yanglao.xls')
read_excel = ExcelFile.sheet_by_name('整理324')

sheets = []
for year in range(2003, 2021):
    text = ''
    for i in range(324):

        rows = read_excel.row_values(i)
        date = rows[0].split(".")[0]
        # print("date是:" + str(date))
        # print(date == str(year))
        # print(len(rows[1]))
        if date == str(year):
            print(rows[0])
            print(len(rows[1]))
            text = text + rows[1]

    sheets.append(wbk.add_sheet(str(year), cell_overwrite_ok=True))
    print("年份是:" + str(year))
    print(len(text))

    sheets[year-2003].write(0, 0, "TF-IDF")
    sheets[year-2003].write(0, 1, "权重")
    sheets[year-2003].write(0, 2, "字段")
    sheets[year-2003].write(0, 3, "次数")

    words = jieba.lcut(text)
    counts = {}
    for word in words:
        if len(word) == 1:
            continue
        else:
            counts[word] = counts.get(word, 0) + 1
    items = list(counts.items())
    items.sort(key=lambda x: x[1], reverse=True)
    length = len(items)

    cnt = 1
    for x, w in jieba.analyse.extract_tags(text, topK=length, withWeight=True):
        sheets[year-2003].write(cnt, 0, x)
        sheets[year-2003].write(cnt, 1, w)
        # print('%s  %s' % (x, w))
        cnt = cnt + 1

    cnt_frequency = 1
    for j in range(length):
        word, count = items[j]
        sheets[year-2003].write(cnt_frequency, 2, word)
        sheets[year-2003].write(cnt_frequency, 3, count)
        # print(word, count)
        cnt_frequency = cnt_frequency + 1

    print("")
    wbk.save('/Users/Oraida/Desktop/first34.xls')
