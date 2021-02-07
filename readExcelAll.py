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
total = 0
sheets = []

for i in range(324):

    rows = read_excel.row_values(i)

    date = rows[0]
    text = rows[1]

    i_0 = total + 0
    i_1 = total + 1
    i_2 = total + 2
    i_3 = total + 3
    i_4 = total + 4

    sheets.append(wbk.add_sheet(date + "(" + str(i) + ")", cell_overwrite_ok=True))

    sheets[i].write(0, i_0, "TF-IDF")
    sheets[i].write(0, i_1, "权重")
    sheets[i].write(0, i_2, "字段")
    sheets[i].write(0, i_3, "次数")
    sheets[i].write(0, i_4, "")

    words = jieba.lcut(text)
    counts = {}
    for word in words:
        if len(word) == 1:
            continue
        else:
            counts[word] = counts.get(word, 0) + 1
    items = list(counts.items())
    items.sort(key=lambda x:x[1], reverse=True)
    length = len(items)
    cnt_frequency = 1
    for j in range(length):
        word, count = items[j]
        sheets[i].write(cnt_frequency, i_2, word)
        sheets[i].write(cnt_frequency, i_3, count)
        sheets[i].write(cnt_frequency, i_4, '')
        print(word, count)
        cnt_frequency = cnt_frequency + 1
    cnt = 1
    for x, w in jieba.analyse.extract_tags(text, topK=length, withWeight=True):
        sheets[i].write(cnt, i_0, x)
        sheets[i].write(cnt, i_1, w)
        print('%s  %s' % (x, w))
        cnt = cnt + 1
wbk.save('/Users/Oraida/Desktop/first52.xls')