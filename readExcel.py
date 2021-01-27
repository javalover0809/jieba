
import jieba.analyse

import jieba.posseg as psg

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
excel = xlwt.Workbook()
new_excel = excel.add_sheet("sheet2",cell_overwrite_ok=True)

import xlrd
ExcelFile = xlrd.open_workbook('/Users/Oraida/Desktop/yanglao.xls')
read_excel = ExcelFile.sheet_by_name('整理324')

total = 0
for i in range(2):
    rows = read_excel.row_values(i)

    date = rows[0]
    text = rows[1]

    i_0 = total + 0
    i_1 = total + 1
    i_2 = total + 2
    i_3 = total + 3
    i_4 = total + 4

    new_excel.write(0, i_0, date)
    new_excel.write(0, i_1, "权重")
    new_excel.write(0, i_2, "字段")
    new_excel.write(0, i_3, "次数")
    new_excel.write(0, i_4, "")

# print("=================")
#
# cnt_textRank = 1
# for x, w in jieba.analyse.textrank(text, topK=50, withWeight=True):
#     new_excel.write(cnt_textRank, 2, x)
#     new_excel.write(cnt_textRank, 3, w)
#     print('%s  %s' % (x, w))
#     cnt_textRank = cnt_textRank + 1

    words = jieba.lcut(text)
    counts = {}
    for word in words:
        if len(word) == 1:
            continue
        else:
            counts[word] = counts.get(word, 0) + 1
    items = list(counts.items())
    items.sort(key=lambda x:x[1], reverse=True)

    cnt_frequency = 1
    for i in range(50):
        word, count = items[i]

        new_excel.write(cnt_frequency, i_2, word)
        new_excel.write(cnt_frequency, i_3, count)
        new_excel.write(cnt_frequency, i_4, '')
        print(word, count)
        cnt_frequency = cnt_frequency + 1

    print("=================")

    cnt = 1
    for x, w in jieba.analyse.extract_tags(text, topK=50, withWeight=True):
        new_excel.write(cnt, i_0, x)
        new_excel.write(cnt, i_1, w)
        print('%s  %s' % (x, w))
        cnt = cnt + 1

# print('分词及词性：')
# result = psg.lcut(text)
# print([(x.word, x.flag) for x in result])
    total = total + 5
    excel.save('/Users/Oraida/Desktop/first11.xls')