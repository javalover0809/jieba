import xlrd
import jieba.analyse
import xlwt
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


excel = xlwt.Workbook()

ExcelFile = xlrd.open_workbook('/Users/Oraida/Desktop/yanglao.xls')
sheet = ExcelFile.sheet_by_name('整理324')

rows = sheet.row_values(13)

date = rows[0]
text = rows[1]

new_excel = excel.add_sheet("sheet1")

new_excel.write(0, 0, "TF-IDF")
new_excel.write(0, 1, "权重")

# new_excel.write(0, 2, "textRank")
# new_excel.write(0, 3, "权重")

new_excel.write(0, 2, "字段")
new_excel.write(0, 3, "次数")

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
for i in range(len(items)):
    word, count = items[i]

    new_excel.write(cnt_frequency, 2, word)
    new_excel.write(cnt_frequency, 3, count)
    new_excel.write(cnt_frequency, 4, '')
    print(word, count)
    cnt_frequency = cnt_frequency + 1

print("=================")

cnt = 1
for x, w in jieba.analyse.extract_tags(text, topK=len(items), withWeight=True):
    new_excel.write(cnt, 0, x)
    new_excel.write(cnt, 1, w)
    print('%s  %s' % (x, w))
    cnt = cnt + 1

# print('分词及词性：')
# result = psg.lcut(text)
# print([(x.word, x.flag) for x in result])

excel.save('/Users/Oraida/Desktop/first7.xls')