import xlrd
import jieba.analyse
import xlwt
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
rows = sheet.row_values(11)

date = rows[0]
text = rows[1]

new_excel = excel.add_sheet(date)

new_excel.write(0, 0, "TF-IDF")
new_excel.write(0, 1, "权重")

new_excel.write(0, 2, "textRank")
new_excel.write(0, 3, "权重")

new_excel.write(0, 4, "字段")
new_excel.write(0, 5, "次数")

cnt = 1
for x, w in jieba.analyse.extract_tags(text, topK=40, withWeight=True):
    new_excel.write(cnt, 0, x)
    new_excel.write(cnt, 1, w)
    print('%s  %s' % (x, w))
    cnt = cnt + 1

print("=================")

cnt_textRank = 1
for x, w in jieba.analyse.textrank(text, topK=40, withWeight=True):
    new_excel.write(cnt_textRank, 2, x)
    new_excel.write(cnt_textRank, 3, w)
    print('%s  %s' % (x, w))
    cnt_textRank = cnt_textRank + 1

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

    new_excel.write(cnt_frequency, 4, word)
    new_excel.write(cnt_frequency, 5, count)
    print(word, count)
    cnt_frequency = cnt_frequency + 1

excel.save('/Users/Oraida/Desktop/first4.xls')