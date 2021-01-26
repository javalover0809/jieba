import xlrd
import jieba.analyse

ExcelFile = xlrd.open_workbook('/Users/Oraida/Desktop/yanglao.xls')

print(ExcelFile.sheet_names())

sheet = ExcelFile.sheet_by_name('整理324')

rows_0 = sheet.row_values(0)
print(rows_0)
print(rows_0[1].strip())
s = rows_0[1]
jieba.suggest_freq('养老服务', True)
jieba.suggest_freq('养老机构', True)
jieba.suggest_freq('机构养老', True)
jieba.suggest_freq('社区养老', True)
jieba.suggest_freq('居家养老', True)
jieba.suggest_freq('养老社区', True)
jieba.suggest_freq('养老公寓', True)
jieba.suggest_freq('养老保险', True)
jieba.suggest_freq('社会养老', True)

print(jieba.lcut(s))

for x, w in jieba.analyse.extract_tags(s, withWeight=True):
    print('%s  %s' % (x, w))

words = jieba.lcut(s)
counts = {}
for word in words:
    if len(word) == 1:  #排除单个字符的分词结果
        continue
    else:
        counts[word] = counts.get(word,0) + 1
items = list(counts.items())
items.sort(key=lambda x:x[1], reverse=True)
for i in range(100):
    word, count = items[i]
    print(word , count)
    # print ("{0:<10}{1:>5}".format(word, count))
