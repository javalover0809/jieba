import jieba
import jieba.analyse
import xlrd

s = '养老问题是关系国计民生和国家长治久安的重大问题，为积极应对人口老龄化，建立起与人口老龄化进程相适应在党中央、国务院领导下，统筹协调全国养老服务工作'
#
# cut=jieba.cut(s)
# print(cut)
#
# #精确模式
# print('精确模式输出：')
# print('，'.join(cut))
#
jieba.suggest_freq('养老服务', True)
print(jieba.lcut(s))

for x, w in jieba.analyse.extract_tags(s, withWeight=True):
    print('%s  %s' % (x,w))