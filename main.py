# coding:utf-8
from py2neo import Graph, Node, Relationship
import time
import xlrd
import xlwt
from datetime import date, datetime

# 打开图数据库
graph = Graph('http://localhost:7474', name='neo4j', password='1q2w3e')
# graph.delete_all()

# 打开 Excel:zzjg
file = "zzjg.xls"
wb = xlrd.open_workbook(filename=file)
sheetn = wb.sheet_by_index(0)
print(sheetn.name, sheetn.nrows, sheetn.ncols)
rowsacount = sheetn.nrows

# rows = sheet1.row_values(0) #获取行内容
# print(rows)

# cols = sheet1.col_values(0) #获取列内容
# print(cols)

rootnode = Node('v组织机构', name='中国石化')
graph.create(rootnode)
amax = 1
for i in range(0, rowsacount):
    # print(i)
    rn = sheetn.row_values(i)
    # print(rn[3].split("/")[0])
    a = len(rn[3].split("/"))
    if a > amax:
        amax = a
        print(str(amax) + str(rn))
        # node = Node('组织机构', name=rn[1])
        # graph.create(node)
        # r_i = Relationship(rootnode, '管理', node)
        # graph.create(r_i)
# 关闭图数据库
