#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:tianxing
import re
import xlrd

#需要处理的文件名
file = 'D:\\test\\工作簿40.xlsx'

#打开excel数据表
workbook = xlrd.open_workbook(filename=file)
#读取电子表到列表
SheetList = workbook.sheet_names()
#读取第一个电子表的名称
SheetName = SheetList[0]
#电子表索引从0开始
Sheet1 = workbook.sheet_by_index(0)
#实例化电子表对象
Sheet1 = workbook.sheet_by_name(SheetName)
#打开关键字文件
fd = open("D:\\test.txt","r",encoding='utf-8')
#输出统计数字的文件
fw = open("D:\\test\\fw.txt","w")


for j in fd:
    name = j.strip()
    d = dict()
    dnew = {name: 0}
    d.update(dnew)
    for i in range(Sheet1.nrows):
        rows = Sheet1.row_values(i)
        # if rows[1] == name:
        if re.search(name,rows[1]):
            d[name] = d[name]+rows[2]
    print(d)
    print(d,file=fw)
fd.close()