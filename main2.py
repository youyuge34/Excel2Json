#!/usr/bin/env python
# encoding: utf-8
"""
@author: yousheng
@contact: 1197993367@qq.com
@site: http://youyuge.cn

@version: 1.0
@license: Apache Licence
@file: main2.py
@time: 17/10/14 下午4:47

"""
import xlrd
from collections import OrderedDict
import json
import codecs
import os

def excel2json(path):
    wb = xlrd.open_workbook(path)

    convert_list = []
    sh = wb.sheet_by_index(0)
    title = sh.row_values(0)
    for rownum in range(1, sh.nrows):
        rowvalue = sh.row_values(rownum)
        single = OrderedDict()
        for colnum in range(0, len(rowvalue)):
            # print(title[colnum], rowvalue[colnum])
            single[title[colnum]] =rowvalue[colnum]
        convert_list.append(single)

    j = json.dumps(convert_list,indent=4,encoding='utf-8',ensure_ascii=False)
    return j

if __name__ == '__main__':
    rootdir =os.path.abspath(os.path.dirname(__file__))
    origindir = os.path.join(rootdir,'origins')
    for path in os.listdir(origindir):
        print 'running...',path
        j = excel2json(os.path.join(origindir,path))
        with codecs.open('output/'+path.split('.')[0].split('/')[-1]+'.json', "w", "utf-8") as f:
            f.write(j)