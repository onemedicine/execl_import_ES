#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd

# from pprint import pprint
from elasticsearch import Elasticsearch
from elasticsearch import helpers

es_hosts = ["192.168.1.252:9200"]
index_name = "jokes"
doc_type = "joke"
body = []
jokelist = []
# 读取execl
def read_xls():
    # 打开文件
    book = xlrd.open_workbook(r'D:\python project\jsonInputES\joke.xlsx')
    print(book)
    sheet = book.sheets()[0]   #  sheet为一个对象，[0]代表输出第一页内容

    # 获取行数和列数
    nrows = sheet.nrows
    ncols = sheet.ncols
    print(nrows, ncols)

    # # 获取整行和整列的值（列表）
    # print(sheet.row_values(2))
    # print(sheet.col_values(2))

    # 获取每个单元格内容
    # for i in range(sheet.nrows):
    #     for j in range(sheet.ncols):
    #         print(sheet.cell(i,j).value, end=' ')
    #     # print('')   # 换行输出
    #         pass

    # 获取每一行的内容
    for i in range(sheet.nrows):
        # 跳过标题行
        if i == 0:
            continue
        # print(i)
        li = sheet.row_values(i)
        jokelist.append(li)
        # for user in URLlist:
        #     print(user[0], user[1])
        print(li)   # 输出内容为列表

    # 获取每一列的内容
    # for i in range(sheet.ncols):
    #     li = sheet.col_values(i)
    #     print(li)    # 输出内容为列表


def main():
    read_xls()
    for joke in jokelist:
        body.append({
            "_index": "jokes",
            "_type": "joke",
            "_id": int(joke[0]),
            "_source": {
                "title": joke[1],
                "class": joke[2],
                "data": joke[3]
            }
        })
    es = Elasticsearch(es_hosts)
    helpers.bulk(es, body)
    #select
    # res = es.search(index='students', body={"query": {"match_all": {}}})
    # pprint(res)
    # pprint(es.info())

if __name__ == '__main__':
    main()



# 测试python使用 clru POST 或PUT均异常
# if __name__ == '__main__':
#     read_xls()
#
#     # headers = {'content-type': 'application/json', 'Accept-Charset': 'UTF-8'}
#     headers = {
#         'Content-Type': 'application/json',
#     }
#
#     allNum = 0
#     faileNum = 0
#     # for each list
#     for joke in jokelist:
#         joke_id = int(joke[0])
#         joke_title = joke[1]
#         joke_class = joke[2]
#         joke_data = joke[3]
#
#         # url split
#         # url = 'http://192.168.1.252:9200/jokes/joke/' + str(joke_id)
#         url = 'http://192.168.1.252:9200/jokes/joke/'
#         print(url)
#         payload = '{  "title":"%s",\t"class":"%s",\t"data":"%s"}' % (joke_title, joke_class, joke_data)
#         allNum = allNum+1
#         print(payload)
#         try:
#             r = requests.post(url, data=payload, headers=headers)
#         except :
#             faileNum = faileNum + 1
#             print("erro")
#             continue
#         print("ALL: %d, FAILED: %d, SUCCESS:%d\n", allNum, faileNum, allNum-faileNum)

