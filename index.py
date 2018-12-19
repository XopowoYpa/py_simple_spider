# -*- coding: utf-8 -*-
from urllib import request
from urllib import parse
from bs4 import BeautifulSoup
import random
import re
import json
import xlsxwriter

#随机user agent 以防被封ip
agentList = ["Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1",
    "Mozilla/5.0 (X11; CrOS i686 2268.111.0) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1092.0 Safari/536.6",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1090.0 Safari/536.6",
    "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/19.77.34.5 Safari/537.1",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.9 Safari/536.5",
    "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
    "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_0) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.0 Safari/536.3",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24",
    "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24"
]

# 获取豆瓣网上的所有标签并且生成dict返回 
def getAllTags():
    dic = {}
    url = 'https://www.douban.com/tag/'
    headers = {
        'User-Agent':random.choice(agentList)
    }

    req = request.Request(url,headers=headers)
    res = request.urlopen(req).read()
    soup = BeautifulSoup(res,'html.parser')
    for item in soup.find_all(class_='topic-list'):
        for i in item.find_all('a'):
            dic[i.string] = i['href']
    return dic

# 整合所需要的图书列表URL并返回 
# tag: 用户输入的标签
# tagUrl: 输入标签所对应的内部url
# bookSum: 所需要的书本最大个数
# return: 与豆瓣网服务器交互的特定url
def generateSearchUrl(tag,tagUrl,bookSum):
    # 例子 https://www.douban.com/j/tag/items?start=0&limit=6&topic_id=238&topic_name=cosplay&mod=book#
    baseUrl = 'https://www.douban.com/j/tag/items?start=0'
    headers = {
        'User-Agent':random.choice(agentList)
    }

    req = request.Request(tagUrl,headers=headers)
    res = request.urlopen(req).read()
    htmlStr = str(res,'utf-8')
    # 利用正则表达式来寻找topic_id这个参数 之后再整合进url 
    searchIdRes = re.search(r'topic_id: \d+,',htmlStr).group()
    tagId = re.search(r'\d+',searchIdRes).group()
    baseUrl = baseUrl + "&limit=" + str(bookSum)
    baseUrl = baseUrl + "&topic_id=" + tagId
    baseUrl = baseUrl + "&topic_name=" + parse.quote(tag) + "&mod=book#"
    return baseUrl

# 清空字符串中的回车和空格 
# s: 指定字符串
# return: 格式化后的字符串
def clear(s):
    res = ''
    for item in s:
        if not(item is ' ' or item is '\n'):
            res = res + item
    return res

# 返回解析的数据
# url: 与豆瓣服务器交互的url
# return: 一个二维数组里面是用户所需求的若干固定格式（书名 评分 价格 出版时间 出版社 作者 封面图链接 具体内容链接） 的数据
def analysisJsonData(url):
    data = []
    headers = {
        'User-Agent':random.choice(agentList)
    }

    req = request.Request(url,headers=headers)
    res = request.urlopen(req).read()
    # 获取json数据后解析 #
    jsonData = json.loads(res)
    if jsonData['r'] is 0:
        soup = BeautifulSoup(jsonData['html'],'html.parser')
        for item in soup.find_all('dl'):
            tmp = []
            tmp.append(clear(item.find(class_='title').string))
            if item.find(class_='rating_nums'):
                tmp.append(clear(item.find(class_='rating_nums').string))
            else:
                tmp.append('')
            strList = item.find(class_='desc').string.split(' / ')
            tmp.append(clear(strList[-1]))
            tmp.append(clear(strList[-2]))
            tmp.append(clear(strList[-3]))
            strTmp = ''
            for i in range(len(strList)-3):
                strTmp = strTmp + clear(strList[i]) +' '
            tmp.append(strTmp)
            tmp.append(item.find('img')['src'])
            tmp.append(item.find('a')['href'])
            data.append(tmp)
    return data

# 输出成表格 
# tag: 用户输入的标签
# data: 已获取分析的数据
def generateXlsx(tag,data):
    name = tag + '.xlsx'
    workbook = xlsxwriter.Workbook(name)
    worksheet = workbook.add_worksheet()
    
    row=1
    col=0
    titles = ['书名','评分','价格','出版时间','出版社','作者','封面图链接','具体内容链接']
    for item in titles:
        worksheet.write(0,col,item)
        col = col + 1
    
    for item in data:
        col = 0
        for i in item:
            worksheet.write(row,col,i)
            col = col + 1
        row = row + 1
    workbook.close()
    

# 主函数 #
def main():
    tagDics = getAllTags()

    print('经过爬取可得到豆瓣一共有以下图书标签：')
    for tag in tagDics:
        print(tag+' ',end='')
    inputTag = input('\n您要选择的标签是：')
    if tagDics.__contains__(inputTag):
        bookSum = int(input('请输入爬取最大的书本数目：'))
        print('正在爬取中...')
        url = generateSearchUrl(inputTag,tagDics[inputTag],bookSum)
        print('正在分析中...')
        data = analysisJsonData(url)
        if not data:
            print('您选择的标签没有图书,请重新运行程序')
            return
        print('正在为您生成图书信息的excel表格,请稍等...')
        generateXlsx(inputTag,data)
        print('生成成功！')
    else:
        print('不存在此标签！')

if __name__ == '__main__':
    main()