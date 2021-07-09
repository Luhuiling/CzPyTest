# -*- coding: UTF-8 -*-
'''=================================================
@File   ：爬取网易云课堂Python课程数据.py
@IDE    ：PyCharm
@Author ：Ms. Lu
@Date   ：2021/6/30 13:39
@Desc   ：网易云课堂Python课程页面。
输入网址“study.163.com”访问网易云课堂首页，在首页搜索栏中输入“Python”关键字，进入Python课程页面，然后单击“全部”选项,显示全部Python课程。
打开谷歌浏览器，使用“检查”功能分析页面。我们发现课程信息没有直接显示页面中，而是保存在studycourse.json文件中。
使用requests 模块来获取课程数据信息，使用xlsxwriter模块将课程信息写入到Excel表格。

实现逻辑：
首先使用requests模块发送POST请求,获取到当前页数的课程信息,
然后使用json()方法获取到Json格式数据，
接下来使用xlsxwriter模块将获取到的当前页数的信息写入到Excel。
最后在依次遍历每一页的课程信息。
=================================================='''

import requests
import xlsxwriter

def get_json(index):
    """
    爬取课程的json数据
    :param index: 当前索引，从0开始
    :return: Json数据
    """
    url = "https://study.163.com/p/search/studycourse.json"

    # payload信息
    payload ={
        "activityId": 0,
        "keyword": "python",
        "orderType": 50,
        "pageIndex": 1,
        "pageSize": 50,
        "priceType": -1,
        "qualityType": 0,
        "relativeOffset": 0,
        "searchTimeType": -1,
    }

    # Headers 信息
    headers = {
        "accept": "application/json",
        "Host": "study.163.com",
        "content-type": "application/json",
        "origin": "https://study.163.com",
        "user-agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Safari/537.36"
    }

    try:
        response = requests.post(url,json=payload,headers=headers) # 发送POST请求
        content_json = response.json() # 获取json数据
        if content_json and content_json["code"] == 0: # 判断数据是否存在
            return  content_json
        return None
    except Exception as e:
        print("ERROR")
        print(e)
        return None


def get_content(content_json):
    """
    获取课程信息列表
    :param content_json: 获取的json格式数据
    :return: 课程数据
    """
    if "result" in content_json:
        return content_json['result']['list'] # 返回课程数据列表

def save_excel(content,index):
    """
    存储到Excel
    :param content: 课程内容
    :param index: 索引值，从0开始
    :return: None
    """
    for num,item in enumerate(content):   # enumerate 函数用于遍历序列中的元素以及它们的下标
        row = 50*index + (num+1)

        # 行内容
        worksheet.write(row, 0, item['productId'])
        worksheet.write(row, 1, item['courseId'])
        worksheet.write(row, 2, item['productName'])
        worksheet.write(row, 3, item['productType'])
        worksheet.write(row, 4, item['provider'])
        worksheet.write(row, 5, item['score'])
        worksheet.write(row, 6, item['scoreLevel'])
        worksheet.write(row, 7, item['learnerCount'])
        worksheet.write(row, 8, item['lessonCount'])
        worksheet.write(row, 9, item['lectorName'])
        worksheet.write(row, 10, item['originalPrice'])
        worksheet.write(row, 11, item['discountPrice'])
        worksheet.write(row, 12, item['discountRate'])
        worksheet.write(row, 13, item['imgUrl'])
        worksheet.write(row, 14, item['bigImgUrl'])
        worksheet.write(row, 15, item['description'])


def main(index):
    """
    程序运行函数
    :param index: 索引值，从0开始
    :return:
    """
    # index1 = str(index)
    content_json = get_json(index)
    content = get_content(content_json)
    save_excel(content,index)


if __name__ == '__main__':
    print('START: 爬取数据中。。。。')
    workbook = xlsxwriter.Workbook("网易云课堂Python课程数据.xlsx") # 创建excel
    worksheet = workbook.add_worksheet("first_sheet") # 创建sheet

    # 行首标题
    worksheet.write(0, 0, '商品ID') #worksheet.write(0, 0, '商品ID')的第一个参数表示行（从0开始），第二个参数表示列（从0开始），第三个参数是该表格的内容。
    worksheet.write(0, 1, '课程ID')
    worksheet.write(0, 2, '商品名称')
    worksheet.write(0, 3, '商品类型')
    worksheet.write(0, 4, '机构名称')
    worksheet.write(0, 5, '评分')
    worksheet.write(0, 6, '评分等级')
    worksheet.write(0, 7, '学习人数')
    worksheet.write(0, 8, '课程节数')
    worksheet.write(0, 9, '讲师名称')
    worksheet.write(0, 10, '原价')
    worksheet.write(0, 11, '折扣价')
    worksheet.write(0, 12, '折扣率')
    worksheet.write(0, 13, '课程小图URL')
    worksheet.write(0, 14, '课程大图URL')
    worksheet.write(0, 15, '课程描述')

    # 获取总页数
    totlePageCount = get_json(1)['result']["query"]["totlePageCount"]

    # 遍历每一页
    for index in range(totlePageCount):
        main(index)

    # 关闭excel写入
    workbook.close()
    print('END：数据爬取完毕！！！')


