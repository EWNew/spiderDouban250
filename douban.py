# -*- coding : utf-8 -*-
# @Time : 2020/6/7 13:48
# @Author : NYW

from bs4 import BeautifulSoup  # 网页解析，获取数据
import re  # 正则表达式，进行文字匹配
import urllib.request, urllib.error  # 制定url，获取网页数据
import xlwt  # 进行excel操作


def main():
    # 1爬取网页
    baseurl = "https://movie.douban.com/top250?start="
    dataList = getData(baseurl)
    savapath = ".\\doubanTop250.xls"
    # 3.保存数据
    savaData(dataList, savapath)


findLink = re.compile(r'<a href="(.*?)">')  # 链接
findTitle = re.compile(r'<span class="title">(.*)</span>')  # 片名
findRating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')  # 评价分数
findJudge = re.compile(r'<span>(\d*)人评价</span>')  # 评价人数
findInq = re.compile(r'<span class="inq">(.*)</span>')  # 概况


def getData(baseurl):
    datalist = []
    # 2解析数据
    for i in range(10):
        html = askUrl(baseurl + "%d" % (i * 25))
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.find_all("div", class_="item"):
            # datalist.append(item.find("div", class_="hd").find("span").string)
            item = str(item)
            link = re.findall(findLink, item)[0]
            titles = re.findall(findTitle, item)
            title01 = titles[0];
            if len(titles) == 2:
                title02 = titles[1].replace("\xa0/\xa0", "")
            else:
                title02 = " ";
            rate = float(re.findall(findRating, item)[0])
            judge = int(re.findall(findJudge, item)[0])
            inq = re.findall(findInq, item)
            if len(inq) != 0:
                inq = inq[0]
            else:
                inq = " "
            datalist.append((link, title01, title02, rate, judge, inq))
    return datalist


def savaData(dataList, savapath):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('豆瓣电影top250', cell_overwrite_ok=True)
    col = ("电影详情链接", "影片中文名", "影片外文名", "评分", "评价人数", "概况")
    for i in range(6):
        sheet.write(0, i, col[i])
    for i in range(len(dataList)):
        for j in range(6):
            sheet.write(i + 1, j, dataList[i][j])
    book.save(savapath)


def askUrl(url):
    head = {
        "User-Agent": "Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 83.0.4103.97Safari / 537.36Edg / 83.0.478.45"
    }
    request = urllib.request.Request(url, headers=head)
    html = ""
    try:
        respense = urllib.request.urlopen(request)
        html = respense.read().decode("utf-8")
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


if __name__ == "__main__":
    main()
