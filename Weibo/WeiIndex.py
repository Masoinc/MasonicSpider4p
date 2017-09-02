# coding=utf-8
import json
import os
from urllib import parse

import requests
import xlrd

Rootdir = os.path.abspath(os.path.dirname(os.getcwd()))
Exceldir = Rootdir + r"\Dataset\登峰杯播放社交全.xls"
Sheet = "Sheet1"
date_start = "2014-01-01"
date_stop = "2017-07-09"


def read_xls_file(directory, sheet):
    xls = xlrd.open_workbook(directory)
    s = xls.sheet_by_name(sheet)
    cols = s.col_values(0)
    return cols


def savejson(soapname, data):
    fl = open('../OutPut/14to17/' + soapname + '.json', 'w')
    fl.write(json.dumps(data, ensure_ascii=False, indent=2))
    fl.close()


def search_name(name):
    url_format = "http://data.weibo.com/index/ajax/hotword?word={}&flag=nolike&_t=0"
    # 伪造cookie
    cookie_header = {
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.81 Safari/537.36",
        "Referer": "http://data.weibo.com/index?sudaref=www.google.com"
    }

    urlname = parse.quote(name)
    first_requests = url_format.format(urlname)
    codes = requests.get(first_requests, headers=cookie_header).json()
    ids = codes["data"]["id"]
    # 伪造包头
    header = {
        "Connection": "keep-alive",
        "Accept-Encoding": "gzip, deflate, sdch",
        "Accept": "*/*",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.81 Safari/537.36",
        "Accept-Language": "zh-CN,zh;q=0.8",
        "Referer": "http://data.weibo.com/index/hotword?wid={}&wname={}".format(ids, urlname),
        "Content-Type": "application/x-www-form-urlencoded",
        "Host": "data.weibo.com"
    }

    # 获取日期
    date_url = "http://data.weibo.com/index/ajax/getdate?month=1&__rnd=1498190033389"
    dc = requests.get(date_url, headers=header).json()
    edate, sdate = dc["edate"], dc["sdate"]

    # 数据返回
    codes = requests.get("http://data.weibo.com/index/ajax/getchartdata?wid={}&sdate=2014-01-01&edate=2017-07-29".format(ids, sdate, edate), headers=header).json()
    return codes


if __name__ == "__main__":
    # 自动爬取新浪微博微指数，结果保存于Output/14to17
    dataless = ["女管家", "飘洋过海来看你", "守护丽人"]
    # 以上为未收录数据
    presoap = ["楚乔传", "醉玲珑", "我的前半生", "上古情歌"]
    for soap in presoap:
        if os.path.exists('../OutPut/14to17/' + soap + '.json'):
            print(soap + "已保存过")
            continue
        if soap == "" or dataless.__contains__(soap):
            continue
        savejson(soap, search_name(soap))
        print(soap + "已保存")
    print("已完成")
    # for soap in read_xls_file(Exceldir, Sheet):
    #     if os.path.exists('../OutPut/14to17/' + soap + '.json'):
    #         print(soap + "已保存过")
    #         continue
    #     if soap == "" or dataless.__contains__(soap):
    #         # 爱情公寓四未收录，以爱情公寓为关键词查询
    #         continue
    #     savejson(soap, search_name(soap))
    #     print(soap + "已保存")
    # print("已完成")
