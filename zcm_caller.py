#coding=utf-8
__author__ = 'kk'

import sys
import time
import random
import xlrd, xlwt
import arrow
from requests import Session
import bs4


LOGIN_INFO = {
    "username": "chenyk",
    "password": ""
}

req_headers = {
    "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Encoding":"gzip, deflate, sdch",
    "Accept-Language":"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4",
    "Connection":"keep-alive",
    "DNT":1,
    "Host":"zcm.zcmlc.com",
    "Referer":"http://zcm.zcmlc.com/zcm/admin/login",
    "Upgrade-Insecure-Requests":1,
    "User-Agent":"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"
}

URL_LOGIN = "http://zcm.zcmlc.com/zcm/admin/login"

# query account purchase info url
URL_QUERY_ACCOUNT_PURCHASE_INFO_WITH_PAGINATION = "http://zcm.zcmlc.com/zcm/admin/userdetailbuy?Page={page}&account={account}"
URL_QUERY_ACCOUNT_PURCHASE_INFO_WITH_TIME_RANGE = "http://zcm.zcmlc.com/zcm/admin/userdetailbuy?"

# xls tel column row name
XLS_TEL_COL_ROW_NAME = "手机号"

# operator name filter
XLS_NAME_FILTER = "分配"
XLS_NAME_FILTER_TO_FILTER = "陈益康"

# timeout for every request
REQ_TIMEOUT = 3


def is_success(code):
    return 200 <= code <= 299


def filter_tels(xls_sheet):
    """
    returns tels to query
    :param xls_sheet:
    :return:
    """
    tels = []
    operator_column_num = xls_sheet.row_values(0).index(XLS_NAME_FILTER)
    tel_col_num = xls_sheet.row_values(0).index(XLS_TEL_COL_ROW_NAME)
    for i in range(len(xls_sheet.col(0))):
        current_row = xls_sheet.row_values(i)
        if current_row[operator_column_num]==XLS_NAME_FILTER_TO_FILTER:
            if current_row[tel_col_num]:
                tels.append(current_row[tel_col_num])
    return tels


def parse_account_info(html):
    """
    parse account purchase info from html page
    :param html:
    :return:
    """
    return


def random_pause(tel_length):
    """
    make a random pause
    :param tel_length:
    :return:
    """
    # pause_range = {
    #     (1,50): (1,5),
    #     (51,100): (5,10),
    #     (150,200): (10,15),
    #     (200, 500): (15,20),
    # }
    time.sleep(random.randint(1,10))


def main():

    sess = Session() # 存放此次登录的 cookie

    # === read xls ===
    print("读xls电话列…")
    print("文件名: " + sys.argv[1])
    wb = xlrd.open_workbook(sys.argv[1])
    sheet1 = wb.sheet_by_index(0)
    tels = filter_tels(sheet1)
    print("搜寻到可用的电话号码数: " + str(len(tels)))

    # === logging ===
    print("登录账户…")
    # user_name = input("用户名: ")
    # password = input("密码: ")
    # if not user_name or not password:
    #     raise Exception("用户名密码为空。")
    # LOGIN_INFO.update({"username":user_name,"password":password})
    resp = sess.post(URL_LOGIN, data=LOGIN_INFO, headers=req_headers)
    if not is_success(resp.status_code):
        raise Exception("登录失败。")

    # === requests ===
    print("查询数据…")
    print("设置时间起始终止, 输入格式为:年年年年-月月-日日, 然后回车。")
    time_begin = input("起始日期: ")
    time_end = input("终止日期: ")
    if time_begin:
        time_begin = arrow.get(time_begin)
        time_begin = time_begin.format("YYYY-MM-DD HH:mm:ss")
        print("起始时间为: " + time_begin)
    if time_end:
        time_end = arrow.get(time_end)
        time_end = time_end.format("YYYY-MM-DD HH:mm:ss")
        print("结束时间为: " + time_end)

    for current_tel in tels:
        # FIXME only fetch the first page
        resp = sess.get(URL_QUERY_ACCOUNT_PURCHASE_INFO_WITH_TIME_RANGE, params={
            "purchaseDatebegin":time_begin,
            "purchaseDateend":time_end,
            "account": current_tel
        }, headers=req_headers, timeout=REQ_TIMEOUT)
        # TODO make a random pause between every request loop


if __name__=="__main__":
    main()