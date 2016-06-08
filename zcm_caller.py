#coding=utf-8
__author__ = 'kk'

import re
import sys
import xlrd, xlwt
import arrow
from requests import Session

sess = Session()

# login info
LOGIN_INFO = {
    "username": "",
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

# login url
URL_LOGIN = "http://zcm.zcmlc.com/zcm/admin/login"

# query account purchase info url
URL_QUERY_ACCOUNT_PURCHASE_INFO_WITH_PAGINATION = "http://zcm.zcmlc.com/zcm/admin/userdetailbuy?Page={page}&account={account}"
URL_QUERY_ACCOUNT_PURCHASE_INFO_WITH_TIME_RANGE = "http://zcm.zcmlc.com/zcm/admin/userdetailbuy?"

# xls tel column row name
XLS_TEL_COL_ROW_NAME = "手机号"


def parse_account_info(html):
    """

    :param html:
    :return:
    """
    return


# === read xls ===
print("读xls电话列…")
print("文件名: " + sys.argv[2])
wb = xlrd.open_workbook(sys.argv[2])
sheet1 = wb.sheet_by_index(0)
tel_col_num = sheet1.row_values(0).index(XLS_TEL_COL_ROW_NAME)
tels = sheet1.col_values(tel_col_num) # tels at its row num

# === logging ===
print("登录账户…")
user_name = input("用户名: ")
password = input("密码: ")
if not user_name or not password:
    raise Exception("用户名密码为空。")
sess.post(URL_LOGIN, data=LOGIN_INFO, headers=req_headers)

# === requests ===
print("查询数据…")
print("设置时间起始终止, 输入格式为:年年年年-月月-日日, 然后回车。")
time_begin = input("起始日期: ")
time_end = input("终止日期: ")
if time_begin:
    time_begin = arrow.get(time_begin)
    time_begin_raw = time_begin.format("YYYY-MM-DD HH:mm:ss")
    print("起始时间为: " + time_begin_raw)
if time_end:
    time_end = arrow.get(time_end)
    time_end_raw = time_end.format("YYYY-MM-DD HH:mm:ss")
    print("结束时间为: " + time_end_raw)

for current_tel in tels:
