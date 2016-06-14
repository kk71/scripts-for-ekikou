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
    "password": "000123"
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

# 列名对应的字段位置(页面中)
PAGE_ROW = (
    (0, "购买时间"),
    (2, "产品名"),
    (3, "金额"),
    (4, "状态")
)


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
    rst = []
    soup = bs4.BeautifulSoup(html, "html5lib")
    trs = soup.select("#theadFix > tbody > tr")
    for tr in trs:
        tr_content = []
        for td in tr.select("td"):
            if td.string and td.string.strip():
                tr_content.append(td.string.strip())
            elif td.string is None:
                tr_content.append(td.string)
        rst.append(tr_content)
    return rst # 返回的是全部数据


def filter_parsed_account_info(one_item_of_info_list):
    """
    从上面的 parse_account_info 中过滤出需要的数据(PAGE_ROW)
    是每条数据
    :param info_list:
    :return:
    """
    rst = []
    for l in PAGE_ROW:
        rst.append(one_item_of_info_list[l[0]])
    return rst


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


def generate_new_xls_filename():
    return sys.argv[1][:-4] + " - 账户导出数据(%s).xls" % arrow.now().format("YYYY-MM-DD HH:mm:ss")


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

    # 产生文件名,然后写入 xls 表的首行
    file_name = generate_new_xls_filename()
    print("输出文件: " + file_name)
    doc = xlwt.Workbook()
    sheet = doc.add_sheet("sheet1")
    sheet.write(0,0, "电话") # 第一列为电话号码
    for n in range(len(PAGE_ROW)):
        sheet.write(0,n+1,PAGE_ROW[n][1])
    doc.save(file_name)
    current_line = 1 # 当前 xls 写的行数

    for current_tel in tels:
        # FIXME only fetch the first page
        resp = sess.get(URL_QUERY_ACCOUNT_PURCHASE_INFO_WITH_TIME_RANGE, params={
            "purchaseDatebegin":time_begin,
            "purchaseDateend":time_end,
            "account": current_tel
        }, headers=req_headers, timeout=REQ_TIMEOUT)
        if not is_success(resp.status_code):
            raise Exception("请求数据时返回状态错误, code: {code}, account: {account}".format(
                code=resp.status_code,
                account=current_tel
            ))
        h = resp.content

        rst = parse_account_info(h)
        data_to_write = [[current_tel] + filter_parsed_account_info(i) for i in rst]
        for i in range(len(data_to_write)):
            for j in range(len(data_to_write[i])):
                sheet.write(current_line, j, data_to_write[i][j])
            current_line += 1 # 行数增加
        doc.save(file_name)
        print("写入%s" % current_tel)
        random_pause(len(tels))


if __name__=="__main__":
    main()